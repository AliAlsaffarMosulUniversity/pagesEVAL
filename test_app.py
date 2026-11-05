import base64, copy, json, pathlib, sys, threading, unittest, urllib.request, urllib.error, zipfile
import pymupdf as fitz
import app

SETTINGS={'count':3,'code':'P3S12D1A2023','size':12,'codeBox':{'x':54,'y':9,'w':19,'h':3.5},'numberBox':{'x':89,'y':82,'w':7,'h':4}}
def fixture(n=3,rotated=False):
    d=fitz.open()
    for i in range(n):
        p=d.new_page(width=595+i*10,height=842)
        p.insert_text((40,200),f'Original content {i+1}')
        p.draw_rect(fitz.Rect(30,210,300,350),color=(0,0,1))
        if rotated:p.set_rotation(90)
    return d.tobytes()

def word_fixture(path):
    with zipfile.ZipFile(path,'w') as z:
        z.writestr('[Content_Types].xml','''<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>''')
        z.writestr('_rels/.rels','''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>''')
        z.writestr('word/document.xml','''<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Word conversion test</w:t></w:r></w:p><w:p><w:r><w:br w:type="page"/><w:t>Second page content</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>''')

class Tests(unittest.TestCase):
    def test_sequence_and_preservation(self):
        source=fixture(); out=app.stamp(source,SETTINGS)
        a=fitz.open(stream=source,filetype='pdf'); b=fitz.open(stream=out,filetype='pdf')
        self.assertEqual(len(b),3)
        for i in range(3):
            self.assertEqual(a[i].rect,b[i].rect)
            self.assertIn(f'Original content {i+1}',b[i].get_text())
            self.assertIn(SETTINGS['code'],b[i].get_text())
            self.assertEqual(b[i].get_text().strip().splitlines()[-1],str(i+1))
            # The whole content region stays pixel-identical.
            region=fitz.Rect(0,150,500,400)
            self.assertEqual(a[i].get_pixmap(clip=region).samples,b[i].get_pixmap(clip=region).samples)
    def test_template(self):
        s=dict(SETTINGS,mode='template'); out=app.stamp(fixture(1),s)
        d=fitz.open(stream=out,filetype='pdf');self.assertEqual(len(d),3)
        self.assertIn('3',d[2].get_text())
    def test_invalid(self):
        for s in [dict(SETTINGS,count=2),dict(SETTINGS,count=0),dict(SETTINGS,code=''),dict(SETTINGS,size=90),dict(SETTINGS,codeBox={'x':99,'y':1,'w':20,'h':4})]:
            with self.assertRaises(ValueError):app.stamp(fixture(),s)
    def test_arabic(self):
        out=app.stamp(fixture(),dict(SETTINGS,code='المحور ١'))
        self.assertGreater(len(out),1000)
    def test_rotation(self):
        data=fixture(1,True); pdf,n=app.load_document('test.pdf',data)
        a=fitz.open(stream=data,filetype='pdf');b=fitz.open(stream=pdf,filetype='pdf')
        self.assertEqual(a[0].get_pixmap().samples,b[0].get_pixmap().samples)
    def test_http(self):
        server=app.ThreadingHTTPServer(('127.0.0.1',0),app.Handler)
        threading.Thread(target=server.serve_forever,daemon=True).start()
        base=f'http://127.0.0.1:{server.server_port}'
        def post(path,data,token=app.TOKEN):
            req=urllib.request.Request(base+path,json.dumps(data).encode(),headers={'X-Session':token,'Content-Type':'application/json'})
            return urllib.request.urlopen(req)
        try:
            self.assertIn('ترقيم',urllib.request.urlopen(base).read().decode())
            with self.assertRaises(urllib.error.HTTPError):post('/export',SETTINGS,'bad')
            self.assertEqual(json.load(post('/upload',{'name':'test.pdf','data':base64.b64encode(fixture()).decode()}))['pages'],3)
            self.assertTrue(post('/render',dict(SETTINGS,page=2,stamped=True)).read().startswith(b'\x89PNG'))
            self.assertTrue(post('/export',SETTINGS).read().startswith(b'%PDF'))
        finally:server.shutdown();server.server_close()

if __name__=='__main__':
    if '--word-fixture' in sys.argv:word_fixture(pathlib.Path(sys.argv[-1]))
    elif '--word-test' in sys.argv:
        path=pathlib.Path(sys.argv[-1]);pdf,n=app.load_document(path.name,path.read_bytes())
        d=fitz.open(stream=pdf,filetype='pdf');assert n==2 and 'Second page content' in d[1].get_text()
        s=dict(SETTINGS,count=2);out=app.stamp(pdf,s);assert len(fitz.open(stream=out,filetype='pdf'))==2
        print('WORD CONVERSION + STAMP: PASS')
    else:unittest.main()
