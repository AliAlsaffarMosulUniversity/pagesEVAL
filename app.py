"""Local document stamping server. Documents never leave this computer."""
import base64, json, os, pathlib, secrets, shutil, subprocess, tempfile, threading, webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import pymupdf as fitz

ROOT = pathlib.Path(__file__).parent
TOKEN = secrets.token_urlsafe(24)
LOCK = threading.Lock()
SOURCE = None
MAX_BYTES = 60 * 1024 * 1024

def office_path():
    candidates = [os.environ.get('SOFFICE'), shutil.which('soffice'),
        r'C:\Program Files\LibreOffice\program\soffice.exe']
    runtime = pathlib.Path.home()/'.cache/codex-runtimes/codex-primary-runtime/dependencies/native'
    if runtime.exists():
        candidates += [str(p) for p in runtime.rglob('soffice.exe')]
    return next((p for p in candidates if p and pathlib.Path(p).is_file()), None)

def load_document(name, data):
    ext = pathlib.Path(name).suffix.lower()
    if ext == '.pdf':
        pdf = data
    elif ext in ('.doc', '.docx'):
        office = office_path()
        with tempfile.TemporaryDirectory(prefix='page-marker-') as folder:
            folder = pathlib.Path(folder)
            source = folder/('source'+ext)
            source.write_bytes(data)
            target = folder/'source.pdf'
            if os.name=='nt':
                command=['powershell.exe','-NoProfile','-ExecutionPolicy','Bypass','-File',str(ROOT/'convert-word.ps1'),str(source),str(target)]
            elif office:
                command=[office, '-env:UserInstallation='+ (folder/'profile').as_uri(), '--headless','--convert-to','pdf','--outdir',str(folder),str(source)]
            else:
                raise ValueError('تحويل Word يحتاج Microsoft Word على Windows أو LibreOffice على الأنظمة الأخرى.')
            subprocess.run(command,capture_output=True, timeout=120,
                creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0))
            if not target.exists():
                raise ValueError('تعذر تحويل Word. تحقق أن الملف غير محمي ويمكن فتحه في LibreOffice.')
            pdf = target.read_bytes()
    else:
        raise ValueError('اختر ملف PDF أو DOC أو DOCX.')
    with fitz.open(stream=pdf, filetype='pdf') as doc:
        if doc.needs_pass or not len(doc):
            raise ValueError('الملف محمي بكلمة مرور أو خالٍ من الصفحات.')
        if len(doc)>1000:
            raise ValueError('الحد الأقصى 1000 صفحة.')
        # Normalize rotation without changing the visual appearance.
        for page in doc:
            page.remove_rotation()
        return doc.tobytes(), len(doc)

def stamp(pdf, settings):
    from html import escape
    doc = fitz.open(stream=pdf, filetype='pdf')
    try:
        count = int(settings['count'])
        if count < 1 or count > 1000: raise ValueError('عدد الصفحات يجب أن يكون بين 1 و1000.')
        mode = settings.get('mode','document')
        if mode == 'template':
            if len(doc)!=1: raise ValueError('وضع القالب يتطلب مستنداً من صفحة واحدة.')
            original = doc
            doc = fitz.open()
            for _ in range(count): doc.insert_pdf(original)
            original.close()
        elif count != len(doc):
            raise ValueError('عدد الصفحات لا يطابق المستند. لا تُحذف صفحات أو تُكرر تلقائياً. استخدم قالب صفحة واحدة للتكرار.')
        code = str(settings['code']).strip()
        if not code or len(code)>120: raise ValueError('أدخل رمز محور من 1 إلى 120 حرفاً.')
        size = float(settings.get('size',12))
        if not 6<=size<=40: raise ValueError('حجم الخط يجب أن يكون بين 6 و40.')
        for i,page in enumerate(doc):
            for key,value in [('code',code),('number',str(i+1))]:
                b = settings[key+'Box']
                x,y,w,h = (float(b[k]) for k in ('x','y','w','h'))
                if not all(map(__import__('math').isfinite,[x,y,w,h])) or min(x,y)<0 or min(w,h)<=0 or x+w>100 or y+h>100:
                    raise ValueError('يجب أن تقع الخانتان بالكامل داخل الصفحة.')
                rect = fitz.Rect(x/100*page.rect.width,y/100*page.rect.height,(x+w)/100*page.rect.width,(y+h)/100*page.rect.height)
                # No white rectangle or redaction: keep all source content intact.
                html = '<div dir="auto">'+escape(value)+'</div>'
                spare,scale = page.insert_htmlbox(rect,html,css=f'* {{font-family: sans-serif; font-size:{size}pt; color:#111;}} div {{text-align:center;}}',scale_low=1)
                if spare<0: raise ValueError('النص لا يتسع داخل الخانة. زد العرض أو الارتفاع أو صغّر الخط.')
        return doc.tobytes(garbage=3,deflate=True)
    finally: doc.close()

class Handler(BaseHTTPRequestHandler):
    def log_message(self,*args): pass
    def reply(self,status,data,kind='application/json'):
        if isinstance(data,dict): data=json.dumps(data,ensure_ascii=False).encode()
        self.send_response(status); self.send_header('Content-Type',kind)
        self.send_header('Content-Length',str(len(data))); self.send_header('Cache-Control','no-store')
        self.send_header('X-Content-Type-Options','nosniff'); self.end_headers(); self.wfile.write(data)
    def do_GET(self):
        if self.path=='/': self.reply(200,(ROOT/'index.html').read_bytes(),'text/html; charset=utf-8')
        elif self.path=='/session': self.reply(200,{'token':TOKEN})
        else: self.reply(404,{'error':'غير موجود'})
    def do_POST(self):
        global SOURCE
        if self.headers.get('X-Session')!=TOKEN:
            self.reply(403,{'error':'جلسة غير صالحة'}); return
        try:
            length=int(self.headers.get('Content-Length','0'))
            if not 0<length<MAX_BYTES*1.4: raise ValueError('حجم الملف يتجاوز 60 ميغابايت.')
            body=json.loads(self.rfile.read(length))
            with LOCK:
                if self.path=='/upload':
                    data=base64.b64decode(body['data'],validate=True)
                    if len(data)>MAX_BYTES: raise ValueError('الحد الأقصى 60 ميغابايت.')
                    pdf,n=load_document(body['name'],data)
                    SOURCE=pdf
                    self.reply(200,{'pages':n}); return
                if SOURCE is None: raise ValueError('ارفع المستند أولاً.')
                if self.path=='/render':
                    pdf=stamp(SOURCE,body) if body.get('stamped') else SOURCE
                    with fitz.open(stream=pdf,filetype='pdf') as doc:
                        index=0 if body.get('mode')=='template' and not body.get('stamped') else int(body.get('page',1))-1
                        if not 0<=index<len(doc): raise ValueError('رقم معاينة غير صالح.')
                        page=doc[index]
                        pix=page.get_pixmap(matrix=fitz.Matrix(1.3,1.3),alpha=False)
                        self.reply(200,pix.tobytes('png'),'image/png'); return
                if self.path=='/export':
                    self.reply(200,stamp(SOURCE,body),'application/pdf'); return
                self.reply(404,{'error':'غير موجود'})
        except Exception as exc:
            self.reply(400,{'error':str(exc)})

if __name__=='__main__':
    server=ThreadingHTTPServer(('127.0.0.1',0),Handler)
    url=f'http://127.0.0.1:{server.server_port}'
    print('PageMarker:',url,flush=True)
    if '--no-browser' not in __import__('sys').argv: webbrowser.open(url)
    try: server.serve_forever()
    except KeyboardInterrupt: pass
    finally: server.server_close()
