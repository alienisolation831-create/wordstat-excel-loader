import asyncio, os, secrets, hashlib, hmac, sqlite3, time, json, re, io, zipfile, threading, calendar, logging
from datetime import date, timedelta
from pathlib import Path
from contextlib import contextmanager, asynccontextmanager
from urllib.parse import urlsplit
import httpx
from openpyxl import load_workbook, Workbook
from starlette.applications import Starlette
from starlette.requests import Request
from starlette.responses import JSONResponse, HTMLResponse, Response, FileResponse
from starlette.routing import Route
from starlette.concurrency import run_in_threadpool

BASE=Path(__file__).resolve().parent
MAX_UPLOAD=5*1024*1024
MAX_ROWS=5000
MAX_COLS=50
API='https://searchapi.api.cloud.yandex.net/v2/wordstat/dynamics'
COOKIE='wordstat_session'
PERIODS={'day':'PERIOD_DAILY','week':'PERIOD_WEEKLY','month':'PERIOD_MONTHLY'}
class APIError(RuntimeError): pass

DEVICES={'DEVICE_ALL','DEVICE_DESKTOP','DEVICE_PHONE','DEVICE_TABLET'}

def integer(name, default, lo, hi):
    val=int(os.getenv(name,str(default)))
    if not lo<=val<=hi: raise ValueError(f'Invalid {name}')
    return val

class Config:
    def __init__(self):
        self.data=Path(os.getenv('DATA_DIR','./data')).resolve()
        self.demo=os.getenv('DEMO_MODE','true').lower()=='true'
        self.password=os.getenv('APP_PASSWORD','')
        self.secure=os.getenv('COOKIE_SECURE','true').lower()=='true'
        self.origin=os.getenv('APP_ORIGIN','').rstrip('/')
        if not self.origin: raise ValueError('APP_ORIGIN is required')
        parsed=urlsplit(self.origin)
        if parsed.scheme not in ('http','https') or not parsed.netloc or parsed.username or parsed.path or parsed.query or parsed.fragment: raise ValueError('Invalid APP_ORIGIN')
        if self.secure and parsed.scheme!='https': raise ValueError('Secure cookie requires HTTPS origin')
        if not self.secure and (parsed.scheme!='http' or parsed.hostname not in ('localhost','127.0.0.1','::1')): raise ValueError('Insecure cookie only allowed on loopback HTTP')
        if len(self.password)<16 or len(self.password)>256: raise ValueError('APP_PASSWORD must contain 16–256 characters')
        self.salt=secrets.token_bytes(16)
        self.password_hash=hashlib.scrypt(self.password.encode(),salt=self.salt,n=2**15,r=8,p=3,maxmem=64*1024*1024)
        self.password=None
        self.key=os.getenv('YANDEX_API_KEY',''); self.folder=os.getenv('YANDEX_FOLDER_ID','')
        if not self.demo and (not self.key or not self.folder): raise ValueError('YANDEX_API_KEY and YANDEX_FOLDER_ID required')
        self.hour=integer('MAX_REQUESTS_HOUR',100,1,10000)
        self.day=integer('MAX_REQUESTS_DAY',200,1,50000)
        self.run=integer('MAX_REQUESTS_RUN',500,1,MAX_ROWS)
        self.concurrent=integer('MAX_HTTP_CONCURRENT',8,2,32)
        self.cost=integer('API_MAX_COST_KOPECKS',0,0,1000000)
        self.budget=integer('API_DAILY_BUDGET_KOPECKS',0,0,100000000)
        if bool(self.cost)!=bool(self.budget) or (not self.demo and not self.cost): raise ValueError('Configure positive API_MAX_COST_KOPECKS and API_DAILY_BUDGET_KOPECKS for real mode')
        if self.cost:
            if self.budget<self.cost: raise ValueError('Daily budget below one request')
            self.day=min(self.day,self.budget//self.cost)
        self.max_date=os.getenv('WORDSTAT_MAX_DATE','2026-10-07')
        if self.max_date: date.fromisoformat(self.max_date)
        self.data.mkdir(mode=0o700,parents=True,exist_ok=True)
        self.db=self.data/'jobs.sqlite3'

@contextmanager
def db(cfg):
    c=sqlite3.connect(cfg.db,timeout=10); c.row_factory=sqlite3.Row
    try:
        c.execute('PRAGMA foreign_keys=ON'); yield c; c.commit()
    except: c.rollback(); raise
    finally: c.close()

def init(cfg):
    with db(cfg) as c:
        c.executescript('''PRAGMA journal_mode=WAL;
        CREATE TABLE IF NOT EXISTS sessions(token TEXT PRIMARY KEY, csrf TEXT NOT NULL, expires REAL NOT NULL);
        CREATE TABLE IF NOT EXISTS attempts(at REAL NOT NULL);
        CREATE TABLE IF NOT EXISTS peer_attempts(peer TEXT NOT NULL, at REAL NOT NULL);
        CREATE INDEX IF NOT EXISTS peer_attempts_at ON peer_attempts(at);
        CREATE INDEX IF NOT EXISTS peer_attempts_peer ON peer_attempts(peer,at);
        CREATE TABLE IF NOT EXISTS jobs(id TEXT PRIMARY KEY, status TEXT NOT NULL, created REAL NOT NULL, params TEXT NOT NULL, demo INTEGER NOT NULL);
        CREATE TABLE IF NOT EXISTS items(job TEXT NOT NULL REFERENCES jobs(id) ON DELETE CASCADE, n INTEGER NOT NULL, phrase TEXT NOT NULL, status TEXT NOT NULL DEFAULT 'pending', result TEXT, error TEXT, PRIMARY KEY(job,n));
        CREATE TABLE IF NOT EXISTS calls(at REAL NOT NULL);
        CREATE INDEX IF NOT EXISTS calls_at ON calls(at);
        ''')
        c.execute("DELETE FROM sessions")
        c.execute("UPDATE items SET status='error',error='Сервер перезапущен; результат неизвестен, повтор не выполнялся' WHERE status='running'")
        c.execute("UPDATE jobs SET status='queued' WHERE status='running'")
    os.chmod(cfg.db,0o600)

def digest(token): return hashlib.sha256(token.encode()).hexdigest()

def validate_xlsx(raw):
    if not raw or len(raw)>MAX_UPLOAD: raise ValueError('Размер XLSX: до 5 МБ')
    try:
        with zipfile.ZipFile(io.BytesIO(raw)) as z:
            infos=z.infolist(); names=[i.filename for i in infos]
            if len(infos)>500 or len(names)!=len(set(names)) or sum(i.file_size for i in infos)>40*1024*1024: raise ValueError('Слишком большой или неоднозначный XLSX')
            for i in infos:
                n=i.filename.lower()
                if i.flag_bits&1 or n.startswith('/') or '..' in n.split('/') or '\\' in n or any(x in n for x in ('vbaproject','externallinks','embeddings','activex')): raise ValueError('Неподдерживаемое содержимое XLSX')
            if '[Content_Types].xml' not in names or 'xl/workbook.xml' not in names: raise ValueError('Требуется XLSX')
        return load_workbook(io.BytesIO(raw),read_only=True,data_only=False,keep_links=False)
    except (zipfile.BadZipFile,KeyError,OSError,TypeError,ValueError) as e:
        raise ValueError('Некорректный или неподдерживаемый XLSX') from e
    except Exception as e: raise ValueError('Не удалось прочитать XLSX') from e

def preview(raw):
    w=validate_xlsx(raw)
    try: return [{'name':s.title,'rows':s.max_row,'columns':s.max_column} for s in w.worksheets][:50]
    finally: w.close()

def phrases(raw,sheet,col,header,limit):
    w=validate_xlsx(raw)
    try:
        if sheet not in w.sheetnames or not 1<=col<=MAX_COLS: raise ValueError('Выберите лист и столбец от 1 до 50')
        s=w[sheet]; out=[]; seen=set()
        if s.max_row and s.max_row>MAX_ROWS+1: raise ValueError('До 5000 строк запросов')
        if s.max_column and s.max_column>MAX_COLS: raise ValueError('До 50 столбцов')
        s.reset_dimensions()
        for idx,row in enumerate(s.iter_rows(min_row=2 if header else 1,min_col=col,max_col=col,max_row=MAX_ROWS+2)):
            cell=row[0]; val=cell.value
            if idx>=MAX_ROWS and val is not None: raise ValueError('До 5000 строк запросов')
            if val is None: continue
            if cell.data_type=='f' or not isinstance(val,str): raise ValueError('В столбце запросов разрешён только текст, без формул')
            val=val.strip()
            if not val: continue
            if len(val)>400 or any(ord(ch)<32 for ch in val): raise ValueError('Запрос до 400 символов, без управляющих символов')
            if val not in seen: seen.add(val);out.append(val)
            if len(out)>limit: raise ValueError(f'Лимит уникальных запросов за запуск: {limit}')
        if not out: raise ValueError('Нет запросов')
        return out
    finally: w.close()

def parameters(cfg,q):
    try: start=date.fromisoformat(q['start']); end=date.fromisoformat(q['end']); period=q['period']
    except (KeyError,ValueError): raise ValueError('Укажите корректные даты и период')
    if period not in PERIODS or start>end or end>date.today() or (end-start).days>730: raise ValueError('Некорректный период (не более 731 дня)')
    if not cfg.demo and cfg.max_date and end>date.fromisoformat(cfg.max_date): raise ValueError('Дата позже доступной статистики. Проверьте WORDSTAT_MAX_DATE')
    if period=='week' and (start.weekday()!=0 or end.weekday()!=6): raise ValueError('Недели: понедельник–воскресенье')
    if period=='month' and (start.day!=1 or end.day!=calendar.monthrange(end.year,end.month)[1]): raise ValueError('Месяцы: первый и последний день')
    regions=q.get('regions','').strip(); device=q.get('device','DEVICE_ALL')
    if regions and not re.fullmatch(r'\d{1,9}(,\d{1,9}){0,19}',regions): raise ValueError('Регионы: до 20 числовых ID через запятую')
    if device not in DEVICES: raise ValueError('Некорректное устройство')
    return {'start':str(start),'end':str(end),'period':period,'regions':regions.split(',') if regions else [],'device':device}

def reserve(cfg,c,now):
    c.execute('DELETE FROM calls WHERE at<=?',(now-86400,))
    n=c.execute('SELECT COUNT(*) FROM calls WHERE at>?',(now-3600,)).fetchone()[0]
    d=c.execute('SELECT COUNT(*) FROM calls').fetchone()[0]
    if n>=cfg.hour or d>=cfg.day: return False
    c.execute('INSERT INTO calls VALUES(?)',(now,)); return True

def api_result(cfg,phrase,p,demo):
    if demo:
        d=date.fromisoformat(p['start']); end=date.fromisoformat(p['end']); out=[]
        while d<=end:
            count=int(hashlib.sha256((phrase+str(d)).encode()).hexdigest()[:6],16)%10000
            out.append({'date':str(d),'count':count})
            d=(d+timedelta(days=calendar.monthrange(d.year,d.month)[1] if p['period']=='month' else 7 if p['period']=='week' else 1))
        return out
    payload={'phrase':phrase,'period':PERIODS[p['period']], 'fromDate':p['start']+'T00:00:00Z','toDate':p['end']+'T00:00:00Z','regions':p['regions'],'devices':[p['device']],'folderId':cfg.folder}
    # Fixed HTTPS host; no redirects, proxies or automatic retries. Never log body or credentials.
    with httpx.Client(timeout=httpx.Timeout(30,connect=10),follow_redirects=False,trust_env=False) as client:
        with client.stream('POST',API,json=payload,headers={'Authorization':'Api-Key '+cfg.key}) as response:
            if response.status_code!=200: raise APIError(f'API HTTP {response.status_code}')
            raw=bytearray(); deadline=time.monotonic()+45
            for chunk in response.iter_bytes():
                if time.monotonic()>deadline: raise APIError('API response deadline exceeded')
                raw.extend(chunk)
                if len(raw)>2*1024*1024: raise APIError('Ответ API слишком большой')
    try:
        data=json.loads(raw)['results']; out=[]; seen=set()
        if not isinstance(data,list) or len(data)>732: raise ValueError()
        for item in data:
            d=date.fromisoformat(item['date'][:10]); n=int(item['count'])
            if not p['start']<=str(d)<=p['end'] or d in seen or not 0<=n<=10**15: raise ValueError()
            seen.add(d);out.append({'date':str(d),'count':n})
        return out
    except (ValueError,TypeError,KeyError,IndexError,OverflowError): raise APIError('Неподдерживаемый формат ответа API') from None

def process_one(cfg):
    with db(cfg) as c:
        c.execute('BEGIN IMMEDIATE')
        j=c.execute("SELECT * FROM jobs WHERE status IN ('queued','running','waiting') ORDER BY created LIMIT 1").fetchone()
        if not j: return False
        item=c.execute("SELECT * FROM items WHERE job=? AND status='pending' ORDER BY n LIMIT 1",(j['id'],)).fetchone()
        if not item:
            c.execute("UPDATE jobs SET status='done' WHERE id=?",(j['id'],));return True
        if not j['demo'] and cfg.demo:
            c.execute("UPDATE jobs SET status='paused' WHERE id=?",(j['id'],));return False
        if not j['demo'] and not reserve(cfg,c,time.time()):
            c.execute("UPDATE jobs SET status='waiting' WHERE id=?",(j['id'],));return False
        c.execute("UPDATE jobs SET status='running' WHERE id=?",(j['id'],));c.execute("UPDATE items SET status='running' WHERE job=? AND n=?",(j['id'],item['n']))
    result=None; error=None
    try: result=api_result(cfg,item['phrase'],json.loads(j['params']),bool(j['demo']))
    except APIError as e: error=str(e) # only fixed local errors above
    except Exception: error='Ошибка связи или обработки ответа API'
    with db(cfg) as c:
        c.execute('BEGIN IMMEDIATE')
        c.execute('UPDATE items SET status=?,result=?,error=? WHERE job=? AND n=?',('error' if error else 'done',json.dumps(result) if result is not None else None,error,j['id'],item['n']))
        if error and not j['demo']:
            # Stop on any remote error, to avoid burning quota on failing calls.
            c.execute("UPDATE jobs SET status='paused' WHERE id=? AND status='running'",(j['id'],))
        elif not c.execute("SELECT 1 FROM items WHERE job=? AND status='pending'",(j['id'],)).fetchone():
            c.execute("UPDATE jobs SET status='done' WHERE id=? AND status='running'",(j['id'],))
    return True

def worker(cfg,stop):
    while not stop.is_set():
        try: worked=process_one(cfg)
        except Exception: worked=False; logging.error('Queue worker failed (details suppressed)')
        stop.wait(0.1 if worked else 2)

def create_app(cfg=None,run_worker=True):
    cfg=cfg or Config(); init(cfg)
    @asynccontextmanager
    async def lifespan(app):
        stop=threading.Event(); t=None
        if run_worker: t=threading.Thread(target=worker,args=(cfg,stop),daemon=True);t.start()
        yield
        stop.set()
        if t: await run_in_threadpool(t.join,35)
    login_lock=asyncio.Lock()
    def answer(data,status=200): return JSONResponse(data,status_code=status)
    def session(r):
        token=r.cookies.get(COOKIE,'')
        if not re.fullmatch(r'[A-Za-z0-9_-]{43}',token): return None
        with db(cfg) as c: return c.execute('SELECT * FROM sessions WHERE token=? AND expires>?',(digest(token),time.time())).fetchone()
    def auth(r,write=False):
        s=session(r)
        if not s: return answer({'error':'Требуется вход'},401)
        if write and (r.headers.get('origin')!=cfg.origin or not hmac.compare_digest(r.headers.get('x-csrf-token','').encode(),s['csrf'].encode())): return answer({'error':'CSRF: запрос отклонён'},403)
        return None
    async def body(r,limit):
        data=bytearray()
        try:
            async with asyncio.timeout(15):
                async for chunk in r.stream():
                    if len(data)+len(chunk)>limit: raise ValueError('Размер запроса превышен')
                    data.extend(chunk)
        except TimeoutError: raise ValueError('Превышено время загрузки') from None
        return bytes(data)
    async def home(r): return FileResponse(BASE/'static'/'index.html',media_type='text/html')
    async def asset(r):
        name=r.path_params['name']
        allowed={'app.js':'text/javascript','style.css':'text/css'}
        if name not in allowed: return answer({'error':'Not found'},404)
        return FileResponse(BASE/'static'/name,media_type=allowed[name])
    async def login(r):
        if r.headers.get('origin')!=cfg.origin: return answer({'error':'Недопустимый источник'},403)
        try:
            p=json.loads(await body(r,2048))['password'];
            if not isinstance(p,str) or len(p)>256: raise ValueError()
            password_bytes=p.encode('utf-8')
        except Exception: return answer({'error':'Некорректный запрос'},400)
        # The ASGI peer is used, not client-controlled forwarding headers.
        peer=digest(r.client.host if r.client else 'unknown')
        def attempt():
            with db(cfg) as c:
                c.execute('BEGIN IMMEDIATE'); now=time.time()
                c.execute('DELETE FROM peer_attempts WHERE at<=?',(now-900,))
                if c.execute('SELECT COUNT(*) FROM peer_attempts WHERE peer=?',(peer,)).fetchone()[0]>=5:
                    return False,429
                # Reserve before KDF to make concurrent attempts atomic.
                c.execute('INSERT INTO peer_attempts VALUES(?,?)',(peer,now))
            return hmac.compare_digest(hashlib.scrypt(password_bytes,salt=cfg.salt,n=2**15,r=8,p=3,maxmem=64*1024*1024),cfg.password_hash),401
        # At most one memory-intensive KDF at a time; no unbounded queue.
        if login_lock.locked(): return answer({'error':'Повторите попытку позже'},429)
        async with login_lock:
            ok,status=await run_in_threadpool(attempt)
            if not ok: return answer({'error':'Слишком много попыток' if status==429 else 'Неверный пароль'},status)
            with db(cfg) as c:
                c.execute('BEGIN IMMEDIATE'); c.execute('DELETE FROM peer_attempts WHERE peer=?',(peer,)); c.execute('DELETE FROM sessions')
                token=secrets.token_urlsafe(32); csrf=secrets.token_urlsafe(32)
                c.execute('INSERT INTO sessions VALUES(?,?,?)',(digest(token),csrf,time.time()+8*3600))
        res=answer({'csrf':csrf,'demo':cfg.demo,'limit':cfg.run,'max_date':cfg.max_date});res.set_cookie(COOKIE,token,httponly=True,secure=cfg.secure,samesite='strict',max_age=28800,path='/');return res
    async def info(r):
        s=session(r)
        if not s:return answer({'error':'Требуется вход'},401)
        return answer({'csrf':s['csrf'],'demo':cfg.demo,'limit':cfg.run,'max_date':cfg.max_date})
    async def logout(r):
        if (err:=auth(r,True)):return err
        with db(cfg) as c:c.execute('DELETE FROM sessions WHERE token=?',(digest(r.cookies[COOKIE]),))
        res=answer({'ok':True});res.delete_cookie(COOKIE,path='/',secure=cfg.secure,httponly=True,samesite='strict');return res
    async def inspect(r):
        if (err:=auth(r,True)):return err
        try:return answer({'sheets':await run_in_threadpool(preview,await body(r,MAX_UPLOAD))})
        except ValueError as e:return answer({'error':str(e)},400)
    async def start(r):
        if (err:=auth(r,True)):return err
        try:
            q=r.query_params;p=parameters(cfg,q)
            if q.get('confirm')!='yes':raise ValueError('Подтвердите отправку запросов')
            items=await run_in_threadpool(phrases,await body(r,MAX_UPLOAD),q.get('sheet',''),int(q.get('col','1')),q.get('header')=='yes',cfg.run)
        except ValueError as e:return answer({'error':str(e)},400)
        if (err:=auth(r,True)):return err
        with db(cfg) as c:
            c.execute('BEGIN IMMEDIATE')
            if c.execute("SELECT 1 FROM jobs WHERE status IN ('queued','running','waiting','paused')").fetchone():return answer({'error':'Сначала завершите или остановите текущую задачу'},409)
            if c.execute('SELECT COUNT(*) FROM jobs').fetchone()[0]>=50:return answer({'error':'Удалите старые задачи (лимит 50)'},409)
            if len(items)*((date.fromisoformat(p['end'])-date.fromisoformat(p['start'])).days+1)>100000:return answer({'error':'Слишком большой объём данных; сократите период или число запросов'},400)
            j=secrets.token_hex(16);c.execute('INSERT INTO jobs VALUES(?,?,?,?,?)',(j,'queued',time.time(),json.dumps(p),int(cfg.demo)))
            c.executemany('INSERT INTO items(job,n,phrase) VALUES(?,?,?)',[(j,n,x) for n,x in enumerate(items)])
        return answer({'id':j,'total':len(items)},201)
    async def jobs(r):
        if (err:=auth(r)):return err
        with db(cfg) as c:
            rows=c.execute('''SELECT j.id,j.status,j.created,j.demo,COUNT(i.n) total,SUM(CASE WHEN i.status IN ('done','error') THEN 1 ELSE 0 END) processed,SUM(CASE WHEN i.status='error' THEN 1 ELSE 0 END) errors FROM jobs j JOIN items i ON i.job=j.id GROUP BY j.id ORDER BY j.created DESC LIMIT 50''').fetchall()
        return answer({'jobs':[dict(x) for x in rows]})
    async def control(r):
        if (err:=auth(r,True)):return err
        j=r.path_params['job'];action=r.path_params['action']
        with db(cfg) as c:
            c.execute('BEGIN IMMEDIATE')
            if action=='stop':c.execute("UPDATE jobs SET status='stopped' WHERE id=? AND status IN ('queued','running','waiting','paused')",(j,))
            elif action=='resume':
                if c.execute("SELECT 1 FROM jobs WHERE id!=? AND status IN ('queued','running','waiting','paused')",(j,)).fetchone():return answer({'error':'Есть активная задача'},409)
                c.execute("UPDATE jobs SET status='queued' WHERE id=? AND status='paused'",(j,))
            elif action=='delete':c.execute("DELETE FROM jobs WHERE id=? AND status IN ('done','stopped') AND NOT EXISTS(SELECT 1 FROM items WHERE job=? AND status='running')",(j,j))
            else:return answer({'error':'Неизвестное действие'},400)
        return answer({'ok':True})
    def export(j):
        with db(cfg) as c:
            job=c.execute('SELECT * FROM jobs WHERE id=?',(j,)).fetchone()
            if not job:return None
            rows=c.execute('SELECT * FROM items WHERE job=? ORDER BY n',(j,)).fetchall()
        w=Workbook(write_only=True); s=w.create_sheet('Динамика');s.append(['Запрос','Дата','Частотность','Режим']);summary=w.create_sheet('Итоги');summary.append(['Запрос','Сумма доступных значений','Статус','Ошибка','Режим'])
        from openpyxl.cell import WriteOnlyCell
        def safe(ws,x):
            cell=WriteOnlyCell(ws,value=x);cell.data_type='s';return cell
        mode='DEMO — НЕ РЕАЛЬНЫЕ ДАННЫЕ' if job['demo'] else 'Wordstat'
        for row in rows:
            vals=json.loads(row['result']) if row['result'] else []
            for v in vals:s.append([safe(s,row['phrase']),safe(s,v['date']),v['count'],mode])
            summary.append([safe(summary,row['phrase']),sum(v['count'] for v in vals) if vals else None,row['status'],safe(summary,row['error'] or ''),mode])
        buf=io.BytesIO();w.save(buf);return buf.getvalue()
    async def download(r):
        if (err:=auth(r)):return err
        b=await run_in_threadpool(export,r.path_params['job'])
        if b is None:return answer({'error':'Не найдено'},404)
        return Response(b,media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={'Content-Disposition':'attachment; filename="wordstat-results.xlsx"'})
    app=Starlette(routes=[Route('/',home),Route('/api/login',login,methods=['POST']),Route('/api/info',info),Route('/api/logout',logout,methods=['POST']),Route('/api/preview',inspect,methods=['POST']),Route('/api/start',start,methods=['POST']),Route('/api/jobs',jobs),Route('/api/jobs/{job}/export',download),Route('/api/jobs/{job}/{action}',control,methods=['POST']),Route('/static/{name}',asset)],lifespan=lifespan,debug=False)
    class SecurityMiddleware:
        def __init__(self,app):
            self.app=app; self.inflight=0
        async def __call__(self,scope,receive,send):
            if scope['type']!='http': return await self.app(scope,receive,send)
            headers=[v for k,v in scope['headers'] if k.lower()==b'host']
            async def secured_send(msg):
                if msg['type']=='http.response.start':
                    additions={'Cache-Control':'no-store','X-Content-Type-Options':'nosniff','X-Frame-Options':'DENY','Referrer-Policy':'no-referrer','Content-Security-Policy':"default-src 'none'; script-src 'self'; style-src 'self'; connect-src 'self'; img-src 'self'; frame-ancestors 'none'; base-uri 'none'; form-action 'self'",'Permissions-Policy':'camera=(), microphone=(), geolocation=()'}
                    if cfg.secure: additions['Strict-Transport-Security']='max-age=31536000'
                    msg['headers']=[(k,v) for k,v in msg.get('headers',[]) if k.decode('latin1').lower() not in {h.lower() for h in additions}]+[(k.lower().encode(),v.encode()) for k,v in additions.items()]
                await send(msg)
            if len(headers)!=1 or headers[0].decode('latin1').lower()!=urlsplit(cfg.origin).netloc.lower():
                return await answer({'error':'Invalid host'},400)(scope,receive,secured_send)
            if self.inflight>=cfg.concurrent:
                return await answer({'error':'Server busy'},503)(scope,receive,secured_send)
            self.inflight+=1
            try: await self.app(scope,receive,secured_send)
            finally: self.inflight-=1
    app.add_middleware(SecurityMiddleware)
    app.state.cfg=cfg
    return app
