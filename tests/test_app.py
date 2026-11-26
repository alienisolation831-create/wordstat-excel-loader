import io, os, tempfile, unittest, time, zipfile
from unittest.mock import patch
from openpyxl import Workbook, load_workbook
from starlette.testclient import TestClient
from app import Config,create_app,db,process_one,reserve,phrases,validate_xlsx,APIError,api_result,init

ORIGIN='http://localhost:8000'
PW='test-only-random-password-123'
def xlsx(values, formula=False):
    w=Workbook();s=w.active;s.title='Queries';s.append(['query'])
    for v in values:
        s.append([v])
        if v.startswith('=') and not formula:s.cell(s.max_row,1).data_type='s'
    b=io.BytesIO();w.save(b);return b.getvalue()

class Tests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.env=patch.dict(os.environ,{'DATA_DIR':self.tmp.name,'APP_PASSWORD':PW,'APP_ORIGIN':ORIGIN,'COOKIE_SECURE':'false','DEMO_MODE':'true','MAX_REQUESTS_HOUR':'2','MAX_REQUESTS_DAY':'3','MAX_REQUESTS_RUN':'10'},clear=True);self.env.start();self.cfg=Config();self.app=create_app(self.cfg,False);self.client=TestClient(self.app,base_url=ORIGIN);self.headers={'Origin':ORIGIN}
    def tearDown(self):self.client.close();self.env.stop();self.tmp.cleanup()
    def login(self):
        r=self.client.post('/api/login',json={'password':PW},headers=self.headers);self.assertEqual(r.status_code,200,r.text);self.headers['X-CSRF-Token']=r.json()['csrf'];return r
    def start(self,values=['one','two']):
        return self.client.post('/api/start',params={'sheet':'Queries','col':'1','header':'yes','start':'2026-09-01','end':'2026-09-30','period':'month','confirm':'yes'},content=xlsx(values),headers=self.headers)
    def test_unauthorized(self):
        self.assertEqual(self.client.get('/api/jobs').status_code,401);self.assertEqual(self.client.post('/api/start').status_code,401);self.assertEqual(self.client.get('/api/jobs/fake/export').status_code,401)
    def test_login_cookie_csrf_host(self):
        r=self.login();self.assertIn('HttpOnly',r.headers['set-cookie']);self.assertIn('SameSite=strict',r.headers['set-cookie']);self.assertEqual(self.client.post('/api/logout',headers={'Origin':ORIGIN}).status_code,403);self.assertEqual(self.client.post('/api/logout',headers={**self.headers,'Origin':'https://evil.example'}).status_code,403);self.assertEqual(self.client.get('/',headers={'Host':'evil.example'}).status_code,400);self.assertEqual(self.client.post('/api/logout',headers=self.headers).status_code,200);self.assertEqual(self.client.get('/api/jobs').status_code,401)
    def test_throttle(self):
        for _ in range(5):self.assertEqual(self.client.post('/api/login',json={'password':'wrong'},headers=self.headers).status_code,401)
        self.assertEqual(self.client.post('/api/login',json={'password':PW},headers=self.headers).status_code,429)
    def test_preview_and_formula_rejected(self):
        self.login();self.assertEqual(self.client.post('/api/preview',content=b'not xlsx',headers=self.headers).status_code,400);self.assertEqual(self.client.post('/api/preview',content=xlsx(['one']),headers=self.headers).json()['sheets'][0]['name'],'Queries')
        with self.assertRaises(ValueError):phrases(xlsx(['=1+1'],True),'Queries',1,True,10)
    def test_zip_unsafe(self):
        b=io.BytesIO();
        with zipfile.ZipFile(b,'w') as z:z.writestr('../evil','x')
        with self.assertRaises(ValueError):validate_xlsx(b.getvalue())
    def test_demo_end_to_end_and_safe_export(self):
        self.login();r=self.start(['=1+1','one','one']);self.assertEqual(r.status_code,201,r.text);self.assertEqual(r.json()['total'],2);j=r.json()['id'];self.assertEqual(self.start().status_code,409)
        with patch('app.httpx.Client',side_effect=AssertionError('network prohibited')):
            self.assertTrue(process_one(self.cfg));self.assertTrue(process_one(self.cfg))
        self.assertEqual(self.client.get('/api/jobs').json()['jobs'][0]['status'],'done')
        r=self.client.get(f'/api/jobs/{j}/export');self.assertEqual(r.status_code,200);w=load_workbook(io.BytesIO(r.content));self.assertEqual(w['Динамика']['A2'].data_type,'s');self.assertEqual(w['Динамика']['A2'].value,'=1+1');self.assertIn('DEMO',w['Динамика']['D2'].value);w.close()
        with db(self.cfg) as c:self.assertEqual(c.execute('SELECT COUNT(*) FROM calls').fetchone()[0],0)
    def test_stop(self):
        self.login();j=self.start().json()['id'];self.client.post(f'/api/jobs/{j}/stop',headers=self.headers);self.assertFalse(process_one(self.cfg));self.assertEqual(self.client.get('/api/jobs').json()['jobs'][0]['processed'],0)
    def test_quota(self):
        now=time.time()
        with db(self.cfg) as c:
            c.execute('BEGIN IMMEDIATE');self.assertTrue(reserve(self.cfg,c,now));self.assertTrue(reserve(self.cfg,c,now));self.assertFalse(reserve(self.cfg,c,now));self.assertTrue(reserve(self.cfg,c,now+3601));self.assertFalse(reserve(self.cfg,c,now+7201))
    def test_real_error_pauses_and_reserves(self):
        self.login();j=self.start().json()['id'];self.cfg.demo=False
        with db(self.cfg) as c:c.execute('UPDATE jobs SET demo=0 WHERE id=?',(j,))
        with patch('app.api_result',side_effect=APIError('API HTTP 429')):self.assertTrue(process_one(self.cfg));self.assertFalse(process_one(self.cfg))
        with db(self.cfg) as c:self.assertEqual(c.execute('SELECT COUNT(*) FROM calls').fetchone()[0],1)
        self.assertEqual(self.client.get('/api/jobs').json()['jobs'][0]['status'],'paused')
    def test_config_reject_insecure_remote(self):
        with patch.dict(os.environ,{'APP_ORIGIN':'http://example.com'}):
            with self.assertRaises(ValueError):Config()
    def test_restart_preserves_quota_and_marks_unknown(self):
        self.login();j=self.start().json()['id']
        with db(self.cfg) as c:
            c.execute("UPDATE jobs SET status='running' WHERE id=?",(j,));c.execute("UPDATE items SET status='running' WHERE job=? AND n=0",(j,));c.execute('INSERT INTO calls VALUES(?)',(time.time(),))
        init(self.cfg)
        with db(self.cfg) as c:
            self.assertEqual(c.execute('SELECT status FROM jobs WHERE id=?',(j,)).fetchone()[0],'queued');self.assertEqual(c.execute('SELECT status FROM items WHERE job=? AND n=0',(j,)).fetchone()[0],'error');self.assertEqual(c.execute('SELECT COUNT(*) FROM calls').fetchone()[0],1)
        self.assertEqual(self.client.get('/api/jobs').status_code,401);self.login()
        self.assertTrue(process_one(self.cfg));self.assertEqual(self.client.get('/api/jobs').json()['jobs'][0]['status'],'done')
    def test_secure_cookie_and_security_headers(self):
        with patch.dict(os.environ,{'APP_ORIGIN':'https://example.test','COOKIE_SECURE':'true'}):
            cfg=Config();app=create_app(cfg,False)
            with TestClient(app,base_url='https://example.test') as client:
                r=client.post('/api/login',json={'password':PW},headers={'Origin':'https://example.test'});self.assertEqual(r.status_code,200);self.assertIn('Secure',r.headers['set-cookie']);self.assertIn("frame-ancestors 'none'",r.headers['content-security-policy']);self.assertEqual(r.headers['cache-control'],'no-store');self.assertIn('strict-transport-security',r.headers)
    def test_request_size(self):
        self.login();r=self.client.post('/api/preview',content=b'x'*(5*1024*1024+1),headers=self.headers);self.assertEqual(r.status_code,400)
    def test_real_payload(self):
        import httpx
        self.cfg.key='not-a-real-key';self.cfg.folder='test-folder';calls=[]
        def handler(r):
            import json
            calls.append(r);self.assertEqual(str(r.url),'https://searchapi.api.cloud.yandex.net/v2/wordstat/dynamics');self.assertEqual(r.headers['authorization'],'Api-Key not-a-real-key');self.assertEqual(json.loads(r.content)['period'],'PERIOD_MONTHLY');return httpx.Response(200,json={'results':[{'date':'2026-09-01T00:00:00Z','count':'123'}]})
        real=httpx.Client
        with patch('app.httpx.Client',side_effect=lambda **kw:real(transport=httpx.MockTransport(handler),**kw)):
            self.assertEqual(api_result(self.cfg,'test',{'start':'2026-09-01','end':'2026-09-30','period':'month','device':'DEVICE_ALL','regions':[]},False),[{'date':'2026-09-01','count':123}]);self.assertEqual(len(calls),1)
    def test_password_change_revokes_sessions(self):
        self.login();token=self.client.cookies.get('wordstat_session')
        with patch.dict(os.environ,{'APP_PASSWORD':'new-unique-test-password-456'}):
            newapp=create_app(Config(),False)
            with TestClient(newapp,base_url=ORIGIN) as other:
                other.cookies.set('wordstat_session',token)
                self.assertEqual(other.get('/api/jobs').status_code,401)
    def test_invalid_unicode(self):
        r=self.client.post('/api/login',content=b'{"password":"\\ud800"}',headers=self.headers)
        self.assertEqual(r.status_code,400)
    def test_peer_isolation_and_no_header_spoofing(self):
        with TestClient(self.app,base_url=ORIGIN,client=('192.0.2.1',12345)) as attacker:
            for _ in range(5):self.assertEqual(attacker.post('/api/login',json={'password':'bad'},headers=self.headers).status_code,401)
            self.assertEqual(attacker.post('/api/login',json={'password':PW},headers={**self.headers,'X-Forwarded-For':'192.0.2.2'}).status_code,429)
        with TestClient(self.app,base_url=ORIGIN,client=('192.0.2.2',12345)) as owner:
            self.assertEqual(owner.post('/api/login',json={'password':PW},headers=self.headers).status_code,200)
    def test_budget_caps_persist(self):
        with patch.dict(os.environ,{'DEMO_MODE':'false','YANDEX_API_KEY':'test-placeholder','YANDEX_FOLDER_ID':'test-folder'}):
            with self.assertRaises(ValueError):Config()
            with patch.dict(os.environ,{'API_MAX_COST_KOPECKS':'25','API_DAILY_BUDGET_KOPECKS':'50'}):
                cfg=Config();self.assertEqual(cfg.day,2)
                with db(cfg) as c:
                    c.execute('BEGIN IMMEDIATE');now=time.time();self.assertTrue(reserve(cfg,c,now));self.assertTrue(reserve(cfg,c,now));self.assertFalse(reserve(cfg,c,now+3601))
    def test_static_allowlist_and_headers(self):
        self.assertEqual(self.client.get('/static/app.js').status_code,200)
        self.assertEqual(self.client.get('/static/not-allowed').status_code,404)
        r=self.client.get('/',headers={'Host':'invalid.example'});self.assertEqual(r.status_code,400);self.assertEqual(r.headers['cache-control'],'no-store')
    def test_concurrency_limit(self):
        import asyncio
        async def exercise():
            self.cfg.concurrent=2
            self.app.middleware_stack=None
            stack=self.app.build_middleware_stack()
            gate=asyncio.Event();entered=asyncio.Event();count=0
            middleware=stack.app
            original=middleware.app
            async def hold(scope,receive,send):
                nonlocal count
                count+=1
                if count==2:entered.set()
                await gate.wait()
                await original(scope,receive,send)
            middleware.app=hold
            import httpx
            async with httpx.AsyncClient(transport=httpx.ASGITransport(app=stack),base_url=ORIGIN) as c:
                tasks=[asyncio.create_task(c.get('/')) for _ in range(2)]
                await asyncio.wait_for(entered.wait(),2)
                try:self.assertEqual((await c.get('/')).status_code,503)
                finally:gate.set();await asyncio.gather(*tasks)
        asyncio.run(exercise())
if __name__=='__main__':unittest.main()
