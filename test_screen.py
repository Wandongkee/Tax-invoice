import unittest
from test_offsets import ns, run_case, row

class FakeStreamlit:
    def __init__(self,result):
        self.session_state={'sales':result,'sales_files':(None,None)}
        self.messages=[]
        self.metrics=[]
        self.downloads=[]
    def __enter__(self): return self
    def __exit__(self,*args): pass
    def columns(self,n): return [self]*n
    def subheader(self,*args): pass
    def file_uploader(self,*args,**kwargs): return None
    def metric(self,label,value): self.metrics.append((label,value))
    def caption(self,*args): pass
    def info(self,text): self.messages.append(('info',text))
    def warning(self,text): self.messages.append(('warning',text))
    def success(self,text): self.messages.append(('success',text))
    def download_button(self,*args,**kwargs): self.downloads.append(kwargs)

class ScreenTests(unittest.TestCase):
    def render(self,result):
        fake=FakeStreamlit(result)
        ns['st']=fake
        ns['render_invoice_section']('매출','홈택스','전산','시작','sales',True,'sales')
        return fake
    def test_missing_is_warning_even_without_amount_errors(self):
        r,_=run_case([row()],[])
        fake=self.render(r)
        self.assertTrue(any(t=='warning' for t,_ in fake.messages))
        self.assertFalse(any(t=='success' for t,_ in fake.messages))
        self.assertIn(('전산 누락',1),fake.metrics)
        self.assertEqual(len(fake.downloads),1)
    def test_offset_summary(self):
        r,_=run_case([row(),row(),row(-100,-10)],[row()])
        fake=self.render(r)
        self.assertIn(('정상 일치',1),fake.metrics)
        self.assertIn(('상쇄 처리(쌍)',1),fake.metrics)
        self.assertTrue(any(t=='success' for t,_ in fake.messages))
    def test_stale_result_disappears_on_removal(self):
        r,_=run_case([row()],[row()])
        fake=FakeStreamlit(r)
        fake.session_state['sales_files']=('old','old')
        ns['st']=fake
        ns['render_invoice_section']('매출','홈택스','전산','시작','sales',True,'sales')
        self.assertIsNone(fake.session_state['sales'])
        self.assertEqual(fake.downloads,[])
    def test_empty_is_not_success(self):
        r,_=run_case([],[])
        fake=self.render(r)
        self.assertFalse(any(t=='success' for t,_ in fake.messages))
        self.assertTrue(any('비교할 자료가 없습니다' in msg for _,msg in fake.messages))
