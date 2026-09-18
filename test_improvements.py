import hashlib
import io
import unittest
from test_offsets import ns, run_case, row
ns['hashlib'] = hashlib

class Improvements(unittest.TestCase):
    def test_same_business_is_not_duplicate(self):
        r,s=run_case([row()],[row(),row(300,30,'2026-09-03')])
        self.assertEqual(r['counts']['전산 중복 의심'],0)
        self.assertEqual(r['counts']['전산 확인 필요'],1)
    def test_exact_surplus_is_duplicate(self):
        r,s=run_case([row()],[row(),row()])
        self.assertEqual(r['counts']['정상 일치'],1)
        self.assertEqual(r['counts']['전산 중복 의심'],1)
    def test_repeated_erp_without_ht(self):
        r,s=run_case([],[row(),row()])
        self.assertEqual(r['counts']['전산 중복 의심'],2)
    def test_date_candidates(self):
        r,s=run_case([row()],[row(d='2026-10-02'),row(d='2026-10-03')])
        self.assertEqual(r['wrong_count'],0)
        self.assertEqual(r['counts']['홈택스 확인 필요'],1)
        self.assertEqual(r['counts']['전산 확인 필요'],2)
        self.assertEqual(len(s['5_확인필요_후보']),2)
    def test_multiple_ht_for_single_erp(self):
        r,s=run_case([row(d='2026-10-02'),row(d='2026-10-03')],[row()])
        self.assertEqual(r['wrong_count'],0)
        self.assertEqual(r['counts']['홈택스 확인 필요'],2)
        self.assertEqual(r['counts']['전산 확인 필요'],1)
        self.assertEqual(r['counts']['전산 누락'],0)
    def test_amount_candidates(self):
        r,s=run_case([row()],[row(200,20),row(300,30)])
        self.assertEqual(r['wrong_count'],0)
        self.assertEqual(r['counts']['홈택스 확인 필요'],1)
    def test_month_match_precedes_amount_candidates(self):
        ht=[row(),row(200,20,'2026-09-02')]
        for h in (ht,ht[::-1]):
            r,s=run_case(h,[row(100,10,'2026-09-02')])
            self.assertEqual(r['wrong_count'],0)
            self.assertEqual(r['counts']['정상 월 기준 일치'],1)
            self.assertEqual(r['counts']['홈택스 확인 필요'],0)
            self.assertEqual(r['counts']['전산 누락'],1)
    def test_unique_errors(self):
        r,s=run_case([row(),row(300,30,'2026-09-03',b='999')],
                     [row(d='2026-10-02'),row(400,40,'2026-09-03',b='999')])
        self.assertEqual(r['wrong_count'],2)
    def test_counts_reconcile(self):
        r,s=run_case([row(),row(),row(-100,-10),row(500,50,b='777'),row(b='')],
                     [row(),row(),row(b='')])
        c=r['counts']
        self.assertEqual(c['홈택스 원본(행)'],sum(c[k] for k in ['정상 일치','월·금액 확인','전산 누락','홈택스 확인 필요','홈택스 비교 제외','홈택스 상쇄(행)']))
        self.assertEqual(c['전산 원본(행)'],sum(c[k] for k in ['정상 일치','월·금액 확인','전산 확인 필요','전산 중복 의심','종이계산서 의심','전산 비교 제외','전산 상쇄(행)']))
        self.assertEqual(len(s['1_홈택스_대조결과']),5)
        self.assertEqual(len(s['2_종이세금계산서_의심']),1)
        self.assertEqual(len(s['6_대조결과_요약']),19)
    def test_empty_inputs(self):
        r,s=run_case([],[])
        self.assertTrue(all(v==0 for v in r['counts'].values()))
    def test_invalid_dates_not_exact(self):
        r,s=run_case([row(d='')],[row(d='')])
        self.assertEqual(r['counts']['정상 일치'],0)
    def test_unrelated_erp(self):
        r,s=run_case([row()],[row(),row(b='999')])
        self.assertEqual(r['counts']['종이계산서 의심'],1)

class UploadTests(unittest.TestCase):
    def upload(self,name,data):
        f=io.BytesIO(data)
        f.name=name
        return f
    def test_same_files_retain(self):
        state={}
        h=self.upload('a.xlsx',b'a'); e=self.upload('b.xlsx',b'b')
        ns['sync_uploaded_files'](state,'sales',h,e)
        state['sales']='result'
        ns['sync_uploaded_files'](state,'sales',h,e)
        self.assertEqual(state['sales'],'result')
    def test_same_name_different_content_clears_own_result(self):
        state={'purchases':'keep'}
        h=self.upload('a.xlsx',b'a'); e=self.upload('b.xlsx',b'b')
        ns['sync_uploaded_files'](state,'sales',h,e)
        state['sales']='old'
        ns['sync_uploaded_files'](state,'sales',self.upload('a.xlsx',b'new'),e)
        self.assertIsNone(state['sales'])
        self.assertEqual(state['purchases'],'keep')
    def test_removal(self):
        state={}
        h=self.upload('a.xlsx',b'a'); e=self.upload('b.xlsx',b'b')
        ns['sync_uploaded_files'](state,'sales',h,e)
        state['sales']='old'
        ns['sync_uploaded_files'](state,'sales',None,e)
        self.assertIsNone(state['sales'])

if __name__=='__main__':
    unittest.main(verbosity=2)
