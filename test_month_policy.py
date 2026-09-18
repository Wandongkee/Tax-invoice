import unittest
from itertools import permutations
from test_offsets import ns,run_case,row
import test_screen

class MonthlyPolicy(unittest.TestCase):
    def assert_reconciled(self,r):
        c=r['counts']
        self.assertEqual(c['정상 일치'],c['정상 완전 일치']+c['정상 월 기준 일치'])
        self.assertEqual(c['홈택스 원본(행)'],sum(c[k] for k in ['정상 일치','월·금액 확인','전산 누락','홈택스 확인 필요','홈택스 비교 제외','홈택스 상쇄(행)']))
        self.assertEqual(c['전산 원본(행)'],sum(c[k] for k in ['정상 일치','월·금액 확인','전산 확인 필요','전산 중복 의심','종이계산서 의심','전산 비교 제외','전산 상쇄(행)']))
    def test_two_repeated_invoices_both_sales_purchases(self):
        for sales in (True,False):
            r,s=run_case([row(d='2026-08-19')]*2,[row(d='2026-08-01')]*2,sales)
            self.assertEqual(r['counts']['정상 월 기준 일치'],2)
            self.assertEqual(r['wrong_count'],0)
            self.assertEqual(len(s['2_종이세금계산서_의심']),0)
            self.assertEqual(len(s['5_확인필요_후보']),0)
            self.assertEqual(len(s['7_정상_월기준_내역']),2)
            self.assertEqual(s['1_홈택스_대조결과']['작성일자'].tolist(),['2026-08-19']*2)
            self.assertEqual(s['7_정상_월기준_내역']['전산_발생일자'].tolist(),['2026-08-01']*2)
            self.assert_reconciled(r)
    def test_exact_match_preserved_before_month_match(self):
        for erp in permutations([row(d='2026-08-01'),row(d='2026-08-19')]):
            r,s=run_case([row(d='2026-08-19')]*2,list(erp))
            self.assertEqual(r['counts']['정상 완전 일치'],1)
            self.assertEqual(r['counts']['정상 월 기준 일치'],1)
            self.assert_reconciled(r)
    def test_erp_surplus(self):
        r,s=run_case([row(d='2026-08-19')]*2,[row(d='2026-08-01'),row(d='2026-08-05'),row(d='2026-08-10')])
        self.assertEqual(r['counts']['정상 일치'],2)
        self.assertEqual(r['counts']['전산 중복 의심'],1)
        self.assert_reconciled(r)
    def test_hometax_surplus(self):
        r,s=run_case([row(d='2026-08-19')]*3,[row(d='2026-08-01')]*2)
        self.assertEqual(r['counts']['정상 일치'],2)
        self.assertEqual(r['counts']['전산 누락'],1)
        self.assert_reconciled(r)
    def test_cross_month_not_normal(self):
        r,s=run_case([row(d='2026-08-31')],[row(d='2026-09-01')])
        self.assertEqual(r['counts']['정상 일치'],0)
        self.assertEqual(r['counts']['귀속월 차이'],1)
        self.assertIn('확인 필요(귀속월 차이)',s['1_홈택스_대조결과']['전산대조결과'][0])
        self.assert_reconciled(r)
    def test_different_year_same_month_not_normal(self):
        r,s=run_case([row(d='2025-08-19')],[row(d='2026-08-19')])
        self.assertEqual(r['counts']['정상 일치'],0)
        self.assertEqual(r['counts']['귀속월 차이'],1)
    def test_invalid_date_not_normal(self):
        for d in ('','invalid'):
            r,s=run_case([row(d=d)],[row(d='2026-08-01')])
            self.assertEqual(r['counts']['정상 일치'],0)
            self.assertEqual(r['counts']['날짜 확인 필요'],1)
    def test_tax_difference_not_normal(self):
        r,s=run_case([row(d='2026-08-19')],[row(t=9,d='2026-08-01')])
        self.assertEqual(r['counts']['정상 일치'],0)
        self.assertEqual(r['counts']['금액·세액 오류'],1)
        self.assert_reconciled(r)
    def test_monthly_total_is_not_sufficient(self):
        r,s=run_case([row(100,10,'2026-08-19'),row(200,20,'2026-08-19')],
                     [row(150,15,'2026-08-01'),row(150,15,'2026-08-01')])
        self.assertEqual(r['counts']['정상 일치'],0)
        self.assertEqual(r['counts']['홈택스 확인 필요'],2)
        self.assert_reconciled(r)
    def test_different_business_not_normal(self):
        r,s=run_case([row(d='2026-08-19')],[row(d='2026-08-01',b='987-65-43210')])
        self.assertEqual(r['counts']['정상 일치'],0)
    def test_offsets_then_month_matching(self):
        r,s=run_case([row(d='2026-08-19')]*2+[row(-100,-10,'2026-08-19')],[row(d='2026-08-01')])
        self.assertEqual(r['counts']['상쇄 처리(쌍)'],1)
        self.assertEqual(r['counts']['정상 월 기준 일치'],1)
        self.assert_reconciled(r)
    def test_signed_invoices_not_absolute_amounts(self):
        r,s=run_case([row(-100,-10,'2026-08-19')],[row(d='2026-08-01')])
        self.assertEqual(r['counts']['정상 일치'],0)
    def test_policy_change_clears_previous_results(self):
        state={'sales':'old','sales_files':(None,None)}
        ns['sync_uploaded_files'](state,'sales',None,None)
        self.assertIsNone(state['sales'])
    def test_monthly_normal_ui(self):
        r,_=run_case([row(d='2026-08-19')]*2,[row(d='2026-08-01')]*2)
        fake=test_screen.ScreenTests().render(r)
        self.assertIn(('정상 일치',2),fake.metrics)
        self.assertTrue(any(t=='success' for t,_ in fake.messages))
        self.assertFalse(any(t=='warning' for t,_ in fake.messages))

if __name__=='__main__':
    unittest.main()
