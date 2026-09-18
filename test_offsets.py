import ast
import io
import hashlib
import unittest
from pathlib import Path
import pandas as pd

source = Path(__file__).with_name('opp.py').read_text(encoding='utf-8')
tree = ast.parse(source)
functions = ast.Module(body=[n for n in tree.body if isinstance(n, ast.FunctionDef)], type_ignores=[])
ns = {'pd': pd, 'io': io, 'hashlib': hashlib}
exec(compile(functions, 'opp.py', 'exec'), ns)

def run_case(ht, erp, sales=True):
    # tuples: supply, tax, date, business number
    def excel(rows, hometax):
        records = []
        for amount, tax, date, biz in rows:
            record = {'공급가액': amount, '세액': tax}
            if hometax:
                record.update({'공급받는자사업자등록번호': biz, '공급자사업자등록번호': biz,
                               '작성일자': date, '상호': '거래처', '상호.1': '거래처'})
            else:
                record.update({'사업자등록번호': biz, '발생일자': date})
            records.append(record)
        cols = (['공급가액','세액','공급받는자사업자등록번호','공급자사업자등록번호','작성일자','상호','상호.1']
                if hometax else ['공급가액','세액','사업자등록번호','발생일자'])
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
            pd.DataFrame(records, columns=cols).to_excel(writer, index=False, startrow=5 if hometax else 1)
        buf.seek(0)
        return buf
    result = ns['process_tax_invoices'](excel(ht, True), excel(erp, False), is_sales=sales)
    sheets = pd.read_excel(io.BytesIO(result['combined_result']), sheet_name=None)
    return result, sheets

def row(a=100, t=10, d='2026-09-01', b='123-45-67890'):
    return (a,t,d,b)

class Tests(unittest.TestCase):
    def test_user_sales_and_purchases(self):
        for sales in (True, False):
            for ht in ([row(),row(),row(-100,-10)], [row(-100,-10),row(),row()]):
                r,s = run_case(ht,[row()],sales)
                self.assertEqual(r['wrong_count'],0)
                self.assertEqual(r['offset_pair_count'],1)
                statuses = s['1_홈택스_대조결과']['전산대조결과'].tolist()
                self.assertEqual(statuses.count('정상(일치)'),1)
                self.assertEqual(sum(x.startswith('상쇄') for x in statuses),2)
                self.assertEqual(len(s['2_종이세금계산서_의심']),0)
                self.assertEqual(len(s['4_상쇄처리_내역']),2)
    def test_erp_offsets(self):
        r,s=run_case([row()],[row(),row(),row(-100,-10)])
        self.assertEqual(r['offset_pair_count'],1)
        self.assertEqual(len(s['2_종이세금계산서_의심']),0)
    def test_both_sides_full_records_preserve_exact_matches(self):
        r,s=run_case([row(),row(),row(-100,-10)],[row(),row(),row(-100,-10)])
        self.assertEqual(r['offset_pair_count'],0)
        self.assertTrue(s['1_홈택스_대조결과']['전산대조결과'].eq('정상(일치)').all())
    def test_incorrect_tax_not_offset(self):
        r,s=run_case([row(),row(),row(-100,10)],[row()])
        self.assertEqual(r['offset_pair_count'],0)
    def test_different_dates_not_offset(self):
        r,s=run_case([row(),row(),row(-100,-10,'2026-09-02')],[row()])
        self.assertEqual(r['offset_pair_count'],0)
    def test_different_business_not_offset(self):
        r,s=run_case([row(),row(),row(-100,-10,b='987-65-43210')],[row()])
        self.assertEqual(r['offset_pair_count'],0)
    def test_blank_dates_not_offset(self):
        r,s=run_case([row(d=''),row(-100,-10,d='')],[])
        self.assertEqual(r['offset_pair_count'],0)
    def test_zero_rows_not_offset(self):
        r,s=run_case([row(0,0),row(0,0)],[])
        self.assertEqual(r['offset_pair_count'],0)
    def test_zero_tax_offset(self):
        r,s=run_case([row(100,0),row(-100,0)],[])
        self.assertEqual(r['offset_pair_count'],1)
    def test_repeated_pairs_and_unmatched(self):
        r,s=run_case([row(),row(-100,-10),row(),row(-100,-10),row()],[])
        self.assertEqual(r['offset_pair_count'],2)
        self.assertEqual(s['1_홈택스_대조결과']['전산대조결과'].str.contains('누락').sum(),1)
    def test_ordinary_amount_error(self):
        r,s=run_case([row()],[row(200,20)])
        self.assertEqual(r['wrong_count'],1)
    def test_cross_month_requires_review(self):
        r,s=run_case([row()],[row(d='2026-10-02')])
        self.assertEqual(r['wrong_count'],1)

if __name__ == '__main__':
    unittest.main(verbosity=2)

