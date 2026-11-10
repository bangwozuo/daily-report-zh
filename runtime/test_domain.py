import unittest,csv,io
from generated_agent.domain import clean_csv,detect,ratio,change

HEADER='date,revenue,orders,visitors,refunds,store_id,currency,promotion,opening_date,data_origin,phone,手机号\n'
def source(day='2026-10-31',revenue='100',orders='40',visitors='200',refunds='7',promotion='false'):
    return HEADER+f'{day},{revenue},{orders},{visitors},{refunds},SIM-STORE-001,CNY,{promotion},2024-01-15,SYNTHETIC,13800000000,13900000000\n'
class Rules(unittest.TestCase):
    def test_utf8_bom_preserves_date_column(self):
        result=clean_csv('\ufeff'+source());self.assertEqual(result['status'],'draft');self.assertEqual(result['rows'][0]['date'],'2026-10-31')
    def test_privacy_deleted_before_rows_and_evidence(self):
        result=clean_csv(source());self.assertEqual(set(result['removed_fields']),{'phone','手机号'})
        self.assertNotIn('phone',result['rows'][0]);self.assertNotIn('手机号',result['rows'][0])
    def test_zero_rates_remain_unknown(self):
        result=clean_csv(source(revenue='0',orders='0',visitors='0',refunds='0'))['rows'][0]
        for key in ['average_order_value','conversion_rate','refund_rate']:self.assertIsNone(result[key])
    def test_negative_not_silently_corrected(self):
        result=clean_csv(source(orders='-1'));self.assertEqual(result['rows'],[]);self.assertTrue(result['quality'])
    def test_duplicate_preserves_quality_does_not_double(self):
        text=source();result=clean_csv(text+text.splitlines()[1]+'\n');self.assertEqual(len(result['rows']),1);self.assertIn('duplicate_day_rows',result['quality'][0]['reason'])
    def test_missing_calendar_not_filled(self):
        first=source(day='2026-10-28');second=source(day='2026-10-31');result=clean_csv(first+second.splitlines()[1]+'\n')
        missing=next(q for q in result['quality'] if q['reason']=='missing_calendar_date');self.assertEqual(missing['days'],2);self.assertFalse(missing['filled_with_zero'])
    def test_low_order_ratio_is_suppressed(self):
        rows=clean_csv(source(orders='12'))['rows'];result=detect(rows,'SIM-STORE-001','2026-10-31','2026-10-31')
        self.assertTrue(any('不足20单只报比率值' in a['suppression_reasons'] for a in result['alerts']))
    def test_promotion_protection(self):
        rows=clean_csv(source(promotion='true'))['rows'];result=detect(rows,'SIM-STORE-001','2026-10-31','2026-10-31')
        self.assertTrue(result['alerts']);self.assertTrue(all(a['state']=='suppressed' for a in result['alerts']))
    def test_yoy_missing_is_unknown(self):
        result=detect(clean_csv(source())['rows'],'SIM-STORE-001','2026-10-31','2026-10-31');self.assertIsNone(result['daily'][0]['revenue_yoy'])
    def test_empty_core_headers_stop(self):self.assertIn(clean_csv('date,orders\n2026-10-31,4\n')['status'],{'blocked','needs_input'})
if __name__=='__main__':unittest.main()
