import unittest,json
from generated_agent.operations import source_check,clean,execute

class BusinessBoundaries(unittest.TestCase):
    def test_no_default_authorization(self):
        self.assertEqual(source_check({'source':{'id':'x'}})['status'],'blocked')
    def test_source_requires_ownership(self):
        self.assertEqual(source_check({'source':{'authorized':True,'access':'user_export'}})['status'],'needs_input')
    def test_no_implicit_connector(self):
        self.assertEqual(source_check({'source':{'authorized':True,'access':'official_api'}})['status'],'blocked')
    def test_synthetic_not_customer(self):
        self.assertEqual(source_check({'source':{'authorized':True,'access':'owned_synthetic','id':'x','owner':'QA','scope':'sim','acquired_at':'2026-11-01','expected_fields':['date']},'data_origin':'customer'})['status'],'blocked')
    def test_no_copy_without_checked_input(self):
        self.assertEqual(execute({'action':'copy'})['status'],'needs_input')
    def test_candidate_not_confirmed(self):
        result=execute({'action':'copy','checked_metrics':{'daily':[{}],'source_sha256':'x'},'confirmed_anomalies':[{'human_confirmation':'pending'}]})
        self.assertEqual(result['status'],'blocked')
    def test_unknown_operation(self):
        with self.assertRaises(ValueError):execute({'action':'send_customer'})
if __name__=='__main__':unittest.main()
