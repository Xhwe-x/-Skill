import unittest
from check_bid import check


class CheckTests(unittest.TestCase):
    def test_current_facts(self):
        self.assertFalse(check('青禾设计服务，20个工作日。', {'project_name':'青禾设计服务','required_literals':['20个工作日']})['issues'][0]['code'] == 'REQUIRED_CUE')

    def test_legacy_is_project_specific(self):
        self.assertTrue(any(i['level']=='BLOCKER' for i in check('90日历天', {'forbidden_literals':['90日历天']})['issues']))
        self.assertFalse(any(i['level']=='BLOCKER' for i in check('90日历天', {})['issues']))

    def test_support_boundary(self):
        self.assertFalse(any(i['code']=='SCOPE' for i in check('我方提供设备采购技术咨询。', {'project_type':'design-consulting'})['issues']))

    def test_direct_boundary(self):
        self.assertTrue(any(i['code']=='SCOPE' for i in check('我方负责全部设备采购。', {'project_type':'design-consulting'})['issues']))

    def test_placeholders(self):
        self.assertEqual(check('【待填写】', {}, 'draft')['issues'][0]['level'],'REVIEW')
        self.assertEqual(check('【待填写】', {}, 'final')['issues'][0]['level'],'BLOCKER')

    def test_semantic_limit(self):
        self.assertTrue(check('已响应', {})['requires_semantic_review'])


if __name__=='__main__': unittest.main()
