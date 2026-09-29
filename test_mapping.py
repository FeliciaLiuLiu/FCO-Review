"""Scalar logic and schema checks only; no synthetic CSV or output records."""
import unittest
from convert import names_need_highlight
from mapping import MAPPING, HEADERS

class MappingChecks(unittest.TestCase):
    def test_names(self):
        self.assertFalse(names_need_highlight('Billy Kim', 'BILLYKIM', None))
        self.assertFalse(names_need_highlight('Billy Kim', None, 'Payment for billy kim Ltd'))
        self.assertFalse(names_need_highlight('Billy Kim', 'Billy\tKim', 'Unrelated'))
        self.assertTrue(names_need_highlight('Billy Kim', 'Billy', 'Kim'))
        self.assertTrue(names_need_highlight('Billy Kim', '', None))
        self.assertTrue(names_need_highlight('Billy Kim', 'Other', 'Unrelated'))
        self.assertTrue(names_need_highlight('Billy Kim', 'Billy-Kim', None))

    def test_schema(self):
        self.assertEqual(len(HEADERS), 45)
        self.assertEqual(len(set(HEADERS)), 45)
        mapping = dict(MAPPING)
        for label, prefix in [('Ultimate Debtor', 'ult_orig'), ('Ultimate Creditor', 'ult_bene'),
                              ('Debtor', 'orig'), ('Creditor', 'bene')]:
            for suffix, source in [('Account ID', 'raw_account_num'), ('Name', 'clean_name'),
                                   ('Address Line 1', 'raw_addr_id'), ('City', 'city'),
                                   ('Country', 'cntry_code')]:
                self.assertEqual(mapping[label + ' ' + suffix], prefix + '_' + source)
        self.assertEqual(sum(source is None for _, source in MAPPING), 10)

if __name__ == '__main__':
    unittest.main()
