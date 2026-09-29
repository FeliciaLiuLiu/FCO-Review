"""Schema/key/preflight tests; no generated input datasets."""
from pathlib import Path
import tempfile
import unittest
from fill_addresses import ADDRESS_MAPPING, checked_headers, normalize_key, combine_batches
from mapping import HEADERS

class AddressChecks(unittest.TestCase):
    def test_mapping(self):
        self.assertEqual(len(ADDRESS_MAPPING), 15)
        self.assertEqual(len(set(target for target, _ in ADDRESS_MAPPING)), 15)
        self.assertEqual({target for target, _ in ADDRESS_MAPPING},
                         {header for header in HEADERS if 'Address Line' in header})
        for party, prefix in [('Ultimate Debtor', 'ult_deb_address_line'),
                              ('Debtor', 'orig_addr_line'),
                              ('Creditor', 'secondary_bene_addr_line'),
                              ('Ultimate Creditor', 'ult_cred_address_line'),
                              ('Initiating Party', 'ip_address_line')]:
            for line in (1, 2, 3):
                self.assertEqual(dict(ADDRESS_MAPPING)[f'{party} Address Line {line}'], f'{prefix}{line}')

    def test_text_keys(self):
        self.assertEqual(normalize_key(' 000123 '), '000123')
        self.assertEqual(normalize_key(None), '')
        self.assertNotEqual(normalize_key('Ab'), normalize_key('ab'))
        with self.assertRaises(ValueError):
            normalize_key(123)

    def test_headers(self):
        checked_headers(['a', 'b'], ['b'], 'test')
        with self.assertRaises(ValueError):
            checked_headers(['a', 'a'], ['a'], 'test')
        with self.assertRaises(ValueError):
            checked_headers(['a'], ['b'], 'test')

    def test_missing_files_do_not_create_combined(self):
        with tempfile.TemporaryDirectory() as folder:
            with self.assertRaisesRegex(ValueError, 'Batch8.csv'):
                combine_batches(Path(folder), 'utf-8-sig')
            self.assertEqual(list(Path(folder).iterdir()), [])

if __name__ == '__main__':
    unittest.main()
