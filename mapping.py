"""Authoritative output order and source column mapping."""
MAPPING = [
    ('Fintech Customer', '$filename'),
    ('Transaction Date', 'value_date'),
    ('Transaction ID (UETR or Other Unique ID)', 'trxn_id'),
    ('Transaction Amount', 'amount'),
]

def party(label, prefix):
    return [
        (label + ' Account ID', prefix + '_raw_account_num'),
        (label + ' Name', prefix + '_clean_name'),
        (label + ' Address Line 1', prefix + '_raw_addr_id'),
        (label + ' Address Line 2', None),
        (label + ' Address Line 3', None),
        (label + ' City', prefix + '_city'),
        (label + ' Country', prefix + '_cntry_code'),
    ]

MAPPING += party('Ultimate Debtor', 'ult_orig')
MAPPING += party('Debtor', 'orig')
for label, prefix in [('Agent Debtor', 'obk'), ('Agent Creditor', 'bbk')]:
    MAPPING += [(label + ' ID', prefix + '_raw_account_num'),
                (label + ' Name', prefix + '_clean_name'),
                (label + ' Country', prefix + '_cntry_code')]
MAPPING += party('Creditor', 'bene')
MAPPING += party('Ultimate Creditor', 'ult_bene')
MAPPING += party('Initiating Party', 'ip')
HEADERS = [label for label, _ in MAPPING]
