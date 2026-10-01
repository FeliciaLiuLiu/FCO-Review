# Fixed Focal Client Accounts

These are fictional account identifiers from client_registry.py, not real company accounts. This reference describes the new generator configuration; existing CSV files have not been regenerated.

Both dataset generators share this mapping. Each account belongs exclusively to one client regardless of whether the client appears as originator, beneficiary, or a related party. Import accounts as text to preserve leading zeros. All individual, business, and bank clients also have fixed account sets.

| Client | Category | Account count | All account numbers |
|---|---|---:|---|
| PAYPAL | fintech | 7 | `01000000010001`, `01000000010002`, `01000000010003`, `01000000010004`, `01000000010005`, `01000000010006`, `01000000010007` |
| VENMO | fintech | 3 | `01000000020001`, `01000000020002`, `01000000020003` |
| PAYONEER | fintech | 5 | `01000000030001`, `01000000030002`, `01000000030003`, `01000000030004`, `01000000030005` |
| TTT MoneyCorp | fintech | 7 | `01000000040001`, `01000000040002`, `01000000040003`, `01000000040004`, `01000000040005`, `01000000040006`, `01000000040007` |
| WISE | fintech | 5 | `01000000050001`, `01000000050002`, `01000000050003`, `01000000050004`, `01000000050005` |
| STRIPE | fintech | 3 | `01000000060001`, `01000000060002`, `01000000060003` |
| AURORA GLOBAL TRADING LTD | business | 3 | `01000000070001`, `01000000070002`, `01000000070003` |
| CEDAR LOGISTICS GROUP | business | 5 | `01000000080001`, `01000000080002`, `01000000080003`, `01000000080004`, `01000000080005` |
| QUARTZ IMPORT EXPORT LTD | business | 7 | `01000000090001`, `01000000090002`, `01000000090003`, `01000000090004`, `01000000090005`, `01000000090006`, `01000000090007` |

## Field correspondence

- Originator: orig_raw_account_num belongs to orig_raw_name and orig_clean_name.
- Beneficiary: bene_raw_account_num belongs to bene_raw_name and bene_clean_name. The existing schema has no bene_client_name column.
- In the wide dataset, std_account_num matches raw_account_num, and raw_name, clean_name, master_name, and entity_id describe the same owner.
- Populated ultimate-party groups reuse their corresponding core leg and selected account. The initiating party reuses a focal core leg when available.
- Accounts, entity IDs, and client names are protected from noise and partial blank injection. Optional groups can be wholly absent.
- Fintech rows retain the 25% rule. The listed business focal clients are selected through the ordinary business pool.
- Edit FOCAL_CLIENTS to change per-client account counts. Keep existing entity numbers stable to preserve ownership across regenerations.
