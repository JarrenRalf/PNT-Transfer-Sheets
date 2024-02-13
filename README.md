# Transfer-Sheets
At Pacific Net & Twine LTD, there is a primary store and shipping center located in Richmond, plus 2 branch stores located in Parksville and Prince Rupert. Products flow between these three locations on a daily basis. When the branch stores need items to either fill their shelves or fulfill a customer order, they use the _Item Search_ tab to find the desired item/s and add them to the _Order_ page. When the shipping team in Richmond picks the items, they fill in a shipped quantity and change the shipment status of that item, which is now moved to the _Shipped_ page, to inform the particular branch store of the date which the items will ship and which method they ship by. When the receiving team at the branch store physically recieve the item, they again change the shipment status to received, which results in the item moving to the _Received_ page. Once an item makes it to the Received tab, this transaction is ready to be recorded in the inventory system.
In addition to handeling the transfer of products between the stores, the _Transfer Sheets_ are used to count inventory, in 3 ways primary. 
1. The _InfoCounts_, which identify which items have a negative inventory and therefore should be verified.
2. The _Manual Counts_, which are items that have been specifically searched for on the _Item Search_ page, that have been added by a user to prompt employess to do a spot check of those items, regardless of recorded inventory values.
3. During the process of ordering, for example, someone from a branch store may order 10 coils of rope, but their inventory value for that SKU indicates that they have 33 units but that number does not match the physical stock, so an employee enters in their true value, which for example could be 1, hence why they are ordering the 10 coils.

Those are the primary functions of these sets of spreadsheets, however, with more than 5000 lines of original Google Apps Script (Javascript) code written, this spreadsheet is feature rich, especially for the primary inventory control manager.
Some of executive features include the following:
- Receiving every item that a is shipped by a particular carrier with 1 click.\
- Handeling a UPC database that allows users to scan barcodes in order to add items to _Order_, _Shipped_, or _Manual Counts_ pages
- Automatic formatting of each displayed page, because with multiple users, various aspects of the sheets change regularly.
- 
