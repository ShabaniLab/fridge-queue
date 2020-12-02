This Google Apps Script is designed to be bound to a gsheet.

From a gsheet, open `Tools > Script Editor` and paste the `.gs` javascript and
`.json` manifest files.  You may need to run the script once from the script
editor to be prompted for permissions.  Then `Edit > Current project's
triggers` to add it as an installable trigger.  The script runs under a single
user's account and can only access user emails *in the same domain* from
accounts with which the gsheet is shared.

The trigger function is named `installableOnEdit()` and will fire when any edit
is made to the gsheet.  If the column headers are properly named (matching the
strings defined in the `.gs`), it will (depending on which column is edited),
1. autofill some fields, and
2. send an email to the next person in the queue.
