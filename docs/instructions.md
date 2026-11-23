# Instructions

Fill in values on the config tab.

Under FOLIO Inventory click **Create inventory**.

On the first execution, you will need to give the apps script permission to run.

Once permissions are set, try **Create inventory** again.

Enter the FOLIO username's password in the provided box.

A new sheet is created with instance data.

To do:
- Control return fields on the config tab?
- Sort by call number order
- CQL query parameters in the API request
- Limit on the number of returned rows (on a large FOLIO instance unlimited rows will timeout)
- Better handling of deeply nested FOLIO JSON structures (see the numerous subjects columns)
- Auth tenant and data tenant differ for SCELC but not members (SCELC Office instance accounts auth against the central tenant)
