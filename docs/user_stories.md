From Pacific Union College:

At Pacific Union College, we conduct regular inventories of our physical materials. We do this by sending our student workers out to our shelves.  They record (usually by scanning the item’s barcode) what books are on the shelves.  Historically, in Sierra, scanning an item would connect with the ILS and enter the current date into an Inventory Date field. We could then run a list of items that had a recent inventoried date and compare them with the same range that didn’t have a recent inventoried date.  This would let us know what items are on the shelf.  

We want to be able to conduct a physical inventory of all our circulating items at least every other year. This ensures that misplaced items are found, that items thought missing that are on the shelf are correctly checked in, and that items that are not on the shelf can be searched for and either reshelved correctly or marked as missing. It also lets us check the condition of items on the shelf for any in need of repair or replacement. 

We need a way to scan an item on the shelf, record in the ILS that that item has been scanned, and then determine items that were not on the shelf based on some criteria.  Currently, FOLIO does not provide a way to automatically add a field when an item is scanned. While we can manually scan items and record barcodes using third-party apps, there is no way natively to have that scan registered in FOLIO for comparisons.  

We need to be able to 

1. Mark in some way (hopefully in bulk) items that have been found on the shelf and when they were last inventoried.  
2. Pull a list of all items that have been recently inventoried based on a date in a call number range and location.  
3. Pull a list of all items that are in a call number range and location that have the current status of "Available" (as in, they should be on the shelf).  
4. Compare the list of items to find items without a recent inventory date or be able to pull a list of items with an available status in a call number range and location without the most recent inventory date  
5. Do this for different locations and call number ranges throughout the course of the year.  
6. Have this process be able to be repeated every few years on an ongoing basis.  
7. Be able to pull these lists without having to know FOLIO’s API or JSON.

Nice to have

* Have the resulting lists be in correct call number order.  
* Be able to pull these reports ourselves instead of having to ask SCELC

When comparing results and then creating a list of missing items, we need the following data points:

* Item UUID   
* Item Barcode  
* Title (from Instance)  
* Call Number  
* Location  
* Item Status  
* Some type of item note field that indicates if an item is or is not on the shelf.    
  * Note: If this field can be filtered on easily within Excel, this may remove the need for pulling two lists for comparison; you could just pull a list of all available items in a call number range and location and filter by items without that date.

Western University of Health Sciences and Menlo

Both of these libraries did not use Sierra, but the process was similar in their previous ILS. Both stated that they used to be able to scan an item, which would connect it to their ILS to enter some kind of indication that the item had been inventoried.  Western and Menlo just finished an inventory before migration, so they will not be working on this again until next year at the earliest.  However, they have indicated that their needs are the same as listed above.

 In addition, Menlo would like to bring in the additional data point of the action holdings note in FOLIO for any item not found on the shelf. This is their note indicating that the item is marked for shared print retention.  All items missing and marked for retention need to either be replaced by the same edition or reported as missing to SCELC for potential transfer of the retention to another shared print library that holds that item.  SCELC can provide this data for the library if needed, but that would require the additional field on the FOLIO report of OCN (pulled from the instance).

