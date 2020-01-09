Trading Card Inventory
======================

This Excel spreadsheet will help organize your card inventory and track sells (*This will not work in Google Sheets*).

## How do I use this?

* The first tab titled  *Inventory Master* is self-explanatory.

* The rest of the tabs *Jan-Dec* are where your sales go. 

* For this to work properly you **MUST** match the following cells between *Inventory Master* and *Jan-Dec* tabs: 
    * `inventory number`
    * `year`
    * `set`
    * `player`
    * `card`

* The `inventory number` is **Column B** on *Inventory Master* and **Column C** on *Jan-Dec* tabs.

## How does it work?

#### Purchase/Sales Channel

These columns use the Data Validation functionality to build the drop-down list. 
You can get to this by clicking 'Data' >> 'Data Validation'. In the validation criteria, list is chosen as the field type and the range uses the 'Sales Channel' table on *Jan* tab (W10:W13)

On *Jan-Dec* tabs, the Sales Channel uses its respective table for the list source.

#### Inventory - Sold Price & Fee

On *Inventory Master* tab, the Sold Price column uses a function that scans *Jan-Dec* tabs looking for rows that match on all the following:

`inventory number` (**Column B** on *Inventory*, **Column C** on *Jan-Dec*)

`year` (**Column D** on *Inventory*, **Column E** on *Jan-Dec*) 

`set` (**Column E** on *Inventory*, **Column F** on *Jan-Dec*)

`player` (**Column F** on *Inventory*, **Column G** on *Jan-Dec*)

`card` (**Column G** on *Inventory*, **Column H** on *Jan-Dec*) 

The text in these fields must match exactly!

This is the function used for Sold Price (Fee is the same, but replace $K$5:$K$1000 with $P$5:$P$1000):

     =IFERROR(
       SUMPRODUCT(
        SUMIFS(
          INDIRECT("'" & Months & "'!$K$5:$K$1000"),
          INDIRECT("'" & Months & "'!$C$5:$C$1000"),$B7,
          INDIRECT("'" & Months & "'!$E$5:$E$1000"),$D7,
          INDIRECT("'" & Months & "'!$F$5:$F$1000"),$E7,
          INDIRECT("'" & Months & "'!$G$5:$G$1000"),$F7,
          INDIRECT("'" & Months & "'!$H$5:$H$1000"),$G7
        )
       ),
      )

**BE WARNED:** Adding or removing columns will likely break all functionality across all tabs! Also, do not rename the tabs!

#### Jan-Dec Card Cost

This method is similar to the Sold Price & Fee columns described above, except it multiplies it by card count.  Here's the function:

     =IFERROR(
       SUMPRODUCT(
        SUMIFS(
          INDIRECT("'" & Months & "'!$L$7:$L$1000"),
          INDIRECT("'" & Months & "'!$B$7:$B$1000"),$C5,
          INDIRECT("'" & Months & "'!$D$7:$D$1000"),$E5,
          INDIRECT("'" & Months & "'!$E$7:$E$1000"),$F5,
          INDIRECT("'" & Months & "'!$F$7:$F$1000"),$G5,
          INDIRECT("'" & Months & "'!$G$7:$G$1000"),$H5
        )*J5
       ),
      )

#### Other Functions

The rest of the cells performing calculations aren't doing anything special, just click around to see what's happening.