# CopyFromFilter-Excel-VBA
Splits one big MS Excel worksheet into many smaller parts using one identifying field.

#### Workflow
Use this macro, if you have one big Excel worksheet with information from many account manager/supplier/customer/etc. (the defining item) and want to split this information regurlaly into smaller parts.
You just need one column with the defining item and every unique item will recieve its own Excel workbook with just its information.

#### Install
Just import `FilterSplit.bas`, `Functions.bas` and `Settings.bas` into your desired workbook and you are set to go.

#### Settings
`Settings.bas` consists of a retriever for variables, you can use to define:
* `Field`:      the column in which your identifier is (numerical).
* `Table`:      the worksheet in which your information is (string).
* `Filter`:     the name of the column of your identifier (string).
* `TrgtRange`:  the range of your new workbook into which the information should be copied (range as string).

#### Flaws, Prospects and Roadmap
This macro is an early version of a general macro based upon a highly specialized macro I wrote at work.
* It does not retrieve the names of the columns, which is the biggest flaw to fix in future releases.
* Furthermore it does not format the new worksheet and just copies bare information (which is sufficient for most, but not all uses).

#### Credits
* `Functions.IsInArray()` by [JimmyPena](http://stackoverflow.com/users/190829/jimmypena "stackoverflow.com") via [StackOverflow](http://stackoverflow.com/a/10952705 "stackoverflow.com")
* the beautiful solution converting a range to recordset by [kulshresthazone](http://en.gravatar.com/kulshresthazone "gravatar.com") via [Useful Gyann](https://usefulgyaan.wordpress.com/2013/07/11/vba-trick-of-the-week-range-to-recordset-without-making-connection/ "VBA Trick of the Week")
