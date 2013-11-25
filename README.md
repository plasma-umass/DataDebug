DataDebug
=========

CheckCell is an add-in for Microsoft Excel based on a technique called [data debugging](https://web.cs.umass.edu/publication/details.php?id=2283 "data debugging") for finding potential _input errors_ in spreadsheets.

Data debugging is an approach that combines data dependence analysis with statistical analysis to find and rank potential data errors. Since it is impossible to know _a priori_ whether data are erroneous or not, data debugging instead reveals data whose impact on the computation is unusual. Data debugging is particularly promising in the context of data-intensive programming environments that intertwine data with programs (in the form of queries or formulas).

CheckCell highlights values in shades proportional to the unusualness of their impact on the spreadsheet's computation, which includes charts and formulas. CheckCell is efficient; its algorithms are asymptotically optimal, and the current prototype runs in seconds for most spreadsheets.


Try CheckCell:
==============

You will need Microsoft Excel 2010.

=== Installing CheckCell ===
A release version of checkCell is available in the "Release" folder. Simply download the contents of the folder, run "setup.exe", and follow the instructions.
You are now ready to use CheckCell. When you open Excel, you will find CheckCell in the "Add-ins" tab in the ribbon.

=== Uninstalling CheckCell ===
In Excel, go to the Office menu, and click on "Options". Click on "Add-Ins" in the menu on the left. Find the "Manage" drop-down menu at the bottom, select "COM Add-ins", and click "Go...". Locate and select "DataDebug" in the list, and click on "Remove". Press "OK".