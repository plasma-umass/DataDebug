DataDebug
=========

CheckCell is an add-in for Microsoft Excel based on a technique called [data debugging](https://web.cs.umass.edu/publication/details.php?id=2283 "data debugging") for finding potential _input errors_ in spreadsheets.

Data debugging is an approach that combines data dependence analysis with statistical analysis to find and rank potential data errors. Since it is impossible to know _a priori_ whether data are erroneous or not, data debugging instead reveals data whose impact on the computation is unusual. Data debugging is particularly promising in the context of data-intensive programming environments that intertwine data with programs (in the form of queries or formulas).

CheckCell highlights suspected errors in red one cell at a time, so that they can be inspected by the user and corrected, or marked as 'OK'. CheckCell is efficient; its algorithms are asymptotically optimal, and the current prototype runs in seconds for most spreadsheets.


Try CheckCell:
==============

You will need Microsoft Excel 2010.

Installing CheckCell
--------------------
A release version of checkCell is available in the "Release.zip" archive. Simply extract the contents of the archive, run "setup.exe", and follow the instructions.
You are now ready to use CheckCell. You will find CheckCell in Excel, under the "Add-Ins" tab in the ribbon.

Uninstalling CheckCell
----------------------
In Excel, go to the Office menu, and click on "Options". Click on "Add-Ins" in the menu on the left. Find the "Manage" drop-down menu at the bottom, select "COM Add-ins", and click "Go...". Locate and select "DataDebug" in the list, and click on "Remove". Press "OK".

Using CheckCell
---------------
To use CheckCell, open the spreadsheet you would like to audit, and click CheckCell's "Analyze" button. (It is located in the "Add-Ins" tab in the ribbon.) CheckCell will perform its analysis, and if any potential errors are found, they will be highlighted one at a time in decreasing order of importance. For each highlighted cell, you will have to decide if it is actually an error. If so, click on the "Fix Error" button and enter the correct value in the box that comes up. Otherwise, click the "Mark as OK" button. After each correction, CheckCell will re-run its analysis using the corrected value.

CheckCell's sensitivity level is adjustable. By default it is set to display the top 5% most unusual values, but you may change it by entering a different value in the box labeled "% Most Unusual to Show".