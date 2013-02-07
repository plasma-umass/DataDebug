DataDebug
=========

CheckCell is an add-in for Microsoft Excel based on a technique called [data debugging](https://web.cs.umass.edu/publication/details.php?id=2283 "data debugging") to find potential errors in spreadsheets.

Data debugging is an approach that combines data dependence analysis with statistical analysis to find and rank potential data errors. Since it is impossible to know _a priori_ whether data are erroneous or not, data debugging instead reveals data whose impact on the computation is unusual. Data debugging is particularly promising in the context of data-intensive programming environments that intertwine data with programs (in the form of queries or formulas).

CheckCell highlights values in shades proportional to the unusualness of their impact on the spreadsheet's computation, which includes charts and formulas. CheckCell is efficient; its algorithms are asymptotically optimal, and the current prototype runs in seconds for most spreadsheets we examine. We perform a case study by employing workers via a crowdsourcing platform, and show that CheckCell is effective at finding actual data entry errors.