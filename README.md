# Miniso Forecasting Tool

This tool basically pulls the Excel data (.xlsx or .xlsm), applies interpolation (Missing Data Handling) and/or exterpolation (Forecasting using Theta function), and prints the results to an Excel file. 

The Excel data should be formatted like this:

First column should have ID's (Unique identifiers of the data)
First row should have time indexes (Day, month, week, etc.)
The sale data should be filled like a matrix in this format.


This software is built and published by Görkem Kısa. This software is a part of my team's Senior Design Project, anyone who has access to this page can inspect and use the code freely (if you want to use this, please reference the contributors of the packages and libraries that I've used).

Our team:
- Görkem Kısa
- Aslı Çahantimur
- Ali Şamil Adıgüzel
- Ecem Akış

I've built this software using C# and R. Used libraries in C#:

- R.NET

Used packages in R:

- forecast
- openxlsx
- imputeTS
- mice
- forecastHybrid

and all of the dependent packages of the above list.


References:
Hyndman R, Athanasopoulos G, Bergmeir C, Caceres G, Chhay L, O'Hara-Wild M, Petropoulos F, Razbash S, Wang E, Yasmeen F (2019). forecast: Forecasting functions for time series and linear models. R package version 8.7, http://pkg.robjhyndman.com/forecast.

Hyndman RJ, Khandakar Y (2008). “Automatic time series forecasting: the forecast package for R.” Journal of Statistical Software, 26(3), 1–22. http://www.jstatsoft.org/article/view/v027i03.
