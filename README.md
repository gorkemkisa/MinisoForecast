# Miniso Forecasting Tool

## WARNING: This software is not affiliated with Miniso/Miniso Turkey. It doesn't contain any kinds of confidential, private data and/or methods. This is just a senior design project that we've created with collaboration of Miniso. They have all the rights to change, or not use at all.

This tool basically pulls the Excel data (.xlsx or .xlsm), applies interpolation (Missing Data Handling) and/or exterpolation (Forecasting using Theta function), and prints the results to an Excel file. 

The Excel data should be formatted like this:

- First column should have ID's (Unique identifiers of the data)
- First row should have time indexes (Day, month, week, etc.)
- The sale data should be filled like a matrix in this format.


This software is built and published by Görkem Kısa. This software is a part of my team's Senior Design Project, anyone who has access to this page can inspect and use the code freely under the use of MIT license(if you want to use this, please reference the contributors of the packages and libraries that I've used). Form1.cs is the main source code.

Our team:
- Görkem Kısa
- Aslı Çahantimur
- Ali Şamil Adıgüzel
- Ecem Akış

I also want to thank Onurhan Akçay for his help in terms of UI in this project.

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
