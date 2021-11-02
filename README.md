# Stock-Analysis

## Overview of Project

### Purpose of this project was to expand the dataset to include the entire stock market over the last few years (2017 and 2018)

## Results! 
![VBA_Challenge_2017](https://user-images.githubusercontent.com/89805399/139947856-44fc37bf-32ee-498a-83ed-470f64293cc7.png)
In 2017, the code took about 1.445 secounds to run before we got our results.If you look at the data above almost all of the total daily volumes are in the positive percentage (green) except for "TERP"
![VBA_Challenge_2018](https://user-images.githubusercontent.com/89805399/139948368-ea0c1c3a-7208-4f8d-98fb-04be210d54e8.png)
In 2018, the code took about 1.457 secounds to run before we got our results. Compared to 2017 all of the data are in the negative (red) except for "ENPH" and "RUN".
The pros of refactoring the code are making the code simpler. On the "VBA_challenge" sheet we added totalVolume where were initializing the variables for starting and ending price. the code looked like this (Dim totalVolumes as Long). We also refactored the code to use (yearValue) worksheet rather than "2018" worksheets. One of the cons and issues I was running into creating the refactored code for "VBA_Challenge" is a run time error for the output section. For example, the code line (Cells(4 + i, 3).Value = tickerEndingPrices / tickerStarting{rices - 1) would run into a run time '6 error. Another issue I was having is that I had "Dim totalEndingPrices as Single" code originally set to "Long" which gave me different total daily value results.
All do refactoring the code made it more presentable and readable it was more difficult to run and a lot more debugging on my end to get to successfully run.
