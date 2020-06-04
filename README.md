# Options_Based_Trading

Disclaimer: When I wrote all of this code, I was very much a novice with python and more practical coding concepts (learncpp.com and w33schools.com only gets you so far). I have not looked at the code contained within these scipts since they were created. If there is interest in this type of analysis, I can look into giving it a much needed rewrite/update. Please let me know in the issues tab.

This program analyzes and evaluates the performance of unusual option orders made by investment banks, hedge funds, etc. You can generate daily or weekly reports, with their respective scripts. The Performance_Evaluation.py script can take raw data from any number of trading days and compute the expected return on those options if they were held to a given expiration day, without any hedging with the opposite option contract.

I am no longer working on the credit_spread_daily.py script, but in its current form it can produce daily reports of unusual credit spreads that were opened for that trading day.

Required folders in the main directory:
For daily_report.py, Performance_Evaluation.py, and weekly_reports.py:
A folder titled ‘output’ is required in the working directory. The excel file produced by the script will be titled ‘output’ and found in the ‘output’ folder.


For daily_report.py and weekly_reports.py:
Download a csv file from https://www.barchart.com/options/unusual-activity/stocks and place it a folder entitled ‘data’ in the working directory.

For Performance_Evaluation.py:
Place the same csv files as the daily and weekly reports and place it in a folder titled ‘data_for_validation’.

For Credit_spreads_daily:
The two folders you will use are ‘spread_data’ and ‘spread_output’. They work in the same fashion as the previous folders.

