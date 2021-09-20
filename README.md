# Stock Performance Analyzer
VBA-based stock performance analysis

###Overall Objective
The purpose of this analysis was to automate the analysis of a large database of stock information for the client.  Initially, a database of stocks and their daily performance at open, the high, low, and close were all provided in a dense list.  I wrote code to automate the summary of this data by analyzing stocks as a whole (by their name or ticker) and reducing the presentation to overall performance of the stock for the entire year of 2017 and 2018.

###Results
Two scripts were generated in the goal of analyzing this data as aforementioned. The first script went line by line for every single entry provided in the database. As a result, run times were approximately 0.8-0.9 seconds for both groups of stock data.  To mitigate this, arrays were employed to group the data so that collection of information was a run overall. This showed a nearly 10 fold decrease in the amount of time it took for the script to run (see \Resources\VBA_Challenge_2017 and ..\VBA_Challenge_2018 for the new run times)

![VBA_Challenge_2017](https://user-images.githubusercontent.com/59892063/133951976-1cb482a0-6075-467f-a8ad-4fe36391c5fb.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/59892063/133951964-43fde59d-fec9-45c7-9731-d00ba7b0cda0.png)
