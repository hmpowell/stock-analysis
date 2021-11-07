# VBA of Wall Street

## Overview of Project

### Purpose

The purpose of this analysis was to help Steve give appropriate recommendations to his parents about their investments in the stock market, particularly in green stocks. The original stock in which they wanted to invest was determined to be losing money, so they wanted to find stocks that had good returns in 2017 and 2018 in order to make a more informed decision. We utilized VBA within Excel to determine the returns and present an easy to read and execute analysis of the results.

## Results


### Comparison of stock performance between 2017 and 2018

![This is a chart showing the return calculated in 2017 for 12 green stocks](https://github.com/hmpowell/stock-analysis/blob/main/Resources/Results_2017.png)

![This is a chart showing the return calculated in 2018 for 12 green stocks](https://github.com/hmpowell/stock-analysis/blob/main/Resources/Results_2018.png)

In comparing the 2017 to the 2018 stocks, there are only two which had positive returns for those two years: RUN and ENPH. ENPH had a very high return in 2017 and still had a high return in 2018. This seems like a great choice for investment if the trend continued. RUN had a much lower, yet still positive, rate of return in 2017 and then had a much higher rate (the highest of these stocks) in 2018. More research would have to be done to make sure this would continue on its upward trend. If they were looking to invest in one stock in 2019, I would have recommended ENPH; however, it is several years later now and the most recent years would need to be taken into consideration as well.

The original stock Steve wanted to consider for his parents was DQ. It is interesting to note that it had around a 200% return in 2017, while dropping significantly to a -62% return in 2018. It would be worth looking into specifics of what might have caused such a significant drop in return, as it went from the highest to the lowest return in one year. Since we are now in 2021, these stock prices would, as mentioned above, need to be considered for the most recent two years to understand whether this is a larger trend or if something occurred at the end of 2017 to cause such a drop and whether or not it recovered in 2019.

TERP was the only stock to have negative returns in both 2017 and 2018. I would recommend staying away from this stock unless more research had been done to prove that an increase in return on investment would be coming in the future.

Finally, it might be worth looking into why 2018 had such a significant drop in the return for these green stocks. It is interesting to see that only two stocks gave positive returns in 2018, while they were positive in 2017 (except for TERP as discussed above). Steve would need to look into what happened in the whole marketplace in general, whether it was only green stocks that had dropped significantly, and if there was an issue with only these stocks specifically. It would seem that something else was at play during these years that caused the stocks to drop this much.

### Comparison of execution times of original and refactored script

As was intended, the refactored script runs much faster than the original script:

#### 2017

![This is the original time it took for the script to run in 2017](https://github.com/hmpowell/stock-analysis/blob/main/Resources/Original_Run_Time_2017.png)
![This is the refactored time it took for the script to run in 2017](https://github.com/hmpowell/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

In 2017, the original script ran in 0.32 seconds, whereas the refactored script ran in 0.078 seconds. This means it ran about 4.1 times faster with the refactored code.

#### 2018

![This is the original time it took for the script to run in 2018](https://github.com/hmpowell/stock-analysis/blob/main/Resources/Original_Run_Time_2018.png)
![This is the refactored time it took for the script to run in 2018](https://github.com/hmpowell/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

In 2018, the original script ran in 0.31 seconds, whereas the refactored script ran in 0.074 seconds, which was also about 4.2 times faster with the refactored code.

## Summary

### Advantages or disadvantages of refactoring code

#### Advantages

There are advantages to refactoring code. The greatest advantage is that the code runs much faster. If there were thousands or hundreds of thousands of lines of data, it would allow the user to run the code and get the output faster than it would take the original code to run. In an example like the stock market, stocks change regularly and quickly. Taking less time might allow a user to calculate their output and make a faster decision than someone using a code that has not been refactored. That faster decision might allow a trader to make much more money than someone working slower.

Another potential advantage is that the code is more readable and more maintainable when it is condensed as refactored code. The next person using the code would be able to keep it running, as well as extend its usability as the amount of data increases. Code should continually be improved upon.

#### Disadvantages

When you refactor, you have the chance of creating bugs in already successful code. If your data is limited, it might be better to have it run for a slightly longer time, as opposed to risking changing your code. 

Another disadvantage would be that it takes more time to refactor and condense the code. It may cost a company or person more time and money to test and retest the code. If you do not have the time before a deadline or the money to continue to pay someone to refactor, then you would not want to continue.

These disadvantages would need to be factored in to any decision about whether or not to refactor a code.

### Application to refactoring the original VBA script

One of the disadvantages to refactoring the code mentioned above is something I experienced first hand within this challenge. During the module, I had the correct code and the correct output; however, once it was refactored, I had an incorrect output, which in this case, would result in the wrong recommendations being given to the investors. It took me over a week of trying and retrying to figure out the small error I had in my code.

The error came because I defined the variable tickerIndex = i and then mistakenly continued to use tickerIndex in my equation. The two different codes can be seen here (the first is incorrect and the second is correct):

![This is the incorrect code](https://github.com/hmpowell/stock-analysis/blob/main/Resources/Code_With_Error.png)

![This is the correct code](https://github.com/hmpowell/stock-analysis/blob/main/Resources/Code_Debugged.png)

A small change in the code resulted in a significant change in the results in the return column.

If I had to choose between refactoring or using the original code in this challenge, I would say that the risks did not offset the rewards. The code ran faster; however, it still ran very quickly in the original script. It would have saved a lot of time if I would have been able to use the original code.
