# Stocks-analysis
## Project Overview
---
***Background***

Steve asked me to analyze a group of 12 green stocks to support his parent’s investment decisions. To do so, I gladly designed an interactive, user friendly, workbook using Visual Basic Application (VBA) within Excel to provide each stocks annual volume and return on investment (ROI).

He loved being able to analyze each stock at the click of a button and now wants to expand his research beyond the 12 green stocks.

***Purpose***

Steve wants to analyze a higher number of stocks and I am here to help. This may increase the amount of time it takes the analysis to produce results and I’d like to maintain or, even better, improve it! Now I will take advantage of any opportunity to improve the workbooks efficiency by refactoring the VBA coding. To ensure that I am going in the right direction, I will compare the new execution time with the original workbook.

###Results

![](/Resources/VBA_Challenge_2017.png)![](/Resources/VBA_Challenge_2018.png)

As it can be seen just by glancing through the findings, the 2017 performance was considerably better as there is a lot more green to be seen in the table. This means there are more returns that came out positive. This is specially true for stocks such as Daqo, that in 2017 came out with a return of 199.4% and in 2018 of -62.6%. Similarly, this same happened for most of the stocks including Atlantica Yield (ticker: AY), Canadian Solar (CSIQ), First Solar (FSLR), Hannon Armstrong Inc. (HASI), Jinko Solar (JKS), SolarEdge (SEDG), SunPower (SPWR), and Vivint Solar (VSLR).

In conclusion, even though the vast majority of them performed well in 2017, most of them performed poorly in 2018. Definitely, the top picks would have been Enphase Energy (ENPH) or SunrRun (RUN), which were the only green energy stocks that gave high returns for both consecutive years. The best investment would have been to buy a lot of ENPH stock at the beginning of 2017 and hold it through 2018 as it had 129.5% and 81.9% return, respectively. Nevertheless, money could also have been made by buying a put option on TERP at the beginning of 2017 as it returned -7.2% and then in 2018 another -5%.

###Execution Times and Code Performance
As it was mentioned in the overview of the project, I first created a VBA script with multiple macros that analyzed the stock data and formatted the results. It did the job, however, I then thought I could refactor the code to be able to do everything in one single macro. This would not only decrease the execution time, but it would also serve me as an efficient design pattern that I could share and use with different stock data.

Before refactoring, my "draftCode" script in module 1 was running the analysis in 0.97 seconds without even including formatting as I had that done in another macro. The code execution time decreased significantly after I refactored.
As you can see, the code run time for both years decreased almost by half to 0.57 seconds. Even though this might not be a huge difference in this case, when we work with significantly larger datasets code performance and execution times will be more important to us.

###Summary
___
Refactoring code has its pros and cons. Just like proofreading, going back and reevaluating what we wrote is a key step for better results. One pro would be that it makes the script easier to understand. A clean code is more versatile which is a huge advantage as code is always shared. On the other hand, the disadvantages behind refactoring revolve around the fact that it is just time consuming. Some people even argue that it is useless, as you are working on a body of code that already works and gets the job done.

In this project, this could be easily seen as the code performance and execution times relatively got better. Refactoring the code that did the analysis of the stocks definitely made the script more coherent. Now, I didn't have the iterations in different for loops or the formatting in another subroutine. Instead, I could just run the refactored macro and everything would be done at once. Even though the advantages were many, it took me a lot of time to come up with the refactored version of the code. I needed one single nested for loop to do the job and store the data I needed into arrays rather than simple variables. But even though it was time consuming, I believe it was worthwhile as now I have a body of clean code that I can easily edit in order to analyze different stock data. The goal should be to always strive to create a design pattern that is versatile, clean, easy to understand and shareable.

In conclusion, the original script had the advantage that if the code was wrong at any point, it would only affect that macro and function. The others would still work. Meanwhile, the disadvantages were that it was a very long and all-over-the-place script. On the other hand, the advantage of the refactored script was that it was more efficient and organized, making the execution times a lot shorter and the script easier to follow.




