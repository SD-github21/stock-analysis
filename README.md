
# Analysis of Green Energy Stocks

## Overview of Green Energy Stocks Analysis

An analysis of green energy stocks was conducted in order to help identify which green energy stocks were the best to invest in:
* The analysis entailed developing a VBA macro to calculate important metrics to determine which green energy stocks performed well according to two target years, i.e., 2017 and 2018. 
* Total daily volume was investigated to provide an indication of how actively a stock was traded, while the yearly return was used to illuminate the percentage difference in the price of the stock from the beginning of the year to the end of the year. 
* By viewing these metrics, we could identify which green energy stocks appeared to perform well and this could help inform decision-making processes around which ones to invest in. 

However, although this VBA macro was successfully developed to conduct this analysis throughout the course of Module 2, the purpose of the current analysis was to determine if refactoring the code of this VBA macro enhanced the ability for that macro to run more quickly and efficiently. 

### Analysis Based on Refactoring

A substantial amount of the coding for the original VBA macro, entitled “yearValueAnalysis”, was utilized for the refactored macro, entitled “AllStocksAnalysisRefactored". For example, the following elements appear in both macros:

![image](https://user-images.githubusercontent.com/85533099/131233173-c0afca8f-b9f4-4d67-9f8e-7ca0a0867e56.png)

However, in the refactored code below, a variable called “tickerIndex” was created as a way to quickly grab each of the tickers for later analyses and set to 0. The metrics of interest, i.e., “tickerVolumes”, “tickerStartingPrices”, and “tickerEndingPrices” were established as array variables, by adding the number 12 in their Dim statements. Setting these variables as an array would later allow for "tickerIndex" to be combined with these variables for subsequent analyses:

![image](https://user-images.githubusercontent.com/85533099/131233070-6aaa17f2-4fa8-4aac-82cd-3bbf78659157.png)

The below section was added to create a “for loop” to set the "tickerIndex" to go through each individual ticker from ticker 0 to ticker 11 and to initialize each of the metric variables to 0, while adding the "tickerIndex" to each. 

![image](https://user-images.githubusercontent.com/85533099/131233077-7129bb9c-e0c8-474e-b428-f95705488ef0.png)

The below sections of code appear in both the original macro as well as the refactored macro with the only change being the addition of (tickerIndex) to each variable across all the calculations as well as the usage of "tickerIndex" for the output Worksheet.  

![image](https://user-images.githubusercontent.com/85533099/131233085-6ed120fc-2e42-4e39-868d-f864423cb3ef.png)

![image](https://user-images.githubusercontent.com/85533099/131233089-e4472048-2780-4627-a621-86b85db0cb6d.png)

## Results

The following screen shots compare the results of using the original macro versus the refactored macro in regards to the time the computer spent running through each analysis respectively:

### ORIGINAL MACRO – RUN TIME FOR YEAR 2017 --> 1.097656 seconds

![image](https://user-images.githubusercontent.com/85533099/131233092-4849695a-15e6-4fe4-9013-3a626c633e86.png)

### REFACTORED MACRO – RUN TIME FOR YEAR 2017 --> 1 second

![image](https://user-images.githubusercontent.com/85533099/131233098-6572ce18-30d5-448b-8996-a0fb8ca33315.png)

### ORIGINAL MACRO – RUN TIME FOR YEAR 2018 --> 0.9804688 seconds

![image](https://user-images.githubusercontent.com/85533099/131233099-6b52f469-5461-45c7-b0dc-8783f7553fd6.png)

### REFACTORED MACRO – RUN TIME FOR YEAR 2018 --> 1 second

![image](https://user-images.githubusercontent.com/85533099/131233104-6562ab11-3e65-417a-a7af-14275f895a4e.png)

In conducting a comparison between the total time each macro needed to complete their respective analysis, it appears that the refactored solution took a little less time for the 2017 analyses but a little longer for the 2018 analyses, although the total time recorded for each macro does change each time the macro is run and varies a bit each time. In this instance, the refactoring process did not yield a result, i.e., a change in run time, that was significantly different from the original macro that had been developed. 

## Summary

### Advantages and Disadvantages of the Refactoring Process in General
There are several advantages that the refactoring process can provide:
* First, refactoring can pare down the amount of code that is written in a macro, thereby creating a “sleeker” solution both visually and with regards to a more efficient implementation of the code. 
* Second, refactoring can significantly reduce the amount of time a computer takes to run the code. In other words, if the original code is longer and more detailed, this will take longer for the computer to read through and execute. 
* Third, if a macro isn’t working properly or there appears to be challenges in running the macro, refactoring that code may illuminate the source of these challenges and allow for coding errors to be detected and eradicated. 

There are also several disadvantages to the refactoring process: 
* First, if there is a large macro with a lot of code, refactoring might actually create coding errors if the analyst conducting the refactoring is not careful in how they are tracking the changes that they are introducing to the macro. 
* Second, it would be important to create a clear flow process of refactoring before it is implemented. For example, designating a single point person to work on refactoring at a time would be crucial, as having many people refactoring code would introduce more error to the process and also potentially duplicate efforts, which would waste time. 
* Third, it might be difficult to determine which parts of the code need refactoring in the first place, i.e., aspects of the code, at first glance, may seem perfectly reasonable to keep but this might actually be where the source of the challenge is. 
* Finally, it is possible that engaging in an elaborate refactoring process might not change anything about the efficiency of a macro. In other words, the original macro and the refactored macro may both take the same amount of time to run. This might not achieve the intended goal of trying to cut down on the amount of time spent for the computer to run the code and complete the analysis.  

### Advantages and Disadvantages of the Refactoring Process Involved in this Analysis
In regards to refactoring the current green stocks analysis macro:
* Each time I ran the analysis, the total run time varied. Sometimes the original macro was a little quicker than the refactored macro, while at other times, the refactored macro was slightly faster. However, I found that the overall amount of time for the computer to run the original macro versus the refactored macro did not change by a significant amount despite the time and effort I spent in working to refactor the original code. 
* I did find that the refactored code I was able to create was “sleeker” in places and that it helped to reduce some of the coding used in the original macro.
* Finally, I was able to catch an error in the original code. For example, in the VBA starter code that was provided to us as part of the Weekly Challenge, there was a “For Loop” at the end of the code, indicating “For i = 0 to 11”:

![image](https://user-images.githubusercontent.com/85533099/131233108-5bd5c952-ec25-42a2-9e6a-2bda7a37f785.png)

 This caused me a great deal of confusion, because I wrote the following code thinking that I actually needed that “For Loop”:

![image](https://user-images.githubusercontent.com/85533099/131233111-defffc6f-ce16-4a0b-8422-e3713435cd9d.png)

I received the following error message:
 
![image](https://user-images.githubusercontent.com/85533099/131233114-199795ff-290d-4c0c-9322-4014c5d2a2f1.png)

I eventually came to the conclusion that the “For Loop” seemed unnecessary and instead included the following code using "tickerIndex":

![image](https://user-images.githubusercontent.com/85533099/131233119-7d25ac2a-830d-42a1-b7d2-40921249e842.png)

Although I was able to achieve the desired results for the Weekly Challenge, I thought there had been something wrong with my deletion of the “For Loop”. I accessed resources through Live Assistants and the TA Office Hours as well as fellow students, and discovered that the “For Loop” was not needed and was clogging my analysis unnecessarily. In this case, I had created a solution to a problem that had been created in the original macro, but had thought my deletion of the “For Loop” was incorrect for the assignment purposes. In other words, I believed I had to work with the “For Loop” and refactor it in order to create a better solution. 

* The above experience illustrates an example of how aspects of the code may seem reasonable to include based on what exists in the original macro, but a person engaging in refactoring would have to be willing to take some risks and delete what had been established previously in order to successfully refactor that code. Working with smaller sections of the code and running them allowed me to identify where the problem existed, so in the end, the desired results were achieved. However, I believe that there is an element of "thinking outside the box" that needs to embraced in order to be truly successful with refactoring.
