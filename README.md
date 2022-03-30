# Stock Market Analysis VBA

## Overview of Project

Steve requested that I look at some VBA code that was built for "green" stocks. This code did several jobs, such as looking at yearly trade volume, stock returns over a year, and formatting of cells and text. The main purpose was to look at these "green" stocks to see if they would be worth investing Steve's money into. After reviewing the selected stocks, Steve realized he may need to look outside of "green" stocks to ensure his client's funding is being invested wisely and wanted to make sure this report could handle that need. 

### Purpose of My Review

Since Steve may want to look outside of these original "green" stock selections, he asked me if I could update the report and VBA script to be used on all stocks in the market. He, wanting to treat every client well, wants to ensure his reporting is correct and efficient. 

Since the goal of this report is to review every stock in the market, I have been asked to make this report as efficient as possible. Reviewing 12 stocks any report would be quick, but if we bring a report like this to the entire stock market being as efficient as possible would be ideal

## VBA Analysis Results

### VBA Script Refactoring

When Steve provided me the original report, I determined I could do several edits or refactoring to it to make it more efficient and have it applied to several more stocks when that would be needed. To do this I determined that I could change how the script ran, by creating four new arrays. tickerIndex, tickerVolumes, tickerStartingPrice, and tickerEndPrice. 

![image](https://user-images.githubusercontent.com/100856534/160916986-ab98c1d3-4a7c-4b23-bde9-de16c1b91259.png)

By doing the refactoring this way, it would pull in the ticker symbol, starting price, ending, price, and volumes once before it goes through the remainder of the script. Whereas before for each stock it would do these processes as it goes for each stock, making it a little slower. 

### Refactored Code Examples

**Refactored Script for Looping Through Iterations**

![image](https://user-images.githubusercontent.com/100856534/160918878-b421d060-3ef4-411f-a202-6afa8d1a91e8.png)

**Original Script for Looping Through Iterations**

![image](https://user-images.githubusercontent.com/100856534/160919585-83e2cba6-22ea-4fa9-852d-52e21825f7fa.png)

This shows a slight difference. In the original script the ticker is the variable where ticker was equal to ticker(i). i was equal to the 12 different stocks. It would need to iterate through each stock one at a time. Where as with the refactored code tickerIndex in referenced with each array, such as totalVolume(tickerIndex) and tickerStartingPrice(tickerIndex). This produces a much quicker reference to the information that is needed to be pulled.

### Speed of Macro Script

When the original script was ran, the speed of the whole process, took arund 1.07 seconds for both the 2017 and 2018 years. When the code was reactored to make it more efficient that same product that was produced was around .13 seconds for both 2017 and 2018. That is .94 seconds faster for each year, or 87% faster. Again, this is not that big of a difference for the data set we are using, but that difference would be much larger if it came to producing this information for the entire stock market. 

**2017 Refactored Results**

![VBA_Challenge_2017](https://user-images.githubusercontent.com/100856534/160921344-50a4cc07-942e-42b9-8e2c-6b4bc42dbbc8.png)

**2018 Refactored Results**

![VBA_Challenge_2018](https://user-images.githubusercontent.com/100856534/160921418-3e1f82f5-aa6c-4419-8761-34e65e42fdfb.png)

## Stock Market VBA Analysis Summary

1) What are the advantages or disadvantages of refactoring code?

One of the major advantages of refactoring code is that you do not need to start from scratch, a subroutine would not need to be built from the first command. This saves a programmer time and effort when it comes to working on new or established projects. However, this also comes with some disadvantages as well, as when it comes to code that you did not write you would need to make sure that the original script works for the data you are working on. Also, not every script writer uses the same syntax when it comes to writing subroutines. 

How do these pros and cons apply to refactoring the original VBA script?

A lot of the work on this script was already done for me. What I needed to do was to add the efficiencies asked of me by Steve. I could focus on writing the script I needed to after a quick review of the code that was already there. However, one disadvantage I found was when none of my code would work when it came to the subAllStocksAnalysisRefactored() subroutine. After debugging line by line I discovered that some of the WorkSheets("All Stocks Analysis").Activate commands did not work for me as I had a typo from my original code to the refactored code that was provided. My sheet name was "All Stock Analysis" with stock missing the s that was in the refactored code. This being a prime example of a disadvantage, since I was using someone else's script with my own worksheet set up, they did not immediately match and errors occurred, thus taking up some more time.
