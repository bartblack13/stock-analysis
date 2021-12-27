# Stock Analysis
## Overview of Project
The purpose of this project was to gain familiarity with using the Microsoft Excel Visual Basic for Applications (VBA) to analyze large sets of data. Since VBA requires the use of coding language, this project also served  to start building a foundation of writing and manipulating coding languages, that can hopefully translate and be applied to quickly learning other codes.  It was done as part of my Week 2 project for my DU coding bootcamp.  My work with the data set included:

* initializing variable of different types
* using loops to automatically loop through large data sets, including nested loops
* using conditional arguments
  a) "if-then" statements
  b) and/or statements
* initiallizing and using arrays and indexes to make writing and executing code more efficient
* reusing, debugging, and commenting on code
* formatting outputs
    a) fonts, interior, cell color, etc
    b) number formats
    c) conditional formatting
* creating and using:
    a) Message Boxes
    b) user input boxes
    c) creating user friendly buttons
* measuring code performance via embedded timer
* refactoring code


## Results
The goal of this project was to analyze stock performance data from 12 different green energy companies for the years 2017 and 2018 for a "friend" named Steve.  Steve's parents were interested in investing into the company DAQO (Ticker: DQ), but were open to other companies that might be a better investment.

The stock data that was provided included daily trade volume and daily pricing (open, high, low, close); a years worth of daily trade info across multiple variables for 12 stocks (3012 rows of data), for two separate years.  By writing VBA macros, I was able to analyze each dataset in aproximately 4.5sec., and automatically generate a table displaying "total daily volume" and "percent return", with positive returns highlighted in green and negative returns in red (see below for tables and figures).  I was also able to write a partially successful program, which was capabale of performing the above anlysis and table generation in between 1.27-1.37sec (see figures below).

In 2017, DAQO had a positive return of 199.4%, but had a negative return of 62.6%.  Across both years, they increased their total daily volume from 35,796,200 - 107,873,900, which shows that interest in the stock has increased, indicating that their company might have a favorable next year.  However, this would require a bounce back, which is risky.  It would be a better to consider some of the other stocks that were analyzed.

Of the 12 stocks, only 2 stocks had positive returns for the year 2018, which were ENPH and RUN.  In 2017, only one stock had a negative return which was **not** either ENPH of RUN.  From the start of 2017 to the end of 2018, ENPH stock increase from $1.05/share to $4.73/share (+350.5% return), and almost trippled their total daily volume by years end.  Across the two years, RUN increased their stock price from $5.59/share to $10.89/share (+83.95% return), aproximately doubled their total daily volume.  Both companies have substantially higher trading volume than DAQO, which suggests that thier is more investor interest in ENPH and RUN than DAQO.  However, ENPH and RUN have single digit prices per share, across the 2yr timeframe, which is usually seen in companies that are new to the public exchange markets.  The lower price per share, meansthat Steve's parents could either invest less money, or they could buy more shares for the same money and have a higher yield if the companies value takes off.

**Conclusion:** Considering this initial and basic stock analysis, Steve's parents would be much wiser to invest in either RUN or ENPH, rather than DAQO.  


embed picture (2017 and 2018 output graphs)

## Summary
In a summary statement, address the following questions.
### **What are the advantages or disadvantages of refactoring code?**
The largest advantage of refactoring code is that a workable solution is already at hand.  There is also a template and pattern to approach the problem with.  This allows the code writer to easily copy and paste or even just reorganize the code to make it more efficient and cleaner.  However, I do see that it could force you into a coding strategy that might not be the easiest or the best, as you are using the pre-exisiting logic as foundation for your editing.

Another advantage is that you can make the code more efficient, allowing it to run quicker; be applied to multiple data file sets; or allow for a data set to dynamically change while still using the same code.  This gives the coder a huge advatage for time management.  If the coder can spend a bit of extra time writing a well thought out code in order to avoid writing multiple macros for similar files, then the coder has won the long game.

### **How do these pros and cons apply to refactoring the original VBA script?**
For me, the original code worked flawlessly.  It was averaging run speeds of 4-4.5sec., which is still pretty quick, but could definitely be faster, especially in today's internet based world.  I was able to refactor the orignal code as I was learning it to make it a bit more efficient.  For example, as soon as I learned how to use an input box, I began editing all of the "worksheet activation" codes, so that I could auto populate the values neccessary to change between analysis and output sheet regardless of year.  

The down side for me was that the predetermined structure of the refactored code pigeonholed me into trying to use coding strategies that I wasn't super confident with.  In a lot of ways, it might have been easier for me to develope my own strategy and write a lot of the code from the ground up.  This is especially true, since I spent a lot of time trying edit prexisting code, make sure it was consistent, make sure it was given to me without errors or misspelled variables, etc.  Sometimes, starting from scratch is better than retooling something, regardless of the industry or application.

In the end, even with the problems that I encountered, and the disfunctional code, unable to generate correct output values, the refactored code was running aout 60% faster.  I think that given enough time and more experience, I could problem solve my exisiting refactored code to work properly and run more quickly than my original code.

