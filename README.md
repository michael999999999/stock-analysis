# VBA of Wall Street

## Steve has recently graduated from his finance class, and his parents have decided to be his first clients. They would like to invest their savings in Daqo New Energy Corporation, a leader in green energy which produces the photovoltaic cells used in harnessing solar power.

### Steve has gathered a few alternatives to Daqo and has requested his Excel workbook be upgraded with macros in order to run several analyses. He will then use this information choose the best fit for his parents. The accompanying documentation includes the workbook with requested macros for easy use, as well as images demonstrating his stocks 2017 and 2018 performance. 

## Analysis

Based on the information provided by Steve, it would appear that Daqo Corp (ticker: "DQ") had a very strong 2017, as did 10 of the other stocks named. 11 out of the 12 options ended the year "in the green", some significantly so, with DQ topping the list at +199.4%! Other big winners include SEDG at +184.5%, ENPH at +129.5%, and FSLR at +101.3%.

![2017_Stock_Analysis](https://github.com/michael999999999/stock-analysis/blob/main/All_Stocks_Analysis_2017.png)

Results would take a turn for the worse in 2018, however, with 10 of those same stocks actually ending the year in the red. DQ stood out again, this time for the wrong reason as the largest loser. DQ ended the year -62.6%, followed by JKS at -60.5% and SPWWR at -44.6%.

![2018_Stock_Analysis](https://github.com/michael999999999/stock-analysis/blob/main/All_Stocks_Analysis_2018.png)

In my unprofessional opinion, it would seem as though ENPH would be the better option over DQ. ENPH showed growth in both years of data, +129.5% with the rest of the sector in 2017, and continuing that trend +81.9% in 2018. This shows strength as it was only one of two "green" stocks up in a year where that sector as a whole suffered. 

Additionally, while the +199.4% for DQ in 2017 would seem promising, they were the least traded of the 12 stocks, returning a volume of 35,000,000. For comparison, the next lowest 2017 volume was HASI at 80,000,000, then VSLR at 109,487,900. Further, in 2018 they were the 3rd lowest in volume at 107,873,900, sandwiched between HASI at 104,340,600 and VSLR at 136,539,100. Being towards the bottom of volume in both years seems to demonstrate a disinterest in the stock, especially considering most of the other stock saw a rise in volume.

- execution times of the original script and the refactored script.

As for the Excel document itself, the macro running the analysis was able to be refactored in a way that significantly improves performance. The prior code relied on nested loops performing mathematical equations, while the refactored code made use of VBA arrays. While mathematical equations must be used either way, arrays can help the process by storing values on the fly where nested loops are slower and can be a resource hog. 

![2017_Refactored_Code](https://github.com/michael999999999/stock-analysis/blob/main/VBA_Challenge_2017.png)

![2018_Refactored_Code](https://github.com/michael999999999/stock-analysis/blob/main/VBA_Challenge_2018.png)

## Results

- What are the advantages or disadvantages of refactoring code?
- How do these pros and cons apply to refactoring the original VBA script?
- The analysis is well described with screenshots and code (4 pt).

## Summary

- There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
- There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).

rough draft
