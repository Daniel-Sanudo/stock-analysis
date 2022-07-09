# Refactored Green Stock Analysis
Previously, a subroutine to analyze a dataset from green energy stocks for years 2017 and 2018 was coded. In this proyect, the previous subroutine was refactored to improve it's performance so it returns a faster result when analyzing bigger datasets.

## Overview of Project
![Runtime Comparison](/Resources/Comparison.png)

This table summarizes the improved runtime for the refactored code; which is now 10 times faster than the original subroutine for a 3012 row dataset that includes the ticker name, date for the data, opening, closing, high and low prices as well as the traded volume.

![2017 results](/Resources/2017_Results.png)

The results for the analyzed stocks for the year 2017 show that all but TERP had a positive return. The highest return belonged to DQ with 199.4%, while the lowest positive return belonged to RUN.
TERP had a negative return of 7.2%. And the most traded stock was SPWR.

![2018 results](/Resources/2018_Results.png)

The results for the analyzed stocks for the year 2018 show that all stocks excep ENPH and RUN had a negative return. RUN had a positive return of 84% while ENPH had a positive return of 81.9%.
The biggest loss belonged to DQ as for this year their return dropped to a value of -62.6%.

### Purpose
The code was refactored since the original one took nearly a second to extract the traded volume and the return percentage for the 3012 row dataset that was used. Considering how Steve wants to look into more stocks, it's valid to assume that the size of the dataset will keep growing and so will the runtime for the subroutine. If 100 stocks for green energy were added, the runtime might increase tenfold; so the original subroutine would take up to 10 seconds to solve, while the new one would have the results in a second.

## Results
After refactoring, the relevant information is extracted much faster than the original code, this is because the new subroutine does not neet a loop to verify which ticker will have to be used because the data is already sorted by ticker name in an ascending order, thus there's no need for a loop but only a conditional that uses the previous and following cell to confirm if the opening and ending price for the current ticker are shown.

Regarding the green energy stock performance, the year 2017 had promising results since the return for almost all of them were positive, some of them returning upwards of 100% in a year which is a promising investment. Given how 11 out of our 12 stocks had a positive return, it's assumed that the year 2017 was a positive year for the green energy industry. DQ taking the lead with its great return value.

Now, for the year 2018, 10 out of the 12 stocks had a negative return. SPWR and JKS had a negative return that outdid the previous year's return and ended up with a lower closing value than their opening value for 2018. (SPWR started with 8.53 and ended with 4.97, while JKS started with 24.12 and ended with 9.89)
Only ENPH and RUN had a positive return this year, it could be assumed that 2018 was a tougher year for green energies.

Given the performance of the analyzed stocks through 2017 and 2018, ENPH and RUN would be the safest investments given how they've yielded a positive return on both years.

### Advantages of refactored code
Refactored code has the advantage of outputing a result much faster since it was modified to have a more specific function for this dataset. Since the data is already sorted then there's no need for an additional loop that verifies which ticker we're working with at each row.

The reduced downtime allows for faster decisions since the results are outputted much faster. In this case, the runtime does not exceed 1 second with the original code, but this 1000% runtime difference can have a hughe impact when working with datasets that exceed the hundred thousands rows.

### Disadvantages of refactored code
The disadvantage for the refactored code would be that it required the dataset to be organized in a certain manner so that the conditionals work. If by some reason, Steve added more stocks to the dataset, and did not sort them alphabetically, the results would be wrong. Moreover, Steve would also need to add the ticker names to the VBA tickers array to get the information for it.

### Uses of refactored code
This analysis proves that while refactored code has a much better performance, it also adds more requirements in order for those optimizations to work; in this case, the requirement is sorting the dataset by ticker name in ascending order. 
Refactored code proves to be a faster solution than coding a subroutine from the beginning, however this comes with the downside of more stric requirements for it to work.