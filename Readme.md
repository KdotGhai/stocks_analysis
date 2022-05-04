# VBA of WALL STREET
## Overview of Project
&nbsp;&nbsp;&nbsp;&nbsp;The purpose of this project is refactored the VBA code to work on larger dataset as efficient as possible to execute easy and fast analysis. To achieve this goal, some researches and manipulations covered during the process.  Some applications could be applied to any code to work faster and more efficiently. Throught the research part I have encountered some of them. For example there are adaptations for output process. You can easily close the screen updating, other irrelavant calculations or any other events from the process. These can give some efficiency. But not as much as rewriting to code to be faster. Instead of nested loops and rechecking all the dataset using an increasement arrays and single loops might give better results. So main purpose is to apply those changes to our code.
## Results
&nbsp;&nbsp;&nbsp;&nbsp;To analyze the different years stocks and their results first we need a Macro that can work on multiple datasets based on user choice. Also we need a timer to see how long it takes to execute the code. The outputs of the code need to be determined. That helps us to see the differences stock based on different years. 
### Stock performance between 2017 and 2018
&nbsp;&nbsp;&nbsp;&nbsp;The analyze shows us to year 2017 was much better in return performance than the year 2018. This resulted was achieved by the comparison of starting and ending prices of the indivual stocks during a year period. Results shown in png files.
### Initial code Vs Refactored Code Performance
&nbsp;&nbsp;&nbsp;&nbsp;Initial code was working and gave us the results that we desire to see. However its efficiency is not as much as expected to work on larger datasets and more stocks. Since we work only a dozen of indivual stocks its hard to get results. Then we refactored the code that gives us much quick and same analyze. Refactored code is almost six times faster than the first one (it may vary on larger groups but still faster). Instead of nested loops and look over and over the same groups, we create a single loop and increasment state based on specific conditions. 
These are two images that the results for each year and also the time that pass to execute the code. 



## Summary
&nbsp;&nbsp;&nbsp;&nbsp;In conclusion refactored code has a faster results and year 2017 was much better in return performance of given stocks. 
### Advantages and Disadvantages of Refactoring
&nbsp;&nbsp;&nbsp;&nbsp;When you are looking for larger datasets and multiple worksheets and thousand of Stocks Refactored code has a huge advantage. Since we are working on small-scaled projects the initial code will most likely be simple but re-usable, is easy to write and to manipulate. However, through refactoring the code it becomes more complex to ensure that it carries out the same task, if not more, with moderate to minuscule time shaved off. This process is a key role to any project since it's more user friendly, optimal, can be updated anytime, or re-used to other projects of this nature. 
