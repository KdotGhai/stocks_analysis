# VBA of WALL STREET
## Overview of Project
&nbsp;&nbsp;&nbsp;&nbsp;The purpose of this project is refactored the VBA code to work on larger dataset as efficient as possible to execute easy and fast analysis. To achieve this goal, some researches and manipulations covered during the process.  Some applications could be applied to any code to work faster and more efficiently. Throught the research part I have encountered some of them. For example there are adaptations for output process. You can easily close the screen updating, other irrelavant calculations or any other events from the process. These can give some efficiency. But not as much as rewriting to code to be faster. Instead of nested loops and rechecking all the dataset using an increasement arrays and single loops might give better results. So main purpose is to apply those changes to our code.
## Results
&nbsp;&nbsp;&nbsp;&nbsp;To analyze the different years stocks and their results first we need a Macro that can work on multiple datasets based on user choice. Also we need a timer to see how long it takes to execute the code. The outputs of the code need to be determined. That helps us to see the differences stock based on different years. 
### Stock performance between 2017 and 2018
&nbsp;&nbsp;&nbsp;&nbsp;The analyze shows us to year 2017 was much better in return performance than the year 2018. This resulted was achieved by the comparison of starting and ending prices of the indivual stocks during a year period. Results shown in png files.
### Initial code Vs Refactored Code Performance
&nbsp;&nbsp;&nbsp;&nbsp;The initial code for our standard yearly analysis was functional and gave us desired results but it leaves a user and programmer wondering if it can be done better.The efficiency of the code is in question here and although our results are small as is as well as, when refactoring the time shaved is even smaller we must remember that its due to the size of the dataset that should be taken into consideration when addressing time complexity. Therefore, we refactored the code that gives us much quick and same analyze. Refactored code is almost six times faster than the first one (it may vary on larger groups but still faster). Instead of nested loops and look overs the same groups, we create a single loop and increasment state based on specific conditions.  
<br>
These are images that the results for each year and also the time that pass to execute the code:
- #### Standard Analysis
<p align="center">  
<img src="https://github.com/KdotGhai/stocks_analysis/blob/f0bf79fc445df299918800bfb7c7c94dfbea8a08/Images/VBA_2017_Standard_Year_Analysis.png" />
  <br>  </br>
</p>
<p align="center">  
<img src="https://github.com/KdotGhai/stocks_analysis/blob/f0bf79fc445df299918800bfb7c7c94dfbea8a08/Images/VBA_2018_Standard_Year_Analysis.png" />
  <br>  </br>
</p>

- #### Refactored Analysis
<p align="center">  
<img src="https://github.com/KdotGhai/stocks_analysis/blob/f0bf79fc445df299918800bfb7c7c94dfbea8a08/Images/VBA_2017_Refactored_Year_Analysis.png" />
  <br>  </br>
</p>
<p align="center">  
<img src="https://github.com/KdotGhai/stocks_analysis/blob/8736dc83eaeac52683952dfe46d760634a4fc6ae/Images/VBA_2018_Refactored_Year_Analysis.png" />
  <br>  </br>
</p>



## Summary
&nbsp;&nbsp;&nbsp;&nbsp;In conclusion refactored code provided faster results and the year 2017 was much better in return performance of given stocks. When we are looking at larger datasets, multiple worksheets, and thousands of Stocks, refactored code is more advantageous. Since we are working on small-scaled projects the initial code will most likely be simple but re-usable, is easy to write and to manipulate. However, through refactoring it becomes more complex to ensure that it carries out the same task, if not more, with moderate to minuscule time shaved off. This process is a key role to any project since it's more user friendly, optimal, can be updated anytime, or re-used on other projects of similar nature. 
