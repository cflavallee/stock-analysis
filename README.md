# Stock Analysis

## Overview of Project
In order to find the best performaing green stocks from 2017 & 2018, VBA code was used to evaluate stock performance.  This project could also be extended to additional stocks or the stock market as a whole. The refactoring code allowed for faster run times, enabling quicker analysis of various stocks. 

## Results
After refactoring the Module 2 code, the script ran roughly twice as fast (See screenshots below).  There are additional improvements that could be made, given more time and knowledge of VBA.

### 2017 Module Code

![2017 Module](https://github.com/cflavallee/stock-analysis/blob/main/Resources/2017%20Module%20Code.PNG)

### 2017 Refactored Code

![2017 Refactored](https://github.com/cflavallee/stock-analysis/blob/main/Resources/2017%20Refactored.PNG)

### 2018 Module Code

![2018 Module](https://github.com/cflavallee/stock-analysis/blob/main/Resources/2018%20Module%20Code.PNG)

### 2018 Refactored Code

![2017 Refactored](https://github.com/cflavallee/stock-analysis/blob/main/Resources/2018%20Refactored.PNG)

### Data Output for 2017 Refactored Code

![2017 Refactored](https://github.com/cflavallee/stock-analysis/blob/main/Resources/2017%20Refactored%20Run%20Time.PNG)

### Data Output for 2018 Refactored Code

![2017 Refactored](https://github.com/cflavallee/stock-analysis/blob/main/Resources/2018%20Refactored%20Run%20Time.PNG)

This was done mainly through two modifications.  First, all data was store in arrays prior to outputting the results in the analysis sheet.  Second, an Exit For was added after the Ending Price was populated, so there was no need to continue looping through the rows. Also, some redundant code was removed.  

Additionally, The output data was populated correctly, matching the Module 2 code output, as can be seen in the last two screenshots.  

## Summary

### Benefits 
The benefits of refactoring the code included having a quicker run time for the script as well as making the code a bit easier to read and understand. Although the original code was simpler in some ways, it had limited functionality.  

### Challenges
The main disadvantage was in the time it took to try new lnes of code and sequences.  It may not always be beneifcial to rework code that is effective, depending on the scope of the project. 

