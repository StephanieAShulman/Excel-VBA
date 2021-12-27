# Challenge: VBA Refractor Code

## Project Overview
While Steve's parents are focused on a single renewable energy stock (DAQO, or DQ), he would like them to expand their horizons. This work employs Visual Basic for Applications (VBA) macro coding to create an automated review of the returns for 11 additional stocks. The final performance provided sufficient information, but - as with anything - time is money, and the next step of the process was to improve macros efficiency. The following demonstrates how the original code was modify to increase productivity.
   ### Initial Approach
   With provided code, a macro was developed to analyze the return for the stock of interest, DQ. After using Range() and Cells() to create necessary headers in a new worksheet and establishing starting and ending price variables, the internet was consulted for the best way to identify the end of the worksheet. For loops and conditionals were used to pull stock-specific information from 2018 data and populate the new worksheet. The final result showed a return of -63% ... Steve was not impressed.
   
   
