# Challenge: VBA Refractor Code
Steve's parents are focused on a single renewable energy stock (DAQO, or DQ). He would like them to expand their horizons. To assist Steve, this work employs Visual Basic for Applications (VBA) coding to create an automated review (a macros) of the returns for 11 additional stocks. The final performance provided sufficient information, but time is money. The next step of the process was to improve macros productivity for the days when Steve wants to review additional stocks. The following demonstrates how the original code was modified to increase macros efficiency.

### Project Overview
With provided code, a macros was developed to analyze the return for the stock of interest, DQ. Coding was then expanded to focus on additional stocks. The final code uses a number of key elements to display highlighted differences among the stocks, including:
1. A timer to document the speed at which stock returns were calculated.
2. A popout box, allowing the user to determine which year of returns to populate.
3. Variables created for stock ticker symbols, starting prices and ending prices.
4. For Loops and Nested Loops allowed for code to repeat the review task, with a If/Then/End If formula directing the flow.
5. Feeding of data directly to an output worksheet, providing stock symbol, total volume and final return.
6. Color coding of return cells to highlight percentage stock increases (green) and decreases (red) over the year, as well as overall formatting of the displayed text.

The run concluded with a final statement of elapsed time from the ticker.

The final result were two tables, which provided an overview of stock changes between 2017 and 2018.
Between 2017 and 2018, only two stocks appeared to maintain sizeable gains: ENPH and RUN. These may be two additional stocks that Steve's parents could consider.
   
   ![Alt text](https://user-images.githubusercontent.com/30667001/147508707-35852e8f-7d4d-4d90-b5b7-04b16a5f398b.png)
   
### Execution Times: Original vs Refractored Script
The original coding required a number of iterations: an inner loop for total volume and starting and ending prices and an outer loop through ticker symbol before populating the new worksheet. The run times for 2017 and 2018 were over a minute (0.88 and 0.94 seconds respectively).

 ![Alt text](https://user-images.githubusercontent.com/30667001/147583905-4e58781a-9da0-4527-89e1-40e3082b5707.png)

The refracted code shorted the run times considerably (0.13 seconds apiece) by reducing the number of interations required. Removing the inner loop and directing the remaining for loops through the arrays optimized total run times.

 ![Alt text](https://user-images.githubusercontent.com/30667001/147586355-669eca00-80c2-408d-8f68-4c9ee1b59649.png)
 
 ### Summary of Findings
 #### Refractoring code: Advantages and disadvantages
 <INSERT HERE>
 #### Implications of advantages and disadvantage to VBA Wall Street script
  <INSERT HERE>



