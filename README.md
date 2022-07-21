# Stock Analysis for Steve
This analysis and project was built for Steve and his parents. They wanted to analyze the stock market performance of DQ as well as multiple other potential investing opportunities. With VBA I created macros that will go through the data set of stock prices and show Steve each stock's performance through total daily volume and return for both 2017 or 2018.

## Results
![VBA_Challenge_2018 png](https://user-images.githubusercontent.com/108035549/180322276-27bf0646-ec0d-4a55-aac4-ba1254e17f69.png)
![VBA_Challenge_2017 png](https://user-images.githubusercontent.com/108035549/180322303-96e8f7dc-b3f6-4251-b092-2ea2d617613c.png)

Our 12 stocks had better returns in 2017 than 2018. Only 2 of the 12 stocks had positive results in both 2017 and 2018. These stocks were RUN and ENPH. After initially doing analysis on DQ only, I refactored the code that gave us the DQ results to include results for all of the 12 stocks. After refactoring the code again to include even more stocks, I managed to make a button for Steve to push to run the Macros, and by using the Start Timer and End Timer functions in VBA, was able to track how fast it took for the code to produce the desired results. Initially it took over 2.5 seconds, but after the refactoring and including arrays i was able to lower that to less than half a second, which is visible in the two screenshots attached above. Some examples of the array codes are included below. The arrays were quite tricky for me to figure out and were easily the most difficult part of this project.

    (Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    tickerVolumes(i) = 0 
    tickerStartingPrices(i) = 0 
    tickerEndingPrices(i) = 0) 
    
    
## Summary
Refactoring code has both advantages and disadvantages. For one, you have a working framework to reference and can copy some parts of the code in other places. However, it can be hard to keep track of what small changes need to be made when refactoring and it is easy to get lost in the code. In addition, some parts of the code that you copied earlier may not work if the proper syntax is not added when refactoring or adding. With more for loops and more if/then statements I had to be extra careful and go slowly, particularly to get the provided code to work with the code I was building.
