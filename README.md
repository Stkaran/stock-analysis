# Stock-analysis

## Purpose
At the start, this code was used only to find the yearly return of a singular stock, DQ, because that is what our friend was intially focused on. However, by going back and both adding and changing some code, we were able to able to expand on what information the code would pull for us all while making it run more effciently.

## Results
The first attempt at this project ended up with a runtime of around 1.13 seconds. We were able to reduce to quite significantly with our refactored as seen below. 

![image](https://user-images.githubusercontent.com/89424470/133722521-5f00fff8-710d-4ce4-9a1e-d3516ad040b3.png)![image](https://user-images.githubusercontent.com/89424470/133722570-2e089e2e-d9c0-4592-9768-5567230aa1e7.png)

The main reason of this increase of speed was due to the use of mutiple arrays in our code. On our first attempt, we used a for loop with a nested loop in order to interate through data, tagging the specfic stock we wanted to focus on. In the refactored code, rather than having just one array for the companies of the stock, we  made three output arrays for the starting and ending prices of the stocks for whatever year we told the code to look through. And with the inclusion of the tickerIndex, the code was able to keep track of the values the code was reading. This made it so the code only had to loop through the data once, which in turn decreased its runtime.

    Dim tickerIndex As Integer
       tickerIndex = 0
    
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
There was also another major change that was added. Before, we used a serparate subroutine in order to format the data to be more presentable to an audience. An extra step can be a bit cumbersome, so the formatting code was also added to the refactored code's subroutine. And even with this addition, the refactored code still ran faster than before.

## Summary

So in conclusion, it is clear that there are some major advantages when refactoring code. The most obvious of these is the increase in effciency. On a larger scale changes like what we did here could save massive amounts of time. At first we just want our code to work, but after that  going back and making it as simple as possible will not only increase performance, but if you ever need to come a edit it again, it will be easier  for you or someone else if it is streamlined. On the other hand, rafactoring can have downsides. One wrong parentheses can cause a code not to work, so when going back and changing a lot things can leave you in a place where you end not only having code theat doesn't work, but you've also created even more problems for yourself than you had before. This can end up being a huge timesink where you only make yourself more frustrated.

In this case specifically, at first I thought my conditionals need to be in a nested for loop. Eventually I figured out they did't need to be and changed things back, however, the code still wasn't working.

    '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        'End If
        End If
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then

            tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
        
        
        '3d Increase the tickerIndex.
       
            tickerIndex = tickerIndex + 1
   
When trying to debug it, the error messages were always going to the first line in this code. After quite a bit of time, I finally realized the tickerEndingPrices was still set to a value of j from when I was doing the nested loop. When compared to a text editor like Vs Code, VBA does make the refactoring process a bit more diffcult when making small errors such as this one. The major upside to VBA has to be the ability to save mutiple macros in one place. A big majority of the final code was just copy pasted from previous work (formatting, orignal ticker array, etc.). This saved lots time makes combining multiple subroutines very easy. 
