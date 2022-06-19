# stock-analysis
# Overview of Project

## Purpose
Our friend Steve recently graduated with a degree in finance and his parents want advice on which stocks to invest in.

The purpose of this project is to assist our friend Steve in analyzing stock data using Excel. We will also build macros using VBA to help automate the process so that Steve can run the analysis much easier on his own. 

## Results

### Performance Comparison 2017-2018
For the year 2017, most stocks saw a good amount of growth, while in 2018 most stocks lost value. It may be a good idea to invest in a stock that saw continuous growth in both years. Below you will find tables illustrating our analysis for each year. 



![results_2017](https://user-images.githubusercontent.com/35434608/174492972-8777dcce-82a1-402f-ac8b-65873776ed4c.png)


![results_2018](https://user-images.githubusercontent.com/35434608/174492981-6fbfcb4f-4a5d-42d9-ac7a-5062d5658758.png)



We can see that the stock ENPH grew by almost 130% in 2017, and almost 82% in 2018. This is probably our best choice as it shows that the stock is growing by relatively huge amounts in both years.

Another good option may be RUN which only grew by 5.5% in 2017, but saw growth of 84% in 2018. This is smaller growth compared with ENPH, but still it is the only other stock to show growth in both years.


### Code vs Code Refactored


We accomplished this analysis in three basic steps: 
1. First we wrote a script to analyze a single stock. The bulk of this work was done by a `for` loop, which ran through every line of the data set and took information based on certain conditions.
```
'loop over all the rows    For i = rowStart To rowEnd			
	'check that we have the right stock        If Cells(i, 1).Value = "DQ" Then            'increase totalVolume by the value in the current row            totalVolume = totalVolume + Cells(i, 8).Value        End If	'check if first instance of stock        If Cells(i - 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then            startingPrice = Cells(i, 6).Value        End If	'check if last instance of stock        If Cells(i + 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then            endingPrice = Cells(i, 6).Value        End If    Next i
```

2. Next we generalized that script so that it analyzed all stocks. This was done by creating an array of all the stocks, and feeding each of those stocks through our initial `for` loop. We have our original `for` loop nested inside the loop of all stocks.

```
' Loop through tickers   For i = 0 To 11       ticker = tickers(i)       totalVolume = 0              ' loop through rows in the data       Sheets(yearValue).Activate       For j = 2 To RowCount                  ' Get total volume for current ticker           If Cells(j, 1).Value = ticker Then               totalVolume = totalVolume + Cells(j, 8).Value           End If                      ' get starting price for current ticker           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then               startingPrice = Cells(j, 6).Value           End If           ' get ending price for current ticker           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then               endingPrice = Cells(j, 6).Value           End If                  Next j              ' Output data for current ticker       Worksheets("All Stocks Analysis").Activate       Cells(4 + i, 1).Value = ticker       Cells(4 + i, 2).Value = totalVolume       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1   Next i
```


3. And last we refactored that script to optimize for more clear and readable code, while also shortening the runtime. The refactored code works a little differently but the logic is very similar. Instead of outputting our results inside the loop, we just stored each result into an array of its own and output those results one time at the end.

# Summary
1. What are the advantages or disadvantages of refactoring code?

Refactoring code is like creating the final draft of an essay. You take the time to think through the process and make things as efficient as possible. A frequent concequence of this process is that it makes the code more clear and readable for whoever is going to look at it next, even if it's your future self. Another advatage of refactoring code is that you may be able to isolate some functions you've created in order to use them in future projects. If we isolate the section of our code that cleans up the formatting of our data sheet for instance, we may be able to use that macro on any future project.

A disadvantage of refactoring code is that you may not always have the time to do it. It's possible that you may be on a very tight schedule and need a quick and easy solution straight away. You can of course always go back and refactor later when you have the time.
 
2. How do these pros and cons apply to refactoring the original VBA script?

In our case, refactoring the code improved our runtime from ~0.64 seconds to only ~0.13 seconds. Now while this may not seem like a huge difference, this would add up if we were to try to run the original code on a much larger data set. Below are the recorded runtimes of the refactored code for years 2017 and 2018, followed by the runtimes of the original code.






![VBA_Challenge_2017](https://user-images.githubusercontent.com/35434608/174492988-87f58d87-dd40-4cf2-8f49-7e41b1741bfa.png)


![VBA_Challenge_2018](https://user-images.githubusercontent.com/35434608/174492998-e39f296e-f91d-483c-bc78-9d0f0cbdd73a.png)


![OG_code_runtime_2017](https://user-images.githubusercontent.com/35434608/174493121-cce9176a-a203-40be-ab74-9addb73287d2.png)


![OG_code_runtime_2018](https://user-images.githubusercontent.com/35434608/174493190-5b623fea-1a2c-4f64-98df-703f36355636.png)
