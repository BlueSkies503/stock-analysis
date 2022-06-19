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
(talk about importance of run times with larger data sets, use images of run time msgBoxes and exapmles of code)

We accomplished this analysis in three basic steps. 
1. First we wrote a script to analyze a single stock.
	- The bulk of this work was done by a `for` loop, which ran through every line of the data set and grabbed information based on certain conditions.
```
'loop over all the rows    For i = rowStart To rowEnd			
	'check that we have the right stock        If Cells(i, 1).Value = "DQ" Then            'increase totalVolume by the value in the current row            totalVolume = totalVolume + Cells(i, 8).Value        End If	'check if first instance of stock        If Cells(i - 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then            startingPrice = Cells(i, 6).Value        End If	'check if last instance of stock        If Cells(i + 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then            endingPrice = Cells(i, 6).Value        End If    Next i
```

2. Next we generalized that script so that it analyzed all stocks.
	- This was done by 
3. And last we refactored that script to optimize for more clear and readable code, while also shortening the runtime.

# Summary
1. What are the advantages or disadvantages of refactoring code?
2. How do these pros and cons apply to refactoring the original VBA script?
Pro - more readable, faster for computer to process
Con - can take more human effort, or more human time to create. opposite of "quick and dirty"






![VBA_Challenge_2017](https://user-images.githubusercontent.com/35434608/174492988-87f58d87-dd40-4cf2-8f49-7e41b1741bfa.png)


![VBA_Challenge_2018](https://user-images.githubusercontent.com/35434608/174492998-e39f296e-f91d-483c-bc78-9d0f0cbdd73a.png)


![OG_code_runtime_2017](https://user-images.githubusercontent.com/35434608/174493121-cce9176a-a203-40be-ab74-9addb73287d2.png)


![OG_code_runtime_2018](https://user-images.githubusercontent.com/35434608/174493190-5b623fea-1a2c-4f64-98df-703f36355636.png)
