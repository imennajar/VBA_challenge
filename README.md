# How to write a script in VBA to manipulate input data from Excel, calculate and display output data: :cd:

In this project we will learn how to use VBA to  loop through all the stocks for one year and produce useful data output

# What we will learn from this project:

    - How to declare and use the right variables
    - How to use the right instructions: data entry, condition, iteration, and data display
    - How to call predefined and created functions and subroutines
    - How to apply the same code in several sheets
    
# Instructions:
    -The ticker symbol: the symbol of each ticker without redundancy
    - Yearly change = opening price at the beginning of a given year - closing price at the end of that year
    - Percentage change = (yearly change/ opening price at the beginning of the year)*100
    - Total stock volume = sum of the stock volumes of each ticker
    - Greatest % increase: the maximum of percentage change
    - Greatest % decrease: the minimum of percentage change
    - Greatest total volume: the biggest total stock volume

# Software used:

MS Excel

#  Program:

## Place to store VBA code 🔎

Sheet Module/ Code Module/ ThisWorkbook Module

ThisWorkbook Module contains Events that are run when the user takes an action in/on the workbook. Our code is stored in ThisWorkbook Module

## Initial interface
![screenshot before](/Screenshot%20(4).png)
## Final interface
![screenshot after](/Screenshot%20(2).png)

## Video to watch
https://drive.google.com/file/d/1Z45YCd-eEwH1AzR3o59wsiavJTMx8C14/view?usp=drive_link

## Code ✍️


``` Function Stock:

Sub stock()

'variables declaration:

Dim i As Double
Dim j As Double
Dim k As Double
Dim x As String

Dim open_value As Double
Dim close_value As Double
Dim total_stock As Double


'initialize a counter to save the position of the beginning of the first ticker

j = 2

'initialize a counter to save the position where we will add the first ticker and its information

k = 2

'save the first open value of the first ticker

open_value = Cells(2, 3).Value

'save the first value of the stock volume

total_stock = Cells(2, 7).Value


For i = j To Cells(Rows.Count, 1).End(xlUp).Row

    If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
    
    'if we still have the same ticker we keep adding the stock volume to the total stock volume of the same ticker
    
        total_stock = total_stock + Cells(i + 1, 7).Value
        
    Else

        'add each ticker to its appropriate cell
        
        Cells(k, 9).Value = Cells(i, 1).Value
             
        'save the close value of that ticker
        
        close_value = Cells(i, 6).Value
        
        'calculate and add the value of the yearly change
        
        Cells(k, 10) = close_value - open_value
        
        'calculate and add the value of the pourcentage of change

        Cells(k, 11).Value = Round((Cells(k, 10).Value / open_value) * 100, 2) & "%"
        
        'save the open value for the next ticker
        
        open_value = Cells(i + 1, 3).Value
        
        'add the total stock volume
        
        Cells(k, 12).Value = total_stock
        
        'initialize the total stock volume for the next ticker by the first value of that ticker
        
        total_stock = Cells(j, 7)
        
       'increment the counter to save the position of the beginning of the next ticker
        
        j = i + 1
        
        'increment the counter to save the position where we will add the next ticker and its information
        
        k = k + 1

    End If
    
Next i

'Call needed subroutines

Call greatest
Call heading
Call coloring
Call sizing

End Sub
```
# Bonus: 🤦‍♀️

- Change the width of created columns for better readability
  
- Write a code for the conditional formatting to apply it to all the other sheets automaticly

# Tip:🪄
To verify if the code is working correctly for the cells Greatest % increase, Greatest % decrease, and Greatest total volume, we can use the functions Min and Max and compare with our coding results :smile:
