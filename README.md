# 1. How to write a script in VBA to manipulate input data form excel to calculate and display output data :smile:


Hello, in this project we will learn how to use VBA to  loop throught all the stock for one year and output useful data

# 2. what we will learn from this project

    - How to declare and use the right variables
    - How to use the right instructions: input of date, condition, iteration, and output of data
    - How to call predefined functions
    - How to call created function
    - How to apply the same code in several sheets

# 3. Software we need:
MS Excel

# Our Program

Result

![screenshot before](/Screenshot%20(4).png)
![screenshot after](/Screenshot%20(1).png)

```Main function
'initialize a counter to save the position of the bigenning of the first ticker

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

        'add the ticker in the cell(k,9)
        
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
        
       'increment the counter to save the position of the bigenning of the next ticker
        
        j = i + 1
        
        'increment the counter to save the position where we will add the next ticker and its information
        
        k = k + 1

    End If
    
Next i

call greatest```

