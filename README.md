# Subset sum problem - Implementation for MS Excel
![Excel meme](/assets/img/Excel_meme.jpg)
## Problem setup
---
[Subset sum problem](https://en.wikipedia.org/wiki/Subset_sum_problem) itself is quite famous and important for many applications. I personally faced it in the following case: given a list of wagons with weight of loaded cargo, is there a set of wagons with total weight of cargo of a given amount. 

Main differences from the classical formulation are that the weights may be (and mostly are) non-integer numbers. However this is quite easily solvable by multiplying weights and target sum by $ 10^n $, where *n* is a number of deciumal figures. 

---

## Solution

Basically there are 2 main algorithms used for solution of the Subset sum problem, although recently some improvements have been made, which I'll probably review in the future. For now we implement the following:
* Recursion
* Dynamic programming
>It is worth mentioning that we not only have to solve the SSP, but also need to return the subset itself. 

Both algorithms have same input data:   
$ arr $ - the array of integers;  
$ n $ - size of the array;  
$ sum $ - target sum as an integer.

### Recursion 

```Mermaid
graph LR
Root("F(arr, n, sum)")
Root --> A(Is sum = 0)
A --> |True| B>Retrun True]
A --> |False| C(Is n = 0)
C --> |True| D>Return False]
C --> |Flase| E("n = n - 1")
C --> |False| F("n = n - 1; sum = sum - arr(n-1)")
E --> Root
F --> Root

```
In order to return the subset, we have to record an element of $ arr $ echa time we decrease the $ sum $. VBA implementation is provided below.

```VB
Function isThereSumRecursive(arr, n, sum) As Boolean
    If sum = 0 Then
        isThereSumRecursive = True
        Exit Function
    End If
    If n = 0 Then
        isThereSumRecursive = False
        Exit Function
    End If
    
    isThereSumRecursive = isThereSumRecursive(arr, n - 1, sum) Or isThereSumRecursive(arr, n - 1, sum - arr(n - 1))
    
    ' Adding elements to subset
    If isThereSumRecursive(arr, n - 1, sum - arr(n - 1)) Then
        If Not subset.Exists(n - 1) Then
            subset.Add n - 1, arr(n - 1)
        End If
    End If
       
End Function
```

### Dynamic programming

Based on input data build the matrix $dp$ of size $ (n+1) \times (sum+1) $ as follows: 
 
$ dp[i, 0] = True, i \in [0;...; n]; $ <br>
$ dp[0, j] = False, j \in [0; ...; sum]; $  <br>

$ \begin{cases} 
dp[i,j] = dp[i-1, j] \\
j < arr(i-1)
\end{cases}$

$ \begin{cases} 
dp[i,j] = (dp[i-1, j] \; \lor \; dp[i-1, j - arr(i-1)]) \\
j \ge arr(i-1)
\end{cases}$

In order to return the subset, we have to walk through the $ dp $ matrix top-down and find $ dp[i,j] $ such as:  

$
\begin{cases} 
    dp[i,j] = True;\\
    dp[i-1,j] = False.
\end{cases}
 $  

VBA implementation is provided below.

```VB
Function isThereSumDynamic(arr, n, sum) As Boolean
    ReDim dp(n + 1, sum + 1)
    
    For i = 0 To n
        dp(i, 0) = True
    Next i
    For i = 1 To sum
        dp(0, i) = False
    Next i
    
    For i = 1 To n
        For j = 1 To sum
            If j < arr(i - 1) Then
                dp(i, j) = dp(i - 1, j)
            End If
            If j >= arr(i - 1) Then
                dp(i, j) = dp(i - 1, j) Or dp(i - 1, j - arr(i - 1))
            End If
        Next j
    Next i
    
    isThereSumDynamic = dp(n, sum)
    
    ' Adding elements to subset
    If isThereSumDynamic Then
        i = n
        j = sum
        While j > 0
            If dp(i, j) And dp(i - 1, j) Then
                i = i - 1
            End If
            If dp(i, j) And Not dp(i - 1, j) Then
                subset.Add i - 1, arr(i - 1)
                j = j - arr(i - 1)
                i = i - 1
            End If
        Wend
        
    End If
    
End Function
```