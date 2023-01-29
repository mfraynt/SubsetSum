Attribute VB_Name = "Subset_sum"
Dim dp(), res As Boolean
Dim i, j, sum As Variant
Dim arr() As Variant
Dim subset As Variant
Dim solution As Range

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

Sub SubsetSum(sum)
    Set subset = CreateObject("Scripting.Dictionary")

    n = Selection.Cells.Count
    ReDim arr(n)
    For i = 0 To UBound(arr) - 1
        arr(i) = Selection(i + 1)
    Next i
    
    recursiveComplexity = 2 ^ n - 1
    dynamicComplexity = (n + 1) * (sum + 1)
    
    If recursiveComplexity > dynamicComplexity Then
        res = isThereSumDynamic(arr, n, sum)
        SubsetSumForm.Label2.Caption = "Algorithm used: Dynamic programming."
    Else:
        res = isThereSumRecursive(arr, n, sum)
        SubsetSumForm.Label2.Caption = "Algorithm used: Recursion."
    End If
    
    If res Then
        MsgBox ("Subset found and will be selected.")
    Else
        MsgBox ("Subset was not found")
    End If
    Unload SubsetSumForm
    
    If res Then
        Set solution = Nothing
        For Each v In subset
            If solution Is Nothing Then
                Set solution = Selection(v + 1)
            Else:
                Set solution = Union(solution, Selection(v + 1))
            End If
            Debug.Print Selection(v + 1)
        Next v
        solution.Select
    End If
    
End Sub

Sub showForm()
    SubsetSumForm.Show
End Sub

