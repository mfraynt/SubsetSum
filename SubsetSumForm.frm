VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SubsetSumForm 
   Caption         =   "Solve Subset Sum Problem"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3765
   OleObjectBlob   =   "SubsetSumForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SubsetSumForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    sum = TextBox1.Value
    If sum = "" Then
        MsgBox ("Input target sum as an integer.")
        Exit Sub
    End If
    
    Call SubsetSum(sum)
End Sub


Private Sub Label2_Click()

End Sub
