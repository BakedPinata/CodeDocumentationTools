VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDemoForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4304
   ClientLeft      =   96
   ClientTop       =   416
   ClientWidth     =   5392
   OleObjectBlob   =   "frmDemoForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDemoForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm1_Initialize1()
    Dim dummyData() As Variant
    
    ' Define a 3-row, 2-column array of dummy data
    dummyData = Array( _
        Array("Item 1", "Category A"), _
        Array("Item 2", "Category B"), _
        Array("Item 3", "Category C") _
    )
    
    With Me.lbxDemoListBox
        .ColumnCount = 2
        .List = dummyData
    End With
End Sub

Private Sub UserForm_Initialize()
    Dim Data(1 To 3, 1 To 2) As Variant
    
    ' Populate the array with dummy values
    Data(1, 1) = "Alpha": Data(1, 2) = 100
    Data(2, 1) = "Beta":  Data(2, 2) = 200
    Data(3, 1) = "Gamma": Data(3, 2) = 300
    
    With Me.lbxDemoListBox
        .ColumnCount = 2
        .List = Data
    End With
End Sub
