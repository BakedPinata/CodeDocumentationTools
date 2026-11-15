Option Compare Database
Option Explicit

' The start-up form. Its Load event leaves a mark, so it's plain whether it ran.
Private Sub Form_Load()
    LogStartup "frmStart"
End Sub

Private Sub cmdCustomers_Click()
    OpenCustomers
End Sub
