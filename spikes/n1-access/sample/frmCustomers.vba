Option Compare Database
Option Explicit

Private Sub Form_Current()
    Me.Caption = "Customer: " & Nz(Me.txtCustomerName, "(new)")
End Sub

Private Sub cboRegion_AfterUpdate()
    Me.Dirty = False
End Sub

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name
End Sub
