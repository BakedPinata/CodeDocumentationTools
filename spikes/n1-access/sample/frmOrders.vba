Option Compare Database
Option Explicit

Private Sub Form_Current()
    Dim totals As clsOrderTotals
    Set totals = New clsOrderTotals
    totals.Add 1, Nz(Me.txtTotal, 0)
End Sub

Private Sub cmdLines_Click()
    DoCmd.OpenForm "frmOrderLines", , , "OrderId = " & Nz(Me.txtOrderId, 0)
End Sub
