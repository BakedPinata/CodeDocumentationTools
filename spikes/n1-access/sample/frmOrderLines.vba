Option Compare Database
Option Explicit

Private Sub txtQuantity_AfterUpdate()
    Me.txtLineTotal = LineTotal(Nz(Me.txtQuantity, 0), Nz(Me.txtPrice, 0))
End Sub
