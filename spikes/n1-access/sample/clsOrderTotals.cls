Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' Adds up an order's lines.

Private m_total As Currency

Public Sub Add(ByVal quantity As Long, ByVal price As Currency)
    m_total = m_total + LineTotal(quantity, price)
End Sub

Public Property Get Total() As Currency
    Total = m_total
End Property
