Attribute VB_Name = "tooltips"
Option Explicit
' member variable for IsTipVisible property
Public m_IsTipVisible As Boolean

Public Sub ShowTip(tip As String, ctrl As Control)
  With frmmain
    .TipCaption = tip
    .TipTop = ctrl.Top - .tipback.ScaleHeight
    .TipLeft = (ctrl.Left + ctrl.Width) - (.tipback.ScaleWidth / 2)
    .TipVisible = True
    .tipback.ZOrder 0
    m_IsTipVisible = True
  End With
End Sub

Public Sub HideTip()
  With frmmain
    m_IsTipVisible = False
    .TipVisible = False
    .tipback.ZOrder 1
  End With
End Sub
