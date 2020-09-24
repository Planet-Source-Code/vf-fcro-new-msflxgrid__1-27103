VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EditableGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private WithEvents txtedit As TextBox
Attribute txtedit.VB_VarHelpID = -1
Private WithEvents fg2 As MSFlexGrid
Attribute fg2.VB_VarHelpID = -1

Public Enum CaseSensitive
caseOFF = 1
CaseON = 2
End Enum

Event BeforeEdit(Row As Long, Col As Long, OldValue As String)
Event AfterEdit(Row As Long, Col As Long, NewValue As String)

Private Sub Class_Terminate()
Set txtedit = Nothing
Set fg2 = Nothing
End Sub

Private Sub Fg2_KeyPress(KeyAscii As Integer)
BeforePROC
MSHFlexGridEdit fg2, txtedit, KeyAscii
End Sub
Private Sub Fg2_DblClick()
BeforePROC
MSHFlexGridEdit fg2, txtedit, 32
End Sub
Private Sub Fg2_LeaveCell()
AfterPROC
End Sub
Private Sub Fg2_GotFocus()
AfterPROC
End Sub
Private Sub txtEdit_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc(vbCr) Then KeyAscii = 0
End Sub
Private Sub txtEdit_KeyDown(KeyCode As Integer, _
Shift As Integer)
EditKeyCode fg2, txtedit, KeyCode, Shift
End Sub
Public Sub SetControl(FLX As Object, Txt As Object)
Set fg2 = FLX
Set txtedit = Txt
End Sub
Private Function CalcAlign() As Single
Select Case fg2.ColAlignment(fg2.Col)
Case Is = 0, 1, 2, 9
CalcAlign = 0
Case Is = 3, 4, 5
CalcAlign = 2
Case Is = 6, 7, 8
CalcAlign = 1
End Select
End Function

Private Sub Align()
Dim ClcA As Single
ClcA = CalcAlign
txtedit.Alignment = ClcA
End Sub
Private Sub Fnt()
txtedit.FontBold = fg2.CellFontBold
txtedit.FontItalic = fg2.CellFontItalic
txtedit.FontName = fg2.CellFontName
txtedit.FontSize = fg2.CellFontSize
txtedit.FontStrikethru = fg2.CellFontStrikeThrough
txtedit.FontUnderline = fg2.CellFontUnderline
txtedit.Font.Charset = fg2.Font.Charset
End Sub

Private Sub BeforePROC()
Align
Fnt
TMCLR = txtedit.BackColor
RaiseEvent BeforeEdit(fg2.Row, fg2.Col, fg2)
End Sub
Private Sub AfterPROC()
If txtedit.Visible = False Then Exit Sub
fg2 = txtedit
txtedit.Visible = False
RaiseEvent AfterEdit(fg2.Row, fg2.Col, fg2)
End Sub
Private Sub txtedit_LostFocus()
AfterPROC
End Sub
Public Sub ReceiveAllData(var() As String, Fixed As Boolean)
Dim rw As Long
Dim cl As Long
Dim pr As Long
Dim u As Long
Dim uu As Long
cl = fg2.Cols
rw = fg2.Rows
pr = 0: pg = 0
If Fixed = True Then pr = fg2.FixedRows: pg = fg2.FixedCols
ReDim var(fg2.FixedCols - pg To cl - 1, fg2.FixedRows - pr To rw - 1)
For u = fg2.FixedRows - pr To rw - 1
For uu = fg2.FixedCols - pg To cl - 1
var(uu, u) = fg2.TextMatrix(u, uu)
Next uu
Next u
End Sub
Public Sub ReceiveColsData(var() As String, Row As Long)
Dim u As Long
ReDim var(0)
For u = 0 To fg2.Cols - 1
var(u) = fg2.TextMatrix(Row, u)
ReDim Preserve var(UBound(var) + 1)
Next u
End Sub
Public Sub ReceiveRowsData(var() As String, Col As Long)
Dim u As Long
ReDim var(0)
For u = 0 To fg2.Rows - 1
var(u) = fg2.TextMatrix(u, Col)
ReDim Preserve var(UBound(var) + 1)
Next u
End Sub

Public Sub FindText(Row() As Long, Col() As Long, Txt As String, CaseSensit As CaseSensitive, Fixed As Boolean)
ReDim Row(0)
ReDim Col(0)
If Txt = "" Then Exit Sub
Dim rw As Long
Dim cl As Long
Dim pr As Long
Dim u As Long
Dim uu As Long
Dim xx As Long
cl = fg2.Cols
rw = fg2.Rows
pr = 0: pg = 0
If Fixed = True Then pr = fg2.FixedRows: pg = fg2.FixedCols
xx = 0
For u = fg2.FixedRows - pr To rw - 1
For uu = fg2.FixedCols - pg To cl - 1
If CaseSensit = CaseON Then
If Txt = fg2.TextMatrix(u, uu) Then
Col(xx) = uu: Row(xx) = u
ReDim Preserve Col(UBound(Col) + 1)
ReDim Preserve Row(UBound(Row) + 1)
xx = xx + 1
End If
Else
If UCase(Txt) = UCase(fg2.TextMatrix(u, uu)) Then
Col(xx) = uu: Row(xx) = u
ReDim Preserve Col(UBound(Col) + 1)
ReDim Preserve Row(UBound(Row) + 1)
xx = xx + 1
End If
End If
Next uu
Next u
End Sub
