VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   11760
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option2 
      BackColor       =   &H00404080&
      Caption         =   "OFF"
      Height          =   255
      Left            =   8040
      TabIndex        =   14
      Top             =   5280
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00404080&
      Caption         =   "ON"
      Height          =   255
      Left            =   8040
      TabIndex        =   13
      Top             =   5040
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   8040
      TabIndex        =   11
      Top             =   5880
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   1
      Left            =   6600
      TabIndex        =   6
      Top             =   6120
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   0
      Left            =   6600
      TabIndex        =   5
      Top             =   5280
      Width           =   495
   End
   Begin VB.ListBox List1 
      Height          =   4155
      Left            =   7800
      TabIndex        =   4
      Top             =   600
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C0C0&
      Caption         =   "Receive All Data"
      Height          =   495
      Index           =   0
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6600
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00CB8472&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2760
      TabIndex        =   0
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid MSH1 
      Height          =   4575
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   8070
      _Version        =   393216
      Rows            =   3
      Cols            =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line2 
      X1              =   3000
      X2              =   3000
      Y1              =   5040
      Y2              =   5880
   End
   Begin VB.Shape Shape5 
      Height          =   855
      Left            =   360
      Top             =   5040
      Width           =   5295
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "EVENT-After Edit"
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   18
      Top             =   5160
      Width           =   2415
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "EVENT-Before Edit"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   17
      Top             =   5160
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   16
      Top             =   5520
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackColor       =   &H00404080&
      Caption         =   "CASE SENSITIVE"
      Height          =   255
      Left            =   8760
      TabIndex        =   15
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Line Line1 
      X1              =   7920
      X2              =   11400
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Label Label5 
      BackColor       =   &H00404080&
      Caption         =   "Search "
      Height          =   255
      Left            =   9360
      TabIndex        =   12
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "AUTHOR:VANJA FUCKAR,EMAIL:INGA@VIP.HR"
      Height          =   255
      Left            =   3600
      TabIndex        =   10
      Top             =   7320
      Width           =   3975
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "View Results"
      Height          =   255
      Left            =   9240
      TabIndex        =   9
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Receive Rows Data"
      Height          =   255
      Index           =   1
      Left            =   6120
      TabIndex        =   8
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Receive Cols Data"
      Height          =   255
      Index           =   0
      Left            =   6120
      TabIndex        =   7
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   5520
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   6000
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   6000
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   4575
      Left            =   7680
      Top             =   240
      Width           =   3975
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00404080&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Left            =   7920
      Top             =   4920
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents FLX As EditableGrid
Attribute FLX.VB_VarHelpID = -1
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal _
    hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long
Const LB_SETTABSTOPS = &H192


Private Sub Command1_Click(Index As Integer)
List1.Clear
Dim FF() As String
Select Case Index
 Case 0
 FLX.ReceiveAllData FF, True
For u = LBound(FF, 2) To UBound(FF, 2)
For uu = LBound(FF, 1) To UBound(FF, 1)
List1.AddItem FF(uu, u) & vbTab & "col:" + Str(uu) & vbTab & "row:" + Str(u)
Next uu
Next u
End Select
End Sub

Private Sub FLX_AfterEdit(row As Long, col As Long, NewValue As String)
Label1(1) = Str(row) + " " + Str(col) + " " + NewValue
End Sub

Private Sub FLX_BeforeEdit(row As Long, col As Long, OldValue As String)
Label1(0) = Str(row) + " " + Str(col) + " " + OldValue
End Sub



Private Sub Form_Load()
Set FLX = New EditableGrid
Option1.Value = True

MSH1.Rows = 7
MSH1.Cols = 7

Me.Top = (Screen.Height - Height) / 2
Me.Left = (Screen.Width - Width) / 2

FLX.SetControl MSH1, Text1 'this will able control to be Editable

MSH1.TextMatrix(0, 1) = "Name"
MSH1.TextMatrix(0, 2) = "NickName"
MSH1.TextMatrix(0, 3) = "Sex"
MSH1.TextMatrix(0, 4) = "Height"

For u = 0 To MSH1.Cols - 1
MSH1.ColAlignment(u) = 2
Next u

Dim tabs(1) As Long
tabs(0) = 100
tabs(1) = 100
ListBoxSetTabStops List1, tabs()
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
If Not IsNumeric(Text2(Index)) Then Exit Sub
If Val(Text2(1)) < 0 Or Val(Text2(1)) > MSH1.Rows Then Exit Sub
If Val(Text2(0)) < 0 Or Val(Text2(0)) > MSH1.Cols Then Exit Sub
List1.Clear
Dim FF() As String
Dim x As Long
Select Case Index
Case 0


FLX.ReceiveColsData FF, Val(Text2(0))
x = 0
Do Until x = UBound(FF)
List1.AddItem FF(x)
x = x + 1
Loop
Case 1


FLX.ReceiveRowsData FF, Val(Text2(1))
x = 0
Do Until x = UBound(FF)
List1.AddItem FF(x)
x = x + 1
Loop
 
End Select
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
Dim tempCS As CaseSensitive
Dim row() As Long
Dim col() As Long
Dim x As Long
If Option1.Value = True Then
tempCS = CaseON
Else
tempCS = caseOFF
End If
If KeyAscii = 13 Then
KeyAscii = 0
List1.Clear
FLX.FindText row, col, Text3, tempCS, True
x = 0
Do Until x = UBound(row)
List1.AddItem Text3 & vbTab & "Col:" + Str(col(x)) & vbTab & "Row:" + Str(row(x))
x = x + 1
Loop
End If
End Sub
Sub ListBoxSetTabStops(lb As ListBox, tabStops() As Long)
    Dim numEls As Long
    numEls = UBound(tabStops) - LBound(tabStops) + 1
    SendMessage lb.hwnd, LB_SETTABSTOPS, numEls, tabStops(LBound(tabStops))
End Sub
