VERSION 5.00
Begin VB.UserControl XPButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1095
   ScaleHeight     =   510
   ScaleWidth      =   1095
   ToolboxBitmap   =   "XPPress.ctx":0000
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1320
      Top             =   120
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      Height          =   255
      Left            =   3480
      Top             =   1080
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label XPpress 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "XPButton"
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   3480
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   3480
      Top             =   0
      Width           =   615
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   3480
      Top             =   480
      Width           =   615
   End
   Begin VB.Shape Shape5 
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   3480
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Shape11 
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   2160
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "XPButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Public Enum eHoverColor
  hovBlack
  hovBlue
  hovCyan
  hovForest
  hovGreen
  hovMagenta
  hovOrange
  hovPurple
  hovRed
  hovYellow
End Enum
Dim m_CurrPoint As POINTAPI
Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    If Shape1.Visible = True Then
    Call orange
    Shape1.Visible = False
    End If
    Timer1.Enabled = True
End Sub
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
     Call Grey
     Shape11.Visible = True
End Sub
Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
     Call orange
     Shape11.Visible = False
End Sub
Private Sub UserControl_gotFocus()
     Shape2.Visible = True
     Timer1.Enabled = True
End Sub
Private Sub UserControl_LostFocus()
 Call white
 Shape2.Visible = False
End Sub
Private Sub Timer1_Timer()
    GetCursorPos m_CurrPoint
    ScreenToClient hwnd, m_CurrPoint
    If m_CurrPoint.X < UserControl.ScaleLeft Or _
        m_CurrPoint.Y < UserControl.ScaleTop Or _
        m_CurrPoint.X > UserControl.ScaleLeft + UserControl.Width / 15 Or _
        m_CurrPoint.Y > UserControl.ScaleTop + UserControl.Height / 15 Then
        If Shape2.Visible = True Then
        Call blue
        End If
        If Shape2.Visible = False Then
        Call white
        End If
        Timer1.Enabled = False
        Shape1.Visible = True
    End If
        End Sub
Private Sub UserControl_Resize()
    Image1.Width = UserControl.Width
    Image1.Height = UserControl.Height
    XPpress.Left = 0
    XPpress.Top = ((UserControl.Height - XPpress.Height) / 2) + 30
    XPpress.Width = UserControl.Width
      Call white
      Call borderline
      UserControl.BackColor = Ambient.BackColor
      Call white
End Sub
Private Function blue()
    UserControl.Cls
    Call blueouter
    Call blueinner
    Call whiteinnerback
    Call bbottom
    Call borderline
    Line (30, 15)-(UserControl.Width - 30, 15), RGB(206, 231, 255)
    PSet (UserControl.Width - 15, UserControl.Height - 15), UserControl.BackColor
    PSet (0, UserControl.Height - 15), UserControl.BackColor
End Function
Private Function orange()
    UserControl.Cls
    Call orangeouter
    Call orangeinner
    Call whiteinnerback
    Call obottom
    Call borderline
    PSet (UserControl.Width - 15, UserControl.Height - 15), UserControl.BackColor
    PSet (0, UserControl.Height - 15), UserControl.BackColor
End Function
Private Function white()
    UserControl.Cls
    Call whiteback
    Call borderline
    PSet (UserControl.Width - 15, UserControl.Height - 15), UserControl.BackColor
    PSet (0, UserControl.Height - 15), UserControl.BackColor
End Function
Private Function Grey()
    UserControl.Cls
    Call Pressed
    Call borderline
End Function
Private Function whiteback()
Dim i As Integer
Dim j As Integer
Dim k As Integer
j = (UserControl.Height - 30) / (15 * 20)
k = 15
For i = 1 To j
    Line (15, k)-(UserControl.Width - 15, k), &HFFFFFF
    k = k + 15
  Next i
For i = 1 To j
    Line (15, k)-(UserControl.Width - 15, k), &HFCFCFC
    k = k + 15
  Next i
For i = 1 To j
    Line (15, k)-(UserControl.Width - 15, k), &HF9F9F9
    k = k + 15
  Next i
For i = 1 To j
    Line (15, k)-(UserControl.Width - 15, k), &HF6F6F6  '
    k = k + 15
  Next i
For i = 1 To j
    Line (15, k)-(UserControl.Width - 15, k), &HF3F3F3   '
    k = k + 15
  Next i
For i = 1 To j
    Line (15, k)-(UserControl.Width - 15, k), &HF0F0F0  '
    k = k + 15
  Next i
For i = 1 To j
    Line (15, k)-(UserControl.Width - 15, k), &HF0F0F0
    k = k + 15
  Next i
For i = 1 To j
    Line (15, k)-(UserControl.Width - 15, k), &HF0F0F0
    k = k + 15
  Next i
For i = 1 To j
    Line (15, k)-(UserControl.Width - 15, k), &HF0F0F0
    k = k + 15
  Next i
For i = 1 To j
    Line (15, k)-(UserControl.Width - 15, k), &HF0F0F0
    k = k + 15
  Next i
For i = 1 To j
    Line (15, k)-(UserControl.Width - 15, k), &HF0F0F0
    k = k + 15
  Next i
For i = 1 To j
    Line (15, k)-(UserControl.Width - 15, k), &HF0F0F0
    k = k + 15
  Next i
For i = 1 To j
    Line (15, k)-(UserControl.Width - 15, k), &HF0F0F0
    k = k + 15
  Next i
For i = 1 To j
    Line (15, k)-(UserControl.Width - 15, k), &HF0F0F0
    k = k + 15
  Next i
For i = 1 To j
    Line (15, k)-(UserControl.Width - 15, k), &HF0F0F0
    k = k + 15
  Next i
For i = 1 To j
    Line (15, k)-(UserControl.Width - 15, k), &HF0F0F0
    k = k + 15
  Next i
For i = 1 To j
    Line (15, k)-(UserControl.Width - 15, k), &HF0F0F0
    k = k + 15
  Next i
For i = 1 To j
    Line (15, k)-(UserControl.Width - 15, k), &HF0F0F0
    k = k + 15
  Next i
For i = 1 To j
    Line (15, k)-(UserControl.Width - 15, k), &HF0F0F0
    k = k + 15
  Next i
For i = 1 To j
    Line (15, k)-(UserControl.Width - 15, k), &HF0F0F0
    k = k + 15
  Next i
For i = 1 To j
    Line (15, k)-(UserControl.Width - 15, UserControl.Height - 15), &HF0F0F0, BF
    k = k + 15
  Next i
      Line (15, UserControl.Height - 30)-(UserControl.Width - 15, UserControl.Height - 30), RGB(214, 208, 197)
      Line (15, UserControl.Height - 45)-(UserControl.Width - 15, UserControl.Height - 45), RGB(226, 223, 214)

End Function
Private Function borderline()
    Line (0, 30)-(0, UserControl.Height - 30), &H6A4206
    Line (30, 0)-(UserControl.Width - 30, 0), &H6A4206
    Line (30, UserControl.Height - 15)-(UserControl.Width - 30, UserControl.Height - 15), &H6A4206
    Line (UserControl.Width - 15, 30)-(UserControl.Width - 15, UserControl.Height - 30), &H6A4206
    PSet (9, 9), RGB(122, 149, 168)
    PSet (8, 9), RGB(122, 149, 168)
    PSet (9, 8), &H6A4206
    PSet (7, 9), RGB(122, 149, 168)
    PSet (9, 7), RGB(122, 149, 168)
    PSet (UserControl.Width - 30, 9), &H6A4206
    PSet (UserControl.Width - 30, 7), RGB(122, 149, 168)
    PSet (UserControl.Width - 20, 14), RGB(122, 149, 168)
    PSet (9, UserControl.Height - 30), &H6A4206
    PSet (7, UserControl.Height - 30), RGB(122, 149, 168)
    PSet (14, UserControl.Height - 20), RGB(122, 149, 168)
    PSet (UserControl.Width - 30, UserControl.Height - 30), &H6A4206
    PSet (UserControl.Width - 20, UserControl.Height - 30), RGB(122, 149, 168)
    PSet (UserControl.Width - 30, UserControl.Height - 20), RGB(122, 149, 168)
End Function
Private Function orangeouter()
Dim L As Integer
Dim m As Integer
Dim n As Integer
m = (UserControl.Height - 30) / (15 * 21)
n = 15
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(254, 223, 154)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(253, 220, 147)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(252, 217, 140)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(251, 214, 133)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(250, 211, 126)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(249, 208, 119)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(248, 205, 112)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(247, 202, 105)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(246, 199, 98)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(245, 196, 91)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(244, 193, 84)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(243, 190, 77)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(242, 187, 70)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(241, 184, 63)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(240, 181, 56)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(239, 178, 49)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(238, 175, 42)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(238, 175, 42)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(238, 175, 42)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(238, 175, 42)    '
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(238, 175, 42)
    n = n + 15
  Next L
    Line (15, n)-(UserControl.Width - 15, UserControl.Height), RGB(238, 175, 42), BF
    Line (0, 15)-(UserControl.Width - 15, 15), RGB(255, 240, 207)
End Function
Private Function orangeinner()
Dim L As Integer
Dim m As Integer
Dim n As Integer
m = (UserControl.Height - 60) / (15 * 20)
n = 30
For L = 1 To m
    Line (30, n)-(UserControl.Width - 30, n), RGB(253, 216, 137)  '
    n = n + 15
  Next L
For L = 1 To m
    Line (30, n)-(UserControl.Width - 30, n), RGB(252, 210, 121)     '
    n = n + 15
  Next L
For L = 1 To m
    Line (30, n)-(UserControl.Width - 30, n), RGB(251, 206, 113)
    n = n + 15
  Next L
For L = 1 To m
    Line (30, n)-(UserControl.Width - 30, n), RGB(250, 202, 105)
    n = n + 15
  Next L
For L = 1 To m
    Line (30, n)-(UserControl.Width - 30, n), RGB(249, 198, 97)
    n = n + 15
  Next L
For L = 1 To m
    Line (30, n)-(UserControl.Width - 30, n), RGB(248, 194, 89)
    n = n + 15
  Next L
For L = 1 To m
    Line (30, n)-(UserControl.Width - 30, n), RGB(247, 190, 81)
    n = n + 15
  Next L
For L = 1 To m
    Line (30, n)-(UserControl.Width - 30, n), RGB(246, 186, 73)
    n = n + 15
  Next L
For L = 1 To m
    Line (30, n)-(UserControl.Width - 30, n), RGB(245, 182, 65)
    n = n + 15
  Next L
For L = 1 To m
    Line (30, n)-(UserControl.Width - 30, n), RGB(244, 178, 57)
    n = n + 15
  Next L
For L = 1 To m
    Line (30, n)-(UserControl.Width - 30, n), RGB(243, 174, 49)
    n = n + 15
  Next L
      Line (15, UserControl.Height - 30)-(UserControl.Width - 15, UserControl.Height - 30), RGB(229, 151, 0)
End Function
Private Function whiteinnerback()
Dim i As Integer
Dim j As Integer
Dim k As Integer
j = (UserControl.Height - 90) / (15 * 20)
k = 45
For i = 1 To j
    Line (45, k)-(UserControl.Width - 45, k), &HFFFFFF '
    k = k + 15
  Next i
For i = 1 To j
    Line (45, k)-(UserControl.Width - 45, k), &HFCFCFC   '
    k = k + 15
  Next i
For i = 1 To j
    Line (45, k)-(UserControl.Width - 45, k), &HF9F9F9      '
    k = k + 15
  Next i
For i = 1 To j
    Line (45, k)-(UserControl.Width - 45, k), &HF6F6F6  '
    k = k + 15
  Next i
For i = 1 To j
    Line (45, k)-(UserControl.Width - 45, k), &HF3F3F3   '
    k = k + 15
  Next i
For i = 1 To j
    Line (45, k)-(UserControl.Width - 45, k), &HF0F0F0  '
    k = k + 15
  Next i
For i = 1 To j
    Line (45, k)-(UserControl.Width - 45, k), &HF0F0F0
    k = k + 15
  Next i
For i = 1 To j
    Line (45, k)-(UserControl.Width - 45, k), &HF0F0F0
    k = k + 15
  Next i
For i = 1 To j
    Line (45, k)-(UserControl.Width - 45, k), &HF0F0F0
    k = k + 15
  Next i
For i = 1 To j
    Line (45, k)-(UserControl.Width - 45, k), &HF0F0F0
    k = k + 15
  Next i
For i = 1 To j
    Line (45, k)-(UserControl.Width - 45, k), &HF0F0F0
    k = k + 15
  Next i
For i = 1 To j
    Line (45, k)-(UserControl.Width - 45, k), &HF0F0F0
    k = k + 15
  Next i
For i = 1 To j
    Line (45, k)-(UserControl.Width - 45, k), &HF0F0F0
    k = k + 15
  Next i
For i = 1 To j
    Line (45, k)-(UserControl.Width - 45, k), &HF0F0F0
    k = k + 15
  Next i
For i = 1 To j
    Line (45, k)-(UserControl.Width - 45, k), &HF0F0F0
    k = k + 15
  Next i
For i = 1 To j
    Line (45, k)-(UserControl.Width - 45, k), &HF0F0F0
    k = k + 15
  Next i
For i = 1 To j
    Line (45, k)-(UserControl.Width - 45, k), &HF0F0F0
    k = k + 15
  Next i
For i = 1 To j
    Line (45, k)-(UserControl.Width - 45, k), &HF0F0F0
    k = k + 15
  Next i
For i = 1 To j
    Line (45, k)-(UserControl.Width - 45, k), &HF0F0F0
    k = k + 15
  Next i
For i = 1 To j
    Line (45, k)-(UserControl.Width - 45, k), &HF0F0F0
    k = k + 15
  Next i
For i = 1 To j
    Line (45, k)-(UserControl.Width - 60, _
    UserControl.Height - 60), &HF0F0F0, BF
 Next i
End Function
Private Function bbottom() As Variant
      Line (30, UserControl.Height - 45)-(UserControl.Width - 30, UserControl.Height - 45), RGB(110, 154, 227)
      Line (30, UserControl.Height - 30)-(UserControl.Width - 30, UserControl.Height - 30), RGB(97, 125, 229)
End Function
Private Function obottom() As Variant
      Line (30, UserControl.Height - 45)-(UserControl.Width - 30, UserControl.Height - 45), RGB(243, 174, 49)
      Line (30, UserControl.Height - 30)-(UserControl.Width - 30, UserControl.Height - 30), RGB(229, 151, 0)
End Function
Private Function blueouter() As Variant
Dim L As Integer
Dim m As Integer
Dim n As Integer
m = (UserControl.Height - 30) / (15 * 20)
n = 15
Line (15, 30)-(UserControl.Width - 15, 30), RGB(206, 231, 251)
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(188, 212, 246)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(186, 211, 246)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(182, 208, 245)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(178, 205, 244)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(174, 202, 243)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(170, 199, 242)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(166, 196, 241)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(162, 193, 240)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(158, 190, 239)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(154, 187, 238)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(150, 184, 237)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(146, 181, 236)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(142, 178, 235)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(138, 175, 234)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(134, 172, 233)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(130, 169, 232)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(126, 166, 231)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(122, 163, 230)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(118, 160, 229)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(114, 157, 228)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(110, 154, 227)
    n = n + 15
  Next L
    Line (15, n)-(UserControl.Width - 15, UserControl.Height), RGB(110, 154, 227)

End Function
Private Function blueinner()
Dim L As Integer
Dim m As Integer
Dim n As Integer
m = (UserControl.Height - 60) / (15 * 20)
n = 15
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(188, 212, 246)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(186, 211, 246)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(182, 208, 245)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(178, 205, 244)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(174, 202, 243)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(170, 199, 242)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(166, 196, 241)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(162, 193, 240)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(158, 190, 239)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(154, 187, 238)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(150, 184, 237)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(146, 181, 236)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(142, 178, 235)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(138, 175, 234)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(134, 172, 233)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(115, 169, 232)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(126, 166, 231)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(122, 163, 215)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(118, 160, 229)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(114, 157, 228)
    n = n + 15
  Next L
For L = 1 To m
    Line (15, n)-(UserControl.Width - 15, n), RGB(110, 154, 227)
    n = n + 15
  Next L
      Line (15, n - 15)-(UserControl.Width, UserControl.Height), RGB(110, 154, 227), BF
      Line (0, n)-(UserControl.Width, UserControl.Height), RGB(110, 154, 227), BF
End Function
Private Function Pressed() As Variant
Dim i As Single
Dim k As Single
k = 15
For i = 1 To UserControl.Height
    Line (15, k)-(UserControl.Width - 15, k), RGB(225, 224, 217)
    k = k + 15
  Next i
End Function
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = Image1.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    Image1.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = XPpress.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    Set XPpress.Font = New_Font
    PropertyChanged "Font"
End Property
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub
Private Sub Image1_Click()
    RaiseEvent Click
End Sub
Private Sub Image1_DblClick()
    RaiseEvent DblClick
End Sub
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = XPpress.Caption
End Property
Public Property Let Caption(ByVal New_Caption As String)
    XPpress.Caption() = New_Caption
    PropertyChanged "Caption"
End Property
Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font that appears in each row for the given level."
    FontName = XPpress.FontName
End Property
Public Property Let FontName(ByVal New_FontName As String)
    XPpress.FontName() = New_FontName
    PropertyChanged "FontName"
End Property
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
    FontSize = XPpress.FontSize
End Property
Public Property Let FontSize(ByVal New_FontSize As Single)
    XPpress.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets underline font styles."
    FontUnderline = XPpress.FontUnderline
End Property
Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    XPpress.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Image1.Enabled = PropBag.ReadProperty("Enabled", True)
    Set XPpress.Font = PropBag.ReadProperty("Font", Ambient.Font)
    XPpress.Caption = PropBag.ReadProperty("Caption", "XPpress")
    XPpress.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", Image1.Enabled, True)
    Call PropBag.WriteProperty("Font", XPpress.Font, Ambient.Font)
    Call PropBag.WriteProperty("Caption", XPpress.Caption, "XPpress")
    Call PropBag.WriteProperty("FontName", XPpress.FontName, "")
    Call PropBag.WriteProperty("FontSize", XPpress.FontSize, 0)
    Call PropBag.WriteProperty("FontUnderline", XPpress.FontUnderline, 0)
End Sub
