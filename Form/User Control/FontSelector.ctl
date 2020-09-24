VERSION 5.00
Begin VB.UserControl FontSelector 
   BackStyle       =   0  '³z©ú
   ClientHeight    =   5265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6765
   ScaleHeight     =   5265
   ScaleWidth      =   6765
   Begin VB.PictureBox picDrop 
      Appearance      =   0  '¥­­±
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   0
      ScaleHeight     =   271
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   273
      TabIndex        =   1
      Top             =   300
      Visible         =   0   'False
      Width           =   4125
      Begin VB.VScrollBar Scr 
         Height          =   4095
         LargeChange     =   12
         Left            =   3840
         TabIndex        =   14
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblFont 
         Appearance      =   0  '¥­­±
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   11
         Left            =   480
         TabIndex        =   13
         Top             =   3720
         Width           =   3615
      End
      Begin VB.Label lblFont 
         Appearance      =   0  '¥­­±
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   10
         Left            =   480
         TabIndex        =   12
         Top             =   3360
         Width           =   3615
      End
      Begin VB.Label lblFont 
         Appearance      =   0  '¥­­±
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   9
         Left            =   480
         TabIndex        =   11
         Top             =   3000
         Width           =   3615
      End
      Begin VB.Label lblFont 
         Appearance      =   0  '¥­­±
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   8
         Left            =   480
         TabIndex        =   10
         Top             =   2640
         Width           =   3615
      End
      Begin VB.Label lblFont 
         Appearance      =   0  '¥­­±
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   7
         Left            =   480
         TabIndex        =   9
         Top             =   2280
         Width           =   3615
      End
      Begin VB.Label lblFont 
         Appearance      =   0  '¥­­±
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   6
         Left            =   480
         TabIndex        =   8
         Top             =   1920
         Width           =   3615
      End
      Begin VB.Label lblFont 
         Appearance      =   0  '¥­­±
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   5
         Left            =   480
         TabIndex        =   7
         Top             =   1560
         Width           =   3615
      End
      Begin VB.Label lblFont 
         Appearance      =   0  '¥­­±
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   4
         Left            =   480
         TabIndex        =   6
         Top             =   1200
         Width           =   3615
      End
      Begin VB.Label lblFont 
         Appearance      =   0  '¥­­±
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   3
         Left            =   360
         TabIndex        =   5
         Top             =   960
         Width           =   3615
      End
      Begin VB.Label lblFont 
         Appearance      =   0  '¥­­±
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   480
         TabIndex        =   4
         Top             =   600
         Width           =   3615
      End
      Begin VB.Label lblFont 
         Appearance      =   0  '¥­­±
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label lblFont 
         Appearance      =   0  '¥­­±
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   0
         Width           =   3615
      End
   End
   Begin VB.ComboBox Cob 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      Style           =   2  '³æ¯Â¤U©Ô¦¡
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "FontSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private K() As cControlFlater, i As Long
Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long

Public Drop As Boolean
Dim Ready As Boolean
Dim ActualWidth As Integer
Dim CurrentFocus As Integer

Dim CurrentTop As Integer

Public Event Click()

Public Sub MakeMeFlat()

    ReDim Preserve K(0 To 0)

            On Error Resume Next
            Set K(0) = New cControlFlater
            K(0).Attach Cob

End Sub

Private Sub Cob_DropDown()
Drop = True
ActualWidth = UserControl.Width
UserControl.Width = picDrop.Width
UserControl.Height = picDrop.Top + picDrop.Height
If Ready = False Then MakeFont
'picDrop.SetFocus
picDrop.Visible = True
SetFocus picDrop.hwnd
End Sub

Sub MakeFont()

picDrop.Height = lblFont(0).Height * 12 * Screen.TwipsPerPixelY

Scr.Min = 0
Scr.Max = Screen.FontCount - 12

Dim scrWidth As Integer

scrWidth = Scr.Width

Scr.Move picDrop.ScaleWidth - Scr.Width, 0, Scr.Width, picDrop.ScaleHeight

For i = 0 To 11
    lblFont(i).Move 0, i * 20, picDrop.ScaleWidth - scrWidth
Next i

PrintFont

End Sub

Sub MakeText()
For i = 0 To Screen.FontCount - 1
    Cob.AddItem Screen.Fonts(i)
Next i
End Sub

Sub PrintFont()

For i = 0 To 11
    lblFont(i).FontName = Screen.Fonts(i + CurrentTop)
    lblFont(i).FontBold = False
    lblFont(i).FontItalic = False
    lblFont(i).FontSize = 12
    lblFont(i).Caption = Screen.Fonts(i + CurrentTop)
Next i

End Sub

Private Sub lblFont_Click(Index As Integer)

Cob.Text = Screen.Fonts(Index + CurrentTop)

Me.UserControl_ExitFocus

RaiseEvent Click

End Sub

Public Property Get Text() As String
Text = Cob.Text
End Property

Public Property Let Text(str As String)
Cob.Text = str
End Property

Private Sub lblFont_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Index <> CurrentFocus Then
    
    For i = 0 To 11
        
        If i <> Index Then
            lblFont(i).BackColor = vbWhite
            lblFont(i).ForeColor = vbBlack
        End If
        
    Next
    
    lblFont(Index).BackColor = &H8000000D
    lblFont(Index).ForeColor = vbWhite
    
End If

CurrentFocus = Index

End Sub

Private Sub picDrop_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
Drop = False
UserControl.Width = ActualWidth
UserControl_Resize
End If
End Sub

Private Sub Scr_Change()

CurrentTop = Scr.Value

PrintFont

picDrop.SetFocus
picDrop.Refresh
End Sub

Private Sub Scr_Scroll()
DoEvents
Scr_Change
End Sub

Public Sub UserControl_ExitFocus()
Drop = False
If ActualWidth <> 0 Then UserControl.Width = ActualWidth
UserControl_Resize
End Sub

Private Sub UserControl_Initialize()
CurrentFocus = -1
MakeText
End Sub

Private Sub UserControl_Resize()
If Drop = False Then
    UserControl.Height = 315 'Cob.Height
    picDrop.Visible = False
    Cob.Move 0, 0, UserControl.Width
End If
End Sub

Public Property Get hwnd() As Long
hwnd = UserControl.hwnd
End Property
