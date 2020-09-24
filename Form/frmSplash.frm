VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   Caption         =   "Welcome to EZ Storekeeper"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15240
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmSplash.frx":0ECA
   ScaleHeight     =   11520
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   3360
      TabIndex        =   2
      Top             =   10560
      Width           =   8775
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ZpyderKeeper - Inventory Management System"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   8265
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   7320
      Top             =   5760
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "(  0 %  ) ..."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   7440
      TabIndex        =   1
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Loading"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   6360
      TabIndex        =   0
      Top             =   5280
      Width           =   1095
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mintCount As Integer, mintPause As Integer
Dim onceflag As Boolean

Private Sub Form_Unload(Cancel As Integer)
frmLogIn.Show
End Sub

Private Sub Timer1_Timer()
mintPause = mintPause + 1
If mintCount < 50 Then
    mintCount = mintCount + 1
    Label2.Caption = "(  " & mintCount & "%  )..."
    frmSplash.Refresh
ElseIf mintCount < 100 Then
    mintCount = mintCount + 2
    Label2.Caption = "(  " & mintCount & "%  )..."
    frmSplash.Refresh
End If
If mintPause = 101 Then
        Label2.Caption = "App..."
        Label1.Caption = "Starting"
ElseIf mintPause > 150 Then
         Unload Me
         frmLogIn.Show
End If
End Sub
