VERSION 5.00
Object = "{BB31661F-0587-11D6-9DD0-00C04F0BD97C}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5850
   ClientLeft      =   2040
   ClientTop       =   1380
   ClientWidth     =   7155
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":0ECA
   ScaleHeight     =   5850
   ScaleWidth      =   7155
   StartUpPosition =   2  'CenterScreen
   Begin prjChameleon.chameleonButton cmdExit 
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Top             =   5280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "Exit"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632319
      FCOL            =   0
      FCOLO           =   0
      MPTR            =   0
      MICON           =   "frmAbout.frx":1734A
   End
   Begin prjChameleon.chameleonButton cmdFlash 
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   4680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "Flash"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12648384
      FCOL            =   0
      FCOLO           =   0
      MPTR            =   0
      MICON           =   "frmAbout.frx":17366
   End
   Begin prjChameleon.chameleonButton cmdCredit 
      Height          =   495
      Left            =   2880
      TabIndex        =   4
      Top             =   4080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "Credit"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16761024
      FCOL            =   0
      FCOLO           =   0
      MPTR            =   0
      MICON           =   "frmAbout.frx":17382
   End
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   1080
      Picture         =   "frmAbout.frx":1739E
      ScaleHeight     =   1635
      ScaleWidth      =   4875
      TabIndex        =   1
      Top             =   480
      Width           =   4935
      Begin VB.Line Line3 
         X1              =   720
         X2              =   4200
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   3720
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line1 
         X1              =   480
         X2              =   3960
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   """The Ultimate Retail Store Inventory Solution"""
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   1200
         Width           =   4935
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "ZpyderKeeper"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmAbout.frx":2BE8B
      Height          =   1575
      Left            =   1080
      TabIndex        =   0
      Top             =   2400
      Width           =   4935
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
frmMain.Frame1.Visible = True
Unload Me
End Sub

Private Sub cmdCredit_Click()
    
With frmCredit
    .Caption = "Credit"
    .brwWebBrowser.Navigate App.Path & "\Credit.htm"
    .Show vbModal
End With
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFlash_Click()
frmFlash.Show

End Sub
