VERSION 5.00
Object = "{BB31661F-0587-11D6-9DD0-00C04F0BD97C}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmSrchName 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Customer by Name"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmSrchName.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjChameleon.chameleonButton cmdSearch 
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   8
      TX              =   "Search"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      MPTR            =   0
      MICON           =   "frmSrchName.frx":0ECA
   End
   Begin prjChameleon.chameleonButton cmdClose 
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   2400
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   8
      TX              =   "Close"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      MPTR            =   0
      MICON           =   "frmSrchName.frx":0EE6
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enter Name"
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "Name:"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmSrchName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
frmMain.Frame2.Visible = True
End Sub

Private Sub cmdSearch_Click()
frmResName.Show
End Sub

Private Sub Form_Load()
frmMain.Frame2.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Frame2.Visible = True
End Sub
