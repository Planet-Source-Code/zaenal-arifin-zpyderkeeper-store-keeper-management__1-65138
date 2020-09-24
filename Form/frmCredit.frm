VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{BB31661F-0587-11D6-9DD0-00C04F0BD97C}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmCredit 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8505
   Icon            =   "frmCredit.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmCredit.frx":0ECA
   ScaleHeight     =   6405
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjChameleon.chameleonButton cmdExit 
      Height          =   495
      Left            =   3480
      TabIndex        =   1
      Top             =   5640
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "Exit"
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
      BCOL            =   16744576
      FCOL            =   0
      FCOLO           =   0
      MPTR            =   0
      MICON           =   "frmCredit.frx":184FE
   End
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   8535
      ExtentX         =   15055
      ExtentY         =   8916
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmCredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Job As String

Private Sub brwWebBrowser_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error Resume Next
Dim n As Single
n = Progress * 100 \ ProgressMax
n = n Mod 100

End Sub

Private Sub cmdClose_Click()

Unload Me
brwWebBrowser.Stop


End Sub



Private Sub cmdExit_Click()
Unload Me
brwWebBrowser.Stop

End Sub

Private Sub Form_Resize()
On Error Resume Next
Me.brwWebBrowser.Move 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Job = "kill" Then Exit Sub

On Error Resume Next
Cancel = 1
Me.Hide
frmAbout.SetFocus
End Sub

