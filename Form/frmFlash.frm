VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{BB31661F-0587-11D6-9DD0-00C04F0BD97C}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmFlash 
   BorderStyle     =   0  'None
   Caption         =   "Flash Movie"
   ClientHeight    =   6690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8550
   Icon            =   "frmFlash.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmFlash.frx":0ECA
   ScaleHeight     =   6690
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjChameleon.chameleonButton cmdStop 
      Height          =   375
      Left            =   5880
      TabIndex        =   2
      Top             =   6120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Stop 'n Exit"
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
      MICON           =   "frmFlash.frx":16E14
   End
   Begin prjChameleon.chameleonButton cmdStart 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   6120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Start"
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
      MICON           =   "frmFlash.frx":16E30
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   5775
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      _cx             =   14208
      _cy             =   10186
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
   End
End
Attribute VB_Name = "frmFlash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStart_Click()
ShockwaveFlash1.Movie = App.Path & "/fahrschule.swf"
ShockwaveFlash1.Play

End Sub

Private Sub cmdStop_Click()
ShockwaveFlash1.Stop
 Unload Me
End Sub

Private Sub ShockwaveFlash1_OnReadyStateChange(newState As Long)
ShockwaveFlash1.Movie = App.Path & "/fahrschule.swf"
End Sub

