VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BB31661F-0587-11D6-9DD0-00C04F0BD97C}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmResName 
   Caption         =   "Results of Search by Name"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5775
   Icon            =   "frmResName.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjChameleon.chameleonButton cmdClose 
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   15
      Top             =   3720
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
      MICON           =   "frmResName.frx":0ECA
   End
   Begin VB.TextBox txtargName 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   240
      Width           =   2535
   End
   Begin VB.Frame Frame3 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5535
      Begin VB.TextBox txtCName 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txtCAddress 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtCPhone 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox txtCCel 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox txtCredit 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   2040
         Width           =   2535
      End
      Begin MSComctlLib.ListView lvwContacts 
         Height          =   2175
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   3836
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         BorderStyle     =   1
         Appearance      =   0
         Enabled         =   0   'False
         NumItems        =   0
      End
      Begin VB.Label Label2 
         Caption         =   "Name:"
         Height          =   255
         Left            =   1800
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Address:"
         Height          =   255
         Left            =   1800
         TabIndex        =   12
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Phone No:"
         Height          =   255
         Left            =   1800
         TabIndex        =   11
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Cell No:"
         Height          =   255
         Left            =   1800
         TabIndex        =   10
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Email:"
         Height          =   255
         Left            =   1800
         TabIndex        =   9
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Credit:"
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   2040
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmResName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As Database
Dim rs As Recordset
Dim lstmain1 As ListItem
Dim clmhead1 As ColumnHeader
Dim a As String

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
a = frmSrchName.txtName
txtargName.Text = a
Viewname
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmSrchName
End Sub

Private Sub lvwContacts_click()
Set rs = db.OpenRecordset("Contacts")
While Not rs.EOF
If lvwContacts.SelectedItem.Text = rs!ContactID Then
    txtCName.Text = rs!Name
    txtCAddress.Text = rs!Address
    txtCPhone.Text = rs!PhoneNo
    txtCCel.Text = rs!Celno
    txtEmail = rs!Email
    txtCredit = rs!Credit
    rs.MoveLast
    rs.MoveNext
Else
    rs.MoveNext
End If
Wend
End Sub

Private Sub Viewname()
Set db = DBEngine.Workspaces(0).OpenDatabase(App.Path & "\db2.mdb")
Set clmhead1 = lvwContacts.ColumnHeaders.Add(, , "Contact ID", lvwContacts.Width)
Set rs = db.OpenRecordset("Contacts")
lvwContacts.ListItems.Clear
    While Not rs.EOF
    If InStr(1, rs!Name, a) <> 0 Then
    Set lstmain1 = lvwContacts.ListItems.Add(, , rs!ContactID)
    End If
    rs.MoveNext
    Wend
    If rs.RecordCount = 0 Then
        MsgBox "There are no records on the database"
    Else
        lvwContacts.Enabled = True
    End If
End Sub



