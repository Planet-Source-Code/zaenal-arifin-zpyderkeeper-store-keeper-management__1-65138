VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BB31661F-0587-11D6-9DD0-00C04F0BD97C}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmResExpiry 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Results of Search by Expiry Date"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5790
   Icon            =   "frmResExpiry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjChameleon.chameleonButton cmdCLose 
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
      MICON           =   "frmResExpiry.frx":0ECA
   End
   Begin VB.TextBox txtargExpiry 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   240
      Width           =   2535
   End
   Begin VB.Frame Frame4 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5535
      Begin VB.TextBox txtExpiry 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txtTotalPrice 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtQuantity 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtUnitPrice 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
      Begin MSComctlLib.ListView lvwProducts 
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
      Begin VB.Label Label13 
         Caption         =   "Expiry Date:"
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Description:"
         Height          =   255
         Left            =   1680
         TabIndex        =   12
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Date Purchased:"
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Total Price:"
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Quantity:"
         Height          =   255
         Left            =   1680
         TabIndex        =   9
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Unit Price:"
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmResExpiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As Database
Dim rs As Recordset
Dim lstmain2 As ListItem
Dim clmhead2 As ColumnHeader
Dim a As Date

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
If (frmSrchExpiry.txtExpiry <> "") Then
a = frmSrchExpiry.txtExpiry
txtargExpiry.Text = a
Else
a = 1 / 1 / 2003
End If
ViewExp
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmSrchExpiry
End Sub

Private Sub lvwProducts_Click()
Set rs = db.OpenRecordset("Products")
While Not rs.EOF
If lvwProducts.SelectedItem.Text = rs!ProductID Then
    txtUnitPrice.Text = rs!UnitPrice
    txtQuantity.Text = rs!Quantity
    txtTotalPrice.Text = rs!TotalPrice
    txtDescription.Text = rs!Description
    txtDate = rs!Date
    txtExpiry = rs!DateExp
    rs.MoveLast
    rs.MoveNext
Else
    rs.MoveNext
End If
Wend
End Sub

Private Sub ViewExp()
Set db = DBEngine.Workspaces(0).OpenDatabase(App.Path & "\db2.mdb")
Set clmhead2 = lvwProducts.ColumnHeaders.Add(, , "Product ID", lvwProducts.Width)
Set rs = db.OpenRecordset("Products")
lvwProducts.ListItems.Clear
    While Not rs.EOF
    If rs!DateExp <= a Then
    Set lstmain2 = lvwProducts.ListItems.Add(, , rs!ProductID)
    End If
    rs.MoveNext
    Wend
    If rs.RecordCount = 0 Then
        MsgBox "There are no records on the database"
    Else
        lvwProducts.Enabled = True
    End If
End Sub
