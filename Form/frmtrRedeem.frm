VERSION 5.00
Object = "{BB31661F-0587-11D6-9DD0-00C04F0BD97C}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmtrRedeem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Redeem Credited Amount from Customer"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4965
   Icon            =   "frmtrRedeem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjChameleon.chameleonButton cmdtrDo 
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   3240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   8
      TX              =   "Update"
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
      MICON           =   "frmtrRedeem.frx":0ECA
   End
   Begin prjChameleon.chameleonButton cmdClose 
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   3240
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
      MICON           =   "frmtrRedeem.frx":0EE6
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4695
      Begin VB.TextBox txtCredpay 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Top             =   1800
         Width           =   2895
      End
      Begin VB.ComboBox cmbcust 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "Amount Paying:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Amount Credited:  Rupees"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label lblcred 
         Caption         =   "0"
         Height          =   255
         Left            =   2160
         TabIndex        =   3
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Customer Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmtrRedeem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As Database
Dim rs, rs3 As Recordset
Dim credtopay, credpaying As Long
Dim vlflag As Boolean

Private Sub cmbcust_Click()
Dim pr As String
pr = cmbcust.Text
rs.MoveFirst
While Not rs.EOF
    If pr = rs!Name Then
    credtopay = rs!Credit
    lblcred.Caption = credtopay
    End If
    rs.MoveNext
Wend
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdtrDo_Click()
If txtCredpay.Text = "" Then
credpaying = 0
Else
credpaying = txtCredpay.Text
End If
vlflag = False
Validate
If vlflag = False Then
    rs.MoveFirst
    While Not rs.EOF
    If cmbcust.Text = rs!Name Then
    rs.Edit
    rs!Credit = rs!Credit - credpaying
    rs.Update
    End If
    rs.MoveNext
    Wend
    rs3.Edit
    rs3!Cash = rs3!Cash + credpaying
    rs3.Update
    StatRefresh
    MsgBox "Transaction Completed Successfully", vbOKOnly, "OK"
    Unload Me
End If
End Sub

Private Sub Form_Load()
frmMain.Frame2.Visible = False
Set db = DBEngine.Workspaces(0).OpenDatabase(App.Path & "\db2.mdb")
Set rs = db.OpenRecordset("Contacts")
Set rs3 = db.OpenRecordset("Cashrec")
rs.MoveFirst
rs3.MoveFirst
While Not rs.EOF
    cmbcust.AddItem rs!Name
    rs.MoveNext
Wend
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Frame2.Visible = True
End Sub

Private Sub txtCredpay_Change()
Dim dig$, i, digi$, digits$
If txtCredpay.Text <> "" Then
    dig$ = Mid(txtCredpay.Text, Len(txtCredpay.Text), 1)
    If Asc(dig$) < 46 Or Asc(dig$) > 57 Then
        For i = 1 To Len(txtCredpay.Text) - 1
            digi$ = Mid(txtCredpay.Text, i, 1)
            digits$ = digits$ & digi$
        Next i
        txtCredpay.Text = digits$
        txtCredpay.SelStart = Len(txtCredpay.Text)
    End If
End If
End Sub

Private Sub Validate()
If credpaying = 0 Then
MsgBox "Please Enter Amount", vbOKOnly, "Error"
vlflag = True
txtCredpay.SetFocus
End If
If credtopay < credpaying Then
MsgBox "Paying more Cash Than Needed", vbOKOnly, "Error"
vlflag = True
txtCredpay = ""
txtCredpay.SetFocus
End If
End Sub
