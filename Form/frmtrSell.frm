VERSION 5.00
Object = "{BB31661F-0587-11D6-9DD0-00C04F0BD97C}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmtrSell 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sell Product to Customer"
   ClientHeight    =   4710
   ClientLeft      =   2925
   ClientTop       =   2520
   ClientWidth     =   5925
   Icon            =   "frmtrSell.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Begin prjChameleon.chameleonButton cmdtrDo 
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   15
      Top             =   4080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   8
      TX              =   "Sell"
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
      MICON           =   "frmtrSell.frx":0ECA
   End
   Begin prjChameleon.chameleonButton cmdCLose 
      Height          =   375
      Left            =   4680
      TabIndex        =   14
      Top             =   4080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   8
      TX              =   "CLose"
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
      MICON           =   "frmtrSell.frx":0EE6
   End
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   5655
      Begin VB.TextBox txtAmt 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2160
         TabIndex        =   2
         Top             =   1920
         Width           =   2895
      End
      Begin VB.TextBox txtCash 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2160
         TabIndex        =   3
         Top             =   2880
         Width           =   2895
      End
      Begin VB.ComboBox cmbprod 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "frmtrSell.frx":0F02
         Left            =   2160
         List            =   "frmtrSell.frx":0F04
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   960
         Width           =   2895
      End
      Begin VB.ComboBox cmbcust 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   2160
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label lblPrice 
         Caption         =   "0"
         Height          =   255
         Left            =   2760
         TabIndex        =   13
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Price per Unit:                    Rupees"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "Money Paying in Cash:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label lblMoney 
         Caption         =   "0"
         Height          =   255
         Left            =   2760
         TabIndex        =   9
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Money to be paid:              Rupees"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2400
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "Amount Buying:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Product Description:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Customer Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Label Label6 
      Caption         =   "Money to be paid:              Rupees"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   1680
      Width           =   2535
   End
End
Attribute VB_Name = "frmtrSell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As Database
Dim rs As Recordset
Dim rs2 As Recordset
Dim rs3 As Recordset
Dim cashed, credited, credited2 As Long
Dim vlflag As Boolean

Private Sub cmbprod_Click()
Dim pr As String
pr = cmbprod.Text
rs2.MoveFirst
While Not rs2.EOF
    If pr = rs2!Description Then
    lblPrice.Caption = rs2!UnitPrice
    If txtAmt.Text <> "" Then lblMoney.Caption = rs2!UnitPrice * CInt(txtAmt.Text)
    End If
    rs2.MoveNext
Wend
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdtrDo_Click()
vlflag = False
Validate
If Not vlflag Then
Finalupdate
StatRefresh
MsgBox "Transaction Completed Successfully.", vbOKOnly, "OK"
Unload Me
End If
End Sub

Private Sub Form_Load()
'vlflag = False
frmMain.Frame2.Visible = False
Set db = DBEngine.Workspaces(0).OpenDatabase(App.Path & "\db2.mdb")
Set rs = db.OpenRecordset("Contacts")
rs.MoveFirst
While Not rs.EOF
    cmbcust.AddItem rs!Name
    rs.MoveNext
Wend
cmbcust.AddItem "xxx"
Set rs2 = db.OpenRecordset("Products")
rs2.MoveFirst
While Not rs2.EOF
    cmbprod.AddItem rs2!Description
    rs2.MoveNext
Wend
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Frame2.Visible = True
End Sub

Private Sub txtAmt_Change()
Dim dig$, i, digi$, digits$
If txtAmt.Text <> "" Then
    dig$ = Mid(txtAmt.Text, Len(txtAmt.Text), 1)
    If Asc(dig$) < 46 Or Asc(dig$) > 57 Then
        For i = 1 To Len(txtAmt.Text) - 1
            digi$ = Mid(txtAmt.Text, i, 1)
            digits$ = digits$ & digi$
        Next i
        txtAmt.Text = digits$
        txtAmt.SelStart = Len(txtAmt.Text)
    End If
End If
Amtupdate
End Sub

Private Sub Amtupdate()
Dim a As Long, B As Long
If txtAmt.Text = "" Or lblPrice.Caption = "0" Then
    lblMoney.Caption = "0"
Else
a = CLng(lblPrice.Caption)
B = CLng(txtAmt.Text)
lblMoney.Caption = a * B
End If
End Sub

Private Sub Validate()
credited2 = 0
Dim custflag As Boolean
Dim prodflag As Boolean
Dim amtt As Long
custflag = False
prodflag = False



Set rs = db.OpenRecordset("Contacts")
Set rs2 = db.OpenRecordset("Products")

If txtCash.Text = "" Then
cashed = 0
Else
cashed = CLng(txtCash.Text)
End If
If lblMoney.Caption = "" Then
credited = 0
Else
credited = CLng(lblMoney.Caption) - cashed
End If

rs.MoveFirst
While Not rs.EOF
If cmbcust.Text = rs!Name Then
custflag = True
credited2 = rs!Credit
End If
rs.MoveNext
Wend
If cmbcust.Text = "xxx" Then custflag = True
If Not custflag Then
MsgBox "Please Choose a Customer from the List ( xxx for Unknown )", vbOKOnly, "Error"
cmbcust.SetFocus
vlflag = True
End If

rs2.MoveFirst
While Not rs2.EOF
If cmbprod.Text = rs2!Description Then prodflag = True
rs2.MoveNext
Wend
If Not prodflag Then
MsgBox "Please Choose a Valid Product from the List", vbOKOnly, "Error"
cmbprod.SetFocus
vlflag = True
End If

If txtAmt.Text = "" Then
MsgBox "Please Choose Amount", vbOKOnly, "Error"
txtAmt.Text = ""
txtAmt.SetFocus
vlflag = True
End If

rs2.MoveFirst
While Not rs2.EOF
If cmbprod.Text = rs2!Description Then
If txtAmt.Text = "" Then
amtt = 0
Else
amtt = CInt(txtAmt.Text)
End If
    If amtt > rs2!Quantity Then
           MsgBox "Amount Not Present at Store. Please Choose a Lesser Amount.", vbOKOnly, "Error"
           txtAmt.Text = ""
           txtCash.Text = ""
           txtAmt.SetFocus
           vlflag = True
    End If
End If
rs2.MoveNext
Wend

credited2 = credited2 + credited
If credited2 > 500 Then
    MsgBox "Credit More than 500 Rupees is Not Possible.", vbOKOnly, "Error"
           txtCash.Text = ""
           txtCash.SetFocus
           vlflag = True
End If

If credited < 0 Then
MsgBox "Paying More Cash Than Needed.", vbOKOnly, "Error"
           txtCash.Text = ""
           txtCash.SetFocus
           vlflag = True
End If

If cmbcust.Text = "xxx" And credited > 0 Then
MsgBox "Cannot Credit to Unknown Customer.", vbOKOnly, "Error"
           txtCash.Text = ""
           txtCash.SetFocus
           vlflag = True
End If
End Sub

Private Sub Finalupdate()
Set rs = db.OpenRecordset("Contacts")
Set rs2 = db.OpenRecordset("Products")
Set rs3 = db.OpenRecordset("Cashrec")

rs.MoveFirst
While Not rs.EOF
If cmbcust.Text = rs!Name Then
    rs.Edit
    rs!Credit = rs!Credit + credited
    rs.Update
End If
rs.MoveNext
Wend
rs3.MoveFirst
rs3.Edit
rs3!Cash = rs3!Cash + cashed
rs3.Update

rs2.MoveFirst
While Not rs2.EOF
If cmbprod.Text = rs2!Description Then
    rs2.Edit
    rs2!Quantity = rs2!Quantity - CLng(txtAmt.Text)
    rs2.Update
End If
rs2.MoveNext
Wend
End Sub

Private Sub txtCash_Change()
Dim dig$, i, digi$, digits$
If txtCash.Text <> "" Then
    dig$ = Mid(txtCash.Text, Len(txtCash.Text), 1)
    If Asc(dig$) < 46 Or Asc(dig$) > 57 Then
        For i = 1 To Len(txtCash.Text) - 1
            digi$ = Mid(txtCash.Text, i, 1)
            digits$ = digits$ & digi$
        Next i
        txtCash.Text = digits$
        txtCash.SelStart = Len(txtCash.Text)
    End If
End If
End Sub
