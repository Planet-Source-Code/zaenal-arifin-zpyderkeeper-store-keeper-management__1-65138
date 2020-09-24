VERSION 5.00
Object = "{BB31661F-0587-11D6-9DD0-00C04F0BD97C}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmAddContacts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Customer"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmAddContacts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjChameleon.chameleonButton cmdClose 
      Height          =   375
      Left            =   3360
      TabIndex        =   16
      Top             =   3600
      Width           =   975
      _ExtentX        =   1720
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
      MICON           =   "frmAddContacts.frx":0ECA
   End
   Begin prjChameleon.chameleonButton cmdAdd 
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   15
      Top             =   3600
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   8
      TX              =   "Add"
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
      MICON           =   "frmAddContacts.frx":0EE6
   End
   Begin VB.Frame Frame1 
      Caption         =   "Customer Information"
      Height          =   3375
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtCredit 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   2640
         Width           =   2895
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   2280
         Width           =   2895
      End
      Begin VB.TextBox txtCel 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   1920
         Width           =   2895
      End
      Begin VB.TextBox txtPhone 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Credit:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Email:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Cell no:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Phone no:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Address:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Contact ID:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmAddContacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As Database
Dim rs As Recordset

Private Sub Form_Load()
frmMain.Frame2.Visible = False
Dim X
Set db = DBEngine.Workspaces(0).OpenDatabase(App.Path & "\db2.mdb")
Set rs = db.OpenRecordset("Contacts")

If rs.RecordCount = 0 Then
txtID.Text = "1"
Else
rs.MoveLast
txtID.Text = rs!ContactID + 1
End If
End Sub

Private Sub cmdAdd_Click()
Set rs = db.OpenRecordset("Contacts")
If txtID.Text = "" Or _
    txtName.Text = "" Or _
    txtAddress.Text = "" Or _
    txtPhone.Text = "" Or _
    txtCel.Text = "" Or _
    txtEmail.Text = "" Or _
    txtCredit.Text = "" Then
    MsgBox "Please fill all the boxes."
Else
    rs.AddNew
    rs!ContactID = txtID.Text
    rs!Name = txtName.Text
    rs!Address = txtAddress.Text
    rs!PhoneNo = txtPhone.Text
    rs!Celno = txtCel.Text
    rs!Email = txtEmail.Text
    rs!Credit = txtCredit.Text
    rs.Update
    Limpyo
    If rs.RecordCount = 0 Then
        txtID.Text = "1"
    Else
        rs.MoveLast
        txtID.Text = rs!ContactID + 1
    End If
        
    frmMain.ContactView
    Unload Me
    frmMain.Frame2.Visible = True
End If
rs.Close
End Sub

Private Sub cmdClose_Click()
frmMain.Frame2.Visible = True
Unload Me
End Sub

Public Function Limpyo()
txtName.Text = ""
txtAddress.Text = ""
txtPhone.Text = ""
txtCel.Text = ""
txtEmail.Text = ""
txtCredit.Text = ""
End Function

Private Sub Form_Unload(Cancel As Integer)
frmMain.Frame2.Visible = True
End Sub

Private Sub txtCel_Change()
Dim dig$, i, digi$, digits$
If txtCel.Text <> "" Then
    dig$ = Mid(txtCel.Text, Len(txtCel.Text), 1)
    If Asc(dig$) < 46 Or Asc(dig$) > 57 Then
        For i = 1 To Len(txtCel.Text) - 1
            digi$ = Mid(txtCel.Text, i, 1)
            digits$ = digits$ & digi$
        Next i
        txtCel.Text = digits$
        txtCel.SelStart = Len(txtCel.Text)
    End If
End If
End Sub

Private Sub txtCredit_Change()
Dim dig$, i, digi$, digits$
If txtCredit.Text <> "" Then
    dig$ = Mid(txtCredit.Text, Len(txtCredit.Text), 1)
    If Asc(dig$) < 46 Or Asc(dig$) > 57 Then
        For i = 1 To Len(txtCredit.Text) - 1
            digi$ = Mid(txtCredit.Text, i, 1)
            digits$ = digits$ & digi$
        Next i
        txtCredit.Text = digits$
        txtCredit.SelStart = Len(txtCredit.Text)
    End If
End If
End Sub

Private Sub txtPhone_Change()
Dim dig$, i, digi$, digits$
If txtPhone.Text <> "" Then
    dig$ = Mid(txtPhone.Text, Len(txtPhone.Text), 1)
    If Asc(dig$) < 46 Or Asc(dig$) > 57 Then
        For i = 1 To Len(txtPhone.Text) - 1
            digi$ = Mid(txtPhone.Text, i, 1)
            digits$ = digits$ & digi$
        Next i
        txtPhone.Text = digits$
        txtPhone.SelStart = Len(txtPhone.Text)
    End If
End If
End Sub
