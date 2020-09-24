VERSION 5.00
Object = "{BB31661F-0587-11D6-9DD0-00C04F0BD97C}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmtrUpdateCash 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update Cash At Store"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmtrUpdateCash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjChameleon.chameleonButton cmdtrDo 
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   2880
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
      MICON           =   "frmtrUpdateCash.frx":0ECA
   End
   Begin prjChameleon.chameleonButton cmdClose 
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   2880
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
      MICON           =   "frmtrUpdateCash.frx":0EE6
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   4455
      Begin VB.TextBox txtwth 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2280
         TabIndex        =   1
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtdep 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2280
         TabIndex        =   0
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblcash 
         Height          =   255
         Left            =   2880
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Current Cash at Store is:    Rupees"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Withdraw Rupees:"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Deposit Rupees:"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   1200
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmtrUpdateCash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As Database
Dim rs As Recordset
Dim vlflag As Boolean
Dim expcash As Long
Dim a, depcash, wthcash As Long

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdtrDo_Click()
vlflag = False
If txtdep = "" Then
depcash = 0
Else
depcash = CLng(txtdep.Text)
End If
If txtwth = "" Then
wthcash = 0
Else
wthcash = CLng(txtwth.Text)
End If
If depcash = 0 And wthcash = 0 Then
MsgBox "Please enter some values.", vbOKOnly, "Error"
txtdep.SetFocus
vlflag = True
End If
expcash = a - wthcash + depcash

If expcash < 1000 Then
MsgBox "This would deplete Cash Store Beyond Permissible Limit.", vbOKOnly, "Error"
txtdep.Text = ""
txtwth.Text = ""
txtdep.SetFocus
vlflag = True
End If

If expcash > 1000 And Not vlflag Then
    rs.MoveFirst
    rs.Edit
    rs!Cash = expcash
    rs.Update
    StatRefresh
MsgBox "Transaction Completed Successfully.", vbOKOnly, "OK"
Unload Me
End If
End Sub

Private Sub Form_Load()
frmMain.Frame2.Visible = False
Set db = DBEngine.Workspaces(0).OpenDatabase(App.Path & "\db2.mdb")
Set rs = db.OpenRecordset("Cashrec")
rs.MoveFirst
a = rs!Cash
lblcash.Caption = a
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Frame2.Visible = True
End Sub

Private Sub txtdep_Change()
Dim dig$, i, digi$, digits$
If txtdep.Text <> "" Then
    dig$ = Mid(txtdep.Text, Len(txtdep.Text), 1)
    If Asc(dig$) < 46 Or Asc(dig$) > 57 Then
        For i = 1 To Len(txtdep.Text) - 1
            digi$ = Mid(txtdep.Text, i, 1)
            digits$ = digits$ & digi$
        Next i
        txtdep.Text = digits$
        txtdep.SelStart = Len(txtdep.Text)
    End If
End If
End Sub

Private Sub txtwth_Change()
Dim dig$, i, digi$, digits$
If txtwth.Text <> "" Then
    dig$ = Mid(txtwth.Text, Len(txtwth.Text), 1)
    If Asc(dig$) < 46 Or Asc(dig$) > 57 Then
        For i = 1 To Len(txtwth.Text) - 1
            digi$ = Mid(txtwth.Text, i, 1)
            digits$ = digits$ & digi$
        Next i
        txtwth.Text = digits$
        txtwth.SelStart = Len(txtwth.Text)
    End If
End If
End Sub
