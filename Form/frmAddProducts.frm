VERSION 5.00
Object = "{BB31661F-0587-11D6-9DD0-00C04F0BD97C}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmAddProducts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Product"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmAddProducts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjChameleon.chameleonButton cmdClose 
      Height          =   375
      Left            =   3480
      TabIndex        =   16
      Top             =   3600
      Width           =   855
      _ExtentX        =   1508
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
      MICON           =   "frmAddProducts.frx":0ECA
   End
   Begin prjChameleon.chameleonButton cmdAdd 
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   15
      Top             =   3600
      Width           =   855
      _ExtentX        =   1508
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
      MICON           =   "frmAddProducts.frx":0EE6
   End
   Begin VB.Frame Frame1 
      Caption         =   "Product Information"
      Height          =   3375
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtExpiry 
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
         TabIndex        =   6
         Top             =   2640
         Width           =   2895
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox txtUnitPrice 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtQuantity 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox txtTotalPrice 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   1920
         Width           =   2895
      End
      Begin VB.TextBox txtDate 
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
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2280
         Width           =   2895
      End
      Begin VB.Label Label7 
         Caption         =   "Expiry Date:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Product ID:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Unit Price:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Quantity:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Total Price:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Description:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Date:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2280
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmAddProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As Database
Dim rs As Recordset

Private Sub Form_Load()
Dim X
Set db = DBEngine.Workspaces(0).OpenDatabase(App.Path & "\db2.mdb")
Set rs = db.OpenRecordset("Products")
txtDate.Text = Date
frmMain.Frame2.Visible = False
End Sub

Private Sub cmdAdd_Click()
Set rs = db.OpenRecordset("Products")
If txtID.Text = "" Or _
    txtUnitPrice.Text = "" Or _
    txtQuantity.Text = "" Or _
    txtTotalPrice.Text = "" Or _
    txtDescription.Text = "" Or _
    txtDate.Text = "" Or _
    txtExpiry.Text = "" Then
    MsgBox "Please fill all the boxes."
Else
    rs.AddNew
    rs!ProductID = txtID.Text
    rs!UnitPrice = txtUnitPrice.Text
    rs!Quantity = txtQuantity.Text
    rs!TotalPrice = txtTotalPrice.Text
    rs!Description = txtDescription.Text
    rs!Date = txtDate.Text
    rs!DateExp = txtExpiry.Text
    rs.Update
    Limpyo
    frmMain.ProductView
    Unload Me
    frmMain.Frame2.Visible = True
End If

rs.Close
End Sub

Private Sub cmdClose_Click()
Unload Me
frmMain.Frame2.Visible = True
End Sub

Public Function Limpyo()
txtID.Text = ""
txtUnitPrice.Text = ""
txtQuantity.Text = ""
txtTotalPrice.Text = ""
txtDescription.Text = ""
txtExpiry.Text = ""
End Function


Private Sub txtQuantity_Change()
Dim dig$, i, digi$, digits$
If txtQuantity.Text <> "" Then
    dig$ = Mid(txtQuantity.Text, Len(txtQuantity.Text), 1)
    If Asc(dig$) < 46 Or Asc(dig$) > 57 Then
        For i = 1 To Len(txtQuantity.Text) - 1
            digi$ = Mid(txtQuantity.Text, i, 1)
            digits$ = digits$ & digi$
        Next i
        txtQuantity.Text = digits$
        txtQuantity.SelStart = Len(txtQuantity.Text)
    End If
End If
txtTotalPrice.Text = Val(txtUnitPrice.Text) * Val(txtQuantity.Text)
End Sub

Private Sub txtUnitPrice_Change()
Dim dig$, i, digi$, digits$
If txtUnitPrice.Text <> "" Then
    dig$ = Mid(txtUnitPrice.Text, Len(txtUnitPrice.Text), 1)
    If Asc(dig$) < 46 Or Asc(dig$) > 57 Then
        For i = 1 To Len(txtUnitPrice.Text) - 1
            digi$ = Mid(txtUnitPrice.Text, i, 1)
            digits$ = digits$ & digi$
        Next i
        txtUnitPrice.Text = digits$
        txtUnitPrice.SelStart = Len(txtUnitPrice.Text)
    End If
End If
txtTotalPrice.Text = Val(txtUnitPrice.Text) * Val(txtQuantity.Text)
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Frame2.Visible = True
End Sub
