VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{BB31661F-0587-11D6-9DD0-00C04F0BD97C}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFC0C0&
   Caption         =   "ZpyderKeeper"
   ClientHeight    =   10515
   ClientLeft      =   60
   ClientTop       =   1065
   ClientWidth     =   15240
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0ECA
   ScaleHeight     =   10515
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin prjChameleon.chameleonButton cmdClose 
      Height          =   735
      Left            =   9240
      TabIndex        =   64
      Top             =   8040
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      BTYPE           =   6
      TX              =   "Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
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
      MICON           =   "frmMain.frx":1BE56
   End
   Begin prjChameleon.chameleonButton cmdMakalah 
      Height          =   735
      Left            =   6840
      TabIndex        =   63
      Top             =   8040
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      BTYPE           =   6
      TX              =   "Makalah"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12648384
      FCOL            =   0
      FCOLO           =   0
      MPTR            =   0
      MICON           =   "frmMain.frx":1BE72
   End
   Begin prjChameleon.chameleonButton cmdAbout 
      Height          =   735
      Left            =   4440
      TabIndex        =   62
      Top             =   8040
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      BTYPE           =   6
      TX              =   "About"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
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
      MICON           =   "frmMain.frx":1BE8E
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3840
      Top             =   7800
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3720
      Top             =   7080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BEAA
            Key             =   "Customers"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D1BC
            Key             =   "Products"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E4CE
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F7E0
            Key             =   "Transactions"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20AF2
            Key             =   "About"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6495
      Left            =   4440
      TabIndex        =   12
      Top             =   1320
      Width           =   6375
      Begin TabDlg.SSTab tabs 
         Height          =   6255
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   11033
         _Version        =   393216
         Style           =   1
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   1058
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "General *"
         TabPicture(0)   =   "frmMain.frx":21E04
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblInput"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label12"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label19"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label20"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "lsStart"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "cmbMenu"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "cmdOK"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).ControlCount=   7
         TabCaption(1)   =   "Customers"
         TabPicture(1)   =   "frmMain.frx":21E20
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cmdGeneral"
         Tab(1).Control(1)=   "cmdCAdd"
         Tab(1).Control(2)=   "cmdCView"
         Tab(1).Control(3)=   "cmdCDelete"
         Tab(1).Control(4)=   "cmdCSave"
         Tab(1).Control(5)=   "cmdCEdit"
         Tab(1).Control(6)=   "Frame3"
         Tab(1).ControlCount=   7
         TabCaption(2)   =   "Products"
         TabPicture(2)   =   "frmMain.frx":21E3C
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "cmdPSave"
         Tab(2).Control(1)=   "cmdPEdit"
         Tab(2).Control(2)=   "cmdPDelete"
         Tab(2).Control(3)=   "cmdPAdd"
         Tab(2).Control(4)=   "cmdPView"
         Tab(2).Control(5)=   "cmdGeneral2"
         Tab(2).Control(6)=   "Frame4"
         Tab(2).ControlCount=   7
         TabCaption(3)   =   "Search"
         TabPicture(3)   =   "frmMain.frx":21E58
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "cmdGeneral3"
         Tab(3).Control(1)=   "cmdPSearch"
         Tab(3).Control(2)=   "cmdCSearch"
         Tab(3).Control(3)=   "prodFrame"
         Tab(3).Control(4)=   "custFrame"
         Tab(3).ControlCount=   5
         TabCaption(4)   =   "Transactions"
         TabPicture(4)   =   "frmMain.frx":21E74
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "cmdGeneral4"
         Tab(4).Control(1)=   "cmdtrSell"
         Tab(4).Control(1).Enabled=   0   'False
         Tab(4).Control(2)=   "cmdtrBuy"
         Tab(4).Control(2).Enabled=   0   'False
         Tab(4).Control(3)=   "cmdtrCash"
         Tab(4).Control(3).Enabled=   0   'False
         Tab(4).Control(4)=   "cmdtrRedeem"
         Tab(4).Control(4).Enabled=   0   'False
         Tab(4).Control(5)=   "Statframe"
         Tab(4).ControlCount=   6
         Begin prjChameleon.chameleonButton cmdGeneral4 
            Height          =   615
            Left            =   -72600
            TabIndex        =   81
            Top             =   5520
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   1085
            BTYPE           =   8
            TX              =   "General"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            FCOL            =   0
            FCOLO           =   0
            MPTR            =   0
            MICON           =   "frmMain.frx":21E90
         End
         Begin prjChameleon.chameleonButton cmdGeneral3 
            Height          =   615
            Left            =   -72600
            TabIndex        =   80
            Top             =   5520
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   1085
            BTYPE           =   8
            TX              =   "General"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            FCOL            =   0
            FCOLO           =   0
            MPTR            =   0
            MICON           =   "frmMain.frx":21EAC
         End
         Begin VB.CommandButton cmdPSearch 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Search Products"
            Height          =   375
            Left            =   -74880
            Style           =   1  'Graphical
            TabIndex        =   79
            Top             =   2520
            Width           =   2535
         End
         Begin VB.CommandButton cmdCSearch 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Search Customers"
            Height          =   375
            Left            =   -74880
            Style           =   1  'Graphical
            TabIndex        =   78
            Top             =   1800
            Width           =   2535
         End
         Begin VB.CommandButton cmdPSave 
            Caption         =   "Save"
            Height          =   375
            Left            =   -70320
            TabIndex        =   77
            Top             =   3480
            Width           =   975
         End
         Begin VB.CommandButton cmdPEdit 
            Caption         =   "Edit"
            Height          =   375
            Left            =   -71400
            TabIndex        =   76
            Top             =   3480
            Width           =   1095
         End
         Begin VB.CommandButton cmdPDelete 
            Caption         =   "Delete"
            Height          =   375
            Left            =   -72480
            TabIndex        =   75
            Top             =   3480
            Width           =   1095
         End
         Begin VB.CommandButton cmdPAdd 
            Caption         =   "Add"
            Height          =   375
            Left            =   -73800
            TabIndex        =   74
            Top             =   3480
            Width           =   1095
         End
         Begin VB.CommandButton cmdPView 
            Caption         =   "View"
            Height          =   375
            Left            =   -74880
            TabIndex        =   73
            Top             =   3480
            Width           =   1095
         End
         Begin prjChameleon.chameleonButton cmdGeneral2 
            Height          =   615
            Left            =   -72600
            TabIndex        =   72
            Top             =   5520
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   1085
            BTYPE           =   8
            TX              =   "General"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            FCOL            =   0
            FCOLO           =   0
            MPTR            =   0
            MICON           =   "frmMain.frx":21EC8
         End
         Begin prjChameleon.chameleonButton cmdGeneral 
            Height          =   615
            Left            =   -72600
            TabIndex        =   71
            Top             =   5520
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   1085
            BTYPE           =   8
            TX              =   "General"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            FCOL            =   0
            FCOLO           =   0
            MPTR            =   0
            MICON           =   "frmMain.frx":21EE4
         End
         Begin VB.CommandButton cmdCAdd 
            Caption         =   "Add"
            Height          =   375
            Left            =   -73800
            TabIndex        =   70
            Top             =   3480
            Width           =   1095
         End
         Begin VB.CommandButton cmdCView 
            Caption         =   "View"
            Height          =   375
            Left            =   -74880
            TabIndex        =   69
            Top             =   3480
            Width           =   1095
         End
         Begin VB.CommandButton cmdCDelete 
            Caption         =   "Delete"
            Height          =   375
            Left            =   -72480
            TabIndex        =   68
            Top             =   3480
            Width           =   1095
         End
         Begin VB.CommandButton cmdCSave 
            Caption         =   "Save"
            Height          =   375
            Left            =   -70320
            TabIndex        =   67
            Top             =   3480
            Width           =   975
         End
         Begin VB.CommandButton cmdCEdit 
            Caption         =   "Edit"
            Height          =   375
            Left            =   -71400
            TabIndex        =   66
            Top             =   3480
            Width           =   1095
         End
         Begin prjChameleon.chameleonButton cmdOK 
            Height          =   495
            Left            =   2640
            TabIndex        =   65
            Top             =   5640
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   873
            BTYPE           =   8
            TX              =   "OK"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            FCOL            =   0
            FCOLO           =   0
            MPTR            =   0
            MICON           =   "frmMain.frx":21F00
         End
         Begin VB.ComboBox cmbMenu 
            BackColor       =   &H00C0FFFF&
            DataField       =   "1"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            ItemData        =   "frmMain.frx":21F1C
            Left            =   5160
            List            =   "frmMain.frx":21F1E
            Style           =   1  'Simple Combo
            TabIndex        =   2
            Top             =   4080
            Width           =   390
         End
         Begin VB.Frame Frame3 
            Height          =   2655
            Left            =   -74880
            TabIndex        =   49
            Top             =   660
            Width           =   5535
            Begin VB.TextBox txtCName 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   2760
               Locked          =   -1  'True
               TabIndex        =   8
               TabStop         =   0   'False
               Top             =   240
               Width           =   2535
            End
            Begin VB.TextBox txtCAddress 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   2760
               Locked          =   -1  'True
               TabIndex        =   7
               TabStop         =   0   'False
               Top             =   600
               Width           =   2535
            End
            Begin VB.TextBox txtCPhone 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   2760
               Locked          =   -1  'True
               TabIndex        =   6
               TabStop         =   0   'False
               Top             =   960
               Width           =   2535
            End
            Begin VB.TextBox txtCCel 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   2760
               Locked          =   -1  'True
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   1320
               Width           =   2535
            End
            Begin VB.TextBox txtEmail 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   2760
               Locked          =   -1  'True
               TabIndex        =   4
               TabStop         =   0   'False
               Top             =   1680
               Width           =   2535
            End
            Begin VB.TextBox txtCredit 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   2760
               Locked          =   -1  'True
               TabIndex        =   3
               TabStop         =   0   'False
               Top             =   2040
               Width           =   2535
            End
            Begin MSComctlLib.ListView lvwContacts 
               Height          =   2175
               Left            =   120
               TabIndex        =   9
               TabStop         =   0   'False
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
               BackColor       =   16777215
               BorderStyle     =   1
               Appearance      =   0
               Enabled         =   0   'False
               NumItems        =   0
            End
            Begin VB.Label Label2 
               Caption         =   "Name:"
               Height          =   255
               Left            =   1800
               TabIndex        =   55
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label3 
               Caption         =   "Address:"
               Height          =   255
               Left            =   1800
               TabIndex        =   54
               Top             =   600
               Width           =   975
            End
            Begin VB.Label Label4 
               Caption         =   "Phone No:"
               Height          =   255
               Left            =   1800
               TabIndex        =   53
               Top             =   960
               Width           =   975
            End
            Begin VB.Label Label5 
               Caption         =   "Cell No:"
               Height          =   255
               Left            =   1800
               TabIndex        =   52
               Top             =   1320
               Width           =   975
            End
            Begin VB.Label Label6 
               Caption         =   "Email:"
               Height          =   255
               Left            =   1800
               TabIndex        =   51
               Top             =   1680
               Width           =   975
            End
            Begin VB.Label Label14 
               Caption         =   "Credit:"
               Height          =   255
               Left            =   1800
               TabIndex        =   50
               Top             =   2040
               Width           =   975
            End
         End
         Begin VB.Frame Frame4 
            Height          =   2655
            Left            =   -74880
            TabIndex        =   42
            Top             =   660
            Width           =   5535
            Begin VB.TextBox txtUnitPrice 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   3000
               Locked          =   -1  'True
               TabIndex        =   18
               TabStop         =   0   'False
               Top             =   240
               Width           =   2295
            End
            Begin VB.TextBox txtQuantity 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   3000
               Locked          =   -1  'True
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   600
               Width           =   2295
            End
            Begin VB.TextBox txtTotalPrice 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   3000
               Locked          =   -1  'True
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   960
               Width           =   2295
            End
            Begin VB.TextBox txtDate 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   3000
               Locked          =   -1  'True
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   1320
               Width           =   2295
            End
            Begin VB.TextBox txtDescription 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   3000
               Locked          =   -1  'True
               TabIndex        =   22
               TabStop         =   0   'False
               Top             =   1680
               Width           =   2295
            End
            Begin VB.TextBox txtExpiry 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   3000
               Locked          =   -1  'True
               TabIndex        =   23
               TabStop         =   0   'False
               Top             =   2040
               Width           =   2295
            End
            Begin MSComctlLib.ListView lvwProducts 
               Height          =   2175
               Left            =   120
               TabIndex        =   17
               TabStop         =   0   'False
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
               BackColor       =   16777215
               BorderStyle     =   1
               Appearance      =   0
               Enabled         =   0   'False
               NumItems        =   0
            End
            Begin VB.Label Label7 
               Caption         =   "Unit Price:"
               Height          =   255
               Left            =   1680
               TabIndex        =   48
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label8 
               Caption         =   "Quantity:"
               Height          =   255
               Left            =   1680
               TabIndex        =   47
               Top             =   600
               Width           =   975
            End
            Begin VB.Label Label9 
               Caption         =   "Total Price:"
               Height          =   255
               Left            =   1680
               TabIndex        =   46
               Top             =   960
               Width           =   975
            End
            Begin VB.Label Label10 
               Caption         =   "Date Purchased:"
               Height          =   255
               Left            =   1680
               TabIndex        =   45
               Top             =   1320
               Width           =   1215
            End
            Begin VB.Label Label11 
               Caption         =   "Description:"
               Height          =   255
               Left            =   1680
               TabIndex        =   44
               Top             =   1680
               Width           =   1215
            End
            Begin VB.Label Label13 
               Caption         =   "Expiry Date:"
               Height          =   255
               Left            =   1680
               TabIndex        =   43
               Top             =   2040
               Width           =   1215
            End
         End
         Begin VB.CommandButton cmdtrSell 
            BackColor       =   &H00C0E0FF&
            Caption         =   "&Sell Product to Customer"
            Height          =   375
            Left            =   -74760
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   1080
            Width           =   2535
         End
         Begin VB.CommandButton cmdtrBuy 
            BackColor       =   &H00FFC0FF&
            Caption         =   "B&uy Product from Supplier"
            Height          =   375
            Left            =   -74760
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   1680
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.CommandButton cmdtrCash 
            BackColor       =   &H00FFC0C0&
            Caption         =   "&Update Cash at Store"
            Height          =   375
            Left            =   -74760
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   2880
            Width           =   2535
         End
         Begin VB.CommandButton cmdtrRedeem 
            BackColor       =   &H00C0FFC0&
            Caption         =   "&Redeem Credited Amount"
            Height          =   375
            Left            =   -74760
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   2280
            Width           =   2535
         End
         Begin VB.Frame Statframe 
            Height          =   3255
            Left            =   -72120
            TabIndex        =   35
            Top             =   720
            Width           =   2655
            Begin VB.TextBox txtprod 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   1920
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   1080
               Width           =   615
            End
            Begin VB.TextBox txtcust 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   1920
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   1680
               Width           =   615
            End
            Begin VB.TextBox txtcash 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   1680
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   2280
               Width           =   855
            End
            Begin VB.Label Label15 
               Caption         =   "Customer Members:"
               Height          =   255
               Left            =   240
               TabIndex        =   39
               Top             =   1680
               Width           =   1455
            End
            Begin VB.Label Label16 
               Caption         =   "Product Types:"
               Height          =   255
               Left            =   240
               TabIndex        =   38
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label Label17 
               Caption         =   "Cash at Store:"
               Height          =   255
               Left            =   240
               TabIndex        =   37
               Top             =   2280
               Width           =   1095
            End
            Begin VB.Label Label18 
               Caption         =   "ZpyderStore"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   240
               TabIndex        =   36
               Top             =   480
               Width           =   2175
            End
         End
         Begin Zpyderkeeper.cpvCoolList lsStart 
            Height          =   3060
            Left            =   0
            TabIndex        =   1
            Top             =   840
            Width           =   6015
            _ExtentX        =   10610
            _ExtentY        =   5398
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ItemHeight      =   40
            ItemHeightAuto  =   0   'False
            SelectModeStyle =   3
         End
         Begin VB.Frame prodFrame 
            Height          =   3255
            Left            =   -72120
            TabIndex        =   40
            Top             =   720
            Visible         =   0   'False
            Width           =   2655
            Begin VB.CommandButton cmdPTotpr 
               BackColor       =   &H00FFFFC0&
               Caption         =   "By &Total Price"
               Height          =   375
               Left            =   480
               Style           =   1  'Graphical
               TabIndex        =   31
               TabStop         =   0   'False
               Top             =   2520
               Width           =   1695
            End
            Begin VB.CommandButton cmdPExpiry 
               BackColor       =   &H00FFFFC0&
               Caption         =   "By &Expiry Date"
               Height          =   375
               Left            =   480
               Style           =   1  'Graphical
               TabIndex        =   30
               TabStop         =   0   'False
               Top             =   1800
               Width           =   1695
            End
            Begin VB.CommandButton cmdPQua 
               BackColor       =   &H00FFFFC0&
               Caption         =   "By &Quantity"
               Height          =   375
               Left            =   480
               Style           =   1  'Graphical
               TabIndex        =   29
               TabStop         =   0   'False
               Top             =   1080
               Width           =   1695
            End
            Begin VB.CommandButton cmdPDesc 
               BackColor       =   &H00FFFFC0&
               Caption         =   "By &Description"
               Height          =   375
               Left            =   480
               Style           =   1  'Graphical
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   360
               Width           =   1695
            End
         End
         Begin VB.Frame custFrame 
            Height          =   3255
            Left            =   -72120
            TabIndex        =   41
            Top             =   720
            Width           =   2655
            Begin VB.CommandButton cmdCName 
               BackColor       =   &H8000000E&
               Caption         =   "By &Name"
               Height          =   375
               Left            =   480
               Style           =   1  'Graphical
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   360
               Width           =   1695
            End
            Begin VB.CommandButton cmdCAddr 
               BackColor       =   &H8000000E&
               Caption         =   "By &Address"
               Height          =   375
               Left            =   480
               Style           =   1  'Graphical
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   1080
               Width           =   1695
            End
            Begin VB.CommandButton cmdCPhone 
               BackColor       =   &H8000000E&
               Caption         =   "By P&hone No"
               Height          =   375
               Left            =   480
               Style           =   1  'Graphical
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   1800
               Width           =   1695
            End
            Begin VB.CommandButton cmdCCredit 
               BackColor       =   &H8000000E&
               Caption         =   "By C&redit"
               Height          =   375
               Left            =   480
               Style           =   1  'Graphical
               TabIndex        =   27
               TabStop         =   0   'False
               Top             =   2520
               Width           =   1695
            End
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "*) P.S. : (Input Menu must arrange to letter c)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   61
            Top             =   4560
            Width           =   5055
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "**) P.S. : (Highlight Marker must put at list Customers (c))"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   4920
            Width           =   5775
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmMain.frx":21F20
            Height          =   375
            Left            =   360
            TabIndex        =   59
            Top             =   5160
            Width           =   4575
         End
         Begin VB.Label lblInput 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Input Menu (First Letter Menu) **                    :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   58
            Top             =   4080
            Width           =   5055
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   3120
      TabIndex        =   10
      Top             =   240
      Width           =   9255
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ZpyderKeeper - Inventory Management System"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   465
         Left            =   360
         TabIndex        =   11
         Top             =   120
         Width           =   8295
      End
   End
   Begin VB.Label lblKeyin2 
      BackStyle       =   0  'Transparent
      Caption         =   "ZpyderKeeper - ""The Ultimate Retail Store Inventory Solution""                Contact Me : zpiderboi@programmer.net"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   -8040
      TabIndex        =   57
      Top             =   9480
      Width           =   8295
   End
   Begin VB.Label lblKeyin1 
      BackStyle       =   0  'Transparent
      Caption         =   "ZpyderKeeper - ""The Ultimate Retail Store Inventory Solution""                Contact Me : zpiderboi@yahoo.com"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   3480
      TabIndex        =   56
      Top             =   9480
      Width           =   8295
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "&Menu"
      Begin VB.Menu mnuGen 
         Caption         =   "&General"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuCust 
         Caption         =   "&Customers"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuProd 
         Caption         =   "&Products"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuSrch 
         Caption         =   "&Search"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuTrans 
         Caption         =   "&Transactions"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuDash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "E&xit"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "A&bout"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim db As Database
Dim rs As Recordset
Dim cflag As Boolean
Dim pflag As Boolean
Dim lstmain1 As ListItem
Dim lstmain2 As ListItem
Dim clmhead1 As ColumnHeader
Dim clmhead2 As ColumnHeader
Dim trc, trp, trch As Long


Private Sub cmbMenu_Change()

cmdOK_Click
End Sub


Private Sub cmdAbout_Click()
Frame1.Visible = False
frmAbout.Show
End Sub

Private Sub cmdCAddr_Click()
frmSrchAddr.Show
End Sub

Private Sub cmdCancel_Click()

End Sub

Private Sub cmdCCredit_Click()
frmSrchCredit.Show
End Sub

Private Sub cmdCEdit_Click()
cmdCSave.Visible = True
        txtCName.Locked = False
        txtCAddress.Locked = False
        txtCPhone.Locked = False
        txtCCel.Locked = False
        txtEmail.Locked = False
        txtCredit.Locked = False
        txtCName.SetFocus
End Sub

Private Sub cmdCName_Click()
frmSrchName.Show
End Sub

Private Sub cmdCPhone_Click()
frmSrchPhone.Show
End Sub

Private Sub cmdCSave_Click()
Set rs = db.OpenRecordset("Contacts")
cmdCSave.Visible = False
        txtCName.Locked = True
        txtCAddress.Locked = True
        txtCPhone.Locked = True
        txtCCel.Locked = True
        txtEmail.Locked = True
        txtCredit.Locked = True
        While Not rs.EOF
        If lvwContacts.SelectedItem.Text = rs!ContactID Then
            rs.Edit
            rs!Name = txtCName.Text
            rs!Address = txtCAddress.Text
            rs!PhoneNo = txtCPhone.Text
            rs!Celno = txtCCel.Text
            rs!Email = txtEmail
            rs!Credit = txtCredit.Text
            rs.Update
            rs.MoveLast
            rs.MoveNext
        Else
            rs.MoveNext
        End If
        Wend
End Sub

Private Sub cmdCSearch_Click()
prodFrame.Visible = False
custFrame.Visible = True
End Sub

Private Sub cmdGeneral_Click()

    DoEvents
    frmMain.Show
    frmMain.tabs.Tab = 0
End Sub

Private Sub cmdGeneral2_Click()

    DoEvents
    frmMain.Show
    frmMain.tabs.Tab = 0
End Sub

Private Sub cmdGeneral3_Click()

    DoEvents
    frmMain.Show
    frmMain.tabs.Tab = 0
End Sub

Private Sub cmdGeneral4_Click()

    DoEvents
    frmMain.Show
    frmMain.tabs.Tab = 0
End Sub

Private Sub cmdMakalah_Click()
On Error GoTo errHandle
    Dim a As Double
    a = Shell("C:\Program Files\Windows NT\Accessories\wordpad.exe C:\Inventory.doc ", vbNormalFocus)
    Exit Sub
errHandle:
    MsgBox "Tidak Terdapat Program WordPad Pada Komputer Anda Sehingga Makalah Tidak Bisa Dibuka", vbInformation, "Error in opening!!!"
    Resume Next
End Sub

Private Sub cmdOK_Click()
Dim c As Integer
Dim p As Integer
Dim s As Integer
Dim t As Integer
Dim a As Integer

Select Case lsStart.ListIndex Or cmbMenu.ListIndex
    Case 0
    
    DoEvents
    frmMain.Show
    frmMain.tabs.Tab = 1
    
    Case 1
    
    DoEvents
    frmMain.Show
    frmMain.tabs.Tab = 2
    
    Case 2
    
    DoEvents
    frmMain.Show
    frmMain.tabs.Tab = 3
    
    Case 3
    
    DoEvents
    frmMain.Show
    frmMain.tabs.Tab = 4
    
    Case 4
    
    DoEvents
    frmAbout.Show
    
    
    Exit Sub

End Select
End Sub

Private Sub cmdPAdd_Click()
frmAddProducts.Show
End Sub

Private Sub cmdCAdd_Click()
frmAddContacts.Show
End Sub

Private Sub cmdPDesc_Click()
frmSrchDesc.Show
End Sub

Private Sub cmdPEdit_Click()
cmdPSave.Visible = True
            txtUnitPrice.Locked = False
            txtQuantity.Locked = False
            txtTotalPrice.Locked = False
            txtDescription.Locked = False
            txtDate.Locked = True
            txtExpiry.Locked = False
            txtUnitPrice.SetFocus
End Sub

Private Sub cmdPExpiry_Click()
frmSrchExpiry.Show
End Sub

Private Sub cmdPQua_Click()
frmSrchQua.Show
End Sub

Private Sub cmdPSave_Click()
Set rs = db.OpenRecordset("Products")
cmdPSave.Visible = False
            txtUnitPrice.Locked = True
            txtQuantity.Locked = True
            txtTotalPrice.Locked = True
            txtDescription.Locked = True
            txtDate.Locked = True
            txtExpiry.Locked = True
            
            While Not rs.EOF
        If lvwProducts.SelectedItem.Text = rs!ProductID Then
            rs.Edit
            rs!UnitPrice = txtUnitPrice.Text
            rs!Quantity = txtQuantity.Text
            rs!TotalPrice = txtTotalPrice.Text
            rs!Description = txtDescription.Text
            rs!Date = txtDate.Text
            rs!DateExp = txtExpiry.Text
            rs.Update
            rs.MoveLast
            rs.MoveNext
        Else
            rs.MoveNext
        End If
        Wend
End Sub

Private Sub cmdPSearch_Click()
prodFrame.Visible = True
custFrame.Visible = False
End Sub

Private Sub cmdPTotpr_Click()
frmSrchTotpr.Show
End Sub



Private Sub cmdtrCash_Click()
frmtrUpdateCash.Show
End Sub

Private Sub cmdtrRedeem_Click()
frmtrRedeem.Show
End Sub

Private Sub cmdtrSell_Click()
frmtrSell.Show
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Form_Initialize()

lsStart.SetImageList Me.ImageList1

lsStart.AddItem "Customers (c)", ImageList1.ListImages("Customers").Index, ImageList1.ListImages("Customers").Index
lsStart.AddItem "Products (p)", ImageList1.ListImages("Products").Index, ImageList1.ListImages("Products").Index
lsStart.AddItem "Search (s)", ImageList1.ListImages("Search").Index, ImageList1.ListImages("Search").Index
lsStart.AddItem "Transactions (t)", ImageList1.ListImages("Transactions").Index, ImageList1.ListImages("Transactions").Index
lsStart.AddItem "About (a)", ImageList1.ListImages("About").Index, ImageList1.ListImages("About").Index


End Sub

Private Sub Form_Load()
Dim i As Integer

For i = 1 To 5
    cmbMenu.AddItem Choose(i, "c", "p", "s", "t", "a")

Next
    
Set db = DBEngine.Workspaces(0).OpenDatabase(App.Path & "\db2.mdb")
Set clmhead1 = lvwContacts.ColumnHeaders.Add(, , "Contact ID", lvwContacts.Width)
Set clmhead2 = lvwProducts.ColumnHeaders.Add(, , "Product ID", lvwProducts.Width)
StatRefresh
cflag = False
pflag = False
frmMain.tabs.Tab = 0
End Sub

Private Sub cmdCView_Click()
If cflag = False Then
cflag = True
ContactView
Else
cflag = False
ContactNoview
End If
End Sub

Private Sub cmdPView_Click()
If pflag = False Then
pflag = True
ProductView
Else
pflag = False
ProductNoview
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub lsStart_DblClick()
cmdOK_Click
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

Public Sub cmdCDelete_Click()
Set rs = db.OpenRecordset("Contacts")
        txtCName.Text = ""
        txtCAddress.Text = ""
        txtCPhone.Text = ""
        txtCCel.Text = ""
        txtEmail.Text = ""
        txtCredit.Text = ""
    cmdCDelete.Enabled = False
    cmdCEdit.Enabled = False
While Not rs.EOF
    If frmMain.lvwContacts.SelectedItem.Text = rs!ContactID Then
    rs.Delete
    ContactView
        If rs.RecordCount = 0 Then
            Exit Sub
        Else
            rs.MoveLast
        End If
    rs.MoveNext
    Else
    rs.MoveNext
    End If
Wend
rs.Close
End Sub

Public Sub cmdPDelete_Click()
Set rs = db.OpenRecordset("Products")
            txtUnitPrice.Text = ""
            txtQuantity.Text = ""
            txtTotalPrice.Text = ""
            txtDescription.Text = ""
            txtDate.Text = ""
            txtExpiry.Text = ""
While Not rs.EOF
    If frmMain.lvwProducts.SelectedItem.Text = rs!ProductID Then
    rs.Delete
    cmdPDelete.Enabled = False
    cmdPEdit.Enabled = False
    ProductView
        If rs.RecordCount = 0 Then
            Exit Sub
        Else
            rs.MoveLast
        End If
    rs.MoveNext
    Else
    rs.MoveNext
    End If
Wend
rs.Close
End Sub

Private Sub mnuAbout_Click()
Frame1.Visible = False
frmAbout.Show
End Sub

Private Sub mnuClose_Click()
End
End Sub

Private Sub cmdClose_Click()
End
End Sub

Public Sub ProductView()
Set rs = db.OpenRecordset("Products")
frmMain.lvwProducts.ListItems.Clear
    While Not rs.EOF
    Set lstmain2 = frmMain.lvwProducts.ListItems.Add(, , rs!ProductID)
    rs.MoveNext
    Wend
    If rs.RecordCount = 0 Then
        MsgBox "There are no records on the database"
    Else
        frmMain.lvwProducts.Enabled = True
        frmMain.cmdPDelete.Enabled = True
        frmMain.cmdPEdit.Enabled = True
    End If
End Sub

Public Sub ContactView()
Set rs = db.OpenRecordset("Contacts")
frmMain.lvwContacts.ListItems.Clear
    While Not rs.EOF
    Set lstmain1 = frmMain.lvwContacts.ListItems.Add(, , rs!ContactID)
    rs.MoveNext
    Wend
    If rs.RecordCount = 0 Then
        MsgBox "There are no records on the database"
    Else
        frmMain.lvwContacts.Enabled = True
        frmMain.cmdCDelete.Enabled = True
        frmMain.cmdCEdit.Enabled = True
    End If
End Sub

Public Sub ContactNoview()
frmMain.lvwContacts.ListItems.Clear
frmMain.lvwContacts.Enabled = False
        frmMain.cmdCDelete.Enabled = False
        frmMain.cmdCEdit.Enabled = False
txtCName.Text = ""
    txtCAddress.Text = ""
    txtCPhone.Text = ""
    txtCCel.Text = ""
    txtEmail.Text = ""
    txtCredit.Text = ""
End Sub

Public Sub ProductNoview()
frmMain.lvwProducts.ListItems.Clear
frmMain.lvwProducts.Enabled = False
        frmMain.cmdPDelete.Enabled = False
        frmMain.cmdPEdit.Enabled = False
txtUnitPrice.Text = ""
    txtQuantity.Text = ""
    txtTotalPrice.Text = ""
    txtDescription.Text = ""
    txtDate.Text = ""
    txtExpiry.Text = ""
End Sub

Private Sub mnuCust_Click()
frmMain.Show
frmMain.tabs.Tab = 1
End Sub

Private Sub mnuGen_Click()
frmMain.Show
frmMain.tabs.Tab = 0
End Sub

Private Sub mnuProd_Click()
frmMain.Show
frmMain.tabs.Tab = 2
End Sub

Private Sub mnuSrch_Click()
frmMain.Show
frmMain.tabs.Tab = 3
End Sub

Private Sub mnuTrans_Click()
frmMain.Show
frmMain.tabs.Tab = 4
End Sub

Public Sub StatRefresh()

Set rs = db.OpenRecordset("Products")
trp = rs.RecordCount
txtprod.Text = trp
Set rs = db.OpenRecordset("Contacts")
trc = rs.RecordCount
txtcust.Text = trc
Set rs = db.OpenRecordset("Cashrec")
rs.MoveFirst
trch = rs!Cash
txtCash.Text = trch
End Sub

Private Sub tabs_DblClick()
StatRefresh
End Sub

Private Sub Timer1_Timer()
DoEvents
    Static lbl2 As Boolean
    Static lbl1 As Boolean
    If lbl1 = False Then
        lblKeyin1.Move lblKeyin1.Left + 20
    End If
    If lbl2 = True Then
        lblKeyin2.Move lblKeyin2.Left + 20
    End If
    If lblKeyin1.Left + lblKeyin1.Width >= Me.Width Then
        If lbl2 = False Then lbl2 = True
    End If
    If lblKeyin2.Left + lblKeyin2.Width >= Me.Width Then
        If lbl1 = True Then lbl1 = False
    End If
    If lblKeyin1.Left >= Me.Width Then
        lblKeyin1.Left = 0 - lblKeyin1.Width
        lbl1 = True
    End If
    If lblKeyin2.Left >= Me.Width Then
        lblKeyin2.Left = 0 - lblKeyin2.Width
        lbl2 = False
    End If
End Sub

