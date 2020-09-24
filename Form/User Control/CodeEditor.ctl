VERSION 5.00
Object = "{ECEDB943-AC41-11D2-AB20-000000000000}#2.0#0"; "CMAX20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl CodeEditor 
   Alignable       =   -1  'True
   ClientHeight    =   5010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6150
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   5010
   ScaleWidth      =   6150
   Begin CodeMaxCtl.CodeMax Code 
      Height          =   2655
      Index           =   0
      Left            =   1680
      OleObjectBlob   =   "CodeEditor.ctx":0000
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSComctlLib.TabStrip Tab1 
      Height          =   3855
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   6800
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "CodeEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public R As CodeMaxCtl.Range
Public Job As String

Private HTMLString(0 To 1000) As String

Dim ECount As Long

Public Event Click()
Public Event ContentMenu()
Public Event SelChange(ByVal Control As CodeMaxCtl.ICodeMax)

Private Function Code_Click(Index As Integer, ByVal Control As CodeMaxCtl.ICodeMax) As Boolean
RaiseEvent Click
End Function

Private Sub Code_SelChange(Index As Integer, ByVal Control As CodeMaxCtl.ICodeMax)
RaiseEvent SelChange(Control)

Set R = Code(Index).GetSel(True)
Code(Index).HighlightedLine = R.EndLineNo

If Code(Index).SelText <> "" Then Exit Sub
If R.EndColNo > Code(Index).GetLineLength(R.EndLineNo) Then
 Code(Index).SetCaretPos R.EndLineNo, Code(Index).GetLineLength(R.EndLineNo)
End If
End Sub

Private Sub Tab1_Click()
Dim dnum As Integer
dnum = CInt(Mid(Tab1.SelectedItem.key, 4))
Code(dnum).ZOrder 0
End Sub

Private Sub Tab1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then RaiseEvent ContentMenu
End Sub

Private Sub UserControl_Initialize()
Tab1.Tabs.Clear

End Sub

Private Sub UserControl_Resize()
Dim i As Integer
On Error Resume Next
Tab1.Move 0, 0, UserControl.Width, UserControl.Height
For i = 0 To Code.UBound
    Code(i).Move Tab1.clientLeft, Tab1.clientTop, Tab1.clientWidth, Tab1.clientHeight
Next
End Sub

Public Property Get EditorSets(Index) As CodeMax

    Set EditorSets = Code(Index)

End Property

Public Property Get EditorUbound() As Long
EditorUbound = Code.UBound
End Property

Public Property Get EditorCount() As Long
EditorCount = ECount
End Property

Public Sub RemoveEditor(Index)

    ECount = ECount - 1
        On Error GoTo 1
        
        Unload Code(Index)
        Tab1.Tabs.Remove "eds" & Index
        
Exit Sub
1
ECount = ECount + 1
Err.Raise Err.Number
End Sub

Public Property Get EditorHTML(Index) As String
    On Error Resume Next
    EditorHTML = HTMLString(Index)
End Property

Public Property Let EditorHTML(Index, str As String)
    HTMLString(Index) = str
End Property

Public Property Get ActiveHTML() As String
On Error Resume Next
ActiveHTML = HTMLString(CInt(Mid(Tab1.SelectedItem.key, 4)))
End Property

Public Property Let ActiveHTML(str As String)
HTMLString(CInt(Mid(Tab1.SelectedItem.key, 4))) = str
End Property

Public Property Get ActiveEditor() As CodeMax
On Error GoTo 1
    Set ActiveEditor = Code(CLng(Mid(Tab1.SelectedItem.key, 4)))
Exit Property
1
End Property

Public Function IsExist(Index) As Boolean
Dim a As String
On Error GoTo 1
a = Tab1.Tabs("eds" & Index)
IsExist = True
Exit Function
1 IsExist = False
End Function

Public Sub ShowEditorSet(Index)
Tab1.Tabs("eds" & Index).Selected = True
Code(Index).ZOrder 0
End Sub

Public Function AddEditorSet(Index, ByVal Caption As String, HTMLCode As String) As CodeMax
    
    ECount = ECount + 1
    
    On Error GoTo 1
    Load Code(Index)
    
    Tab1.Tabs.Add , "eds" & Index, Caption
    Tab1.Tabs("eds" & Index).Selected = True
    
    'ReDim Preserve HTMLString(Index To Index)
    HTMLString(Index) = HTMLCode
    
    With Code(Index)
        .Visible = True
        .ZOrder 0
        .text = HTMLCode
        .Move Tab1.clientLeft, Tab1.clientTop, Tab1.clientWidth, Tab1.clientHeight
    End With

    
    UserControl_Resize
    
Exit Function
1
ECount = ECount - 1
MsgBox Error
'Err.Raise Err.Number
End Function
