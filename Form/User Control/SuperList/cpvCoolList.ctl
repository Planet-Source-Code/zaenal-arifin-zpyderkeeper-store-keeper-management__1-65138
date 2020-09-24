VERSION 5.00
Begin VB.UserControl cpvCoolList 
   Appearance      =   0  '¥­­±
   BackColor       =   &H80000005&
   ClientHeight    =   2640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2580
   ClipControls    =   0   'False
   FillStyle       =   0  '¹ê¤ß
   KeyPreview      =   -1  'True
   ScaleHeight     =   176
   ScaleMode       =   3  '¹³¯À
   ScaleWidth      =   172
   ToolboxBitmap   =   "cpvCoolList.ctx":0000
   Begin VB.VScrollBar sbUC 
      Height          =   1125
      Left            =   1800
      Max             =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox iUC 
      Appearance      =   0  '¥­­±
      BackColor       =   &H80000005&
      BorderStyle     =   0  '¨S¦³®Ø½u
      ClipControls    =   0   'False
      FillStyle       =   0  '¹ê¤ß
      ForeColor       =   &H80000008&
      Height          =   1590
      Left            =   15
      ScaleHeight     =   106
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   78
      TabIndex        =   0
      Top             =   90
      Width           =   1170
   End
   Begin VB.Menu SelectionMenu 
      Caption         =   "SelectionMenu"
      Begin VB.Menu optSelectionMenu 
         Caption         =   "&Select all"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu optSelectionMenu 
         Caption         =   "&Unselect all"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu optSelectionMenu 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu optSelectionMenu 
         Caption         =   "&Invert"
         Enabled         =   0   'False
         Index           =   3
      End
   End
End
Attribute VB_Name = "cpvCoolList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
'=================================================================================
'                                                              cpvCoolList OCX 1.0
'
'                                                                      Carles P.V.
'                                                               carles_pv@terra.es
'---------------------------------------------------------------------------------
'                                                        Last modified: 12/09/2001
'=================================================================================













'#################################################################################
'## UserControl constants / types / variables / events
'#################################################################################

Public Enum Alignment
            [AlignLeft] = DT_LEFT
            [AlignCenter] = DT_CENTER
            [AlignRight] = DT_RIGHT
End Enum

Public Enum Appearance
            [Flat]
            [3D]
End Enum

Public Enum BorderStyle
            [None]
            [Fixed Single]
End Enum

Public Enum OrderType
            [Ascendent]
            [Descendent]
End Enum

Public Enum SelectMode
            [Single]
            [multiple]
End Enum

Public Enum SelectModeStyle
            [Standard]
            [Dither]
            [Gradient_V]
            [Gradient_H]
            [Box]
            [Underline]
            [byPicture]
End Enum

'#################################################################################

Private Type Item
            Text As String
            Icon As Integer
            IconSelected As Integer
End Type

'#################################################################################

Dim List() As Item                  '# List array of items (Text, icons)
Dim Selected() As Boolean           '# List array of items (Selected/Unselected)

Dim LastListIndex As Integer        '# Last selected item
Dim LastY As Single                 '# Last Y value [pixels] (prevents item repaint)
Dim AnchorItemState As Boolean      '# Anchor item value (multiple selection).
                                    '  Case extended selection: all selected items
                                    '  will be set to Anchor selection state.
Dim EnsureVisibleItem As Boolean    '# Ensure visible last selected item (ListIndex)

Dim ItemRct As RECT                 '# Item rectangle area
Dim tmpItemHeight As Integer        '# Item height [pixels]
Dim VisibleRows As Integer          '# Visible rows in control area
Dim Scrolling As Boolean            '# Scrolling by mouse
Dim YScrolling As Long              '# Y scrolling coordinate flag (scroll speed = f(Y))
Dim HasFocus As Boolean             '# Control has focus

Dim IL As Object                    '# Will point to ImageList control
Dim iScale As Integer               '# ImageList parent scale mode

Dim WithEvents p_Font As StdFont    '# Font object
Attribute p_Font.VB_VarHelpID = -1

Dim cBackNrm As Long                '# Back color [Normal]
Dim cBackSel As Long                '# Back color [Selected]
Dim cFontNrm As Long                '# Font color [Normal]
Dim cFontSel As Long                '# Font color [Selected]
Dim cGrad1 As RGB                   '# Gradient color from [Selected]
Dim cGrad2 As RGB                   '# Gradient color  to  [Selected]
Dim cBox As Long                    '# Box border color

'#################################################################################

Dim UC_Alignment As Alignment
Dim UC_Apeareance As Appearance
Dim UC_BackNormal As OLE_COLOR
Dim UC_BackSelected As OLE_COLOR
Dim UC_BackSelectedG1 As OLE_COLOR
Dim UC_BackSelectedG2 As OLE_COLOR
Dim UC_BoxBorder As OLE_COLOR
Dim UC_BoxOffset As Integer
Dim UC_BoxRadius As Integer
Dim UC_Focus As Boolean
Dim UC_FontNormal As OLE_COLOR
Dim UC_FontSelected As OLE_COLOR
Dim UC_HoverSelection As Boolean
Dim UC_ItemHeight As Integer
Dim UC_ItemHeightAuto As Boolean
Dim UC_ItemOffset As Integer
Dim UC_ItemTextLeft As Integer
Dim UC_ListIndex As Integer
Dim UC_OrderType As OrderType
Dim UC_ScrollBarWidth As Integer
Dim UC_SelectionPicture As Picture
Dim UC_SelectMode As SelectMode
Dim UC_SelectModeStyle As SelectModeStyle
Dim UC_ShowMenu As Boolean
Dim UC_TopIndex As Integer
Dim UC_WordWrap As Boolean

'#################################################################################

Const UCdf_Appearance = 1
Const UCdf_Alignment = DT_LEFT
Const UCdf_BackNormal = vbWindowBackground
Const UCdf_BackSelected = vbHighlight
Const UCdf_BackSelectedG1 = vbHighlight
Const UCdf_BackSelectedG2 = vbWindowBackground
Const UCdf_BorderStyle = 1
Const UCdf_BoxBorder = vbHighlightText
Const UCdf_BoxOffset = 1
Const UCdf_BoxRadius = 0
Const UCdf_Focus = True
Const UCdf_FontNormal = vbWindowText
Const UCdf_FontSelected = vbHighlightText
Const UCdf_HoverSelection = False
Const UCdf_ItemHeightAuto = True
Const UCdf_ItemOffset = 0
Const UCdf_ItemTextLeft = 2
Const UCdf_OrderType = 0
Const UCdf_ScrollBarWidth = 13
Const UCdf_SelectMode = 0
Const UCdf_SelectModeStyle = 0
Const UCdf_ShowMenu = False
Const UCdf_WordWrap = True

'#################################################################################

Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event ListIndexChange()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
Event Scroll()
Event TopIndexChange()




'#################################################################################
'## Init/Read/Write properties
'#################################################################################

Private Sub UserControl_InitProperties()

    UserControl.Appearance = UCdf_Appearance
    UserControl.BorderStyle = UCdf_BorderStyle
    UC_ScrollBarWidth = UCdf_ScrollBarWidth

    Set iUC.Font = Ambient.Font
    Set p_Font = Ambient.Font
    
    UC_FontNormal = UCdf_FontNormal
    UC_FontSelected = UCdf_FontSelected
    UC_BackNormal = UCdf_BackNormal
    UC_BackSelected = UCdf_BackSelected
    UC_BackSelectedG1 = UCdf_BackSelectedG1
    UC_BackSelectedG2 = UCdf_BackSelectedG2
    
    UC_BoxBorder = UCdf_BoxBorder
    UC_BoxOffset = UCdf_BoxOffset
    UC_BoxRadius = UCdf_BoxRadius
    
    UC_Alignment = UCdf_Alignment
    UC_Focus = UCdf_Focus
    UC_HoverSelection = UCdf_HoverSelection
    UC_WordWrap = UCdf_WordWrap
    
    UC_ItemHeight = iUC.TextHeight("TextHeight")
    UC_ItemHeightAuto = UCdf_ItemHeightAuto
    UC_ItemOffset = UCdf_ItemOffset
    UC_ItemTextLeft = UCdf_ItemTextLeft
    
    UC_OrderType = UCdf_OrderType
    Set UC_SelectionPicture = Nothing
    UC_SelectMode = UCdf_SelectMode
    UC_SelectModeStyle = UCdf_SelectModeStyle
    UC_ShowMenu = UCdf_ShowMenu
    
    UC_ListIndex = -1
    UC_TopIndex = -1
    
    SetColors

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Appearance = PropBag.ReadProperty("Appearance", UCdf_Appearance)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", UCdf_BorderStyle)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    UC_ScrollBarWidth = PropBag.ReadProperty("ScrollBarWidth", UCdf_ScrollBarWidth)
    sbUC.Width = PropBag.ReadProperty("ScrollBarWidth", UCdf_ScrollBarWidth)
    
    Set iUC.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set p_Font = PropBag.ReadProperty("Font", Ambient.Font)
    
    UC_FontNormal = PropBag.ReadProperty("FontNormal", UCdf_FontNormal)
    UC_FontSelected = PropBag.ReadProperty("FontSelected", UCdf_FontSelected)
    UC_BackNormal = PropBag.ReadProperty("BackNormal", UCdf_BackNormal)
    iUC.BackColor = PropBag.ReadProperty("BackNormal", UCdf_BackNormal)
    UC_BackSelected = PropBag.ReadProperty("BackSelected", UCdf_BackSelected)
    UC_BackSelectedG1 = PropBag.ReadProperty("BackSelectedG1", UCdf_BackSelectedG1)
    UC_BackSelectedG2 = PropBag.ReadProperty("BackSelectedG2", UCdf_BackSelectedG2)
    
    UC_BoxBorder = PropBag.ReadProperty("BoxBorder", UCdf_BoxBorder)
    UC_BoxOffset = PropBag.ReadProperty("BoxOffset", UCdf_BoxOffset)
    UC_BoxRadius = PropBag.ReadProperty("BoxRadius", UCdf_BoxRadius)
    
    UC_Alignment = PropBag.ReadProperty("Alignment", UCdf_Alignment)
    UC_Focus = PropBag.ReadProperty("Focus", UCdf_Focus)
    UC_HoverSelection = PropBag.ReadProperty("HoverSelection", UCdf_HoverSelection)
    UC_WordWrap = PropBag.ReadProperty("WordWrap", UCdf_WordWrap)

    UC_ItemOffset = PropBag.ReadProperty("ItemOffset", UCdf_ItemOffset)
    UC_ItemHeightAuto = PropBag.ReadProperty("ItemHeightAuto", UCdf_ItemHeightAuto)
    UC_ItemTextLeft = PropBag.ReadProperty("ItemTextLeft", UCdf_ItemTextLeft)

    UC_OrderType = PropBag.ReadProperty("OrderType", UCdf_OrderType)
    Set UC_SelectionPicture = PropBag.ReadProperty("SelectionPicture", Nothing)
    UC_SelectMode = PropBag.ReadProperty("SelectMode", UCdf_SelectMode)
    UC_SelectModeStyle = PropBag.ReadProperty("SelectModeStyle", UCdf_SelectModeStyle)
    UC_ShowMenu = PropBag.ReadProperty("ShowMenu", UCdf_ShowMenu)
    
    iUC.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set iUC.MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    
    Dim tmp As String
    tmp = PropBag.ReadProperty("ItemHeight", 0)
    If tmp < iUC.TextHeight("TextHeight") Then
       UC_ItemHeight = iUC.TextHeight("TextHeight")
    Else
       UC_ItemHeight = tmp
    End If
    
    If UC_SelectMode = [multiple] Then
       optSelectionMenu(0).Enabled = True
       optSelectionMenu(1).Enabled = True
       optSelectionMenu(3).Enabled = True
    End If
    
    UC_ListIndex = -1
    UC_TopIndex = -1

    SetColors

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("ScrollBarWidth", UC_ScrollBarWidth, UCdf_ScrollBarWidth)
    
    Call PropBag.WriteProperty("Font", iUC.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontNormal", UC_FontNormal, UCdf_FontNormal)
    Call PropBag.WriteProperty("FontSelected", UC_FontSelected, UCdf_FontSelected)
    Call PropBag.WriteProperty("BackNormal", UC_BackNormal, UCdf_BackNormal)
    Call PropBag.WriteProperty("BackSelected", UC_BackSelected, UCdf_BackSelected)
    Call PropBag.WriteProperty("BackSelectedG1", UC_BackSelectedG1, UCdf_BackSelectedG1)
    Call PropBag.WriteProperty("BackSelectedG2", UC_BackSelectedG2, UCdf_BackSelectedG2)
    
    Call PropBag.WriteProperty("BoxBorder", UC_BoxBorder, UCdf_BoxBorder)
    Call PropBag.WriteProperty("BoxOffset", UC_BoxOffset, UCdf_BoxOffset)
    Call PropBag.WriteProperty("BoxRadius", UC_BoxRadius, UCdf_BoxRadius)
    
    Call PropBag.WriteProperty("Alignment", UC_Alignment, UCdf_Alignment)
    Call PropBag.WriteProperty("Focus", UC_Focus, UCdf_Focus)
    Call PropBag.WriteProperty("HoverSelection", UC_HoverSelection, UCdf_HoverSelection)
    Call PropBag.WriteProperty("WordWrap", UC_WordWrap, UCdf_WordWrap)
    
    Call PropBag.WriteProperty("ItemHeight", UC_ItemHeight, 0)
    Call PropBag.WriteProperty("ItemHeightAuto", UC_ItemHeightAuto, UCdf_ItemHeightAuto)
    Call PropBag.WriteProperty("ItemOffset", UC_ItemOffset, UCdf_ItemOffset)
    Call PropBag.WriteProperty("ItemTextLeft", UC_ItemTextLeft, UCdf_ItemTextLeft)
    
    Call PropBag.WriteProperty("OrderType", UC_OrderType, UCdf_OrderType)
    Call PropBag.WriteProperty("SelectionPicture", UC_SelectionPicture, Nothing)
    Call PropBag.WriteProperty("SelectMode", UC_SelectMode, UCdf_SelectMode)
    Call PropBag.WriteProperty("SelectModeStyle", UC_SelectModeStyle, UCdf_SelectModeStyle)
    Call PropBag.WriteProperty("ShowMenu", UC_ShowMenu, UCdf_ShowMenu)

    Call PropBag.WriteProperty("MousePointer", iUC.MousePointer, 0)
    Call PropBag.WriteProperty("MouseIcon", iUC.MouseIcon, Nothing)

End Sub

'#################################################################################
'## UserControl initialitation, focus, size, refresh, termination
'#################################################################################

Private Sub UserControl_Initialize()
    
   '# Initialize arrays
    ReDim List(0)
    ReDim Selected(0)
   '# Initialize position flags
    EnsureVisibleItem = True    '# Ensure visible last selected
    LastListIndex = -1          '# Last selected
    LastY = -1                  '# Last Y coordinate
   '# Initialize font object
    Set p_Font = New StdFont
       
End Sub

Private Sub UserControl_EnterFocus()
    HasFocus = True
    DrawFocus UC_ListIndex
End Sub

Private Sub UserControl_ExitFocus()
    HasFocus = False
    DrawItem UC_ListIndex
End Sub

Private Sub UserControl_Resize()
    
   '## Set item height
    If UC_ItemHeightAuto Then
        tmpItemHeight = iUC.TextHeight("TextHeight")
    Else
        If UC_ItemHeight < iUC.TextHeight("TextHeight") Then
           tmpItemHeight = iUC.TextHeight("TextHeight")
        Else
           tmpItemHeight = UC_ItemHeight
        End If
    End If
    
   '## Get visible rows a readjust control height
    VisibleRows = ScaleHeight \ tmpItemHeight
    Height = (VisibleRows) * tmpItemHeight * 15 + (Height - ScaleHeight * 15)
    
   '## Locate and resize drawing area & scroll bar
    iUC.Move 0, 0, ScaleWidth - IIf(sbUC.Visible, sbUC.Width, 0), ScaleHeight
    sbUC.Move ScaleWidth - sbUC.Width, 0, sbUC.Width, ScaleHeight
    ReadjustScrollBar
        
End Sub

Private Sub iUC_Paint()

    If Not Ambient.UserMode Then
        
        iUC.Cls

        Select Case UC_Alignment
               Case 0: iUC.CurrentX = UC_ItemTextLeft + UC_ItemOffset
               Case 1: iUC.CurrentX = (ScaleWidth - iUC.TextWidth(Ambient.DisplayName)) * 0.5
               Case 2: iUC.CurrentX = (ScaleWidth - iUC.TextWidth(Ambient.DisplayName)) - UC_ItemOffset
        End Select
                       iUC.CurrentY = UC_ItemOffset
                    
        SetTextColor iUC.hdc, cFontNrm
        iUC.Print Ambient.DisplayName
        
        Dim IR As RECT
        SetRect IR, 0, 0, ScaleWidth, tmpItemHeight
        DrawFocusRect iUC.hdc, IR

    Else: DrawList
        
    End If
    
End Sub

Private Sub UserControl_Terminate()
    Erase List
    Erase Selected
    Set IL = Nothing
    Scrolling = False
End Sub

'#################################################################################
'## ScrollBar
'#################################################################################

Private Sub sbUC_Change()
    LastY = -1
    If UC_ListIndex = LastListIndex Then DrawList
    RaiseEvent Scroll
End Sub

Private Sub sbUC_Scroll()
    sbUC_Change
    RaiseEvent Scroll
End Sub










'#################################################################################
'## Scrolling/Events
'#################################################################################

'## Click() ----------------------------------------------------------------------
Private Sub iUC_Click()
    If UC_ListIndex > -1 Then RaiseEvent Click
End Sub

'## DblClick() -------------------------------------------------------------------
Private Sub iUC_DblClick()
    If UC_ListIndex > -1 Then RaiseEvent DblClick
End Sub

'## KeyDown(KeyCode, Shift) ------------------------------------------------------
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If UBound(List) = 0 Or UC_ListIndex = -1 Then
        RaiseEvent KeyDown(KeyCode, Shift)
        Exit Sub
    End If
    
    Select Case KeyCode
    
        Case 38 '{Up arrow}
            If UC_ListIndex > 0 Then ListIndex = ListIndex - 1
            
        Case 40 '{Down arrow}
            If UC_ListIndex < UBound(List) - 1 Then ListIndex = ListIndex + 1
            
        Case 33 '{PageDown}
            If UC_ListIndex > VisibleRows Then
                ListIndex = ListIndex - VisibleRows
            Else
                ListIndex = 0
            End If
            
        Case 34 '{PageUp}
            If UC_ListIndex < UBound(List) - VisibleRows - 1 Then
                ListIndex = ListIndex + VisibleRows
            Else
                ListIndex = UBound(List) - 1
            End If
            
        Case 36 '{Start}
            ListIndex = 0
            
        Case 35 '{End}
            ListIndex = UBound(List) - 1
            
        Case 32 '{Space} Select/Unselect
            If UC_SelectMode <> 0 And UC_ListIndex > -1 Then
                Selected(UC_ListIndex) = Not Selected(UC_ListIndex)
                DrawItem UC_ListIndex
                DrawFocus UC_ListIndex
            End If
            RaiseEvent Click
            
    End Select
    
    RaiseEvent KeyDown(KeyCode, Shift)
       
End Sub

'## KeyPress(KeyAscii) -----------------------------------------------------------
Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

'## KeyPress(KeyCode, Shift) -----------------------------------------------------
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'## MouseDown(Button, Shift, X, Y) -----------------------------------------------
Private Sub iUC_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    
    If Button = vbRightButton Then
       If UC_ShowMenu And _
          y >= 0 And y < ScaleHeight Then
         '# Show context menu
          PopupMenu SelectionMenu
       End If
       RaiseEvent MouseDown(Button, Shift, X, y)
       Exit Sub
    End If
   
    Dim SelectedListIndex As Integer
    SelectedListIndex = sbUC + Int(y / tmpItemHeight)
    
    If SelectedListIndex >= 0 And SelectedListIndex < UBound(List) Then
    
        Select Case UC_SelectMode
            Case 0 '[Single]
                    Selected(SelectedListIndex) = True
            Case 1 '[Multiple]
                    Selected(SelectedListIndex) = Not Selected(SelectedListIndex)
                    AnchorItemState = Selected(SelectedListIndex)
        End Select
        
        LastY = y
        ListIndex = SelectedListIndex
        
    End If
    
    Scrolling = True
    RaiseEvent MouseDown(Button, Shift, X, y)
    
End Sub

'## MouseMove(Button, Shift, X, Y) -----------------------------------------------
Private Sub iUC_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    
    YScrolling = y
    
    If y < 0 Then
        ScrollUP
        RaiseEvent MouseMove(Button, Shift, X, y)
        Exit Sub
    End If
    If y > ScaleHeight Then
        ScrollDown
        RaiseEvent MouseMove(Button, Shift, X, y)
        Exit Sub
    End If
                
    If (UC_HoverSelection Or Button) And _
       Int(y / tmpItemHeight) <> Int(LastY / tmpItemHeight) Then
     
        Dim SelectedListIndex As Integer
        SelectedListIndex = sbUC + Int(y / tmpItemHeight)
        
        If SelectedListIndex >= 0 And SelectedListIndex < UBound(List) Then
           Selected(SelectedListIndex) = AnchorItemState
           ListIndex = SelectedListIndex
           LastY = y
        End If
        
    End If
    
    RaiseEvent MouseMove(Button, Shift, X, y)
    
End Sub

'## MouseUp(Button, Shift, X, Y) -------------------------------------------------
Private Sub iUC_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    Scrolling = False
    AnchorItemState = True
    RaiseEvent MouseUp(Button, Shift, X, y)
End Sub










'#################################################################################
'## Methods
'#################################################################################

'## SetImageList -----------------------------------------------------------------
Public Sub SetImageList(ImageListControl)
    
    Set IL = ImageListControl
    On Error Resume Next
    iScale = IL.Parent.ScaleMode
    
    iUC_Paint
    
End Sub

'## AddItem ----------------------------------------------------------------------
Public Sub AddItem(Text As Variant, _
                   Optional Icon As Integer, _
                   Optional IconSelected As Integer)
                   
    '# 0,...,n-1 [n = ListCount]
    
        List(UBound(List)).Text = CStr(Text)
        List(UBound(List)).Icon = Icon
        List(UBound(List)).IconSelected = IconSelected
    
    ReDim Preserve List(0 To UBound(List) + 1)
    ReDim Preserve Selected(0 To UBound(List) + 1)
    
    ReadjustScrollBar
    If UBound(List) < VisibleRows + 1 Then DrawItem UBound(List) - 1
    
End Sub

'## InsertItem -------------------------------------------------------------------
Public Sub InsertItem(Index As Integer, _
                      Text As Variant, _
                      Optional Icon As Integer, _
                      Optional IconSelected As Integer)
     
    If UBound(List) = 0 Or _
       Index > UBound(List) Then Err.Raise 381
       '# (Empty List or "Item added")
    
        ReDim Preserve List(UBound(List) + 1)
        ReDim Preserve Selected(UBound(List))
    
        Dim i As Integer
        For i = UBound(List) - 1 To Index Step -1
           List(i + 1) = List(i)
           Selected(i + 1) = Selected(i)
        Next i
         
        List(Index).Text = CStr(Text)
        List(Index).Icon = Icon
        List(Index).IconSelected = IconSelected
        Selected(Index) = False
        
    ReadjustScrollBar
    EnsureVisibleItem = False
    If UC_ListIndex > -1 And Index <= UC_ListIndex Then ListIndex = ListIndex + 1
    iUC_Paint
    
End Sub

'## ModifyItem -------------------------------------------------------------------
Public Sub ModifyItem(Index As Integer, _
                      Text As Variant, _
                      Optional Icon As Integer = -1, _
                      Optional IconSelected As Integer = -1)
    
    If UBound(List) = 0 Then Err.Raise 381
    '# (Empty List)
    
        List(Index).Text = CStr(Text)
        If Icon > -1 Then List(Index).Icon = Icon
        If IconSelected > -1 Then List(Index).IconSelected = IconSelected
    
    DrawItem Index
    DrawFocus UC_ListIndex

End Sub

'## RemoveItem -------------------------------------------------------------------
Public Sub RemoveItem(Index As Integer)

    If UBound(List) = 0 Or Index > UBound(List) - 1 Then Err.Raise 381 _
    '# (Empty List)
        
        If Index < UBound(List) - 1 Then
            Dim i As Integer
            For i = Index To UBound(List) - 1
               List(i) = List(i + 1)
               Selected(i) = Selected(i + 1)
            Next i
        End If
         
        ReDim Preserve List(UBound(List) - 1)
        ReDim Preserve Selected(UBound(List))
        
    ReadjustScrollBar
    EnsureVisibleItem = False
    
    If Index < UC_ListIndex Then
        If UC_ListIndex > -1 Then ListIndex = ListIndex - 1
    ElseIf Index = UC_ListIndex Then
        ListIndex = -1
    End If
    
    If UBound(List) = 0 Then
        iUC.Cls
    Else
        iUC_Paint
    End If
    
End Sub

'## GetItem ----------------------------------------------------------------------
Public Function GetItem(Index As Integer) As String
    If Index < 0 Then Err.Raise 381
    GetItem = List(Index).Text
End Function

'## IsSelected -------------------------------------------------------------------
Public Function IsSelected(Index As Integer) As Boolean
    If Index < 0 Then Err.Raise 381
    IsSelected = Selected(Index)
End Function

'## FindFirst -------------------------------------------------------------------
Public Function FindFirst(fndString As String, _
                          Optional StartIndex As Integer = 0, _
                          Optional StartWith As Boolean = False) As Integer
    
    If UBound(List) = 0 Then
       Err.Raise 2, , "Empty list"
       Exit Function
    End If
    
    Dim i As Integer
    For i = StartIndex To UBound(List)
        If StartWith Then
            If InStr(1, UCase(List(i).Text), UCase(fndString)) = 1 Then FindFirst = i: Exit Function
        Else
            If InStr(1, UCase(List(i).Text), UCase(fndString)) > 1 Then FindFirst = i: Exit Function
        End If
    Next i
    
   '#fndString not found:
    FindFirst = -1

End Function

'## Clear -------------------------------------------------------------------
Public Sub Clear()
    
   '# Hide scroll bar
    sbUC.Visible = False
   '# Clear and resize drawing area
    iUC.Cls
    iUC.Move 0, 0, ScaleWidth, ScaleHeight
   '# Reset Item arrays
    ReDim List(0)
    ReDim Selected(0)
    
    LastListIndex = -1
    UC_ListIndex = -1
    UC_TopIndex = -1
    
End Sub

'## Order ------------------------------------------------------------------------
Public Sub Order()

    If UBound(List) < 2 Then Exit Sub
    
    If UC_SelectMode = [Single] Then
        If UC_ListIndex > -1 Then Selected(UC_ListIndex) = False
    End If
    
    Dim Index As Integer, Index2 As Integer
    Dim FirstItem As Integer, NumberOfItems As Integer
    Dim Distance As Integer, Value As Item
    Dim Desc As Boolean
    
    FirstItem = LBound(List)
    NumberOfItems = UBound(List)
    If OrderType = Descendent Then Desc = True
    
    Do: Distance = Distance * 3 + 1
    Loop Until Distance > NumberOfItems
    
    Do
        Distance = Distance \ 3
        
        For Index = Distance + FirstItem To NumberOfItems + FirstItem - 1
        
            Value = List(Index)
            Index2 = Index
            
            Do While (List(Index2 - Distance).Text > Value.Text) Xor Desc
                List(Index2) = List(Index2 - Distance)
                Index2 = Index2 - Distance
                If Index2 - Distance < FirstItem Then Exit Do
            Loop
            
            List(Index2) = Value
            
        Next
        
    Loop Until Distance = 1
    
    ListIndex = -1
    sbUC = 0
    
   '# Unselect all and refresh
    ReDim Selected(0 To UBound(List))
    iUC_Paint
    
End Sub

'## SelectItem -------------------------------------------------------------------
Public Sub SelectItem(Index As Integer)
    
    If UC_SelectMode = [Single] Then
        ListIndex = Index
    Else
        Selected(Index) = True
        DrawItem Index
        If Index = UC_ListIndex Then DrawFocus Index
    End If
    
End Sub

'## UnselectItem -----------------------------------------------------------------
Public Sub UnselectItem(Index As Integer)
    
    If UC_SelectMode = [Single] Then
    Else
        Selected(Index) = False
        DrawItem Index
        If Index = UC_ListIndex Then DrawFocus Index
    End If

End Sub










'#################################################################################
'## Draw List / Item / Focus
'#################################################################################

'## DrawList ---------------------------------------------------------------------
Private Sub DrawList()
    
    On Error Resume Next
    If Not Extender.Visible Then Exit Sub
    
    '## Draw visible rows
    Dim i As Integer
    For i = sbUC To sbUC + VisibleRows
        DrawItem i
    Next i
    
    '## Draw focus
    DrawFocus UC_ListIndex

End Sub

'## DrawItem ---------------------------------------------------------------------
Private Sub DrawItem(Index As Integer)

   '# Item out of area ?
    If Index < sbUC Or _
       Index > sbUC + VisibleRows Then Exit Sub
    
        iUC.FontUnderline = False
        
       '##
       '## Define Item area
       '##
        SetRect ItemRct, _
                0, (Index - sbUC) * tmpItemHeight, _
                ItemRct.Right, (Index - sbUC) * tmpItemHeight + tmpItemHeight
       
       '##
       '## Draw selected Item...
       '##
        If Selected(Index) Then
           '# Draw back area
            Select Case UC_SelectModeStyle
                
                Case 0 '[Standard]
                        DrawBack iUC.hdc, ItemRct, cBackSel
                        SetTextColor iUC.hdc, cFontSel
                
                Case 1 '[Dither] *(Effect will be applied after drawing icon)
                        DrawBack iUC.hdc, ItemRct, cBackNrm
                        SetTextColor iUC.hdc, cFontSel
                        
                Case 2 '[Gradient-V]
                        DrawBackGrad iUC.hdc, ItemRct, cGrad1, cGrad2, GRADIENT_FILL_RECT_V
                        SetTextColor iUC.hdc, cFontSel
                        
                Case 3 '[Gradient-H]
                        DrawBackGrad iUC.hdc, ItemRct, cGrad1, cGrad2, GRADIENT_FILL_RECT_H
                        SetTextColor iUC.hdc, cFontSel
                                
                Case 4 '[Box]
                        DrawBack iUC.hdc, ItemRct, cBackNrm
                        DrawBox iUC.hdc, ItemRct, UC_BoxOffset, UC_BoxRadius, cBackSel, cBox
                        SetTextColor iUC.hdc, cFontSel
                                                
                Case 5 '[Underline]
                        DrawBack iUC.hdc, ItemRct, cBackNrm
                        SetTextColor iUC.hdc, cFontSel
                        iUC.FontUnderline = True
                                        
                Case 6 '[byPicture]
                        If Not SelectionPicture Is Nothing Then
                           iUC.PaintPicture SelectionPicture, _
                                            0, ItemRct.Top, _
                                            ItemRct.Right, tmpItemHeight
                        Else
                           DrawBack iUC.hdc, ItemRct, cBackSel
                        End If
                        SetTextColor iUC.hdc, cFontSel
            End Select
           
           '# Draw icon
            If Not IL Is Nothing Then
               On Error Resume Next
               If UC_WordWrap Then
                  IL.ListImages(List(Index).IconSelected).Draw iUC.hdc, _
                     ScaleX(UC_ItemOffset, vbPixels, iScale), _
                     ScaleY(ItemRct.Top + UC_ItemOffset, vbPixels, iScale), 1
               Else
                  IL.ListImages(List(Index).IconSelected).Draw iUC.hdc, _
                     ScaleX(UC_ItemOffset, vbPixels, iScale), _
                     ScaleY(ItemRct.Top + (tmpItemHeight - IL.ImageHeight) * 0.5, vbPixels, iScale), 1
               End If
            End If
           '# If case, apply dither effect (*)
            If UC_SelectModeStyle = 1 Then DrawDither iUC.hdc, ItemRct, cBackSel
       
       '##
       '## Draw unselected Item...
       '##
        Else
           '# Draw back area
            DrawBack iUC.hdc, ItemRct, cBackNrm
            SetTextColor iUC.hdc, cFontNrm
            
           '# Draw icon
            If Not IL Is Nothing Then
               On Error Resume Next
               If UC_WordWrap Then
                  IL.ListImages(List(Index).Icon).Draw iUC.hdc, _
                     ScaleX(UC_ItemOffset, vbPixels, iScale), _
                     ScaleY(ItemRct.Top + UC_ItemOffset, vbPixels, iScale), 1
               Else
                  IL.ListImages(List(Index).Icon).Draw iUC.hdc, _
                     ScaleX(UC_ItemOffset, vbPixels, iScale), _
                     ScaleY(ItemRct.Top + (tmpItemHeight - IL.ImageHeight) * 0.5, vbPixels, iScale), 1
               End If
            End If
        End If
       
       '##
       '## Draw text...
       '##
        On Error Resume Next
        If UC_WordWrap Then
               SetRect ItemRct, _
                       UC_ItemOffset + UC_ItemTextLeft, ItemRct.Top + UC_ItemOffset, _
                       ItemRct.Right - UC_ItemOffset, ItemRct.Bottom
              DrawText iUC.hdc, _
                       List(Index).Text, Len(List(Index).Text), _
                       ItemRct, _
                       UC_Alignment Or DT_WORDBREAK
               SetRect ItemRct, _
                       0, ItemRct.Top - UC_ItemOffset, _
                       ItemRct.Right + UC_ItemOffset, ItemRct.Bottom
        Else
               SetRect ItemRct, _
                       UC_ItemOffset + UC_ItemTextLeft, ItemRct.Top, _
                       ItemRct.Right - UC_ItemOffset, ItemRct.Bottom
              DrawText iUC.hdc, _
                       List(Index).Text, Len(List(Index).Text), _
                       ItemRct, _
                       DT_SINGLELINE Or DT_VCENTER
               SetRect ItemRct, _
                       0, ItemRct.Top - UC_ItemOffset, _
                       ItemRct.Right + UC_ItemOffset, ItemRct.Bottom
        End If
         
End Sub

'## DrawFocus --------------------------------------------------------------------
Private Sub DrawFocus(Index As Integer)
    
    If Not UC_Focus Or Not HasFocus Then Exit Sub
    
   '# Item out of area ?
    If Index < sbUC Or _
       Index > sbUC + VisibleRows Then Exit Sub
       
       '# Draw it
        SetRect ItemRct, _
                0, (Index - sbUC) * tmpItemHeight, _
                ItemRct.Right, (Index - sbUC) * tmpItemHeight + tmpItemHeight

        SetTextColor iUC.hdc, cBackSel
        DrawFocusRect iUC.hdc, ItemRct

End Sub

'## ReadjustScrollBar ------------------------------------------------------------
Private Sub ReadjustScrollBar()
     
    If UBound(List) > VisibleRows Then
    
        If Not sbUC.Visible Then
           '# Show scroll bar
            sbUC.Visible = True
            sbUC.Refresh
            On Error Resume Next
            sbUC.LargeChange = VisibleRows
           '# Update item area rectangle right margin
            ItemRct.Right = ScaleWidth - sbUC.Width
           '# Repaint control area
            iUC_Paint
        End If
        
    Else
    
      '# Hide scroll bar
       sbUC.Visible = False
      '# Update item area rectangle right margin
       ItemRct.Right = ScaleWidth
       
    End If
    
   '# Update sbUC max value
    sbUC.Max = UBound(List) - VisibleRows
       
End Sub











'#################################################################################
'## Scroll Up/Down by mouse & multiple select
'#################################################################################

'## ScrollUP ---------------------------------------------------------------------
Private Sub ScrollUP()

    Dim t As Long     '# Timer counter
    Dim Delay As Long '# Scrolling delay
    
    Delay = 500 + 20 * YScrolling
    If Delay < 40 Then Delay = 40
    
   '# Scroll while MouseDown and mouse pos. < "Control top"
    Do While Scrolling And YScrolling < 0
        If GetTickCount - t > Delay Then
            t = GetTickCount
            If UC_ListIndex > 0 Then
                If UC_SelectMode = [multiple] Then
                    Selected(UC_ListIndex - 1) = AnchorItemState
                End If
                ListIndex = ListIndex - 1
            End If
        End If
        DoEvents
    Loop

End Sub

'## ScrollDown -------------------------------------------------------------------
Private Sub ScrollDown()

    Dim t As Long     '# Timer counter
    Dim Delay As Long '# Scrolling delay
    
    Delay = 500 - 20 * (YScrolling - ScaleHeight - 1)
    If Delay < 40 Then Delay = 40
    
   '# Scroll while MouseDown and mouse pos. > "Control bottom"
    Do While Scrolling And YScrolling > ScaleHeight - 1
        If GetTickCount - t > Delay Then
            t = GetTickCount
            If UC_ListIndex < UBound(List) - 1 Then
                If UC_SelectMode = [multiple] Then
                    Selected(UC_ListIndex + 1) = AnchorItemState
                End If
                ListIndex = ListIndex + 1
            End If
        End If
        DoEvents
    Loop

End Sub











'#################################################################################
'## Colors
'#################################################################################

'## SetColors --------------------------------------------------------------------
Private Sub SetColors()
    
    '## Item back color [Normal]
    cBackNrm = GetLngColor(UC_BackNormal)

    '## Item back color [Selected]
    cBackSel = GetLngColor(UC_BackSelected)
    
    '## Item back color 1 [Selected (Gradient style)]
    cGrad1 = GetRGBColors(GetLngColor(UC_BackSelectedG1))
    
    '## Item back color 2 [Selected (Gradient style)]
    cGrad2 = GetRGBColors(GetLngColor(UC_BackSelectedG2))
    
    '## Item box border color [Selected (Box style)]
    cBox = GetLngColor(UC_BoxBorder)
    
    '## Item font color [Normal]
    cFontNrm = GetLngColor(UC_FontNormal)
    
    '## Item font color [Selected]
    cFontSel = GetLngColor(UC_FontSelected)
    
End Sub

'## GetLngColor ------------------------------------------------------------------
Private Function GetLngColor(Color As Long) As Long

    If Color And &H80000000 Then
       GetLngColor = GetSysColor(Color And &H7FFFFFFF)
    Else
       GetLngColor = Color
    End If

End Function

'## GetRGBColors -----------------------------------------------------------------
Private Function GetRGBColors(Color As Long) As RGB

    Dim HexColor As String
        
    HexColor = String(6 - Len(Hex(Color)), "0") & Hex(Color)
    GetRGBColors.R = "&H" & Mid(HexColor, 5, 2) & "00"
    GetRGBColors.G = "&H" & Mid(HexColor, 3, 2) & "00"
    GetRGBColors.B = "&H" & Mid(HexColor, 1, 2) & "00"

End Function










'#################################################################################
'## Context menu
'#################################################################################

Private Sub optSelectionMenu_Click(Index As Integer)
    
    Dim i As Integer
    Select Case Index
    
        Case 0 '# Select all
                For i = 0 To UBound(List) - 1
                    Selected(i) = True
                Next i
            
        Case 1 '# Unselect all
                'For i = 0 To UBound(List) - 1
                '    Selected(i) = False
                'Next i
                ReDim Selected(0 To UBound(List))
    
        Case 3 '# Invert
                For i = 0 To UBound(List) - 1
                    Selected(i) = Not Selected(i)
                Next i
    End Select
    iUC_Paint

End Sub










'#################################################################################
'## Properties
'#################################################################################

'## Alignment --------------------------------------------------------------------
Public Property Get Alignment() As Alignment
    Alignment = UC_Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As Alignment)
    UC_Alignment = New_Alignment
    iUC_Paint
    PropertyChanged "Alignment"
End Property

'## Appearance -------------------------------------------------------------------
Public Property Get Appearance() As Appearance
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Appearance)
    UserControl.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

'## BackNormal -------------------------------------------------------------------
Public Property Get BackNormal() As OLE_COLOR
    BackNormal = UC_BackNormal
End Property

Public Property Let BackNormal(ByVal New_BackNormal As OLE_COLOR)
    UC_BackNormal = New_BackNormal
    cBackNrm = GetLngColor(UC_BackNormal)
    iUC.BackColor = cBackNrm
    iUC_Paint
    PropertyChanged "BackNormal"
End Property

'## BackSelected -----------------------------------------------------------------
Public Property Get BackSelected() As OLE_COLOR
    BackSelected = UC_BackSelected
End Property

Public Property Let BackSelected(ByVal New_BackSelected As OLE_COLOR)
    
    UC_BackSelected = New_BackSelected
    cBackSel = GetLngColor(UC_BackSelected)
    iUC_Paint
    PropertyChanged "BackSelected"
    
End Property

'## BackSelectedG1 ---------------------------------------------------------------
Public Property Get BackSelectedG1() As OLE_COLOR
    BackSelectedG1 = UC_BackSelectedG1
End Property

Public Property Let BackSelectedG1(ByVal New_BackSelectedG1 As OLE_COLOR)
    UC_BackSelectedG1 = New_BackSelectedG1
    cGrad1 = GetRGBColors(GetLngColor(UC_BackSelectedG1))
    iUC_Paint
    PropertyChanged "BackSelectedG1"
End Property

'## BackSelectedG2 ---------------------------------------------------------------
Public Property Get BackSelectedG2() As OLE_COLOR
    BackSelectedG2 = UC_BackSelectedG2
End Property

Public Property Let BackSelectedG2(ByVal New_BackSelectedG2 As OLE_COLOR)
    UC_BackSelectedG2 = New_BackSelectedG2
    cGrad2 = GetRGBColors(GetLngColor(UC_BackSelectedG2))
    iUC_Paint
    PropertyChanged "BackSelectedG2"
End Property

'## BorderStyle ------------------------------------------------------------------
Public Property Get BorderStyle() As BorderStyle
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyle)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'## BoxBorder --------------------------------------------------------------------
Public Property Get BoxBorder() As OLE_COLOR
    BoxBorder = UC_BoxBorder
End Property

Public Property Let BoxBorder(ByVal New_BoxBorder As OLE_COLOR)
    UC_BoxBorder = New_BoxBorder
    cBox = GetLngColor(UC_BoxBorder)
    iUC_Paint
    PropertyChanged "BoxBorder"
End Property

'## BoxOffset --------------------------------------------------------------------
Public Property Get BoxOffset() As Integer
    BoxOffset = UC_BoxOffset
End Property

Public Property Let BoxOffset(ByVal New_BoxOffset As Integer)
    If New_BoxOffset <= tmpItemHeight * 0.5 Then
        UC_BoxOffset = New_BoxOffset
    End If
    iUC_Paint
    PropertyChanged "BoxOffset"
End Property

'## BoxRadius --------------------------------------------------------------------
Public Property Get BoxRadius() As Integer
    BoxRadius = UC_BoxRadius
End Property

Public Property Let BoxRadius(ByVal New_BoxRadius As Integer)
    UC_BoxRadius = New_BoxRadius
    iUC_Paint
    PropertyChanged "BoxRadius"
End Property

'## Enabled ----------------------------------------------------------------------
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    sbUC.Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

'## Focus ------------------------------------------------------------------------
Public Property Get Focus() As Boolean
    Focus = UC_Focus
End Property

Public Property Let Focus(ByVal New_Focus As Boolean)
    UC_Focus = New_Focus
    If New_Focus Then DrawFocus UC_ListIndex Else DrawItem UC_ListIndex
    PropertyChanged "Focus"
End Property

'## Font -------------------------------------------------------------------------
Public Property Get Font() As Font
    Set Font = p_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    With p_Font
        .Name = New_Font.Name
        .Size = New_Font.Size
        .Bold = New_Font.Bold
        .Italic = New_Font.Italic
        .Underline = New_Font.Underline
        .Strikethrough = New_Font.Strikethrough
    End With
    iUC_Paint
    PropertyChanged "Font"
End Property

Private Sub p_Font_FontChanged(ByVal PropertyName As String)
    Set iUC.Font = p_Font
    UserControl_Resize
End Sub

'## FontNormal -------------------------------------------------------------------
Public Property Get FontNormal() As OLE_COLOR
    FontNormal = UC_FontNormal
End Property

Public Property Let FontNormal(ByVal New_FontNormal As OLE_COLOR)
    UC_FontNormal = New_FontNormal
    cFontNrm = GetLngColor(UC_FontNormal)
    SetTextColor iUC.hdc, cFontNrm
    iUC_Paint
    PropertyChanged "FontNormal"
End Property

'## FontSelected -----------------------------------------------------------------
Public Property Get FontSelected() As OLE_COLOR
    FontSelected = UC_FontSelected
End Property

Public Property Let FontSelected(ByVal New_FontSelected As OLE_COLOR)
    UC_FontSelected = New_FontSelected
    cFontSel = GetLngColor(UC_FontSelected)
    iUC_Paint
    PropertyChanged "FontSelected"
End Property

'## HoverSelection ---------------------------------------------------------------
Public Property Get HoverSelection() As Boolean
    HoverSelection = UC_HoverSelection
End Property

Public Property Let HoverSelection(ByVal New_HoverSelection As Boolean)
    UC_HoverSelection = New_HoverSelection
    DrawItem UC_ListIndex
    DrawFocus UC_ListIndex
    PropertyChanged "HoverSelection"
End Property

'## ItemHeight -------------------------------------------------------------------
Public Property Get ItemHeight() As Integer
    ItemHeight = UC_ItemHeight
End Property

Public Property Let ItemHeight(ByVal New_ItemHeight As Integer)
    UC_ItemHeight = New_ItemHeight
    UserControl_Resize
    iUC_Paint
    PropertyChanged "ItemHeight"
End Property

'## ItemHeightAuto ---------------------------------------------------------------
Public Property Get ItemHeightAuto() As Boolean
    ItemHeightAuto = UC_ItemHeightAuto
End Property

Public Property Let ItemHeightAuto(ByVal New_ItemHeightAuto As Boolean)
    UC_ItemHeightAuto = New_ItemHeightAuto
    UserControl_Resize
    iUC_Paint
    PropertyChanged "ItemHeightAuto"
End Property

'## ItemOffset -------------------------------------------------------------------
Public Property Get ItemOffset() As Integer
    ItemOffset = UC_ItemOffset
End Property

Public Property Let ItemOffset(ByVal New_ItemOffset As Integer)
    If New_ItemOffset <= tmpItemHeight Then
        UC_ItemOffset = New_ItemOffset
    End If
    iUC_Paint
    PropertyChanged "ItemOffset"
End Property

'## ItemTextLeft -----------------------------------------------------------------
Public Property Get ItemTextLeft() As Integer
    ItemTextLeft = UC_ItemTextLeft
End Property

Public Property Let ItemTextLeft(ByVal New_ItemTextLeft As Integer)
    UC_ItemTextLeft = New_ItemTextLeft
    iUC_Paint
    PropertyChanged "ItemTextLeft"
End Property

'## <ListCount> ------------------------------------------------------------------
Public Property Get ListCount() As Integer
    ListCount = UBound(List)
End Property

'## ListIndex --------------------------------------------------------------------
Public Property Get ListIndex() As Integer
Attribute ListIndex.VB_MemberFlags = "400"
    ListIndex = UC_ListIndex
End Property

Public Property Let ListIndex(ByVal New_ListIndex As Integer)
    
    If New_ListIndex < -1 Or New_ListIndex > UBound(List) - 1 Then Err.Raise 380
    
    If New_ListIndex < 0 Or UBound(List) = 0 Then
        UC_ListIndex = -1
        LastY = -1
    Else
        UC_ListIndex = New_ListIndex
    End If
    
   '# Unselect last / Select actual [Single selection mode]
    If UC_SelectMode = [Single] Then
        If LastListIndex > -1 Then Selected(LastListIndex) = False
        If UC_ListIndex > -1 Then Selected(UC_ListIndex) = True
    End If
    
   '# Draw last (delete Focus) ...
    DrawItem LastListIndex
    LastListIndex = UC_ListIndex
   '# ... and draw actual (draw Focus)
    DrawItem UC_ListIndex
    DrawFocus UC_ListIndex

   '# Ensure visible actual selected item
    If EnsureVisibleItem Then
        If UC_ListIndex < sbUC And UC_ListIndex > -1 Then
           sbUC = UC_ListIndex
        ElseIf UC_ListIndex >= sbUC + VisibleRows Then
           sbUC = UC_ListIndex - VisibleRows + 1
        End If
    Else
        EnsureVisibleItem = True
    End If
    
    PropertyChanged "ListIndex"
    RaiseEvent ListIndexChange

End Property

'## MouseIcon --------------------------------------------------------------------
Public Property Get MouseIcon() As Picture
    Set MouseIcon = iUC.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set iUC.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'## MousePointer -----------------------------------------------------------------
Public Property Get MousePointer() As MousePointerConstants
    MousePointer = iUC.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    iUC.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'## OrderType --------------------------------------------------------------------
Public Property Get OrderType() As OrderType
    OrderType = UC_OrderType
End Property

Public Property Let OrderType(ByVal New_OrderType As OrderType)
    UC_OrderType = New_OrderType
    PropertyChanged "OrderType"
End Property

'## ScrollBarWidth ---------------------------------------------------------------
Public Property Get ScrollBarWidth() As Integer
    ScrollBarWidth = UC_ScrollBarWidth
End Property

Public Property Let ScrollBarWidth(ByVal New_ScrollBarWidth As Integer)
    
   '# Check Min value width...
    If New_ScrollBarWidth < 9 Then
        UC_ScrollBarWidth = 9
        sbUC.Width = 9
   '# Check Max value width...
    ElseIf New_ScrollBarWidth > ScaleWidth * 0.5 Then
        UC_ScrollBarWidth = ScaleWidth * 0.5
        sbUC.Width = ScaleWidth * 0.5
   '# Set new value...
    Else
        UC_ScrollBarWidth = New_ScrollBarWidth
        sbUC.Width = New_ScrollBarWidth
    End If
    sbUC.Visible = False
    ReadjustScrollBar
    UserControl_Resize
    
    PropertyChanged "ScrollBarWidth"
    
End Property

'## <SelectedCount> --------------------------------------------------------------
Public Property Get SelectedCount() As Integer
    
    Dim i As Integer
    
    SelectedCount = 0
    For i = 0 To UBound(List)
        If Selected(i) Then SelectedCount = SelectedCount + 1
    Next i
    
End Property

'## SelectionPicture -------------------------------------------------------------
Public Property Get SelectionPicture() As Picture
    Set SelectionPicture = UC_SelectionPicture
End Property

Public Property Set SelectionPicture(ByVal New_SelectionPicture As Picture)
    Set UC_SelectionPicture = New_SelectionPicture
    iUC_Paint
    PropertyChanged "SelectionPicture"
End Property

'## SelectMode -------------------------------------------------------------------
Public Property Get SelectMode() As SelectMode
    SelectMode = UC_SelectMode
End Property

Public Property Let SelectMode(ByVal New_SelectMode As SelectMode)
    
    UC_SelectMode = New_SelectMode
    
    If Ambient.UserMode Then
        If New_SelectMode = [Single] Then
           '# Unselect all and select actual
            If UC_ListIndex > -1 Then
               Dim i As Integer
               For i = LBound(List) To UBound(List)
                   If i <> UC_ListIndex Then Selected(i) = False
               Next i
               Selected(UC_ListIndex) = True
               DrawItem UC_ListIndex
               DrawFocus UC_ListIndex
            End If
           '# Disable selection menu items
            optSelectionMenu(0).Enabled = False
            optSelectionMenu(1).Enabled = False
            optSelectionMenu(3).Enabled = False
        Else
           '# Enable selection menu items
            optSelectionMenu(0).Enabled = True
            optSelectionMenu(1).Enabled = True
            optSelectionMenu(3).Enabled = True
        End If
   End If
   
   ReadjustScrollBar
   iUC_Paint
   
   PropertyChanged "SelectMode"

End Property

'## SelectModeStyle --------------------------------------------------------------
Public Property Get SelectModeStyle() As SelectModeStyle
    SelectModeStyle = UC_SelectModeStyle
End Property

Public Property Let SelectModeStyle(ByVal New_SelectModeStyle As SelectModeStyle)
    UC_SelectModeStyle = New_SelectModeStyle
    iUC_Paint
    PropertyChanged "SelectModeStyle"
End Property

'## ShowMenu ---------------------------------------------------------------------
Public Property Get ShowMenu() As Boolean
    ShowMenu = UC_ShowMenu
End Property

Public Property Let ShowMenu(ByVal New_ShowMenu As Boolean)
    UC_ShowMenu = New_ShowMenu
    PropertyChanged "ShowMenu"
End Property

'## TopIndex ---------------------------------------------------------------------
Public Property Get TopIndex() As Integer
Attribute TopIndex.VB_MemberFlags = "400"
    TopIndex = sbUC
End Property

Public Property Let TopIndex(ByVal New_TopIndex As Integer)
    
    If New_TopIndex < 0 Or _
       New_TopIndex > UBound(List) - VisibleRows Then Err.Raise 380

    UC_TopIndex = New_TopIndex
    sbUC = New_TopIndex
    
    PropertyChanged "TopIndex"
    RaiseEvent TopIndexChange
        
End Property

'## WordWrap ---------------------------------------------------------------------
Public Property Get WordWrap() As Boolean
    WordWrap = UC_WordWrap
End Property

Public Property Let WordWrap(ByVal New_WordWrap As Boolean)
    UC_WordWrap = New_WordWrap
    iUC_Paint
    PropertyChanged "WordWrap"
End Property





'                                  *  *  *  *  *

