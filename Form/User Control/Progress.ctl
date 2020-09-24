VERSION 5.00
Begin VB.UserControl Progress 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "Progress.ctx":0000
   Begin VB.Label lblCaption 
      Alignment       =   2  '¸m¤¤¹ï»ô
      BackStyle       =   0  '³z©ú
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   2775
   End
   Begin VB.Shape Rec 
      BorderColor     =   &H80000009&
      BorderStyle     =   0  '³z©ú
      FillColor       =   &H8000000D&
      FillStyle       =   0  '¹ê¤ß
      Height          =   855
      Left            =   360
      Top             =   1200
      Width           =   2295
   End
End
Attribute VB_Name = "Progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event Change(ByVal Percentage As Long)


Private Type TypeRGB
    R As Integer
    G As Integer
    B As Integer
End Type

Private _
    PMax As Long, _
    PMin As Long, _
    PValue As Long, _
    FromCol As Long, _
    ToCol As Long

Private Sub UserControl_Initialize()
Rec.Width = UserControl.Width
UserControl.Height = 20 * Screen.TwipsPerPixelY
Rec.Height = UserControl.Height

End Sub

Public Property Let Max(lng As Long)
PMax = lng
End Property

Public Property Get Max() As Long
Max = PMax
End Property

Public Property Let Min(lng As Long)
PMin = lng
End Property

Public Property Get Min() As Long
Min = PMin
End Property

Public Property Let Value(lng As Long)

PValue = lng

If Value > Me.Max Or Value < Me.Min Then Exit Property

Dim Percent As Long
If lng = Me.Min Then GoTo 2
Percent = Int((lng - Me.Min) / (Me.Max - Me.Min) * 100)
2
RaiseEvent Change(Percent)

'The Main code is there, Color-Fading!!!
If Me.Max = Me.Min Then GoTo 3
Rec.Width = (lng - Me.Min) * (UserControl.Width \ (Me.Max - Me.Min))
3
Dim RGBColor1 As TypeRGB, RGBColor2 As TypeRGB
Dim RGBFinal As TypeRGB

RGBColor1 = RGB2TypeRGB(Me.FromColor)
RGBColor2 = RGB2TypeRGB(Me.ToColor)

With RGBFinal
    .R = RGBColor1.R + ((lng - 1) * ((RGBColor2.R - RGBColor1.R) \ (Me.Max - 1)))
    .G = RGBColor1.G + ((lng - 1) * ((RGBColor2.G - RGBColor1.G) \ (Me.Max - 1)))
    .B = RGBColor1.B + ((lng - 1) * ((RGBColor2.B - RGBColor1.B) \ (Me.Max - 1)))
    
        Rec.FillColor = RGB(.R, .G, .B)
    
End With

End Property

Public Property Get Value() As Long
Value = PValue
End Property

Public Property Set Font(fnt As StdFont)
Set lblCaption.Font = fnt
End Property

Public Property Get Font() As StdFont
Set Font = lblCaption.Font
End Property

Public Property Let Caption(str As String)
lblCaption.Caption = str
End Property

Public Property Get Caption() As String
Caption = lblCaption.Caption
End Property

Public Property Let FromColor(Col As OLE_COLOR)
FromCol = Col
End Property

Public Property Get FromColor() As OLE_COLOR
FromColor = FromCol
End Property

Public Property Let ToColor(Col As OLE_COLOR)
ToCol = Col
End Property

Public Property Get ToColor() As OLE_COLOR
ToColor = ToCol
End Property

Private Function RGB2TypeRGB(ByVal Color As Long) As TypeRGB
    With RGB2TypeRGB
    .R = Color Mod 256
    .G = (Color \ 256) Mod 256
    .B = Color \ 65536
    End With
End Function

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With Me

    .Caption = PropBag.ReadProperty("Caption", "")
    .Font = PropBag.ReadProperty("Font", lblCaption.Font)
    .FromColor = PropBag.ReadProperty("FromColor", vbBlue)
    .ToColor = PropBag.ReadProperty("ToColor", vbBlue)
    .Max = PropBag.ReadProperty("Max", 100)
    .Min = PropBag.ReadProperty("Min", 0)
    .Value = PropBag.ReadProperty("Value", 0)

End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag

    .WriteProperty "Caption", Me.Caption
    .WriteProperty "Font", Me.Font
    .WriteProperty "FromColor", Me.FromColor
    .WriteProperty "ToColor", Me.ToColor
    .WriteProperty "Max", Me.Max
    .WriteProperty "Min", Me.Min
    .WriteProperty "Value", Me.Value
    
End With
End Sub

Private Sub UserControl_Resize()
Rec.Move 0, 0
Rec.Height = UserControl.Height
lblCaption.Move 0, 0, UserControl.Width, UserControl.Height
End Sub


