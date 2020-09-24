Attribute VB_Name = "ModItem"
Option Explicit
'#################################################################################
'## Item effects
'#################################################################################


Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal y As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Function SelectObject Lib "gdi32" _
                       (ByVal hdc As Long, _
                        ByVal hObject As Long) As Long

Public Declare Function DeleteObject Lib "gdi32" _
                       (ByVal hObject As Long) As Long
                       
Public Declare Function GetSysColor Lib "user32" _
                       (ByVal nIndex As Long) As Long

Public Declare Function CreatePen Lib "gdi32" _
                       (ByVal nPenStyle As Long, _
                        ByVal nWidth As Long, _
                        ByVal crColor As Long) As Long
                        
Public Const PS_SOLID = 0

Public Declare Function CreateSolidBrush Lib "gdi32" _
                       (ByVal crColor As Long) As Long
                        
Public Declare Function SetTextColor Lib "gdi32" _
                       (ByVal hdc As Long, _
                        ByVal crColor As Long) As Long

Public Declare Function SetRect Lib "user32" _
                       (lpRect As RECT, _
                        ByVal X1 As Long, ByVal Y1 As Long, _
                        ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Declare Function RoundRect Lib "gdi32" _
                       (ByVal hdc As Long, _
                        ByVal X1 As Long, ByVal Y1 As Long, _
                        ByVal X2 As Long, ByVal Y2 As Long, _
                        ByVal X3 As Long, ByVal Y3 As Long) As Long

Public Declare Function FillRect Lib "user32" _
                       (ByVal hdc As Long, _
                        lpRect As RECT, _
                        ByVal hBrush As Long) As Long
                        
Public Declare Function PatBlt Lib "gdi32" _
                       (ByVal hdc As Long, _
                        ByVal X As Long, ByVal y As Long, _
                        ByVal nWidth As Long, ByVal nHeight As Long, _
                        ByVal dwRop As Long) As Long

Public Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" _
                       (ByVal hdc As Long, _
                        pVertex As TRIVERTEX, _
                        ByVal dwNumVertex As Long, _
                        pMesh As GRADIENT_RECT, _
                        ByVal dwNumMesh As Long, _
                        ByVal dwMode As Long) As Long

Public Type TRIVERTEX
            X As Long
            y As Long
            R As Integer
            G As Integer
            B As Integer
            Alpha As Integer
End Type

Public Type RGB
            R As Integer
            G As Integer
            B As Integer
End Type

Public Type GRADIENT_RECT
            UpperLeft As Long
            LowerRight As Long  '
End Type

Public Const GRADIENT_FILL_RECT_H As Long = &H0
Public Const GRADIENT_FILL_RECT_V  As Long = &H1

Public Declare Function DrawText Lib "user32" Alias "DrawTextA" _
                       (ByVal hdc As Long, _
                        ByVal lpStr As String, ByVal nCount As Long, _
                        lpRect As RECT, _
                        ByVal wFormat As Long) As Long
               
Public Const DT_CENTER = &H1
Public Const DT_LEFT = &H0
Public Const DT_RIGHT = &H2
Public Const DT_VCENTER = &H4
Public Const DT_WORDBREAK = &H10
Public Const DT_SINGLELINE = &H20

Public Declare Function DrawFocusRect Lib "user32" _
                       (ByVal hdc As Long, _
                        lpRect As RECT) As Long

Public Declare Function InflateRect Lib "user32" _
                       (lpRect As RECT, _
                        ByVal dx As Long, ByVal dy As Long) As Long

Public Type RECT
            Left As Long
            Top As Long
            Right As Long
            Bottom As Long
End Type

'
'## Paint item back area (Standard)
'
Public Sub DrawBack(ByVal hdc As Long, _
                    R As RECT, _
                    ByVal Color As Long)

    Dim hBrush As Long
    Dim Ret As Long
    
    hBrush = CreateSolidBrush(Color)
    Ret = FillRect(hdc, R, hBrush)
    Ret = DeleteObject(hBrush)

End Sub
'
'## Dither effect
'
Public Sub DrawDither(ByVal hdc As Long, _
                      R As RECT, _
                      ByVal Color As Long)

    Dim hBrush As Long
    Dim Ret As Long
    
    hBrush = CreateSolidBrush(Color)
    hBrush = SelectObject(hdc, hBrush)

    PatBlt hdc, R.Left, _
                R.Top, _
                R.Right - R.Left, _
                R.Bottom - R.Top, _
                &HA000C9
        
    Ret = DeleteObject(hBrush)

End Sub
'
'## Paint item back area (Gradient)
'
Public Sub DrawBackGrad(ByVal hdc As Long, _
                        R As RECT, _
                        Color1 As RGB, _
                        Color2 As RGB, _
                        Direction As Long)

    Dim V(1) As TRIVERTEX
    Dim GRct As GRADIENT_RECT
    
    '# from
    With V(0)
        .X = R.Left
        .y = R.Top
        .R = Color1.R
        .G = Color1.G
        .B = Color1.B
        .Alpha = 0
    End With
    '# to
    With V(1)
        .X = R.Right
        .y = R.Bottom
        .R = Color2.R
        .G = Color2.G
        .B = Color2.B
        .Alpha = 0
    End With
    
    GRct.UpperLeft = 0
    GRct.LowerRight = 1

    GradientFillRect hdc, V(0), 2, GRct, 1, Direction

End Sub
'
'## Paint box
'
Public Sub DrawBox(ByVal hdc As Long, _
                   R As RECT, _
                   ByVal Offset As Integer, _
                   ByVal Radius As Integer, _
                   ByVal Color1 As Long, _
                   ByVal Color2 As Long)

    Dim hPen As Long
    Dim hBrush As Long
    Dim Ret As Long
    
    hBrush = CreateSolidBrush(Color1)
    hBrush = SelectObject(hdc, hBrush)
    hPen = CreatePen(PS_SOLID, 1, Color2)
    hPen = SelectObject(hdc, hPen)
    
    InflateRect R, -Offset, -Offset
    RoundRect hdc, R.Left, _
                   R.Top, _
                   R.Right, _
                   R.Bottom, _
                   Radius, Radius
    InflateRect R, Offset, Offset
    
    Ret = DeleteObject(hPen)
    Ret = DeleteObject(hBrush)
    
End Sub



