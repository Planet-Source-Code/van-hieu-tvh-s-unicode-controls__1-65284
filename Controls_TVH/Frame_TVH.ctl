VERSION 5.00
Begin VB.UserControl Frame_TVH 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2955
   ControlContainer=   -1  'True
   FillStyle       =   0  'Solid
   ScaleHeight     =   129
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   197
   ToolboxBitmap   =   "Frame_TVH.ctx":0000
End
Attribute VB_Name = "Frame_TVH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'//////////////////////////////// Code by Truong Van Hieu ///////////////////////////////////////////////////
'////////////////////////////////// tvhhh2003@yahoo.com ////////////////////////////////////////////////////
'////////////////////////////////// Special for Vietnamese ////////////////////////////////////////////////////

Option Explicit

Enum E_FrameStyle
    vhFrame_Standard = 0
    vhFrame_Standard2 = 1
    vhFrame_Gradiant = 2
    vhFrame_JCStyle = 3
    vhFrame_Windows = 4
    vhFrame_Messenger = 5
    vhFrame_None = 6
End Enum

Enum E_FrameThemeColor
    vhFrameThemeBlue = 0
    vhFrameThemeSilver = 1
    vhFrameThemeOlive = 2
    vhFrameThemeVisual2005 = 3
    vhFrameThemeNorton2005 = 5
End Enum

Enum E_FrameAlignmentText
    vhFrameAlign_Left = 0
    vhFrameAlign_Right = 1
    vhFrameAlign_Center = 2
End Enum

Private b_ChangeProperty As Boolean

Private b_ColorFrom As OLE_COLOR
Private b_ColorTo As OLE_COLOR

Private m_Caption As String
Private m_TiengViet As Boolean
Private m_FrameStyle As E_FrameStyle
Private m_AlignmentText As E_FrameAlignmentText
Private m_ShadowText As Boolean
Private m_ShadowColor As OLE_COLOR
Private m_ThemeColor As E_FrameThemeColor ' Only for --> JC_Style, Messenger Style
Private m_BorderColor As OLE_COLOR ' Only for --> Windows Style
Private m_BackTitleColor As OLE_COLOR ' Only for --> Windows Style
Private m_ForeTitleColor As OLE_COLOR
Private m_BackFrameColor As OLE_COLOR ' Only for --> Windows Style, Standard1, Standard2
Private m_TitleHeight As Integer
Private m_RoundCorner As Boolean
Private m_BackColor As OLE_COLOR
Private m_BorderGradientWidth As Long
Private m_GradientOutColor As OLE_COLOR
Private m_GradientInColor As OLE_COLOR

Event Click()
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Property Get Font() As StdFont
    Set Font = UserControl.Font
End Property

Public Property Set Font(new_Font As StdFont)
    Set UserControl.Font = new_Font
    PropertyChanged "Font"
    Refresh
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(new_Caption As String)
    m_Caption = new_Caption
    PropertyChanged "Caption"
    Refresh
End Property

Public Property Get TiengViet() As Boolean
    TiengViet = m_TiengViet
End Property

Public Property Let TiengViet(new_TiengViet As Boolean)
    m_TiengViet = new_TiengViet
    PropertyChanged "TiengViet"
    Refresh
End Property

Public Property Get FrameStyle() As E_FrameStyle
    FrameStyle = m_FrameStyle
End Property

Public Property Let FrameStyle(new_FrameStyle As E_FrameStyle)
    m_FrameStyle = new_FrameStyle
    PropertyChanged "FrameStyle"
    b_ChangeProperty = True
    Select Case m_FrameStyle
        Case vhFrame_JCStyle, vhFrame_Messenger
            SetThemeColor m_ThemeColor
        Case vhFrame_Windows
            BackFrameColor = &HE0FFFF
            BackTitleColor = &HB0EFF0
            BorderColor = 0
    End Select
    b_ChangeProperty = False
    Refresh
End Property

Public Property Get AlignmentText() As E_FrameAlignmentText
    AlignmentText = m_AlignmentText
End Property

Public Property Let AlignmentText(new_AlignmentText As E_FrameAlignmentText)
    m_AlignmentText = new_AlignmentText
    PropertyChanged "AlignmentText"
    Refresh
End Property

Public Property Get ShadowText() As Boolean
    ShadowText = m_ShadowText
End Property

Public Property Let ShadowText(new_ShadowText As Boolean)
    m_ShadowText = new_ShadowText
    PropertyChanged "ShadowText"
    Refresh
End Property

Public Property Get ShadowColor() As OLE_COLOR
    ShadowColor = m_ShadowColor
End Property

Public Property Let ShadowColor(new_ShadowColor As OLE_COLOR)
    m_ShadowColor = new_ShadowColor
    PropertyChanged "ShadowColor"
    Refresh
End Property

Public Property Get ThemeColor() As E_FrameThemeColor
    ThemeColor = m_ThemeColor
End Property

Public Property Let ThemeColor(new_ThemeColor As E_FrameThemeColor)
    m_ThemeColor = new_ThemeColor
    SetThemeColor m_ThemeColor
    PropertyChanged "ThemeColor"
    Refresh
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(new_BorderColor As OLE_COLOR)
    m_BorderColor = new_BorderColor
    PropertyChanged "BorderColor"
    Refresh
End Property

Public Property Get BackTitleColor() As OLE_COLOR
    BackTitleColor = m_BackTitleColor
End Property

Public Property Let BackTitleColor(new_BackTitleColor As OLE_COLOR)
    m_BackTitleColor = new_BackTitleColor
    PropertyChanged "BackTitleColor"
    Refresh
End Property

Public Property Get ForeTitleColor() As OLE_COLOR
    ForeTitleColor = m_ForeTitleColor
End Property

Public Property Let ForeTitleColor(new_ForeTitleColor As OLE_COLOR)
    m_ForeTitleColor = new_ForeTitleColor
    PropertyChanged "ForeTitleColor"
    Refresh
End Property

Public Property Get BackFrameColor() As OLE_COLOR
    BackFrameColor = m_BackFrameColor
End Property

Public Property Let BackFrameColor(new_BackFrameColor As OLE_COLOR)
    m_BackFrameColor = new_BackFrameColor
    PropertyChanged "BackFrameColor"
    Refresh
End Property

Public Property Get TitleHeight() As Integer
    TitleHeight = m_TitleHeight
End Property

Public Property Let TitleHeight(new_TitleHeight As Integer)
    m_TitleHeight = new_TitleHeight
    PropertyChanged "TitleHeight"
    Refresh
End Property

Public Property Get RoundCorner() As Boolean
    RoundCorner = m_RoundCorner
End Property

Public Property Let RoundCorner(new_RoundCorner As Boolean)
    m_RoundCorner = new_RoundCorner
    PropertyChanged "RoundCorner"
    Refresh
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(new_BackColor As OLE_COLOR)
    m_BackColor = new_BackColor
    PropertyChanged "BackColor"
    UserControl.BackColor = m_BackColor
    Refresh
End Property

Public Property Get BorderGradientWidth() As Long
    BorderGradientWidth = m_BorderGradientWidth
End Property

Public Property Let BorderGradientWidth(new_BorderGradientWidth As Long)
    m_BorderGradientWidth = new_BorderGradientWidth
    PropertyChanged "BorderGradientWidth"
    Refresh
End Property

Public Property Get GradientOutColor() As OLE_COLOR
    GradientOutColor = m_GradientOutColor
End Property

Public Property Let GradientOutColor(new_GradientOutColor As OLE_COLOR)
    m_GradientOutColor = new_GradientOutColor
    PropertyChanged "GradientOutColor"
    Refresh
End Property

Public Property Get GradientInColor() As OLE_COLOR
    GradientInColor = m_GradientInColor
End Property

Public Property Let GradientInColor(new_GradientInColor As OLE_COLOR)
    m_GradientInColor = new_GradientInColor
    PropertyChanged "GradientInColor"
    Refresh
End Property

Public Property Get Hwnd() As Long
    Hwnd = UserControl.Hwnd
End Property

Sub Refresh()
Dim r As RECT
Dim h&, w&, tLeft&
With UserControl
    If b_ChangeProperty Then Exit Sub
    .Cls
    Dim s As String
    s = IIf(m_TiengViet, VNI_Unicode(m_Caption), m_Caption)
    h = TextHeight(s) + 2
    w = UniTextWidth(.hdc, s)
    Select Case m_FrameStyle
        Case vhFrame_Standard
            b_ChangeProperty = True
            BackColor = m_BackFrameColor
            b_ChangeProperty = False
            
            Line (0, h \ 2)-(.ScaleWidth - 1, h \ 2), RGB(167, 166, 170)
            Line (0, h \ 2)-(0, .ScaleHeight - 1), RGB(167, 166, 170)
            Line (.ScaleWidth - 1, h \ 2)-(.ScaleWidth - 1, .ScaleHeight), RGB(255, 255, 255)
            Line (0, .ScaleHeight - 1)-(.ScaleWidth, .ScaleHeight - 1), RGB(255, 255, 255)

            Line (1, h \ 2 + 1)-(.ScaleWidth - 2, h \ 2 + 1), RGB(255, 255, 255)
            Line (1, h \ 2 + 1)-(1, .ScaleHeight - 2), RGB(255, 255, 255)
            Line (.ScaleWidth - 2, h \ 2 + 1)-(.ScaleWidth - 2, .ScaleHeight - 1), RGB(167, 166, 170)
            Line (1, .ScaleHeight - 2)-(.ScaleWidth - 1, .ScaleHeight - 2), RGB(167, 166, 170)
            
            If m_AlignmentText = vhFrameAlign_Left Then
                tLeft = 6
            ElseIf m_AlignmentText = vhFrameAlign_Right Then
                tLeft = .ScaleWidth - w - 10
            Else
                tLeft = .ScaleWidth \ 2 - w \ 2 - 2
            End If
            
            Line (tLeft, h \ 2)-(tLeft + w + 4, h \ 2), m_BackFrameColor
            Line (tLeft, h \ 2 + 1)-(tLeft + w + 4, h \ 2 + 1), m_BackFrameColor
            
            SetRect r, 7, 1, .ScaleWidth - 9, h - 1
        Case vhFrame_Standard2
            .BackColor = m_BackFrameColor
            .Forecolor = RGB(255, 255, 255)
            Rectangle .hdc, 0, h \ 2 - 1, .ScaleWidth, .ScaleHeight
            Rectangle .hdc, 2, h \ 2 + 1, .ScaleWidth - 2, .ScaleHeight - 2
            .Forecolor = RGB(167, 166, 170)
            Rectangle .hdc, 1, h \ 2, .ScaleWidth - 1, .ScaleHeight - 1
            
            If m_AlignmentText = vhFrameAlign_Left Then
                tLeft = 6
            ElseIf m_AlignmentText = vhFrameAlign_Right Then
                tLeft = .ScaleWidth - w - 10
            Else
                tLeft = .ScaleWidth \ 2 - w \ 2 - 2
            End If
            
            Line (tLeft, h \ 2 - 1)-(tLeft + w + 4, h \ 2 - 1), m_BackFrameColor
            Line (tLeft, h \ 2)-(tLeft + w + 4, h \ 2), m_BackFrameColor
            Line (tLeft, h \ 2 + 1)-(tLeft + w + 4, h \ 2 + 1), m_BackFrameColor
            
            SetRect r, 7, 1, .ScaleWidth - 9, h - 1
        Case vhFrame_Gradiant
            .FillStyle = 0
            .FillColor = m_BackFrameColor
            .Forecolor = m_BorderColor
            If m_RoundCorner Then
                RoundRect .hdc, 0&, 0&, .ScaleWidth, .ScaleHeight, 10&, 10&
            Else
                RoundRect .hdc, 0&, 0&, .ScaleWidth, .ScaleHeight, 0&, 0&
            End If
            
            .FillStyle = 1
            
            SetRect r, 0, 0, .ScaleWidth, m_TitleHeight
            DrawGradBorderRect vbWhite, m_GradientInColor, r, m_GradientOutColor, False
            
            DrawGradOutBorder m_GradientOutColor, m_GradientInColor, m_BorderGradientWidth
            SetRect r, 0, m_TitleHeight - m_BorderGradientWidth, .ScaleWidth, m_TitleHeight
            DrawGradLine2 m_GradientOutColor, m_GradientInColor, r
            
            SetRect r, m_BorderGradientWidth + 4, m_BorderGradientWidth, .ScaleWidth - m_BorderGradientWidth - 4, m_TitleHeight - m_BorderGradientWidth
        Case vhFrame_JCStyle
            .FillStyle = 0
            m_BackFrameColor = BlendColors(b_ColorFrom, vbWhite)
            .FillColor = m_BackFrameColor
            .Forecolor = m_BorderColor
            If m_RoundCorner Then
                RoundRect .hdc, 0&, 0&, .ScaleWidth, .ScaleHeight, 10&, 10&
            Else
                RoundRect .hdc, 0&, 0&, .ScaleWidth, .ScaleHeight, 0&, 0&
            End If
            .FillStyle = 1
            SetRect r, 0, 0, .ScaleWidth, m_TitleHeight
            DrawGradBorderRect b_ColorTo, b_ColorFrom, r, m_BorderColor

            SetRect r, 0, 0, .ScaleWidth, 4
            DrawGradBorderRect b_ColorTo, b_ColorFrom, r, m_BorderColor

            SetRect r, 0, m_TitleHeight - 4, .ScaleWidth, m_TitleHeight
            DrawGradBorderRect b_ColorTo, b_ColorFrom, r, m_BorderColor
            
            SetRect r, 1, m_TitleHeight, .ScaleWidth - 1, .ScaleHeight - .ScaleHeight * 0.2
            DrawGradBorderRect b_ColorTo, m_BackFrameColor, r, m_BorderColor, False
            
            SetRect r, 6, 4, .ScaleWidth - 6, m_TitleHeight - 4
        Case vhFrame_Windows
            .FillStyle = 0
            .FillColor = m_BackFrameColor
            .Forecolor = m_BorderColor
            If m_RoundCorner Then
                RoundRect .hdc, 0&, 0&, .ScaleWidth, .ScaleHeight, 10&, 10&
            Else
                RoundRect .hdc, 0&, 0&, .ScaleWidth, .ScaleHeight, 0&, 0&
            End If
            .FillColor = m_BackTitleColor
            Rectangle .hdc, 0, 0, .ScaleWidth, m_TitleHeight
            .FillStyle = 1
            
            SetRect r, 6, 2, .ScaleWidth - 6, m_TitleHeight - 2
        Case vhFrame_Messenger
            .FillStyle = 0
            m_BackFrameColor = BlendColors(b_ColorFrom, vbWhite)
            .FillColor = m_BackFrameColor
            .Forecolor = m_BorderColor
            
            If m_RoundCorner Then
                RoundRect .hdc, 0&, 0&, .ScaleWidth, .ScaleHeight, 10&, 10&
            Else
                RoundRect .hdc, 0&, 0&, .ScaleWidth, .ScaleHeight, 0&, 0&
            End If
            
            .FillStyle = 1
                        
            SetRect r, 0, 0, .ScaleWidth, 7
            DrawGradBorderRect vbWhite, b_ColorFrom, r, m_BorderColor
            DrawPilots r, BlendColors(vbBlack, b_ColorFrom)
            
            Line (0, m_TitleHeight - 1)-(.ScaleWidth, m_TitleHeight - 1), m_BorderColor
            
            SetRect r, 1, m_TitleHeight, .ScaleWidth - 1, .ScaleHeight - .ScaleHeight * 0.2
            DrawGradBorderRect b_ColorTo, m_BackFrameColor, r, m_BorderColor, False
            SetRect r, 6, 7, .ScaleWidth - 6, m_TitleHeight - 2
        Case Else
            SetRect r, 4, 2, .ScaleWidth - 4, m_TitleHeight + 2
    End Select
    If m_ShadowText Then
        Offset r, 1
        .Forecolor = m_ShadowColor
        DrawText .hdc, s, r
        Offset r, -1
    End If
    .Forecolor = m_ForeTitleColor
    DrawText .hdc, s, r
End With
End Sub

Private Function DrawText(hdc As Long, s As String, tRect As RECT) As Long
    If Trim(s) = "" Then Exit Function
    Dim Flags As Long
    Flags = DT_NOCLIP Or DT_VCENTER Or DT_SINGLELINE
    Flags = Flags Or IIf(m_AlignmentText = vhFrameAlign_Left, 0, IIf(m_AlignmentText = vhFrameAlign_Right, DT_RIGHT, DT_CENTER))
    DrawText = DrawTextW(hdc, StrPtr(s), Len(s), tRect, Flags)
End Function


Private Sub DrawOutline(hdc As Long, s As String, tRect As RECT, Size As Integer)
    Dim tx As Integer, ty As Integer
    Dim t As RECT
    t = tRect
    For ty = tRect.Top - Size To tRect.Top + Size
        For tx = tRect.Left - Size To tRect.Left + Size
            t.Right = t.Right - (t.Left - tx)
            t.Bottom = t.Bottom - (t.Top - ty)
            t.Left = tx
            t.Top = ty
            DrawText hdc, s, t
        Next tx
    Next ty
End Sub

Private Sub DrawShadow(s As String, tR As RECT, Color_S As Long, Color_E As Long, Depth As Integer, Style As E_ShadowStyle)
Dim AColor() As Long
Dim t As RECT
Dim i As Integer, dx As Integer, dy As Integer
    t = tR
    GradientColor Color_E, Color_S, Depth + 1, AColor
    Select Case Style
        Case 0: dx = -1: dy = -1 'LeftTop
        Case 1: dx = 1: dy = -1 'RightTop
        Case 2: dx = -1: dy = 1 'LeftBottom
        Case 3: dx = 1: dy = 1  'RightBottom
    End Select
    For i = Depth To 1 Step -1
        UserControl.Forecolor = AColor(Depth - i)
        t.Left = tR.Left + i * dx
        t.Top = tR.Top + i * dy
        t.Right = tR.Right + i * dx
        t.Bottom = tR.Bottom + i * dy
        DrawText UserControl.hdc, s, t
    Next i
End Sub

Private Sub SetThemeColor(bTheme As E_FrameThemeColor, Optional iRefresh As Boolean = False)
    b_ChangeProperty = Not (iRefresh)
    Select Case bTheme
'        Case vhFrameThemeBlue
'            b_ColorFrom = RGB(129, 169, 226)
'            b_ColorTo = RGB(221, 236, 254)
'            m_BorderColor = RGB(0, 0, 128)
        Case vhFrameThemeSilver
            b_ColorFrom = RGB(153, 151, 180)
            b_ColorTo = RGB(244, 244, 251)
            BorderColor = RGB(75, 75, 111)
        Case vhFrameThemeOlive
            b_ColorFrom = RGB(181, 197, 143)
            b_ColorTo = RGB(247, 249, 225)
            BorderColor = RGB(63, 93, 56)
        Case vhFrameThemeVisual2005
            b_ColorFrom = RGB(194, 194, 171)
            b_ColorTo = RGB(248, 248, 242)
            BorderColor = RGB(145, 145, 115)
        Case vhFrameThemeNorton2005
            b_ColorFrom = RGB(217, 172, 1)
            b_ColorTo = RGB(255, 239, 165)
            BorderColor = RGB(117, 91, 30)
        Case Else
            b_ColorFrom = RGB(129, 169, 226)
            b_ColorTo = RGB(221, 236, 254)
            BorderColor = RGB(0, 0, 128)
    End Select
    BackFrameColor = BlendColors(b_ColorFrom, b_ColorFrom)
    b_ChangeProperty = False
End Sub

Private Function BlendColors(ByVal lcolor1 As Long, ByVal lcolor2 As Long)
    BlendColors = RGB(((lcolor1 And &HFF) + (lcolor2 And &HFF)) / 2, (((lcolor1 \ &H100) And &HFF) + ((lcolor2 \ &H100) And &HFF)) / 2, (((lcolor1 \ &H10000) And &HFF) + ((lcolor2 \ &H10000) And &HFF)) / 2)
End Function

Private Sub SetForeColor(dc&, Color&)
Dim t&
    t = CreatePen(0, 1, Color)
    DeleteObject SelectObject(dc, t)
    DeleteObject t
End Sub

Private Sub DrawGradLine(bColorOut As OLE_COLOR, bColorIn As OLE_COLOR, r As RECT)
Dim i As Integer, c() As Long, C2(2) As Long
    C2(0) = bColorOut
    C2(1) = bColorIn
    C2(2) = bColorOut
    SplitGradientColor C2, r.Bottom - r.Top, c
    For i = 0 To r.Bottom - r.Top - 1
        Line (r.Left, r.Top + i)-(r.Right, r.Top + i), c(i)
    Next i
End Sub

Private Sub DrawGradLine2(bColorOut As OLE_COLOR, bColorIn As OLE_COLOR, r As RECT)
Dim i As Integer, c() As Long, C2(2) As Long, w&
    w = r.Bottom - r.Top
    C2(0) = bColorOut
    C2(1) = bColorIn
    C2(2) = bColorOut
    SplitGradientColor C2, r.Bottom - r.Top, c
    For i = 0 To (r.Bottom - r.Top + 1) \ 2 - 1
        Line (r.Left + w - i, r.Top + i)-(r.Right - w + i, r.Top + i), c(w - i - 1)
        Line (r.Left + w - i, r.Bottom - i - 1)-(r.Right - w + i, r.Bottom - i - 1), c(w - i - 1)
    Next i
    
End Sub

Private Sub DrawGradOutBorder(bColorOut As OLE_COLOR, bColorIn As OLE_COLOR, bThick As Long)
    Dim tc&(2), tc2&(), i
    tc(0) = bColorOut: tc(1) = bColorIn: tc(2) = bColorOut
    SplitGradientColor tc, bThick, tc2
    For i = 0 To bThick - 1
        Forecolor = tc2(i)
        Rectangle UserControl.hdc, 0 + i, 0 + i, UserControl.ScaleWidth - i, UserControl.ScaleHeight - i
    Next i
End Sub

Private Sub DrawGradBorderRect(bSColor As Long, bEColor As Long, r As RECT, bBorderColor As Long, Optional iBorder As Boolean = True)
Dim i As Integer
Dim c() As Long
With UserControl
    GradientColor2 bSColor, bEColor, r.Bottom - r.Top, c
    For i = 0 To r.Bottom - r.Top - 1
        Line (r.Left, r.Top + i)-(r.Right, r.Top + i), c(i)
    Next i
    If iBorder Then
        .Forecolor = bBorderColor
        Rectangle .hdc, r.Left, r.Top, r.Right, r.Bottom
    End If
End With
End Sub

Private Sub DrawPilots(r As RECT, c As OLE_COLOR, Optional Num As Integer = 9)
Dim i As Integer, X&, Y&
With UserControl
    .DrawWidth = 2
    .Forecolor = c
    X = (r.Right - r.Left) \ 2 - (Num * 4 - 1) \ 2
    Y = (r.Bottom - r.Top) \ 2 - 1
    For i = 0 To Num - 1
        PSet (X + i * 4 + 1, Y + 1), vbWhite
        PSet (X + i * 4, Y), c
    Next i
    .DrawWidth = 1
End With
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
    b_ChangeProperty = True
    TitleHeight = 29
    FrameStyle = vhFrame_JCStyle
    BorderGradientWidth = 7
    GradientOutColor = vbBlack
    GradientInColor = vbWhite
    BackFrameColor = vbWhite
    AlignmentText = vhFrameAlign_Center
    Font.Name = "Arial"
    TiengViet = True
    Caption = "Khung"
    
    b_ChangeProperty = False
    Refresh
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
    Set UserControl.Font = .ReadProperty("Font", Parent.Font)
    m_Caption = .ReadProperty("Caption", "Khung")
    m_TiengViet = .ReadProperty("TiengViet", True)
    m_FrameStyle = .ReadProperty("FrameStyle", vhFrame_Standard)
    m_AlignmentText = .ReadProperty("AlignmentText", vhFrameAlign_Left)
    m_ShadowText = .ReadProperty("ShadowText", True)
    m_ShadowColor = .ReadProperty("ShadowColor", vbWhite)
    m_ThemeColor = .ReadProperty("ThemeColor", vhFrameThemeBlue)
    
    If m_FrameStyle = vhFrame_JCStyle Or m_FrameStyle = vhFrame_Messenger Then
        SetThemeColor m_ThemeColor
    End If
    
    m_BorderColor = .ReadProperty("BorderColor", vbBlack)
    m_BackTitleColor = .ReadProperty("BorderColor", &HC0FFFF)
    m_ForeTitleColor = .ReadProperty("ForeTitleColor", vbBlack)
    m_BackFrameColor = .ReadProperty("BackFrameColor", &HE0FFFF)
    m_TitleHeight = .ReadProperty("TitleHeight", 29)
    m_RoundCorner = .ReadProperty("RoundCorner", True)
    m_BackColor = .ReadProperty("BackColor", &H8000000F)
    m_BorderGradientWidth = .ReadProperty("BorderGradientWidth", 7)
    m_GradientOutColor = .ReadProperty("GradientOutColor", vbBlack)
    m_GradientInColor = .ReadProperty("GradientInColor", vbWhite)
    
    Refresh
End With
End Sub

Private Sub UserControl_Resize()
    Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    .WriteProperty "Font", UserControl.Font, Parent.Font
    .WriteProperty "Caption", m_Caption, "Khung"
    .WriteProperty "TiengViet", m_TiengViet, True
    .WriteProperty "FrameStyle", m_FrameStyle, vhFrame_Standard
    .WriteProperty "AlignmentText", m_AlignmentText, vhFrameAlign_Left
    .WriteProperty "ShadowText", m_ShadowText, True
    .WriteProperty "ShadowColor", m_ShadowColor, vbWhite
    .WriteProperty "ThemeColor", m_ThemeColor, vhFrameThemeBlue
    .WriteProperty "BorderColor", m_BorderColor, vbBlack
    .WriteProperty "BackTitleColor", m_BackTitleColor, &HC0FFFF
    .WriteProperty "ForeTitleColor", m_ForeTitleColor, vbBlack
    .WriteProperty "BackFrameColor", m_BackFrameColor, &HE0FFFF
    .WriteProperty "TitleHeight", m_TitleHeight, 29
    .WriteProperty "RoundCorner", m_RoundCorner, True
    .WriteProperty "BackColor", m_BackColor, &H8000000F
    .WriteProperty "BorderGradientWidth", m_BorderGradientWidth, 7
    .WriteProperty "GradientOutColor", m_GradientOutColor, vbBlack
    .WriteProperty "GradientInColor", m_GradientInColor, vbWhite
End With
End Sub

Public Sub About()
Attribute About.VB_UserMemId = -552
On Error Resume Next
    fAbout.Show 1
End Sub

