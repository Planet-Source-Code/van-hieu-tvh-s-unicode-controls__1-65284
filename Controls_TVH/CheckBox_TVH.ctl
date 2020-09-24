VERSION 5.00
Begin VB.UserControl CheckBox_TVH 
   AutoRedraw      =   -1  'True
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1695
   ScaleHeight     =   22
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   113
   ToolboxBitmap   =   "CheckBox_TVH.ctx":0000
   Begin VB.PictureBox TPic 
      Height          =   135
      Left            =   1935
      ScaleHeight     =   75
      ScaleWidth      =   90
      TabIndex        =   0
      Top             =   225
      Visible         =   0   'False
      Width           =   150
   End
End
Attribute VB_Name = "CheckBox_TVH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'///////////////////////////////////////// Truong Van Hieu ////////////////////////////////////////
'////////////////////////////////// tvhhh2003@yahoo.com /////////////////////////////////////
'//////////////////////////////////// Special for Vietnamese /////////////////////////////////////
Option Explicit

Const TransColor = &H8000000F

Private Enum E_CheckStatus
    eNormal = 0
    eGotFocus = 1
    eMoveOver = 2
    eClickDown = 3
    eDisabled = 4
End Enum

Enum E_AlignmentCheckBox
    ecbLeft = 0
    ecbRight = 1
    ecbTop = 2
    ecbBottom = 3
End Enum

Private Const d = 13
Private Const D_Box_Text = 3

Private bFocus As Boolean
Private MouseEvent As E_MouseEvent
Private bPrevButton&

Private m_TiengViet As Boolean
Private m_Caption As String
Private m_Forecolor As OLE_COLOR
Private m_BackColor As OLE_COLOR
Private m_Transparent As Boolean
Private m_Alignment As E_AlignmentCheckBox
Private m_Shadow As Boolean
Private m_ShadowColor As OLE_COLOR
Private m_Enabled As Boolean
Private m_Checked As Boolean
Private m_TooltipTiengViet As Boolean
Event Click()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseLeave(Button As Integer, Shift As Integer, x As Single, y As Single)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)


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

Public Property Get Forecolor() As OLE_COLOR
    Forecolor = m_Forecolor
End Property

Public Property Let Forecolor(new_Forecolor As OLE_COLOR)
    m_Forecolor = new_Forecolor
    PropertyChanged "Forecolor"
    Refresh
End Property

Public Property Get Backcolor() As OLE_COLOR
    Backcolor = m_BackColor
End Property

Public Property Let Backcolor(new_Backcolor As OLE_COLOR)
    m_BackColor = new_Backcolor
    PropertyChanged "Backcolor"
    Refresh
End Property

Public Property Get Transparent() As Boolean
    Transparent = m_Transparent
End Property

Public Property Let Transparent(new_Transparent As Boolean)
    m_Transparent = new_Transparent
    PropertyChanged "Transparent"
    Refresh
End Property

Public Property Get Alignment() As E_AlignmentCheckBox
    Alignment = m_Alignment
End Property

Public Property Let Alignment(new_Alignment As E_AlignmentCheckBox)
    m_Alignment = new_Alignment
    PropertyChanged "Alignment"
    Refresh
End Property

Public Property Get Font() As StdFont
    Set Font = UserControl.Font
End Property

Public Property Set Font(new_Font As StdFont)
    Set UserControl.Font = new_Font
    PropertyChanged "Font"
    Refresh
End Property

Public Property Get Shadow() As Boolean
    Shadow = m_Shadow
End Property

Public Property Let Shadow(new_Shadow As Boolean)
    m_Shadow = new_Shadow
    PropertyChanged "Shadow"
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

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(new_Enabled As Boolean)
    m_Enabled = new_Enabled
    PropertyChanged "Enabled"
    Refresh
End Property

Public Property Get Checked() As Boolean
    Checked = m_Checked
End Property

Public Property Let Checked(new_Checked As Boolean)
    m_Checked = new_Checked
    PropertyChanged "Checked"
    Refresh
End Property

Public Property Get TooltipTiengViet() As Boolean
    TooltipTiengViet = m_TooltipTiengViet
End Property

Public Property Let TooltipTiengViet(new_TooltipTiengViet As Boolean)
    m_TooltipTiengViet = new_TooltipTiengViet
    PropertyChanged "ToolTipTiengViet"
End Property

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

'-------------------------------------------------------------------------------------------------------------------------------------

Sub Refresh()
Dim st As E_CheckStatus
With UserControl
    .Cls
    .Backcolor = IIf(m_Transparent, TransColor, m_BackColor)
    Set .Picture = Nothing
    If MouseEvent = eMouseLeaving Then
        st = eNormal
    ElseIf MouseEvent = eMouseLeavingClicking Or MouseEvent = eMouseMoving Then
        st = eMoveOver
    ElseIf MouseEvent = eMouseMovingClicking Then
        st = eClickDown
    Else
       Exit Sub
    End If
    .Enabled = m_Enabled
    DrawCheckBoxXp2005 IIf(m_Enabled, st, eDisabled)
    DrawCaption
    If m_Transparent Then
        .MaskColor = TransColor
        .MaskPicture = .Image
        .BackStyle = 0
    Else
        .BackStyle = 1
    End If
End With
End Sub

Private Sub DrawCaption()
    Dim s As String
    Dim t As RECT
    Dim Flag&
    Const D_Edge_Text = 1
With UserControl
    Select Case m_Alignment
        Case ecbLeft:
            t.Left = d + D_Box_Text
            t.Right = .ScaleWidth - D_Edge_Text
            t.Top = 0
            t.Bottom = .ScaleHeight
            Flag = DT_LEFT
        Case ecbRight:
            t.Left = D_Edge_Text
            t.Right = .ScaleWidth - d - D_Box_Text
            t.Top = 0
            t.Bottom = .ScaleHeight
            Flag = DT_RIGHT
        Case ecbTop:
            t.Left = D_Edge_Text
            t.Right = .ScaleWidth - D_Edge_Text
            t.Top = d
            t.Bottom = .ScaleHeight
            Flag = DT_CENTER
        Case ecbBottom:
            t.Left = D_Edge_Text
            t.Right = .ScaleWidth - D_Edge_Text
            t.Top = 0
            t.Bottom = .ScaleHeight - d
            Flag = DT_CENTER
    End Select
    Set TPic.Font = UserControl.Font
    s = IIf(m_TiengViet, mUnicode.VNI_Unicode(m_Caption), m_Caption)
    Dim dong&
    dong = DrawTextW(TPic.hdc, StrPtr(s), Len(s), t, Flag Or DT_NOCLIP Or DT_WORDBREAK)
    t.Top = t.Top + (t.Bottom - t.Top - dong) \ 2
    If m_Shadow And m_Enabled Then
        Offset t, 1
        .Forecolor = m_ShadowColor
        DrawTextW hdc, StrPtr(s), Len(s), t, Flag Or DT_NOCLIP Or DT_WORDBREAK
        Offset t, -1
    End If
    .Forecolor = IIf(m_Enabled, m_Forecolor, RGB(167, 166, 170))
    DrawTextW hdc, StrPtr(s), Len(s), t, Flag Or DT_NOCLIP Or DT_WORDBREAK
    If bFocus Then
        t.Left = t.Left - 1
        t.Top = t.Top - 2
        t.Bottom = t.Top + dong + 3
        If dong / TextHeight(Left(s, 1)) = 1 Then
            t.Right = t.Left + TextWidthW(hdc, s) + 2
        Else
            t.Right = t.Right + 1
        End If
        .Forecolor = 0
        DrawFocusRect hdc, t
    End If
End With
End Sub

Private Sub DrawCheckBoxXp2005(c As E_CheckStatus, Optional iCheck As Byte = 0)
'iCheck =1 : Checked
'iCheck =2 : UnChecked
Dim y&, x&
Dim i As Byte
Dim ArrC&()
With UserControl
    If iCheck <> 1 And iCheck <> 2 Then
        iCheck = IIf(m_Checked, 1, 2)
    End If
    Select Case m_Alignment
        Case ecbLeft:
            x = 0
            y = (.ScaleHeight - d) \ 2
        Case ecbRight:
            x = .ScaleWidth - d
            y = (.ScaleHeight - d) \ 2
        Case ecbTop:
            x = (.ScaleWidth - d) \ 2
            y = 0
        Case ecbBottom:
            x = (.ScaleWidth - d) \ 2
            y = .ScaleHeight - d
    End Select
    Select Case c
        Case eNormal:
            PSet (x + 1, y + 1), RGB(226, 226, 221)
            Line (x + 1, y + 2)-(x + 3, y), RGB(226, 226, 221)
            Line (x + 1, y + 3)-(x + 4, y), RGB(226, 226, 221)
            
            PSet (x + d - 2, y + d - 2), RGB(255, 255, 255)
            Line (x + d - 3, y + d - 2)-(x + d - 1, y + d - 4), RGB(255, 255, 255)
            Line (x + d - 4, y + d - 2)-(x + d - 1, y + d - 5), RGB(255, 255, 255)
            
            GradientColor2 RGB(226, 226, 221), RGB(255, 255, 255), (d - 4) * 2 - 1, ArrC
            
            For i = 0 To d - 6
                Line (x + 1, y + i + 4)-(x + i + 5, y), ArrC(i + 1)
                Line (x + 1 + i, y + d - 2)-(x + d - 1, y + i), ArrC(d - 5 + i)
            Next i
        Case eMoveOver:
            GradientColor2 RGB(255, 240, 207), RGB(248, 179, 48), (d - 2) * 2 - 1, ArrC
            For i = 0 To d - 3
                Line (x + 1, y + i + 1)-(x + i + 2, y), ArrC(i)
                Line (x + 1 + i, y + d - 2)-(x + d - 1, y + i), ArrC(d - 3 + i)
            Next i
            For i = 0 To 6
                Line (x + (d - 7) \ 2, y + (d - 7) \ 2 + i)-(x + (d - 7) \ 2 + 7, y + (d - 7) \ 2 + i), RGB(247, 247, 245)
            Next i
        Case eClickDown:
            GradientColor2 RGB(176, 176, 167), RGB(241, 239, 223), (d - 2) * 2 - 1, ArrC
            For i = 0 To d - 3
                Line (x + 1, y + i + 1)-(x + i + 2, y), ArrC(i)
                Line (x + 1 + i, y + d - 2)-(x + d - 1, y + i), ArrC(d - 3 + i)
            Next i
        Case eDisabled:
            .Forecolor = RGB(198, 197, 201)
            .FillStyle = 0
            .FillColor = vbWhite
            Rectangle hdc, x, y, x + d, y + d
            .FillStyle = 1
    End Select
    .Forecolor = IIf(c = eDisabled, RGB(198, 197, 201), RGB(28, 81, 128))
    Rectangle hdc, x, y, x + d, y + d
    If iCheck = 1 Then
        Line (x + (d - 7) \ 2, y + (d - 7) \ 2 + 2)-(x + (d - 7) \ 2, y + (d - 7) \ 2 + 5), IIf(c = eDisabled, RGB(198, 197, 201), RGB(33, 161, 33))
        Line (x + (d - 7) \ 2 + 1, y + (d - 7) \ 2 + 3)-(x + (d - 7) \ 2 + 1, y + (d - 7) \ 2 + 6), IIf(c = eDisabled, RGB(198, 197, 201), RGB(33, 161, 33))
        For i = 0 To 4
            Line (x + (d - 7) \ 2 + 2 + i, y + (d - 7) \ 2 + 4 - i)-(x + (d - 7) \ 2 + 2 + i, y + (d - 7) \ 2 + 7 - i), IIf(c = eDisabled, RGB(198, 197, 201), RGB(33, 161, 33))
        Next i
    End If
End With
End Sub

Private Sub UserControl_Click()
    If bPrevButton = 1 Then
        m_Checked = Not (m_Checked)
        Refresh
        RaiseEvent Click
    End If
End Sub

Private Sub UserControl_DblClick()
    If bPrevButton = 1 Then
        UserControl_MouseDown 1, 0, 1, 1
    End If
End Sub

Private Sub UserControl_ExitFocus()
    bFocus = False
    Refresh
End Sub

Private Sub UserControl_GotFocus()
    bFocus = True
    If MouseEvent <> eMouseMovingClicking Then Refresh
End Sub

Private Sub UserControl_Initialize()
    Font.Name = "Arial"
    m_Caption = "Nu1t kie63m"
    m_TiengViet = True
    m_Forecolor = 0
    m_Shadow = True
    m_ShadowColor = vbWhite
    m_BackColor = &H8000000F
    m_Enabled = True
    Refresh
End Sub

Private Sub UserControl_InitProperties()
    'Refresh
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        UserControl_Click
    End If
    If KeyCode = vbKeyRight Then
        SendKeys "{TAB}"
    ElseIf KeyCode = vbKeyLeft Then
        SendKeys "+{TAB}"
    End If
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    If Button = 1 Then
        bFocus = True
        MouseEvent = eMouseMovingClicking
        Refresh
    End If
    bPrevButton = Button
    RaiseEvent MouseDown(Button, Shift, x, y)
    UserControl.Parent.SetFocus
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
With UserControl
    If Button = 0 Then
        FTip.SetTip .hwnd, .Extender.ToolTipText, , , m_TooltipTiengViet, "Verdana"
    End If
    If x < 0 Or y < 0 Or x > .ScaleWidth Or y > .ScaleHeight Then
re:
        If Button = 1 And MouseEvent = eMouseMovingClicking Then
            MouseEvent = eMouseLeavingClicking
            Refresh
        ElseIf MouseEvent <> eMouseLeavingClicking Then
            MouseEvent = eMouseLeaving
            Refresh
        End If
        If Button <> 1 Then
            ReleaseCapture
        End If
        RaiseEvent MouseLeave(Button, Shift, x, y)
    Else
        Dim t2 As POINTAPI
        GetCursorPos t2
        If WindowFromPoint(t2.x, t2.y) <> .hwnd Then
            GoTo re
        Else
            SetCapture hwnd
        End If
        If Button = 1 And MouseEvent = eMouseLeavingClicking Then
            MouseEvent = eMouseMovingClicking
            Refresh
        ElseIf Button = 1 Then
            MouseEvent = eMouseMovingClicking
            Refresh
        Else
            Static t As POINTAPI
            If t2.x = t.x And t2.y = t.y Then
                RaiseEvent MouseMove(Button, Shift, x, y)
                Exit Sub
            End If
            MouseEvent = eMouseMoving
            Refresh
            GetCursorPos t
        End If
        RaiseEvent MouseMove(Button, Shift, x, y)
    End If
End With
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If MouseEvent = eMouseMovingClicking Then
            MouseEvent = eMouseMoving
        Else
            MouseEvent = eMouseLeaving
        End If
        Refresh
    End If
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
    m_TiengViet = .ReadProperty("TiengViet", True)
    m_Caption = .ReadProperty("Caption", "Nu1t kie63m")
    m_Forecolor = .ReadProperty("Forecolor", 0)
    m_BackColor = .ReadProperty("BackColor", Parent.Backcolor)
    m_ShadowColor = .ReadProperty("ShadowColor", vbWhite)
    m_Shadow = .ReadProperty("Shadow", True)
    m_Alignment = .ReadProperty("Alignment", ecbLeft)
    m_Transparent = .ReadProperty("Transparent", False)
    Set UserControl.Font = .ReadProperty("Font", Parent.Font)
    m_Enabled = .ReadProperty("Enabled", True)
    m_Checked = .ReadProperty("Checked", False)
    m_TooltipTiengViet = .ReadProperty("ToolTipTiengViet", True)
    Refresh
End With
End Sub

Private Sub UserControl_Resize()
    Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    Call .WriteProperty("TiengViet", m_TiengViet, True)
    Call .WriteProperty("Caption", m_Caption, "Nu1t kie63m")
    Call .WriteProperty("Forecolor", m_Forecolor, 0)
    Call .WriteProperty("BackColor", m_BackColor, Parent.Backcolor)
    Call .WriteProperty("ShadowColor", m_ShadowColor, vbWhite)
    Call .WriteProperty("Shadow", m_Shadow, True)
    Call .WriteProperty("Alignment", m_Alignment, ecbLeft)
    Call .WriteProperty("Transparent", m_Transparent, False)
    Call .WriteProperty("Font", UserControl.Font, Parent.Font)
    Call .WriteProperty("Enabled", m_Enabled, True)
    Call .WriteProperty("Checked", m_Checked, False)
    Call .WriteProperty("ToolTipTiengViet", m_TooltipTiengViet, True)
    Refresh
End With
End Sub

Public Sub About()
Attribute About.VB_UserMemId = -552
On Error Resume Next
    fAbout.Show 1
End Sub
