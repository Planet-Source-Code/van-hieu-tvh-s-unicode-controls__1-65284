VERSION 5.00
Begin VB.UserControl OptionBox_TVH 
   AutoRedraw      =   -1  'True
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1695
   ScaleHeight     =   22
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   113
   ToolboxBitmap   =   "OptionBox.ctx":0000
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Index           =   5
      Left            =   1560
      Picture         =   "OptionBox.ctx":0312
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Index           =   4
      Left            =   1440
      Picture         =   "OptionBox.ctx":03A4
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   1200
      Picture         =   "OptionBox.ctx":0436
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   960
      Picture         =   "OptionBox.ctx":0680
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   720
      Picture         =   "OptionBox.ctx":08CA
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   480
      Picture         =   "OptionBox.ctx":0B14
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   195
   End
End
Attribute VB_Name = "OptionBox_TVH"
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
    eMoveOver = 1
    eClickDown = 2
    eDisabled = 3
End Enum

Enum E_AlignmentOpt
    eobLeft = 0
    eobRight = 1
    eobTop = 2
    eobBottom = 3
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
Private m_Alignment As E_AlignmentOpt
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

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(new_BackColor As OLE_COLOR)
    m_BackColor = new_BackColor
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

Public Property Get Alignment() As E_AlignmentOpt
    Alignment = m_Alignment
End Property

Public Property Let Alignment(new_Alignment As E_AlignmentOpt)
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

Public Property Get Hwnd() As Long
    Hwnd = UserControl.Hwnd
End Property


'-------------------------------------------------------------------------------------------------------------------------------------

Sub Refresh()
Dim st As E_CheckStatus
With UserControl
    .Cls
    .BackColor = IIf(m_Transparent, TransColor, m_BackColor)
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
    DrawOptionBoxXp2005 IIf(m_Enabled, st, eDisabled)
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
    Set Pic(0).Font = UserControl.Font
    s = IIf(m_TiengViet, mUnicode.VNI_Unicode(m_Caption), m_Caption)
    Dim dong&
    dong = DrawTextW(Pic(0).hdc, StrPtr(s), Len(s), t, Flag Or DT_NOCLIP Or DT_WORDBREAK)
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
            't.Right = t.Left + TextWidthW(.hdc, s) + 2
        Else
            't.Right = t.Right + 1
        End If
        .Forecolor = 0
        DrawFocusRect hdc, t
    End If
End With
End Sub

Private Sub DrawOptionBoxXp2005(c As E_CheckStatus, Optional iCheck As Byte = 0)
Dim y&, x&, id As Byte
With UserControl
    If iCheck <> 1 And iCheck <> 2 Then
        iCheck = IIf(m_Checked, 1, 2)
    End If
    Select Case m_Alignment
        Case eobLeft:
            x = 0
            y = (.ScaleHeight - d) \ 2
        Case eobRight:
            x = .ScaleWidth - d
            y = (.ScaleHeight - d) \ 2
        Case eobTop:
            x = (.ScaleWidth - d) \ 2
            y = 0
        Case eobBottom:
            x = (.ScaleWidth - d) \ 2
            y = .ScaleHeight - d
    End Select
    Select Case c
        Case eClickDown:
            id = 2
        Case eDisabled:
            id = 3
        Case eMoveOver
            id = 1
        Case eNormal:
            id = 0
    End Select
    TransparentBlt .hdc, x, y, Pic(0).ScaleWidth, Pic(0).ScaleWidth, _
                        Pic(id).hdc, 0, 0, Pic(0).ScaleWidth, Pic(0).ScaleWidth, vbRed
    If iCheck = 1 Then
        TransparentBlt .hdc, x + Pic(0).ScaleWidth \ 2 - Pic(5).ScaleWidth \ 2, y + Pic(0).ScaleWidth \ 2 - Pic(5).ScaleHeight \ 2, Pic(5).ScaleWidth, Pic(5).ScaleHeight, _
                        Pic(IIf(c = eDisabled, 5, 4)).hdc, 0, 0, Pic(5).ScaleWidth, Pic(5).ScaleHeight, vbRed
    End If
End With
End Sub

Sub ClearChecks()
Dim o As Control
    
    For Each o In UserControl.Parent.Controls
        If (TypeOf o Is OptionBox_TVH) Then
            If o.Container.Hwnd = UserControl.ContainerHwnd Then
                If o.Hwnd <> UserControl.Hwnd Then
                    If (o.Checked) Then o.Checked = False
                End If
            End If
        End If
    Next
End Sub

Private Sub UserControl_Click()
    If bPrevButton = 1 Then
        ClearChecks
        m_Checked = True
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
    m_Caption = "Nu1t cho5n"
    m_TiengViet = True
    m_Forecolor = 0
    m_Shadow = True
    m_ShadowColor = vbWhite
    m_BackColor = &H8000000F
    m_Enabled = True
    Pic(0).ScaleWidth = Pic(0).ScaleWidth
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
        FTip.SetTip .Hwnd, .Extender.ToolTipText, , , m_TooltipTiengViet, "Verdana"
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
        If WindowFromPoint(t2.x, t2.y) <> .Hwnd Then
            GoTo re
        Else
            SetCapture Hwnd
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
    m_BackColor = .ReadProperty("BackColor", Parent.BackColor)
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
    Call .WriteProperty("BackColor", m_BackColor, Parent.BackColor)
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

