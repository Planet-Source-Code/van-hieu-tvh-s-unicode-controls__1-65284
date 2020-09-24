VERSION 5.00
Begin VB.UserControl Button_TVH 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1290
   ClipControls    =   0   'False
   DefaultCancel   =   -1  'True
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   24
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   86
   ToolboxBitmap   =   "Button_TVH.ctx":0000
   Begin VB.PictureBox PSrc 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   990
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   0
      Top             =   165
      Visible         =   0   'False
      Width           =   810
   End
End
Attribute VB_Name = "Button_TVH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'///////////////////////////////////////// Truong Van Hieu ////////////////////////////////////////
'////////////////////////////////// tvhhh2003@yahoo.com /////////////////////////////////////
'//////////////////////////////////// Special for Vietnamese /////////////////////////////////////
Option Explicit

Public Enum E_ButtonStyle
    btXP = 0
End Enum

Private Enum E_ButtonEvents
    eNormal = 0
    eGotFocus = 1
    eMoveOver = 2
    eClickDown = 3
    eDisabled = 4
End Enum

Enum E_AlignmentPicture
    eLeft = 0
    eRight = 1
    eCenter = 2
    eTopCenter = 3
    eBottomCenter = 4
End Enum

Enum E_PictureSize
    e_16x16 = 0
    e_32x32 = 1
    e_64x64 = 2
    e_FullStretch = 3
    e_ScaleStrech = 4
    e_Custom = 5
End Enum

Enum E_MouseEvent
    eMouseLeaving = 0
    eMouseLeavingClicking = 1
    eMouseMoving = 2
    eMouseMovingClicking = 3
End Enum

Private bKeySpaceDown As Boolean
Private MouseEvent As E_MouseEvent
Private bFocus As Boolean
Private bPrevButton As Integer
Private imgX As Long
Private imgY As Long
Private imgW As Long
Private imgH As Long
Const SideSpace = 7
Const d_Pic_Caption = 1 'Khoang cach trong tu Picture den Caption
Const min_d_Text = 15 'Khoang cach toi thieu co`n du* de co' the DrawCaption
Const DisabledColor1 = &HC0C0C0
Const DisabledColor2 = &HE0E0E0
Private m_ButtonStyle As E_ButtonStyle
Private m_Enabled As Boolean
Private m_TiengViet As Boolean
Private m_Caption As String
Private m_CaptionTV As String
Private m_Forecolor As OLE_COLOR
Private m_ShadowColor As OLE_COLOR
Private m_Shadow As Boolean
Private m_AlignmentText As E_Alignment
Private m_Font As StdFont
Private m_AlignmentPicture As E_AlignmentPicture
Private m_PictureSize As E_PictureSize
Private m_TransPicture As Boolean
Private m_TransColor As OLE_COLOR
Private m_PictureWidth As Long
Private m_PictureHeight As Long
Private m_TooltipTiengViet As Boolean

'-------------------------------------------------------
Event Click()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseLeave(Button As Integer, Shift As Integer, x As Single, y As Single)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)

'------------------------------------------------------
Public Property Get ButtonStyle() As E_ButtonStyle
    ButtonStyle = m_ButtonStyle
End Property

Public Property Let ButtonStyle(new_ButtonStyle As E_ButtonStyle)
    m_ButtonStyle = new_ButtonStyle
    PropertyChanged "ButtonStyle"
    Fresh
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(new_Enabled As Boolean)
    m_Enabled = new_Enabled
    UserControl.Enabled = m_Enabled
    PropertyChanged "Enabled"
    Fresh
End Property

Public Property Get TiengViet() As Boolean
    TiengViet = m_TiengViet
End Property

Public Property Let TiengViet(new_TiengViet As Boolean)
    m_TiengViet = new_TiengViet
    PropertyChanged "TiengViet"
    Fresh
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(new_Caption As String)
    m_Caption = new_Caption
    m_CaptionTV = mUnicode.VNI_Unicode(new_Caption)
    PropertyChanged "Caption"
    Fresh
End Property

Public Property Get CaptionTV() As String
    CaptionTV = m_CaptionTV
End Property

Public Property Get Forecolor() As OLE_COLOR
    Forecolor = m_Forecolor
End Property

Public Property Let Forecolor(new_Forecolor As OLE_COLOR)
    m_Forecolor = new_Forecolor
    PropertyChanged "Forecolor"
    Fresh
End Property

Public Property Get ShadowColor() As OLE_COLOR
    ShadowColor = m_ShadowColor
End Property

Public Property Let ShadowColor(new_ShadowColor As OLE_COLOR)
    m_ShadowColor = new_ShadowColor
    PropertyChanged "ShadowColor"
    Fresh
End Property

Public Property Get Shadow() As Boolean
    Shadow = m_Shadow
End Property

Public Property Let Shadow(new_Shadow As Boolean)
    m_Shadow = new_Shadow
    PropertyChanged "Shadow"
    Fresh
End Property

Public Property Get AlignmentText() As E_Alignment
    AlignmentText = m_AlignmentText
End Property

Public Property Let AlignmentText(new_AlignmentText As E_Alignment)
    m_AlignmentText = new_AlignmentText
    PropertyChanged "AlignmentText"
    Fresh
End Property

Public Property Get Font() As StdFont
    Set Font = UserControl.Font
End Property

Public Property Set Font(new_Font As StdFont)
    Set UserControl.Font = new_Font
    PropertyChanged "Font"
    Fresh
End Property

Public Property Get Picture() As StdPicture
    Set Picture = PSrc.Picture
End Property

Public Property Set Picture(new_Picture As StdPicture)
    Set PSrc.Picture = new_Picture
    PropertyChanged "Picture"
    Fresh
End Property

Public Property Get AlignmentPicture() As E_AlignmentPicture
    AlignmentPicture = m_AlignmentPicture
End Property

Public Property Let AlignmentPicture(new_AlignmentPicture As E_AlignmentPicture)
    m_AlignmentPicture = new_AlignmentPicture
    PropertyChanged "AlignmentPicture"
    Fresh
End Property

Public Property Get PictureSize() As E_PictureSize
    PictureSize = m_PictureSize
End Property

Public Property Let PictureSize(new_PictureSize As E_PictureSize)
    m_PictureSize = new_PictureSize
    PropertyChanged "PictureSize"
    Fresh
End Property

Public Property Get TransPicture() As Boolean
    TransPicture = m_TransPicture
End Property

Public Property Let TransPicture(new_TransPicture As Boolean)
    m_TransPicture = new_TransPicture
    PropertyChanged "TransPicture"
    Fresh
End Property

Public Property Get TransColor() As OLE_COLOR
    TransColor = m_TransColor
End Property

Public Property Let TransColor(new_TransColor As OLE_COLOR)
    m_TransColor = new_TransColor
    PropertyChanged "TransColor"
    Fresh
End Property

Public Property Get PictureWidth() As Long
    PictureWidth = m_PictureWidth
End Property

Public Property Let PictureWidth(new_PictureWidth As Long)
    m_PictureWidth = new_PictureWidth
    PropertyChanged "PictureWidth"
    Fresh
End Property

Public Property Get PictureHeight() As Long
    PictureHeight = m_PictureHeight
End Property

Public Property Let PictureHeight(new_PictureHeight As Long)
    m_PictureHeight = new_PictureHeight
    PropertyChanged "PictureHeight"
    Fresh
End Property

Public Property Get TooltipTiengViet() As Boolean
    TooltipTiengViet = m_TooltipTiengViet
End Property

Public Property Let TooltipTiengViet(new_TooltipTiengViet As Boolean)
    m_TooltipTiengViet = new_TooltipTiengViet
    PropertyChanged "ToolTipTiengViet"
End Property




Public Sub About()
Attribute About.VB_UserMemId = -552
    On Error Resume Next
    fAbout.Show 1
End Sub

'------------------------------------------------------
Private Sub Fresh() '(Optional BtEvents As E_ButtonEvents = eNormal)
Dim t As E_ButtonEvents
    On Error Resume Next
    Cls
    Select Case MouseEvent
        Case eMouseLeaving:
            t = IIf(bFocus = True, eGotFocus, eNormal)
        Case eMouseLeavingClicking:
            t = eMoveOver
        Case eMouseMoving:
            t = eMoveOver
        Case eMouseMovingClicking:
            t = eClickDown
    End Select
    DrawButton m_ButtonStyle, t
    DrawPicture t
    DrawCaption t
    
End Sub

Private Sub DrawPicture(Optional BtEvents As E_ButtonEvents = eNormal)
With UserControl
    If PSrc.Picture Then
        Select Case m_PictureSize
            Case e_16x16:
                imgW = 16
                imgH = 16
            Case e_32x32:
                imgW = 32
                imgH = 32
            Case e_64x64:
                imgW = 64
                imgH = 64
            Case e_FullStretch:
                imgW = .ScaleWidth ' - 2 * SideSpace
                imgH = .ScaleHeight ' - 2 * SideSpace
            Case e_ScaleStrech:
                SignPic PSrc.ScaleWidth, PSrc.ScaleHeight, .ScaleWidth - 2 * SideSpace, .ScaleHeight - 2 * SideSpace, imgW, imgH
            Case e_Custom:
                imgW = m_PictureWidth
                imgH = m_PictureHeight
        End Select
        Select Case m_AlignmentPicture
            Case eLeft:
                imgX = SideSpace
                imgY = (.ScaleHeight - imgH) \ 2
            Case eRight:
                imgX = .ScaleWidth - SideSpace - imgW - 1
                imgY = (.ScaleHeight - imgH) \ 2
            Case eCenter:
                imgX = (.ScaleWidth - imgW) \ 2
                imgY = (.ScaleHeight - imgH) \ 2
            Case eTopCenter:
                imgX = (.ScaleWidth - imgW) \ 2
                imgY = SideSpace
            Case eBottomCenter:
                imgX = (.ScaleWidth - imgW) \ 2
                imgY = .ScaleHeight - SideSpace - imgH - 1
        End Select
        If BtEvents = eClickDown Then imgX = imgX + 1: imgY = imgY + 1
        If PSrc.Picture.Type = 3 Then
            DrawIconEx hdc, imgX, imgY, PSrc.Picture.Handle, imgW, imgH, 0, 0, 3
        Else
            If m_TransPicture Then
                TransparentBlt .hdc, imgX, imgY, imgW, imgH, PSrc.hdc, 0, 0, PSrc.ScaleWidth, PSrc.ScaleHeight, TransColor
            Else
                TransparentBlt .hdc, imgX, imgY, imgW, imgH, PSrc.hdc, 0, 0, PSrc.ScaleWidth, PSrc.ScaleHeight, -1
            End If
        End If
    End If
End With
End Sub

Sub SignPic(WSrc As Long, HSrc As Long, WDes As Long, HDes As Long, WResult As Long, HResult As Long)
    If (HDes * WSrc) / HSrc <= WDes Then
        HResult = HDes
        WResult = (HDes * WSrc) / HSrc
    Else
        WResult = WDes
        HResult = (WDes * HSrc) / WSrc
    End If
End Sub

Private Sub DrawCaption(Optional BtEvents As E_ButtonEvents = eNormal)
Dim t As RECT
Dim s As String
Dim Flag As Long
With UserControl
    s = IIf(m_TiengViet, m_CaptionTV, m_Caption)
    t.Top = 0
    t.Bottom = .ScaleHeight - 1
    If m_PictureSize = e_FullStretch Or m_AlignmentPicture = eCenter Or m_AlignmentPicture = eTopCenter Or m_AlignmentPicture = eBottomCenter Then
        t.Left = SideSpace
        t.Right = .ScaleWidth - 1 - SideSpace
        If m_AlignmentPicture = eTopCenter Then
            t.Top = imgY + imgH
        ElseIf m_AlignmentPicture = eBottomCenter Then
            t.Bottom = imgY - 1
        End If
    ElseIf m_PictureSize = e_ScaleStrech Then
        If .ScaleWidth - imgW - 2 * SideSpace - d_Pic_Caption >= min_d_Text Then
            If m_AlignmentPicture = eLeft Then
                t.Left = SideSpace + imgW + d_Pic_Caption
                t.Right = .ScaleWidth - 1 - SideSpace
            ElseIf m_AlignmentPicture = eRight Then
                t.Left = SideSpace
                t.Right = .ScaleWidth - 1 - SideSpace - imgW - d_Pic_Caption
            End If
        Else
            t.Left = SideSpace
            t.Right = .ScaleWidth - 1 - SideSpace
        End If
    ElseIf m_PictureSize <> e_FullStretch Then
        If m_AlignmentPicture = eLeft Then
            t.Left = SideSpace + imgW + d_Pic_Caption
            t.Right = .ScaleWidth - 1 - SideSpace
        ElseIf m_AlignmentPicture = eRight Then
            t.Left = SideSpace
            t.Right = .ScaleWidth - 1 - SideSpace - imgW - d_Pic_Caption
        End If
    End If
    Flag = DT_NOCLIP Or DT_WORDBREAK
    Flag = Flag Or IIf(m_AlignmentText = aRight, DT_RIGHT, 0)
    Flag = Flag Or IIf(m_AlignmentText = aCenter, DT_CENTER, 0)
    Dim d As Long
    Set PSrc.Font = UserControl.Font
    d = DrawTextW(PSrc.hdc, StrPtr(s), Len(s), t, Flag)
    t.Top = t.Top + (t.Bottom - t.Top - d) \ 2 + 1
    t.Bottom = .ScaleHeight - 1
    If BtEvents = eClickDown Then Offset t, 1
    Offset t, 1
    .Forecolor = IIf(m_Enabled, m_ShadowColor, DisabledColor2)
    If m_Shadow Then DrawTextW .hdc, StrPtr(s), Len(s), t, Flag
    .Forecolor = IIf(m_Enabled, m_Forecolor, DisabledColor1)
    Offset t, -1
    DrawTextW .hdc, StrPtr(s), Len(s), t, Flag
End With
End Sub

Private Sub DrawButton(BtStyle As E_ButtonStyle, BtEvents As E_ButtonEvents)
Dim c(18) As Long
Dim i As Integer
Dim r() As Long
With UserControl
    Select Case BtStyle
        Case btXP:
            DrawXP2005Button IIf(m_Enabled, BtEvents, eDisabled)
        Case Else:
    End Select
End With
End Sub

Private Sub DrawXP2005Button(BtEvents As E_ButtonEvents)
Dim c(18) As Long
Dim i As Integer
Dim r() As Long
With UserControl
    Select Case BtEvents
        Case eNormal:
            c(1) = RGB(168, 183, 202)
            c(2) = RGB(84, 115, 150)
            c(3) = RGB(78, 109, 147)
            c(4) = RGB(123, 145, 180)
            c(5) = RGB(43, 79, 130)
            c(6) = RGB(213, 222, 236)
            c(7) = RGB(223, 232, 245)
            c(8) = RGB(152, 174, 209)
            c(9) = RGB(144, 168, 205)
            c(10) = RGB(248, 251, 255)
            c(11) = RGB(205, 215, 228)
            c(12) = RGB(243, 248, 255)
            c(13) = RGB(249, 253, 255)
            c(14) = RGB(241, 248, 249)
            c(15) = RGB(208, 220, 235)
            c(16) = RGB(166, 184, 207)
        Case eGotFocus:
            c(2) = RGB(45, 86, 176)
            c(3) = c(2)
            c(5) = c(2)
            c(6) = RGB(139, 192, 234)
            c(7) = RGB(113, 161, 217)
            c(8) = RGB(55, 99, 176)
            c(9) = RGB(63, 109, 177)
            c(10) = RGB(163, 185, 224)
            c(11) = c(9)
            
            c(17) = c(10)
            c(18) = RGB(94, 129, 188)
            
            c(1) = RGB(168, 183, 202)
            c(4) = RGB(123, 145, 180)
            c(12) = RGB(243, 248, 255)
            c(13) = RGB(249, 253, 255)
            c(14) = RGB(241, 248, 249)
            c(15) = RGB(208, 220, 235)
            c(16) = RGB(166, 184, 207)
            
        Case eMoveOver:
            c(6) = RGB(245, 232, 169)
            c(7) = c(6)
            c(8) = RGB(218, 121, 38)
            c(9) = c(8)
            c(10) = c(6)
            c(11) = RGB(230, 170, 69)
            c(17) = RGB(239, 207, 134)
            c(18) = RGB(241, 184, 96)
            
            c(1) = RGB(168, 183, 202)
            c(2) = RGB(84, 115, 150)
            c(3) = RGB(78, 109, 147)
            c(4) = RGB(123, 145, 180)
            c(5) = RGB(43, 79, 130)
            c(12) = RGB(243, 248, 255)
            c(13) = RGB(249, 253, 255)
            c(14) = RGB(241, 248, 249)
            c(15) = RGB(208, 220, 235)
            c(16) = RGB(166, 184, 207)
        Case eClickDown:
            c(6) = RGB(213, 222, 236)
            c(7) = RGB(119, 146, 186)
            c(8) = c(7)
            c(9) = RGB(249, 253, 255)
            c(13) = RGB(136, 158, 188)
            c(14) = RGB(199, 213, 231)
            c(15) = RGB(237, 244, 247)
            c(16) = c(9)
            c(17) = c(7)
            c(18) = RGB(243, 248, 255)
            
            c(1) = RGB(168, 183, 202)
            c(2) = RGB(84, 115, 150)
            c(3) = RGB(78, 109, 147)
            c(4) = RGB(123, 145, 180)
            c(5) = RGB(43, 79, 130)
        Case eDisabled:
            .Backcolor = RGB(255, 255, 255)
            .Forecolor = RGB(198, 197, 201)
            Rectangle hdc, 0, 0, .ScaleWidth, .ScaleHeight
            PSetR 0, 0, RGB(228, 233, 238)
            Exit Sub
    End Select
    PSetR 0, 0, c(1)
    PSetR 1, 0, c(2)
    PSetR 0, 1, c(3)
    PSetR 1, 1, c(4)
    
    Line (2, 0)-(.ScaleWidth - 2, 0), c(5)
    Line (2, .ScaleHeight - 1)-(.ScaleWidth - 2, .ScaleHeight - 1), c(5)
    Line (0, 2)-(0, .ScaleHeight - 2), c(5)
    Line (.ScaleWidth - 1, 2)-(.ScaleWidth - 1, .ScaleHeight - 2), c(5)
    
    Line (2, 1)-(.ScaleWidth - 2, 1), c(6)
    Line (2, .ScaleHeight - 2)-(.ScaleWidth - 2, .ScaleHeight - 2), c(9)
    
    GradientColor2 c(13), c(14), 7, r
    For i = 0 To 6
        Line (2, 2 + i)-(.ScaleWidth - 2, 2 + i), r(i)
    Next i
    GradientColor2 c(15), c(16), 7, r
    For i = 0 To 6
        Line (2, .ScaleHeight - 9 + i)-(.ScaleWidth - 2, .ScaleHeight - 9 + i), r(i)
    Next i
    GradientColor2 c(14), c(15), .ScaleHeight - 16, r
    For i = 0 To .ScaleHeight - 16 - 1
        Line (2, 8 + i)-(.ScaleWidth - 2, 8 + i), r(i)
    Next i
    LineYG 1, 2, .ScaleHeight - 2, c(7), c(8)
    LineYG .ScaleWidth - 2, 2, .ScaleHeight - 2, c(7), c(8)
    If BtEvents <> eClickDown Then
        LineYG 2, 2, .ScaleHeight - 2, c(10), c(11)
        LineYG .ScaleWidth - 3, 2, .ScaleHeight - 2, c(10), c(11)
        Line (3, 3)-(.ScaleWidth - 3, 3), c(12)
    End If
    
    If BtEvents <> eNormal Then
        Line (3, 2)-(.ScaleWidth - 3, 2), c(17)
        Line (3, .ScaleHeight - 3)-(.ScaleWidth - 3, .ScaleHeight - 3), c(18)
    End If
End With
End Sub

Sub PSetR(x As Integer, y As Integer, Color As Long)
With UserControl
    PSet (x, y), Color
    PSet (.ScaleWidth - x - 1, y), Color
    PSet (x, .ScaleHeight - y - 1), Color
    PSet (.ScaleWidth - x - 1, .ScaleHeight - y - 1), Color
End With
End Sub

'Line Gradient
Sub LineXG(x1 As Integer, y As Integer, x2 As Integer, c1 As Long, C2 As Long)
Dim t() As Long
    If x2 - x1 < 1 Then Exit Sub
    GradientColor2 c1, C2, x2 - x1, t
    Dim i As Integer
    For i = 0 To UBound(t)
        PSet (x1 + i, y), t(i)
    Next i
End Sub

Sub LineYG(x As Integer, y1 As Integer, y2 As Integer, c1 As Long, C2 As Long)
Dim t() As Long
    If y2 - y1 < 1 Then Exit Sub
    GradientColor2 c1, C2, y2 - y1, t
    Dim i As Integer
    For i = 0 To UBound(t)
        PSet (x, y1 + i), t(i)
    Next i
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    bPrevButton = vbLeftButton
    UserControl_Click
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    Beep
End Sub

Private Sub UserControl_Click()
    If bPrevButton = vbLeftButton Then
        RaiseEvent Click
    End If
End Sub

Private Sub UserControl_DblClick()
    If bPrevButton = vbLeftButton Then
        UserControl_MouseDown 1, 0, 1, 1
    End If
End Sub

'------------------ UserControl Processing ---------------------------------------------------------------------

Private Sub UserControl_ExitFocus()
    Debug.Print UserControl.Ambient.DisplayName & " - False " & Rnd * 1000
    bFocus = False
    MouseEvent = eMouseLeaving
    Fresh
End Sub

Private Sub UserControl_GotFocus()
    Debug.Print UserControl.Ambient.DisplayName & " - True " & Rnd * 1000
    bFocus = True
    If MouseEvent <> eMouseMovingClicking Then Fresh 'eGotFocus
End Sub

Private Sub UserControl_Initialize()
    m_Enabled = True
    m_TiengViet = True
    m_AlignmentText = aCenter
    m_AlignmentPicture = eLeft
    m_ShadowColor = vbWhite
    m_Shadow = True
    Font.Name = "Arial"
    Caption = "Nu1t ba61m"
    Fresh
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyRight Then
        SendKeys "{TAB}"
    ElseIf KeyCode = vbKeyLeft Then
        SendKeys "+{TAB}"
    ElseIf KeyCode = vbKeySpace Then
        bKeySpaceDown = True
        UserControl_MouseDown vbLeftButton, 0, 0, 0
    End If
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace And bKeySpaceDown Then
        UserControl_MouseUp vbLeftButton, 0, 0, 0
        UserControl_Click
        bKeySpaceDown = False
    End If
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    If Button = 1 Then
        MouseEvent = eMouseMovingClicking
        bFocus = True
        Fresh
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
            Fresh
        ElseIf MouseEvent <> eMouseLeavingClicking Then
            MouseEvent = eMouseLeaving
            Fresh
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
            Fresh
        ElseIf Button = 1 Then
            MouseEvent = eMouseMovingClicking
            Fresh
        Else
            MouseEvent = eMouseMoving
            Fresh
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
        Fresh
    End If
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_Resize()
    Fresh
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
    m_Enabled = .ReadProperty("Enabled", True)
    m_ButtonStyle = .ReadProperty("ButtonStyle", 0)
    m_TiengViet = .ReadProperty("TiengViet", True)
    m_Caption = .ReadProperty("Caption", "Nu1t ba61m")
    m_CaptionTV = mUnicode.VNI_Unicode(m_Caption)
    m_Forecolor = .ReadProperty("Forecolor", 0)
    m_ShadowColor = .ReadProperty("ShadowColor", 0)
    m_Shadow = .ReadProperty("Shadow", True)
    m_AlignmentText = .ReadProperty("AlignmentText", eCenter)
    Set UserControl.Font = .ReadProperty("Font", Parent.Font)
    m_AlignmentPicture = .ReadProperty("AlignmentPicture", eLeft)
    m_PictureSize = .ReadProperty("PictureSize", e_32x32)
    m_TransPicture = .ReadProperty("TransPicture", True)
    m_TransColor = .ReadProperty("TransColor", vbWhite)
    Set PSrc.Picture = .ReadProperty("Picture", Nothing)
    m_PictureWidth = .ReadProperty("PictureWidth", 32)
    m_PictureHeight = .ReadProperty("PictureHeight", 32)
    m_TooltipTiengViet = .ReadProperty("ToolTipTiengViet", True)
    UserControl.Enabled = m_Enabled
    Fresh
End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    Call .WriteProperty("Enabled", m_Enabled, True)
    Call .WriteProperty("ButtonStyle", m_ButtonStyle, 0)
    Call .WriteProperty("TiengViet", m_TiengViet, True)
    Call .WriteProperty("Caption", m_Caption, "Nu1t ba61m")
    Call .WriteProperty("Forecolor", m_Forecolor, 0)
    Call .WriteProperty("ShadowColor", m_ShadowColor, 0)
    Call .WriteProperty("Shadow", m_Shadow, True)
    Call .WriteProperty("AlignmentText", m_AlignmentText, eCenter)
    Call .WriteProperty("Font", UserControl.Font, Parent.Font)
    Call .WriteProperty("AlignmentPicture", m_AlignmentPicture, eLeft)
    Call .WriteProperty("PictureSize", m_PictureSize, e_32x32)
    Call .WriteProperty("TransPicture", m_TransPicture, True)
    Call .WriteProperty("TransColor", m_TransColor, vbWhite)
    Call .WriteProperty("Picture", PSrc.Picture, Nothing)
    Call .WriteProperty("PictureWidth", m_PictureWidth, Nothing)
    Call .WriteProperty("PictureHeight", m_PictureHeight, Nothing)
    Call .WriteProperty("ToolTipTiengViet", m_TooltipTiengViet, Nothing)
End With
End Sub
