VERSION 5.00
Begin VB.Form FTip 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2475
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   26
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmWaitStart 
      Interval        =   400
      Left            =   900
      Top             =   90
   End
   Begin VB.Timer tmProcess 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1395
      Top             =   30
   End
   Begin VB.Timer tmWaitEnd 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   2010
      Top             =   0
   End
End
Attribute VB_Name = "FTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////// Truong Van Hieu ////////////////////////////////////////
'////////////////////////////////// tvhhh2003@yahoo.com /////////////////////////////////////
'//////////////////////////////////// Special for Vietnamese /////////////////////////////////////

Option Explicit
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Const DT_TOP = &H0
Const DT_LEFT = &H0
Const DT_CENTER = &H1
Const DT_RIGHT = &H2
Const DT_VCENTER = &H4
Const DT_BOTTOM = &H8
Const DT_WORDBREAK = &H10
Const DT_NOCLIP = &H100

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Const LWA_COLORKEY = &H1
Const LWA_ALPHA = &H2
Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000


Const LB_ITEMFROMPOINT = &H1A9

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function DrawTextW Lib "user32" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40

Private Const SW_SHOWNOACTIVATE = 4


Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32W" (ByVal hdc As Long, ByVal lpsz As Long, ByVal cbString As Long, lpSize As Size) As Long

Private Type Size
    cx As Long
    cy As Long
End Type


Private HwndCur As Long
Private STip As String
Private TV As Boolean
Private WaitEnd As Byte
Private WaitStart As Byte
Private EndPos As POINTAPI
Private MaxWidth As Integer
Private m_TextColor As OLE_COLOR
Private m_BackColor As OLE_COLOR
Private m_TransparentLen As Byte

Const WaitEndTime = 10
Const WaitStartTime = 1

Const D_Edge = 3

Public Property Let TransparentLen(n_TL As Byte)
    m_TransparentLen = n_TL
End Property

Public Property Get TransparentLen() As Byte
    TransparentLen = m_TransparentLen
End Property


Public Property Let TiengViet(n_TV As Boolean)
    TV = n_TV
End Property

Public Property Get TiengViet() As Boolean
    TiengViet = TV
End Property

Public Property Let TipBackColor(n_BC As OLE_COLOR)
    m_BackColor = n_BC
End Property

Public Property Get TipBackColor() As OLE_COLOR
    TipBackColor = m_BackColor
End Property

Public Property Let TipTextColor(n_TC As OLE_COLOR)
    m_TextColor = n_TC
End Property

Public Property Get TipTextColor() As OLE_COLOR
    TipTextColor = m_TextColor
End Property


Sub HideTip()
    tmProcess.Enabled = False
    tmWaitEnd.Enabled = False
    Hide
End Sub

Sub ShowTip()
    Dim r As RECT
    Dim T As POINTAPI
    Dim d&
    Dim ts As String
    ts = IIf(TV, mUnicode.VNI_Unicode(STip), STip)
    GetCursorPos T
    r.Top = D_Edge
    r.Left = D_Edge
    r.Bottom = ScaleHeight
    r.Right = r.Left + TextWidthW(hdc, ts)
    Cls
    If TextWidthW(hdc, ts) > MaxWidth Then
        r.Right = MaxWidth + D_Edge * 2
    End If
    ForeColor = m_TextColor
    BackColor = m_BackColor
    d = DrawTextW(hdc, StrPtr(ts), Len(ts), r, DT_NOCLIP Or DT_WORDBREAK Or DT_CENTER)
    Call ShowWindow(Me.hWnd, SW_SHOWNOACTIVATE)
    Dim X&, Y&, w&, h&
    X = T.X + 1
    Y = T.Y + 22
    w = r.Right + D_Edge
    h = d + 2 * D_Edge
    If X + w > Screen.Width / 15 Then X = Screen.Width / 15 - w
    SetWindowPos Me.hWnd, HWND_TOPMOST, X, Y, w, h, SWP_NOACTIVATE Or SWP_SHOWWINDOW
    Line (0, 0)-(ScaleWidth, 0), RGB(220, 223, 228)
    Line (0, 0)-(0, ScaleHeight), RGB(220, 223, 228)
    Line (0, ScaleHeight - 1)-(ScaleWidth, ScaleHeight - 1), 0
    Line (ScaleWidth - 1, 0)-(ScaleWidth - 1, ScaleHeight), 0
    SetLayeredWindowAttributes Me.hWnd, 0, m_TransparentLen, LWA_ALPHA
End Sub

Sub SetTip(tHwnd As Long, s As String, Optional TTextColor As OLE_COLOR = 0, Optional TBackColor As OLE_COLOR = -1, Optional TiengViet As Boolean = False, Optional tFont, Optional tTipWidthMax As Integer = 300)
Dim tp As POINTAPI
    If s = "" Then Exit Sub
    MaxWidth = tTipWidthMax
    If Not IsMissing(tFont) Then
        On Error Resume Next
        Set Font = StandFont
        Set Font = tFont
        Font.Name = tFont
    Else
        Set Font = StandFont
    End If
    GetCursorPos tp
    If tp.X = EndPos.X And tp.Y = EndPos.Y And tHwnd = HwndCur And s = STip Then
        Exit Sub
    End If
    If tHwnd = HwndCur And s = STip And Visible Then
        Exit Sub
    End If
    GetCursorPos EndPos
    m_TextColor = TTextColor
    m_BackColor = IIf(TBackColor = -1, RGB(255, 255, 225), TBackColor)
    TV = TiengViet
    HwndCur = tHwnd
    STip = s
    tmProcess.Enabled = True
    If Visible Then
        WaitStart = WaitStartTime
        tmWaitStart_Timer
    Else
        WaitStart = 0
        tmWaitStart.Enabled = True
    End If
    SetLayeredWindowAttributes Me.hWnd, 0, m_TransparentLen, LWA_ALPHA
End Sub

Sub SetTipObject(Obj As Object, Optional TTextColor As OLE_COLOR = 0, Optional TBackColor As OLE_COLOR = -1, Optional TiengViet As Boolean = False, Optional tFont, Optional tTipWidthMax As Integer = 300)
Dim tp As POINTAPI
Dim s As String
Dim tHwnd&
    tHwnd = Obj.hWnd
    s = Obj.Tag
    If s = "" Then Exit Sub
    MaxWidth = tTipWidthMax
    If Not IsMissing(tFont) Then
        On Error Resume Next
        Set Font = StandFont
        Set Font = tFont
        Font.Name = tFont
    Else
        Set Font = StandFont
    End If
    GetCursorPos tp
    If tp.X = EndPos.X And tp.Y = EndPos.Y And tHwnd = HwndCur And s = STip Then
        Exit Sub
    End If
    If tHwnd = HwndCur And s = STip And Visible Then
        Exit Sub
    End If
    GetCursorPos EndPos
    m_TextColor = TTextColor
    m_BackColor = IIf(TBackColor = -1, RGB(255, 255, 225), TBackColor)
    TV = TiengViet
    HwndCur = tHwnd
    STip = s
    tmProcess.Enabled = True
    If Visible Then
        WaitStart = WaitStartTime
        tmWaitStart_Timer
    Else
        WaitStart = 0
        tmWaitStart.Enabled = True
    End If
End Sub

Function StandFont() As StdFont
Dim T As New StdFont
    T.Bold = False
    T.Italic = False
    T.Name = "Arial"
    T.Size = 10
    T.Strikethrough = False
    T.Underline = False
    Set StandFont = T
End Function

Private Sub Form_Load()
    m_BackColor = RGB(255, 255, 225)
    m_TextColor = 0
    MaxWidth = 500
    TransparentLen = 255
    Dim Ret As Long
    Ret = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong hWnd, GWL_EXSTYLE, Ret
    m_TransparentLen = 255
End Sub

Private Sub tmProcess_Timer()
Dim T As POINTAPI
Dim h&
    GetCursorPos T
    h = WindowFromPoint(T.X, T.Y)
    If h <> HwndCur Then
        tmWaitStart.Enabled = False
        HideTip
    End If
End Sub

Private Sub tmWaitEnd_Timer()
    If WaitEnd = WaitEndTime Then
        'GetCursorPos EndPos
        HideTip
    End If
    WaitEnd = WaitEnd + 1
End Sub

Private Sub tmWaitStart_Timer()
    If WaitStart = WaitStartTime Then
        WaitEnd = 0
        ShowTip
        tmWaitEnd.Enabled = True
        tmWaitStart.Enabled = False
    End If
    
    WaitStart = WaitStart + 1
End Sub

Sub SetListTooltip(Mylst As ListBox, Button As Integer, X As Single, Y As Single, Optional lMaxWidth As Integer = 200)
Dim lXPoint As Long
Dim lYPoint As Long
Dim lIndex As Long
If Button = 0 Then
    lXPoint = CLng(X / Screen.TwipsPerPixelX)
    lYPoint = CLng(Y / Screen.TwipsPerPixelY)
    With Mylst
        lIndex = SendMessage(.hWnd, _
        LB_ITEMFROMPOINT, 0, ByVal _
        ((lYPoint * 65536) + lXPoint))
        If (lIndex >= 0) And _
        (lIndex <= .ListCount) Then
            .Tag = .List(lIndex)
        Else
            .Tag = ""
        End If
    End With
    FTip.SetTip Mylst.hWnd, Mylst.Tag, , , True, , lMaxWidth
End If
End Sub

Function TextWidthW&(hdc&, s As String)
Dim sz As Size
    GetTextExtentPoint32 hdc, StrPtr(s), Len(s), sz
    TextWidthW = sz.cx
End Function



