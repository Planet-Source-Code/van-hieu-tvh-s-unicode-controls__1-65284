Attribute VB_Name = "modProcess"
'///////////////////////////////////////// Truong Van Hieu ////////////////////////////////////////
'////////////////////////////////// tvhhh2003@yahoo.com /////////////////////////////////////
Option Explicit

Public MenuFontHeight As Long
Public MenuFontName  As String
Public PHdc As Long
Public Hw As Long
Public m_TiengViet As Boolean

Const StrSep As String = "[sep]" '  String SEPARATOR

Const d_Caption_Hotkey = 12

Private Enum E_Alignment
    aLeft = 0
    aCenter = &H1
    aRight = &H2
End Enum

Const FW_DONTCARE = 0
Const FW_THIN = 100
Const FW_EXTRALIGHT = 200
Const FW_LIGHT = 300
Const FW_NORMAL = 400
Const FW_MEDIUM = 500
Const FW_SEMIBOLD = 600
Const FW_BOLD = 700
Const FW_EXTRABOLD = 800
Const FW_HEAVY = 900
Const FW_BLACK = FW_HEAVY
Const FW_DEMIBOLD = FW_SEMIBOLD
Const FW_REGULAR = FW_NORMAL
Const FW_ULTRABOLD = FW_EXTRABOLD
Const FW_ULTRALIGHT = FW_EXTRALIGHT
'used with fdwCharSet
Const ANSI_CHARSET = 0
Const DEFAULT_CHARSET = 1
Const SYMBOL_CHARSET = 2
Const SHIFTJIS_CHARSET = 128
Const HANGEUL_CHARSET = 129
Const CHINESEBIG5_CHARSET = 136
Const OEM_CHARSET = 255
'used with fdwOutputPrecision
Const OUT_CHARACTER_PRECIS = 2
Const OUT_DEFAULT_PRECIS = 0
Const OUT_DEVICE_PRECIS = 5
'used with fdwClipPrecision
Const CLIP_DEFAULT_PRECIS = 0
Const CLIP_CHARACTER_PRECIS = 1
Const CLIP_STROKE_PRECIS = 2
'used with fdwQuality
Const DEFAULT_QUALITY = 0
Const DRAFT_QUALITY = 1
Const PROOF_QUALITY = 2
'used with fdwPitchAndFamily
Const DEFAULT_PITCH = 0
Const FIXED_PITCH = 1
Const VARIABLE_PITCH = 2


Private MsgText As String

Public oldWndProc As Long
Public oldDlgProc As Long
Public hDlgHook As Long
Public sDlgCaption As String
Public sMenuCaption() As String
Public bMaxWidth() As Long
Const MaxItem = 1000

Public Const FORM_CAPTION = "Nga2y mai ra d9i"
Public Const FONT_FILE = ""
Public Const FONT_FACE = "Arial"

Public Const MF_SEPARATOR = &H800&
Public Const MF_DISABLED As Long = &H2&

Public Const GWL_WNDPROC = (-4&)
Public Const WH_CBT = 5
Public Const HCBT_ACTIVATE = 5

Public Const WM_ERASEBKGND As Long = &H14
Public Const WM_NCPAINT = &H85
Public Const WM_NCACTIVATE = &H86
Public Const WM_COMMAND As Long = &H111
Public Const WM_DRAWITEM = &H2B
Public Const WM_MEASUREITEM = &H2C
Public Const WM_SETTEXT = &HC
Public Const WM_SETFONT = &H30
Public Const WM_MENUSELECT = &H11F

Public Const SM_CYMENU = 15
Public Const SM_CYCAPTION = 4&
Public Const SM_CXDLGFRAME = 7
Public Const SM_CYDLGFRAME = 8
Public Const SM_CXSMICON = 49
Public Const SM_CXFRAME = 32
Public Const SM_CYFRAME = 33

Public Const NEWTRANSPARENT = 3

Const DT_SINGLELINE = &H20
Const DT_VCENTER = &H4
Const DT_EXPANDTABS = &H40
Const DT_CENTER = &H1

Public Const MIIM_TYPE = &H10
Public Const MIIM_DATA = &H20
Public Const MF_OWNERDRAW = &H100&
Public Const MIM_BACKGROUND = &H2
Public Const MIM_APPLYTOSUBMENUS = &H80000000

Public Const ODS_CHECKED = &H8
Public Const ODS_DISABLED = &H4
Public Const ODS_FOCUS = &H10
Public Const ODS_GRAYED = &H2
Public Const ODS_SELECTED = &H1
Public Const ODS_DEFAULT = &H20
Public Const ODS_HOTLIGHT = &H40
Public Const ODS_NOACCEL = &H100

Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_MENU = 4
Public Const COLOR_MENUTEXT = 7
Public Const COLOR_CAPTIONTEXT = 9
Public Const COLOR_INACTIVECAPTIONTEXT = 19
Public Const BDR_SUNKENOUTER = &H2
Public Const BF_RECT = &HF

Public Enum E_MenuStyle
    iDefault = 0
    iOffice2003 = 1
End Enum

Public Enum E_MenuStatus
    iLeave = 0
    iMoving = 1
    iHightlight = 2
    iSelected = 3
End Enum

Private Type MENUINFO
  cbSize As Long
  fMask As Long
  dwStyle As Long
  cyMax As Long
  hbrBack As Long
  dwContextHelpID As Long
  dwMenuData As Long
End Type

Type MENUITEMINFO
   cbSize As Long
   fMask As Long
   fType As Long
   fState As Long
   wID As Long
   hSubMenu As Long
   hbmpChecked As Long
   hbmpUnchecked As Long
   dwItemData As Long
   dwTypeData As String
   cch As Long
End Type

Type MEASUREITEMSTRUCT
   CtlType As Long
   CtlID As Long
   itemID As Long
   itemWidth As Long
   itemHeight As Long
   ItemData As Long
End Type

Type DRAWITEMSTRUCT
   CtlType As Long
   CtlID As Long
   itemID As Long
   itemAction As Long
   itemState As Long
   hwndItem As Long
   hdc As Long
   rcItem As RECT
   ItemData As Long
End Type



Declare Function MessageBox Lib "user32" Alias "MessageBoxW" (ByVal hwnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal wType As Long) As Long
Attribute MessageBox.VB_UserMemId = -552
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long

Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuW" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal h As Long, ByVal W As Long, ByVal E As Long, ByVal o As Long, ByVal W As Long, ByVal i As Long, ByVal U As Long, ByVal s As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Declare Function DrawTextEx Lib "user32" Alias "DrawTextExW" (ByVal hdc As Long, ByVal lpsz As Long, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As Any) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal ParenthWnd As Long, ByVal ChildhWnd As Long, ByVal ClassName As String, ByVal Caption As String) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd&, ByVal nIndex&, ByVal dwNewLong&) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Declare Function GetActiveWindow Lib "user32" () As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Declare Function SetMenuInfo Lib "user32.dll" (ByVal hMenu As Long, ByRef LPCMENUINFO As MENUINFO) As Long
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal B As Long, lpMenuItemInfo As MENUITEMINFO) As Long
Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Boolean
Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long


Function NewWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim hdc As Long, uRCT As RECT
Dim uDIS As DRAWITEMSTRUCT, uMIS As MEASUREITEMSTRUCT
Dim hBrush As Long, hPen As Long
Dim hOldBrush As Long, hOldPen As Long
Dim iTextColor As Long, iMenuColor As Long
Dim bSelected As Boolean
Dim mColor&
NewWndProc = CallWindowProc(oldWndProc, hwnd, uMsg, wParam, lParam)
mColor = RGB(235, 233, 237)
Select Case uMsg
'Case WM_ERASEBKGND

Case WM_DRAWITEM
   CopyMemory uDIS, ByVal lParam, LenB(uDIS)
   DrawMenu uDIS, iDefault
Case WM_MEASUREITEM
    CopyMemory uMIS, ByVal lParam, Len(uMIS)
    If uMIS.ItemData Mod MaxItem = 0 Then
        uMIS.itemWidth = TextWidth(sMenuCaption(uMIS.ItemData \ MaxItem, uMIS.ItemData Mod MaxItem))
    Else
        uMIS.itemWidth = bMaxWidth(uMIS.ItemData \ MaxItem) + 33 '16
    End If
    uMIS.itemHeight = GetSystemMetrics(SM_CYMENU)
    CopyMemory ByVal lParam, uMIS, Len(uMIS)
Case WM_MENUSELECT
    SetAllOwnerDraw Hw
End Select

End Function

Sub SetColorPen(hdc As Long, color As Long, hPen As Long)
    hPen = CreatePen(0, 1, color)
    DeleteObject SelectObject(hdc, hPen)
End Sub

Sub SetColorFill(hdc As Long, color As Long, hBrush As Long)
    hBrush = CreateSolidBrush(color)
    DeleteObject SelectObject(hdc, hBrush)
End Sub

Sub RectOffset(R As RECT, d As Long)
    R.Left = R.Left + d
    R.Right = R.Right + d
    R.Bottom = R.Bottom + d
    R.Top = R.Top + d
End Sub

Sub RectFill(hdc As Long, R As RECT, CPen As Long, CBrush As Long)
Dim hPen&, hBrush&
    SetColorPen hdc, CPen, hPen
    SetColorFill hdc, CBrush, hBrush
    Rectangle hdc, R.Left, R.Top, R.Right, R.Bottom
    DeleteObject hPen
    DeleteObject hBrush
End Sub

Sub RectGradientFill(hdc As Long, R As RECT, c1 As Long, C2 As Long)
Dim i%, AColor&()
    GradientColor c1, C2, R.Bottom - R.Top - 1, AColor
    For i = 0 To R.Bottom - R.Top - 1
        LineApi hdc, R.Left, R.Top + i, R.Right, R.Top + i, AColor(i)
    Next i
End Sub

Sub RectGradientFill2(hdc As Long, R As RECT, c1 As Long, C2 As Long)
Dim i%, AColor&(), c&(2)
    c(0) = c1: c(1) = C2: c(2) = c1
    SplitGradientColor c, R.Bottom - R.Top, AColor
    For i = 0 To R.Bottom - R.Top - 1
        LineApi hdc, R.Left, R.Top + i, R.Right, R.Top + i, AColor(i)
    Next i
End Sub


Private Sub DrawMenu(t As DRAWITEMSTRUCT, mStyle As E_MenuStyle)
Dim iTextColor&, iBackColor&, hPen&, hBrush&
Dim cBackMenu&, iAlign As E_Alignment
Dim m_TextColor&, m_TextColorSelect&, m_BackColorSelect&
Dim s As String
    s = sMenuCaption(t.ItemData \ MaxItem, t.ItemData Mod MaxItem)
Select Case mStyle
    Case iDefault:
        cBackMenu = RGB(235, 233, 237)
        SetBkMode t.hdc, 1
        RectFill t.hdc, t.rcItem, cBackMenu, cBackMenu
        If t.ItemData Mod MaxItem = 0 Then
            If t.itemState And ODS_SELECTED Then
                RectGradientFill t.hdc, t.rcItem, RGB(166, 167, 170), vbWhite
                DrawEdge t.hdc, t.rcItem, BDR_SUNKENOUTER, BF_RECT
                RectOffset t.rcItem, 1
            ElseIf t.itemState And ODS_HOTLIGHT Then
                RectGradientFill2 t.hdc, t.rcItem, RGB(166, 167, 170), vbWhite
                DrawEdge t.hdc, t.rcItem, &H4, BF_RECT
            'ElseIf t.itemState And ODS_NOACCEL Then
            End If
            RectOffset t.rcItem, 1
            DrawString t.hdc, t.rcItem, GetCaption(s), False, vbWhite, aCenter
            RectOffset t.rcItem, -1
            iTextColor = vbBlack
            iAlign = aCenter
        Else
            If t.itemState And ODS_DISABLED Then
                If Not isSeparator(s) Then
                    iTextColor = RGB(167, 166, 170)
                    t.rcItem.Left = t.rcItem.Left + 23
                    RectOffset t.rcItem, 1
                    DrawString t.hdc, t.rcItem, GetCaption(s), False, vbWhite, aLeft
                    t.rcItem.Left = t.rcItem.Left - 23
                    RectOffset t.rcItem, -1
                Else
                    iTextColor = vbBlack
                End If
            ElseIf t.itemState And ODS_SELECTED Then
                'RectFill t.hdc, t.rcItem, RGB(51, 94, 168), RGB(51, 94, 168) 'RGB(0, 255, 255), &HC00000
                RectGradientFill t.hdc, t.rcItem, RGB(51, 94, 168), vbWhite
                RectOffset t.rcItem, 1
                t.rcItem.Left = t.rcItem.Left + 23
                DrawString t.hdc, t.rcItem, GetCaption(s), False, vbBlack, aLeft
                t.rcItem.Left = t.rcItem.Left - 23
                RectOffset t.rcItem, -1
                iTextColor = vbWhite
            Else
                iTextColor = vbBlack
            End If
            If t.itemState And ODS_CHECKED Then
                Const d = 13
                Dim x&, y&, i As Byte
                x = t.rcItem.Left + 4
                y = t.rcItem.Top + (t.rcItem.Bottom - t.rcItem.Top - d) \ 2
                LineApi t.hdc, x, y, x + d, y, RGB(167, 166, 170)
                LineApi t.hdc, x, y, x, y + d, RGB(167, 166, 170)
                
                LineApi t.hdc, x + d, y, x + d, y + d + 1, vbWhite
                LineApi t.hdc, x, y + d, x + d + 1, y + d, vbWhite
                
                LineApi t.hdc, x + (d - 7) \ 2, y + (d - 7) \ 2 + 2, x + (d - 7) \ 2, y + (d - 7) \ 2 + 5, vbBlack
                LineApi t.hdc, x + (d - 7) \ 2 + 1, y + (d - 7) \ 2 + 3, x + (d - 7) \ 2 + 1, y + (d - 7) \ 2 + 6, vbBlack
                For i = 0 To 4
                    LineApi t.hdc, x + (d - 7) \ 2 + 2 + i, y + (d - 7) \ 2 + 4 - i, x + (d - 7) \ 2 + 2 + i, y + (d - 7) \ 2 + 7 - i, vbBlack
                Next i
            End If
            t.rcItem.Left = t.rcItem.Left + 23
            iAlign = aLeft
        End If
End Select
    If isSeparator(s) Then
        Dim oldHeight As Integer, oldName As String
        t.rcItem.Left = t.rcItem.Left - 23
        If Mid(s, Len(StrSep) + 1) = "" Then
            LineApi t.hdc, 5, t.rcItem.Top + (t.rcItem.Bottom - t.rcItem.Top) \ 2, bMaxWidth(t.ItemData \ MaxItem) + 40, t.rcItem.Top + (t.rcItem.Bottom - t.rcItem.Top) \ 2, vbBlack 'RGB(167, 166, 170)
            LineApi t.hdc, 5, t.rcItem.Top + (t.rcItem.Bottom - t.rcItem.Top) \ 2 + 1, bMaxWidth(t.ItemData \ MaxItem) + 40, t.rcItem.Top + (t.rcItem.Bottom - t.rcItem.Top) \ 2 + 1, vbWhite 'RGB(235, 233, 237)
        Else
            'Ve Separator co' chu~
            oldHeight = MenuFontHeight
            oldName = MenuFontName
            MenuFontHeight = 11
            MenuFontName = "Tahoma"
            RectOffset t.rcItem, 1
            DrawString t.hdc, t.rcItem, Mid(s, Len(StrSep) + 1), False, vbWhite, aCenter
            RectOffset t.rcItem, -1
            DrawString t.hdc, t.rcItem, Mid(s, Len(StrSep) + 1), False, iTextColor, aCenter
            
            LineApi t.hdc, 5, t.rcItem.Top + (t.rcItem.Bottom - t.rcItem.Top) \ 2, t.rcItem.Left + (t.rcItem.Right - t.rcItem.Left - UniTextWidth(t.hdc, Mid(s, Len(StrSep) + 1))) \ 2 - 2, t.rcItem.Top + (t.rcItem.Bottom - t.rcItem.Top) \ 2, vbBlack
            LineApi t.hdc, t.rcItem.Left + UniTextWidth(t.hdc, Mid(s, Len(StrSep) + 1)) \ 2 + (t.rcItem.Right - t.rcItem.Left) \ 2 + 2, t.rcItem.Top + (t.rcItem.Bottom - t.rcItem.Top) \ 2, t.rcItem.Right - 5, t.rcItem.Top + (t.rcItem.Bottom - t.rcItem.Top) \ 2, vbBlack
            LineApi t.hdc, 5, t.rcItem.Top + (t.rcItem.Bottom - t.rcItem.Top) \ 2 + 1, t.rcItem.Left + (t.rcItem.Right - t.rcItem.Left - UniTextWidth(t.hdc, Mid(s, Len(StrSep) + 1))) \ 2 - 2, t.rcItem.Top + (t.rcItem.Bottom - t.rcItem.Top) \ 2 + 1, vbWhite
            LineApi t.hdc, t.rcItem.Left + UniTextWidth(t.hdc, Mid(s, Len(StrSep) + 1)) \ 2 + (t.rcItem.Right - t.rcItem.Left) \ 2 + 2, t.rcItem.Top + (t.rcItem.Bottom - t.rcItem.Top) \ 2 + 1, t.rcItem.Right - 5, t.rcItem.Top + (t.rcItem.Bottom - t.rcItem.Top) \ 2 + 1, vbWhite
            MenuFontHeight = oldHeight
            MenuFontName = oldName
            SetFontToHdc t.hdc, MenuFontHeight, MenuFontName
        End If
    Else
        DrawString t.hdc, t.rcItem, GetCaption(s), False, iTextColor, iAlign
        t.rcItem.Right = t.rcItem.Left + bMaxWidth(t.ItemData \ MaxItem)
        DrawString t.hdc, t.rcItem, GetHotkey(s), False, iTextColor, aRight
    End If
    If t.ItemData Mod MaxItem = 0 Then
        'DrawMenuBar Hw
    End If
End Sub

Function isSeparator(s As String) As Boolean
    isSeparator = (LCase(Mid(s, 1, Len(StrSep))) = StrSep)
End Function

Sub LineApi(hdc&, x1&, y1&, x2&, y2&, c&)
    Dim t&, tp As POINTAPI, old&
    t = CreatePen(0, 1, c)
    old = SelectObject(hdc, t)
    MoveToEx hdc, x1, y1, tp
    LineTo hdc, x2, y2
    SelectObject hdc, old
    DeleteObject t
End Sub

Function GetCaption(s As String) As String
    If InStr(1, s, vbTab) = 0 Then
        GetCaption = s
    Else
        GetCaption = Mid(s, 1, InStr(1, s, vbTab) - 1)
    End If
End Function

Function GetHotkey(s As String) As String
    If InStrRev(s, vbTab) = 0 Then
        GetHotkey = ""
    Else
        GetHotkey = Mid(s, InStrRev(s, vbTab) + 1)
    End If
End Function

Function GetWidth(s As String) As String
    If GetHotkey(s) = "" Then
        GetWidth = TextWidth(GetCaption(s))
    Else
        GetWidth = TextWidth(GetCaption(s)) + d_Caption_Hotkey + TextWidth(GetHotkey(s))
    End If
End Function

Sub SelectFont(hdc As Long)
Dim hFont&, hOldFont&
    hFont = CreateFont(MenuFontHeight, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, MenuFontName)
    hOldFont = SelectObject(hdc, hFont)
    DeleteObject hFont
End Sub

Sub SetAllOwnerDraw(hwnd As Long)
Dim h&, i&, j&, m() As Long, tm() As Long
    h = GetMenu(hwnd)
    ReDim m(GetMenuItemCount(h))
    m(0) = h
    ReDim tm(UBound(m))
    tm(0) = h
    For i = 0 To GetMenuItemCount(h) - 1
        m(i + 1) = GetSubMenu(h, i)
        tm(i + 1) = m(i + 1)
    Next i
    For i = 1 To UBound(tm)
        GetAllSubMenu tm(i), m
    Next i
    Dim Max As Long
    Max = 0
    For i = 0 To UBound(m)
        If GetMenuItemCount(m(i)) > Max Then
            Max = GetMenuItemCount(m(i))
        End If
    Next i
    ReDim Preserve sMenuCaption(UBound(m), Max)
    ReDim Preserve bMaxWidth(UBound(m))
    For i = 0 To GetMenuItemCount(tm(0)) - 1
        SetOwnerDraw tm(0), 0, i, True
    Next i
    For i = 1 To UBound(m)
        For j = 0 To GetMenuItemCount(m(i)) - 1
            If i > UBound(tm) Then
                SetOwnerDraw m(i), j + 1, i - 1
            Else
                SetOwnerDraw m(i), j + 1, i - 1
            End If
        Next j
    Next i
End Sub

Function IsArrayEmpty(arr() As Long) As Boolean
    On Error GoTo er
    Dim t&
    t = UBound(arr)
    IsArrayEmpty = False
    Exit Function
er:
    IsArrayEmpty = True
End Function

Sub GetAllSubMenu(hMenu As Long, ByRef m() As Long)
    Dim n As Long, i As Long, hChild As Long
    n = GetMenuItemCount(hMenu)
    If IsArrayEmpty(m) Then
        ReDim m(0)
        m(0) = hMenu
    End If
    For i = 0 To n - 1
        hChild = GetSubMenu(hMenu, i)
        If hChild <> 0 Then
            Dim t As Long
            t = UBound(m) + 1
            ReDim Preserve m(t)
            m(t) = hChild
            GetAllSubMenu hChild, m
        End If
    Next i
End Sub

Sub SetFontToHdc(hdc&, Height&, FName As String)
Dim hFont As Long
SetBkMode hdc, NEWTRANSPARENT
hFont = CreateFont(Height, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, FName)
SelectObject hdc, hFont
DeleteObject hFont
End Sub

Private Sub DrawString(hdc As Long, uRCT As RECT, sText As String, bBold As Boolean, iColor As Long, Optional Align As E_Alignment = aCenter)
Dim hOldColor&
hOldColor = SetTextColor(hdc, iColor)
SetFontToHdc hdc, MenuFontHeight, MenuFontName
DrawTextEx hdc, StrPtr(sText), Len(sText), uRCT, DT_SINGLELINE Or DT_VCENTER Or DT_EXPANDTABS Or Align, ByVal 0&
SetTextColor hdc, hOldColor
End Sub

Sub SetOwnerDraw(hMenu As Long, iMenuID As Long, nSub As Long, Optional isSub As Boolean = False)
Dim uMII As MENUITEMINFO
On Error Resume Next
uMII.cbSize = LenB(uMII)
uMII.fMask = MIIM_TYPE Or MIIM_DATA
uMII.dwTypeData = String(256, 0)
uMII.cch = Len(uMII.dwTypeData)
GetMenuItemInfo hMenu, IIf(isSub, nSub, iMenuID - 1), True, uMII
If uMII.dwTypeData = "" And uMII.fType <> MF_SEPARATOR Then
    Exit Sub
End If
If uMII.fType = MF_SEPARATOR Then
    sMenuCaption(nSub, iMenuID) = StrSep
Else
    If isSeparator(Left(uMII.dwTypeData, InStr(uMII.dwTypeData, vbNullChar) - 1)) Then
        uMII.fType = uMII.fType Or MF_SEPARATOR
    End If
    sMenuCaption(nSub, iMenuID) = IIf(m_TiengViet, mUnicode.VNI_Unicode(Left(uMII.dwTypeData, InStr(uMII.dwTypeData, vbNullChar) - 1)), Left(uMII.dwTypeData, InStr(uMII.dwTypeData, vbNullChar) - 1))
End If
uMII.fType = uMII.fType Or MF_OWNERDRAW
uMII.dwItemData = nSub * MaxItem + iMenuID
If isSub = False Then
    Dim s As String
    s = sMenuCaption(nSub, iMenuID)
    If GetWidth(s) > bMaxWidth(nSub) Then
        bMaxWidth(nSub) = GetWidth(s)
    End If
End If
SetMenuItemInfo hMenu, IIf(isSub, nSub, iMenuID - 1), True, uMII
End Sub

Function TextWidth(s As String) As Long
    TextWidth = UniTextWidth(PHdc, s)
End Function

Sub SetupMenu(h As Long)
    DrawMenuSystem h
    oldWndProc = SetWindowLong(h, GWL_WNDPROC, AddressOf NewWndProc)
    SetAllOwnerDraw h
    
End Sub

Sub DrawMenuSystem(h As Long)
Const MF_STRING = &H0&
    Dim t1 As Long, t2 As Long
    t1 = GetSystemMenu(h, 0)
    
    t2 = GetMenuItemID(t1, 0)
    ModifyMenu t1, t2, MF_STRING, t2, StrPtr(mUnicode.VNI_Unicode("Kho6i phu5c"))
    
    t2 = GetMenuItemID(t1, 1)
    ModifyMenu t1, t2, MF_STRING, t2, StrPtr(mUnicode.VNI_Unicode("Di chuye63n"))
    
    t2 = GetMenuItemID(t1, 2)
    ModifyMenu t1, t2, MF_STRING, t2, StrPtr(mUnicode.VNI_Unicode("Ki1ch co74"))
    
    t2 = GetMenuItemID(t1, 3)
    ModifyMenu t1, t2, MF_STRING, t2, StrPtr(mUnicode.VNI_Unicode("Thu nho3"))
    
    t2 = GetMenuItemID(t1, 4)
    ModifyMenu t1, t2, MF_STRING, t2, StrPtr(mUnicode.VNI_Unicode("Pho1ng to"))
    
    t2 = GetMenuItemID(t1, 6)
    ModifyMenu t1, t2, MF_STRING, t2, StrPtr(mUnicode.VNI_Unicode("Ke61t thu1c"))
End Sub

Sub RefreshMenu(h As Long)
    SetAllOwnerDraw h
    DrawMenuBar h
End Sub

Sub Restore(h As Long)
    SetWindowLong h, GWL_WNDPROC, oldWndProc
End Sub

