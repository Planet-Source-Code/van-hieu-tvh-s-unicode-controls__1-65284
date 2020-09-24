Attribute VB_Name = "modProcess"
Option Explicit

Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
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

Const MF_SEPARATOR = &H800&
Const MF_STRING = &H0&

Const GWL_WNDPROC = (-4&)

Const MIIM_TYPE = &H10
Const MIIM_DATA = &H20
Const MF_OWNERDRAW = &H100&

Const NEWTRANSPARENT = 3

Const DT_SINGLELINE = &H20
Const DT_VCENTER = &H4
Const DT_EXPANDTABS = &H40

Const WM_NCPAINT = &H85
Const WM_NCACTIVATE = &H86
Const WM_DRAWITEM = &H2B
Const WM_MEASUREITEM = &H2C
Const WM_SETTEXT = &HC
Const WM_SETFONT = &H30


Const SM_CYMENU = 15
Const SM_CYCAPTION = 4&
Const SM_CXDLGFRAME = 7
Const SM_CYDLGFRAME = 8
Const SM_CXSMICON = 49
Const SM_CXFRAME = 32
Const SM_CYFRAME = 33

Const ODS_SELECTED = &H1

Const COLOR_HIGHLIGHT = 13
Const COLOR_HIGHLIGHTTEXT = 14
Const COLOR_MENU = 4
Const COLOR_MENUTEXT = 7
Const COLOR_CAPTIONTEXT = 9
Const COLOR_INACTIVECAPTIONTEXT = 19

Const BDR_SUNKENOUTER = &H2
Const BF_RECT = &HF

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd&, ByVal nIndex&, ByVal dwNewLong&) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal h As Long, ByVal w As Long, ByVal E As Long, ByVal O As Long, ByVal w As Long, ByVal i As Long, ByVal U As Long, ByVal s As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExW" (ByVal hdc As Long, ByVal lpsz As Long, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Long, lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Boolean
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long

Private old_Wnd As Long
Private mhWnd As Long
Private sMenuCaption() As String

Public Sub KhoiTao(hwnd As Long)
    Exit Sub
    old_Wnd = SetWindowLong(mhWnd, GWL_WNDPROC, AddressOf NewWndProc)
End Sub

Function NewWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim hdc As Long, uRCT As RECT
Dim uDIS As DRAWITEMSTRUCT, uMIS As MEASUREITEMSTRUCT
Dim hBrush As Long, hPen As Long
Dim hOldBrush As Long, hOldPen As Long
Dim iTextColor As Long, iMenuColor As Long
Dim bSelected As Boolean
NewWndProc = CallWindowProc(old_Wnd, hwnd, uMsg, wParam, lParam)
Select Case uMsg
'Case WM_NCACTIVATE, WM_NCPAINT
'   hdc = GetWindowDC(hWnd)
'   uRCT.Left = GetSystemMetrics(SM_CXFRAME) + GetSystemMetrics(SM_CXSMICON) + 4
'   uRCT.Top = GetSystemMetrics(SM_CYFRAME)
'   uRCT.Right = uRCT.Left + 200
'   uRCT.Bottom = uRCT.Top + GetSystemMetrics(SM_CYCAPTION) - 1
'   If uMsg = WM_NCACTIVATE Then
'      If wParam Then iTextColor = COLOR_CAPTIONTEXT Else iTextColor = COLOR_INACTIVECAPTIONTEXT
'   Else
'      If hWnd = GetActiveWindow() Then iTextColor = COLOR_CAPTIONTEXT Else iTextColor = COLOR_INACTIVECAPTIONTEXT
'   End If
'   DrawString hdc, uRCT, FORM_CAPTION, True, iTextColor
'   ReleaseDC hWnd, hdc
Case WM_DRAWITEM
   CopyMemory uDIS, ByVal lParam, LenB(uDIS)
   If (uDIS.itemState And ODS_SELECTED) Then
      iMenuColor = COLOR_HIGHLIGHT
      iTextColor = COLOR_HIGHLIGHTTEXT
      bSelected = True
   Else
      iMenuColor = COLOR_MENU
      iTextColor = COLOR_MENUTEXT
   End If
   hBrush = CreateSolidBrush(GetSysColor(iMenuColor))
   hPen = CreatePen(0, 1, GetSysColor(iMenuColor))
   hOldBrush = SelectObject(uDIS.hdc, hBrush)
   hOldPen = SelectObject(uDIS.hdc, hPen)
   If uDIS.ItemData = 0 Then
      iTextColor = COLOR_MENUTEXT
      If bSelected Then
         DrawEdge uDIS.hdc, uDIS.rcItem, BDR_SUNKENOUTER, BF_RECT
      Else
         Rectangle uDIS.hdc, uDIS.rcItem.Left, uDIS.rcItem.Top, uDIS.rcItem.Right, uDIS.rcItem.Bottom
      End If
      uDIS.rcItem.Left = uDIS.rcItem.Left + 8
   Else
      Rectangle uDIS.hdc, uDIS.rcItem.Left, uDIS.rcItem.Top, uDIS.rcItem.Right, uDIS.rcItem.Bottom
      uDIS.rcItem.Left = uDIS.rcItem.Left + GetSystemMetrics(SM_CYMENU)
   End If
   DrawString uDIS.hdc, uDIS.rcItem, sMenuCaption(uDIS.ItemData), False, iTextColor
   SelectObject uDIS.hdc, hOldBrush
   SelectObject uDIS.hdc, hOldPen
   DeleteObject hBrush
   DeleteObject hPen
Case WM_MEASUREITEM
   CopyMemory uMIS, ByVal lParam, Len(uMIS)
   If uMIS.ItemData = 0 Then uMIS.itemWidth = 35 Else uMIS.itemWidth = 100
   uMIS.itemHeight = GetSystemMetrics(SM_CYMENU)
   CopyMemory ByVal lParam, uMIS, Len(uMIS)
End Select
End Function


Private Sub DrawString(hdc As Long, uRCT As RECT, sText As String, bBold As Boolean, iColor As Long)
Dim hFont As Long, hOldFont As Long, hOldColor As Long
sText = mUnicode.Text_To_Unicode(sText)
SetBkMode hdc, NEWTRANSPARENT
hOldColor = SetTextColor(hdc, GetSysColor(iColor))
hFont = CreateFont(16, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, FONT_FACE)
hOldFont = SelectObject(hdc, hFont)
DrawTextEx hdc, StrPtr(sText), Len(sText), uRCT, DT_SINGLELINE Or DT_VCENTER Or DT_EXPANDTABS, ByVal 0&
SetTextColor hdc, hOldColor
SelectObject hdc, hOldFont
DeleteObject hFont
End Sub

Sub SetOwnerDraw(hMenu As Long, iMenuID As Long)
Dim uMII As MENUITEMINFO
    uMII.cbSize = LenB(uMII)
    uMII.fMask = MIIM_TYPE Or MIIM_DATA
    uMII.dwTypeData = String(256, 0)
    uMII.cch = Len(uMII.dwTypeData)
    GetMenuItemInfo hMenu, iMenuID, True, uMII
    If uMII.fType <> MF_SEPARATOR Then
        ReDim Preserve sMenuCaption(UBound(sMenuCaption) + 1)
        sMenuCaption(UBound(sMenuCaption)) = Left(uMII.dwTypeData, InStr(uMII.dwTypeData, vbNullChar) - 1)
    End If
    uMII.fType = uMII.fType Or MF_OWNERDRAW
    uMII.dwItemData = UBound(sMenuCaption) + 1
    SetMenuItemInfo hMenu, iMenuID, True, uMII
End Sub

Sub SetAllOwnerDraw(hwnd As Long)
Dim h&, i&, j&, n&, m() As Long
    ReDim sMenuCaption(0)
    h = GetMenu(hwnd)
    GetAllSubMenu h, m
    n = 1
    For i = 0 To UBound(m)
        For j = 0 To GetMenuItemCount(m(i)) - 1
            SetOwnerDraw m(i), j
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

Sub Setup(hwnd As Long)
    mhWnd = hwnd
    SetAllOwnerDraw mhWnd
    old_Wnd = SetWindowLong(mhWnd, GWL_WNDPROC, AddressOf NewWndProc)
    
End Sub

Sub Restore()
    SetWindowLong mhWnd, GWL_WNDPROC, old_Wnd
End Sub

