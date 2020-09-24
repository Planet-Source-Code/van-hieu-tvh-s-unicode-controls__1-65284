VERSION 5.00
Begin VB.UserControl Menu_TVH 
   ClientHeight    =   840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   840
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   InvisibleAtRuntime=   -1  'True
   Picture         =   "Menu_TVH.ctx":0000
   ScaleHeight     =   56
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   56
   ToolboxBitmap   =   "Menu_TVH.ctx":2502
End
Attribute VB_Name = "Menu_TVH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'///////////////////////////////////////// Truong Van Hieu ////////////////////////////////////////
'////////////////////////////////// tvhhh2003@yahoo.com /////////////////////////////////////
'//////////////////////////////////// Special for Vietnamese /////////////////////////////////////
Option Explicit

Public Property Get TiengViet() As Boolean
    TiengViet = m_TiengViet
End Property

Public Property Let TiengViet(new_TiengViet As Boolean)
    m_TiengViet = new_TiengViet
    PropertyChanged "TiengViet"
End Property

Public Property Get FontName() As StdFont
    UserControl.FontName = MenuFontName
    Set FontName = UserControl.Font
    PropertyChanged "FontName"
End Property

Public Property Set FontName(new_FontName As StdFont)
    Set UserControl.Font = new_FontName
    MenuFontName = UserControl.FontName
End Property

Public Property Get FontHeight() As Long
    FontHeight = MenuFontHeight
    PropertyChanged "FontHeight"
End Property

Public Property Let FontHeight(new_FontHeight As Long)
    MenuFontHeight = new_FontHeight
End Property


Public Sub Setup(h As Long)
    SelectFont UserControl.hdc
    PHdc = UserControl.hdc
    Hw = h
    SetupMenu h
End Sub

Public Sub Refresh()
    RefreshMenu Hw
End Sub

Private Sub UserControl_Initialize()
    UserControl_Resize
    MenuFontHeight = 16
    MenuFontName = "Tahoma"
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
    m_TiengViet = .ReadProperty("TiengViet", True)
    MenuFontName = .ReadProperty("FontName.FontName", "Tahoma")
    MenuFontHeight = .ReadProperty("FontHeight", 16)
End With
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = 56 * 15
    UserControl.Height = 56 * 15
End Sub

Private Sub UserControl_Terminate()
    Restore Hw
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    Call .WriteProperty("TiengViet", m_TiengViet, True)
    Call .WriteProperty("FontName.FontName", MenuFontName, "Tahoma")
    Call .WriteProperty("FontHeight", MenuFontHeight, 16)
End With
End Sub


Public Sub About()
Attribute About.VB_UserMemId = -552
On Error Resume Next
    fAbout.Show 1
End Sub

'Public Sub About()
'    MessageBox Parent.hwnd, StrPtr(VNI_Unicode("Chu7o7ng tri2nh d9u7o75c vie16t bo73i Tru7o7ng Va8n Hie61u." & vbCrLf & "Email Contact: tvhhh2003@yahoo.com")), StrPtr("About TVH Menu"), vbOKOnly
'End Sub
