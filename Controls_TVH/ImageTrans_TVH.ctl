VERSION 5.00
Begin VB.UserControl ImageTrans_TVH 
   AutoRedraw      =   -1  'True
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   285
   ScaleHeight     =   26
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   19
   Begin VB.PictureBox TPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   120
      Left            =   825
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   7
      TabIndex        =   0
      Top             =   930
      Visible         =   0   'False
      Width           =   105
   End
End
Attribute VB_Name = "ImageTrans_TVH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'///////////////////////////////////////// Truong Van Hieu ////////////////////////////////////////
'////////////////////////////////// tvhhh2003@yahoo.com /////////////////////////////////////
'//////////////////////////////////// Special for Vietnamese /////////////////////////////////////
Option Explicit

Private m_ColorTransparent As OLE_COLOR
Private m_Transparent As Boolean
Private m_Stretch As Boolean

Public Property Get Picture() As StdPicture
    Set Picture = TPic.Picture
End Property

Public Property Set Picture(new_Picture As StdPicture)
    Set TPic.Picture = new_Picture
    PropertyChanged "Picture"
    Refresh
End Property

Public Property Get ColorTransparent() As OLE_COLOR
    ColorTransparent = m_ColorTransparent
End Property

Public Property Let ColorTransparent(new_ColorTransparent As OLE_COLOR)
    m_ColorTransparent = new_ColorTransparent
    PropertyChanged "ColorTransparent"
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

Public Property Get Stretch() As Boolean
    Stretch = m_Stretch
End Property

Public Property Let Stretch(new_Stretch As Boolean)
    m_Stretch = new_Stretch
    PropertyChanged "Stretch"
    Refresh
End Property

Private Sub Refresh()
With UserControl
    .Cls
    If TPic.Picture Is Nothing Then
        Exit Sub
    End If
    Dim w&, h&
    w = IIf(m_Stretch, .ScaleWidth, TPic.ScaleWidth)
    h = IIf(m_Stretch, .ScaleHeight, TPic.ScaleHeight)
    If m_Transparent Then
        UserControl.BackColor = m_ColorTransparent
        TransparentBlt .hdc, 0, 0, w, h, TPic.hdc, 0, 0, TPic.ScaleWidth, TPic.ScaleHeight, m_ColorTransparent
        .MaskColor = m_ColorTransparent
        .MaskPicture = .Image
        .BackStyle = 0
    Else
        .BackStyle = 1
        BitBlt .hdc, 0, 0, w, h, TPic.hdc, 0, 0, vbSrcCopy
    End If
End With
End Sub

Private Sub UserControl_Initialize()
    Refresh
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
    m_ColorTransparent = .ReadProperty("ColorTransparent", 0)
    m_Transparent = .ReadProperty("Transparent", False)
    m_Stretch = .ReadProperty("Stretch", False)
    Set TPic.Picture = .ReadProperty("Picture", Nothing)
    Refresh
End With
End Sub

Private Sub UserControl_Show()
    Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    .WriteProperty "ColorTransparent", m_ColorTransparent, 0
    .WriteProperty "Transparent", m_Transparent, False
    .WriteProperty "Stretch", m_Stretch, False
    .WriteProperty "Picture", TPic.Picture, Nothing
End With
End Sub

Private Sub UserControl_Resize()
    If m_Stretch Then
        Refresh
    Else
        UserControl.Width = TPic.Width * 15
        UserControl.Height = TPic.Height * 15
    End If
End Sub

Public Sub About()
Attribute About.VB_UserMemId = -552
On Error Resume Next
    fAbout.Show 1
End Sub

