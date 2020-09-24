VERSION 5.00
Begin VB.Form fAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   111
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   344
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Controls_TVH.Button_TVH bOK 
      Height          =   375
      Left            =   1830
      TabIndex        =   0
      Top             =   1185
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   661
      Caption         =   "D9o62ng y1"
      ShadowColor     =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureWidth    =   32
      PictureHeight   =   32
      ToolTipTiengViet=   -1  'True
   End
   Begin Controls_TVH.Label_TVH Label_TVH1 
      Height          =   525
      Index           =   0
      Left            =   0
      Top             =   30
      Width           =   5115
      _ExtentX        =   6271
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Control ActiveX TVH"
      WordWrap        =   -1  'True
      BorderStyle     =   2
      OutlineColor    =   65535
      Shadow          =   -1  'True
      ShadowDepth     =   3
      ShadowStyle     =   0
      ShadowColorStart=   65535
      ShadowColorEnd  =   0
      Alignment       =   2
      GradientBackColorStyle=   1
      GradientBackColorStart=   12648447
      GradientBackColorEnd=   16776960
   End
   Begin Controls_TVH.Label_TVH Label_TVH1 
      Height          =   315
      Index           =   1
      Left            =   810
      Top             =   510
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "by Tru7o7ng Va8n Hie61u 2006"
      WordWrap        =   -1  'True
      OutlineColor    =   65535
      Shadow          =   -1  'True
      ShadowDepth     =   2
      ShadowStyle     =   0
      ShadowColorEnd  =   16776960
      Alignment       =   2
      GradientBackColorStyle=   1
      GradientBackColorStart=   12648447
      GradientBackColorEnd=   16776960
   End
   Begin Controls_TVH.Label_TVH Label_TVH1 
      Height          =   315
      Index           =   2
      Left            =   840
      Top             =   795
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Email: tvhhh2003@yahoo.com"
      WordWrap        =   -1  'True
      ForeColor       =   4210752
      OutlineColor    =   65535
      Shadow          =   -1  'True
      ShadowStyle     =   0
      ShadowColorStart=   0
      Alignment       =   2
      GradientBackColorStyle=   1
      GradientBackColorStart=   12648447
      GradientBackColorEnd=   16776960
   End
   Begin VB.Shape Shape1 
      Height          =   1665
      Left            =   0
      Top             =   0
      Width           =   5160
   End
End
Attribute VB_Name = "fAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bOK_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Then
        bOK_Click
    End If
End Sub

Private Sub Form_Load()
Dim c(2) As Long
Dim t() As Long
    DoEvents
    c(0) = vbBlack
    c(1) = vbWhite
    c(2) = vbBlack
    SplitGradientColor c, ScaleHeight, t
Dim i As Integer
    For i = 0 To ScaleHeight - 1
        Line (0, i)-(ScaleWidth, i), t(i)
    Next i
End Sub

