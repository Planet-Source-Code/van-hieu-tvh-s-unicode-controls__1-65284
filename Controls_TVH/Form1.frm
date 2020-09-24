VERSION 5.00
Object = "*\AControls_TVH.vbp"
Begin VB.Form Form1 
   BackColor       =   &H0000FF00&
   BorderStyle     =   0  'None
   Caption         =   "XP2005"
   ClientHeight    =   1980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1980
   ScaleWidth      =   3510
   StartUpPosition =   2  'CenterScreen
   Begin Controls_TVH.Frame_TVH Frame_TVH1 
      Height          =   1980
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   3493
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "TVH's Controls"
      FrameStyle      =   2
      AlignmentText   =   2
      ShadowColor     =   16777152
      BorderColor     =   8388608
      BackTitleColor  =   8388608
      ForeTitleColor  =   255
      BackFrameColor  =   8438015
      TitleHeight     =   40
      RoundCorner     =   0   'False
      BackColor       =   0
      BorderGradientWidth=   9
      GradientInColor =   33023
      Begin Controls_TVH.Button_TVH Button_TVH3 
         Height          =   375
         Left            =   285
         TabIndex        =   1
         Top             =   1290
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   661
         Caption         =   "Test 3"
         ShadowColor     =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureSize     =   0
         TransPicture    =   0   'False
         TransColor      =   0
         PictureWidth    =   0
         PictureHeight   =   0
         ToolTipTiengViet=   0   'False
      End
      Begin Controls_TVH.Button_TVH Button_TVH2 
         Height          =   375
         Left            =   1890
         TabIndex        =   2
         Top             =   750
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   661
         Caption         =   "Test 2"
         ShadowColor     =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureSize     =   0
         TransPicture    =   0   'False
         TransColor      =   0
         PictureWidth    =   0
         PictureHeight   =   0
         ToolTipTiengViet=   0   'False
      End
      Begin Controls_TVH.Button_TVH Button_TVH1 
         Height          =   375
         Left            =   285
         TabIndex        =   3
         Top             =   750
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   661
         Caption         =   "Test 1"
         ShadowColor     =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureSize     =   0
         TransPicture    =   0   'False
         TransColor      =   0
         PictureWidth    =   0
         PictureHeight   =   0
         ToolTipTiengViet=   0   'False
      End
      Begin Controls_TVH.Button_TVH Button_TVH4 
         Height          =   375
         Left            =   1890
         TabIndex        =   4
         Top             =   1290
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   661
         Caption         =   "END"
         Forecolor       =   255
         ShadowColor     =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureSize     =   0
         TransPicture    =   0   'False
         TransColor      =   0
         PictureWidth    =   0
         PictureHeight   =   0
         ToolTipTiengViet=   0   'False
      End
      Begin Controls_TVH.Label_TVH Label_TVH1 
         Height          =   1245
         Left            =   135
         Top             =   600
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   2196
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         Transparent     =   0   'False
         Shadow          =   -1  'True
         ShadowStyle     =   0
         BackColorStyle  =   1
         GradientBackColorStyle=   9
         GradientBackColorStart=   12640511
         GradientBackColorEnd=   33023
         ToolTipTiengViet=   0   'False
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Changed As Boolean
Private Sub Button_TVH1_Click()
    fTest.Show
End Sub

Private Sub Button_TVH2_Click()
    fTest2.Show
End Sub

Private Sub Button_TVH3_Click()
    fTest3.Show
End Sub

Private Sub Command1_Click()
    Me.Caption = TT.gotiengviet
End Sub

Private Sub Button_TVH4_Click()
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub



