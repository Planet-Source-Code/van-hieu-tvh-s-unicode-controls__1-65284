VERSION 5.00
Object = "*\AControls_TVH.vbp"
Begin VB.Form fTest2 
   Caption         =   "Button, Check Box, Option Box with Winxp205 Style"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   ScaleHeight     =   368
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   586
   StartUpPosition =   3  'Windows Default
   Begin Controls_TVH.ImageTrans_TVH ImageTrans_TVH1 
      Height          =   765
      Left            =   270
      TabIndex        =   19
      Top             =   120
      Width           =   930
      _ExtentX        =   1535
      _ExtentY        =   1244
      Transparent     =   -1  'True
      Picture         =   "fTest2.frx":0000
   End
   Begin Controls_TVH.Button_TVH Button_TVH1 
      Height          =   360
      Left            =   7080
      TabIndex        =   18
      Top             =   4785
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   635
      Caption         =   "Thoa1t (Exit)"
      ShadowColor     =   16777215
      AlignmentText   =   1
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
      Picture         =   "fTest2.frx":0B89
      PictureWidth    =   0
      PictureHeight   =   0
      ToolTipTiengViet=   0   'False
   End
   Begin Controls_TVH.Frame_TVH Frame_TVH3 
      Height          =   4650
      Left            =   4995
      TabIndex        =   12
      Top             =   75
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   8202
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Frame with Messenger Style"
      FrameStyle      =   5
      AlignmentText   =   2
      ShadowColor     =   65535
      ThemeColor      =   5
      BorderColor     =   1989493
      BackTitleColor  =   1989493
      ForeTitleColor  =   255
      BackFrameColor  =   8443628
      BackColor       =   12648447
      Begin Controls_TVH.Label_TVH Label_TVH5 
         Height          =   285
         Left            =   255
         Top             =   1095
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "Picture Box"
         AutoSize        =   -1  'True
         Shadow          =   -1  'True
         ShadowStyle     =   0
         GradientBackColorEnd=   0
         ToolTipTiengViet=   0   'False
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   870
         Left            =   270
         Picture         =   "fTest2.frx":15CFB
         ScaleHeight     =   810
         ScaleWidth      =   900
         TabIndex        =   20
         Tag             =   "Tooltip Transparent"
         Top             =   1395
         Width           =   960
      End
      Begin VB.CommandButton Command1 
         Caption         =   "VB's Button - Move here"
         Height          =   435
         Left            =   210
         TabIndex        =   17
         Tag             =   "Use Tag Property to Show Tooltip"
         Top             =   4065
         Width           =   3135
      End
      Begin VB.ListBox L 
         Height          =   1230
         ItemData        =   "fTest2.frx":166AE
         Left            =   1725
         List            =   "fTest2.frx":166D0
         TabIndex        =   13
         Top             =   1395
         Width           =   1695
      End
      Begin Controls_TVH.Label_TVH Label_TVH2 
         Height          =   360
         Left            =   -30
         Top             =   720
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "Move on All Onjects"
         AutoSize        =   -1  'True
         WordWrap        =   -1  'True
         Shadow          =   -1  'True
         ShadowStyle     =   1
         Alignment       =   2
         GradientBackColorEnd=   0
         ToolTipTiengViet=   0   'False
      End
      Begin Controls_TVH.Label_TVH Label_TVH4 
         Height          =   285
         Left            =   1710
         Top             =   1050
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "List Box:"
         AutoSize        =   -1  'True
         Shadow          =   -1  'True
         ShadowStyle     =   0
         GradientBackColorEnd=   0
         ToolTipTiengViet=   0   'False
      End
      Begin Controls_TVH.CheckBox_TVH CheckBox_TVH1 
         Height          =   360
         Left            =   195
         TabIndex        =   14
         ToolTipText     =   "Nu1t kie63m kie63u XP 2005"
         Top             =   2820
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   635
         Caption         =   "Check box - Xp2005 Style"
         BackColor       =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Checked         =   -1  'True
      End
      Begin Controls_TVH.CheckBox_TVH C 
         Height          =   360
         Left            =   180
         TabIndex        =   15
         Top             =   3555
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   635
         Caption         =   "Disabled"
         BackColor       =   14737632
         Transparent     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Checked         =   -1  'True
      End
      Begin Controls_TVH.CheckBox_TVH CheckBox_TVH2 
         Height          =   360
         Left            =   195
         TabIndex        =   16
         ToolTipText     =   "Chec Box trong suo61t (Transparent)"
         Top             =   3225
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   635
         Caption         =   "Trong suo61t (Transparent)"
         BackColor       =   14737632
         Transparent     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Checked         =   -1  'True
      End
      Begin Controls_TVH.Label_TVH Label_TVH6 
         Height          =   360
         Left            =   0
         Top             =   450
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "___.- Tooltip by Form -.___"
         AutoSize        =   -1  'True
         WordWrap        =   -1  'True
         ForeColor       =   255
         Shadow          =   -1  'True
         ShadowStyle     =   0
         ShadowColorStart=   0
         Alignment       =   2
         GradientBackColorEnd=   0
         ToolTipTiengViet=   0   'False
      End
   End
   Begin Controls_TVH.Frame_TVH Frame_TVH2 
      Height          =   2460
      Left            =   90
      TabIndex        =   5
      Top             =   90
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   4339
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "    Frame with Gradient Style"
      FrameStyle      =   2
      AlignmentText   =   2
      ShadowColor     =   16761024
      BorderColor     =   8388608
      BackTitleColor  =   8388608
      BackFrameColor  =   12648447
      TitleHeight     =   40
      RoundCorner     =   0   'False
      BackColor       =   0
      GradientOutColor=   8421631
      GradientInColor =   12648447
      Begin Controls_TVH.OptionBox_TVH OptionBox_TVH6 
         Height          =   270
         Left            =   1440
         TabIndex        =   11
         Top             =   1605
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   476
         Caption         =   "Trong suo61t (Transparent)"
         Transparent     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ToolTipTiengViet=   0   'False
      End
      Begin Controls_TVH.OptionBox_TVH OptionBox_TVH3 
         Height          =   255
         Left            =   1455
         TabIndex        =   6
         Top             =   1935
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         Caption         =   "Disabled"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         ToolTipTiengViet=   0   'False
      End
      Begin Controls_TVH.OptionBox_TVH OptionBox_TVH1 
         Height          =   270
         Left            =   255
         TabIndex        =   7
         Top             =   765
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   476
         Caption         =   "Be6n tra1i (Left side)"
         BackColor       =   12648384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ToolTipTiengViet=   0   'False
      End
      Begin Controls_TVH.OptionBox_TVH OptionBox_TVH2 
         Height          =   270
         Left            =   2490
         TabIndex        =   8
         Top             =   765
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   476
         Caption         =   "Be6n pha3i (Right side)"
         BackColor       =   12648384
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ToolTipTiengViet=   0   'False
      End
      Begin Controls_TVH.OptionBox_TVH OptionBox_TVH4 
         Height          =   465
         Left            =   255
         TabIndex        =   9
         Top             =   1065
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   820
         Caption         =   "O73 du7o71i (Bottom)"
         BackColor       =   12648384
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ToolTipTiengViet=   0   'False
      End
      Begin Controls_TVH.OptionBox_TVH OptionBox_TVH5 
         Height          =   465
         Left            =   2490
         TabIndex        =   10
         Top             =   1065
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   820
         Caption         =   "Be6n tre6n (Top)"
         BackColor       =   12648384
         Alignment       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ToolTipTiengViet=   0   'False
      End
   End
   Begin Controls_TVH.Frame_TVH Frame_TVH1 
      Height          =   2100
      Left            =   105
      TabIndex        =   0
      Top             =   2625
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   3704
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Frame with JC_Style"
      FrameStyle      =   3
      AlignmentText   =   2
      ShadowColor     =   16777152
      BorderColor     =   8388608
      BackTitleColor  =   8388608
      ForeTitleColor  =   16711680
      BackFrameColor  =   15783104
      BackColor       =   0
      Begin Controls_TVH.Button_TVH Button_TVH4 
         Default         =   -1  'True
         Height          =   570
         Left            =   2700
         TabIndex        =   1
         ToolTipText     =   "Tooltip"
         Top             =   495
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1005
         Caption         =   "Button"
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
         TransPicture    =   0   'False
         TransColor      =   0
         Picture         =   "fTest2.frx":1679D
         PictureWidth    =   32
         PictureHeight   =   32
         ToolTipTiengViet=   -1  'True
      End
      Begin Controls_TVH.Button_TVH Button_TVH3 
         Cancel          =   -1  'True
         Height          =   390
         Left            =   2700
         TabIndex        =   2
         ToolTipText     =   "Kho6ng co1 hi2nh"
         Top             =   1125
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   688
         Caption         =   "Empy"
         Shadow          =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureSize     =   0
         TransPicture    =   0   'False
         TransColor      =   0
         PictureWidth    =   32
         PictureHeight   =   32
         ToolTipTiengViet=   -1  'True
      End
      Begin Controls_TVH.Button_TVH B1 
         Height          =   1470
         Left            =   90
         TabIndex        =   3
         ToolTipText     =   "Tooltip cho Button na2y ba82ng Tie61ng Vie65t"
         Top             =   480
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   2593
         Caption         =   "Move here"
         Forecolor       =   8454016
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlignmentPicture=   2
         PictureSize     =   4
         TransColor      =   16580348
         Picture         =   "fTest2.frx":16AB7
         PictureWidth    =   32
         PictureHeight   =   32
         ToolTipTiengViet=   -1  'True
      End
      Begin Controls_TVH.Button_TVH Button_TVH2 
         Height          =   390
         Left            =   2685
         TabIndex        =   4
         ToolTipText     =   "D9a6u ma61t tie6u rui2?"
         Top             =   1560
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   688
         Enabled         =   0   'False
         Caption         =   "Disabled"
         ShadowColor     =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureSize     =   0
         TransPicture    =   0   'False
         TransColor      =   16580348
         Picture         =   "fTest2.frx":1AEB8
         PictureWidth    =   32
         PictureHeight   =   32
         ToolTipTiengViet=   -1  'True
      End
   End
   Begin Controls_TVH.Label_TVH Label_TVH3 
      Height          =   540
      Left            =   5385
      Top             =   2250
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   953
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "<-- Tooltip Trong suo61t, move vo6 d9a6y"
      WordWrap        =   -1  'True
      BackColor       =   12648447
      Transparent     =   0   'False
      Shadow          =   -1  'True
      ShadowStyle     =   0
      Alignment       =   2
      GradientBackColorEnd=   0
      ToolTipTiengViet=   0   'False
   End
   Begin Controls_TVH.Label_TVH Label_TVH1 
      Height          =   330
      Left            =   15
      ToolTipText     =   "D9i5a chi3 Mail: tvhhh2003@yahoo.com"
      Top             =   5190
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "TVH's Unicode Controls -  Copyright© Tru7o7ng Va8n Hie61u 2006"
      BorderSize      =   2
      BorderStyle     =   4
      Transparent     =   0   'False
      Shadow          =   -1  'True
      ShadowStyle     =   0
      Alignment       =   2
      BackColorStyle  =   1
      GradientBackColorStart=   12648447
      GradientBackColorEnd=   16777152
   End
End
Attribute VB_Name = "fTest2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Button_TVH1_Click()
    Unload Me
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FTip.TransparentLen = 255
    'FTip.SetTip Command1.hWnd, "Du2ng Tooltip cho ca1c Control kha1c", vbWhite, vbBlack, True, "Tahoma"
    FTip.SetTipObject Command1, vbWhite, vbBlack, True, "Tahoma"
End Sub



Private Sub L_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FTip.SetListTooltip L, Button, X, Y
End Sub


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim T As New StdFont
    T.Bold = True
    T.Italic = True
    T.Name = "Arial"
    T.Size = 25
    FTip.TransparentLen = 128
    FTip.SetTipObject Picture1, , , True, T
    
End Sub
