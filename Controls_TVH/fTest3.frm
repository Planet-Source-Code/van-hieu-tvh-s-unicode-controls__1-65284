VERSION 5.00
Object = "*\AControls_TVH.vbp"
Begin VB.Form fTest3 
   Caption         =   "Test Menu_TVH"
   ClientHeight    =   2535
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4485
   LinkTopic       =   "Form2"
   ScaleHeight     =   2535
   ScaleWidth      =   4485
   StartUpPosition =   3  'Windows Default
   Begin Controls_TVH.Menu_TVH M 
      Left            =   2385
      Top             =   630
      _ExtentX        =   1482
      _ExtentY        =   1482
   End
   Begin VB.Menu mFile 
      Caption         =   "Ta65p tin"
      Begin VB.Menu mOpen 
         Caption         =   "Mo73 (Open)"
         Shortcut        =   ^O
      End
      Begin VB.Menu mSave 
         Caption         =   "Lu7u la5i... (Save as..)"
         Shortcut        =   ^R
      End
      Begin VB.Menu mStep1 
         Caption         =   "-"
      End
      Begin VB.Menu mRecent 
         Caption         =   "Ca1c ta65p tin ga62n d9a62y (History)"
         Begin VB.Menu mStep4 
            Caption         =   "[Sep]Ca1c ta65p tin d9a4 d9u7o75c mo73"
         End
         Begin VB.Menu mFile1 
            Caption         =   "Ta65p tin 1"
            Checked         =   -1  'True
            Shortcut        =   +^{F9}
         End
         Begin VB.Menu mFile2 
            Caption         =   "Ta65p tin 2"
         End
         Begin VB.Menu mFile3 
            Caption         =   "Ta65p tin 3"
         End
      End
      Begin VB.Menu mStep2 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "Thoa1t (Exit)"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mEdit 
      Caption         =   "Soa5n tha3o (Editor)"
      Begin VB.Menu mCut 
         Caption         =   "Ca81t (Cut)"
         Shortcut        =   ^C
      End
      Begin VB.Menu mPaste 
         Caption         =   "Da1n (Paste)"
         Shortcut        =   ^V
      End
      Begin VB.Menu mStep3 
         Caption         =   "[Sep]Co6ng cu5 (Tools)"
      End
      Begin VB.Menu mThuoc 
         Caption         =   "Thu7o71c d9o (Ruler)"
         Checked         =   -1  'True
         Shortcut        =   +^{F12}
      End
   End
   Begin VB.Menu mAbout 
      Caption         =   "Tho6ng tin (About)"
   End
End
Attribute VB_Name = "fTest3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Private Sub Form_Load()
    M.Setup hWnd
End Sub

Private Sub mAbout_Click()
    M.About
End Sub

Private Sub mExit_Click()
    Unload Me
End Sub
