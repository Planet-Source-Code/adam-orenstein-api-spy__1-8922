VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuDoStuff 
      Caption         =   "Do stuff"
      Begin VB.Menu mnuWndFind 
         Caption         =   "Find &Window"
      End
      Begin VB.Menu mnuSnapShot 
         Caption         =   "Sa&ve Snapshot"
      End
      Begin VB.Menu mnuEnableDisable 
         Caption         =   "Enable/ Disable"
         Begin VB.Menu mnuEnable 
            Caption         =   "Enable"
         End
         Begin VB.Menu mnuDisable 
            Caption         =   "Disable"
         End
      End
      Begin VB.Menu mnuSet 
         Caption         =   "Set"
         Begin VB.Menu mnuOnTopVals 
            Caption         =   "On top values"
            Begin VB.Menu mnuWinOnTop 
               Caption         =   "Window on top"
            End
            Begin VB.Menu mnuWinNotOnTop 
               Caption         =   "Window not on top"
            End
         End
         Begin VB.Menu mnuZOrderPos 
            Caption         =   "Z-Order position"
            Begin VB.Menu mnuZTop 
               Caption         =   "Top of Z-order"
            End
            Begin VB.Menu mnuZbottom 
               Caption         =   "Bottom of Z-order"
            End
         End
         Begin VB.Menu mnuChangeCText 
            Caption         =   "Change control text"
         End
         Begin VB.Menu mnuChangeText 
            Caption         =   "Change window text"
         End
      End
      Begin VB.Menu mnuSendMessage 
         Caption         =   "Send Message to window"
         Begin VB.Menu mnuCreate 
            Caption         =   "Create"
         End
         Begin VB.Menu mnuDestroy 
            Caption         =   "Destroy"
         End
         Begin VB.Menu mnuClose 
            Caption         =   "Close"
         End
         Begin VB.Menu mnuRefresh 
            Caption         =   "Refresh"
         End
         Begin VB.Menu mnuClick 
            Caption         =   "Click"
            Begin VB.Menu mnuLeftClick 
               Caption         =   "Left click"
            End
            Begin VB.Menu mnuLeftDblClick 
               Caption         =   "Left click (double)"
            End
            Begin VB.Menu mnuRightClick 
               Caption         =   "Right click"
            End
            Begin VB.Menu mnuRightDblClick 
               Caption         =   "Right click (double)"
            End
         End
      End
      Begin VB.Menu mnuShowWindow 
         Caption         =   "Show window"
         Begin VB.Menu mnuHideWin 
            Caption         =   "Hide window"
         End
         Begin VB.Menu mnuShowWin 
            Caption         =   "Show window"
         End
         Begin VB.Menu mnuMinimize 
            Caption         =   "Minimize"
         End
         Begin VB.Menu mnuMaximize 
            Caption         =   "Maximize"
         End
         Begin VB.Menu mnuRestore 
            Caption         =   "Restore"
         End
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RetVal As Long

Private Sub mnuChangeCText_Click()
    Dim Input_ As String
    
    NotOnTop Form1
    Input_ = InputBox("Change control text to:", "Change window text", MainText)
    OnTop Form1
    Call SetWindowText(Win, Input_)
    Call SendMessage(Win, WM_SETTEXT, ByVal CLng(0), ByVal Input_)
End Sub

Private Sub mnuChangeText_Click()
    Dim Input_ As String
    
    NotOnTop Form1
    Input_ = InputBox("Change text to:", "Change window text", MainText)
    OnTop Form1
    Call SetWindowText(Win, Input_)
    Call SendMessage(Win, WM_PAINT, 0&, 0&)
End Sub

Private Sub mnuClose_Click()
    Call SendMessage(Win, WM_CLOSE, 0, 0)
End Sub

Private Sub mnuCreate_Click()
    Call SendMessage(Win, WM_CREATE, 0&, 0&)
End Sub

Private Sub mnuDestroy_Click()
    Call SendMessage(Win, WM_DESTROY, 0&, 0&)
End Sub

Private Sub mnuDisable_Click()
    Call EnableWindow(Win, EW_DISABLE)
End Sub

Private Sub mnuEnable_Click()
    Call EnableWindow(Win, EW_Enable)
End Sub

Private Sub mnuHideWin_Click()
    Call ShowWindow(Win, SW_HIDE)
End Sub

Private Sub mnuLeftClick_Click()
    Call PostMessage(Win, WM_LBUTTONDOWN, ByVal CLng(0), ByVal CLng(0))
    Call PostMessage(Win, WM_LBUTTONUP, ByVal CLng(0), ByVal CLng(0))
End Sub

Private Sub mnuLeftDblClick_Click()
   Call PostMessage(Win, WM_LBUTTONDBLCLK, ByVal CLng(0), ByVal CLng(0))
End Sub

Private Sub mnuMaximize_Click()
    Call ShowWindow(Win, SW_MAXIMIZE)
End Sub

Private Sub mnuMinimize_Click()
    Call ShowWindow(Win, SW_MINIMIZE)
End Sub

Private Sub mnuRefresh_Click()
    Call SendMessage(Win, WM_PAINT, 0, 0)
End Sub

Private Sub mnuRestore_Click()
    Call ShowWindow(Win, SW_RESTORE)
End Sub

Private Sub mnuRightClick_Click()
    Call PostMessage(Win, WM_RBUTTONDOWN, ByVal CLng(0), ByVal CLng(0))
    Call PostMessage(Win, WM_RBUTTONUP, ByVal CLng(0), ByVal CLng(0))
End Sub

Private Sub mnuRightDblClick_Click()
    Call PostMessage(Win, WM_RBUTTONDBLCLK, ByVal CLng(0), ByVal CLng(0))
End Sub

Private Sub mnuShowWin_Click()
    Call ShowWindow(Win, SW_SHOW)
End Sub

Private Sub mnuSnapShot_Click()
    Dim SavePic As String
            
    Call GetWindowShot("ThunderRT6FormDC", Form1.Picture2)
    'IMPORTANT!!!!-> 1st Param would be "ThunderRT5Form" for VB 5.  The snapshot will not work properly when the program is run in design-time.  Only in EXE form becuase the form changes its name from design-time to when it is an EXE.
          
    NotOnTop Form1
    SavePic = InputBox("Save Snapshot As:", "Save Snapshot", App.Path & "\Snapshot.BMP")
        
    
    If Not InStr(1, SavePic$, ".Bmp", vbTextCompare) Then
        SavePic = SavePic & ".Bmp"
    End If
        
    OnTop Form1
    Call SavePicture(Form1.Picture2.Image, SavePic)
    
End Sub


Private Sub mnuWinNotOnTop_Click()
    Call SetWindowPos(Win, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_FLAGS)
End Sub

Private Sub mnuWinOnTop_Click()
    Call SetWindowPos(Win, HWND_TOPMOST, 0, 0, 0, 0, SWP_FLAGS)
End Sub

Private Sub mnuWndFind_Click()
    Seen = 1
    Unload Form1
    frmWindowFind.Show
End Sub

Private Sub mnuZbottom_Click()
  Call SetWindowPos(Win, HWND_BOTTOM, 0, 0, 0, 0, SWP_FLAGS)
End Sub

Private Sub mnuZTop_Click()
    Call SetWindowPos(Win, HWND_TOP, 0, 0, 0, 0, SWP_FLAGS)
End Sub
