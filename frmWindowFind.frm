VERSION 5.00
Begin VB.Form frmWindowFind 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "§pIdEr's API Spy"
   ClientHeight    =   4200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   2790
      Left            =   3000
      TabIndex        =   1
      Top             =   720
      Width           =   2655
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   2790
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "CHILDREN"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "PARENTS"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5160
      TabIndex        =   7
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5400
      TabIndex        =   6
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "§pIdEr's API Spy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5775
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      X1              =   0
      X2              =   6000
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "REFRESH"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BACK TO MAIN"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   3840
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FIND CHILDREN"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   3480
      Width           =   2655
   End
End
Attribute VB_Name = "frmWindowFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim RetVal As Long
    
    OnTop frmWindowFind
        
    EnumWindows AddressOf EnumWindowsProc, 0

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     
     Label1.ForeColor = RGB(0, 255, 0)
     Label2.ForeColor = RGB(0, 255, 0)
     Label3.ForeColor = RGB(0, 255, 0)

End Sub

Private Sub Form_Resize()
    
    OnTop frmWindowFind

End Sub

Private Sub Label1_Click()
    
    List2.Clear
        
    EnumChildWindows Thing(List1.ListIndex + 1), AddressOf EnumChildProc, 0
    
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Label1.ForeColor = &H4000&
    Label2.ForeColor = RGB(0, 255, 0)
    Label3.ForeColor = RGB(0, 255, 0)

End Sub

Private Sub Label2_Click()
    
    Unload frmWindowFind
    Form1.Show

End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Label2.ForeColor = &H8000&
    Label1.ForeColor = RGB(0, 255, 0)
    Label3.ForeColor = RGB(0, 255, 0)

End Sub

Private Sub Label3_Click()
    
    List1.Clear
    List2.Clear
    Call Form_Load

End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     
     Label3.ForeColor = &H4000&
     Label1.ForeColor = RGB(0, 255, 0)
     Label2.ForeColor = RGB(0, 255, 0)

End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    DragForm frmWindowFind

End Sub

Private Sub Label5_Click()
    
    NotOnTop frmWindowFind
    AskForExit
    OnTop frmWindowFind

End Sub

Private Sub Label6_Click()
    
    Call ShowWindow(frmWindowFind.hWnd, SW_MINIMIZE)

End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     
     Label1.ForeColor = RGB(0, 255, 0)
     Label2.ForeColor = RGB(0, 255, 0)
     Label3.ForeColor = RGB(0, 255, 0)

End Sub

Private Sub List2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     
     Label1.ForeColor = RGB(0, 255, 0)
     Label2.ForeColor = RGB(0, 255, 0)
     Label3.ForeColor = RGB(0, 255, 0)

End Sub
