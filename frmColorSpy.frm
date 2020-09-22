VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "§pIdEr's API Spy"
   ClientHeight    =   5415
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Lock Mode"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   0
      TabIndex        =   26
      Top             =   3240
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   2415
      Left            =   5640
      ScaleHeight     =   2355
      ScaleWidth      =   1635
      TabIndex        =   25
      Top             =   960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Do stuff"
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   5160
      Width           =   5415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "E&xit"
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sto&p"
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "S&tart"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5040
      Top             =   3000
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1395
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   3480
      Width           =   5415
   End
   Begin VB.Label Label19 
      BackColor       =   &H00000000&
      Caption         =   " X"
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
      Left            =   5160
      TabIndex        =   24
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label18 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "  _"
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
      Left            =   4920
      TabIndex        =   23
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label20 
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
      TabIndex        =   22
      Top             =   0
      Width           =   5415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      X1              =   0
      X2              =   5880
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1920
      TabIndex        =   21
      Top             =   3000
      Width           =   3495
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1920
      TabIndex        =   20
      Top             =   2760
      Width           =   3495
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1920
      TabIndex        =   19
      Top             =   2520
      Width           =   3495
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1920
      TabIndex        =   17
      Top             =   2280
      Width           =   3495
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1920
      TabIndex        =   16
      Top             =   2040
      Width           =   3495
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1920
      TabIndex        =   15
      Top             =   1800
      Width           =   3495
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1920
      TabIndex        =   14
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1920
      TabIndex        =   13
      Top             =   1320
      Width           =   3735
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1920
      TabIndex        =   12
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000FF00&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1920
      TabIndex        =   11
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1920
      TabIndex        =   10
      Top             =   840
      Width           =   3735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   600
      Width           =   3735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   1920
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bSpying As Boolean

Sub DoSpy()
    Dim ParentWindow As String, Parent As Long, parent_ As Long
    Dim ParentTextLen As Long, ParentText As String, RectMain As Rect
    Dim RetValue As Long
            
    On Error Resume Next
    
    'Get main Handle
    Label8.Caption = "Main handle:  " & WindowFromPoint(CurPos.X, CurPos.Y)
    
    'Get mouse coordinates
    Label15.Caption = "Mouse pos. X:  " & CurPos.X
    Label16.Caption = "Mouse pos. Y:  " & CurPos.Y
    
    'Get parent's handle
    Label9.Caption = "Parent's handle:  " & GetParent(Win)
    
    'Get main text
    MainTextLen = GetWindowTextLength(Win) + 1
    MainText = Space(MainTextLen)
    Call GetWindowText(Win, MainText, MainTextLen)
    Label10.Caption = "Window Text:  " & MainText
    
    'Get control text
    RetValue = SendMessage(Win, WM_GETTEXTLENGTH, ByVal CLng(0), ByVal CLng(0)) + 1
    MainText = Space(RetValue)
    RetValue = SendMessage(Win, WM_GETTEXT, ByVal RetValue, ByVal MainText)
    Label17.Caption = "Control text:  " & MainText
    'Get parent's text
    ParentTextLen = GetWindowTextLength(GetParent(Win)) + 1
    ParentText = Space(ParentTextLen)
    Call GetWindowText(GetParent(Win), ParentText, ParentTextLen)
    Label11.Caption = "Parent Text:  " & ParentText
    
    'Get main class name
    Win = WindowFromPoint(CurPos.X, CurPos.Y)
    MainClassName = Space(255)
    Call GetCursorPos(CurPos)
    Call GetClassName(Win, MainClassName, 255)
    Label6.Caption = "Class name:  " & MainClassName
        
    'Get parent class name
    ParentWindow = Space(255)
    Parent = GetParent(Win)
    parent_ = GetClassName(Parent, ParentWindow, 255)
    Label7.Caption = "Parent class name:  " & ParentWindow
    
    'Get Main Width and Height
    Call GetWindowRect(Win, RectMain)
    Label12.Caption = "Window width:  " & RectMain.Right - RectMain.Left
    Label13.Caption = "Window height:  " & RectMain.Bottom - RectMain.Top
    
    'Get main window state
    If (Not IsIconic(Win)) And (Not IsZoomed(Win)) Then Label14.Caption = "Window state:  General"
    If IsIconic(Win) Then Label14.Caption = "Window state:  Minimized"
    If IsZoomed(Win) Then Label14.Caption = "Window state:  Maximized"
    
    'Color stuff
    CurColor = GetRGB(Picture1.BackColor)
    Call GetCursorPos(CurPos)
    MainDC = GetDC(0)
    Picture1.BackColor = GetPixel(MainDC, CurPos.X, CurPos.Y)
    Call ReleaseDC(0, MainDC)
    Label1.Caption = "Hexadecimal:  " & Hex(Picture1.BackColor)
    Label2.Caption = "Red:  " & CurColor.Red
    Label3.Caption = "Green:  " & CurColor.Green
    Label4.Caption = "Blue:  " & CurColor.Blue
    Label5.Caption = "Decimal:  " & Picture1.BackColor

End Sub

Private Sub Check1_Click()
    Timer1.Enabled = True
    If Command1.Enabled = False Then
    Command1.Enabled = True
    Timer1.Enabled = False
    Exit Sub
    End If
    Command1.Enabled = False
End Sub

Private Sub Command1_Click()
    Timer1.Enabled = True
    Command2.Enabled = True
    Command1.Enabled = False
    Check1.Enabled = False
    bSpying = True
End Sub

Private Sub Command2_Click()
    Timer1.Enabled = False
    Command1.Enabled = True
    Command2.Enabled = False
    Check1.Enabled = True
End Sub

Private Sub Command3_Click()
    NotOnTop Me
    AskForExit
    OnTop Me
End Sub

Private Sub Command4_Click()
    Form1.PopupMenu Form2.mnuDoStuff
End Sub

Private Sub Command5_Click()
    Call ShowWindow(Form1.hWnd, SW_MINIMIZE)
End Sub

Private Sub Form_Load()
    Dim Sugar As Long
    Command2.Enabled = False
    With Picture2
        .Width = Form1.Width + 50
        .Height = Form1.Height + 50
    End With
    
    OnTop Me
End Sub

Private Sub Form_LostFocus()
    OnTop Me
End Sub

Private Sub Form_Resize()
    OnTop Me
End Sub


Private Sub Label18_Click()
    Call ShowWindow(Form1.hWnd, SW_MINIMIZE)
End Sub

Private Sub Label19_Click()
    NotOnTop Me
    AskForExit
    OnTop Me
End Sub

Private Sub Label20_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then DragForm Form1
End Sub

Private Sub Timer1_Timer()
    If Check1.Value = vbChecked Then
        
        If GetAsyncKeyState(vbKeyL) Then
'wrote it a few times to be sure all the info get updated in one keystroke
'they weren't all getting updated at the same time when it was written only once
            DoSpy
            DoSpy
            DoSpy
        End If
        
    Else
                
        If bSpying = True Then
            
            DoSpy
        End If
    End If
End Sub
