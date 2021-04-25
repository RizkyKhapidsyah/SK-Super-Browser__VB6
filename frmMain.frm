VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "SuperBrowser"
   ClientHeight    =   5385
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   6990
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2370
      Top             =   1620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open"
   End
   Begin SuperBrowser.TaskBar TaskBar1 
      Align           =   3  'Align Left
      Height          =   5385
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   9499
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New Navigation"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu Divider1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsTaskbar 
         Caption         =   "&Taskbar"
         Begin VB.Menu mnuOptionsTaskbarLeft 
            Caption         =   "&Left"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuOptionsTaskbarRight 
            Caption         =   "&Right"
         End
         Begin VB.Menu mnuOptionsTaskbarTop 
            Caption         =   "&Top"
         End
         Begin VB.Menu mnuOptionsTaskbarBottom 
            Caption         =   "&Botton"
         End
         Begin VB.Menu Divider2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOptionsTaskbarDisable 
            Caption         =   "&Disable"
         End
         Begin VB.Menu mnuOptionsTaskbarEnable 
            Caption         =   "&Enable"
         End
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu Divider3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu Divider4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpFeatures 
         Caption         =   "&Features"
      End
      Begin VB.Menu Divider5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About SuperBrowser"
      End
   End
   Begin VB.Menu mnuInvisible 
      Caption         =   "Invisible"
      Visible         =   0   'False
      Begin VB.Menu mnuInvisibleRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mnuInvisibleExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuVisible 
      Caption         =   "Visible"
      Visible         =   0   'False
      Begin VB.Menu mnuVisibleHide 
         Caption         =   "&Hide"
      End
      Begin VB.Menu mnuVisibleExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public BrowserCount
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hWnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Private Sub MDIForm_Load()
    LoadNewDoc
    With nid
        .cbSize = Len(nid)
        .hWnd = Me.hWnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "SuperBrowser" & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, nid
End Sub

Private Sub LoadNewDoc()
    Dim frmD As frmChild

    BrowserCount = BrowserCount + 1
    Set frmD = New frmChild
    frmD.Hide
    frmD.Caption = "Navigation " & BrowserCount
    
    If BrowserCount = 1 Then
        frmMain.Caption = "SuperBrowser - " & BrowserCount & " active navigation"
    Else
        frmMain.Caption = "SuperBrowser - " & BrowserCount & " active navigations"
    End If
    
    frmD.Height = 4770
    frmD.Width = 7005
    Open "c:\blank.html" For Output As #1
    Close #1
    frmD.WebBrowser1.Navigate "c:\blank.html"
    frmD.Show
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Result As Long
    Dim Msg As Long
    
    Msg = x / Screen.TwipsPerPixelX
      
    Select Case Msg
        Case WM_LBUTTONUP
            Me.WindowState = vbMaximized
            Result = SetForegroundWindow(Me.hWnd)
        Case WM_RBUTTONUP
            If frmMain.Visible = False Then
                PopupMenu mnuInvisible
            Else
                PopupMenu mnuVisible
            End If
        Case WM_LBUTTONDBLCLK
            If frmMain.Visible = False Then
                frmMain.WindowState = vbMinimized
                frmMain.Visible = True
                frmMain.WindowState = vbMaximized
            Else
                frmMain.WindowState = vbMinimized
                frmMain.Visible = False
            End If
        Case Else
            Me.WindowState = vbMaximized
            Result = SetForegroundWindow(Me.hWnd)
    End Select
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Me.WindowState = vbMinimized
    Me.Visible = False
End Sub

Private Sub mnuFileOpen_Click()
    On Error GoTo ErrHandle
    
    CommonDialog1.ShowOpen

    Dim frmD As frmChild

    Set frmD = New frmChild
    frmD.Hide
    frmD.Caption = CommonDialog1.FileName
    frmD.Height = 4770
    frmD.Width = 7005
    frmD.WebBrowser1.Navigate CommonDialog1.FileName
    frmD.Show
    
ErrHandle:
    Dim intPress As Integer
    intPress = MsgBox("An error has occured while trying to open this file.", vbOKOnly + vbCritical, "Error")
    Exit Sub
End Sub

Private Sub mnuHelpAbout_Click()
    Dim intPress As Integer
    intPress = MsgBox("SuperBrowser is a completely self-contained internet application, allowing the user to navigate multiple websites without cluttering the windows taskbar. It also utilizes an adaptable taskbar independant of the windows taskbar, allowing the user to change from navigation to navigation with only one click.", vbOKOnly + vbQuestion, "About")
End Sub

Private Sub mnuHelpFeatures_Click()
    Dim intPress
    intPress = MsgBox("To create a new navigation, either click File|New Navigation, or click F1. To change the location of the taskbar, choice the desired location from the submenu in Options|Taskbar. To open an image or file, click File|Open. For additional help, email Bryan Healey at Healeynator@yahoo.com", vbOKOnly + vbQuestion, "Features")
End Sub

Private Sub mnuInvisibleExit_Click()
    Shell_NotifyIcon NIM_DELETE, nid
    End
End Sub

Private Sub mnuInvisibleRestore_Click()
    Dim Result As Long
    Result = SetForegroundWindow(Me.hWnd)
    
    Me.Visible = True
    
    Me.WindowState = vbMaximized
End Sub

Private Sub mnuOptionsTaskbarBottom_Click()
    TaskBar1.Align = vbAlignBottom
    TaskBar1.Height = 390
    mnuOptionsTaskbarRight.Checked = False
    mnuOptionsTaskbarLeft.Checked = False
    mnuOptionsTaskbarTop.Checked = False
    mnuOptionsTaskbarBottom.Checked = True
End Sub

Private Sub mnuOptionsTaskbarDisable_Click()
    TaskBar1.Visible = False
    mnuOptionsTaskbarDisable.Visible = False
    mnuOptionsTaskbarEnable.Visible = True
End Sub

Private Sub mnuOptionsTaskbarEnable_Click()
    TaskBar1.Visible = True
    mnuOptionsTaskbarDisable.Visible = True
    mnuOptionsTaskbarEnable.Visible = False
End Sub

Private Sub mnuOptionsTaskbarLeft_Click()
    TaskBar1.Align = vbAlignLeft
    TaskBar1.Width = 1680
    mnuOptionsTaskbarRight.Checked = False
    mnuOptionsTaskbarLeft.Checked = True
    mnuOptionsTaskbarTop.Checked = False
    mnuOptionsTaskbarBottom.Checked = False
End Sub

Private Sub mnuOptionsTaskbarRight_Click()
    TaskBar1.Align = vbAlignRight
    TaskBar1.Width = 1680
    mnuOptionsTaskbarRight.Checked = True
    mnuOptionsTaskbarLeft.Checked = False
    mnuOptionsTaskbarTop.Checked = False
    mnuOptionsTaskbarBottom.Checked = False
End Sub

Private Sub mnuOptionsTaskbarTop_Click()
    TaskBar1.Align = vbAlignTop
    TaskBar1.Height = 390
    mnuOptionsTaskbarRight.Checked = False
    mnuOptionsTaskbarLeft.Checked = False
    mnuOptionsTaskbarTop.Checked = True
    mnuOptionsTaskbarBottom.Checked = False
End Sub

Private Sub mnuVisibleExit_Click()
    Shell_NotifyIcon NIM_DELETE, nid
    End
End Sub

Private Sub mnuVisibleHide_Click()
    Me.Visible = False
    
    Me.WindowState = vbMinimized
End Sub

Private Sub TaskBar1_OnTaskbarMenuRequired(menu As Object)
    Set menu = mnuWindow
End Sub

Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuFileExit_Click()
    Shell_NotifyIcon NIM_DELETE, nid
    End
End Sub

Private Sub mnuFileNew_Click()
    LoadNewDoc
End Sub
