VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmChild 
   Caption         =   "Connecting..."
   ClientHeight    =   4365
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   6885
   Icon            =   "frmChild.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4365
   ScaleWidth      =   6885
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   2955
      Top             =   1605
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   4110
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2822
            MinWidth        =   2822
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2280
      Top             =   1590
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open"
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      DragMode        =   1  'Automatic
      Height          =   1290
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   2275
      BandCount       =   2
      _CBWidth        =   6885
      _CBHeight       =   1290
      _Version        =   "6.0.8169"
      Child1          =   "Toolbar1"
      MinHeight1      =   840
      Width1          =   4740
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      MinHeight2      =   360
      NewRow2         =   -1  'True
      AllowVertical2  =   0   'False
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   165
         TabIndex        =   0
         Top             =   915
         Width           =   4305
      End
      Begin VB.CommandButton Command1 
         Caption         =   "GO"
         Default         =   -1  'True
         Height          =   315
         Left            =   4500
         MouseIcon       =   "frmChild.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   915
         Width           =   555
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   840
         Left            =   165
         TabIndex        =   4
         Top             =   30
         Width           =   6630
         _ExtentX        =   11695
         _ExtentY        =   1482
         ButtonWidth     =   1191
         ButtonHeight    =   1376
         Wrappable       =   0   'False
         ImageList       =   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   11
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Back"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Forward"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Stop"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Refresh"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Home"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Search"
               ImageIndex      =   6
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   5
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Google"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Yahoo!"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "AltaVista"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Lycos"
                  EndProperty
                  BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Ask Jeeves!"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Print"
               ImageIndex      =   7
            EndProperty
         EndProperty
      End
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1590
      Left            =   15
      TabIndex        =   2
      Top             =   1305
      Width           =   5580
      ExtentX         =   9842
      ExtentY         =   2805
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3525
      Top             =   3705
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":0884
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":0CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":112C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":1580
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":1980
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":1DD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":2260
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        WebBrowser1.Navigate2 Combo1.Text
        Combo1.AddItem Combo1.Text
    End If
End Sub

Private Sub Command1_Click()
    WebBrowser1.Navigate2 Combo1.Text
    Combo1.AddItem Combo1.Text
End Sub

Private Sub Form_GotFocus()
    WebBrowser1.SetFocus
End Sub

Private Sub Form_Resize()
    If WindowState <> 1 Then
        If (Height >= 3615) Then
            If (Width >= 5760) Then
                WebBrowser1.Height = Height - 1950
                WebBrowser1.Width = Width - 150
                Combo1.Width = Width - 930
                Command1.Left = Width - 740
            Else
                Width = 5760
                WebBrowser1.Height = Height - 1950
                WebBrowser1.Width = Width - 150
                Combo1.Width = Width - 930
                Command1.Left = Width - 740
            End If
        Else
            Height = 3615
            WebBrowser1.Height = Height - 1950
            WebBrowser1.Width = Width - 150
            Combo1.Width = Width - 930
            Command1.Left = Width - 740
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.BrowserCount = frmMain.BrowserCount - 1
    
    If frmMain.BrowserCount = 1 Then
        frmMain.Caption = "SuperBrowser - " & frmMain.BrowserCount & " active navigation"
    Else
        frmMain.Caption = "SuperBrowser - " & frmMain.BrowserCount & " active navigations"
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error GoTo ErrHandle
    Select Case Button.Index
        Case 1:
            On Error Resume Next
            WebBrowser1.GoBack
        Case 2:
            On Error Resume Next
            WebBrowser1.GoForward
        Case 4:
            WebBrowser1.Stop
        Case 6:
            WebBrowser1.Refresh2
        Case 8:
            Open "c:\blank.html" For Output As #1
            Close #1
            WebBrowser1.Navigate "c:\blank.html"
        Case 9:
            WebBrowser1.Navigate2 "http://www.yahoo.com"
        Case 11:
            DoEvents
            WebBrowser1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER
    End Select

ErrHandle:
    Dim intPress As Integer
    intPress = MsgBox("An error has occured.", vbOKOnly + vbCritical, "Error")
    Exit Sub
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    On Error GoTo ErrHandle
    Select Case ButtonMenu.Index
        Case 1:
            WebBrowser1.Navigate2 "www.google.com"
        Case 2:
            WebBrowser1.Navigate2 "www.yahoo.com"
        Case 3:
            WebBrowser1.Navigate2 "www.altavista.com"
        Case 4:
            WebBrowser1.Navigate2 "www.lycos.com"
        Case 5:
            WebBrowser1.Navigate2 "www.ask.com"
    End Select

ErrHandle:
    Dim intPress As Integer
    intPress = MsgBox("An error has occured.", vbOKOnly + vbCritical, "Error")
    Exit Sub
End Sub

Private Sub WebBrowser1_DownloadComplete()
    StatusBar1.Panels.Item(1).Text = "Complete."
    StatusBar1.Panels.Item(2).Text = ""
End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    Combo1.Text = WebBrowser1.LocationURL
    Caption = WebBrowser1.LocationName
End Sub

Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
    On Error GoTo ErrHandle
    Dim frm As frmChild
    Set frm = New frmChild
    frm.Height = 4770
    frm.Width = 7005
    Set ppDisp = frm.WebBrowser1.Object
    frm.Show
    Set frm = Nothing
    
ErrHandle:
    Dim intPress As Integer
    intPress = MsgBox("An error has occured.", vbOKOnly + vbCritical, "Error")
    Exit Sub
End Sub

Private Sub WebBrowser1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
    If Progress > 0 Then
        StatusBar1.Panels.Item(1).Text = "Downloading site..."
        StatusBar1.Panels.Item(2).Text = Progress & "/" & ProgressMax
    Else
        StatusBar1.Panels.Item(1).Text = "Complete."
        StatusBar1.Panels.Item(2).Text = ""
    End If
End Sub
