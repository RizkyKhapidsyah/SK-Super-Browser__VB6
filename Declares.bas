Attribute VB_Name = "Declares"
Option Explicit

Public Const SM_CXEDGE = 45
Public Const SM_CYEDGE = 46
Public Const SM_CYCAPTION = 4

Public Const BDR_INNER = &HC
Public Const BDR_OUTER = &H3
Public Const BDR_RAISED = &H5
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKEN = &HA
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2
Public Const BF_ADJUST = &H2000
Public Const BF_BOTTOM = &H8
Public Const BF_DIAGONAL = &H10
Public Const BF_FLAT = &H4000
Public Const BF_LEFT = &H1
Public Const BF_MIDDLE = &H800
Public Const BF_MONO = &H8000
Public Const BF_RIGHT = &H4
Public Const BF_SOFT = &H1000
Public Const BF_TOP = &H2
Public Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Public Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Public Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Public Const BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
Public Const BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)
Public Const BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
Public Const BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
Public Const DT_BOTTOM = &H8
Public Const DT_CALCRECT = &H400
Public Const DT_CHARSTREAM = 4
Public Const DT_CENTER = &H1
Public Const DT_DISPFILE = 6
Public Const DT_EXPANDTABS = &H40
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_INTERNAL = &H1000
Public Const DT_LEFT = &H0
Public Const DT_METAFILE = 5
Public Const DT_NOCLIP = &H100
Public Const DT_NOPREFIX = &H800
Public Const DT_PLOTTER = 0
Public Const DT_RASCAMERA = 3
Public Const DT_RASDISPLAY = 1
Public Const DT_RASPRINTER = 2
Public Const DT_RIGHT = &H2
Public Const DT_SINGLELINE = &H20
Public Const DT_TABSTOP = &H80
Public Const DT_TOP = &H0
Public Const DT_VCENTER = &H4
Public Const DT_WORDBREAK = &H10
Public Const DT_PATH_ELLIPSIS = &H4000
Public Const DT_END_ELLIPSIS = &H8000

Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)

Public Const SW_HIDE = 0
Public Const SW_NORMAL = 1
Public Const SW_SHOW = 5
Public Const SW_RESTORE = 9


Const WM_MDIICONARRANGE = &H228
Const WM_MDIACTIVATE = &H222
Const WM_MDICREATE = &H220
Const WM_MDIDESTROY = &H221
Const WM_MDINEXT = &H224
Const WM_MDIMAXIMIZE = &H225
Const WM_MDIRESTORE = &H223

Const WM_MDIGETACTIVE = &H229
Const WM_MDIREFRESHMENU = &H234
Const WM_MDISETMENU = &H230

Const GWL_WNDPROC = (-4)

Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hDC As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As DRAWTEXTPARAMS) As Long
Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long


Type DRAWTEXTPARAMS
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private objTB As TaskBar
Private pfWndProc As Long
Private hSubclassed As Long

Public Sub SubClassParentWnd(ByRef obj As TaskBar)
    Dim hwnd As Long
    If Not obj.Parent Is Nothing Then
        hwnd = obj.Parent.hwnd
        Set objTB = obj
        EnumChildWindows hwnd, AddressOf CatchMDIClientWND, 0
    End If
End Sub

Public Sub UnSubClassParentWnd(ByRef obj As TaskBar)
    If Not obj.Parent Is Nothing Then
        SetWindowLong hSubclassed, GWL_WNDPROC, pfWndProc
        Set objTB = Nothing
    End If
End Sub

Function MDI_ParentWndProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim lRet As Long
    lRet = CallWindowProc(pfWndProc, hwnd, Msg, wParam, lParam)
    If Msg = WM_MDIGETACTIVE Then
            objTB.OnRefresh lRet
    End If
    MDI_ParentWndProc = lRet
End Function

Function CatchMDIClientWND(ByVal hwnd As Long, ByRef lParam As Long) As Long
    Dim lStat As Long
    Dim lpClassName As String
    Dim nLen As Long
        
    lpClassName = String(128, 32)
    nLen = Len(lpClassName)
    
    CatchMDIClientWND = 1
    
    lStat = GetClassName(hwnd, lpClassName, nLen)
    If (lStat > 0) Then
        Dim szName As String
        szName = Left$(lpClassName, lStat)
        If szName = "MDIClient" Then
            ' that's it
            CatchMDIClientWND = 0
            hSubclassed = hwnd
            pfWndProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf MDI_ParentWndProc)
        End If
    End If
End Function
