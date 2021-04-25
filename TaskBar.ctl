VERSION 5.00
Begin VB.UserControl TaskBar 
   Alignable       =   -1  'True
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6840
   ScaleHeight     =   29
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   456
End
Attribute VB_Name = "TaskBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const default_ITEM_WIDTH As Single = 155
Private Const FIRST_OFFSET As Single = 1
Private Const STANDARD_OFFSET As Single = 3
Private Const ICON_WIDTH As Single = 18

Public Event OnTaskbarMenuRequired(ByRef menu As Object)
Public Event OnWindowMenuRequired(ByVal frm As Form, ByRef menu As Object)

Private m_nIndexBeingSelected As Integer
Private m_bInsetSelected As Boolean

Private m_maxCount As Integer
Private m_colIcons As Collection
Private m_refActive As clsIcon

Private m_cxBorder As Long
Private m_cyBorder As Long

Private m_nOptimalHeight As Long
Private m_nAlign As AlignConstants

Private m_strOriginalTooltip As String
Private m_bTooltip As Boolean

Friend Sub OnRefresh(Optional ByVal hWndActive As Long = 0)
    UpdateIconsCollection hWndActive
    MapIconCollection
End Sub

Friend Property Get Parent() As MDIForm
    On Error GoTo ErrorTrap
    Set Parent = UserControl.Parent
    Exit Property
ErrorTrap:
    Debug.Print "No parent or parent is not of class MDIForm"
End Property

Private Sub UserControl_Hide()
    If Ambient.UserMode() Then
        UnSubClassParentWnd Me
        ClearCollection
    End If
End Sub

Private Sub UserControl_InitProperties()
    If Parent Is Nothing Then
        Err.Raise 20000, "TaskBar", "TaskBar control may be placed on MDI froms only"
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not Ambient.UserMode() Then Exit Sub
    
    If Button = vbLeftButton Then
        m_nIndexBeingSelected = PointInElement(x, y)
        m_bInsetSelected = (m_nIndexBeingSelected > 0)
        If (m_nIndexBeingSelected > 0) Then
            SetCapture UserControl.hwnd
            InvalidateElement m_nIndexBeingSelected
        End If
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not Ambient.UserMode() Then Exit Sub
    
    If m_nIndexBeingSelected > 0 Then
        Dim bNewStatus As Boolean
        bNewStatus = IsPointInElement(x, y, m_nIndexBeingSelected)
        If m_bInsetSelected <> bNewStatus Then
            m_bInsetSelected = bNewStatus
            InvalidateElement m_nIndexBeingSelected
        End If
    ElseIf Button = 0 Then
        Dim nElPointed As Integer
        nElPointed = PointInElement(x, y)
        If nElPointed > 0 Then
            Dim rc As RECT
            Dim bDisp As Boolean
            bDisp = False
            If ItemRect(nElPointed, rc) Then
                If rc.Left + m_cxBorder < x And rc.Right - m_cxBorder > x And _
                    rc.Top + m_cyBorder < y And rc.Bottom - m_cyBorder > y Then
                    bDisp = UserControl.TextWidth(m_colIcons(nElPointed).Title) > rc.Right - rc.Left - ICON_WIDTH - 6
                End If
            End If
          
            If bDisp Then
                UserControl.Extender.ToolTipText = m_colIcons(nElPointed).Title
                m_bTooltip = True
            ElseIf m_bTooltip Then
                UserControl.Extender.ToolTipText = m_strOriginalTooltip
                m_bTooltip = False
            End If
        ElseIf m_bTooltip Then
            UserControl.Extender.ToolTipText = m_strOriginalTooltip
            m_bTooltip = False
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not Ambient.UserMode() Then Exit Sub
    
    If m_nIndexBeingSelected > 0 Then
        ReleaseCapture
        If IsPointInElement(x, y, m_nIndexBeingSelected) Then
            ActivateWindow m_nIndexBeingSelected
        End If
        If m_bInsetSelected Then InvalidateElement m_nIndexBeingSelected
        m_nIndexBeingSelected = 0
        m_bInsetSelected = False
    ElseIf Button = vbRightButton Then
        On Error Resume Next
        Dim menu As Object
        Dim nNewActive As Integer
        nNewActive = PointInElement(x, y)
        If 0 < nNewActive Then
            Dim frm As Form
            ActivateWindow nNewActive
            InvalidateElement nNewActive
            
            For Each frm In Forms
                If frm.hwnd = m_refActive.hwnd Then
                    frm.SetFocus
                End If
            Next
            RaiseEvent OnWindowMenuRequired(frm, menu)
            If Not menu Is Nothing Then frm.PopupMenu menu
        Else
            RaiseEvent OnTaskbarMenuRequired(menu)
            If Not menu Is Nothing Then Parent.PopupMenu menu
        End If
    End If
End Sub

Private Sub UserControl_Paint()
    Dim i As Integer
    Dim rcItem As RECT
    Dim icn As clsIcon
    Dim lEdgeParam As Long
    
    If Not Ambient.UserMode() Then
        Exit Sub
    End If
    
    If m_colIcons Is Nothing Then
        UserControl.Cls
        Exit Sub
    End If
    
    
    i = 0
    
    For Each icn In m_colIcons
        If ItemRect(i + 1, rcItem) Then
            If icn Is m_refActive Then
                DrawEdge UserControl.hDC, rcItem, EDGE_SUNKEN, BF_RECT
                UserControl.Line (rcItem.Left + m_cxBorder, rcItem.Top + m_cyBorder) _
                -(rcItem.Right - m_cxBorder - 1, rcItem.Bottom - m_cyBorder - 1), vb3DHighlight, BF
                UserControl.FontBold = True
            ElseIf i + 1 = m_nIndexBeingSelected And m_bInsetSelected Then
                DrawEdge UserControl.hDC, rcItem, EDGE_SUNKEN, BF_RECT
                UserControl.Line (rcItem.Left + m_cxBorder, rcItem.Top + m_cyBorder) _
                -(rcItem.Right - m_cxBorder - 1, rcItem.Bottom - m_cyBorder - 1), vbButtonFace, BF
                UserControl.FontBold = False
            Else
                DrawEdge UserControl.hDC, rcItem, EDGE_RAISED, BF_RECT
                UserControl.Line (rcItem.Left + m_cxBorder, rcItem.Top + m_cyBorder) _
                -(rcItem.Right - m_cxBorder - 1, rcItem.Bottom - m_cyBorder - 1), vbButtonFace, BF
                UserControl.FontBold = False
            End If
            
            rcItem.Left = rcItem.Left + m_cxBorder + 1
            rcItem.Top = rcItem.Top + m_cyBorder
            rcItem.Right = rcItem.Right - m_cxBorder - 1
            rcItem.Bottom = rcItem.Bottom - m_cyBorder - 1
            
            Dim nDiff As Single
            nDiff = rcItem.Bottom - rcItem.Top
            
            If Not icn.Icon Is Nothing Then
                Dim nIconTop As Single
                nIconTop = rcItem.Top + (nDiff - ICON_WIDTH) \ 2
                UserControl.PaintPicture icn.Icon, rcItem.Left, nIconTop, ICON_WIDTH, ICON_WIDTH
            End If
            
            rcItem.Left = rcItem.Left + ICON_WIDTH + 2
            Dim lpDrawTextParams As DRAWTEXTPARAMS
            lpDrawTextParams.iLeftMargin = 1
            lpDrawTextParams.iRightMargin = 1
            lpDrawTextParams.iTabLength = 2
            lpDrawTextParams.cbSize = 20
            
            Dim nTextH As Single
            nTextH = UserControl.TextHeight(icn.Title)
            If nTextH < nDiff Then
                nDiff = (nDiff - nTextH) \ 2
                rcItem.Bottom = rcItem.Bottom - nDiff
                rcItem.Top = rcItem.Top + nDiff
            End If
            
            
            DrawTextEx UserControl.hDC, icn.Title, Len(icn.Title), rcItem, _
            DT_LEFT + DT_VCENTER + DT_END_ELLIPSIS, lpDrawTextParams
        End If
        i = i + 1
    Next
End Sub

Private Sub UserControl_Show()
    If Ambient.UserMode() Then
        SubClassParentWnd Me
        m_strOriginalTooltip = UserControl.Extender.ToolTipText
        OnRefresh
    End If
    m_cxBorder = GetSystemMetrics(SM_CXEDGE)
    m_cyBorder = GetSystemMetrics(SM_CYEDGE)
    m_nOptimalHeight = GetSystemMetrics(SM_CYCAPTION)
    m_nOptimalHeight = m_nOptimalHeight + 2 * m_cyBorder + 3
    
    m_nAlign = UserControl.Extender.Align
    If m_nAlign = vbAlignBottom Or m_nAlign = vbAlignTop Then
        UserControl.Extender.Height = ScaleY(m_nOptimalHeight, vbPixels, vbTwips)
    End If
End Sub

Private Sub UpdateIconsCollection(ByVal hActive As Long)
On Error GoTo UpdateIconsErrorTrap
    
    Dim frm As Form
    
    InitCollection
    For Each frm In Forms
        If IsMDIChild(frm) Then
            VerifyItem frm
        End If
    Next
    
    Dim icn As clsIcon
    Dim iColIdx As Integer
    
    iColIdx = 1
    Set m_refActive = Nothing
    Do While iColIdx <= m_colIcons.Count
        Set icn = m_colIcons.Item(iColIdx)
        
        If icn.hwnd = hActive Then
            Set m_refActive = icn
            iColIdx = iColIdx + 1
        ElseIf Not icn.IsTaught Then
            m_colIcons.Remove iColIdx
        Else
            iColIdx = iColIdx + 1
        End If
    Loop
    
    m_maxCount = m_colIcons.Count
    Exit Sub
    
UpdateIconsErrorTrap:
    Debug.Print "Error occured in UpdateIconsCollection; code" + Str$(Err.Number) + " " + Err.Description
End Sub

Private Sub InitCollection()
    If m_colIcons Is Nothing Then
        Set m_colIcons = New Collection
        m_maxCount = 0
        Exit Sub
    End If
    
    Dim icn As clsIcon
    m_maxCount = 0
    For Each icn In m_colIcons
        icn.ClearTouch
    Next
    m_maxCount = m_colIcons.Count
End Sub

Private Sub MapIconCollection()
    
    Dim icn As clsIcon
    Dim strState As String
    
    Static nLastPaintedCnt As Integer
    Static nLastPaintedAct As Integer
    
    If Not nLastPaintedCnt = m_colIcons.Count Then
        UserControl.Refresh
        
    Else
        Dim nElementIndex As Integer
        nElementIndex = 1
        For Each icn In m_colIcons
            If icn.IsNew Then
                If icn.State = vbMinimized Then
                    ShowWindow icn.hwnd, SW_HIDE
                End If
                UserControl.Refresh
                Exit For
            ElseIf icn.IsChanged Then
                If icn.State = vbMinimized Then
                    ShowWindow icn.hwnd, SW_HIDE
                End If
                InvalidateElement nElementIndex
            ElseIf icn Is m_refActive And nElementIndex <> nLastPaintedAct Then
                InvalidateElement nLastPaintedAct
                InvalidateElement nElementIndex
                nLastPaintedAct = nElementIndex
            End If
            
            nElementIndex = nElementIndex + 1
        Next
    End If
    nLastPaintedCnt = m_colIcons.Count
End Sub

Private Sub VerifyItem(ByRef frm As Form)
    Dim icn As clsIcon
    For Each icn In m_colIcons
        If frm.hwnd = icn.hwnd Then
            icn.Title = frm.Caption
            icn.State = frm.WindowState
            icn.Touch
            Exit Sub
        End If
    Next
    Set icn = New clsIcon
    icn.Title = frm.Caption
    icn.State = frm.WindowState
    icn.hwnd = frm.hwnd
    Set icn.Icon = frm.Icon
    
    m_colIcons.Add icn
End Sub

Private Function IsMDIChild(ByRef frm As Form) As Boolean
On Error GoTo SureItIsInvalid
    
    IsMDIChild = (frm.MDIChild And frm.ShowInTaskbar)
    Exit Function
SureItIsInvalid:
    IsMDIChild = False
End Function

Private Sub InvalidateElement(ByVal nElIdx As Integer)
    If nElIdx < 1 Then Exit Sub
    
    Dim nAllCnt As Integer
    nAllCnt = m_colIcons.Count
    If nElIdx > nAllCnt Then Exit Sub
    
    Dim lpRect As RECT
    If ItemRect(nElIdx, lpRect) Then
        InvalidateRect UserControl.hwnd, lpRect, False
    End If
End Sub

Private Function ItemRect(ByVal itmIdx As Integer, ByRef rItem As RECT) As Boolean
    Debug.Assert itmIdx > 0 And itmIdx <= m_colIcons.Count
    
    m_nAlign = UserControl.Extender.Align
    
    If m_nAlign = vbAlignLeft Or m_nAlign = vbAlignRight Then
        Dim nItemH As Long
        nItemH = m_nOptimalHeight - 3
        rItem.Left = UserControl.ScaleLeft + 1
        rItem.Right = UserControl.ScaleWidth - UserControl.ScaleLeft - 1
        If rItem.Right - rItem.Left > 0 Then
            rItem.Top = FIRST_OFFSET + (itmIdx - 1) * (nItemH + STANDARD_OFFSET)
            rItem.Bottom = rItem.Top + nItemH
            If rItem.Bottom > UserControl.ScaleTop + UserControl.ScaleHeight Then
                rItem.Left = 0
                rItem.Right = 0
                rItem.Top = 0
                rItem.Bottom = 0
                ItemRect = False
            Else
                ItemRect = True
            End If
        Else
            rItem.Left = 0
            rItem.Right = 0
            ItemRect = False
        End If
    ElseIf m_nAlign = vbAlignBottom Or m_nAlign = vbAlignTop Then
        rItem.Top = UserControl.ScaleTop + 1
        rItem.Bottom = UserControl.ScaleHeight - UserControl.ScaleTop - 1
        If rItem.Bottom - rItem.Top > 0 Then
            Dim nItemW As Long
            nItemW = (UserControl.ScaleWidth - 3) / m_maxCount - 3
            nItemW = IIf(nItemW > default_ITEM_WIDTH, default_ITEM_WIDTH, nItemW)
            rItem.Left = UserControl.ScaleLeft + FIRST_OFFSET + (itmIdx - 1) * (nItemW + STANDARD_OFFSET)
            rItem.Right = rItem.Left + nItemW
            ItemRect = True
        Else
            rItem.Top = 0
            rItem.Bottom = 0
            ItemRect = False
        End If
    End If
End Function


Private Function PointInElement(ByVal x As Single, ByVal y As Single) As Integer
    m_nAlign = UserControl.Extender.Align
    Dim nEl As Integer
    If m_nAlign = vbAlignLeft Or m_nAlign = vbAlignRight Then
        PointInElement = 0
        If x > UserControl.ScaleLeft + 1 And x < UserControl.ScaleWidth - UserControl.ScaleLeft - 1 Then
            Dim nItemH As Long
            nItemH = m_nOptimalHeight - 3
            
            nEl = Int((y - UserControl.ScaleTop - FIRST_OFFSET) / (nItemH + STANDARD_OFFSET)) + 1
            If Not (nEl > m_maxCount Or nEl < 0) Then
                If (y - UserControl.ScaleTop - FIRST_OFFSET) - (nEl - 1) * (nItemH + STANDARD_OFFSET) > -2 Then PointInElement = nEl
            End If
        End If
    ElseIf m_nAlign = vbAlignBottom Or m_nAlign = vbAlignTop Then
        PointInElement = 0
        If y > UserControl.ScaleTop + 1 And y < UserControl.ScaleHeight - UserControl.ScaleTop - 1 Then
            Dim nItemW As Long
            
            nItemW = (UserControl.ScaleWidth - 3) / m_maxCount - 3
            nItemW = IIf(nItemW > default_ITEM_WIDTH, default_ITEM_WIDTH, nItemW)
            
            nEl = Int((x - UserControl.ScaleLeft - FIRST_OFFSET) / (nItemW + STANDARD_OFFSET)) + 1
            If Not (nEl > m_maxCount Or nEl < 0) Then
                If (x - UserControl.ScaleLeft - FIRST_OFFSET) - (nEl - 1) * (nItemW + STANDARD_OFFSET) > -2 Then PointInElement = nEl
            End If
        End If
    End If
End Function

Private Function IsPointInElement(ByVal x As Single, ByVal y As Single, ByVal idx As Integer) As Boolean
    m_nAlign = UserControl.Extender.Align
    Dim nEl As Integer
    
    If m_nAlign = vbAlignLeft Or m_nAlign = vbAlignRight Then
        IsPointInElement = False
        If x > UserControl.ScaleLeft + 1 And x < UserControl.ScaleWidth - UserControl.ScaleLeft - 1 Then
            Dim nItemH As Long
            nItemH = m_nOptimalHeight - 3
            
            Dim yOffs As Single
            yOffs = y - UserControl.ScaleTop - FIRST_OFFSET
            
            IsPointInElement = (y > (idx - 1) * (nItemH + STANDARD_OFFSET)) And (y < idx * (nItemH + STANDARD_OFFSET) - STANDARD_OFFSET)
        End If
    ElseIf m_nAlign = vbAlignBottom Or m_nAlign = vbAlignTop Then
        IsPointInElement = False
        If y > UserControl.ScaleTop + 1 And y < UserControl.ScaleHeight - UserControl.ScaleTop - 1 Then
            Dim nItemW As Long
            
            nItemW = (UserControl.ScaleWidth - 3) / m_maxCount - 3
            nItemW = IIf(nItemW > default_ITEM_WIDTH, default_ITEM_WIDTH, nItemW)
                        
            Dim xOffs As Single
            xOffs = x - UserControl.ScaleLeft - FIRST_OFFSET
            
            IsPointInElement = (x > (idx - 1) * (nItemW + STANDARD_OFFSET)) And (x < idx * (nItemW + STANDARD_OFFSET) - STANDARD_OFFSET)
        End If
    End If
End Function

Private Sub ActivateWindow(ByVal nEl As Integer)
    
    If nEl < 1 Or nEl > m_maxCount Then Exit Sub
    On Error GoTo ActivateFailed
    
    Set m_refActive = m_colIcons(nEl)
    If m_refActive.State = vbMinimized Then
        ShowWindow m_refActive.hwnd, SW_SHOW
        ShowWindow m_refActive.hwnd, SW_RESTORE
    Else
        Dim frm As Form
        For Each frm In Forms
            If frm.hwnd = m_refActive.hwnd Then
                frm.SetFocus
            End If
        Next
    End If
    
ActivateFailed:
End Sub

Private Sub ClearCollection()
On Error GoTo ClearCollectionError
    m_nIndexBeingSelected = 0
    m_bInsetSelected = False

    Dim icn As clsIcon
    For Each icn In m_colIcons
        If icn.State = vbMinimized Then
            ShowWindow icn.hwnd, SW_SHOW
        End If
    Next
    Exit Sub
ClearCollectionError:
    Debug.Print "Error code: " + Str$(Err.Number) + " in cleanup"
End Sub
