VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_strTitle As String
Private m_hwnd As Long
Private m_nState As FormWindowStateConstants
Private m_refIcon As IPictureDisp
Private m_bTouch As Boolean
Private m_bNew As Boolean
Private m_bChanged As Boolean

Public Property Get IsNew() As Boolean
    IsNew = m_bNew
End Property

Public Property Get Title() As String
    Title = m_strTitle
End Property

Public Property Let Title(ByVal strNewValue As String)
    m_strTitle = strNewValue
    m_bChanged = True
End Property

Public Property Get hwnd() As Long
    hwnd = m_hwnd
End Property

Public Property Let hwnd(ByVal hNewValue As Long)
    m_hwnd = hNewValue
End Property

Public Property Get State() As FormWindowStateConstants
    State = m_nState
End Property

Public Property Let State(ByVal newStateValue As FormWindowStateConstants)
    m_nState = newStateValue
End Property

Friend Sub ClearTouch()
    m_bNew = False
    m_bTouch = False
    m_bChanged = False
End Sub

Friend Sub Touch()
    m_bTouch = True
End Sub

Friend Function IsTaught() As Boolean
    IsTaught = m_bTouch
End Function

Private Sub Class_Initialize()
    m_bNew = True
    m_bTouch = True
    m_bChanged = True
End Sub

Public Property Get IsChanged() As Boolean
    IsChanged = m_bChanged
End Property

Public Property Set Icon(ByRef newIcon As IPictureDisp)
    On Error GoTo SetNoIcon
    
    Set m_refIcon = newIcon
    Exit Property
    
SetNoIcon:
    Set m_refIcon = Nothing
End Property

Public Property Get Icon() As IPictureDisp
    On Error GoTo GetNoIcon
    
    Set Icon = m_refIcon
    Exit Property
    
GetNoIcon:
    
End Property
