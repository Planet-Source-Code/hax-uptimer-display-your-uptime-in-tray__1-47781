VERSION 5.00
Begin VB.UserControl TrayIcon 
   CanGetFocus     =   0   'False
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   540
   InvisibleAtRuntime=   -1  'True
   Picture         =   "Trayicon.ctx":0000
   ScaleHeight     =   35
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   36
   ToolboxBitmap   =   "Trayicon.ctx":00F0
   Begin VB.Image imgIcon 
      Height          =   240
      Left            =   960
      Picture         =   "Trayicon.ctx":0402
      Top             =   120
      Width           =   240
   End
End
Attribute VB_Name = "TrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type NOTIFYICONDATA
  cbSize            As Long
  hwnd              As Long
  uID               As Long
  uFlags            As Long
  uCallbackMessage  As Long
  hIcon             As Long
  szTip             As String * 64
End Type
    
Private Const NIM_ADD = 0
Private Const NIM_MODIFY = 1
Private Const NIM_DELETE = 2
Private Const NIF_MESSAGE = 1
Private Const NIF_ICON = 2
Private Const NIF_TIP = 4

Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208
Private Const WM_MBUTTONDBLCLK = &H209

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Private pobj_Icon     As IPictureDisp
Private pstr_Tooltip  As String

Public Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer)
Public Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer)
Public Event DblClick(ByVal Button As Integer, ByVal Shift As Integer)

Public Function HideIcon() As Boolean
  Dim tNID As NOTIFYICONDATA
  Dim lRet As Long
  
  With tNID
    .cbSize = Len(tNID)
    .uID = 1&
    .hwnd = UserControl.hwnd
  End With
    
  lRet = Shell_NotifyIcon(NIM_DELETE, tNID)
  
  If lRet <> 0 Then HideIcon = True
End Function

Public Property Set Icon(ByVal New_Icon As IPictureDisp)
  Set pobj_Icon = New_Icon
  pChangeIcon
End Property

Public Property Get Icon() As IPictureDisp
  On Error Resume Next
  If pobj_Icon Is Nothing Then
    Set Icon = imgIcon.Picture
  Else
    Set Icon = pobj_Icon
  End If
End Property

Public Property Let InfoTip(ByVal New_Value As String)
  If Len(New_Value) > 63 Then
    pstr_Tooltip = Left$(New_Value, 63)
  Else
    pstr_Tooltip = New_Value
  End If
  
  pChangeIcon
End Property

Public Property Get InfoTip() As String
  InfoTip = pstr_Tooltip
End Property

Private Sub pChangeIcon()
  Dim tNID As NOTIFYICONDATA
  Dim lRet As Long

  If (Ambient.UserMode = False) Then Exit Sub
  
  With tNID
    .cbSize = Len(tNID)
    .uID = 1&
    .hwnd = UserControl.hwnd
    .hIcon = Icon.Handle
    .szTip = pstr_Tooltip & Chr(0)
    .uFlags = NIF_ICON Or NIF_TIP
  End With
    
  lRet = Shell_NotifyIcon(NIM_MODIFY, tNID)
End Sub

Public Sub PopupMenu(ByRef Menu As Object, Optional ByVal Flags As MenuControlConstants, Optional ByRef Default As Variant)
  SetForegroundWindow UserControl.Parent.hwnd
  
  If IsMissing(Default) Then
    UserControl.Parent.PopupMenu Menu, Flags
  Else
    UserControl.Parent.PopupMenu Menu, Flags, , , Default
  End If
End Sub

Public Function ShowIcon() As Boolean
  Dim tNID As NOTIFYICONDATA
  Dim lRet As Long
  
  With tNID
    .cbSize = Len(tNID)
    .hwnd = UserControl.hwnd
    .uID = 1&
    .szTip = pstr_Tooltip & Chr(0)
    .uCallbackMessage = WM_MOUSEMOVE
    .hIcon = Icon.Handle
    .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
  End With
  lRet = Shell_NotifyIcon(NIM_ADD, tNID)
  
  If lRet <> 0 Then ShowIcon = True
End Function

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim lMsg        As Long
  
  lMsg = X

  Select Case lMsg
    Case WM_LBUTTONDOWN
      RaiseEvent MouseDown(vbLeftButton, Shift)
      
    Case WM_LBUTTONUP
      RaiseEvent MouseUp(vbLeftButton, Shift)
    
    Case WM_LBUTTONDBLCLK
      RaiseEvent DblClick(vbLeftButton, Shift)
      
    Case WM_MBUTTONDOWN
      RaiseEvent MouseDown(vbMiddleButton, Shift)
      
    Case WM_MBUTTONUP
      RaiseEvent MouseUp(vbMiddleButton, Shift)
    
    Case WM_MBUTTONDBLCLK
      RaiseEvent DblClick(vbMiddleButton, Shift)
      
    Case WM_RBUTTONDOWN
      RaiseEvent MouseDown(vbRightButton, Shift)
      
    Case WM_RBUTTONUP
      RaiseEvent MouseUp(vbRightButton, Shift)
    
    Case WM_RBUTTONDBLCLK
      RaiseEvent DblClick(vbRightButton, Shift)

  End Select
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  With PropBag
    pstr_Tooltip = .ReadProperty("InfoTip", "")
    Set pobj_Icon = .ReadProperty("Icon", Nothing)
  End With
  pChangeIcon
End Sub

Private Sub UserControl_Resize()
  UserControl.Size 510, 510
End Sub

Private Sub UserControl_Terminate()
  HideIcon
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  With PropBag
    .WriteProperty "InfoTip", pstr_Tooltip, ""
    .WriteProperty "Icon", pobj_Icon
  End With
End Sub
