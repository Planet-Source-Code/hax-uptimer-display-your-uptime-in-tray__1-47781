VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Uptimer"
   ClientHeight    =   1560
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5070
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command2 
      Caption         =   "Quit Uptimer"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Minimize to Tray"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   480
      Left            =   120
      Picture         =   "Form1.frx":08CA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   120
      Width           =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   960
      Top             =   2640
   End
   Begin uptimer.TrayIcon TrayIcon1 
      Left            =   360
      Top             =   2640
      _extentx        =   900
      _extenty        =   900
      icon            =   "Form1.frx":1194
   End
   Begin VB.Label Label2 
      Caption         =   "Uptimer by hax"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   1260
      Width           =   1095
   End
   Begin VB.Label label_Homepage 
      Caption         =   "http://www.hax-online.de"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   2280
      MouseIcon       =   "Form1.frx":1A70
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   3
      Top             =   1260
      Width           =   1875
   End
   Begin VB.Label Label1 
      Caption         =   "MyUptime"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   3975
   End
   Begin VB.Menu mnuTray 
      Caption         =   "mnuTray"
      Visible         =   0   'False
      Begin VB.Menu Show 
         Caption         =   "Show"
      End
      Begin VB.Menu Beenden 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyUptime As String

Private Sub Beenden_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Me.Hide
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Hide
  With TrayIcon1
    .InfoTip = MyUptime
    .ShowIcon
  End With
End Sub

Private Sub label_Homepage_Click()
Dim SearchVal
    Dim browserstring As String
    Dim launchbrowser As String
    browserstring = "rundll32.exe url.dll,FileProtocolHandler "
    launchbrowser = browserstring + "http://www.hax-online.de"
    SearchVal = Shell(launchbrowser, 0)
End Sub

Private Sub Show_Click()
    Me.WindowState = vbNormal
    Me.Show
End Sub

Private Sub Timer1_Timer()
MyUptime = FormatCount(GetTickCount, DaysHoursMinutesSeconds)
Label1.Caption = MyUptime
  With TrayIcon1
    .InfoTip = MyUptime
  End With
End Sub
Private Sub TrayIcon1_DblClick(ByVal Button As Integer, ByVal Shift As Integer)
  If Button = vbLeftButton Then
    Me.WindowState = vbNormal
    Me.Show
  End If
End Sub
Private Sub TrayIcon1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer)
  If Button = vbRightButton Then
    TrayIcon1.PopupMenu mnuTray, vbPopupMenuRightButton
  End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Unload Form1
End Sub
