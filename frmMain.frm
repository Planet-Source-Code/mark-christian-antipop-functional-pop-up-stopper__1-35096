VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Antipop"
   ClientHeight    =   1575
   ClientLeft      =   1110
   ClientTop       =   1545
   ClientWidth     =   4950
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrAnimate 
      Enabled         =   0   'False
      Interval        =   32
      Left            =   2280
      Top             =   600
   End
   Begin VB.Timer tmrCheck 
      Interval        =   32
      Left            =   1920
      Top             =   600
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdHide 
      Cancel          =   -1  'True
      Caption         =   "Hide"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin vbpAntipop.cSysTray SysTray 
      Left            =   0
      Top             =   0
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   0   'False
      TrayIcon        =   "frmMain.frx":10FA
      TrayTip         =   "Antipop - Pop-ups Allowed"
   End
   Begin VB.CommandButton cmdWeb 
      Caption         =   "Web"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdEmail 
      Caption         =   "Email"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin vbpAntipop.IEEvents IEEvents 
      Left            =   960
      Top             =   840
      _ExtentX        =   2831
      _ExtentY        =   344
      Enabled         =   -1  'True
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   240
      Top             =   300
      Width           =   480
   End
   Begin VB.Label lblAbout 
      Caption         =   "Includes IEEVENTS by Crazyman."
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   6
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label lblAbout 
      Caption         =   "© 2002 Mark Christian. Freeware."
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   5
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label lblAbout 
      AutoSize        =   -1  'True
      Caption         =   "Antipop v1.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   960
      TabIndex        =   4
      Top             =   120
      Width           =   1170
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim allowPops As Boolean
Dim animateStep As Integer
Dim CTRLDown As Boolean
Private Sub cmdClose_Click()
If MsgBox("Are you sure you want to stop Antipop?", vbQuestion + vbYesNo) = vbYes Then
  SysTray.InTray = False
  End
End If
End Sub

Private Sub cmdEmail_Click()
ShellExecute Me.hwnd, "OPEN", "mailto:mark.christian@bigfoot.com?subject=antipop", "", App.Path, 0
End Sub

Private Sub cmdHide_Click()
'This doesn't actually unload the form, because
'Form_Unload cancels the action and hides itself.
Unload Me
End Sub

Private Sub cmdWeb_Click()
ShellExecute Me.hwnd, "OPEN", "http://nexxus.dhs.org", "", App.Path, 0
End Sub


Private Sub Form_Load()
IEEvents.Enabled = True
allowPops = False
Set SysTray.TrayIcon = LoadResPicture("NOPOPS", vbResIcon)
SysTray.TrayTip = "Antipop - Pop-ups Not Allowed"
SysTray.InTray = True
imgIcon.Picture = Me.Icon
Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = True
Me.Hide
End Sub


Private Sub IEEvents_NewWindow(Browser As SHDocVw.InternetExplorer, ppDisp As Object, Cancel As Boolean)
If Not allowPops And Not CTRLDown Then
  'Show animation
  tmrAnimate.Enabled = True 'Enable animation timer
  animateStep = 1 'Set animation to first frame
  tmrAnimate_Timer 'Show first frame
  Cancel = True 'Cancel pop-up
End If
End Sub


Private Sub SysTray_MouseDblClick(Button As Integer, Id As Long)
frmMain.Show
End Sub

Private Sub SysTray_MouseDown(Button As Integer, Id As Long)
allowPops = (allowPops = False)
If allowPops Then
  SysTray.TrayTip = "Antipop - Pop-ups Allowed"
  Set SysTray.TrayIcon = LoadResPicture("POPS", vbResIcon)
Else
  SysTray.TrayTip = "Antipop - Pop-ups Not Allowed"
  Set SysTray.TrayIcon = LoadResPicture("NOPOPS", vbResIcon)
End If
End Sub



Private Sub tmrAnimate_Timer()
If animateStep = 13 Then 'End of animation
  animateStep = 1
  tmrAnimate.Enabled = False
  Set SysTray.TrayIcon = LoadResPicture("NOPOPS", vbResIcon)
End If

thisRes = Trim(Str(animateStep + 1))
If Len(thisRes) = 1 Then thisRes = "0" & thisRes
Set SysTray.TrayIcon = LoadResPicture(thisRes, vbResIcon)
animateStep = animateStep + 1
End Sub

Private Sub tmrCheck_Timer()
CTRLDown = (GetAsyncKeyState(vbKeyControl) <> 0)
End Sub


