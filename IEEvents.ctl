VERSION 5.00
Begin VB.UserControl IEEvents 
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1695
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   255
   ScaleWidth      =   1695
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IE Event Control"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1605
   End
End
Attribute VB_Name = "IEEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents oBrowserEvents As cBrowserEvents
Attribute oBrowserEvents.VB_VarHelpID = -1
Private m_Browsers As cBrowsers
Event BrowserNavigating(Browser As SHDocVw.InternetExplorer, ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
Event DocumentComplete(Browser As SHDocVw.InternetExplorer, pDisp As Object, URL As Variant)
Event DownLoadBegin(Browser As SHDocVw.InternetExplorer)
Event DownLoadComplete(Browser As SHDocVw.InternetExplorer)
Event FileDownload(Browser As SHDocVw.InternetExplorer, Cancel As Boolean)
Event NavigateComplete(Browser As SHDocVw.InternetExplorer, ByVal pDisp As Object, URL As Variant)
Event NavigateError(Browser As SHDocVw.InternetExplorer, ByVal pDisp As Object, URL As Variant, Frame As Variant, StatusCode As Variant, Cancel As Boolean)
Event NewWindow(Browser As SHDocVw.InternetExplorer, ppDisp As Object, Cancel As Boolean)
Event OnFullScreen(Browser As SHDocVw.InternetExplorer, ByVal FullScreen As Boolean)
Event ProgressChange(Browser As SHDocVw.InternetExplorer, ByVal Progress As Long, ByVal ProgressMax As Long)
Event TitleChange(Browser As SHDocVw.InternetExplorer, ByVal Text As String)
Event WindowClosing(Browser As SHDocVw.InternetExplorer, ByVal IsChildWindow As Boolean, Cancel As Boolean)
Event BrowserCreated(Browser As SHDocVw.InternetExplorer)
Event BrowserDestroyed()
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Property Get Enabled() As Boolean
    Enabled = Not oBrowserEvents Is Nothing
    PropertyChanged "Enabled"
End Property
'Destroy our browser collection and get new one
Public Sub Refresh()
    Set oBrowserEvents = Nothing
    Set oBrowserEvents = New cBrowserEvents
    oBrowserEvents.SetOwnerBrowserCollection m_Browsers
    oBrowserEvents.SyncCollection
End Sub
'Must set enabled to get events
Public Property Let Enabled(ByVal blnNewValue As Boolean)
    If blnNewValue Then
        If oBrowserEvents Is Nothing Then
            'Setting enabled when already enabled does nothing
            Set oBrowserEvents = New cBrowserEvents
            oBrowserEvents.SetOwnerBrowserCollection m_Browsers
            oBrowserEvents.SyncCollection
        End If
    Else
        Set oBrowserEvents = Nothing
    End If
    PropertyChanged "Enabled"
End Property

Private Sub oBrowserEvents_BrowserCreated(Browser As SHDocVw.InternetExplorer)
    RaiseEvent BrowserCreated(Browser)
End Sub

Private Sub oBrowserEvents_BrowserDestroyed()
    RaiseEvent BrowserDestroyed
End Sub

Private Sub oBrowserEvents_BrowserNavigating(Browser As SHDocVw.InternetExplorer, ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    RaiseEvent BrowserNavigating(Browser, pDisp, URL, Flags, TargetFrameName, PostData, Headers, Cancel)
End Sub

Private Sub oBrowserEvents_DocumentComplete(Browser As SHDocVw.InternetExplorer, pDisp As Object, URL As Variant)
    RaiseEvent DocumentComplete(Browser, pDisp, URL)
End Sub

Private Sub oBrowserEvents_DownLoadBegin(Browser As SHDocVw.InternetExplorer)
    RaiseEvent DownLoadBegin(Browser)
End Sub

Private Sub oBrowserEvents_DownLoadComplete(Browser As SHDocVw.InternetExplorer)
    RaiseEvent DownLoadComplete(Browser)
End Sub

Private Sub oBrowserEvents_FileDownload(Browser As SHDocVw.InternetExplorer, Cancel As Boolean)
    RaiseEvent FileDownload(Browser, Cancel)
End Sub

Private Sub oBrowserEvents_NavigateComplete(Browser As SHDocVw.InternetExplorer, ByVal pDisp As Object, URL As Variant)
    RaiseEvent NavigateComplete(Browser, pDisp, URL)
End Sub

Private Sub oBrowserEvents_NavigateError(Browser As SHDocVw.InternetExplorer, ByVal pDisp As Object, URL As Variant, Frame As Variant, StatusCode As Variant, Cancel As Boolean)
    RaiseEvent NavigateError(Browser, pDisp, URL, Frame, StatusCode, Cancel)
End Sub

Private Sub oBrowserEvents_NewWindow(Browser As SHDocVw.InternetExplorer, ppDisp As Object, Cancel As Boolean)
    RaiseEvent NewWindow(Browser, ppDisp, Cancel)
End Sub

Private Sub oBrowserEvents_OnFullScreen(Browser As SHDocVw.InternetExplorer, ByVal FullScreen As Boolean)
    RaiseEvent OnFullScreen(Browser, FullScreen)
End Sub

Private Sub oBrowserEvents_ProgressChange(Browser As SHDocVw.InternetExplorer, ByVal Progress As Long, ByVal ProgressMax As Long)
    RaiseEvent ProgressChange(Browser, Progress, ProgressMax)
End Sub

Private Sub oBrowserEvents_TitleChange(Browser As SHDocVw.InternetExplorer, ByVal Text As String)
    RaiseEvent TitleChange(Browser, Text)
End Sub

Private Sub oBrowserEvents_WindowClosing(Browser As SHDocVw.InternetExplorer, ByVal IsChildWindow As Boolean, Cancel As Boolean)
    RaiseEvent WindowClosing(Browser, IsChildWindow, Cancel)
End Sub


Private Sub UserControl_Initialize()
    Set m_Browsers = New cBrowsers
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = Label1.Width
    UserControl.Height = Label1.Height
End Sub

Private Sub UserControl_Terminate()
    Set oBrowserEvents = Nothing
    Set m_Browsers = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Enabled", oBrowserEvents Is Nothing, False
End Sub

Public Property Get Browsers() As cBrowsers
    Set Browsers = m_Browsers
End Property
