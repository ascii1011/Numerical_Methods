VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Modern Consumer - Collections Application"
   ClientHeight    =   8505
   ClientLeft      =   1815
   ClientTop       =   2040
   ClientWidth     =   13215
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMain.frx":37D2
   Begin VB.PictureBox picBackdrop 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      Height          =   315
      Left            =   0
      Picture         =   "frmMain.frx":3E54
      ScaleHeight     =   255
      ScaleWidth      =   13155
      TabIndex        =   2
      Top             =   420
      Visible         =   0   'False
      Width           =   13215
      Begin VB.PictureBox picOriginal 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   19440
         Left            =   1560
         Picture         =   "frmMain.frx":40D6
         ScaleHeight     =   19440
         ScaleWidth      =   25920
         TabIndex        =   4
         Top             =   600
         Width           =   25920
      End
      Begin VB.PictureBox picStretched 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   7260
         Left            =   2040
         ScaleHeight     =   7260
         ScaleWidth      =   4095
         TabIndex        =   3
         Top             =   300
         Width           =   4095
      End
   End
   Begin VB.Timer Timer2 
      Left            =   300
      Top             =   2520
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   300
      Top             =   1920
   End
   Begin MSComctlLib.ImageList ilToolbar 
      Left            =   240
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   25
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9085
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B837
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F019
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F8F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":68AF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":693D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":694E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":695F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":69751
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":698AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":699BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":69ACF
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6DB49
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6DC5B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":71CD5
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":71DE7
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":75E61
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":75F73
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":79FED
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A0FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A211
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7B493
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7F50D
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7F95F
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7FDB1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ilToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Collections"
            Object.ToolTipText     =   "Collections"
            Object.Tag             =   "mnuEdit_COLLECTIONS"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh"
            Object.Tag             =   "mnuEdit_REFRESH"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Reporting"
            Object.ToolTipText     =   "Reporting"
            Object.Tag             =   "mnu_Reporting"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Profile Manager"
            Object.ToolTipText     =   "Profile Manager"
            Object.Tag             =   "mnu_ProfileManager"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Search"
            Object.ToolTipText     =   "Search"
            Object.Tag             =   "mnuEdit_Search"
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Help"
            Object.ToolTipText     =   "Launch help file"
            Object.Tag             =   "mnuHELP_START"
         EndProperty
      EndProperty
      Enabled         =   0   'False
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   8250
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "5/18/2006"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "4:20 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "TSRTime"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7126
            MinWidth        =   7126
            Object.Tag             =   "mgr"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5362
            MinWidth        =   5362
            Object.Tag             =   "refresh"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Index           =   1
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Assignments"
      Index           =   2
      Begin VB.Menu mnuAssGp1_P1 
         Caption         =   "P1"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuTools_Nevilles 
         Caption         =   "Nevilles"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuAssign_RungaKuta 
         Caption         =   "Runga Kuta"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEulers 
         Caption         =   "Eulers"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuGradDesc 
         Caption         =   "Gradient Descent"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuAssign_CleverPivoting 
         Caption         =   "Clever Pivoting"
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Index           =   3
      Begin VB.Menu mnuTools_Websites 
         Caption         =   "WebSites"
         Shortcut        =   +{F1}
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnu_ClearWindows 
         Caption         =   "&Clear Windows"
      End
      Begin VB.Menu mnuArrangeWindows 
         Caption         =   "&Arrange Windows"
      End
      Begin VB.Menu mnuWindowspace2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuTileVertical 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowspace 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_about 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'versions
Dim sVersion_Current        As String
Dim sVersion_Current_Name   As String
Dim sVersion_Current_DateTime   As String

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long

Private Type RECT
    Left As Long
    top As Long
    Right As Long
    Bottom As Long
End Type

Private Sub MDIForm_Load()
    
    frmMain.StatusBar1.Panels.Item(5).Text = ""
    
    frmMain.Caption = "Numerical Methods - Main Application - Version " & App.Major & "." & App.Minor & "." & App.Revision
        
    'Get_User_Name
    'prcGetSysInfo
    'frmGraph.Show
    'frmGaussianReduction.Show
    'frmRungaKuta.Show
    'frmEulers.Show
    'frmGradientDescent.Show
    frmCleverPivoting.Show
                   
    Me.WindowState = 2
    'frmMain.StatusBar1.Panels.Item(3).Text = sUser
    
End Sub

Private Sub MDIForm_Resize()

On Error Resume Next
    
    Dim client_rect As RECT
    Dim client_hwnd As Long '''''

    picStretched.Move 0, 0, _
        ScaleWidth, ScaleHeight

    '''''''''''''''''''' Copy the original picture into picStretched.
    picStretched.PaintPicture _
        picOriginal.Picture, _
        0, 0, _
        picStretched.ScaleWidth, _
        picStretched.ScaleHeight, _
        0, 0, _
        picOriginal.ScaleWidth, _
        picOriginal.ScaleHeight
    
    ''''''''''''''''''''' Set the MDI form's picture.
    Picture = picStretched.Image '

    ''''''''''''''''''''''''' Invalidate the picture.
    client_hwnd = FindWindowEx(Me.hwnd, 0, "MDIClient", _
        vbNullChar)
    GetClientRect client_hwnd, client_rect
    InvalidateRect client_hwnd, client_rect, 1

End Sub


Private Sub mnu_about_Click()
    'frmAbout.Show
End Sub

Private Sub mnuAssGp1_P1_Click()
    frmP1.Show
End Sub

Private Sub mnuAssign_CleverPivoting_Click()
    frmCleverPivoting.Show
End Sub

Private Sub mnuAssign_RungaKuta_Click()
    frmRungaKuta.Show
End Sub

Private Sub mnuEulers_Click()
    frmEulers.Show
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Sub ErrorLog(sMsg As String, Err As ErrObject)
    Dim f
    
    f = FreeFile
        Open "C:\Output.txt" For Append As #f
        Print #f, Now & " Err = " & Err.Description & " " & Err.Number & " ; User:, Message:" & sMsg
        Close #f
        Screen.MousePointer = vbDefault
        
End Sub


Private Sub mnuArrangeWindows_Click()
   
    MousePointer = vbHourglass
    Arrange vbArrangeIcons
    MousePointer = vbDefault

End Sub

Private Sub mnuGradDesc_Click()
    frmGradientDescent.Show
End Sub

Private Sub mnuTileHorizontal_Click()
    MousePointer = vbHourglass
    Arrange vbTileHorizontal
    MousePointer = vbDefault
End Sub


Private Sub mnuTileVertical_Click()
   MousePointer = vbHourglass
    Arrange vbTileVertical
    MousePointer = vbDefault
End Sub

Private Sub mnuTools_Nevilles_Click()
    frmNevillesMethod.Show
End Sub

Private Sub mnuTools_Websites_Click()
    frmWebSites.Show
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Response

On Error GoTo errHandle
    
    Select Case (Button.Key)
        'Case ("Reporting"): prcShowFrmReporting
        'Case ("Tools"):   'prcOrderForm
        'Case ("Search"):  'prcVerifier
        'Case ("Help"):
    End Select
    Exit Sub

errHandle:
    'Debug.Print err.number
    Select Case (Err.Number)
        Case (91):
            Resume Next
        Case (364):
            Resume Next 'error closing qb connection while unloading
        Case (438):
            MsgBox ActiveForm.Caption & " does not support this operation", vbInformation, "SYSTEM"
        Case Else
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Login run time error")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
End Sub





