VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmWebSites 
   Caption         =   "Web Sites"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   9060
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   60
      Width           =   8895
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6915
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   8895
      ExtentX         =   15690
      ExtentY         =   12197
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmWebSites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sSiteName As String
Dim sSiteAddress As String

Private Sub Combo1_Click()
    prcNavigate
End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        prcNavigate
    End If
End Sub

Private Sub Form_Load()
    Me.Width = 9180
    Me.Height = 8085
    
    Combo1.Text = "Integral>>http://www.hostsrv.com/webmab/app1/MSP/quickmath/02/pageGenerate?site=quickmath&s1=calculus&s2=integrate&s3=basic"
    Combo1.AddItem "Integral>>http://www.hostsrv.com/webmab/app1/MSP/quickmath/02/pageGenerate?site=quickmath&s1=calculus&s2=integrate&s3=basic"
    Combo1.AddItem "differentiate>>http://www.hostsrv.com/webmab/app1/MSP/quickmath/02/pageGenerate?site=mathcom&s1=calculus&s2=differentiate&s3=basic"
    Combo1.AddItem "Integral>>http://www.webmath.com/cgi-bin/gopoly.cgi?s=x%5E3%2B5x%5E2-2&wrt=x&action=integrate&back=integrate.html"
    Combo1.AddItem "differentiate:>>http://www.calc101.com/webMathematica/derivatives.jsp#topdoit"
    Combo1.AddItem "School-OnlineUtilities>>http://www.zweigmedia.com/ThirdEdSite/utilsindex.html"
    Combo1.AddItem "School-Integral>>http://www.zweigmedia.com/ThirdEdSite/integral/numint.html"
    
    prcNavigate
End Sub

Sub prcNavigate()
    Dim sTmp() As String
    
    sTmp = Split(Combo1.Text, ">>")
    
    If UBound(sTmp) = 1 Then
        sSiteName = sTmp(0)
        sSiteAddress = sTmp(1)
        
        Me.Caption = sSiteName
        WebBrowser1.Navigate sSiteAddress
    End If

End Sub

Private Sub Form_Resize()

On Error Resume Next

    WebBrowser1.Width = Me.Width - 350
    WebBrowser1.Height = Me.Height - 1400
End Sub
