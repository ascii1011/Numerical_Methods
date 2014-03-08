VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmMatrices_Reduction 
   Caption         =   "Form1"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7530
   ScaleWidth      =   7110
   Begin VB.CommandButton Command2 
      Caption         =   "Stage2"
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   1560
      Width           =   915
   End
   Begin VB.TextBox Text2 
      Height          =   1155
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   1380
      Width           =   4995
   End
   Begin VB.TextBox Text1 
      Height          =   1155
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   60
      Width           =   4995
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Stage1"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   240
      Width           =   915
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4695
      Left            =   60
      TabIndex        =   0
      Top             =   2700
      Width           =   6915
      ExtentX         =   12197
      ExtentY         =   8281
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
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmMatrices_Reduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sHtml As String

Private Sub Command1_Click()
    Dim AMtx As Matrix_Struct
    Dim sFileName As String
    Dim sHtml As String
    
    sFileName = "ReduceMatrix.html"
    
    sHtml = "<html><head><title>" & head & "</title>" & _
            "<style type=""text/css"">" & _
            ".bold {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 9px;font-weight: bold;}" & _
            ".ct {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 9px;font-weight: normal;}" & _
            "</style>" & _
            "</head><body>"
    
    AMtx = funParseString2Matrix(Trim(Text1.Text))
    sHtml = sHtml & PrintMatrix_Html(AMtx, False, "", 1, "", "Value", 0)
    
    'sHtml = sHtml & "<br>" & MatrixReduction(AMtx)
    'sHtml = sHtml & "<br>" & PrintMatrix_Html(AMtx, True, "", 1, "", "Value", 0)
        
        
    sHtml = sHtml & "</body></html>"
    
    prcFile App.Path, sFileName, sHtml
    
    WebBrowser1.Navigate App.Path & "\" & sFileName
    
End Sub

Private Sub Form_Load()
    
    Text1.Text = "0,2,3,-4,1;" & _
                "0,0,2,3,4;" & _
                "2,2,-5,2,4;" & _
                "2,0,-6,9,7;"
End Sub
