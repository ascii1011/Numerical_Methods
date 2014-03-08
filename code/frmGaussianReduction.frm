VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmGaussianReduction 
   Caption         =   "Gaussian Reduction"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4935
   ScaleWidth      =   9615
   Begin VB.CommandButton Command2 
      Caption         =   "Print Preview"
      Height          =   315
      Left            =   4980
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Compute"
      Height          =   315
      Left            =   3540
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   1035
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   180
      Width           =   4515
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1095
      Left            =   4800
      TabIndex        =   0
      Top             =   120
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   1931
      _Version        =   393216
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3075
      Left            =   120
      TabIndex        =   2
      Top             =   1740
      Width           =   9015
      ExtentX         =   15901
      ExtentY         =   5424
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
Attribute VB_Name = "frmGaussianReduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim grAry() As String
Dim grAry_tmp() As String

Dim StartCol As Long    'x
Dim StartRow As Long    'y
Dim EndCol As Long      'x
Dim EndRow As Long      'y

Private Sub Command1_Click()
    prcCompute
End Sub

Sub prcCompute()

    Dim GMtx As Matrix_Struct
    Dim sFileName As String
    Dim sHtml As String
    
    sFileName = "GaussianReduction.html"
    
    sHtml = "<html><head><title></title>" & _
            "<style type=""text/css"">" & _
            ".bold {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 9px;font-weight: bold;}" & _
            ".ct {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 9px;font-weight: normal;}" & _
            "</style>" & _
            "</head><body>"
    
    GMtx = funParseString2Matrix(Trim(Text1.Text))
    sHtml = sHtml & PrintMatrix_Html(GMtx, False, "", 1, "", "Value", 0)
    
    'sHtml = sHtml & "<br>" & MatrixReduction(gmtx)
    'sHtml = sHtml & "<br>" & PrintMatrix_Html(gmtx, True, "", 1, "", "Value", 0)
        
        
    sHtml = sHtml & "</body></html>"
    
    prcFile App.Path, sFileName, sHtml
    
    WebBrowser1.Navigate App.Path & "\" & sFileName
End Sub

Private Sub Command2_Click()
    WebBrowser1.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER
End Sub

Private Sub Form_Load()
    Me.Show
    
    
    Text1.Text = "0.0001,0.01,-0.212;" & _
                "0.0004,0.02,-0.085;"
    
    EndCol = 3
    EndRow = 2
    
    If EndRow <> 0 Or EndCol <> 0 Then
    
        ReDim grAry(EndCol, EndRow)
        ReDim grAry_tmp(EndCol, EndRow)
        
        grAry(0, 0) = 0.0001
        grAry(0, 1) = 0.01
        grAry(0, 2) = -0.121
        grAry(1, 0) = 0.0004
        grAry(1, 1) = 0.02
        grAry(1, 2) = -0.085
        
        MSFlexGrid1.Cols = EndCol + 1
        MSFlexGrid1.rows = EndRow + 1
        
        prcInsertMatrices
        prcGaussianReduction
    End If
    
End Sub

Sub prcInsertMatrices()
    Dim X As Integer, y As Integer
    
    For y = 1 To EndRow
    
        For X = 1 To EndCol
        
            MSFlexGrid1.Col = X
            MSFlexGrid1.Row = y
            
            MSFlexGrid1.Text = grAry(y - 1, X - 1)
        
        Next X
    
    Next y
    
End Sub


Sub prcGaussianReduction()
    Dim iDown As Integer
    Dim iAcross As Integer
    Dim i As Integer, j As Integer
    Dim bFound As Boolean
    Dim iPivot As Integer
    Dim dDivideBy As Double
    
    
    'step 1 (find row with all non-zeros or without a zero in the first spot(increment))
    For iDown = 0 To EndRow - 1
    
        bFound = False
        
        'check for a row with no zeros... first one found will be the pivot
        For iAcross = 0 To EndCol - 1
            If grAry(iDown, iAcross) = 0 Then
                bFound = True
                iAcross = EndCol
            End If
        Next iAcross
                
        If bFound = True Then
            bFound = False
            'check for a row with with a zero in idown - 1
            For iAcross = 1 To EndCol - 1
                If grAry(iDown, iAcross - 1) = 0 Then
                    bFound = True
                    iAcross = EndCol
                End If
            Next iAcross
        End If
        
        If bFound = True Then
            For i = 0 To EndCol - 1
                grAry_tmp(0, i) = grAry(0, i)
                grAry(0, i) = grAry(iPivot, i)
                grAry(iPivot, i) = grAry_tmp(0, i)
            Next i
        End If
        
        If bFound = False Then
            'now do the math for the top row
            dDivideBy = grAry(0, 0)
            
            For i = 0 To EndCol - 1
                grAry(0, i) = grAry(0, i) / dDivideBy
            Next i
            prcInsertMatrices
        End If
    
    Next iDown
    
    
End Sub














