VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmCleverPivoting 
   Caption         =   "Clever Pivoting"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7530
   ScaleWidth      =   7650
   Begin VB.TextBox txtQuestion 
      Height          =   255
      Left            =   2460
      TabIndex        =   6
      Top             =   1260
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print Preview"
      Height          =   315
      Left            =   3840
      TabIndex        =   5
      Top             =   1260
      Width           =   1155
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   960
      TabIndex        =   4
      Text            =   "5"
      Top             =   1260
      Width           =   435
   End
   Begin VB.TextBox Text1 
      Height          =   1035
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   120
      Width           =   7335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Compute"
      Height          =   315
      Left            =   6360
      TabIndex        =   0
      Top             =   1260
      Width           =   1095
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5715
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   7395
      ExtentX         =   13044
      ExtentY         =   10081
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
   Begin VB.Label Label9 
      Caption         =   "Question#"
      Height          =   195
      Left            =   1680
      TabIndex        =   7
      Top             =   1260
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Round To:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   795
   End
End
Attribute VB_Name = "frmCleverPivoting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WhichRowAreWeOn As Long
Dim whichColAreWeOn As Long
Dim iRnd As Long
Dim sFileName As String

Dim GMtx As Matrix_Struct

Private Sub Command1_Click()
    prcCompute
End Sub

Sub prcCompute2()
    Dim sSubHtml As String
    Dim lRow As Long, lCol As Long
    Dim dTmp1 As Double, dTmp2 As Double
    
    sSubHtml = "<html><head><title></title>" & _
            "<style type=""text/css"">" & _
            ".bold {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 9px;font-weight: bold;}" & _
            ".ct {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 9px;font-weight: normal;}" & _
            "</style>" & _
            "</head><body>"
    
    GMtx = funParseString2Matrix(Trim(Text1.Text))
    sSubHtml = sSubHtml & PrintMatrix_Html(GMtx, False, "", 1, "", "Value", 0)
    
        
    
    prcFile App.Path, sFileName, sSubHtml
    
    WebBrowser1.Navigate App.Path & "\" & sFileName
    
End Sub

Sub prcCompute()
    Dim sHtml As String
    
    WhichRowAreWeOn = 0
    whichColAreWeOn = 0
    
    iRnd = Trim(Text2.Text)
    
    
    sHtml = "<html><head><title></title>" & _
            "<style type=""text/css"">" & _
            ".bold {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 9px;font-weight: bold;}" & _
            ".ct {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 9px;font-weight: normal;}" & _
            "</style>" & _
            "</head><body>"
            
    sHtml = sHtml & "<table border1><tr><td colspan=2 class=bold>Christopher Harty</td><tr>" & _
                    "<tr><td>Question: " & Trim(txtQuestion.Text) & "</td></tr></table><br>"
    
    GMtx = funParseString2Matrix(Trim(Text1.Text))
    sHtml = sHtml & PrintMatrix_Html(GMtx, False, "", 1, "", "Value", 0)
    
    sHtml = sHtml & "<br>" & funCleverPivoting
            
    sHtml = sHtml & "</body></html>"
    
    prcFile App.Path, sFileName, sHtml
    
    WebBrowser1.Navigate App.Path & "\" & sFileName
End Sub

Function funCleverPivoting() As String
    Dim lRow As Long, lCol As Long
    Dim sSubHtml As String
    Dim tmpMtx As Matrix_Struct
    
On Error GoTo Error:
    
    tmpMtx = GMtx
    
    
    sSubHtml = "<html><head><title></title>" & _
            "<style type=""text/css"">" & _
            ".bold {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 9px;font-weight: bold;}" & _
            ".ct {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 9px;font-weight: normal;}" & _
            "</style>" & _
            "</head><body>"

    '''''''''''''''''''''''''''''''''
    'First Calculate each row-maximum
    Dim dMax As Double, dTmpMax() As Double
    
    ReDim dTmpMax(GMtx.columns)
    
    
    sSubHtml = sSubHtml & PrintMatrix_Html(GMtx, False, "", 1, "", "Value", 0)
    
    'sSubHtml = sSubHtml & "<table border=1>"
    For lRow = 0 To GMtx.rows - 1
        For lCol = 0 To GMtx.columns - 1
            'sSubHtml = sSubHtml & "<tr><td class=ct>Current:" & dTmpMax(lRow) & " vs " & Abs(GMtx.values(lRow, lCol))
            If dTmpMax(lRow) < Abs(GMtx.values(lRow, lCol)) Then dTmpMax(lRow) = Abs(GMtx.values(lRow, lCol))
            'sSubHtml = sSubHtml & " = " & dTmpMax(lRow) & "</td></tr>"
        Next lCol
        If dMax < dTmpMax(lRow) Then dMax = dTmpMax(lRow)
    Next lRow
    'sSubHtml = sSubHtml & "</table>"
    
    sSubHtml = sSubHtml & "<table border=1>"
    For lRow = 0 To GMtx.rows - 1
        sSubHtml = sSubHtml & "<tr><td class=bold>=" & dTmpMax(lRow) & "</td></tr>"
    Next lRow
    sSubHtml = sSubHtml & "</table>"
    
    sSubHtml = sSubHtml & "<table border=1><tr><td class=bold>Max = " & dMax & "</td></tr></table>"
    
    If dMax = 0 Then
        sSubHtml = "<table border=1><tr><td class=bold>Singular Matrix = Done!</td></tr></table>"
        Exit Function
    End If
    'print result
    'end of first
    '''''''''''''
    
    
    prcFile App.Path, sFileName, sSubHtml
    WebBrowser1.Navigate App.Path & "\" & sFileName
    
    
    
    '''''''''''''
    'Now we begin to pivot
    Dim MaxPivot As Double
    Dim Pivot_Row As Long
    Dim Pivot_Col As Long
    Dim Swap As Long
    Dim Value As Double, dTmp As Double
    Dim tmpCol As Long
    
    'reduction part
    Dim lReductRow As Long, lReductcol As Long
    Dim dReduced As Double, dTmpColVal As Double
    
    For lRow = 0 To GMtx.rows - 1
        'first, look for a candidate to use as pivot in the column below the diagonal
        MaxPivot = round(Abs(GMtx.values(lRow, lRow) / dTmpMax(lRow)), iRnd)
        sSubHtml = sSubHtml & "<br><span class=ct>MaxPivot: " & MaxPivot & " = " & Abs(GMtx.values(lRow, lRow)) & " / " & dTmpMax(lRow) & "</span>"
        
        Pivot_Row = lRow
        Swap = 0
        
        For lCol = lRow + 1 To GMtx.rows - 1
            If GMtx.values(lCol, lRow) <> 0 And dTmpMax(lCol) <> 0 Then
                'sSubHtml = sSubHtml & "<br>zero was found... skipping this row item for this pivot."
                'Exit For
            'End If
                Value = round(Abs(GMtx.values(lCol, lRow) / dTmpMax(lCol)), iRnd)
                sSubHtml = sSubHtml & "<br><span class=ct>value: " & Value & " = " & Abs(GMtx.values(lCol, lCol)) & " / " & dTmpMax(lCol) & "</span>"
                
                sSubHtml = sSubHtml & "<br><span class=ct>Checking if " & Value & " > " & MaxPivot & "</span>"
                
                
        'prcFile App.Path, sFileName, sSubHtml
        'WebBrowser1.Navigate App.Path & "\" & sFileName
                'Exit Function
                
                
                If Value > MaxPivot Then
                    MaxPivot = Value
                    Pivot_Row = lCol
                    Swap = 1
                    sSubHtml = sSubHtml & " = True"
                                    
                
                    If MaxPivot = 0 Then
                        'announce singular matrix
                        sSubHtml = sSubHtml & " Row " & lRow & " being bypassed because of Zero"
                        Exit For
                    End If
                    
                    If Swap = 1 Then
                    
                        sSubHtml = sSubHtml & "<br><span class=ct>Swapping Row " & lRow & " for " & Pivot_Row & "</span>"
                        For tmpCol = 0 To GMtx.columns - 1
                            GMtx.values(lRow, tmpCol) = tmpMtx.values(Pivot_Row, tmpCol)
                            GMtx.values(Pivot_Row, tmpCol) = tmpMtx.values(lRow, tmpCol)
                        Next tmpCol
                        
                        dTmp = dTmpMax(lRow)
                        dTmpMax(lRow) = dTmpMax(Pivot_Row)
                        dTmpMax(Pivot_Row) = dTmp
                        
                        tmpMtx = GMtx
                        
                        sSubHtml = sSubHtml & "<br>" & PrintMatrix_Html(GMtx, False, "", 1, "", "Value", 0)
                                            
        'prcFile App.Path, sFileName, sSubHtml
        'WebBrowser1.Navigate App.Path & "\" & sFileName
                                
                    
                    
                    End If
                    
                    
                Else
                    sSubHtml = sSubHtml & " = False"
                End If
            
            End If
            
        Next lCol
        
                        'now reduce here
                        'Dim lReductRow As Long, lReductcol As Long
                        'Dim dReduced As Double, dTmpColVal As Double
                        For lReductRow = lRow + 1 To GMtx.rows - 1
                            If tmpMtx.values(lReductRow, lRow) <> 0 Then
                            
                                dTmpColVal = tmpMtx.values(lReductRow, lRow)
                                For lReductcol = lRow To GMtx.columns - 1
                                
                                    dReduced = GMtx.values(lReductRow, lReductcol) - ((dTmpColVal / GMtx.values(lRow, lRow)) * GMtx.values(lRow, lReductcol))
                                    GMtx.values(lReductRow, lReductcol) = round(dReduced, iRnd)
                                    
                                Next lReductcol
                                
                            End If
                            
                        Next lReductRow
                        
                        sSubHtml = sSubHtml & "<br>before reduction<br>" & PrintMatrix_Html(tmpMtx, False, "", 1, "", "Value", 0)
                        sSubHtml = sSubHtml & "<br>After reduction<br>" & PrintMatrix_Html(GMtx, False, "", 1, "", "Value", 0)
                        
        'prcFile App.Path, sFileName, sSubHtml
        'WebBrowser1.Navigate App.Path & "\" & sFileName
        'Exit Function
                        tmpMtx = GMtx
        If lRow = GMtx.rows - 1 Then
            Exit For
        End If
    Next lRow
    Dim lTmp As Long
    Dim lBS As Long
    Dim lRow2 As Long
    Dim Xs() As Double
    Dim Xtmp() As Double
    Dim dTmp2 As Double
    Dim lNewPoint As Long
    Dim ltmpPoint As Long
    Dim lStep As Long
    
    ReDim Xs(GMtx.rows)
    ReDim Xtmp(GMtx.rows)
    
    lBS = 0
    Debug.Print "---------------------------------"
    Debug.Print "rows:" & GMtx.rows
    For lRow = GMtx.rows - 1 To 0 Step -1
    
        Xs(lRow) = GMtx.values(lRow, GMtx.columns - 1)
        Debug.Print vbTab & Xs(lRow)
        
        
        lStep = GMtx.rows - 1
        Debug.Print vbTab & "lbs:" & lBS
        lTmp = 0
        For lRow2 = GMtx.rows - 1 To 0 Step -1
        
            If lBS = lTmp Then Exit For
            Debug.Print vbTab & vbTab & " - (" & GMtx.values(lRow, lStep) & " * " & Xs(lStep) & ")"
            Xs(lRow) = Xs(lRow) - (GMtx.values(lRow, lStep) * Xs(lStep))
            lStep = lStep - 1
            lTmp = lTmp + 1
            
        Next lRow2
                
        
        Xs(lRow) = Xs(lRow) / GMtx.values(lRow, lRow)
        Debug.Print vbTab & " / " & GMtx.values(lRow, lRow)
        Debug.Print "X(" & lRow & ") = " & Xs(lRow)
        Debug.Print "********"
        
        lNewPoint = lRow
        lBS = lBS + 1
    Next lRow
    
    
    sSubHtml = sSubHtml & "<br><table border=1>" & _
                    "<tr><td colspan=3 class=bold>Back-Substitution</td></tr>"
    For lTmp = 0 To GMtx.rows - 1
        sSubHtml = sSubHtml & "<tr><td class=ct>X(" & lTmp & ")</td><td class=ct>" & Xs(lTmp) & "</td><td class=ct> Rounded:" & round(Xs(lTmp), iRnd) & "</td></tr>"
    Next lTmp
    sSubHtml = sSubHtml & "</table>"
    
    funCleverPivoting = sSubHtml
    
        'prcFile App.Path, sFileName, sSubHtml
        'WebBrowser1.Navigate App.Path & "\" & sFileName
    
    Exit Function
    
Error:
    sSubHtml = sSubHtml & "<br>Error: " & Err.Number & "<br>" & Err.Description & "<br>"
End Function

Private Sub Command2_Click()
    
    WebBrowser1.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER
End Sub

Private Sub Form_Load()
    Me.Width = 7770
    Me.Height = 8040
    Me.Show
    
    
    sFileName = "CleverPivoting.html"
        
    Text1.Text = "0.11,-43.2,1,0;" & _
                "43.2,-0.11,-2,2;" & _
                "2,1,1,4;"
    
    'Text1.Text = "0.0001,1,10.1,2;" & _
    '            "0,100000,20,1;" & _
    '            "1,50.01,0,7;"
End Sub


Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        prcCompute
    End If
End Sub
