VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmRungaKuta 
   Caption         =   "Runga Kuta / Multiple Regression"
   ClientHeight    =   10155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10155
   ScaleWidth      =   9930
   Begin VB.TextBox txtQuestion 
      Height          =   255
      Left            =   9060
      TabIndex        =   11
      Top             =   1860
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Print Preview"
      Height          =   315
      Left            =   8340
      TabIndex        =   10
      Top             =   2460
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   6180
      TabIndex        =   9
      Text            =   " a,b,a^-1*b"
      Top             =   2820
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   1155
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Top             =   1980
      Width           =   5835
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close Excel"
      Height          =   375
      Left            =   8220
      TabIndex        =   7
      Top             =   420
      Width           =   1515
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   1500
      Width           =   1635
   End
   Begin VB.TextBox Text2 
      Height          =   1035
      Left            =   5460
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   360
      Width           =   2595
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Continue -> SSE"
      Height          =   375
      Left            =   6420
      TabIndex        =   2
      Top             =   1500
      Width           =   1635
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6795
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   9675
      ExtentX         =   17066
      ExtentY         =   11986
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
   Begin VB.TextBox Text1 
      Height          =   1035
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   360
      Width           =   5115
   End
   Begin VB.Label Label9 
      Caption         =   "Question#"
      Height          =   195
      Left            =   8280
      TabIndex        =   12
      Top             =   1860
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "web reverse answer:"
      Height          =   195
      Left            =   5700
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Question:"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmRungaKuta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim grAry() As String

Dim EndCol As Long      'x
Dim EndRow As Long      'y

Dim Alphabet() As String


    Dim xlapp As Excel.Application
    Dim wbxl As Excel.Workbook
    Dim wsxl As Excel.Worksheet
    
    
Private Sub Command2_Click()
    Command2.Enabled = False
    Command1.Enabled = True
    Command3.Enabled = True
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Visible = True
    xlapp.WindowState = xlMaximized
    Set wbxl = xlapp.Workbooks.Add
    Set wsxl = xlapp.ActiveSheet
    
    prcParseMatrix
    prcCreateExcelDoc
End Sub

Private Sub Command3_Click()
    wbxl.Close
End Sub

Private Sub Command4_Click()

    WebBrowser1.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER
End Sub

Private Sub Form_Load()
    Command1.Enabled = False
    Command3.Enabled = False
    
    Me.Width = 10050
    Me.Height = 10665
    
    prcInit
    WebBrowser1.Navigate App.Path & "\matrix_calc.html"
    'WebBrowser1.Navigate "http://www.zweigmedia.com/ThirdEdSite/fancymatrixalg.html?input=a"
End Sub

Sub prcInit()
    ReDim Alphabet(27)
    Alphabet(0) = "A"
    Alphabet(1) = "B"
    Alphabet(2) = "C"
    Alphabet(3) = "D"
    Alphabet(4) = "E"
    Alphabet(5) = "F"
    Alphabet(6) = "G"
    Alphabet(7) = "H"
    Alphabet(8) = "I"
    Alphabet(9) = "J"
    Alphabet(10) = "K"
    Alphabet(11) = "L"
    Alphabet(12) = "M"
    Alphabet(13) = "N"
    Alphabet(14) = "O"
    Alphabet(15) = "P"
    Alphabet(16) = "Q"
    Alphabet(17) = "R"
    Alphabet(18) = "S"
    Alphabet(19) = "T"
    Alphabet(20) = "U"
    Alphabet(21) = "V"
    Alphabet(22) = "W"
    Alphabet(23) = "X"
    Alphabet(24) = "Y"
    Alphabet(25) = "Z"
    Text1.Text = "x0,x1,x2,x3,y;1,1,1,1,1;1,0,0,1,2;1,1,0,0,1;1,0,0,0,2;1,1,0,1,3"
End Sub


Sub prcCreateExcelDoc()

    Dim i As Long
    Dim tmp As String, sTxt As String
    
    'options
    'wsxl.Columns(intcol).AutoFit
    
    'wsxl.Range("A1:G1").Select  ' Replace with your  Range
    'With Selection
    '   .Font.Bold = True
    '   .HorizontalAlignment = xlCenter
    'End With
    
    'If you want to the Excel sheet data to Fit to one Page widh & Height
     'wsxl.PageSetup.FitToPagesWide = 1
     'wsxl.PageSetup.FitToPagesTall = 1
     
    With wsxl
        '.Cells(1, Alphabet(0)) = "X0"
        'For i = 2 To rk.endY
        '    .Cells(i, Alphabet(0)) = "1"
        'Next i
        
        'fill in array
        For rk.y = 0 To rk.endY
            For rk.X = 0 To rk.endX
                .Cells(rk.y + 1, Alphabet(rk.X)) = aryMtx_Tmp(rk.y, rk.X)
            Next rk.X
        Next rk.y
        
        'y-predicted
        .Cells(1, Alphabet(rk.X - 1)) = "Y-predicted"
        
        'goal statement
        tmp = "Goal is y = b0"
        For i = 1 To rk.X - 3
            tmp = tmp & "+b" & i & "X" & i
        Next i
        .Cells(rk.y + 2, Alphabet(0)) = tmp
        
        'template part 1
        For rk.y = 0 To rk.endY - 3
            For rk.X = 0 To rk.endX - 2
                .Cells(rk.y + rk.endY + 6, Alphabet(rk.X)) = "SX" & rk.y & "X" & rk.X
            Next rk.X
        Next rk.y
        'template part 2
        For rk.y = 0 To rk.endY - 3
            .Cells(rk.y + rk.endY + 6, Alphabet(rk.endX)) = "SX" & rk.y & "Y"
        Next rk.y
        
        Dim tmpLen As Long, place As Long, offset As Long
        'Calc template part 1
        ReDim aryMtx_main(rk.endX - 1, rk.endX - 1)
        For rk.y = 0 To rk.endY - 3
            tmpLen = rk.endX - 1
            place = 0
            For rk.X = 0 To rk.endX - 2
                If rk.X >= rk.y Then
                    tmp = funCalcMrtxSum
                    .Cells(rk.y + (2 * rk.endY) + 6, Alphabet(rk.X)) = tmp
                    aryMtx_main(rk.y, rk.X) = tmp
                    
                    If place > 0 Then
                        .Cells(rk.y + (2 * rk.endY) + 6 + place, Alphabet(rk.X - place)) = tmp
                        aryMtx_main(rk.X, rk.y) = tmp
                    End If
                    tmpLen = tmpLen - 1
                    place = place + 1
                End If
            Next rk.X
            offset = offset + 1
        Next rk.y
        
        
        
        sTxt = ""
        'build first part of webstring input
        sTxt = "a=["
        For rk.y = 0 To rk.endY - 3
            For rk.X = 0 To rk.endX - 2
                sTxt = sTxt & aryMtx_main(rk.y, rk.X) & vbTab
            Next rk.X
            sTxt = sTxt & vbCrLf
        Next rk.y
        sTxt = sTxt & "]"
        
        'Calc template part 2
        sTxt = sTxt & vbCrLf & vbCrLf & "b=["
        For rk.y = 0 To rk.endY - 3
            tmp = funCalcSmallMrtxSum
            .Cells(rk.y + (2 * rk.endY) + 6, Alphabet(rk.endX)) = tmp
            sTxt = sTxt & tmp & vbCrLf
        Next rk.y
        sTxt = sTxt & "]"
        Text3.Text = Trim(sTxt)
        
        
        
        
        '=MMULT(MINVERSE(A11:D14),F11:F14)
        'Inverse product
        Dim sFirst, sSecond, sThird, sFourth
        tmpLen = rk.X - 1
        
        sFirst = Alphabet(0) & (2 * rk.endY) + 6
        sSecond = Alphabet(rk.endX - 2) & (2 * rk.endY) + 6 + tmpLen
        sThird = Alphabet(rk.endX + 2) & (2 * rk.endY) + 6
        sFourth = Alphabet(rk.endX + 2) & (2 * rk.endY) + 6 + tmpLen
                
        'backcolor
        wsxl.Range(sFirst & ":" & sSecond).Select  ' Replace with your  Range
        With Selection
            With Selection.Interior
                .ColorIndex = 6
                .Pattern = xlSolid
            End With
           '.Font.Bold = True
           .HorizontalAlignment = xlCenter
        End With
        
        'backcolor
        wsxl.Range(sThird & ":" & sFourth).Select  ' Replace with your  Range
        With Selection
            With Selection.Interior
                .ColorIndex = 7
                .Pattern = xlSolid
            End With
           '.Font.Bold = True
           .HorizontalAlignment = xlCenter
        End With
        
        sThird = Alphabet(rk.endX) & (2 * rk.endY) + 6
        sFourth = Alphabet(rk.endX) & (2 * rk.endY) + 6 + tmpLen
        'backcolor
        wsxl.Range(sThird & ":" & sFourth).Select  ' Replace with your  Range
        With Selection
            With Selection.Interior
                .ColorIndex = 36
                .Pattern = xlSolid
            End With
           '.Font.Bold = True
           .HorizontalAlignment = xlCenter
        End With
        
        tmp = "\{=MMULT(MINVERSE(" & sFirst & ":" & sSecond & ")," & sThird & ":" & sFourth & ")}\"
        
        For rk.y = 0 To rk.endY - 3
            .Cells(rk.y + (2 * rk.endY) + 6, Alphabet(rk.endX + 3)) = tmp
        Next rk.y
        
        Dim sU1 As String, sU2 As String
        sU1 = Alphabet(0) & rk.endY + (2 * rk.endY) + 7
        sU2 = Alphabet(2) & rk.endY + (2 * rk.endY) + 8
        .Cells(rk.endY + (2 * rk.endY) + 7, 1) = "Christopher Harty"
        .Cells(rk.endY + (2 * rk.endY) + 8, 1) = "This is Question: " & Trim(txtQuestion.Text)
        wsxl.Range(sU1 & ":" & sU2).Select  ' Replace with your  Range
        With Selection
            With Selection.Interior
                .ColorIndex = 4
                .Pattern = xlSolid
            End With
           .Font.Bold = True
           '.HorizontalAlignment = xlCenter
        End With
        
        .Visible = True
        
     End With
     
     'wbxl.Close
End Sub

Private Sub Command1_Click()
    Command1.Enabled = False
    prcProcessWebAnswer
    prcDisplayWebAnswer
    prcFinishSSe
    Command2.Enabled = True
End Sub

Sub prcProcessWebAnswer()
    Dim sTmp
    Dim i As Integer
    
    sTmp = Split(Trim(Text2.Text), vbCrLf)
    If UBound(sTmp) > 0 Then
        ReDim aryMtx_Web(UBound(sTmp) + 1)
        For i = 0 To UBound(sTmp)
            aryMtx_Web(i) = Trim(Replace(sTmp(i), vbTab, ""))
        Next i
    End If
End Sub

Sub prcDisplayWebAnswer()
    Dim sLen As Long
    
    sLen = rk.endX - 1

    With wsxl
        
        'fill in array
        For rk.y = 0 To rk.endY - 3
            .Cells(rk.y + (2 * rk.endY) + 6, Alphabet(rk.endX + 2)) = aryMtx_Web(rk.y)
        Next rk.y
        
    End With
End Sub

Sub prcFinishSSe()
    Dim sLen As Long
    Dim sSum As Double
    Dim sTotal As Double
    sSum = 0
    Dim sTmp As String
    sLen = rk.endX - 1

    With wsxl
        
        'fill in array
        For rk.y = 0 To rk.endY - 2
            For rk.X = 0 To rk.endX - 2
                If rk.X = 0 Then
                    sSum = sSum + aryMtx_Web(rk.X)
                Else
                    sSum = sSum + aryMtx_Web(rk.X) * aryMtx_Tmp(rk.y + 1, rk.X)
                End If
            Next rk.X
            .Cells(2 + rk.y, Alphabet(rk.X + 1)) = sSum
            sSum = (aryMtx_Tmp(rk.y + 1, rk.endX - 1) - sSum) ^ 2
            .Cells(2 + rk.y, Alphabet(rk.X + 2)) = sSum
            sTotal = sTotal + sSum
            sSum = 0
        Next rk.y
        .Cells(2 + rk.y, Alphabet(rk.X + 2)) = sTotal
        .Cells(2 + rk.y, Alphabet(rk.X + 1)) = "SSE ="
        'backcolor
        sTmp = Alphabet(rk.X + 1) & 2 + rk.y & ":" & Alphabet(rk.X + 2) & 2 + rk.y
        .Range(sTmp).Select  ' Replace with your  Range
        With Selection
            With Selection.Interior
                .ColorIndex = 14
                .Pattern = xlSolid
            End With
           '.Font.Bold = True
           .HorizontalAlignment = xlCenter
        End With
    End With
End Sub


Function funCalcSmallMrtxSum() As Long
    Dim X As Long, y As Long
    Dim sum As Long
    
    sum = 0
    
    For y = 1 To rk.endY - 1
        sum = sum + aryMtx_Tmp(y, rk.y) * aryMtx_Tmp(y, rk.endX - 1)
    Next y
    
    funCalcSmallMrtxSum = sum
End Function

Function funCalcMrtxSum() As Long
    Dim X As Long, y As Long
    Dim sum As Long
    sum = 0
    For y = 1 To rk.endY - 1
        sum = sum + aryMtx_Tmp(y, rk.y) * aryMtx_Tmp(y, rk.X)
    Next y
    
    funCalcMrtxSum = sum
End Function

Sub prcParseMatrix()
    Dim sMatrx As String
    Dim X As Long, y As Long
    Dim MP1, MP2
    
'Public rk As RK_ST

'Public aryMtx_main() As String
'Public aryMtx_X() As String
'Public aryMtx_Y() As String
    
    sMatrx = Trim(Text1.Text)
    
    MP1 = Split(sMatrx, ";")
    
    rk.endY = UBound(MP1) + 1
    If UBound(MP1) > 0 Then
        For y = 0 To UBound(MP1)
            MP2 = Split(MP1(y), ",")
            
            If y = 0 Then
                rk.endX = UBound(MP2) + 1
                ReDim aryMtx_Tmp(rk.endY, rk.endX)
            End If
            
            For X = 0 To UBound(MP2)
                aryMtx_Tmp(y, X) = MP2(X)
            Next X
        Next y
    End If
    
End Sub

Private Sub txtFirstX_KeyUp(KeyCode As Integer, Shift As Integer)
    'prcRedrawGrid
End Sub

Private Sub txtFirstY_KeyUp(KeyCode As Integer, Shift As Integer)
    'prcRedrawGrid
End Sub
