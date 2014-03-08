VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmGradientDescent 
   Caption         =   "Gradient Descent - multiple variables"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8925
   ScaleWidth      =   11190
   Begin VB.TextBox txtMtrx2 
      Height          =   315
      Left            =   2580
      TabIndex        =   40
      Text            =   " a,b,a^-1*b"
      Top             =   2100
      Width           =   1215
   End
   Begin VB.TextBox txtMtrx1 
      Height          =   315
      Left            =   1260
      TabIndex        =   39
      Top             =   2100
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Matrix Calc"
      Height          =   315
      Left            =   120
      TabIndex        =   38
      Top             =   2100
      Width           =   975
   End
   Begin VB.TextBox txtUX 
      Height          =   315
      Left            =   1560
      TabIndex        =   35
      Top             =   1620
      Width           =   675
   End
   Begin VB.TextBox txtUY 
      Height          =   315
      Left            =   3660
      TabIndex        =   34
      Top             =   1620
      Width           =   675
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Get Diff"
      Height          =   315
      Left            =   120
      TabIndex        =   33
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtQuestion 
      Height          =   255
      Left            =   10380
      TabIndex        =   31
      Top             =   60
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Print Preview"
      Height          =   315
      Left            =   9900
      TabIndex        =   30
      Top             =   420
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Try Next"
      Height          =   255
      Left            =   8640
      TabIndex        =   29
      Top             =   2220
      Width           =   975
   End
   Begin VB.TextBox txtABResult 
      Height          =   285
      Left            =   7860
      TabIndex        =   26
      Text            =   " "
      Top             =   1500
      Width           =   1695
   End
   Begin VB.TextBox txtB 
      Height          =   285
      Left            =   7380
      TabIndex        =   24
      Text            =   " "
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Compute"
      Height          =   255
      Left            =   7020
      TabIndex        =   22
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtA 
      Height          =   285
      Left            =   7380
      TabIndex        =   20
      Text            =   " "
      Top             =   420
      Width           =   2175
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser2 
      Height          =   2895
      Left            =   120
      TabIndex        =   19
      Top             =   6000
      Width           =   10635
      ExtentX         =   18759
      ExtentY         =   5106
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
   Begin VB.TextBox txtRound 
      Height          =   315
      Left            =   4140
      TabIndex        =   17
      Top             =   420
      Width           =   1095
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3435
      Left            =   120
      TabIndex        =   16
      Top             =   2520
      Width           =   10635
      ExtentX         =   18759
      ExtentY         =   6059
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
   Begin VB.CommandButton Command1 
      Caption         =   "Compute All"
      Height          =   315
      Left            =   5460
      TabIndex        =   15
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtDY 
      Height          =   315
      Left            =   3360
      TabIndex        =   14
      Text            =   "(2 * Y)*cos(x^2+y^2)"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtDX 
      Height          =   315
      Left            =   1260
      TabIndex        =   11
      Text            =   "(2 * X)*cos(x^2+y^2)"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtError 
      Height          =   315
      Left            =   2220
      TabIndex        =   9
      Text            =   "0.0001"
      Top             =   420
      Width           =   1095
   End
   Begin VB.TextBox txtStep 
      Height          =   315
      Left            =   540
      TabIndex        =   7
      Text            =   "0.01"
      Top             =   420
      Width           =   1095
   End
   Begin VB.TextBox txtFirstY 
      Height          =   315
      Left            =   5940
      TabIndex        =   5
      Text            =   "10"
      Top             =   60
      Width           =   675
   End
   Begin VB.TextBox txtFirstX 
      Height          =   315
      Left            =   4740
      TabIndex        =   2
      Text            =   "10"
      Top             =   60
      Width           =   675
   End
   Begin VB.TextBox txtEquation 
      Height          =   315
      Left            =   840
      TabIndex        =   0
      Text            =   "sin(x^2 + y^2)"
      Top             =   60
      Width           =   2715
   End
   Begin VB.Label Label11 
      Caption         =   "X = "
      Height          =   195
      Left            =   1260
      TabIndex        =   37
      Top             =   1680
      Width           =   315
   End
   Begin VB.Label Label8 
      Caption         =   "Y = "
      Height          =   195
      Left            =   3360
      TabIndex        =   36
      Top             =   1680
      Width           =   315
   End
   Begin VB.Label Label2 
      Caption         =   "Question#"
      Height          =   195
      Left            =   9600
      TabIndex        =   32
      Top             =   60
      Width           =   735
   End
   Begin VB.Label Label66 
      Height          =   195
      Left            =   7080
      TabIndex        =   28
      Top             =   1860
      Width           =   3555
   End
   Begin VB.Label Label65 
      Caption         =   "-B / 2A ="
      Height          =   195
      Left            =   7080
      TabIndex        =   27
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label64 
      Caption         =   "B = "
      Height          =   195
      Left            =   7080
      TabIndex        =   25
      Top             =   780
      Width           =   315
   End
   Begin VB.Label Label63 
      Caption         =   "A = "
      Height          =   195
      Left            =   7080
      TabIndex        =   23
      Top             =   480
      Width           =   315
   End
   Begin VB.Label Label62 
      Caption         =   "Step 5 Enter A && B"
      Height          =   195
      Left            =   7080
      TabIndex        =   21
      Top             =   180
      Width           =   1635
   End
   Begin VB.Label Label61 
      Caption         =   "RND:"
      Height          =   195
      Left            =   3660
      TabIndex        =   18
      Top             =   480
      Width           =   435
   End
   Begin VB.Line Line4 
      X1              =   6720
      X2              =   6720
      Y1              =   900
      Y2              =   0
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6720
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Label Label10 
      Caption         =   "df/dy"
      Height          =   195
      Left            =   3960
      TabIndex        =   13
      Top             =   960
      Width           =   435
   End
   Begin VB.Label Label9 
      Caption         =   "df/dx"
      Height          =   195
      Left            =   2040
      TabIndex        =   12
      Top             =   960
      Width           =   435
   End
   Begin VB.Label Label7 
      Caption         =   "Error:"
      Height          =   195
      Left            =   1740
      TabIndex        =   10
      Top             =   480
      Width           =   435
   End
   Begin VB.Label Label6 
      Caption         =   "Step:"
      Height          =   195
      Left            =   60
      TabIndex        =   8
      Top             =   480
      Width           =   435
   End
   Begin VB.Label Label5 
      Caption         =   "Y = "
      Height          =   195
      Left            =   5640
      TabIndex        =   6
      Top             =   120
      Width           =   315
   End
   Begin VB.Label Label4 
      Caption         =   "X = "
      Height          =   195
      Left            =   4440
      TabIndex        =   4
      Top             =   120
      Width           =   315
   End
   Begin VB.Label Label3 
      Caption         =   "Start at:"
      Height          =   195
      Left            =   3660
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Equation:"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmGradientDescent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    prcStartComputing
End Sub

Sub prcStartComputing()

    prcPopulateStruct
    
    '1.) compute df/dx & df/dy
    'prcComputeDYDX
    
    '2.) x & y start values are set negative
    'GD.newX = -val(GD.startX)
    'GD.newY = -val(GD.startY)
    
    '3.) computer first 3 steps
    prcComputeFirst3Steps
    
    
    prcFile App.Path, "Gradient.html", funWeb
    
    WebBrowser1.Navigate App.Path & "\Gradient.html"
    WebBrowser2.Navigate App.Path & "\matrix_calc.html"
End Sub

Sub prcComputeFirst3Steps()
    txtRound.Text = "0"
    GD.round = 0
    
    '3.a) compute h = 0 (believe it is always zero)
    GD.tmpEQuation = GD.equation
    GD.tmpEQuation = funReplaceVarWithValue_ReturnEQ(GD.tmpEQuation, "X", GD.newX)
    GD.tmpEQuation = funReplaceVarWithValue_ReturnEQ(GD.tmpEQuation, "Y", GD.newY)
    'GD.round = Len(txtError.Text) - 2
    GD.round = 8
    txtRound.Text = GD.round
    GD.step1Val = round(val(CalcIt(GD.tmpEQuation)), GD.round)
    
End Sub

Function funReplaceVarWithValue_ReturnEQ(sEQ As String, sVar As String, dValue As Double) As String
    
    funReplaceVarWithValue_ReturnEQ = Replace(UCase(sEQ), UCase(sVar), dValue)

End Function

Sub prcPopulateStruct()
prcCalculateNewXY
    GD.equation = Trim(txtEquation.Text)
    GD.startX = Trim(txtFirstX.Text)
    GD.startY = Trim(txtFirstY.Text)
    GD.step = Trim(txtStep.Text)
    GD.Error = Trim(txtError.Text)
    GD.dxEQ = Trim(txtDX.Text)
    GD.dyEQ = Trim(txtDY.Text)
    
    GD.step1Val = 0
    GD.step1Val = GD.step
    GD.step1Val = GD.step * 2
    
End Sub

Function funWeb() As String
    Dim sTxt As String
    Dim tableWidth As Long
    
    tableWidth = 400
    
    sTxt = "<html><head><title></title>" & _
            "<style type=""text/css"">" & _
            ".bold {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 9px;font-weight: bold;}" & _
            ".ct {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 9px;font-weight: normal;}" & _
            "</style>" & _
            "</head><body>"
    
    sTxt = sTxt & "<table border1><tr><td colspan=2 class=bold>Christopher Harty</td><tr>" & _
                    "<tr><td>Question: " & Trim(txtQuestion.Text) & "</td></tr></table><br>"
    
    
    'general info
    sTxt = sTxt & "<table border=1>" & _
            "<tr><td colspan=2 align=center class=bold>Multi-variable Gradient Descent</td></tr>" & _
            "<tr><td class=ct width=40>Equation:</td><td align=center class=ct>" & GD.equation & "&nbsp;</td></tr>" & _
            "<tr><td class=ct>Start X:</td><td align=center class=ct>" & GD.startX & "&nbsp;</td></tr>" & _
            "<tr><td class=ct>Start Y:</td><td align=center class=ct>" & GD.startY & "&nbsp;</td></tr>" & _
            "<tr><td class=ct>Step:</td><td align=center class=ct>" & GD.step & "&nbsp;</td></tr>" & _
            "<tr><td class=ct>Error::</td><td align=center class=ct>" & GD.Error & "&nbsp;</td></tr>" & _
            "</table><br>"
            
    'Step 1
    sTxt = sTxt & "<table border=1 width=" & tableWidth & ">" & _
            "<tr><td class=bold>Step 1 (Calculate df/dx & df/dy)</td></tr>" & _
            "<tr><td class=ct><&nbsp;df/dx&nbsp;,&nbsp;df/dy&nbsp;></td></tr>" & _
            "<tr><td class=ct><&nbsp;" & GD.dxEQ & "&nbsp;,&nbsp;" & GD.dyEQ & "&nbsp;></td></tr>" & _
            "</table><br>"
            
    'Step 2
    sTxt = sTxt & "<table border=1 width=" & tableWidth & ">" & _
            "<tr><td class=bold>Step 2 (x and y signs switch - opposite of Gradient)</td></tr>" & _
            "<tr><td class=ct>-<&nbsp;X&nbsp;,&nbsp;&nbsp;Y&nbsp;></td></tr>" & _
            "<tr><td class=ct>&nbsp;<&nbsp;" & GD.newX & "&nbsp;,&nbsp;" & GD.newY & "&nbsp;></td></tr>" & _
            "</table><br>"
            
    'Step 3 Part 1
    GD.step1Val = 0
    GD.Stp3x = GD.startX
    GD.Stp3y = GD.startY
    GD.tmpX = GD.startX - GD.step1Val
    GD.tmpY = GD.startY - GD.step1Val
    GD.tmpEQuation = GD.equation
    GD.tmpEQuation = funReplaceVarWithValue_ReturnEQ(UCase(GD.tmpEQuation), "X", GD.tmpX)
    GD.tmpEQuation = funReplaceVarWithValue_ReturnEQ(UCase(GD.tmpEQuation), "Y", GD.tmpY)
    GD.Stp1Answer = -val(round(-CalcIt(GD.tmpEQuation), GD.round))
    sTxt = sTxt & "<table border=1 width=" & tableWidth & ">" & _
            "<tr><td align=left class=bold>Step 3A- § = h = " & GD.step1Val & "</td></tr>" & _
            "<tr><td align=left class=ct>§ = h = " & GD.step1Val & "</td></tr>" & _
            "<tr><td class=ct>f(" & GD.Stp3x & "," & GD.Stp3x & ") = " & GD.Stp1Answer & "</td></tr>" & _
            "</table><br>"
            
    'Step 3 Part 2
    GD.step2Val = round(GD.step, GD.round)
    GD.Stp3x = GD.startX & "-h"
    GD.Stp3y = GD.startY & "-h"
    GD.tmpX = GD.startX - GD.step2Val
    GD.tmpY = GD.startY - GD.step2Val
    GD.tmpEQuation = GD.equation
    GD.tmpEQuation = funReplaceVarWithValue_ReturnEQ(UCase(GD.tmpEQuation), "X", GD.tmpX)
    GD.tmpEQuation = funReplaceVarWithValue_ReturnEQ(UCase(GD.tmpEQuation), "Y", GD.tmpY)
    GD.Stp2Answer = -val(round(-CalcIt(GD.tmpEQuation), GD.round))
    sTxt = sTxt & "<table border=1 width=" & tableWidth & ">" & _
            "<tr><td align=left class=bold>Step 3B- § = h = " & GD.step2Val & "</td></tr>" & _
            "<tr><td align=left class=ct>§ = h = " & GD.step2Val & "</td></tr>" & _
            "<tr><td class=ct>f(" & GD.Stp3x & "," & GD.Stp3x & ") = " & GD.Stp2Answer & "</td></tr>" & _
            "</table><br>"
            
    'Step 3 Part 3
    GD.step3Val = round(GD.step * 2, GD.round)
    GD.Stp3x = GD.startX & "-2h"
    GD.Stp3y = GD.startY & "-2h"
    GD.tmpX = GD.startX - GD.step3Val
    GD.tmpY = GD.startY - GD.step3Val
    GD.tmpEQuation = GD.equation
    GD.tmpEQuation = funReplaceVarWithValue_ReturnEQ(UCase(GD.tmpEQuation), "X", GD.tmpX)
    GD.tmpEQuation = funReplaceVarWithValue_ReturnEQ(UCase(GD.tmpEQuation), "Y", GD.tmpY)
    GD.Stp3Answer = -val(round(-CalcIt(GD.tmpEQuation), GD.round))
    sTxt = sTxt & "<table border=1 width=" & tableWidth & ">" & _
            "<tr><td align=left class=bold>Step 3C- § = h = " & GD.step3Val & "</td></tr>" & _
            "<tr><td align=left class=ct>§ = h = " & GD.step3Val & "</td></tr>" & _
            "<tr><td class=ct>f(" & GD.Stp3x & "," & GD.Stp3x & ") = " & GD.Stp3Answer & "</td></tr>" & _
            "</table><br>"
            
    'Mapping
    sTxt = sTxt & "<table border=1>" & _
            "<tr><td colspan=3 align=left class=bold>Mapping of points</td></tr>" & _
            "  <tr> " & _
            "     <td class=ct align=center>§</td> " & _
            "     <td class=ct align=center><i>f</i></td> " & _
            "     <td class=ct>&nbsp;</td> " & _
            "  </tr>" & _
            "  <tr> " & _
            "     <td class=ct align=center>" & GD.step1Val & "&nbsp;</td> " & _
            "     <td class=ct align=center>" & GD.Stp1Answer & "&nbsp;</td> " & _
            "     <td class=bold> = c (Y intercept)</td> " & _
            "  </tr>" & _
            "  <tr> " & _
            "     <td class=ct align=center>" & GD.step2Val & "&nbsp;</td> " & _
            "     <td class=ct align=center>" & GD.Stp2Answer & "&nbsp;</td> " & _
            "     <td class=ct>&nbsp;</td> " & _
            "  </tr>" & _
            "  <tr> " & _
            "     <td class=ct align=center>" & GD.step3Val & "&nbsp;</td> " & _
            "     <td class=ct align=center>" & GD.Stp3Answer & "&nbsp;</td> " & _
            "     <td class=ct>&nbsp;</td> " & _
            "  </tr>" & _
            "</table><br>"
            
            GD.mtxA.rows = 2
            GD.mtxA.columns = 2
            ReDim GD.mtxA.values(GD.mtxA.rows, GD.mtxA.columns)
            GD.mtxA.values(0, 0) = GD.step2Val * GD.step2Val
            GD.mtxA.values(0, 1) = GD.step2Val
            GD.mtxA.values(1, 0) = GD.step3Val * GD.step3Val
            GD.mtxA.values(1, 1) = GD.step3Val
            
            GD.mtxB.rows = 2
            GD.mtxB.columns = 1
            ReDim GD.mtxB.values(GD.mtxB.rows, GD.mtxB.columns)
            GD.mtxB.values(0, 0) = GD.Stp2Answer - GD.Stp1Answer
            GD.mtxB.values(1, 0) = GD.Stp3Answer - GD.Stp1Answer
            
    sTxt = sTxt & "<table border=1>" & _
            "<tr><td align=left class=bold>Step 4 </td></tr>" & _
            "<tr><td align=left class=bold>Trying to find -b/2a for " & GD.equation & "</td></tr>" & _
            "<tr><td class=ct>a(" & GD.step2Val & ")^2 + b(" & GD.step2Val & ") + (" & GD.Stp1Answer & ") = " & GD.Stp2Answer & "</td></tr>" & _
            "<tr><td class=ct>a(" & GD.step3Val & ")^2 + b(" & GD.step3Val & ") + (" & GD.Stp1Answer & ") = " & GD.Stp3Answer & "</td></tr>" & _
            "<tr><td align=left class=bold>Results as follow:</td></tr>" & _
            "<tr><td class=ct>" & GD.mtxA.values(0, 0) & "a + " & GD.mtxA.values(0, 1) & "b = " & GD.mtxB.values(0, 0) & "</td></tr>" & _
            "<tr><td class=ct>" & GD.mtxA.values(1, 0) & "a + " & GD.mtxA.values(1, 1) & "b = " & GD.mtxB.values(1, 0) & "</td></tr>" & _
            "</table><br>"
            
    sTxt = sTxt & "<table border=1>" & _
            "<tr><td align=left class=bold>Step 5 </td></tr>" & _
            "<tr><td align=left class=bold>Use the following Matrices calculations to get A and B then enter them:</td></tr>" & _
            "<tr><td class=ct>[[" & GD.mtxA.values(0, 0) & "," & GD.mtxA.values(0, 1) & "] [" & GD.mtxA.values(1, 0) & "," & GD.mtxA.values(1, 1) & "]]^-1 * [[" & GD.mtxB.values(0, 0) & "][" & GD.mtxB.values(1, 0) & "]]</td></tr>" & _
            "</table><br>"
            
            
            txtMtrx1.Text = "a=[" & GD.mtxA.values(0, 0) & "," & GD.mtxA.values(0, 1) & vbNewLine & GD.mtxA.values(1, 0) & "," & GD.mtxA.values(1, 1) & "]" & vbNewLine & "b=[" & GD.mtxB.values(0, 0) & vbNewLine & GD.mtxB.values(1, 0) & "]"
            ''GD.mtxA_Inverse = GetInverseMatrix(GD.mtxA)
            'sTxt = sTxt & PrintMatrix_Html(GD.mtxA_Inverse, False, "", 1, "", "Value", 0)
            
            
    sTxt = sTxt & "</body></html>"
    
    funWeb = sTxt
    
    
End Function

Private Sub Command2_Click()
    txtABResult.Text = round((-val(txtB.Text) / (2 * val(txta.Text))), GD.round + 4)
    Label66.Caption = "if " & Trim(txtABResult.Text) & " > " & Trim(txtStep.Text) & "then press next"
End Sub

Private Sub Command3_Click()
    txtStep.Text = Trim(txtABResult.Text)
    prcStartComputing
End Sub

Private Sub Command4_Click()

    WebBrowser1.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER
End Sub

Private Sub Command5_Click()
    WebBrowser2.Navigate "http://www.hostsrv.com/webmab/app1/MSP/quickmath/02/pageGenerate?site=mathcom&s1=calculus&s2=differentiate&s3=basic"
End Sub


Sub prcCalculateNewXY()
    Dim sDx As String, sDy As String
    
    GD.startX = -Trim(txtFirstX.Text)
    GD.startY = -Trim(txtFirstY.Text)
    
    sDx = Replace(UCase(txtDX.Text), "Y", GD.startY)
    GD.newX = MathIt(sDx, GD.startX)
    txtUX.Text = GD.newX
    
    sDy = Replace(UCase(txtDY.Text), "Y", GD.startY)
    GD.newY = MathIt(sDy, GD.startX)
    txtUY.Text = GD.newY
End Sub
Private Sub Command7_Click()
    WebBrowser2.Navigate App.Path & "\matrix_calc.html"
End Sub

Private Sub Form_Load()
    
    Me.Width = 11205
    Me.Height = 9435
    
    WebBrowser2.Navigate "http://www.hostsrv.com/webmab/app1/MSP/quickmath/02/pageGenerate?site=mathcom&s1=calculus&s2=differentiate&s3=basic"
End Sub

