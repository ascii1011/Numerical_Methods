VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmEulers 
   Caption         =   "Eulers with Lip"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   8385
   Begin VB.TextBox txtQuestion 
      Height          =   255
      Left            =   5400
      TabIndex        =   20
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox txtRound 
      Height          =   255
      Left            =   3840
      TabIndex        =   18
      Text            =   "4"
      Top             =   840
      Width           =   675
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print Preview"
      Height          =   315
      Left            =   6180
      TabIndex        =   17
      Top             =   1020
      Width           =   1155
   End
   Begin VB.TextBox txtRespect 
      Height          =   255
      Left            =   5340
      TabIndex        =   15
      Top             =   540
      Width           =   675
   End
   Begin VB.TextBox txtAnswer 
      Height          =   255
      Left            =   4620
      TabIndex        =   14
      Top             =   180
      Width           =   2535
   End
   Begin VB.TextBox txtM 
      Height          =   255
      Left            =   3840
      TabIndex        =   12
      Top             =   540
      Width           =   675
   End
   Begin VB.TextBox txtL 
      Height          =   255
      Left            =   2340
      TabIndex        =   10
      Top             =   840
      Width           =   675
   End
   Begin VB.TextBox txtalpha 
      Height          =   255
      Left            =   2340
      TabIndex        =   8
      Top             =   540
      Width           =   675
   End
   Begin VB.TextBox txta 
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   840
      Width           =   675
   End
   Begin VB.TextBox txth 
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   540
      Width           =   675
   End
   Begin VB.TextBox txtFunction 
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   180
      Width           =   3675
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   7335
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   7995
      ExtentX         =   14102
      ExtentY         =   12938
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
      Caption         =   "Compute"
      Height          =   315
      Left            =   6180
      TabIndex        =   0
      Top             =   540
      Width           =   1155
   End
   Begin VB.Label Label9 
      Caption         =   "Question#"
      Height          =   195
      Left            =   4620
      TabIndex        =   21
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Round:"
      Height          =   195
      Left            =   3120
      TabIndex        =   19
      Top             =   840
      Width           =   675
   End
   Begin VB.Label Label7 
      Caption         =   "Respect:"
      Height          =   195
      Left            =   4620
      TabIndex        =   16
      Top             =   540
      Width           =   675
   End
   Begin VB.Label Label6 
      Caption         =   "M:"
      Height          =   195
      Left            =   3120
      TabIndex        =   13
      Top             =   540
      Width           =   675
   End
   Begin VB.Label Label5 
      Caption         =   "L:"
      Height          =   195
      Left            =   1620
      TabIndex        =   11
      Top             =   840
      Width           =   675
   End
   Begin VB.Label Label4 
      Caption         =   "alpha:"
      Height          =   195
      Left            =   1620
      TabIndex        =   9
      Top             =   540
      Width           =   675
   End
   Begin VB.Label Label3 
      Caption         =   "a:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   675
   End
   Begin VB.Label Label2 
      Caption         =   "h:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   540
      Width           =   675
   End
   Begin VB.Label Label1 
      Caption         =   "Function:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   180
      Width           =   675
   End
End
Attribute VB_Name = "frmEulers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Dim IT As Double

Private Sub Command1_Click()
    prcPopStruct
    prcCalculateYGrid
    prcCreateGridStruct
    
    Euler.equation = funFormatFunction(Euler.function)
    'txtAnswer.Text = MathIt(-10)
    
    prcCalcXColumn
    prcCalcYColumn
    prcCalcLipColumn
    prcCalcY_Minus_ErrColumn
    prcCalcY_Plus_ErrColumn
        
    prcFile App.Path, "Eulers.html", funWeb
    
    WebBrowser1.Navigate App.Path & "\Eulers.html"
End Sub

Function funWeb() As String
    Dim i As Long
    
On Error GoTo webErr:

    funWeb = "<html><body>" & _
        "<head><title></title>" & _
        "<style type=""text/css"">" & _
            ".bold {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 8px;font-weight: bold;}" & _
            ".ct {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 8px;font-weight: normal;}" & _
            "</style></head><body>"
    
    funWeb = funWeb & "<table border1><tr><td colspan=2 class=bold>Christopher Harty</td><tr>" & _
                    "<tr><td>Question: " & Trim(txtQuestion.Text) & "</td></tr></table><br>"
    
        'general variables
        funWeb = funWeb & "<table border=1>" & _
            "<tr><td class=bold colspan=2>General Values</td></tr>" & _
            "<tr><td class=bold>h</td><td class=bold>" & Euler.h & "</td></tr>" & _
            "<tr><td class=bold>A</td><td class=bold>" & Euler.A & "</td></tr>" & _
            "<tr><td class=bold>alpha</td><td class=bold>" & Euler.alpha & "</td></tr>" & _
            "<tr><td class=bold>M</td><td class=bold>" & Euler.M & "</td></tr>" & _
            "<tr><td class=bold>L</td><td class=bold>" & Euler.L & "</td></tr>" & _
            "<tr><td class=bold colspan=2>Columns</td></tr>" & _
            "<tr><td class=bold>X</td><td class=bold>" & Euler.ColX & "</td></tr>" & _
            "<tr><td class=bold>Y</td><td class=bold>" & Euler.ColY & "</td></tr>" & _
            "<tr><td class=bold>Lipschitz(err)</td><td class=bold>" & Euler.ColLip & "</td></tr>" & _
            "<tr><td class=bold>y-err</td><td class=bold>" & Euler.ColYNegLip & "</td></tr>" & _
            "<tr><td class=bold>y+err</td><td class=bold>" & Euler.ColYPosLip & "</td></tr></table>"
    
    
        'First part
        funWeb = funWeb & "<table border=1>" & _
            "<tr><td class=bold>X</td><td class=bold>Y</td><td class=bold>Lipschitz(err)</td><td class=bold>Y-err</td><td class=bold>Y+err</td></tr>"
        For i = 0 To Euler.rows
            funWeb = funWeb & "<tr><td class=ct>" & Euler.values(i, 0) & "</td><td class=ct>" & Euler.values(i, 1) & "</td><td class=ct>" & Euler.valuesRef(i, 2) & Euler.values(i, 2) & "</td><td class=ct>" & Euler.values(i, 3) & "</td><td class=ct>" & Euler.values(i, 4) & "</td></tr>"
        Next i
        funWeb = funWeb & "</table>"
            
    funWeb = funWeb & "</body></html>"
    
    Exit Function
    
webErr:
    MsgBox Err.Description
    Resume Next

End Function



Sub prcCalcXColumn()
    Dim i As Double
    
    Euler.values(0, 0) = Euler.A
    
    For i = 1 To Euler.rows
        IT = (Euler.h * i)
        Euler.values(i, 0) = round((Euler.h * i), Euler.round)
    Next i
End Sub

Sub prcCalcYColumn()
    Dim i As Double
    Dim sTempOriginalFunction As String
    Dim sTempFunction As String
            
    sTempOriginalFunction = Euler.function
    sTempFunction = Euler.function
    Euler.ColY = "(Y<i><u>i-1</u></i>+h)*(" & sTempFunction & ")"
    
    Euler.values(0, 1) = Euler.alpha
    
    For i = 1 To Euler.rows
        sTempFunction = sTempOriginalFunction
        sTempFunction = Replace(UCase(sTempFunction), "X", Euler.values(i - 1, 0))
        sTempFunction = Replace(UCase(sTempFunction), "Y", Euler.values(i - 1, 1))
        sTempFunction = Euler.values(i - 1, 1) + Euler.h & "*(" & sTempFunction & ")"
        Euler.function = sTempFunction
        Euler.values(i, 1) = round(MathIt(Euler.function, Euler.values(i - 1, 1)), Euler.round)
    Next i
End Sub


Sub prcCalcLipColumn()
    Dim i As Double
    Dim sLip1 As String
    Dim sLip2 As String
    Dim sUseFunction As String
    
    sLip1 = (Euler.h * Euler.M) / (2 * Euler.L)
    
    sUseFunction = "((h*M)/(2*L))*(EXP((L*(x-a)))-1)"
    
    sUseFunction = Replace(sUseFunction, "h", Euler.h)
    sUseFunction = Replace(sUseFunction, "M", Euler.M)
    sUseFunction = Replace(sUseFunction, "L", Euler.L)
    sUseFunction = Replace(sUseFunction, "a", Euler.A)
       
    For i = 1 To Euler.rows
        Euler.valuesRef(i, 2) = Replace(sUseFunction, "x", Euler.values(i, 0)) & " = "
        Euler.values(i, 2) = round(sLip1 * (Exp(Euler.L * (Euler.values(i, 0) - Euler.A)) - 1), Euler.round)
        'Euler.values(i, 2) = MathIt(sUseFunction, Euler.values(i, 0))
    Next i
End Sub

Sub prcCalcY_Minus_ErrColumn()

    For i = 1 To Euler.rows
        Euler.valuesRef(i, 3) = "(" & Euler.values(i, 1) & " - " & Euler.values(i, 2) & " = "
        Euler.values(i, 3) = round(Euler.values(i, 1) - Euler.values(i, 2), Euler.round)
    Next i

End Sub

Sub prcCalcY_Plus_ErrColumn()

    For i = 1 To Euler.rows
        Euler.valuesRef(i, 4) = "(" & Euler.values(i, 1) & " + " & Euler.values(i, 2) & " = "
        Euler.values(i, 4) = round(Euler.values(i, 1) + Euler.values(i, 2), Euler.round)
    Next i

End Sub

Function funFormatFunction(sEQ As String) As String
    Euler.function = UCase(sEQ)
    Euler.respect = UCase(Euler.respect)
    sEQ = UCase(sEQ)
    
    'substitute
    funFormatFunction = Replace(sEQ, Euler.respect, Euler.alpha)
End Function

Sub prcCreateGridStruct()
    ReDim Euler.values(Euler.rows, 5)
    ReDim Euler.valuesRef(Euler.rows, 5)
End Sub

Sub prcCalculateYGrid()
    Euler.rows = (Euler.alpha - Euler.A) / Euler.h
End Sub

Sub prcPopStruct()
    Euler.function = Trim(txtFunction.Text)
    Euler.h = Trim(txth.Text)
    Euler.A = Trim(txta.Text)
    Euler.alpha = Trim(txtalpha.Text)
    Euler.L = Trim(txtL.Text)
    Euler.M = Trim(txtM.Text)
    Euler.respect = Trim(txtRespect.Text)
    Euler.round = Trim(txtRound.Text)
    Euler.ColX = "X<i><u>i-1</u></i>+h"
    Euler.ColLip = "((h*M)/(2*L))*(EXP((L*(x<i><u>i</u></i>-a)))-1)"
    Euler.ColYNegLip = "Y<i><u>i</u></i>-Lipschitz<i><u>i</u></i>"
    Euler.ColYPosLip = "Y<i><u>i</u></i>+Lipschitz<i><u>i</u></i>"
End Sub

Private Sub Command2_Click()
    WebBrowser1.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER
End Sub

Private Sub Form_Load()
    prcInitVars
End Sub

Sub prcInitVars()
    Me.Width = 8500
    Me.Height = 9510
    'txtFunction.Text = "sqr(5^2-(x)^2)"
    txtFunction.Text = "1/(1+(SIN(x+y))^2)"
    txtRespect.Text = "X"
    txth.Text = "0.1"
    txta.Text = "0"
    txtalpha.Text = "3"
    txtL.Text = "1"
    txtM.Text = "2"
    txtRound.Text = "4"
End Sub
