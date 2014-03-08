VERSION 5.00
Begin VB.Form frmNevillesMethod 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nevilles Method"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3270
   ScaleWidth      =   8115
   Begin VB.CommandButton Command5 
      Caption         =   "Create"
      Height          =   255
      Left            =   180
      TabIndex        =   15
      ToolTipText     =   "Create Tree Structure"
      Top             =   1500
      Width           =   915
   End
   Begin VB.ListBox List2 
      Height          =   1035
      Left            =   2520
      TabIndex        =   13
      Top             =   360
      Width           =   5415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   315
      Left            =   1800
      TabIndex        =   12
      Top             =   2820
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Calculate"
      Height          =   255
      Left            =   1500
      TabIndex        =   10
      Top             =   1500
      Width           =   915
   End
   Begin VB.TextBox Text9 
      Height          =   1455
      Left            =   2520
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Top             =   1680
      Width           =   5415
   End
   Begin VB.TextBox Text3 
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Text            =   "5"
      Top             =   1080
      Width           =   555
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-->"
      Height          =   255
      Left            =   900
      TabIndex        =   4
      Top             =   600
      Width           =   435
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<--"
      Height          =   255
      Left            =   900
      TabIndex        =   3
      ToolTipText     =   "Add to list"
      Top             =   180
      Width           =   435
   End
   Begin VB.TextBox Text2 
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   600
      Width           =   555
   End
   Begin VB.TextBox Text1 
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   180
      Width           =   555
   End
   Begin VB.ListBox List1 
      Height          =   1230
      ItemData        =   "frmNevillesMethod.frx":0000
      Left            =   180
      List            =   "frmNevillesMethod.frx":000D
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Output:"
      Height          =   195
      Left            =   2520
      TabIndex        =   14
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Tree"
      Height          =   195
      Left            =   2520
      TabIndex        =   11
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "N ="
      Height          =   195
      Left            =   1440
      TabIndex        =   8
      Top             =   1140
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Y ="
      Height          =   195
      Left            =   1440
      TabIndex        =   6
      Top             =   660
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "X ="
      Height          =   195
      Left            =   1440
      TabIndex        =   5
      Top             =   240
      Width           =   255
   End
End
Attribute VB_Name = "frmNevillesMethod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim pts() As String
Dim Max As Long
Dim lN As Long

Dim tree() As String



Private Sub Command1_Click()
    If IsNumeric(Trim(Text1.Text)) And IsNumeric(Trim(Text2.Text)) Then
        List1.AddItem Trim(Text1.Text) & "," & Trim(Text2.Text)
        Text1.Text = ""
        Text2.Text = ""
    End If
End Sub

Private Sub Command2_Click()
    Dim val
    
    val = Split(List1.List(List1.ListIndex), ",")
    
    Text1.Text = val(0)
    Text2.Text = val(1)
    
    List1.RemoveItem (List1.ListIndex)
End Sub

'Sub prcSaveTree(lvl As Long)
'    Dim j As Integer
'
'    For j = 0 To Max - 1
'        tree(lvl, j) = ""
'    Next j
'
'End Sub

'Function funNevillesCalc(Xo As Long, Yo As Long, X1 As Long, Y1 As Long) As Long
'    Dim lResult As Long
'
'    lResult = (((lN - Xo) * Y1) - ((lN - X1) * Yo)) / (X1 - Xo)
'
'    funNevillesCalc = lResult
'
'End Function

'Function funNevillesCalc(Xo As Long, Yo As Long, X1 As Long, Y1 As Long) As Long
'    Dim lResult As Long
'    'Dim sOutPut As String
'
'    lResult = (((lN - Xo) * Y1) - ((lN - X1) * Yo)) / (X1 - Xo)
'
'    'sOutPut = "[ ((" & lN & " - " & Xo & ") * " & Y1 & ") - ((" & lN & " - " & X1 & ") * " & Yo & ") ] / (" & X1 & " - " & Xo & ") = " & lResult
'    'List2.AddItem sOutPut
'
'    funNevillesCalc = lResult
'
'End Function

'Function funNevillesCalc_works(Xo As Long, Yo As Long, X1 As Long, Y1 As Long) As Long
'    Dim lResult As Long
'    Dim sOutPut As String
'
'    lResult = (((lN - Xo) * Y1) - ((lN - X1) * Yo)) / (X1 - Xo)
'
'    sOutPut = "[ ((" & lN & " - " & Xo & ") * " & Y1 & ") - ((" & lN & " - " & X1 & ") * " & Yo & ") ] / (" & X1 & " - " & Xo & ") = " & lResult
'    List2.AddItem sOutPut
'
'    funNevillesCalc = lResult
'
'End Function

Private Sub Command4_Click()
    Unload Me
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Command3_Click()
    prcInitNevillesVars
    prcNevillesProcess
End Sub

Sub prcInitNevillesVars()
    Dim i As Long, j As Long
    Dim pt_val As String
    Dim ary
    Dim sLine As String

    List2.Clear
    Text9.Text = ""
    lN = Trim(Text3.Text)

    Max = List1.ListCount
    
    ReDim tree(Max, Max, 2)
        
    sLine = ""
    For i = 0 To Max - 1
        pt_val = List1.List(i)
        ary = Split(pt_val, ",")
        
        tree(i, 0, 0) = ary(0)
        tree(i, 0, 1) = ary(1)
        sLine = sLine & tree(i, 0, 0) & "," & tree(i, 0, 1)
        
        For j = 1 To Max - 1
            If j > 0 Then
                tree(i, j, 0) = "-"
                tree(i, j, 1) = "-"
                sLine = sLine & vbTab & vbTab & tree(i, j, 0)
            End If
        Next j
        
        sLine = sLine & vbNewLine
        
    Next i
    
    'Text9.Text = sLine
    prcRefreshTree
End Sub

Sub prcNevillesProcess()
    Dim i As Long, j As Long
    'Dim x As Long, y As Long
    
    'x = Max - 1
    For i = 1 To Max - 1
        'y = Max - 1
        For j = 1 To i
            tree(i, j, 0) = funCalculate(i, j)
            'y = y - 1
        Next j
        'x = x - 1
    Next i
    
End Sub

Function funCalculate(i As Long, j As Long) As Long
    Dim lResult As Long
    Dim Po As Long
    Dim P1 As Long
    Dim Xo As Long
    Dim X1 As Long
    Dim sOutPut As String
    
On Error Resume Next
    
    If j = 1 Then
        Po = tree(i - 1, 0, 1)
        P1 = tree(i, 0, 1)
        Xo = tree(i - 1, 0, 0)
        X1 = tree(i, 0, 0)
    Else
        Po = tree(i - 1, j - 1, 0)
        P1 = tree(i, j - 1, 0)
        Xo = tree(i - j, j - j, 0)
        X1 = tree(i, j - j, 0)
    End If
    
    '   tree(i, j, 0)
    tree(i, j, 0) = (((lN - Xo) * P1) - ((lN - X1) * Po)) / (X1 - Xo)
    prcRefreshTree
    
    sOutPut = "[ ((" & lN & " - " & Xo & ") * " & P1 & ") - ((" & lN & " - " & X1 & ") * " & Po & ") ] / (" & X1 & " - " & Xo & ") = " & tree(i, j, 0)
    List2.AddItem sOutPut
    
    funCalculate = tree(i, j, 0)
    
End Function


Sub prcRefreshTree()
    Dim i As Long, j As Long
    Dim sLine As String
    
    For i = 0 To UBound(tree) - 1
        For j = 0 To UBound(tree) - 1
            If j <> 0 Then
                sLine = sLine & vbTab & vbTab
            End If
            sLine = sLine & tree(i, j, 0) & "," & tree(i, j, 1)
        Next j
        sLine = sLine & vbNewLine
    Next i
    
    Text9.Text = sLine
    
End Sub

Private Sub Command5_Click()
    prcInitNevillesVars
End Sub
