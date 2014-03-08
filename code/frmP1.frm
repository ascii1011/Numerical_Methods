VERSION 5.00
Begin VB.Form frmP1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "P1"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5940
   ScaleWidth      =   9060
   Begin VB.CheckBox Check1 
      Caption         =   "use sin(x)"
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   3120
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   5535
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   8595
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   3360
         TabIndex        =   14
         Top             =   3300
         Width           =   2115
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   2820
         TabIndex        =   11
         Top             =   3300
         Width           =   375
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   1800
         TabIndex        =   7
         Text            =   "5"
         Top             =   1320
         Width           =   315
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   720
         TabIndex        =   3
         Top             =   2580
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   2820
         TabIndex        =   2
         Top             =   2100
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Calc"
         Height          =   315
         Left            =   5700
         TabIndex        =   1
         Top             =   2580
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "k"
         Height          =   195
         Left            =   2940
         TabIndex        =   13
         Top             =   3060
         Width           =   135
      End
      Begin VB.Label Label7 
         Caption         =   "Using Chopping Method:"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   2820
         TabIndex        =   12
         Top             =   2820
         Width           =   2715
      End
      Begin VB.Label Label6 
         Caption         =   "Using Visual Basic rounding functions:"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   2820
         TabIndex        =   10
         Top             =   1860
         Width           =   2775
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmP1.frx":0000
         Height          =   855
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label4 
         Caption         =   "decimal places."
         Height          =   195
         Left            =   2160
         TabIndex        =   8
         Top             =   1380
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Round to "
         Height          =   195
         Left            =   1020
         TabIndex        =   6
         Top             =   1380
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Sin("
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2580
         Width           =   435
      End
      Begin VB.Label Label2 
         Caption         =   ")  ="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1980
         TabIndex        =   4
         Top             =   2580
         Width           =   435
      End
   End
End
Attribute VB_Name = "frmP1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dRndPlaces As Double
Dim dNum As Double

Private Sub Command1_Click()
    prcBegin
End Sub

Sub prcBegin()
    dRndPlaces = Trim(Text3.Text)
    
    If Not IsNumeric(Trim(Text1.Text)) Then Exit Sub
    
    If Check1.Value = 1 Then
        dNum = Sin(Trim(Text1.Text))
    Else
        dNum = Trim(Text1.Text)
    End If
    
    prcVBRounding
    prcChopping
End Sub

Sub prcVBRounding()
    Dim sum As Double
    Dim X As Double
    Dim iRndBy As Integer
    
On Error GoTo Err:
    
    iRndBy = Trim(Text3.Text)
    X = Trim(Text1.Text)
    sum = round(dNum, iRndBy)
    
    Text2.Text = sum
    
    Exit Sub
    
Err:
    MsgBox Err.Number & vbNewLine & Err.Description
End Sub


Sub prcChopping()
    Dim sValue As String
    Dim k As Double

    '=ROUND(LOG(ABS(A5))-0.5,0)+1-$B$1
    k = Abs(dNum)
    k = Log(k)
    k = Rnd(k - 0.5)
    k = k + 1 - dRndPlaces
    k = Rnd(Log(Abs(dNum)) - 0.5) + 1 - dRndPlaces
    Text4.Text = k
    
    '=10^B5*ROUND(10^(-B5)*ABS(A5),0)*A5/ABS(A5)
    sValue = (10 ^ k) * Rnd((10 ^ -k) * Abs(dNum)) * dNum / Abs(dNum)
    Text5.Text = sValue
End Sub


Function funShiftRight(sTheNum, k) As String
    Dim sValue As String
    Dim i As Integer
    
    If k = 0 Then
        funShiftRight = sTheNum
    Else
        Dim k2 As Integer
        Dim sNum As String
        k2 = 1
        sNum = k
        
        If sNum < 0 Then
            sNum = -sNum
        End If
        
        For i = 0 To sNum
            k2 = k2 * 10
        Next i
        
    End If
    
    If k > 0 Then
        sValue = k2 * sTheNum
    Else
        sValue = sTheNum / k2
    End If
        
    funShiftRight = sValue
        
        
End Function

Private Sub Text1_Change()
    prcBegin
End Sub
