Attribute VB_Name = "mod_vars"
Option Explicit

Dim yyyyy As Double
Dim xywidth As Integer
Dim xmid As Integer
Dim ymid As Integer
Dim ywert As Double
Dim ERF As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'/ Runga Kuta Struct
Type RK_ST
    name As String
    startX As Long
    startY As Long
    endX As Long
    endY As Long
    X As Long
    y As Long
End Type

Public rk As RK_ST

Public aryMtx_main() As String
Public aryMtx_X() As String
Public aryMtx_Y() As String
Public aryMtx_Tmp() As String
Public aryMtx_Web() As String

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'/ Euler
Type Euler_ST
    function As String
    equation As String
    respect As String
    h As Double
    A As Double
    alpha As Double
    L As Double
    M As Double
    rows As Long
    round As Integer
    ColX As String
    ColY As String
    ColLip As String
    ColYNegLip As String
    ColYPosLip As String
    values() As Double
    valuesRef() As String
End Type

Public Euler As Euler_ST

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Type Matrix_Struct
    columns As Long
    rows As Long
    values() As Double
    valuesRef() As String
End Type

Public AMatrix As Matrix_Struct
Public BMatrix As Matrix_Struct
Public CMatrix As Matrix_Struct

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'/ Gradient Descent
Type GradDesc_ST
    round As Long
    equation As String
    tmpEQuation As String
    startX As Double
    startY As Double
    newX As Double
    newY As Double
    tmpX As Double
    tmpY As Double
    tmpH As Double
    step As Double
    Error As Double
    dxEQ As String
    dyEQ As String
    step1Val As Double
    step2Val As Double
    step3Val As Double
    Stp3x As String
    Stp3y As String
    Stp1Answer As Double
    Stp2Answer As Double
    Stp3Answer As Double
    mtxA As Matrix_Struct
    mtxB As Matrix_Struct
    mtxA_Inverse As Matrix_Struct
End Type

Public GD As GradDesc_ST

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function DoesFileExist(filename As String) As Boolean
    DoesFileExist = Dir(filename) <> ""
End Function

Function funParseString2Matrix(sMatrix As String) As Matrix_Struct
    Dim lRows As Long, lCols As Long
    Dim lRowCnt As Long, lColCnt As Long
    Dim X As Long, y As Long
    Dim mtxTmp As Matrix_Struct
    Dim MP1, MP2
        
    MP1 = Split(sMatrix, ";")
        
    lRowCnt = UBound(MP1)
    mtxTmp.rows = lRowCnt
        
    For lRows = 0 To lRowCnt - 1
        MP2 = Split(MP1(lRows), ",")
            
        If lRows = 0 Then
            lColCnt = UBound(MP2) + 1
            mtxTmp.columns = lColCnt
            ReDim mtxTmp.values(lRowCnt, lColCnt)
            ReDim mtxTmp.valuesRef(lRowCnt, lColCnt)
        End If
            
        For lCols = 0 To lColCnt - 1
            mtxTmp.values(lRows, lCols) = MP2(lCols)
        Next lCols
    Next lRows
    
    funParseString2Matrix = mtxTmp
    
End Function



Sub prcFile(sPath As String, sFileName As String, sPage As String)
    Dim f As Integer
    Dim strTemp As String

On Error GoTo errhandler
               
    strTemp = sPath & "\" & sFileName
    
    f = FreeFile
    Open strTemp For Output As #f
    Print #f, sPage
    Close #f
    Exit Sub
    
errhandler:
    Close #f
    MsgBox "An error occured by creating file."
    
End Sub



Function MathIt(sFunction As String, u As Double) As Double
On Error GoTo MathFehl
Dim eswert, A$, bb$, s, d$, c$
Dim b As Double
Dim X As Double
Dim sFunc As String

sFunc = sFunction

ERF = True
X = round(u, 4)
eswert = Str$(X)
'sFunction = Replace(UCase(sFunction), "EXP", "EZP")
A$ = Replace(sFunc, "EXP", "EZP")

While InStr(UCase$(A$), "X") > 0
  If InStr(UCase$(A$), "X") > 0 Then
     bb$ = eswert
     s = InStr(UCase$(A$), "X")
     's = Replace(UCase$(A$), "EZP", "EXP")
     d$ = Left$(A$, s - 1)
     If Len(A$) > s Then c$ = bb$ + Mid$(A$, s + 1, 100)
     A$ = d$ + c$
  End If
Wend

A$ = Replace(UCase$(A$), "EZP", "EXP")
yyyyy = round(val(CalcIt(A$)), 4)

MathIt = yyyyy
Exit Function

MathFehl:
ERF = -1
Resume Next
      
End Function


Function MathIt_ByVariable(sFunction As String, u As Double, sVariable As String) As Double
On Error GoTo MathFehl
Dim eswert, A$, bb$, s, d$, c$
Dim b As Double
Dim X As Double
Dim sFunc As String

sFunc = sFunction

ERF = True
X = round(u, 4)
eswert = Str$(X)
'sFunction = Replace(UCase(sFunction), "EXP", "EZP")
A$ = Replace(sFunc, "EXP", "EZP")

While InStr(UCase$(A$), sVariable) > 0
  If InStr(UCase$(A$), sVariable) > 0 Then
     bb$ = eswert
     s = InStr(UCase$(A$), sVariable)
     's = Replace(UCase$(A$), "EZP", "EXP")
     d$ = Left$(A$, s - 1)
     If Len(A$) > s Then c$ = bb$ + Mid$(A$, s + 1, 100)
     A$ = d$ + c$
  End If
Wend

A$ = Replace(UCase$(A$), "EZP", "EXP")
yyyyy = round(val(CalcIt(A$)), 4)

MathIt_ByVariable = yyyyy
Exit Function

MathFehl:
ERF = -1
Resume Next
      
End Function






Function MathIt_Options(u As Double, sEQ As String) As Double
On Error GoTo MathFehl
Dim eswert, A$, bb$, s, d$, c$
Dim b As Double
Dim X As Double

ERF = True
X = round(u, 4)
eswert = Str$(X)
A$ = sEQ

While InStr(UCase$(A$), "X") > 0
  If InStr(UCase$(A$), "X") > 0 Then
     bb$ = eswert
     s = InStr(UCase$(A$), "X")
     d$ = Left$(A$, s - 1)
     If Len(A$) > s Then c$ = bb$ + Mid$(A$, s + 1, 100)
     A$ = d$ + c$
  End If
Wend
yyyyy = round(val(CalcIt(A$)), 4)

MathIt_Options = yyyyy
Exit Function

MathFehl:
ERF = -1
Resume Next
      
End Function


Function CalcIt(CTerm As String) As String

    Dim RESTR$, KW$, RKA, PR, A$, i, MKT, PL, wert
    Dim Ready As Boolean
 Dim CalcPath(0 To 30) As String
 ConvertKlammer CTerm
 If CheckKlammer(CTerm) = True Then
    RESTR$ = MakeFullSyntax(CTerm)
    While Ready = False
        If InStr(RESTR$, "(") > 0 Then
            For i = 1 To Len(RESTR$)
              KW$ = Mid$(RESTR$, i, 1)
              If KW$ = "(" Then
                 RKA = RKA + 1
                 If RKA = CountKlammer(RESTR$) Then MKT = True: PL = i - 1
              End If
              If MKT = True Then
                 If KW$ <> "(" And KW$ <> ")" Then A$ = A$ + KW$
                 If KW$ = ")" Then PR = i + 1: Exit For
              End If
            Next i
            wert = val(CALC(A$))

            If PL = 0 Then RESTR$ = Str$(wert) + Mid$(RESTR, PR, 32000)
            If PL > 0 Then RESTR$ = Mid$(RESTR$, 1, PL) + Str$(wert) + Mid$(RESTR, PR, 32000)
            RESTR$ = KurzTerm(RESTR$)
            A$ = ""
            RKA = 0
            MKT = False
            If InStr(RESTR$, "(") < 1 Then Ready = True
        Else
           Ready = True
        End If
    Wend
    CalcIt = CALC(RESTR$)
 End If

End Function


Function CheckKlammer(STRG As String) As Boolean
    Dim A$, KW$, RKA, EKA, GKA, i
    A$ = STRG
    For i = 1 To Len(STRG)
      KW$ = Mid$(STRG, i, 1)
      Select Case KW$
        Case "("
           RKA = RKA + 1
        Case "["
           EKA = EKA + 1
        Case "{"
           GKA = GKA + 1
        Case ")"
           RKA = RKA - 1
        Case "]"
           EKA = EKA - 1
        Case "}"
           GKA = GKA - 1
      End Select
    Next i
    If RKA = 0 And EKA = 0 And GKA = 0 Then CheckKlammer = True
End Function



Function ConvertKlammer(STRG As String) As String
    Dim i, KW$
  For i = 1 To Len(STRG)
      KW$ = Mid$(STRG, i, 1)
      If KW$ = "[" Then Mid$(STRG, i, 1) = "("
      If KW$ = "{" Then Mid$(STRG, i, 1) = "("
      If KW$ = "]" Then Mid$(STRG, i, 1) = ")"
      If KW$ = "}" Then Mid$(STRG, i, 1) = ")"
  Next i
  ConvertKlammer = STRG
End Function
Function CountKlammer(Term As String) As Integer
    Dim i, KW$, RKA
    For i = 1 To Len(Term)
      KW$ = Mid$(Term, i, 1)
      Select Case KW$
        Case "("
           RKA = RKA + 1
      End Select
    Next i
    CountKlammer = RKA
End Function

Function MakeFullSyntax(sFunction As String) As String
    Dim i As Long
    Dim KW$, KWA$, A$, s$
    Dim found As Boolean
    
    For i = 1 To Len(sFunction)
      KW$ = Mid$(sFunction, i, 1)
      Select Case KW$
        Case "²"
           A$ = A$ + "^2"
           KWA$ = 2
        Case "³"
           A$ = A$ + "^3"
           KWA$ = 3
        Case "("
           If KWA$ > Chr$(47) And KWA$ < Chr$(58) Then
              A$ = A$ + "*("
              KWA$ = "("
           Else
              A$ = A$ + "("
              KWA$ = "("
           End If
        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
           If KWA$ = ")" Then
              A$ = A$ + "*" + KW$
              KWA$ = KW$
           Else
              A$ = A$ + KW$
              KWA$ = KW$
           End If
        Case Else
           A$ = A$ + KW$
           KWA$ = KW$
      End Select
    Next i
    A$ = UCase$(A$)
    While Not found
       s = InStr(UCase$(A$), "SIN")
       If s > 0 Then
          A$ = Left$(A$, s) + Mid$(A$, s + 3, 32000)
       Else
           found = True
       End If
    Wend
    found = False
    While Not found
       s = InStr(UCase$(A$), "COS")
       If s > 0 Then
          A$ = Left$(A$, s) + Mid$(A$, s + 3, 32000)
       Else
           found = True
       End If
    Wend
    found = False
    While Not found
       s = InStr(UCase$(A$), "SQR")
       If s > 0 Then
          A$ = Mid$(A$, 1, s - 1) + "Q" + Mid$(A$, s + 3, 32000)
       Else
           found = True
       End If
    Wend
    found = False
    While Not found
       s = InStr(UCase$(A$), "ABS")
       If s > 0 Then
          A$ = Left$(A$, s) + Mid$(A$, s + 3, 32000)
       Else
           found = True
       End If
    Wend
    MakeFullSyntax = A$
End Function




Function CALC(Term As String)

    Dim u, ZE$, OK, s, i, Z1$, PL, Z2$, PR, wert, MK

For u = 0 To 8
  If u = 0 Then ZE$ = "^"
  If u = 1 Then ZE$ = "S"
  If u = 2 Then ZE$ = "C"
  If u = 3 Then ZE$ = "Q"
  If u = 4 Then ZE$ = "A"
  If u = 5 Then ZE$ = "*"
  If u = 6 Then ZE$ = "/"
  If u = 7 Then ZE$ = "+"
  If u = 8 Then ZE$ = "-"
  While InStr(Term, ZE$) > 0 And OK = 0
    If Left$(Term, 1) = "-" And ZE$ = "-" Then
       s = InStr(Mid$(Term, 2, 32000), ZE$) + 1
    Else
       s = InStr(Term, ZE$)
    End If

    If u <> 1 Then
        For i = s - 1 To 1 Step -1
           If Mid$(Term, i, 1) > Chr$(47) And Mid$(Term, i, 1) < Chr$(58) Or Mid$(Term, i, 1) = "." Then
              Z1$ = Mid$(Term, i, 1) + Z1$
           Else
              If Mid$(Term, i, 1) = "-" Then
                 If i - 1 > 0 Then
                    If Mid$(Term, i - 1, 1) = "(" Then Z1$ = "-" + Z1$: i = i - 1
                    If Mid$(Term, i - 1, 1) = "-" Then Z1$ = "-" + Z1$: i = i - 1
                    If Mid$(Term, i - 1, 1) = "+" Then Z1$ = "-" + Z1$: i = i - 1
                    If Mid$(Term, i - 1, 1) = "*" Then Z1$ = "-" + Z1$: i = i - 1
                    If Mid$(Term, i - 1, 1) = "/" Then Z1$ = "-" + Z1$: i = i - 1
                    If Mid$(Term, i - 1, 1) = "S" Then Z1$ = "-" + Z1$: i = i - 1
                    If Mid$(Term, i - 1, 1) = "C" Then Z1$ = "-" + Z1$: i = i - 1
                    If Mid$(Term, i - 1, 1) = "A" Then Z1$ = "-" + Z1$: i = i - 1
                    If Mid$(Term, i - 1, 1) = "Q" Then Z1$ = "-" + Z1$: i = i - 1
                    
                 Else
                    Z1$ = "-" + Z1$
                    i = i - 1
                 End If
              End If
              Exit For
           End If
        Next i
        PL = i
    Else
        PL = s - 1
    End If
    For i = s + 1 To Len(Term)
       
       If Mid$(Term, i, 1) > Chr$(47) And Mid$(Term, i, 1) < Chr$(58) Or Mid$(Term, i, 1) = "." Or Mid$(Term, i, 1) = "." Or (Mid$(Term, i, 1) = "-" And Z2$ = "") Then
              Z2$ = Z2$ + Mid$(Term, i, 1)
       Else
         Exit For
       End If
    Next i
    PR = i
    If Z1$ = "" And u = 5 Then OK = 1
    If OK <> 1 Then
        If u = 0 Then wert = val(Z1$) ^ val(Z2$)
        If u = 1 Then wert = Sin(val(Z2$))
        If u = 2 Then wert = Cos(val(Z2$))
        If u = 3 Then wert = Sqr(val(Abs(Z2$)))
        If u = 4 Then wert = Abs(val(Z2$))
        If u = 5 Then wert = val(Z1$) * val(Z2$)
        If u = 6 Then wert = val(Z1$) / val(Z2$)
        If u = 7 Then wert = val(Z1$) + val(Z2$)
        If u = 8 Then wert = val(Z1$) - val(Z2$)
        MK = 0

        If PL = 0 Then Term = Str$(wert) + Mid$(Term, PR, 32000): MK = 1
        If PL > 0 Then Term = Mid$(Term, 1, PL) + Str$(wert) + Mid$(Term, PR, 32000): MK = 1
        If MK = 0 Then Term = Str$(wert)

        If InStr(Term, "E") > 0 Then Term = Str$(round(val(Term), 4))
        Term = KurzTerm(Term)
        If CZTerm(Term) = True Then OK = 1
    End If
  Z1$ = ""
  Z2$ = ""
  Wend
Next u
CALC = Term
End Function
Function KurzTerm(STRG As String) As String
    Dim t, A$, b$
   For t = 1 To Len(STRG)
       A$ = Mid$(STRG, t, 1)
       If A$ <> " " Then b$ = b$ + A$
   Next t
   KurzTerm = b$
End Function

Function CZTerm(STRG As String) As Boolean
    Dim Start, t, A$
    
    If Left$(STRG, 1) = "+" Or Left$(STRG, 1) = "-" Then Start = 2 Else Start = 1
    For t = Start To Len(STRG)
        A$ = Mid$(STRG, t, 1)
        If A$ > Chr$(47) And A$ < Chr$(58) Or A$ = "." Then
           CZTerm = True
        Else
           CZTerm = False
           Exit For
        End If
    Next t
End Function

