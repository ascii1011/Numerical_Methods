Attribute VB_Name = "Mod_InverseMatrix"
Option Explicit





Function GetInverseMatrix(A As Matrix_Struct) As Matrix_Struct
    Dim b As Matrix_Struct
    Dim c As Matrix_Struct
    Dim tmp As Matrix_Struct
    Dim RowsDeep As Long
    Dim lRows As Long, lCols As Long
    
    'output inverse matrix
    ReDim tmp.valuesRef(A.rows, A.columns)
    ReDim tmp.values(A.rows, A.columns)
    
    'inverse matrix
    ReDim b.valuesRef(A.rows, A.columns)
    b.rows = A.rows
    b.columns = A.columns
    b.valuesRef(0, 0) = "a"
    b.valuesRef(0, 1) = "b"
    b.valuesRef(1, 0) = "c"
    b.valuesRef(1, 1) = "d"
    
    'identity
    ReDim c.valuesRef(A.rows, A.columns)
    c.rows = A.rows
    c.columns = A.columns
    c.valuesRef(0, 0) = "1"
    c.valuesRef(0, 1) = "0"
    c.valuesRef(1, 0) = "1"
    c.valuesRef(1, 1) = "0"
    
    For lRows = 0 To A.rows
        For lCols = 0 To A.columns
                
            'For RowsDeep = 0 To A.rows - 1
                If tmp.valuesRef(lRows, lCols) = "" Then
                    tmp.valuesRef(lRows, lCols) = "(" & A.values(lRows, lCols) & " * " & b.valuesRef(lRows, lCols) & ")"
                Else
                    tmp.valuesRef(lRows, lCols) = tmp.valuesRef(lRows, lCols) & " + (" & A.values(lRows, lCols) & " * " & b.valuesRef(lRows, lCols) & ")"
                End If
            'Next RowsDeep
                
            tmp.values(lRows, lCols) = MathIt(tmp.valuesRef(lRows, lCols), 6)
                
        Next lCols
    Next lRows
    
    GetInverseMatrix = tmp
    
End Function

Function PrintMatrix_Html(Mtx As Matrix_Struct, html_syntax As Boolean, head As String, border As Long, tableName As String, Type_Value_ValueRef As String, tableWidth As Long) As String
    Dim lRows As Long, lCols As Long
    Dim html As String
    
    If html_syntax = True Then
        html = "<html><head><title>" & head & "</title>" & _
            "<style type=""text/css"">" & _
            ".bold {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 9px;font-weight: bold;}" & _
            ".ct {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 9px;font-weight: normal;}" & _
            "</style>" & _
            "</head><body>"
    End If
    
    html = html & "<table"
    
    If border > 0 Then
        html = html & " border=" & border
    End If
    
    If tableWidth > 0 Then
        html = html & " width=" & tableWidth
    End If
    
    html = html & ">"
    
    If tableName <> "" Then
        html = html & "<tr><td class=bold>" & tableName & "</td></tr>"
    End If
    
    For lRows = 0 To Mtx.rows - 1
        html = html & "<tr>"
        For lCols = 0 To Mtx.columns - 1
            If Type_Value_ValueRef = "Value" Then
                html = html & "<td class=ct align=center>" & Mtx.values(lRows, lCols) & "</td>"
            ElseIf Type_Value_ValueRef = "ValueRef" Then
                html = html & "<td class=ct align=center>" & Mtx.valuesRef(lRows, lCols) & "</td>"
            End If
        Next lCols
        html = html & "</tr>"
    Next lRows
    
    html = html & "</table>"
    
    If html_syntax = True Then
        html = html & "</body></html>"
    End If
        
    PrintMatrix_Html = html
End Function


Function ResetMatrixValues(Mtx As Matrix_Struct) As Matrix_Struct
    Dim lRows As Long, lCols As Long
    
    ReDim Mtx.values(Mtx.rows, Mtx.columns)
    
    For lRows = 0 To Mtx.rows
        For lCols = 0 To Mtx.columns
            Mtx.values(lRows, lCols) = 0
        Next lCols
    Next lRows
        
    ResetMatrixValues = Mtx
End Function
