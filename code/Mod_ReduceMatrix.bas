Attribute VB_Name = "Mod_ReduceMatrix"
Option Explicit

'Reduced row echelon form for matrices
Function MatrixReduction(Mtx As Matrix_Struct) As String
    Dim bDone As Boolean
    Dim lPP As Long     'pivot point
    Dim lLastRow As Long
    Dim lRows As Long, lCols As Long
    Dim lSwpRow As Long, lSwpCol As Long
    Dim mtxMain As Matrix_Struct
    Dim mtxTmp As Matrix_Struct
    
    Dim tempPlace As Double
    Dim sResult As String
    
    Dim bStep1 As Boolean, bStep2 As Boolean, bStep3 As Boolean
    Dim bStep4 As Boolean, bStep5 As Boolean, bStep6 As Boolean
    
    mtxMain = Mtx
    mtxTmp = Mtx
    lPP = 0
    
    
    bDone = False
    While bDone = False
    
        bStep1 = False
        bStep2 = False
        bStep3 = False
        bStep4 = False
        bStep5 = False
        bStep6 = False
        lLastRow = 0
    
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '1.) find pivot point and swap
        For lRows = 0 To mtxMain.rows - 1
                
            If mtxMain.values(lRows, lPP) <> 0 Then
                If lRows = 0 Then
                    'do nothing & we are done
                    bStep1 = True
                Else
                    'swap this row with the zero row
                    For lSwpCol = 0 To mtxMain.columns - 1
                        mtxMain.values(lLastRow, lSwpCol) = mtxTmp.values(lRows, lSwpCol)
                        mtxMain.values(lRows, lSwpCol) = mtxTmp.values(lLastRow, lSwpCol)
                    Next lSwpCol
                    
                    mtxTmp = mtxMain
                    bStep1 = True
                    Exit For
                End If
            End If
            
        Next lRows
        
        sResult = sResult & PrintMatrix_Html(mtxMain, False, "", 1, "", "Value", 0)
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '2.) reduce the current row
        tempPlace = mtxMain.values(lLastRow, lPP)
        For lCols = 0 To mtxMain.columns - 1
            mtxMain.values(lLastRow, lCols) = mtxMain.values(lLastRow, lCols) / tempPlace
            mtxMain.valuesRef(lLastRow, lCols) = mtxMain.values(lLastRow, lCols) & "/" & tempPlace
        Next lCols
        
        sResult = sResult & "<br>" & PrintMatrix_Html(mtxMain, False, "", 1, "", "Value", 0)
        sResult = sResult & "<br>" & PrintMatrix_Html(mtxMain, False, "", 1, "", "ValueRef", 0)
       
        bDone = True
        lPP = lPP + 1
    Wend
    
    MatrixReduction = sResult
    
    Exit Function
    
    
ReductionErr:
    MsgBox Err.Number & "" & Err.Description
    Exit Function
    
End Function
