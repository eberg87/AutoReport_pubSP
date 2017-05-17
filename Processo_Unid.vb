Sub ProcessosUnidades()
    Dim NCol         As Integer
    Dim NRow         As Integer
    Dim i            As Integer
    Dim j            As Integer
    Dim k            As Integer
    Dim p            As Integer
    Dim AnoLen       As Integer
    Dim UnidLen      As Integer
    Dim MesesLen     As Integer
    Dim Unidades()   As Variant
    Dim AnosAux()    As Variant
    Dim Anos()       As Variant
    Dim Meses()      As Variant

    NCol = 1
    NRow = 2
    i = 0
    UnidLen = 0
    Unidades() = Array("")

    'laço para criar uma array com o nome de todas as unidades disponíveis
    Worksheets("Dados").Activate
    Do While Cells(NRow, NCol).Value <> ""
        Unidades(i) = Cells(NRow, NCol).Value
        If i = 0 Then
            i = i + 1
            UnidLen = i 'equação que irá definir o tamanho do vetor
            ReDim Preserve Unidades(0 To i) 'Array começa com 0 sempre
        ElseIf Unidades(i) <> Unidades(i - 1) Then
            i = i + 1
            UnidLen = i 'equação que irá definir o tamanho do vetor
            ReDim Preserve Unidades(0 To i) 'Array começa com 0 sempre
        End If
        NRow = NRow + 1
        'MsgBox NRow & ";" & NCol & ";" & i & ";" & Unidades(i - 1)
    Loop
    If Unidades(i) = Unidades(i - 1) Then
        i = i - 1
        UnidLen = i 'equação que irá definir o tamanho do vetor
        ReDim Preserve Unidades(0 To i)
    End If
    
    
    NCol = 2
    NRow = 2
    i = 0
    j = 0
    AnoLen = 0
    Anos() = Array("")
    AnosAux() = Array("")
    
    'laço para criar uma array com todos os anos disponíveis
    Worksheets("Dados").Activate
    Do While Cells(NRow, NCol).Value <> ""
        AnosAux(i) = Cells(NRow, NCol).Value
        i = i + 1
        AnoLen = i 'equação que irá definir o tamanho do vetor
        ReDim Preserve AnosAux(0 To i) 'Array começa com 0 sempre
        NRow = NRow + 1
    Loop
    
    'Laço para organizar anos em ordem crescente
    For i = 0 To AnoLen - 1
        For j = i + 1 To AnoLen
            If AnosAux(i) > AnosAux(j) Then
                AnoAux = AnosAux(i)
                AnosAux(i) = AnosAux(j)
                AnosAux(j) = AnoAux
            End If
        Next j
    Next i
    
    k = 0
    
    'Laço para remover anos repetidos e escrever em um novo vetor
    For i = 1 To AnoLen ' i = 1 para excluir a célula em branco
        If AnosAux(i - 1) <> AnosAux(i) Then
            Anos(k) = AnosAux(i)
            k = k + 1
            ReDim Preserve Anos(0 To k)
        End If
    Next i
    
    AnosLen = k
    
    'Criar tabela
    Worksheets("ProcessosUnidades").Activate
    
    Cells.Select
    Selection.Clear
    ActiveWindow.FreezePanes = False
    
    ActiveSheet.Cells.Select 'Select all cells
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -4.99893185216834E-02
        .PatternTintAndShade = 0
    End With
    
    'Escreve Unidades
    For i = 0 To UnidLen
        ActiveSheet.Cells(i + 3, 1) = Unidades(i)
    Next i
    
    'Escreve meses
    Meses = Array("Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez")
    MesesLen = 11 'Array vai de 0 a 11, totalizando os 12 meses
    For i = 0 To AnosLen - 1
        For j = 0 To 11
            p = (j + 2) + 12 * i
            ActiveSheet.Cells(2, p) = Meses(j)
        Next j
    Next i
    
    'Escreve anos
    For i = 0 To AnosLen - 1
        p = 12 * i + 2
        ActiveSheet.Cells(1, p) = Anos(i)
        ActiveSheet.Range(Cells(1, p), Cells(1, p + 11)).Select
        Selection.Merge
    Next i
    
    p = 2
    NRow = 2
    
    'Escreve Total
    Sheets("ProcessosUnidades").Range(Cells(1, 12 * AnosLen + 3), Cells(2, 12 * AnosLen + 3)).Merge
    Sheets("ProcessosUnidades").Range(Cells(1, 12 * AnosLen + 3), Cells(2, 12 * AnosLen + 3)) = "TOTAL"
    Sheets("ProcessosUnidades").Range(Cells(1, 12 * AnosLen + 3), Cells(2, 12 * AnosLen + 3)).Select
    Call PUTitleColor
    
    'Escreve quantidade de documentos de 'Dados' em cada célula
    For i = 0 To UnidLen
        For j = 0 To AnosLen - 1
            For k = 0 To MesesLen
                If Sheets("Dados").Cells(NRow, 1).Value = Unidades(i) And Sheets("Dados").Cells(NRow, 2).Value = Anos(j) And Sheets("Dados").Cells(NRow, 3).Value = k + 1 Then
                    Sheets("ProcessosUnidades").Cells(i + 3, p) = Sheets("Dados").Cells(NRow, 4).Value
                    NRow = NRow + 1
                ElseIf Sheets("Dados").Cells(NRow, 1).Value <> Unidades(i) Or Sheets("Dados").Cells(NRow, 2).Value <> Anos(j) Or Sheets("Dados").Cells(NRow, 3).Value <> k + 1 Then
                    Sheets("ProcessosUnidades").Cells(i + 3, p) = "-"
                End If
            p = p + 1
            Next k
        Next j
        p = 2
        Sheets("ProcessosUnidades").Cells(i + 3, 12 * AnosLen + 3).FormulaR1C1 = "=SUM(R" & i + 3 & "C2:R" & i + 3 & "C" & 12 * AnosLen + 1 & ")"
    Next i
    
    ActiveSheet.Columns.AutoFit
    ActiveSheet.Range(Cells(1, 12 * AnosLen + 3), Cells(UnidLen + 3, 12 * AnosLen + 3)).Select
    Call PUCreateBorders
    ActiveSheet.Range(Cells(1, 1), Cells(UnidLen + 3, 13 + 12 * (AnosLen - 1))).Select
    Call PUCreateBorders
    ActiveSheet.Range(Cells(3, 2), Cells(UnidLen + 3, 13 + 12 * (AnosLen - 1))).Select
    Call PUAlignmentRight
    ActiveSheet.Range(Cells(1, 2), Cells(1, 13 + 12 * (AnosLen - 1))).Select
    Call PUAlignmentCenter
    ActiveSheet.Range(Cells(1, 2), Cells(1, 13 + 12 * (AnosLen - 1))).Select
    Call PUTitleColor
    ActiveSheet.Range(Cells(2, 2), Cells(2, 13 + 12 * (AnosLen - 1))).Select
    Call PUTitleColor
    
    For i = 4 To UnidLen + 3 Step 2
        ActiveSheet.Range(Cells(i - 1, 1), Cells(i - 1, 13 + 12 * (AnosLen - 1))).Select
        Call PULinesColorWhite
        ActiveSheet.Range(Cells(i, 1), Cells(i, 13 + 12 * (AnosLen - 1))).Select
        Call PULinesColor
        ActiveSheet.Cells(i - 1, 12 * AnosLen + 3).Select
        Call PULinesColorWhite
        ActiveSheet.Cells(i, 12 * AnosLen + 3).Select
        Call PULinesColor
    Next i
    
    Range("B3").Select
    ActiveWindow.FreezePanes = True
    
    ActiveSheet.Range(Cells(1, 1), Cells(2, 1)).Select
    Call PUTitleColor
    Selection.Merge
End Sub

Sub PUCreateBorders()
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub

Sub PUAlignmentRight()
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
End Sub

Sub PUAlignmentCenter()
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
End Sub

Sub PUTitleColor()
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 12611584
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Sub PULinesColor()
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
End Sub

Sub PULinesColorWhite()
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
