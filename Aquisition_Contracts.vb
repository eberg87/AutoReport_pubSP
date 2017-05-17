Sub AquisicoesContratacoes()
    Call CreateGeral
    Call CreateDF
    Call CreateStsProcesso
    Worksheets("Geral").Activate
End Sub

Sub CreateGeral()
    Worksheets("Geral").Activate
    Cells.Select
    Selection.Clear
    ActiveWindow.FreezePanes = False
    Call ChangeSheetColor
    Call GeralChangeLinesColor
    Call GeralCreateTableTitle
    Call GeralCreateBorders
    Call GeralCreateTableText
    Call GeralCreateEquationCells
    Call GeralCellWidthAdjust
    Cells(1, 1).Select
    Sheets("Tabela Dinâmica_Geral").Select
    ActiveSheet.PivotTables("Tabela dinâmica2").PivotCache.Refresh
    Cells(1, 1).Select
End Sub

Sub CreateDF()
    Worksheets("Documentos Faltantes").Activate
    Cells.Select
    Selection.Clear
    ActiveWindow.FreezePanes = False
    Call ChangeSheetColor
    Call DFChangeLinesColor
    Call DFCreateTableTitle
    Call DFCreateBorders
    Call DFCreateTableText
    Call DFCreateEquationCells
    Call DFCellWidthAdjust
    Cells(1, 1).Select
End Sub

Sub CreateStsProcesso()
    Worksheets("Status dos Processos").Activate
    Cells.Select
    Selection.Clear
    ActiveWindow.FreezePanes = False
    Call ChangeSheetColor
    Call SPChangeLinesColor
    Call SPCreateTableTitle
    Call SPCreateBorders
    Call SPCreateTableText
    Call SPCreateEquationCells
    Call SPCellWidthAdjust
    Cells(1, 1).Select
End Sub

Sub GeralCellWidthAdjust()
    'Adjust width
    ActiveSheet.Cells.EntireColumn.AutoFit
End Sub

Sub ChangeSheetColor()
    'Change color of every table title
    ActiveSheet.Cells.Select 'Select all cells
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -4.99893185216834E-02
        .PatternTintAndShade = 0
    End With
End Sub

Sub GeralChangeLinesColor()
    ActiveSheet.Range("A2:E16").Select
    Call WhiteBackground
    ActiveSheet.Range("A17").Select
    Call SoftBlueBackground
End Sub

Sub WhiteBackground()
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Sub SoftBlueBackground()
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
End Sub

Sub GeralCreateTableTitle()
    Dim i          As Integer
    Dim TableTitle As Variant
    TableTitle = Array("Tipo Processo", "Quantidade de Processos", "Quantidade de processos com documentos", "Quantidade de processos que faltam documentos", "Quantidade de processos que faltam documentos (%)")
    'Criar títulos
    For i = 1 To 5
        ActiveSheet.Cells(1, i).Select
        Selection = TableTitle(i - 1)
        Call GeralChangeTitleColor
        Selection.Font.Bold = True
    Next i
End Sub

Sub GeralChangeTitleColor()
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub

Sub GeralCreateBorders()
    'Criar grelha das tabelas
    ActiveSheet.Range("A1:E17").Select
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
    
Sub GeralCreateTableText()
    ActiveSheet.Range("A2") = "Compras: Acionamento de Ata de Registro de Preços Própria/Participante"
    ActiveSheet.Range("A3") = "Compras: Adesão a Ata de Registro de Preços - Não Participante"
    ActiveSheet.Range("A4") = "Licitação: Concorrência"
    ActiveSheet.Range("A5") = "Licitação: Concorrência-Registro de Preço"
    ActiveSheet.Range("A6") = "licitação: Concurso"
    ActiveSheet.Range("A7") = "Licitação: Convite"
    ActiveSheet.Range("A8") = "Licitação: Dispensa"
    ActiveSheet.Range("A9") = "Licitação: Inexigibilidade"
    ActiveSheet.Range("A10") = "Licitação: Inexigibilidade-Registro de Preço"
    ActiveSheet.Range("A11") = "Licitação: Leilão"
    ActiveSheet.Range("A12") = "Licitação: Pregão Eletrônico"
    ActiveSheet.Range("A13") = "Licitação: Pregão Eletrônico-Registro de Preço"
    ActiveSheet.Range("A14") = "Licitação: Pregão Presencial"
    ActiveSheet.Range("A15") = "Licitação: Requisição Inicial"
    ActiveSheet.Range("A16") = "Licitação: Tomada de Preços"
    ActiveSheet.Range("A17") = "TOTAL"
End Sub

Sub GeralCreateEquationCells()
    Dim i As Integer
    'Criar equações para Quantidade de Processos (todas categorias)
    For i = 2 To 16
        ActiveSheet.Cells(i, 2).FormulaR1C1 = "=COUNTIF('Dados Brutos'!C,RC[-1])"
        ActiveSheet.Cells(i, 3).FormulaR1C1 = "=COUNTIF('Dados Brutos'!C[-1],Geral!RC[-2])-COUNTIFS('Dados Brutos'!C[-1],Geral!RC[-2],'Dados Brutos'!C[46],""Processo iniciado"")"
        ActiveSheet.Cells(i, 4).FormulaR1C1 = "=COUNTIFS('Dados Brutos'!C[-2],Geral!RC[-3],'Dados Brutos'!C[45],""Falta documento"")"
        ActiveSheet.Cells(i, 5).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-2],0)" 'Usado IFERROR para evitar erros de divisão por zero
        ActiveSheet.Cells(i, 5).NumberFormat = "0.00%"
    Next i
    
    'Criar equações para Total
    For i = 2 To 4
        ActiveSheet.Cells(17, i).FormulaR1C1 = "=SUM(R[-15]C:R[-1]C)"
    Next i
    
    ActiveSheet.Cells(17, 5).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-2],0)"
    ActiveSheet.Cells(17, 5).NumberFormat = "0.00%"
End Sub

Sub DFCellWidthAdjust()
    'Adjust width
    Columns(1).ColumnWidth = 2
    Columns(11).ColumnWidth = 8
    ActiveSheet.Columns.AutoFit
End Sub

Sub DFChangeLinesColor()
    Dim i As Integer
    Dim Col1 As Integer
    Col1 = 10
    For i = 0 To 39 Step 5
        ActiveSheet.Range(Cells(i + 1, 2), Cells(i + 1, Col1)).Select
        Call WhiteBackground
        ActiveSheet.Range(Cells(i + 2, 2), Cells(i + 2, Col1)).Select
        Call SoftBlueBackground
        ActiveSheet.Range(Cells(i + 3, 2), Cells(i + 3, Col1)).Select
        Call WhiteBackground
        Col1 = Col1 - 1
    Next i
    i = 0
    Col1 = 20
    ActiveSheet.Range(Cells(i + 1, Col1 - 8), Cells(i + 1, Col1)).Select
    Call WhiteBackground
    ActiveSheet.Range(Cells(i + 2, Col1 - 8), Cells(i + 2, Col1)).Select
    Call SoftBlueBackground
    ActiveSheet.Range(Cells(i + 3, Col1 - 8), Cells(i + 3, Col1)).Select
    Call WhiteBackground
End Sub

Sub DFCreateTableTitle()
    Dim k          As Integer
    Dim Col1       As Integer
    Dim TableTitle As Variant
    TableTitle = Array("Empenho", "Licitação Homologada", "Pregão em Andamento", "Despacho Autorizatório", "Reserva", "Pesquisa de Preço", "Especificação", "Requisição", "TOTAIS")

    'Criar títulos
    k = 1
    Col1 = 10
    For i = 0 To 8
        If i < 8 Then
            ActiveSheet.Range(Cells(k, 2), Cells(k, Col1)).Select
            Call DFChangeTitleColor
            Selection.Merge
            Selection = TableTitle(i)
            Selection.Font.Bold = True
            k = k + 5
            Col1 = Col1 - 1
        Else
            ActiveSheet.Range(Cells(1, 12), Cells(1, 20)).Select
            Call DFChangeTitleColor
            Selection.Merge
            Selection = TableTitle(i)
            Selection.Font.Bold = True
        End If
    Next i
End Sub

Sub DFChangeTitleColor()
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub

Sub DFCreateBorders()
    'Criar grelha das tabelas
    ActiveSheet.Range("B1:J4,B6:I9,B11:H14,B16:G19,B21:F24,B26:E29,B31:D34,B36:C39,L1:T4").Select
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
    
Sub DFCreateTableText()
    'Escrever textos descritivos das tabelas
    Range("C2,C7,C12,C17,C22,C27,C32,C37,M2").Select
    Selection.Font.Bold = True
    Selection = "Formulário de Requisição"
    
    Range("D2,D7,D12,D17,D22,D27,D32,N2").Select
    Selection.Font.Bold = True
    Selection = "Especificação Técnica"
    
    Range("E2,E7,E12,E17,E22,E27,O2").Select
    Selection.Font.Bold = True
    Selection = "pesquisa_preco"
    
    Range("F2,F7,F12,F17,F22,P2").Select
    Selection.Font.Bold = True
    Selection = "reserva"
    
    Range("G2,G7,G12,G17,Q2").Select
    Selection.Font.Bold = True
    Selection = "despacho_autorizatorio"
    
    Range("H2,H7,H12,R2").Select
    Selection.Font.Bold = True
    Selection = "pregao_andamento"
    
    Range("I2,I7,S2").Select
    Selection.Font.Bold = True
    Selection = "licitacao_homologada"
    
    Range("J2,T2").Select
    Selection.Font.Bold = True
    Selection = "empenho_realizado"
    
    Range("B3,B8,B13,B18,B23,B28,B33,B38,L3").Select
    Selection.Font.Bold = True
    Selection = "TOTAL"
    
    Range("B4,B9,B14,B19,B24,B29,B34,B39,L4").Select
    Selection = "SMC"
End Sub

Sub DFCreateEquationCells()
    Dim i          As Integer
    Dim k          As Integer
    Dim Col1       As Integer
    Dim itable     As Integer
    Dim TableTitle As Variant
    TableTitle = Array("Empenho", "Licitação Homologada", "Pregão em Andamento", "Despacho Autorizatório", "Reserva", "Pesquisa de Preço", "Especificação", "Processo Iniciado")
    itable = 0
    Col1 = 8
    'Write equations
    For i = 4 To 40 Step 5
        For k = Col1 To 1 Step -1
            ActiveSheet.Cells(i, k + 2).FormulaR1C1 = "=COUNTIFS('Dados Brutos'!C47,""" & TableTitle(itable) & """,'Dados Brutos'!C[35],""0"",'Dados Brutos'!C4,'Documentos faltantes'!RC[-" & k & "])"
            'ActiveSheet.Cells(i, k + 2) = Application.WorksheetFunction.Sum(ActiveSheet.Cells(i, k + 2).Offset(1, 0))
            ActiveSheet.Cells(i - 1, k + 2).FormulaR1C1 = "= SUM(R[1]C:R[1]C)"
            ActiveSheet.Cells(i - 1, k + 2).Font.Bold = True
        Next k
        ActiveSheet.Cells(i - 2, 2).FormulaR1C1 = "= SUM(R[+1]C[+1]:R[+1]C[+" & Col1 & "])"
        ActiveSheet.Cells(i - 2, 2).Font.Bold = True
        itable = itable + 1
        Col1 = Col1 - 1
    Next i
    
    i = 4
    For k = 18 To 11 Step -1
        ActiveSheet.Cells(i, k + 2).FormulaR1C1 = "=SUMIF(C2:C2,RC12,C[-10]:C[-10])"
        'ActiveSheet.Cells(i, k + 2) = Application.WorksheetFunction.Sum(ActiveSheet.Cells(i, k + 2).Offset(1, 0))
        ActiveSheet.Cells(i - 1, k + 2).FormulaR1C1 = "= SUM(R[1]C:R[1]C)"
        ActiveSheet.Cells(i - 1, k + 2).Font.Bold = True
    Next k
End Sub

Sub SPChangeLinesColor()
    ActiveSheet.Range("B2:V4").Select
    Call SoftBlueBackground
    ActiveSheet.Range("B4, D4, F4, H4, J4, L4, N4, P4, R4, T4, V4").Select
    Call WhiteBackground
    ActiveSheet.Range("B3, D3, F3, H3, J3, L3, N3, P3, R3, T3, V3").Select
    Call DarkBlueBackground
End Sub

Sub DarkBlueBackground()
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Sub SPCreateTableTitle()
    ActiveSheet.Cells(2, 2) = "STATUS DOS PRCESSOS NAS UNIDADES"
    ActiveSheet.Range(Cells(2, 2), Cells(2, 22)).Select
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection.Font
        .Name = "Calibri"
        .Size = 18
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
End Sub

Sub SPCreateTableText()
    Dim i         As Integer
    Dim TableText As Variant
    TableText = Array("Órgãos", "Processo Iniciado", "Em Requisição de Compras", "Especificação Técnica", "Pesquisa de Preço", "Despacho Autorizatório", "Reserva", "Pregão em Andamento", "Licitação Homologada", "Empenho", "Total de Processos na Unidade")
    
    For i = 1 To 11
        ActiveSheet.Cells(3, 2 * i) = TableText(i - 1)
        ActiveSheet.Cells(3, 2 * i).Select
        Call TableTextAdjust
    Next i
End Sub

Sub TableTextAdjust()
        With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Rows("3:3").RowHeight = 90
End Sub

Sub SPCreateBorders()
    Range("B2:V4").Select
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
    
    Range("C3:C4,E3:E4,G3:G4,I3:I4,K3:K4,M3:M4,O3:O4,Q3:Q4,S3:S4,U3:U4").Select
    Selection.Merge
    
    Cells(4, 2) = "SMC"
End Sub

Sub SPCreateEquationCells()
    Dim i As Integer

    For i = 4 To 20 Step 2
        Cells(4, i).Select
        ActiveCell.FormulaR1C1 = "=COUNTIFS('Dados Brutos'!C47,'Status dos Processos'!R3C" & i & ",'Dados Brutos'!C4,'Status dos Processos'!RC[-" & i - 2 & "])"
    Next i

    Cells(4, 22).FormulaR1C1 = "=SUM(RC[-18], RC[-16], RC[-14], RC[-12], RC[-10], RC[-8], RC[-6], RC[-4], RC[-2])"
End Sub

Sub SPCellWidthAdjust()
    Range("C:C,E:E,G:G,I:I,K:K,M:M,O:O,Q:Q,S:S,U:U").ColumnWidth = 1.5
    Range("B:B,D:D,F:F,H:H,J:J,L:L,N:N,P:P,R:R,T:T,V:V").ColumnWidth = 14
    Columns("A:A").ColumnWidth = 1
End Sub

