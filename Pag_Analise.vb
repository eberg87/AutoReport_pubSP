Sub Pagamentos()
    Worksheets("Análise dos Dados").Activate
    Cells.Select
    Selection.Clear
    Call ChangeSheetColor
    Call CreateTablePagComp
    Call CreateEquationPagComp
    Range("B2:C5").Select
    Call CreateBorder
    Call CreateTableProcRel
    Cells(1, 1).Select
End Sub

Sub ChangeSheetColor()
    'Change color of entire sheet
    ActiveSheet.Cells.Select 'Select all cells
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -4.99893185216834E-02
        .PatternTintAndShade = 0
    End With
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

Sub DarkBlueBackground()
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Sub CreateTablePagComp()
    Dim Unidade As String
    
    Unidade = "SMC"
    
    'Criar descrição
    ActiveSheet.Range(Cells(2, 2), Cells(2, 3)).Merge
    ActiveSheet.Range(Cells(2, 2), Cells(2, 3)).Select
    Selection = "Pagamentos: Compras"
    Call DarkBlueBackground
    Call SetTitle
    Selection.Font.Bold = True
    ActiveSheet.Cells(3, 2).Select
    Selection = "Tipo de Processo"
    Selection.Font.Bold = True
    Call SoftBlueBackground
    ActiveSheet.Cells(4, 2).Select
    Selection = "Quantidade de Processos"
    Selection.Font.Bold = True
    Call SoftBlueBackground
    ActiveSheet.Cells(5, 2).Select
    Selection = "Quantidade de Processos Não Relacionados"
    Selection.Font.Bold = True
    Call SoftBlueBackground
    
    'Criar unidade
    ActiveSheet.Cells(3, 3).Select
    Selection = Unidade
    Selection.Font.Bold = True
    Call SoftBlueBackground
    
    'Formatar células para receber equações
    ActiveSheet.Cells(4, 3).Select
    Selection.Font.Bold = False
    Call WhiteBackground
    ActiveSheet.Cells(5, 3).Select
    Selection.Font.Bold = False
    Call WhiteBackground
End Sub

Sub CreateEquationPagComp()
    Range("C4").Select
    ActiveCell.FormulaR1C1 = "=COUNTIFS(Pagamentos!C3,'Análise dos Dados'!R[-1]C)"
    Range("C5").Select
    ActiveCell.FormulaR1C1 = "=COUNTIFS(Pagamentos!C3,'Análise dos Dados'!R[-2]C,Pagamentos!C[7],""ERR_SEM_TIPO_DE_PROCESSO"")"
End Sub

Sub CreateBorder()
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

Sub SetTitle()
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
End Sub

Sub CreateTableProcRel()
    Dim NCol        As Integer
    Dim NRow        As Integer
    Dim i           As Integer
    Dim ProcLen     As Integer
    Dim Processos() As Variant
    Dim Unidade     As String

    NCol = 10
    NRow = 2
    i = 0
    ProcLen = 0
    Processos() = Array("")
    Unidade = "SMC"

    Call OrganizeProc 'Organiza todos os processos em ordem alfabética
    Call FillEmptyProc 'Para os processos existentes, coloca '-' quando não houver Tipo de Processo
    
    'laço para criar uma array com o nome de todas os tipos de processo disponíveis
    Worksheets("Pagamentos").Activate
    Do While Cells(NRow, NCol).Value <> ""
        Processos(i) = Cells(NRow, NCol).Value
        If i = 0 Then
            i = i + 1
            ProcLen = i 'equação que irá definir o tamanho do vetor
            ReDim Preserve Processos(0 To i) 'Array começa com 0 sempre
        ElseIf Processos(i) <> Processos(i - 1) Then
            i = i + 1
            ProcLen = i 'equação que irá definir o tamanho do vetor
            ReDim Preserve Processos(0 To i) 'Array começa com 0 sempre
        End If
        NRow = NRow + 1
    Loop
    If Processos(i) = Processos(i - 1) Then
        i = i - 1
        ProcLen = i 'equação que irá definir o tamanho do vetor
        ReDim Preserve Processos(0 To i)
    End If
    If Processos(i) = "" Then
        i = i - 1
        ProcLen = i 'equação que irá definir o tamanho do vetor
        ReDim Preserve Processos(0 To i)
    End If
    
    'Laço para criar unidades em 'Análise de Dados'
    Worksheets("Análise dos Dados").Activate
    ActiveSheet.Range(Cells(7, 2), Cells(7, 3)).Merge
    ActiveSheet.Range(Cells(7, 2), Cells(7, 3)).Select
    Selection = "Processos Relacionados aos Processos de 'Pagamentos: Compras'"
    Call DarkBlueBackground
    Call SetTitle
    Selection.Font.Bold = True
    'Criar guia para tipos de processo
    ActiveSheet.Cells(8, 2).Select
    Selection = "Tipo de Processo"
    Selection.Font.Bold = True
    Call SoftBlueBackground
    'Criar unidade
    ActiveSheet.Cells(8, 3).Select
    Selection = Unidade
    Selection.Font.Bold = True
    Call SoftBlueBackground
    'Criar Lista de Processos
    For i = 0 To ProcLen
        ActiveSheet.Cells(i + 9, 2).Select
        Selection = Processos(i)
        Selection.Font.Bold = True
        Call SoftBlueBackground
        ActiveSheet.Cells(i + 9, 3).Select
        Selection.Font.Bold = False
        Selection.FormulaR1C1 = "=COUNTIFS(Pagamentos!C10,'Análise dos Dados'!RC[-1],Pagamentos!C3,'Análise dos Dados'!R8C3)"
        Call WhiteBackground
    Next i
    
    Range(Cells(7, 2), Cells(i + 8, 3)).Select
    Call CreateBorder
    Cells(1, 1).Select
End Sub

Sub OrganizeProc()
    Worksheets("Pagamentos").AutoFilter.Sort.SortFields.Clear
    Worksheets("Pagamentos").AutoFilter.Sort.SortFields.Add Key:= _
        Range("J:J"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With Worksheets("Pagamentos").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub FillEmptyProc()
    Dim NCol        As Integer
    Dim NRow        As Integer
    Dim i           As Integer

    NCol = 10
    NRow = 2
    i = 0

    Worksheets("Pagamentos").Activate
    Do While Cells(NRow, NCol).Value = ""
        If Cells(NRow, NCol).Value = "" Then
            Cells(NRow, NCol) = "ERR_SEM_TIPO_DE_PROCESSO"
        End If
    NRow = NRow + 1
    Loop
End Sub

