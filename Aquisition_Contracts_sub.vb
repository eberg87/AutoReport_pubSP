Public ProcLen    As Integer 'tamanho do vetor de 'Tipo de Processo'. Pode ser usado UBound
Public UnidLen    As Integer 'tamanho do vetor de 'Unidades'. Pode ser usado UBound
Public TipoProc   As Variant 'vetor 'Tipo de Processos'
Public Unidades   As Variant 'vetor 'Unidades'
Public LenDocType As Integer 'tamanho do vetor de 'DocType'. Pode ser usado UBound
Public DocType    As Variant 'vetor 'DocType'
Public Ref        As Integer 'variavel de referência para escrever a segunda tabela na posição

Sub AquisicoesContratacoes() 'programa principal para o arquivo *.xlsm
    Call CreateGeral 'para planilha 'Geral'
    Call PorOrgao 'para planilha 'Por Órgão'
    Call DocFaltante 'para planilha 'Documentos Faltantes'
    
    Worksheets("Geral").Activate
    ActiveSheet.Cells(1, 1).Select
End Sub


'Ver arquivos em \\nas.prodam\CGTIC\Relatórios SEI - SMC\Aquisições e Contratações SMC
Sub CreateGeral()
    Worksheets("Geral").Activate
    Cells.Select
    Selection.Clear
    ActiveWindow.FreezePanes = False
    Call ChangeSheetColor
    Call GeralChangeLinesColor
    Call GeralCreateTableTitle
    ActiveSheet.Range("A1:E17").Select
    Call CreateBorders
    Call GeralCreateTableText
    Call GeralCreateEquationCells
    Call GeralCellWidthAdjust
    Sheets("Tabela Dinâmica_Geral").Select
    ActiveSheet.PivotTables("Tabela dinâmica2").PivotCache.Refresh
End Sub

Sub GeralCellWidthAdjust()
    'Adjust width
    ActiveSheet.Cells.EntireColumn.AutoFit
End Sub

Sub GeralChangeLinesColor()
    ActiveSheet.Range("A2:E16").Select
    Call WhiteBackground
    ActiveSheet.Range("A17").Select
    Call SoftBlueBackground
End Sub

Sub GeralCreateTableTitle()
    Dim i          As Integer
    Dim TableTitle As Variant
    TableTitle = Array("Tipo Processo", "Quantidade de Processos", "Quantidade de processos com documentos", "Quantidade de processos que faltam documentos", "Quantidade de processos que faltam documentos (%)")
    'Criar títulos
    For i = 1 To 5
        ActiveSheet.Cells(1, i).Select
        Selection = TableTitle(i - 1)
        Call ChangeTitleColor
        Selection.Font.Bold = True
    Next i
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
        ActiveSheet.Cells(i, 5).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-2],0)"  'Usado IFERROR para evitar erros de divisão por zero
        ActiveSheet.Cells(i, 5).NumberFormat = "0.00%"
    Next i
    
    'Criar equações para Total
    For i = 2 To 4
        ActiveSheet.Cells(17, i).FormulaR1C1 = "=SUM(R[-15]C:R[-1]C)"
    Next i
    
    ActiveSheet.Cells(17, 5).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-2],0)" 'escreve a divisão e caso seja uma divisão por zero, escreve zero
    ActiveSheet.Cells(17, 5).NumberFormat = "0.00%" 'altera formato do número para o descrito
End Sub

Sub PorOrgao()
    Worksheets("Por Órgão").Activate 'seleciona WorkSheet
    Cells.Select 'seleciona todas as células
    Selection.Clear 'limpa a seleção
    ActiveWindow.FreezePanes = False
    Call ChangeSheetColor 'altera fundo da planilha
    Call SortTP 'organiza em ordem alfabética os tipos de processos
    Call CreateTPArray
    Call SortOG 'organiza em ordem alfabética os orgãos geradores
    Call CreateOGArray
    
    'criar tabela principal
    Call CreateFirstTable
    Call CreateEquationsFT
    
    'criar tabela 'Status dos Processos nas Unidades'
    Call CreateSecondTable
    Call CreateEquationsST
    
    'criar grades das tabelas
    Range(Cells(1, 1), Cells(UnidLen + 4, 2 * ProcLen + 5)).Select
    Call CreateBorders
    
    Range(Cells(Ref, 1), Cells(Ref + UnidLen + 2, LenDocType + 2)).Select
    Call CreateBorders
    
    ActiveSheet.Cells(1, 1).Select
End Sub

Sub SortTP() 'algumas subs são feitas utilizando a ferramenta de gravar macro. Ela gera um código, geralmente, funcional e que exige poucas alterações.
    ActiveWorkbook.Worksheets("Dados Brutos").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Dados Brutos").AutoFilter.Sort.SortFields.Add Key _
        :=Range("B1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    Call AlphaSort
End Sub

Sub SortOG()
    ActiveWorkbook.Worksheets("Dados Brutos").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Dados Brutos").AutoFilter.Sort.SortFields.Add Key _
        :=Range("D1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    Call AlphaSort
End Sub

Sub AlphaSort()
    With ActiveWorkbook.Worksheets("Dados Brutos").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub CreateTPArray()
    Dim i    As Integer
    Dim NRow As Integer
    Dim NCol As Integer
    
    NRow = 2
    NCol = 2
    ProcLen = 0
    TipoProc = Array("")
    
    'laço para criar uma array com o nome de todos os tipos de processo disponíveis
    Worksheets("Dados Brutos").Activate
    Do While Cells(NRow, NCol).Value <> ""
        TipoProc(i) = Cells(NRow, NCol).Value
        If i = 0 Then
            i = i + 1
            ProcLen = i 'equação que irá definir o tamanho do vetor
            ReDim Preserve TipoProc(0 To i) 'Array começa com 0 sempre
        ElseIf TipoProc(i) <> TipoProc(i - 1) Then
            i = i + 1
            ProcLen = i 'equação que irá definir o tamanho do vetor
            ReDim Preserve TipoProc(0 To i) 'Array começa com 0 sempre
        End If
        NRow = NRow + 1
    Loop
    If TipoProc(i) = TipoProc(i - 1) Then
        i = i - 1
        ProcLen = i 'equação que irá definir o tamanho do vetor
        ReDim Preserve TipoProc(0 To i)
    End If
    If TipoProc(i) = "" Then
        i = i - 1
        ProcLen = i 'equação que irá definir o tamanho do vetor
        ReDim Preserve TipoProc(0 To i)
    End If
End Sub

Sub CreateOGArray()
    Dim i    As Integer
    Dim NRow As Integer
    Dim NCol As Integer
    
    NRow = 2
    NCol = 4
    UnidLen = 0
    Unidades = Array("")
    
    'laço para criar uma array com o nome de todas as unidades disponíveis
    Worksheets("Dados Brutos").Activate
    Do While Cells(NRow, NCol).Value <> "" 'laço que se repete até o conteúdo da célula ser nulo
        Unidades(i) = Cells(NRow, NCol).Value 'copia o valor da célula para o array
        If i = 0 Then 'na primeira posição do vetor sempre é escrita uma unidade
            i = i + 1
            UnidLen = i 'equação que irá definir o tamanho do vetor. Pode ser usado UBound após ter todo o vetor. Foi feito dessa forma para fins de aprendizado.
                        'Contudo, pode ser mais interessante redefinir o código usando UBound, pois para cada nova unidade uma operação deve ser feita, consumindo tempo de máquina
            ReDim Preserve Unidades(0 To i) 'Array começa com 0 sempre. Preserva os valores e redimensiona o vetor
        ElseIf Unidades(i) <> Unidades(i - 1) Then 'se o conteúdo do vetor na posição i for diferente do conteúdo do vetor na sua posição anterior (i-1), _
                                                   'então é escrito um novo conteúdo
            i = i + 1
            UnidLen = i 'equação que irá definir o tamanho do vetor
            ReDim Preserve Unidades(0 To i) 'Array começa com 0 sempre
        End If 'se o conteúdo for igual, então o programa ignora. Pode ser interessante adicionar um condicional cobrindo essa possibilidade para evitar alguma má execução
        NRow = NRow + 1
    Loop
    If Unidades(i) = Unidades(i - 1) Then 'se i conteúdo de i for igual ao conteúdo de i-1, então ele apaga a última posição. Isso deve ser feito pois o programa estava duplicando a última unidade
        i = i - 1
        UnidLen = i 'equação que irá definir o tamanho do vetor
        ReDim Preserve Unidades(0 To i) 'redimensiona o vetor para o tamanho que não inclua a unidade repetida
    End If
    If Unidades(i) = "" Then 'após o final do laço, o programa vez ou outra acabava inserindo um valor nulo na última posição do vetor.
        i = i - 1
        UnidLen = i 'equação que irá definir o tamanho do vetor
        ReDim Preserve Unidades(0 To i)
    End If
End Sub

Sub CreateFirstTable() 'cria a primeira tabela da planilha "Por Órgãos".
    Dim i As Integer
    
    'cria textos estáticos
    Worksheets("Por Órgão").Activate
    ActiveSheet.Cells(2, 1).Select
    Selection = "Órgãos"
    Call SoftBlueBackground
    Call BreakText
    
    ActiveSheet.Cells(UnidLen + 4, 1).Select
    Selection = "TOTAL"
    Call ChangeTitleColor
    Call BreakText
    
    'cria coluna de 'unidades'
    For i = 0 To UnidLen 'laço que percorre o vetor de unidades de zero até o seu tamanho. Pode ser usado UBound no lugar dos incrementos usados acima.
        ActiveSheet.Cells(i + 3, 1).Select
        Selection = Unidades(i)
        Call SoftBlueBackground
        Call BreakText
    Next i
    
    'cria linha 'tipo de processo'
    For i = 0 To ProcLen
        ActiveSheet.Cells(1, 2 * i + 2) = TipoProc(i)
        ActiveSheet.Range(Cells(1, 2 * i + 2), Cells(1, 2 * i + 3)).Select
        Selection.Merge
        Call ChangeTitleColor
        Call BreakText
        ActiveSheet.Cells(2, 2 * i + 2).Select
        Selection = "Número de Processos"
        Call SoftBlueBackground
        Call BreakText
        ActiveSheet.Cells(2, 2 * i + 3).Select
        Selection = "Processos que faltam documentos"
        Call SoftBlueBackground
        Call BreakText
    Next i
    
        ActiveSheet.Cells(1, 2 * ProcLen + 4) = "TOTAL"
        ActiveSheet.Range(Cells(1, 2 * ProcLen + 4), Cells(1, 2 * ProcLen + 5)).Select
        Selection.Merge
        Call ChangeTitleColor
        Call BreakText
        ActiveSheet.Cells(2, 2 * ProcLen + 4).Select
        Selection = "Número de Processos"
        Call SoftBlueBackground
        Call BreakText
        ActiveSheet.Cells(2, 2 * ProcLen + 5).Select
        Selection = "Processos que faltam documentos"
        Call SoftBlueBackground
        Call BreakText
    
End Sub

Sub CreateEquationsFT()
    Dim i           As Integer
    Dim j           As Integer
    Dim SumProcCDoc As Integer
    Dim SumProcSDoc As Integer
    
    SumProcCDoc = 0
    SumProcSDoc = 0

    'criar fórmulas de soma horinzontais da primeira tabela
    For i = 0 To ProcLen + 1
       ActiveSheet.Cells(UnidLen + 4, 2 * i + 2).Select
       Selection.FormulaR1C1 = "=SUM(R3C:R" & UnidLen + 3 & "C)"
       Call SoftBlueBackground
       ActiveSheet.Cells(UnidLen + 4, 2 * i + 3).Select
       Selection.FormulaR1C1 = "=SUM(R3C:R" & UnidLen + 3 & "C)"
       Call SoftBlueBackground
    Next i
    
    'criar formulas do interior da tabela. Pode ser usadas equações do VBA para que só se exiba os valores, uma alteração que pode ser feita no futuro
    For j = 0 To ProcLen
        For i = 0 To UnidLen
            Cells(i + 3, 2 * j + 2).Select
            Selection.FormulaR1C1 = "=COUNTIFS('Dados Brutos'!C4,'Por Órgão'!RC1,'Dados Brutos'!C2,'Por Órgão'!R1C:R1C[1])"
            Call WhiteBackground
            Cells(i + 3, 2 * j + 3).Select
            Selection.FormulaR1C1 = "=COUNTIFS('Dados Brutos'!C4,'Por Órgão'!RC1,'Dados Brutos'!C2,'Por Órgão'!R1C[-1],'Dados Brutos'!C49,""Falta documento"")"
            Call WhiteBackground
        Next i
    Next j
    
    'criar fórmulas de soma verticais da primeira tabela
    For i = 0 To UnidLen
        For j = 0 To ProcLen
            SumProcCDoc = SumProcCDoc + Cells(i + 3, 2 * j + 2).Value
            SumProcSDoc = SumProcSDoc + Cells(i + 3, 2 * j + 3).Value
        Next j
    ActiveSheet.Cells(i + 3, 2 * ProcLen + 4).Select
    Selection = SumProcCDoc
    Call SoftBlueBackground
    ActiveSheet.Cells(i + 3, 2 * ProcLen + 5).Select
    Selection = SumProcSDoc
    Call SoftBlueBackground
    SumProcCDoc = 0
    SumProcSDoc = 0
    Next i
    
End Sub

Sub CreateSecondTable()
    Dim Ref        As Integer 'variavel de referência para escrever a segunda tabela na posição
    
    Ref = UnidLen + 7
    DocType = Array("Processo Iniciado", "Requisição de Compras", "Especificação Técnica", "Pesquisa de Preço", "Despacho Autorizatório", "Reserva", "Pregão em Andamento", "Licitação Homologada", "Empenho")
    LenDocType = UBound(DocType) 'tamanho do vetor DocType
    
    Worksheets("Por Órgão").Activate
    ActiveSheet.Range(Cells(Ref, 1), Cells(Ref, LenDocType + 3)).Select
    Selection.Merge
    Selection = "STATUS DOS PROCESSOS NAS UNIDADES"
    Call ChangeTitleColor
    Call BreakText
    
    ActiveSheet.Cells(Ref + 1, 1).Select
    Selection.Merge
    Selection = "Órgãos"
    Call SoftBlueBackground
    Call BreakText
    
    ActiveSheet.Cells(Ref + 1, 2).Select
    Selection.Merge
    Selection = "Total de Processos na Unidade"
    Call SoftBlueBackground
    Call BreakText
    
    For i = 0 To LenDocType
        ActiveSheet.Cells(Ref + 1, i + 3).Select
        Selection = DocType(i)
        Call SoftBlueBackground
        Call BreakText
    Next i
    
    'cria coluna 'unidades'
    For i = 0 To UnidLen
        ActiveSheet.Cells(Ref + i + 2, 1).Select
        Selection = Unidades(i)
        Call SoftBlueBackground
        Call BreakText
    Next i
End Sub

Sub CreateEquationsST()
    Dim i As Integer
    Dim j As Integer
    
    Ref = UnidLen + 7 'referência é utilizada no caso das tabelas que já existem e que seja definida dinamicamente. Caso contrário, deveriamos ter uma quantidade _
                      'pré-definida de unidades para que saibamos qual posição deveriamos ter a nova tabela. Caso sejam adicionadas ou removidas unidades, o código _
                      'irá gerar essa referência e plotar a nova tabela onde se deve
    
    For j = 0 To LenDocType 'como deve-se percorrer a tabela integralmente, é necessário dois 'for' para que seja possível caminhar pelas linhas e colunas.
        For i = 0 To UnidLen
            Cells(Ref + i + 2, j + 3).Select
            Selection.FormulaR1C1 = "=COUNTIFS('Dados Brutos'!C47,'Por Órgão'!R[" & -i - 1 & "]C,'Dados Brutos'!C4,'Por Órgão'!RC1)"
            Call WhiteBackground
        Next i
    Next j
    
    For i = 0 To UnidLen
        ActiveSheet.Cells(Ref + i + 2, 2).Select
        Selection.FormulaR1C1 = "=SUM(RC[1]:RC[" & LenDocType + 1 & "])"
        'Call WhiteBackground
    Next i
End Sub

Sub DocFaltante()
    Worksheets("Documentos faltantes").Activate 'seleciona WorkSheet
    Cells.Select 'seleciona todas as células
    Selection.Clear 'limpa a seleção
    ActiveWindow.FreezePanes = False
    Call ChangeSheetColor 'altera fundo da planilha
    Call TablesDF
    Call EquationDF
    Call ChangeOrderDocType
    Call BordersDF
    
    ActiveSheet.Cells(1, 1).Select
End Sub

Sub TablesDF()
    Dim i             As Integer
    Dim j             As Integer
    Dim LenDocTypeAUX As Integer
    
    Ref = UnidLen + 5
    
    LenDocTypeAUX = LenDocType
    
    Worksheets("Documentos faltantes").Activate
    
    'operação para trocar a posição de 'Despacho Autorizatório' e 'Reserva'
    Call ChangeOrderDocType
    
    For i = 0 To LenDocType - 1
        ActiveSheet.Range(Cells(i * Ref + 1, 2), Cells(i * Ref + 1, LenDocType - i + 2)).Select
        Selection.Merge
        Selection = DocType(LenDocType - i)
        Call ChangeTitleColor
        Call BreakText
    Next i
    
    For i = 0 To LenDocType - 1
        For j = 0 To LenDocTypeAUX - 1
            ActiveSheet.Cells(i * Ref + 2, j + 3).Select
            Selection.Merge
            Selection = DocType(j + 1)
            Call SoftBlueBackground
            Call BreakText
        Next j
        LenDocTypeAUX = LenDocTypeAUX - 1
    Next i
    
    For i = 0 To LenDocType - 1
        For j = 0 To UnidLen
            ActiveSheet.Cells(i * Ref + j + 3, 2).Select
            Selection.Merge
            Selection = Unidades(j)
            Call SoftBlueBackground
            Call BreakText
        Next j
        ActiveSheet.Cells(i * Ref + j + 3, 2).Select
        Selection = "TOTAL"
        Call ChangeTitleColor
    Next i
    

    ActiveSheet.Range(Cells(1, LenDocType + 4), Cells(1, 2 * LenDocType + 4)).Select
    Selection.Merge
    Selection = "TOTAIS"
    Call ChangeTitleColor
    Call BreakText
    
    For i = 0 To LenDocType - 1
        ActiveSheet.Cells(2, i + LenDocType + 5).Select
        Selection.Merge
        Selection = DocType(i + 1)
        Call SoftBlueBackground
        Call BreakText
    Next i
    
    For i = 0 To UnidLen
        ActiveSheet.Cells(i + 3, LenDocType + 4).Select
        Selection.Merge
        Selection = Unidades(i)
        Call SoftBlueBackground
        Call BreakText
    Next i
        
    ActiveSheet.Cells(UnidLen + 4, LenDocType + 4).Select
    Selection = "TOTAL"
    Call ChangeTitleColor


End Sub

Sub ChangeOrderDocType() 'sub para trocar a ordem de elementos em DocType. Foi feito como sub para controle de uso de memória
    Dim DocTypeAUX As String
    
    Worksheets("Documentos faltantes").Activate
    
    'operação para trocar a posição de 'Despacho Autorizatório' e 'Reserva'
    DocTypeAUX = DocType(4)
    DocType(4) = DocType(5)
    DocType(5) = DocTypeAUX
End Sub

Sub EquationDF()
    Dim i             As Integer
    Dim j             As Integer
    Dim k             As Integer
    Dim LenDocTypeAUX As Integer
    
    LenDocTypeAUX = LenDocType
    
    For i = 0 To LenDocType - 2
        For j = 0 To LenDocTypeAUX - 1
            For k = 0 To UnidLen
                Cells(i * Ref + k + 3, j + 3).Select
                Selection.FormulaR1C1 = "=COUNTIFS('Dados Brutos'!C47,""" & DocType(LenDocType - i) & """,'Dados Brutos'!C[35],""0"",'Dados Brutos'!C4,'Documentos faltantes'!RC2)"
                Call WhiteBackground
                Call BreakText
            Next k
            
            Cells(i * Ref + UnidLen + 4, j + 3).Select
            Selection.FormulaR1C1 = "=SUM(R[-" & UnidLen + 1 & "]C:R[-1]C)"
            Call SoftBlueBackground
            Call BreakText
            
            Cells(i * Ref + 2, 2).Select
            Selection.FormulaR1C1 = "=SUM(R[1]C[1]:R[" & UnidLen + 1 & "]C[" & LenDocTypeAUX & "])"
            Call ChangeTitleColor
            Call BreakText
        Next j
        
        LenDocTypeAUX = LenDocTypeAUX - 1
    Next i
    
    For i = 0 To UnidLen
        Cells((LenDocType - 1) * Ref + i + 3, 3).Select
        Selection.FormulaR1C1 = "=COUNTIFS('Dados Brutos'!C47,""" & DocType(0) & """,'Dados Brutos'!C[35],""0"",'Dados Brutos'!C4,'Documentos faltantes'!RC2)"
        Call WhiteBackground
        Call BreakText
    Next i
    
    'escrever ultima tabela
    Cells((LenDocType - 1) * Ref + UnidLen + 4, 3).Select
    Selection.FormulaR1C1 = "=SUM(R[-" & UnidLen + 1 & "]C:R[-1]C)"
    Call SoftBlueBackground
    Call BreakText
            
    Cells((LenDocType - 1) * Ref + 2, 2).Select
    Selection.FormulaR1C1 = "=SUM(R[1]C[1]:R[" & UnidLen + 1 & "]C[" & LenDocTypeAUX & "])"
    Call ChangeTitleColor
    Call BreakText
    
    'escrever tabela de totais
    For i = 0 To LenDocType - 1
        For j = 0 To UnidLen
            Cells(j + 3, LenDocType + i + 5).Select
            Selection.FormulaR1C1 = "=SUMIF(C2,RC" & LenDocType + 4 & ",C[-" & LenDocType + 2 & "])"
            Call WhiteBackground
            Call BreakText
        Next j
        
        For j = 0 To LenDocType - 1
            Cells(UnidLen + 4, LenDocType + j + 5).Select
            Selection.FormulaR1C1 = "=SUM(R[-" & UnidLen + 1 & "]C:R[-1]C)"
            Call SoftBlueBackground
            Call BreakText
        Next j
            
            Cells(2, LenDocType + 4).Select
            Selection.FormulaR1C1 = "=SUM(R3C" & LenDocType + 5 & ":R" & UnidLen + 3 & "C" & 2 * LenDocType + 4 & ")"
            Call ChangeTitleColor
            Call BreakText
    Next i
End Sub

Sub BordersDF()
    Dim i             As Integer
    Dim j             As Integer
    Dim LenDocTypeAUX As Integer

    LenDocTypeAUX = LenDocType
    
    For i = 0 To LenDocType - 1
        For j = 0 To LenDocTypeAUX - 1
            ActiveSheet.Range(Cells(i * Ref + 1, 2), Cells(i * Ref + UnidLen + 4, LenDocTypeAUX + 2)).Select
            Call CreateBorders
        Next j
        LenDocTypeAUX = LenDocTypeAUX - 1
    Next i
    
    ActiveSheet.Range(Cells(1, LenDocType + 5), Cells(UnidLen + 4, 2 * LenDocType + 4)).Select
    Call CreateBorders
End Sub

'--------------------------------------
Sub BreakText()
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
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

Sub ChangeTitleColor()
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

Sub CreateBorders()
    'Criar grelha das tabelas
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

