Public UnidLen      As Integer 'tamanho do vetor de 'Unidades'. Pode ser usado UBound
Public Unidades     As Variant 'vetor 'Unidades'
Public UnidOfLen    As Integer 'tamanho do vetor de 'Unidades' oficiais. Pode ser usado UBound
Public UnidadesOf   As Variant 'vetor 'Unidades' oficiais

Sub Listagem()
    Worksheets("SEI").Activate
    Call OrganizarOrgaos
    Call OrgaoArray
    Call OrgaoOfArray
    Call EscreverUnid
    Call FormatarCond
    
    ActiveSheet.Cells(1, 1).Select
End Sub

Sub OrganizarOrgaos()
    Worksheets("SEI").Activate
    ActiveWorkbook.Worksheets("SEI").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SEI").AutoFilter.Sort.SortFields.Add Key:=Range( _
        "D1:D15596"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("SEI").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub OrgaoArray()
    Dim i    As Integer
    Dim NRow As Integer
    Dim NCol As Integer
    
    NRow = 2
    NCol = 4
    UnidLen = 0
    Unidades = Array("")
    
    'laço para criar uma array com o nome de todas as unidades disponíveis
    Worksheets("SEI").Activate
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
    If Unidades(i) = "" Then 'após o final do laço, o acaba inserindo um valor nulo na última posição do vetor.
        i = i - 1
        UnidLen = i 'equação que irá definir o tamanho do vetor
        ReDim Preserve Unidades(0 To i)
    End If
End Sub

Sub OrgaoOfArray()
    Dim i    As Integer
    Dim NRow As Integer
    Dim NCol As Integer
    
    NRow = 2
    NCol = 1
    UnidOfLen = 0
    UnidadesOf = Array("")
    
    'laço para criar uma array com o nome de todas as unidades disponíveis
    Worksheets("Lista de Órgãos").Activate
    Do While Cells(NRow, NCol).Value <> "" 'laço que se repete até o conteúdo da célula ser nulo
        UnidadesOf(i) = Cells(NRow, NCol).Value 'copia o valor da célula para o array
        If i = 0 Then 'na primeira posição do vetor sempre é escrita uma unidade
            i = i + 1
            UnidOfLen = i 'equação que irá definir o tamanho do vetor. Pode ser usado UBound após ter todo o vetor. Foi feito dessa forma para fins de aprendizado.
                        'Contudo, pode ser mais interessante redefinir o código usando UBound, pois para cada nova unidade uma operação deve ser feita, consumindo tempo de máquina
            ReDim Preserve UnidadesOf(0 To i) 'Array começa com 0 sempre. Preserva os valores e redimensiona o vetor
        ElseIf UnidadesOf(i) <> UnidadesOf(i - 1) Then 'se o conteúdo do vetor na posição i for diferente do conteúdo do vetor na sua posição anterior (i-1), _
                                                   'então é escrito um novo conteúdo
            i = i + 1
            UnidOfLen = i 'equação que irá definir o tamanho do vetor
            ReDim Preserve UnidadesOf(0 To i) 'Array começa com 0 sempre
        End If 'se o conteúdo for igual, então o programa ignora. Pode ser interessante adicionar um condicional cobrindo essa possibilidade para evitar alguma má execução
        NRow = NRow + 1
    Loop
    If UnidadesOf(i) = UnidadesOf(i - 1) Then 'se i conteúdo de i for igual ao conteúdo de i-1, então ele apaga a última posição. Isso deve ser feito pois o programa estava duplicando a última unidade
        i = i - 1
        UnidOfLen = i 'equação que irá definir o tamanho do vetor
        ReDim Preserve UnidadesOf(0 To i) 'redimensiona o vetor para o tamanho que não inclua a unidade repetida
    End If
    If UnidadesOf(i) = "" Then 'após o final do laço, o acaba inserindo um valor nulo na última posição do vetor.
        i = i - 1
        UnidOfLen = i 'equação que irá definir o tamanho do vetor
        ReDim Preserve UnidadesOf(0 To i)
    End If
End Sub

Sub EscreverUnid()
    'a primeira coluna na planilha representa as unidades que foram digitadas pelos usuários
    'a segunda coluna representa a comparação com as unidades registradas na guia "Lista de Órgãos"
    'a terceira coluna representa a correção usando IF de forma não otimizada
    'caso não haja nenhuma unidades nas segunda e terceira colunas, então não há nenhuma unidade
    'que corresponda ao que foi posto.
    Dim i         As Integer
    Dim j         As Integer
    Dim UnidAUX   As String
    Dim UnidOfAUX As String

    Worksheets("Lista de Alterações").Activate
    Cells.Select
    Selection.Clear
    ActiveWindow.FreezePanes = False
    For i = 0 To UnidLen
        ActiveSheet.Cells(i + 1, 1) = Unidades(i)
        For j = 0 To UnidOfLen
            If Unidades(i) = UnidadesOf(j) Then 'condição de unidades iguais
                ActiveSheet.Cells(i + 1, 2) = UnidadesOf(j)
                Exit For
            End If
            
            If Unidades(i) <> UnidadesOf(j) Then 'condição de unidades diferentes
                UnidAUX = UCase(Unidades(i))
                UnidAUX = Replace(UnidAUX, " ", "")
                UnidAUX = Replace(UnidAUX, "/", "")
                UnidAUX = Replace(UnidAUX, "-", "")
                UnidOfAUX = UCase(UnidadesOf(j))
                UnidOfAUX = Replace(UnidOfAUX, " ", "")
                UnidOfAUX = Replace(UnidOfAUX, "/", "")
                UnidOfAUX = Replace(UnidOfAUX, "-", "")
                
                If UnidAUX <> UnidOfAUX Then
                    If UnidAUX = "COHAB" Then
                        UnidAUX = "COHABSP"
                    End If
                    
                    If UnidAUX = "DRE" Then
                        UnidAUX = "SME"
                    End If
                    
                    If UnidAUX = "PRODAM" Then
                        UnidAUX = "PRODAMSP"
                    End If
                    
                    If UnidAUX = "SMEDRECS" Then
                        UnidAUX = "SME"
                    End If
                    
                    If UnidAUX = "SMGSIURB" Then
                        UnidAUX = "SMG"
                    End If
                    
                    If UnidAUX = "SMGHSPM" Then
                        UnidAUX = "SMG"
                    End If
                    
                    If UnidAUX = "SMSAHM" Then
                        UnidAUX = "SMS"
                    End If
                    
                    If UnidAUX = "SMSCRSS" Then
                        UnidAUX = "SMS"
                    End If
                    
                    If UnidAUX = "SMSHSPM" Then
                        UnidAUX = "SMS"
                    End If
                    
                    If UnidAUX = "SMSCRS" Then
                        UnidAUX = "SMS"
                    End If
                    
                    If UnidAUX = "SMSCRSLESTE" Then
                        UnidAUX = "SMS"
                    End If
                    
                    If UnidAUX = "SMSG" Then
                        UnidAUX = "SMS"
                    End If
                    
                    If UnidAUX = "SMSPSPPR" Then
                        UnidAUX = "SPPR"
                    End If
                    
                    If UnidAUX = "SPARICANDUVA" Then
                        UnidAUX = "SPAF"
                    End If
                    
                    If UnidAUX = "SPCAMPOLIMPO" Then
                        UnidAUX = "SPCL"
                    End If
                    
                    If UnidAUX = "SPCASAVERDE" Then
                        UnidAUX = "SPCV"
                    End If
                    
                    If UnidAUX = "SPCIDADEADEMAR" Then
                        UnidAUX = "SPAD"
                    End If
                    
                    If UnidAUX = "SPCIDADETIRADENTES" Then
                        UnidAUX = "SPCT"
                    End If
                    
                    If UnidAUX = "SPERMELINOMATARAZZO" Then
                        UnidAUX = "SPEM"
                    End If
                    
                    If UnidAUX = "SPGUAIANASES" Then
                        UnidAUX = "SPG"
                    End If
                    
                    If UnidAUX = "SPITAIMPAULISTA" Then
                        UnidAUX = "SPIT"
                    End If
                    
                    If UnidAUX = "SPITAQUERA" Then
                        UnidAUX = "SPIQ"
                    End If
                    
                    If UnidAUX = "SPJABAQUARA" Then
                        UnidAUX = "SPJA"
                    End If
                    
                    If UnidAUX = "SPJAÇANÃTREMEMBÉ" Then
                        UnidAUX = "SPJT"
                    End If
                    
                    If UnidAUX = "SPLAPA" Then
                        UnidAUX = "SPLA"
                    End If
                    
                    If UnidAUX = "SPM'BOIMIRIM" Then
                        UnidAUX = "SPMB"
                    End If
                    
                    If UnidAUX = "SPPARELHEIROS" Then
                        UnidAUX = "SPPA"
                    End If
                    
                    If UnidAUX = "SPPIRITUBAJARAGUÁ" Then
                        UnidAUX = "SPPJ"
                    End If
                    
                    If UnidAUX = "SPSANTOAMARO" Then
                        UnidAUX = "SPSA"
                    End If
                    
                    If UnidAUX = "SPSÃOMIGUELPAULISTA" Then
                        UnidAUX = "SPMP"
                    End If
                    
                    If UnidAUX = "SPSAPOPEMBA" Then
                        UnidAUX = "SPSB"
                    End If
                    
                    If UnidAUX = "SPSÉ" Then
                        UnidAUX = "SPSE"
                    End If
                    
                    If UnidAUX = "SPURBAN." Then
                        UnidAUX = "SPURBANISMO"
                    End If
                    
                    If UnidAUX = "SPVILAMARIAVILAGUILHERME" Then
                        UnidAUX = "SPMG"
                    End If
                    
                    If UnidAUX = "SPVILAMARIANA" Then
                        UnidAUX = "SPVM"
                    End If
                    
                    If UnidAUX = "SPAF(SMSP)" Then
                        UnidAUX = "SPAF"
                    End If
                End If
                
                If UnidAUX = UnidOfAUX Then
                    ActiveSheet.Cells(i + 1, 2) = UnidadesOf(j)
                End If
    
            End If
        Next j
    Next i
End Sub

Sub FormatarCond()
    Range(Cells(1, 2), Cells(UnidLen + 1, 2)).Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=NÚM.CARACT(ARRUMAR(B1))=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .Pattern = xlSolid
        .PatternColor = 255
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub

