VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Utilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database

'Atualizada em 11/10/2016
'Desenvolvida por Gustavo Diniz da Corte (+55 31 998503203)
'gustavodacorte@gmail.com

'Abre um relatório especificado por nome
Public Function AbreRelatorio(Nome As String, Optional Argumentos As String = "")
    DoCmd.OpenReport Nome, acViewReport, , , , Argumentos
End Function

'Fecha um Relatório especificado por nome
Public Function FechaRelatorio(Nome As String)
    DoCmd.Close acReport, Nome, acSaveYes
End Function

'Abre um relatório para edição
Public Function EditaRelatorio(Nome As String)
    DoCmd.OpenReport Nome, acViewDesign
End Function

'Abre um formulário especificado por nome
Public Function AbreFormulario(Nome As String, Optional Argumentos As String = "")
    DoCmd.OpenForm Nome, acNormal, , , , , Argumentos
End Function

'Abre um formulário especificado por nome
Public Function AbreFormularioExclusivo(Nome As String, Optional Argumentos As String = "")
    DoCmd.OpenForm Nome, , , , , acDialog, Argumentos
End Function

'Fecha um formulário especificado por nome
Public Function FechaFormulario(Nome As String)
    DoCmd.Close acForm, Nome, acSaveYes
End Function

'Define consulta de interface
Public Function CriaConsulta(Consulta, NomeConsulta)
    'Define consulta
    If Not IsNull(DLookup("Type", "MSYSObjects", "Name='" + NomeConsulta + "'")) Then
        CurrentDb.QueryDefs(NomeConsulta).SQL = Consulta
    Else
        CurrentDb.CreateQueryDef NomeConsulta, Consulta
    End If
End Function

'Merge os nomes
Public Function NomeCompleto(PrimeiroNome As String, Sobrenome As String)
    NomeCompleto = PrimeiroNome + " " + Sobrenome
End Function

'Executa uma consulta de selecao
Public Function ExecutaBusca(Consulta As String) As Recordset
    Debug.Print Consulta
    Set ExecutaBusca = CurrentDb.OpenRecordset(Consulta)
End Function

'Executa uma consulta de Inserção, Atualização ou Deleção
Public Function ExecutaAcao(Consulta As String)
    Debug.Print Consulta
    CurrentDb.Execute Consulta
End Function

'Pega ultima id adicionada em uma tabela
Public Function PegaUltimaID(tabela) As Long
    Set rs = ExecutaBusca("SELECT @@IDENTITY FROM " + tabela)
    id_sub = rs(0)
    PegaUltimaID = id_sub
End Function

'Troca uma String nulo por um " "
Public Function SeNulo(Valor)
    If IsNulo(Valor) Then
        SeNulo = ""
    Else
        SeNulo = Trim(Valor)
    End If
End Function

'Troca um valor Nulo por 0
Public Function NumeroNulo(Valor) As Double
    If IsNulo(Valor) Then
        Valor = 0
    End If
    
    NumeroNulo = Valor
End Function

'Troca um valor 0 por 1 (Para evitar divisões por zero
Public Function SeZero(Valor)
    NumeroNulo Valor
    If Valor = 0 Then
        SeZero = 1
    Else
        SeZero = Valor
    End If
End Function

'Calcula a idade com base em uma data
Public Function CalculaIdade(dataAlvo As Date) As Double
    CalculaIdade = (Year(Date) - Year(dataAlvo)) + (Month(Date) - Month(dataAlvo)) / 12 + (Day(Date) - Day(dataAlvo)) / 365
End Function

'Converte formato de data para formato de String numérica
Public Function ConvDataToStr(data) As String
    If IsDate(data) Then
        ConvDataToStr = Format(data, "yyyymmdd")
    Else
        ConvDataToStr = data
    End If
End Function

'Converte formato de formato numérico para data
Public Function ConvDataFromStr(strData As String, Optional DataBrasileira As Boolean = False) As String
    On Error GoTo ERR1:
    'Declara Variáveis
    Dim final As String
    Dim ano As Integer, mes As Integer, dia As Integer

    ano = CInt((Mid(strData, 1, 4)))
    mes = CInt((Mid(strData, 5, 2)))
    dia = CInt((Mid(strData, 7, 2)))
    
    If Not DataBrasileira Then 'Formato de data conforme linguagem padrão do computador
        final = CStr(DateSerial(ano, mes, dia))
    Else 'Formato de data brasileira
        final = dia & "/" & mes & "/" & ano
    End If
    
    ConvDataFromStr = final
    Exit Function
    
'Condição de Erro
ERR1:
    ConvDataFromStr = ""
    Exit Function
End Function

'Converte formato de data para formato numérico
Public Function ConvHoraToStr(hora As Date) As String
    ConvHoraToStr = Format(hora, "hh:mm")
End Function

'Converte formato de formato numérico para data
Public Function ConvHoraFromStr(strHora As String) As Date
    'Declara Variáveis
    Dim final As Date
    Dim hora As Integer, minuto As Integer

    hora = Mid(strHora, 1, 2)
    minuto = Mid(strHora, 4, 2)
    final = TimeSerial(hora, minuto, 0)

    ConvHoraFromStr = final
End Function

'Verifica se o objeto contem nulo ou string vazia
Public Function IsNulo(objeto) As Boolean
    On Error GoTo ERR1:
    If (objeto = "") Or IsNull(objeto) Or objeto = vbNullString Then
        IsNulo = True
        Exit Function
    End If

    If Len(Trim(objeto)) = 0 Then
        IsNulo = True
        Exit Function
    End If

    IsNulo = False
    Exit Function

'Condição de Erro
ERR1:
    IsNulo = True
    Exit Function
End Function

'Método de inserção de objetos
Public Function SQLInsere(TabelaDestino As String, Campos As String, Valores As String, IDAdicionada As Long) As Boolean
    'Define tratamento de erro
    On Error GoTo ERR1
    Dim Consulta As String
    
    Consulta = "INSERT INTO " + TabelaDestino + "(" + Campos + ") VALUES (" + Valores + ")"
    ExecutaAcao (Consulta)
    
    IDAdicionada = PegaUltimaID(TabelaDestino)
    
    SQLInsere = True
    Exit Function
    
'Condição de Erro
ERR1:
    SQLInsere = False
    Exit Function
End Function

'Método de Edição de objetos
Public Function SQLAtualiza(TabelaDestino As String, CamposAlterados As String, Restricao As String) As Boolean
    'Define tratamento de erro
    On Error GoTo ERR1
    Dim Consulta As String
    
    Consulta = "UPDATE " + TabelaDestino + " SET " + CamposAlterados + " WHERE " + Restricao
    ExecutaAcao Consulta
    
    SQLAtualiza = True
    Exit Function
    
'Condição de Erro
ERR1:
    SQLAtualiza = False
    Exit Function
End Function

'Método de Deleção de objetos
Public Function SQLDeleta(TabelaDestino As String, Optional Restricao As String = "Todos") As Boolean
    'Define tratamento de erro
    On Error GoTo ERR1
    Dim Consulta As String
    
    If Restricao = "Todos" Then
         Consulta = "DELETE FROM " + TabelaDestino
    Else
        Consulta = "DELETE FROM " + TabelaDestino + " WHERE " + Restricao
    End If
    ExecutaAcao Consulta
    
    SQLDeleta = True
    Exit Function
'Tratamento de erros
ERR1:
    SQLDeleta = False
    Exit Function
End Function

'Limpa os valores de um Listbox
Public Function LimpaLista(lista As ListBox)
    lista.RowSourceType = "Value List"
    Dim Contador As Long
    Dim auxiliarFor As Long
    
    Contador = lista.ListCount
    For auxiliarFor = 1 To Contador Step 1
        lista.RemoveItem 0
    Next auxiliarFor
End Function

'Limpa os valores de um Combobox
Public Function LimpaCombinacao(lista As ComboBox)
    lista.RowSourceType = "Value List"
    Dim Contador As Long
    Dim auxiliarFor As Long
    
    Contador = lista.ListCount
    For auxiliarFor = 1 To Contador Step 1
        lista.RemoveItem 0
    Next auxiliarFor
End Function

'Método que realiza a conversão de um número para seu extenso correspondente
Public Function ConvNumeroToExtenso(Valor As Double)
    'Declara Variáveis
    Dim ParteInteira As Long
    Dim ValorStr As String
    Dim ValorInt As String
    Dim unidade, dezena, centena, milhar, outros, diferenciacao
    Dim extenso As String
    
    If CDbl(Valor) = 0 Then
        ConvNumeroToExtenso = "Zero Reais"
        Exit Function
    End If
    
    'Ajusta o valor de parte inteira para execução do método
    ParteInteira = Len(CStr(Format(Valor, ".00"))) - 3
    If ((ParteInteira Mod 3) <> 0) Then
        ParteInteira = ParteInteira + 3 - (ParteInteira Mod 3)
    End If
    
    extenso = ""
    Negativo = (Valor < 0)
    ValorStr = Format(Abs(Valor), String(ParteInteira, "0") + ".00")
    
    'Prepara formatação de letras
    unidade = Array("0", "Um", "Dois", "Tres", "Quatro", "Cinco", "Seis", "Sete", "Oito", "Nove", "Dez", "Onze", "Doze", "Treze", "Quatorze" _
    , "Quinze", "Dezesseis", "Dezessete", "Dezoito", "Dezenove")
    dezena = Array("0", "0", "Vinte", "Trinta", "Quarenta", "Cinqueta", "Sessenta", "Setenta", "Oitenta", "Noventa")
    centena = Array("0", "Cento", "Duzentos", "Trezentos", "Quatrocentos", "Quinhentos", "Seiscentos", "Setecentos", "Oitocentos", "Novecentos")
    milhar = "Mil"
    outros = Array("Mi", "Bi", "Tri", "Quadri", "Quinti")
    diferenciacao = Array("0", "lhão", "lhões", "Cem", "Real", "Reais", "Centavo", "Centavos")
    
    'Analiza parte inteira
    ValorInt = Left(ValorStr, ParteInteira)
    i = 1
    
    'Executa a função através de agrupamentos de três algarismos para a parte inteira do número
    While i <= (ParteInteira / 3)
        'Se é o primeiro agrupamento, não coloca quantificador no valor extenso
        If (i = 1) Then
            pedaco = CInt(Mid(ValorInt, ParteInteira - (3 * i - 1), 3))
        Else: pedaco = CInt(Mid(ValorInt, ParteInteira - (3 * i - 1), 3))
        End If
        'Se é o segundo agrupamento, adiciona quantificador de milhar no valor extenso
        If (i = 2) And (pedaco <> 0) Then
            extenso = " " + milhar + " e " + extenso
        End If
        'Se é um agrupamento superior, coloca o quantificador correspondente no valor extenso
        If (i >= 3) And (pedaco <> 0) Then
            If pedaco = 1 Then
                extenso = outros(i - 3) + diferenciacao(1) + ", " + extenso
            Else: extenso = outros(i - 3) + diferenciacao(2) + ", " + extenso
            End If
            
        End If
        'Atualiza o valor extenso com os números, em extenso, do agrupamento correspondente
        extenso = ConvToExtensoAnalisaTres(Mid(ValorInt, ParteInteira - (3 * i - 1), 3), 3) + extenso
        
        i = i + 1
    Wend
    
    'Define se real deve estar no plural, Caso valor seja igual ou superior a R$1,00
    If (ParteInteira <> 0) Then
        If CLng(ValorInt) = 1 Then
            extenso = extenso + " " + diferenciacao(4)
        Else: extenso = extenso + " " + diferenciacao(5)
        End If
    End If
    
    'Pega os centavos do número
    cents = Right(ValorStr, 2)
    'Se centavos for diferente de zero, analisa os centavos como um número independente
    If (CLng(cents) <> 0) Then
        'Acrescenta um zero à esquerda do número e cria seu extenso
        cents = Format(cents, "000")
        If (ParteInteira <> 0) Then
            extenso = extenso + " e " + ConvToExtensoAnalisaTres(cents, 3)
        Else
            extenso = extenso + ConvToExtensoAnalisaTres(cents, 3)
        End If
        
        'Define se será utilizado centavo ou centavos
        If CLng(cents) = 1 Then
            extenso = extenso + " " + diferenciacao(6)
        Else: extenso = extenso + " " + diferenciacao(7)
        End If
    End If
    
    ConvNumeroToExtenso = extenso
    Debug.Print extenso
End Function

'Pega tês algarismos e define seu extenso
Public Function ConvToExtensoAnalisaTres(ValorInt, tamanho) As String
    'Prepara formatação de letras
    unidade = Array("0", "Um", "Dois", "Tres", "Quatro", "Cinco", "Seis", "Sete", "Oito", "Nove", "Dez", "Onze", "Doze", "Treze", "Quatorze" _
    , "Quinze", "Dezesseis", "Dezessete", "Dezoito", "Dezenove")
    dezena = Array("0", "0", "Vinte", "Trinta", "Quarenta", "Cinqueta", "Sessenta", "Setenta", "Oitenta", "Noventa")
    centena = Array("0", "Cento", "Duzentos", "Trezentos", "Quatrocentos", "Quinhentos", "Seiscentos", "Setecentos", "Oitocentos", "Novecentos")
    milhar = "Mil"
    outros = Array("0", "Mi", "Bi", "Tri", "Quadri", "Quinti")
    diferenciacao = Array("0", "lhão", "lhões", "Cem", "Real", "Reais", "Centavo", "Centavos")

    i = 1
    
    'Enquanto não analisou todos os tres algarismos:
    While i < 3
        
        'Se é a primeira iteração, pega os dois primeiros algarismos, senão pega o terceiro
        If (i = 1) Then
            pedaco = CInt(Mid(ValorInt, tamanho - i, 2))
        Else: pedaco = CInt(Mid(ValorInt, tamanho - i, 1))
        End If
        
        'Se é a terceira iteração, adiciona o extenso correspondente
        If (i = 2) And (pedaco <> 0) Then
            extenso = centena(pedaco) + " e " + extenso
        End If
        
        'Se é a primeira iteração, processa os dois primeiros algarismos
        If (i = 1) Then
            'Se o número é diferente de 0, processa seu extenso
            If (pedaco <> 0) Then
                'Se estiver entre 1~19, adiciona o extenso correspondente
                If (pedaco < 20) Then
                    extenso = extenso + unidade(pedaco)
                Else
                    p1 = Int(pedaco / 10)
                    p2 = pedaco - p1 * 10
                    'Se o primeiro algarismo for 0, adiciona somente o quantificador, senão, adiciona o extenso correspondente
                    If p2 = 0 Then
                        extenso = extenso + dezena(p1)
                    Else: extenso = extenso + dezena(p1) + " e " + unidade(p2)
                    End If
                End If
            Else 'Se o número = 0, analisa o terceiro algarismo
                i = i + 1
                pedaco = CInt(Mid(ValorInt, tamanho - i, 3))
                'Diferencia cem de cento
                If pedaco = 100 Then
                    extenso = extenso + diferenciacao(3)
                Else
                    pedaco = Int(pedaco / 100)
                    extenso = centena(pedaco) + " " + extenso
                End If
            End If
        End If
        
        i = i + 1
    Wend
    'Retorna o valor
    ConvToExtensoAnalisaTres = extenso
    
End Function

'Retorna um valor abreviado para o mês em formato brasileiro
Public Function MesAbreviado(Referencia As Integer) As String
    Select Case Referencia
        Case 1: MesAbreviado = "Jan"
        Case 2: MesAbreviado = "Fev"
        Case 3: MesAbreviado = "Mar"
        Case 4: MesAbreviado = "Abr"
        Case 5: MesAbreviado = "Mai"
        Case 6: MesAbreviado = "Jun"
        Case 7: MesAbreviado = "Jul"
        Case 8: MesAbreviado = "Ago"
        Case 9: MesAbreviado = "Set"
        Case 10: MesAbreviado = "Out"
        Case 11: MesAbreviado = "Nov"
        Case 12: MesAbreviado = "Dez"
        Case Else: MesAbreviado = "NA"
    End Select
End Function

'Retorna o nome completo de um mes de referência na língua portuguesa
Public Function MesCompleto(Referencia As Integer) As String
    Select Case Referencia
        Case 1: MesCompleto = "Janeiro"
        Case 2: MesCompleto = "Fevereiro"
        Case 3: MesCompleto = "Março"
        Case 4: MesCompleto = "Abril"
        Case 5: MesCompleto = "Maio"
        Case 6: MesCompleto = "Junnho"
        Case 7: MesCompleto = "Julho"
        Case 8: MesCompleto = "Agosto"
        Case 9: MesCompleto = "Setembro"
        Case 10: MesCompleto = "Outubro"
        Case 11: MesCompleto = "Novembro"
        Case 12: MesCompleto = "Dezembro"
        Case Else: MesCompleto = "NA"
    End Select
End Function

'Cria uma data no formato: '00' de XXXXXXXX de '0000'
Public Function DataRelatorio(arg01 As Date)
    Dim mes As Integer
    
    dia = Format(arg01, "dd")
    mes = CInt(Format(arg01, "mm"))
    ano = Format(arg01, "yyyy")
    
    'Coloca o dia em numero de 2 digitos
    DataExtenso = dia + " de "
    
    'Coloca o mes em extenso
    DataExtenso = DataExtenso + MesCompleto(mes)
    
    'coloca o ano em numero de 4 digitos
    DataExtenso = DataExtenso + " de " + ano
    
    'retorna o resultado
    DataRelatorio = DataExtenso
End Function

'Cria no HD um arquivo em formato PDF com a janela ativa
Public Function CriaPDF(NomeArquivo As String, Categoria, NomeObjeto As String) As Boolean
    'Tratamento de Erro
    On Error GoTo ERR1
    Dim CaminhoArquivo As String, Caminho As String
    
    Caminho = Application.CurrentProject.Path
    CaminhoArquivo = Caminho & "\" & NomeArquivo & ".pdf"
    Debug.Print CaminhoArquivo
    DoCmd.OutputTo Categoria, NomeObjeto, acFormatPDF, CaminhoArquivo, False
    Debug.Print CaminhoArquivo
    
    Exit Function
'Condição de Erro
ERR1:
    'MsgBox "Erro ao salvar PDF,caminho especificado não existe ou um arquivo de mesmo nome está aberto. Arquivo não Criado", , "ERRO"
    Exit Function
End Function

'Abre um arquivo em formato PDF
Public Function AbrePDF(NomeArquivo As String, Categoria, NomeObjeto As String) As Boolean
    'Tratamento de Erro
    On Error GoTo ERR1
    Dim CaminhoArquivo As String, Caminho As String
    
    Caminho = Application.CurrentProject.Path
    CaminhoArquivo = Caminho & "\" & NomeArquivo & ".pdf"
    Application.FollowHyperlink CaminhoArquivo
    Exit Function
'Condição de Erro
ERR1:
    MsgBox "Erro ao Abrir PDF,caminho especificado não existe", , "ERRO"
    Exit Function
End Function

'Salva uma página específica em PDF e abre
Public Function CriaAbrePDF(NomeArquivo As String, Categoria, NomeObjeto As String)
    If Not CriaPDF(NomeArquivo, Categoria, NomeObjeto) Then
        Exit Function
    End If
    
    If Not AbrePDF(NomeArquivo, Categoria, NomeObjeto) Then
        Exit Function
    End If
End Function

'Redimensiona uma janela horizontalmente para manter todos os controles centralizados
Public Function RedimensionaHorizontal(LarguraOriginal As Double, LarguraAtual As Double, NovaLargura As Double, NomeForm As String) As Double
    On Error GoTo ERR1
    'Declara Variáveis
    Dim Delta As Double
    Dim conta As Long
    
    If NovaLargura < LarguraOriginal Then
        NovaLargura = LarguraOriginal
    End If
    
    'Define componente de ajuste
    Delta = (NovaLargura - LarguraAtual) / 2
    
    'Define numero de componentes do formulário
    conta = Forms(NomeForm).Count - 1
    
    'Ajusta os controles
    For i = 0 To conta Step 1
        Forms(NomeForm).Controls(i).Left = Forms(NomeForm).Controls(i).Left + Delta
    Next i
    
    RedimensionaHorizontal = NovaLargura
    Exit Function
ERR1: 'Condição de erro
    RedimensionaHorizontal = LarguraAtual
    MsgBox "Impossível ajustar a janela nos limites desejados", , "ERRO"
    Exit Function
End Function
                                        
'Redimensiona uma janela Verticalmente para manter todos os controles centralizados
Public Function RedimensionaVertical(AlturaOriginal As Double, AlturaAtual As Double, NovaAltura As Double, NomeForm As String) As Double
    On Error GoTo ERR1
    'Declara Variáveis
    Dim Delta As Double
    Dim conta As Long
    
    'Verifica se a nova altura é compatível sem comprometer a estrutura original do formulário
    If (AlturaOriginal > NovaAltura) Then
        RedimensionaVertical = AlturaOriginal
        Exit Function
    End If
    
    'Calcula variações e componentes a serem ajustados
    Delta = Fix((NovaAltura - AlturaAtual) / 2)
    contaDetalhe = NumeroControles(NomeForm, acDetail) - 1
    contaCabecalho = NumeroControles(NomeForm, acHeader) - 1
    
    'Define se será ajustado o cabeçalho ou a seção de detalhes
    If (SecaoExiste(NomeForm, acHeader) And (contaCabecalho > -1)) Then 'Se existir cabeçalho e existirem controles nele
        'Define a ordenação dos ajustes
        If Delta < 0 Then 'Se delta é negativo
            For i = 0 To contaCabecalho Step 1
                Forms(NomeForm).Controls(i).Top = Forms(NomeForm).Controls(i).Top + Delta
            Next i
            Forms(NomeForm).Section(acHeader).Height = Forms(NomeForm).Section(acHeader).Height + Delta
        Else 'se delta é positivo
            Forms(NomeForm).Section(acHeader).Height = Forms(NomeForm).Section(acHeader).Height + Delta
            For i = 0 To contaCabecalho Step 1
                Forms(NomeForm).Controls(i).Top = Forms(NomeForm).Controls(i).Top + Delta
            Next i
        End If
    Else 'Se não existir cabeçalho ou se não houver controles nele
        For i = 0 To contaDetalhe Step 1 'Ajusta os controles
            Forms(NomeForm).Controls(i).Top = Forms(NomeForm).Controls(i).Top + Delta
        Next i
    End If
    
    RedimensionaVertical = NovaAltura
    Exit Function
ERR1: 'Caso exista erro
    RedimensionaVertical = AlturaAtual
    MsgBox "Impossível ajustar a janela nos limites desejados", , "ERRO"
    Exit Function
End Function
                                          
'Redimensiona uma janela Verticalmente para manter todos os controles centralizados
Public Function RedimensionaDetalheVertical(AlturaOriginal As Double, AlturaAtual As Double, NovaAltura As Double, NomeForm As String) As Double
    On Error GoTo ERR1
    'Declara Variáveis
    Dim Delta As Double
    
    'Verifica se a nova altura é compatível sem comprometer a estrutura original do formulário
    If (AlturaOriginal > NovaAltura) Then
        RedimensionaDetalheVertical = AlturaOriginal
        Exit Function
    End If
    
    'Calcula variações e componentes a serem ajustados
    Delta = Fix((NovaAltura - AlturaAtual) / 2)
    contaDetalhe = NumeroControles(NomeForm, acDetail) - 1
    
    If Delta < 0 Then 'Se delta é negativo
        For i = 0 To contaDetalhe Step 1 'Ajusta os controles
            Forms(NomeForm).Section(acDetail).Controls(i).Top = Forms(NomeForm).Section(acDetail).Controls(i).Top + Delta
        Next i
        Forms(NomeForm).Section(acDetail).Height = Forms(NomeForm).Section(acDetail).Height + Delta
    Else 'Se delta é positivo
        Forms(NomeForm).Section(acDetail).Height = Forms(NomeForm).Section(acDetail).Height + Delta
        For i = 0 To contaDetalhe Step 1 'Ajusta os controles
            Forms(NomeForm).Section(acDetail).Controls(i).Top = Forms(NomeForm).Section(acDetail).Controls(i).Top + Delta
        Next i
    End If
    
    RedimensionaDetalheVertical = NovaAltura
    Exit Function
ERR1: 'Caso exista erro
    RedimensionaDetalheVertical = AlturaAtual
    MsgBox "Impossível ajustar a janela nos limites desejados", , "ERRO"
    Exit Function
End Function
                                            
'Verifica se a seção do formulário existe
Public Function SecaoExiste(FormName As String, TipoSecao As Long) As Boolean
    On Error GoTo ERR1
    Dim Nome As String
    
    Nome = Forms(FormName).Section(TipoSecao).Name
    SecaoExiste = True
    Exit Function
ERR1:
    SecaoExiste = False
    Exit Function
End Function

'Calcula numero de controles em determinada seção do formulário
Public Function NumeroControles(FormName As String, TipoSecao As Long) As Long
    On Error GoTo ERR1
    Dim Nome As String
    
    NumeroControles = Forms(FormName).Section(TipoSecao).Controls.Count
    Exit Function
ERR1:
    NumeroControles = 0
    Exit Function
End Function
