VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database

'Abre um relat�rio especificado por nome
Public Function AbreRelatorio(Nome As String, Optional Argumentos As String = "")
    DoCmd.OpenReport Nome, acViewReport, , , , Argumentos
End Function

'Fecha um Relat�rio especificado por nome
Public Function FechaRelatorio(Nome As String)
    DoCmd.Close acReport, Nome, acSaveYes
End Function

'Abre um relat�rio para edi��o
Public Function EditaRelatorio(Nome As String)
    DoCmd.OpenReport Nome, acViewDesign
End Function

'Abre um formul�rio especificado por nome
Public Function AbreFormulario(Nome As String, Optional Argumentos As String = "")
    DoCmd.OpenForm Nome, acNormal, , , , , Argumentos
End Function

'Abre um formul�rio especificado por nome
Public Function AbreFormularioExclusivo(Nome As String, Optional Argumentos As String = "")
    DoCmd.OpenForm Nome, , , , , acDialog, Argumentos
End Function

'Fecha um formul�rio especificado por nome
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

'Executa uma consulta de Inser��o, Atualiza��o ou Dele��o
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

'Troca um valor 0 por 1 (Para evitar divis�es por zero
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

'Converte formato de data para formato de String num�rica
Public Function ConvDataToStr(data) As String
    If IsDate(data) Then
        ConvDataToStr = Format(data, "yyyymmdd")
    Else
        ConvDataToStr = data
    End If
End Function

'Converte formato de formato num�rico para data
Public Function ConvDataFromStr(strData As String, Optional DataBrasileira As Boolean = False) As String
    On Error GoTo ERR1:
    'Declara Vari�veis
    Dim final As String
    Dim ano As Integer, Mes As Integer, dia As Integer

    ano = CInt((Mid(strData, 1, 4)))
    Mes = CInt((Mid(strData, 5, 2)))
    dia = CInt((Mid(strData, 7, 2)))
    
    If Not DataBrasileira Then 'Formato de data conforme linguagem padr�o do computador
        final = CStr(DateSerial(ano, Mes, dia))
    Else 'Formato de data brasileira
        final = dia & "/" & Mes & "/" & ano
    End If
    
    ConvDataFromStr = final
    Exit Function
    
'Condi��o de Erro
ERR1:
    ConvDataFromStr = ""
    Exit Function
End Function

'Converte formato de data para formato num�rico
Public Function ConvHoraToStr(hora As Date) As String
    ConvHoraToStr = Format(hora, "hh:mm")
End Function

'Converte formato de formato num�rico para data
Public Function ConvHoraFromStr(strHora As String) As Date
    'Declara Vari�veis
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

'Condi��o de Erro
ERR1:
    IsNulo = True
    Exit Function
End Function

'M�todo de inser��o de objetos
Public Function SQLInsere(TabelaDestino As String, Campos As String, Valores As String, IDAdicionada As Long) As Boolean
    'Define tratamento de erro
    On Error GoTo ERR1
    Dim Consulta As String
    
    Consulta = "INSERT INTO " + TabelaDestino + "(" + Campos + ") VALUES (" + Valores + ")"
    ExecutaAcao (Consulta)
    
    IDAdicionada = PegaUltimaID(TabelaDestino)
    
    SQLInsere = True
    Exit Function
    
'Condi��o de Erro
ERR1:
    SQLInsere = False
    Exit Function
End Function

'M�todo de Edi��o de objetos
Public Function SQLAtualiza(TabelaDestino As String, CamposAlterados As String, Restricao As String) As Boolean
    'Define tratamento de erro
    On Error GoTo ERR1
    Dim Consulta As String
    
    Consulta = "UPDATE " + TabelaDestino + " SET " + CamposAlterados + " WHERE " + Restricao
    ExecutaAcao Consulta
    
    SQLAtualiza = True
    Exit Function
    
'Condi��o de Erro
ERR1:
    SQLAtualiza = False
    Exit Function
End Function

'M�todo de Dele��o de objetos
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

'M�todo que realiza a convers�o de um n�mero para seu extenso correspondente
Public Function ConvNumeroToExtenso(Valor As Double)
    'Declara Vari�veis
    Dim ParteInteira As Long
    Dim ValorStr As String
    Dim ValorInt As String
    Dim unidade, dezena, centena, milhar, outros, diferenciacao
    Dim extenso As String
    
    If CDbl(Valor) = 0 Then
        ConvNumeroToExtenso = "Zero Reais"
        Exit Function
    End If
    
    'Ajusta o valor de parte inteira para execu��o do m�todo
    ParteInteira = Len(CStr(Format(Valor, ".00"))) - 3
    If ((ParteInteira Mod 3) <> 0) Then
        ParteInteira = ParteInteira + 3 - (ParteInteira Mod 3)
    End If
    
    extenso = ""
    Negativo = (Valor < 0)
    ValorStr = Format(Abs(Valor), String(ParteInteira, "0") + ".00")
    
    'Prepara formata��o de letras
    unidade = Array("0", "Um", "Dois", "Tres", "Quatro", "Cinco", "Seis", "Sete", "Oito", "Nove", "Dez", "Onze", "Doze", "Treze", "Quatorze" _
    , "Quinze", "Dezesseis", "Dezessete", "Dezoito", "Dezenove")
    dezena = Array("0", "0", "Vinte", "Trinta", "Quarenta", "Cinqueta", "Sessenta", "Setenta", "Oitenta", "Noventa")
    centena = Array("0", "Cento", "Duzentos", "Trezentos", "Quatrocentos", "Quinhentos", "Seiscentos", "Setecentos", "Oitocentos", "Novecentos")
    milhar = "Mil"
    outros = Array("Mi", "Bi", "Tri", "Quadri", "Quinti")
    diferenciacao = Array("0", "lh�o", "lh�es", "Cem", "Real", "Reais", "Centavo", "Centavos")
    
    'Analiza parte inteira
    ValorInt = Left(ValorStr, ParteInteira)
    i = 1
    
    'Executa a fun��o atrav�s de agrupamentos de tr�s algarismos para a parte inteira do n�mero
    While i <= (ParteInteira / 3)
        'Se � o primeiro agrupamento, n�o coloca quantificador no valor extenso
        If (i = 1) Then
            pedaco = CInt(Mid(ValorInt, ParteInteira - (3 * i - 1), 3))
        Else: pedaco = CInt(Mid(ValorInt, ParteInteira - (3 * i - 1), 3))
        End If
        'Se � o segundo agrupamento, adiciona quantificador de milhar no valor extenso
        If (i = 2) And (pedaco <> 0) Then
            extenso = " " + milhar + " e " + extenso
        End If
        'Se � um agrupamento superior, coloca o quantificador correspondente no valor extenso
        If (i >= 3) And (pedaco <> 0) Then
            If pedaco = 1 Then
                extenso = outros(i - 3) + diferenciacao(1) + ", " + extenso
            Else: extenso = outros(i - 3) + diferenciacao(2) + ", " + extenso
            End If
            
        End If
        'Atualiza o valor extenso com os n�meros, em extenso, do agrupamento correspondente
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
    
    'Pega os centavos do n�mero
    cents = Right(ValorStr, 2)
    'Se centavos for diferente de zero, analisa os centavos como um n�mero independente
    If (CLng(cents) <> 0) Then
        'Acrescenta um zero � esquerda do n�mero e cria seu extenso
        cents = Format(cents, "000")
        If (ParteInteira <> 0) Then
            extenso = extenso + " e " + ConvToExtensoAnalisaTres(cents, 3)
        Else
            extenso = extenso + ConvToExtensoAnalisaTres(cents, 3)
        End If
        
        'Define se ser� utilizado centavo ou centavos
        If CLng(cents) = 1 Then
            extenso = extenso + " " + diferenciacao(6)
        Else: extenso = extenso + " " + diferenciacao(7)
        End If
    End If
    
    ConvNumeroToExtenso = extenso
    Debug.Print extenso
End Function

'Pega t�s algarismos e define seu extenso
Public Function ConvToExtensoAnalisaTres(ValorInt, tamanho) As String
    'Prepara formata��o de letras
    unidade = Array("0", "Um", "Dois", "Tres", "Quatro", "Cinco", "Seis", "Sete", "Oito", "Nove", "Dez", "Onze", "Doze", "Treze", "Quatorze" _
    , "Quinze", "Dezesseis", "Dezessete", "Dezoito", "Dezenove")
    dezena = Array("0", "0", "Vinte", "Trinta", "Quarenta", "Cinqueta", "Sessenta", "Setenta", "Oitenta", "Noventa")
    centena = Array("0", "Cento", "Duzentos", "Trezentos", "Quatrocentos", "Quinhentos", "Seiscentos", "Setecentos", "Oitocentos", "Novecentos")
    milhar = "Mil"
    outros = Array("0", "Mi", "Bi", "Tri", "Quadri", "Quinti")
    diferenciacao = Array("0", "lh�o", "lh�es", "Cem", "Real", "Reais", "Centavo", "Centavos")

    i = 1
    
    'Enquanto n�o analisou todos os tres algarismos:
    While i < 3
        
        'Se � a primeira itera��o, pega os dois primeiros algarismos, sen�o pega o terceiro
        If (i = 1) Then
            pedaco = CInt(Mid(ValorInt, tamanho - i, 2))
        Else: pedaco = CInt(Mid(ValorInt, tamanho - i, 1))
        End If
        
        'Se � a terceira itera��o, adiciona o extenso correspondente
        If (i = 2) And (pedaco <> 0) Then
            extenso = centena(pedaco) + " e " + extenso
        End If
        
        'Se � a primeira itera��o, processa os dois primeiros algarismos
        If (i = 1) Then
            'Se o n�mero � diferente de 0, processa seu extenso
            If (pedaco <> 0) Then
                'Se estiver entre 1~19, adiciona o extenso correspondente
                If (pedaco < 20) Then
                    extenso = extenso + unidade(pedaco)
                Else
                    p1 = Int(pedaco / 10)
                    p2 = pedaco - p1 * 10
                    'Se o primeiro algarismo for 0, adiciona somente o quantificador, sen�o, adiciona o extenso correspondente
                    If p2 = 0 Then
                        extenso = extenso + dezena(p1)
                    Else: extenso = extenso + dezena(p1) + " e " + unidade(p2)
                    End If
                End If
            Else 'Se o n�mero = 0, analisa o terceiro algarismo
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

'Retorna um valor abreviado para o m�s em formato brasileiro
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

'Retorna o nome completo de um mes de refer�ncia na l�ngua portuguesa
Public Function MesCompleto(Referencia As Integer) As String
    Select Case Referencia
        Case 1: MesCompleto = "Janeiro"
        Case 2: MesCompleto = "Fevereiro"
        Case 3: MesCompleto = "Mar�o"
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
    dia = Format(arg01, "dd")
    Mes = CInt(Format(arg01, "mm"))
    ano = Format(arg01, "yyyy")
    
    'Coloca o dia em numero de 2 digitos
    DataExtenso = dia + " de "
    
    'Coloca o mes em extenso
    DataExtenso = DataExtenso + MesCompleto
    
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
'Condi��o de Erro
ERR1:
    'MsgBox "Erro ao salvar PDF,caminho especificado n�o existe ou um arquivo de mesmo nome est� aberto. Arquivo n�o Criado", , "ERRO"
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
'Condi��o de Erro
ERR1:
    MsgBox "Erro ao Abrir PDF,caminho especificado n�o existe", , "ERRO"
    Exit Function
End Function

'Salva uma p�gina espec�fica em PDF e abre
Public Function CriaAbrePDF(NomeArquivo As String, Categoria, NomeObjeto As String)
    If Not CriaPDF(NomeArquivo, Categoria, NomeObjeto) Then
        Exit Function
    End If
    
    If Not AbrePDF(NomeArquivo, Categoria, NomeObjeto) Then
        Exit Function
    End If
End Function