Attribute VB_Name = "Mdl_Main"
Option Compare Database

'********************************************Objetos globais************************************************************
Public Principal As New cls_Utilities
Public listas As New cls_Listas

'********************************************Constantes comuns**********************************************************
Public Const ConstEquipStatusAtivo As String = "Ativo"
Public Const ConstEquipStatusInativo As String = "Inativo"
Public Const ConstEquipStatusTodos As String = "Todos"

Public Const ConstManutencaoTipoPreventiva As String = "Preventiva"
Public Const ConstManutencaoTipoCorretiva As String = "Corretiva"

'***********************************Métodos comuns a todos os formulários***********************************************
Public Function AbreFormulario(nome As String)
    Principal.AbreFormulario nome
End Function

Public Function Fecha(nome As String)
    Principal.FechaFormulario nome
End Function

Private Function Redimensiona(LarguraOriginal, AlturaOriginal, LarguraAtual, AlturaAtual, NovaLarguraNovaAltura, NomeForm)
    LarguraAtual = Principal.RedimensionaHorizontal(LarguraOriginal, LarguraAtual, NovaLargura, NomeForm)
    AlturaAtual = Principal.RedimensionaDetalheVertical(AlturaOriginal, AlturaAtual, NovaAltura, NomeForm)
End Function

'********************************************Outros Métodos*************************************************************
Public Sub atualiza_data_seq()
    Dim resultado As DAO.Recordset

    Consulta = "SELECT * FROM manutencao"
    Set resultado = CurrentDb.OpenRecordset(Consulta)
    
    While Not (resultado.EOF)
        n_seq = (Format(resultado![data_manutencao], "YYYYMMDD"))
        Consulta = "UPDATE manutencao SET data_seq = " + CStr(n_seq) + " WHERE Manutencao_ID = " + CStr(resultado![Manutencao_ID])
        CurrentDb.Execute Consulta
        resultado.MoveNext
    Wend
End Sub

Function teste()

    'Variáveis associadas ao banco de dados
    Dim str As String
    Dim resultado As DAO.Recordset
    Dim resultado2 As DAO.Recordset
    Dim banco As DAO.Database
    Dim data0 As String
    Dim data1 As String
    Set banco = CurrentDb
    
    ''''
    ''' Teste de data
    ''''
    data1 = "30/8/2015"
    data0 = "01/01/2000"
    ''''''''''''''''''''''''''''''''
    
    'Variáveis associadas ao processamento de dados
    Dim equip()
    
    'Consulta para pegar a data de ultima manutenção e setor atual do equipamento
    str1 = "SELECT Equipamento.nome_equipamento, Equipamento.Status, " + _
    "Setor.SIGLA " + _
    "FROM Equipamento " + _
    "INNER JOIN Setor ON Equipamento.Setor = Setor.Setor_ID " + _
    "ORDER BY Equipamento.nome_equipamento DESC "
    
    Set resultado = banco.OpenRecordset(str1)
    
    'Consulta para geração e avaliação de custos
    str2 = "SELECT Equipamento.nome_equipamento, SUM(Manutencao.custo) as custo, " & _
    "Setor.SIGLA " & _
    "FROM (Equipamento " & _
    "INNER JOIN Manutencao ON Equipamento.Equip_ID = Manutencao.equipamento_verificado) " & _
    "INNER JOIN Setor ON Manutencao.Setor = Setor.Setor_ID " & _
    "WHERE Manutencao.data_manutencao BETWEEN #" & data0 & "# AND #" & data1 & "# " & _
    "GROUP BY Equipamento.nome_equipamento, Setor.SIGLA "
    
    Set resultado2 = banco.OpenRecordset(str2)
    
    'Consulta para pegar a ultima manutencao feita no equipamento
    str3 = "SELECT Equipamento.nome_equipamento, " + _
    "max(Manutencao.data_manutencao) as data_manutencao " + _
    "FROM Equipamento " + _
    "INNER JOIN Manutencao ON Equipamento.Equip_ID = Manutencao.equipamento_verificado " + _
    "GROUP BY Equipamento.nome_equipamento"
    
     Set resultado3 = banco.OpenRecordset(str3)

    'Consulta para pegar o custo total do equipamento no setor analisado
    str4 = "SELECT Equipamento.nome_equipamento, SUM(Manutencao.custo) as custo, " & _
    "Setor.SIGLA " & _
    "FROM (Equipamento " & _
    "INNER JOIN Manutencao ON Equipamento.Equip_ID = Manutencao.equipamento_verificado) " & _
    "INNER JOIN Setor ON Manutencao.Setor = Setor.Setor_ID " & _
    "GROUP BY Equipamento.nome_equipamento, Setor.SIGLA "
    
    Set resultado4 = banco.OpenRecordset(str4)
    
    '''Processamento de dados
     'Definicao das dimensoes do problema
     
    
    
    'Mostra Resultados na janela de testes
    Do While Not resultado.EOF
        Debug.Print "Equipamento: " & resultado![nome_equipamento] & " Setor: " & resultado![Sigla]
        resultado.MoveNext
    Loop
    
    Do While Not resultado2.EOF
        Debug.Print "Equipamento: " & resultado2![nome_equipamento] & " Setor: " & resultado2![Sigla] & " Custo: " & resultado2![custo]
        resultado2.MoveNext
    Loop
    
    Do While Not resultado3.EOF
        Debug.Print "Equipamento: " & resultado3![nome_equipamento] & " Manutencao: " & resultado3![data_manutencao]
        resultado3.MoveNext
    Loop
    
    Do While Not resultado4.EOF
        Debug.Print "Equipamento: " & resultado4![nome_equipamento] & " Setor: " & resultado4![Sigla] & " Custo: " & resultado4![custo]
        resultado4.MoveNext
    Loop
    
End Function


Function define_consultas()
    
    'Cria consulta para pegar dados atuais do equipamento
    query1 = "dados_atuais"
    str1 = "SELECT Equipamento.nome_equipamento, Equipamento.Status, " + _
    "Setor.SIGLA " + _
    "FROM Equipamento " + _
    "INNER JOIN Setor ON Equipamento.Setor = Setor.Setor_ID " + _
    "ORDER BY Equipamento.nome_equipamento DESC "
    
    If Not IsNull(DLookup("Type", "MSYSObjects", "Name='" + query1 + "'")) Then
        CurrentDb.QueryDefs(query1).SQL = str1
    Else
        CurrentDb.CreateQueryDef query1, str1
    End If
    
    
    'Consulta para geração e avaliação de custos
    query2 = "dados_custo"
    str2 = "SELECT Equipamento.nome_equipamento, SUM(Manutencao.custo) as custo, " & _
    "Setor.SIGLA " & _
    "FROM (Equipamento " & _
    "INNER JOIN Manutencao ON Equipamento.Equip_ID = Manutencao.equipamento_verificado) " & _
    "INNER JOIN Setor ON Manutencao.Setor = Setor.Setor_ID " & _
    "WHERE Manutencao.data_manutencao BETWEEN [Enter Start Date] AND [Enter End Date] " & _
    "GROUP BY Equipamento.nome_equipamento, Setor.SIGLA "
    
    If Not IsNull(DLookup("Type", "MSYSObjects", "Name='" + query2 + "'")) Then
       CurrentDb.QueryDefs(query2).SQL = str2
    Else
        CurrentDb.CreateQueryDef query2, str2
    End If
    
    
    'Consulta para data de ultima manutencao
    query3 = "dados_man"
    str3 = "SELECT Equipamento.nome_equipamento, " + _
    "max(Manutencao.data_manutencao) as data_manutencao " + _
    "FROM Equipamento " + _
    "INNER JOIN Manutencao ON Equipamento.Equip_ID = Manutencao.equipamento_verificado " + _
    "GROUP BY Equipamento.nome_equipamento"
    
    If Not IsNull(DLookup("Type", "MSYSObjects", "Name='" + query3 + "'")) Then
       CurrentDb.QueryDefs(query3).SQL = str3
    Else
        CurrentDb.CreateQueryDef query3, str3
    End If
    
    'Consulta para pegar o custo total do equipamento no setor analisado
    query4 = "custo_total"
    str4 = "SELECT Equipamento.nome_equipamento, SUM(Manutencao.custo) as custo, " & _
    "Setor.SIGLA " & _
    "FROM (Equipamento " & _
    "INNER JOIN Manutencao ON Equipamento.Equip_ID = Manutencao.equipamento_verificado) " & _
    "INNER JOIN Setor ON Manutencao.Setor = Setor.Setor_ID " & _
    "GROUP BY Equipamento.nome_equipamento, Setor.SIGLA "
    
    If Not IsNull(DLookup("Type", "MSYSObjects", "Name='" + query4 + "'")) Then
       CurrentDb.QueryDefs(query4).SQL = str4
    Else
        CurrentDb.CreateQueryDef query4, str4
    End If
    
End Function

