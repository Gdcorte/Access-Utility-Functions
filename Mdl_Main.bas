Attribute VB_Name = "Mdl_Main"
Option Compare Database

'********************************************Objetos globais************************************************************
'Objeto com classe principal de métodos utilizados internamente a um formulário ou controle
Public Principal As New cls_Utilities

'********************************************Constantes comuns**********************************************************

'***********************************Métodos comuns a todos os formulários***********************************************
'Método de abertura de um formulário
Public Function AbreFormulario(nome As String)
    Principal.AbreFormulario nome
End Function

'Método de fechamento de um formulário
Public Function Fecha(nome As String)
    Principal.FechaFormulario nome
End Function

'Método que redimensiona a janela do formulário. NOTA: Formulário deve conter as seguintes variáveis comuns a todo o formulário e listadas como públicas
'Public LarguraAtual As Double, LarguraOriginal As Double
'Public AlturaAtual As Double, AlturaOriginal As Double
Public Function Redimensiona(NomeForm As String)
    'Atualiza campos
    Forms(NomeForm).LarguraAtual = Principal.RedimensionaHorizontal(Forms(NomeForm).LarguraOriginal, Forms(NomeForm).LarguraAtual, Forms(NomeForm).InsideWidth, NomeForm)
    Forms(NomeForm).AlturaAtual = Principal.RedimensionaDetalheVertical(Forms(NomeForm).AlturaOriginal, Forms(NomeForm).AlturaAtual, Forms(NomeForm).InsideHeight, NomeForm)
End Function


