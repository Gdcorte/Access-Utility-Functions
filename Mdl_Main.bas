Attribute VB_Name = "Mdl_Main"
Option Compare Database

'********************************************Objetos globais************************************************************
Public Principal As New cls_Utilities

'********************************************Constantes comuns**********************************************************

'***********************************Métodos comuns a todos os formulários***********************************************
Public Function AbreFormulario(nome As String)
    Principal.AbreFormulario nome
End Function

Public Function Fecha(nome As String)
    Principal.FechaFormulario nome
End Function

