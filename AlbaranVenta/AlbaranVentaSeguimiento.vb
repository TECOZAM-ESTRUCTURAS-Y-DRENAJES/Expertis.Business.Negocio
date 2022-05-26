﻿Public Class AlbaranVentaSeguimiento

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub


    Private Const cnEntidad As String = "tbAlbaranVentaSeguimiento"

#End Region

#Region "Eventos Entidad"
    'Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
    '    MyBase.RegisterAddnewTasks(addnewProcess)
    '    addnewProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
    'End Sub

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then data("IDLineaSeguimiento") = AdminData.GetAutoNumeric
    End Sub

#End Region

End Class