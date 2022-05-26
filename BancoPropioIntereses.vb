Public Class BancoPropioIntereses

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroBancoPropioIntereses"

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarNDiasDesde)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarNDiasHasta)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarPorcentaje)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClave)
    End Sub

    <Task()> Public Shared Sub AsignarNDiasDesde(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("NDiasDesde")) = 0 Then data("NDiasDesde") = 0
    End Sub

    <Task()> Public Shared Sub AsignarNDiasHasta(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("NDiasHasta")) = 0 Then data("NDiasHasta") = 0
    End Sub

    <Task()> Public Shared Sub AsignarPorcentaje(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("Porcentaje")) = 0 Then data("Porcentaje") = 0
    End Sub

    <Task()> Public Shared Sub AsignarClave(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then data("IDInteres") = AdminData.GetAutoNumeric
    End Sub

#End Region

End Class