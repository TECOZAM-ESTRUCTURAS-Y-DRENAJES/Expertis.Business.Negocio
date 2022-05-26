Public Class LineaAvalCabecera

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroLineaAvalCabecera"

#End Region

#Region "Eventos RegisterAddNewTasks"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        data("IDLineaAval") = AdminData.GetAutoNumeric
        data("ImpAval") = 0
        data("MesesCargo") = 0
        data("ComApertura") = 0
        data("AperturaMinimo") = 0
        data("ComEstudio") = 0
        data("EstudioMinimo") = 0
    End Sub

#End Region

#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("FechaFinalizacion", AddressOf CambioFechas)
        oBrl.Add("FechaApertura", AddressOf CambioFechas)
        Return oBrl
    End Function

    <Task()> Public Shared Sub CambioFechas(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("FechaApertura")) > 0 AndAlso Length(data.Current("FechaFinalizacion")) > 0 Then
            If data.Current("FechaApertura") > data.Current("FechaFinalizacion") Then
                ApplicationService.GenerateError("La Fecha de Apertura no puede ser Mayor que la Fecha de Finalización.")
            ElseIf data.Current("FechaFinalizacion") < data.Current("FechaApertura") Then
                ApplicationService.GenerateError("La Fecha de Finalización no puede ser Menor que la Fecha de Apertura.")
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("NumAval")) <= 0 Then ApplicationService.GenerateError("El Número de Aval es un dato Obligatorio.")
        If Length(data("IDBanco")) <= 0 Then ApplicationService.GenerateError("El Banco es un dato Obligatorio.")
        If Length(data("FechaApertura")) <= 0 Then ApplicationService.GenerateError("La Fecha de Apertura es un dato Obligatorio.")
        If Length(data("FechaFinalizacion")) <= 0 Then ApplicationService.GenerateError("La Fecha de Finalizacion es un dato Obligatorio.")
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClave)
    End Sub

    <Task()> Public Shared Sub AsignarClave(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDLineaAval")) = 0 Then data("IDLineaAval") = AdminData.GetAutoNumeric()
        End If
    End Sub

#End Region

End Class