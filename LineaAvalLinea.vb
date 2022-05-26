Public Class LineaAvalLinea

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroLineaAvalLinea"

#End Region

#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("FechaDesde", AddressOf CambioFechas)
        oBrl.Add("Fechahasta", AddressOf CambioFechas)
        oBrl.Add("IDAvalEstado", AddressOf CambioAvalEstado)
        Return oBrl
    End Function

    <Task()> Public Shared Sub CambioFechas(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("FechaDesde")) > 0 AndAlso Length(data.Current("FechaHasta")) > 0 Then
            If data.Current("FechaDesde") > data.Current("FechaHasta") Then
                ApplicationService.GenerateError("La Fecha Desde no puede ser Mayor que la Fecha Hasta")
            ElseIf data.Current("FechaHasta") < data.Current("FechaDesde") Then
                ApplicationService.GenerateError("La Fecha Hasta no puede ser Menor que la Fecha Desde")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioAvalEstado(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current("FechaEstado") = IIf(Not IsDBNull(data.Value), Today.Date, Nothing)
    End Sub

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDAvalClase")) <= 0 Then ApplicationService.GenerateError("La Clase del Aval es Obligatoria.")
        If Length(data("IDObra")) <= 0 Then ApplicationService.GenerateError("El Número de Obra es Obligatorio.")
        If Length(data("IDAvalEstado")) <= 0 Then ApplicationService.GenerateError("El tipo de Estado del Aval es un dato Obligatorio.")
        If Length(data("FechaEstado")) <= 0 Then ApplicationService.GenerateError("La Fecha del Estado es un dato Obligatorio.")
        If Length(data("IDAvalClase")) <= 0 Then ApplicationService.GenerateError("El tipo de Clase del Aval es un dato Obligatorio.")
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClave)
    End Sub

    <Task()> Public Shared Sub AsignarClave(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then data("IDAval") = AdminData.GetAutoNumeric()
    End Sub

#End Region

End Class