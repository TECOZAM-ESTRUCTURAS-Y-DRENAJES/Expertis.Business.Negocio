Public Class LineaAvalIntereses

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroLineaAvalIntereses"

#End Region

#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("ImporteDesde", AddressOf CambioImportes)
        oBrl.Add("ImporteHasta", AddressOf CambioImportes)
        Return oBrl
    End Function

    <Task()> Public Shared Sub CambioImportes(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("ImporteHasta")) > 0 AndAlso Not AreEquals(data.Current("ImporteHasta"), 0) Then
            If data.Current("ImporteDesde") > data.Current("ImporteHasta") Then
                ApplicationService.GenerateError("El Importe Desde no puede ser mayor que el Importe Hasta")
            ElseIf data.Current("ImporteHasta") < data.Current("ImporteDesde") Then
                ApplicationService.GenerateError("El Importe Hasta no puede ser menos que el Importe Desde")
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
        If Not Length(data("IDLineaAval")) > 0 Then ApplicationService.GenerateError("El ID de Linea de Aval es un Dato Obligatorio.")
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClave)
    End Sub

    <Task()> Public Shared Sub AsignarClave(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then data("IDAvalInteres") = AdminData.GetAutoNumeric()
    End Sub

#End Region

End Class