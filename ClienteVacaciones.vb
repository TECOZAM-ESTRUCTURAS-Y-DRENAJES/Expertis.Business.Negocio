Public Class ClienteVacaciones

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub
    Private Const cnEntidad As String = "tbClienteVacaciones"

#End Region
    
#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim Obrl As New BusinessRules
        Obrl.Add("FechaDessde", AddressOf CambioFecha)
        Obrl.Add("FechaHasta", AddressOf CambioFecha)
        Return Obrl
    End Function

    <Task()> Public Shared Sub CambioFecha(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If IsDate(data.Value) Then
            data.Current(data.ColumnName) = data.Value
            If Length(data.Current("FechaDesde")) > 0 And Length(data.Current("FechaHasta")) > 0 Then
                If data.Current("FechaHasta") < data.Current("FechaDesde") Then
                    ApplicationService.GenerateError("La Fecha Hasta debe ser mayor que la Fecha Desde.")
                End If
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarFechas)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarCliente)
    End Sub

    <Task()> Public Shared Sub ValidarFechas(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("FechaDesde")) = 0 OrElse Length(data("FechaHasta")) = 0 Then
            ApplicationService.GenerateError("La Fecha Desde y Fecha Hasta son obligatorias.")
        End If
        If Length(data("FechaDivision")) = 0 AndAlso Length(data("FechaAlternativa")) = 0 Then
            ApplicationService.GenerateError("Debe indicar Fecha Alternativa o Fecha División.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarCliente(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDCliente")) = 0 Then ApplicationService.GenerateError("El Cliente es un dato obligatorio.")
        End If
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarPrimaryKey)
    End Sub

    <Task()> Public Shared Sub AsignarPrimaryKey(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If IsDBNull(data("IDVacacion")) OrElse data("IDVacacion") = 0 Then
                data("IDVacacion") = AdminData.GetAutoNumeric
            End If
        End If
    End Sub

#End Region

End Class