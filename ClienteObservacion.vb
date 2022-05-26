Public Class ClienteObservacion

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub
    Private Const cnEntidad As String = "tbClienteObservacion"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarIDCliente)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarIDObservacion)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarObservacion)
    End Sub

    <Task()> Public Shared Sub ValidarIDCliente(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDCliente")) = 0 Then ApplicationService.GenerateError("El Cliente es un dato obligatorio.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarIDObservacion(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDObservacion")) = 0 Then ApplicationService.GenerateError("Observación es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarObservacion(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim StrWhere As String
        If data.RowState = DataRowState.Modified Then
            StrWhere = "IDClienteObservacion <>" & data("IDClienteObservacion")
        End If
        If Len(StrWhere) > 0 Then StrWhere = StrWhere & " AND "
        StrWhere = "IDCliente='" & data("IDCliente") & "' AND IDObservacion='" & data("IDObservacion") & "'"
        Dim Dt As DataTable = New ClienteObservacion().Filter("IDClienteObservacion", StrWhere)
        If Not Dt Is Nothing AndAlso Dt.Rows.Count > 0 Then
            ApplicationService.GenerateError("Ya existe esta observación para el cliente actual.")
        End If
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
    End Sub

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDClienteObservacion")) = 0 Then data("IDClienteObservacion") = AdminData.GetAutoNumeric
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function ObtenerEntidades(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
        Return New BE.DataEngine().Filter("frmEntidadObservacion", "*", "", "Entidad")
    End Function

#End Region

End Class