Public Class ProveedorTarifa

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub
    Private Const cnEntidad As String = "tbProveedorTarifa"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarProveedor)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarIDTarifa)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarPrimaryKey)
    End Sub

    <Task()> Public Shared Sub ValidarProveedor(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDProveedor")) = 0 Then ApplicationService.GenerateError("El Proveedor es un dato obligatorio.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarIDTarifa(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IdTarifa")) = 0 Then ApplicationService.GenerateError("La Tarifa es un dato obligatorio.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarPrimaryKey(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dtTemp As DataTable = New ProveedorTarifa().SelOnPrimaryKey(data("IDProveedor"), data("IdTarifa"))
            If Not dtTemp Is Nothing AndAlso dtTemp.Rows.Count > 0 Then
                ApplicationService.GenerateError("Este Proveedor ya tiene asociada esa Tarifa")
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarOrden)
    End Sub

    <Task()> Public Shared Sub AsignarOrden(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Nz(data("Orden"), 0) = 0 Then data("Orden") = 1
    End Sub

#End Region

End Class