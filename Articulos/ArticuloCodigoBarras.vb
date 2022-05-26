Public Class ArticuloCodigoBarras

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbArticuloCodigoBarras"

#End Region

#Region "Tareas RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarArticulo)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarCodigoBarras)
    End Sub

    <Task()> Public Shared Sub ValidarArticulo(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDArticulo")) = 0 Then ApplicationService.GenerateError("El Artículo es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarCodigoBarras(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("CodigoBarras")) = 0 Then ApplicationService.GenerateError("El Código de barras es un obligatorio.")
        Dim DtCodBarra As DataTable = New ArticuloCodigoBarras().Filter(New FilterItem("CodigoBarras", FilterOperator.Equal, data("CodigoBarras")))
        If Not DtCodBarra Is Nothing AndAlso DtCodBarra.Rows.Count > 0 Then
            ApplicationService.GenerateError("El Código de Barras: |, ya existe para el artículo: |.", data("CodigoBarras"), DtCodBarra.Rows(0)("IDArticulo"))
        End If
    End Sub

#End Region

#Region "Tareas RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
    End Sub

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("ID")) = 0 Then data("ID") = AdminData.GetAutoNumeric
        End If
    End Sub

#End Region

End Class