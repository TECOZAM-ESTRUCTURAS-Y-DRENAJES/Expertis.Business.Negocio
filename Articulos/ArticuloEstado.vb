Public Class ArticuloEstado

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroArticuloEstado"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDEstado")) = 0 Then ApplicationService.GenerateError("El Estado es un dato obligatorio.")
        If Length(data("DescEstado")) = 0 Then ApplicationService.GenerateError("El descripción del estado es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New ArticuloEstado().SelOnPrimaryKey(data("IDEstado"))
            If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then ApplicationService.GenerateError("El código de Estado ya existe en la Base de Datos")
        End If
    End Sub

#End Region

End Class