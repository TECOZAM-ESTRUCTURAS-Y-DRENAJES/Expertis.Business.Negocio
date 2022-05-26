Public Class TipoPago

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroTipoPago"

#End Region

#Region "Eventos RegisterDeleteTasks"

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf ComprobarSistema)
    End Sub

    <Task()> Public Shared Sub ComprobarSistema(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data("Sistema") Then ApplicationService.GenerateError("No se puede realizar esta operación sobre un Estado del Sistema.")
    End Sub

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarTipoPago)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarDescripcionTipoPago)
    End Sub

    <Task()> Public Shared Sub ComprobarTipoPago(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDTipoPago")) = 0 Then ApplicationService.GenerateError("El Tipo de Pago es obligatorio.")
            Dim DtDatos As DataTable = New TipoPago().SelOnPrimaryKey(data("IDTipoPago"))
            If Not DtDatos Is Nothing AndAlso DtDatos.Rows.Count > 0 Then ApplicationService.GenerateError("El Tipo Pago introducido ya existe.")
        End If
    End Sub

    <Task()> Public Shared Sub ComprobarDescripcionTipoPago(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("DescTipo")) = 0 Then ApplicationService.GenerateError("La Descripción del Tipo Pago es obligatoria.")
    End Sub

#End Region

End Class