Public Class TipoApunte

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroTipoApunte"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarTipoApunte)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarDuplicados)
    End Sub

    <Task()> Public Shared Sub ComprobarTipoApunte(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDTipoApunte")) = 0 Then ApplicationService.GenerateError("Tipo Apunte es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ComprobarDuplicados(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim DtDatos As DataTable = New TipoApunte().SelOnPrimaryKey(data("IDTipoApunte"))
            If Not DtDatos Is Nothing AndAlso DtDatos.Rows.Count > 0 Then
                ApplicationService.GenerateError("El Tipo Apunte introducido ya existe en la Base de Datos.")
            End If
        End If
    End Sub

#End Region

End Class