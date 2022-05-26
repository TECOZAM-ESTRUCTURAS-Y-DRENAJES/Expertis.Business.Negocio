Public Class TipoModificacionRiesgo
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub
    Private Const cnEntidad As String = "tbMaestroTipoModificacionRiesgo"

#Region " RegisterValidateTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarRegistroExistente)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDTipoModificacion")) = 0 Then ApplicationService.GenerateError("El Identificador es un dato obligatorio.")
        If Length(data("DescTipoModificacion")) = 0 Then ApplicationService.GenerateError("La Descripción es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarRegistroExistente(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New TipoModificacionRiesgo().SelOnPrimaryKey(data("IDTipoModificacion"))
            If dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("Ya existe el Registro indicado en el sistema.")
            End If
        End If
    End Sub

#End Region

End Class
