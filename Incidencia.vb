Public Class Incidencia

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroIncidencia"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New Incidencia().SelOnPrimaryKey(data("IDIncidencia"))
            If Not IsNothing(dt) AndAlso Not dt.Rows.Count.Equals(0) Then
                ApplicationService.GenerateError("La Incidencia ya existe.")
            End If
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function ValidaHora(ByVal data As String, ByVal services As ServiceProvider) As DataTable
        Dim dt As DataTable = New Incidencia().SelOnPrimaryKey(data)
        If IsNothing(dt) OrElse dt.Rows.Count = 0 Then ApplicationService.GenerateError("La Incidencia | no existe.", data)
        Return dt
    End Function

#End Region

End Class