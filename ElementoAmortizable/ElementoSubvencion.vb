Public Class ElementoSubvencion

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbElementoSubvencion"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim f As New Filter(FilterUnionOperator.And)
            f.Add("IDElemento", FilterOperator.Equal, data("IDElemento"))
            f.Add("IDSubvencion", FilterOperator.Equal, data("IDSubvencion"))
            Dim DtDatos As DataTable = New ElementoSubvencion().Filter(f)
            If Not DtDatos Is Nothing AndAlso DtDatos.Rows.Count > 0 Then
                ApplicationService.GenerateError("Subvención-elemento duplicado.")
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClave)
    End Sub

    <Task()> Public Shared Sub AsignarClave(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDElementoSubvencion")) = 0 Then data("IDElementoSubvencion") = AdminData.GetAutoNumeric
    End Sub

#End Region

End Class