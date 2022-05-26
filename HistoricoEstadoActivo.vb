Public Class HistoricoEstadoActivo
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbHistoricoEstadoActivo"

#Region "Eventos RegisterAddNewTasks"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        data("IDHistoricoActivo") = AdminData.GetAutoNumeric
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClave)
    End Sub

    <Task()> Public Shared Sub AsignarClave(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDHistoricoActivo")) = 0 Then data("IDHistoricoActivo") = AdminData.GetAutoNumeric
        End If
    End Sub

#End Region

    <Serializable()> _
    Public Class dataEstadoActivoRetorno
        Public IDActivo As String
        Public EstadoPendiente As String
        Public FechaAlquiler As Date
        Public IDOperario As String
        Public Texto As String

        Public Sub New(ByVal IDActivo As String, ByVal EstadoPendiente As String, ByVal FechaAlquiler As Date, ByVal IDOperario As String, ByVal Texto As String)
            Me.IDActivo = IDActivo
            Me.EstadoPendiente = EstadoPendiente
            Me.FechaAlquiler = FechaAlquiler
            Me.IDOperario = IDOperario
            Me.Texto = Texto
        End Sub
    End Class
    <Task()> Public Shared Sub NuevoEstadoActivoHistorico(ByVal data As dataEstadoActivoRetorno, ByVal services As ServiceProvider)
        Dim dt As DataTable = New HistoricoEstadoActivo().AddNew
        Dim dr As DataRow = dt.NewRow
        dr("IDHistoricoActivo") = AdminData.GetAutoNumeric
        dr("IDActivo") = data.IDActivo
        If Length(data.EstadoPendiente) = 0 Then
            data.EstadoPendiente = ProcessServer.ExecuteTask(Of Object, String)(AddressOf Activo.EstadoPendiente, Nothing, services)
        End If
        dr("IDEstadoActivo") = data.EstadoPendiente
        dr("FechaEstado") = data.FechaAlquiler
        dr("IDOperario") = data.IDOperario
        dr("Texto") = data.Texto
        dt.Rows.Add(dr.ItemArray)

        BusinessHelper.UpdateTable(dt)
    End Sub

End Class