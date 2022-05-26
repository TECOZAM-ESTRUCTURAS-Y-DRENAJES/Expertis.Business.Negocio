Public Class SolicitudTransferenciaCabecera

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbSolicitudTransferenciaCabecera"

#End Region

#Region "Eventos AddNewForm"

    Public Overrides Function AddNewForm() As DataTable
        Dim dt As DataTable = MyBase.AddNewForm
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf FillDefaultValues, dt.Rows(0), New ServiceProvider)
        Return dt
    End Function

    <Task()> Friend Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarContadorDefecto, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarFechaSolicitud, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarEstado, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarCentroGestion, data, services)
    End Sub

    <Task()> Public Shared Sub AsignarContadorDefecto(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim StDatos As New Contador.DatosDefaultCounterValue
        StDatos.Row = data
        StDatos.EntityName = "SolicitudTransferenciaCabecera"
        StDatos.FieldName = "NSolicitud"
        ProcessServer.ExecuteTask(Of Contador.DatosDefaultCounterValue)(AddressOf Contador.LoadDefaultCounterValue, StDatos, New ServiceProvider)
    End Sub

    <Task()> Public Shared Sub AsignarFechaSolicitud(ByVal data As DataRow, ByVal services As ServiceProvider)
        data("FechaSolicitud") = Date.Today
    End Sub

    <Task()> Public Shared Sub AsignarEstado(ByVal data As DataRow, ByVal services As ServiceProvider)
        data("EstadoCabecera") = 0
    End Sub

    <Task()> Public Shared Sub AsignarCentroGestion(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim f As New Filter
        f.Add(New StringFilterItem("IDUsuario", ProcessServer.ExecuteTask(Of Object, Guid)(AddressOf Operario.ObtenerIDUsuario, New Object, services).ToString()))
        Dim dtUsuarioCentro As DataTable = New UsuarioCentroGestion().Filter(f)
        If Not IsNothing(dtUsuarioCentro) AndAlso dtUsuarioCentro.Rows.Count > 0 AndAlso Length(dtUsuarioCentro.Rows(0)("IDCentroGestion")) Then
            data("IDCentroGestionSolicitado") = dtUsuarioCentro.Rows(0)("IDCentroGestion")
            Dim STC As New SolicitudTransferenciaCabecera
            STC.ApplyBusinessRule("IDCentroGestionSolicitado", data("IDCentroGestionSolicitado"), data)
        End If
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Public Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarIdentificador)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarNSolicitud)
    End Sub

    <Task()> Public Shared Sub AsignarIdentificador(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IdSolicitud")) = 0 Then data("IdSolicitud") = AdminData.GetAutoNumeric
        End If
    End Sub

    <Task()> Public Shared Sub AsignarNSolicitud(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added AndAlso Not IsDBNull(data("IDContador")) Then
            data("NSolicitud") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data("IDContador"), services)
        End If
    End Sub

#End Region

#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBRL As New BusinessRules
        oBRL.Add("IDCentroGestionSolicitante", AddressOf CambioCentroGestionSolicitante)
        oBRL.Add("IDCentroGestionSolicitado", AddressOf CambioCentroGestionSolicitado)
        Return oBRL
    End Function

    <Task()> Public Shared Sub CambioCentroGestionSolicitante(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim dtal As DataTable = New Almacen().Filter(New StringFilterItem("IDCentroGestion", data.Value))
        If dtal.Rows.Count > 0 Then
            data.Current("IDAlmacenOrigen") = dtal.Rows(0)("IDAlmacen")
        End If
    End Sub

    <Task()> Public Shared Sub CambioCentroGestionSolicitado(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim dtal As DataTable = New Almacen().Filter(New StringFilterItem("IDCentroGestion", data.Value))
        If dtal.Rows.Count > 0 Then
            data.Current("IDAlmacenDestino") = dtal.Rows(0)("IDAlmacen")
        End If
    End Sub

#End Region

End Class