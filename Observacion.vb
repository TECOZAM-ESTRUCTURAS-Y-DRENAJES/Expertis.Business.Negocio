Public Class Observacion

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroObservacion"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("DescObservacion")) = 0 Then ApplicationService.GenerateError("Introduzca la descripción de la observación")
        If Length(data("Entidad")) = 0 Then ApplicationService.GenerateError("Introduzca la entidad correspondiente a la observación")
        Dim DtDatos As DataTable = New BE.DataEngine().Filter("xEntity", New FilterItem("Entidad", FilterOperator.Equal, data("Entidad")), , , , True)
        If DtDatos Is Nothing OrElse DtDatos.Rows.Count = 0 Then ApplicationService.GenerateError("La entidad introducida no está en la lista")
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarContador)
    End Sub

    <Task()> Public Shared Sub AsignarContador(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If data("IDObservacion").ToString.Trim.Length = 0 Then
                Dim contObservaciones As String = New Parametro().ContadorObservaciones()
                Dim dt As DataTable = New Contador().SelOnPrimaryKey(contObservaciones)
                If dt Is Nothing OrElse dt.Rows.Count = 0 Then
                    ApplicationService.GenerateError("No ha introducido el código de la observación y el contador por defecto para observaciones no existe")
                Else : data("IDObservacion") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, CStr(dt.Rows(0)("IDContador")), services)
                End If
            End If
            If Length(data("IDObservacion")) > 0 Then
                Dim dtTemp As DataTable = New Observacion().SelOnPrimaryKey(data("IDObservacion"))
                If Not dtTemp Is Nothing AndAlso dtTemp.Rows.Count > 0 Then
                    ApplicationService.GenerateError("Observación duplicada. -")
                End If
            End If
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Serializable()> _
    Public Class DatosObv
        Public IDEntidad As String
        Public IDPrimaryKey As String
    End Class

    <Task()> Public Shared Function ObtenerObservacionesCliente(ByVal data As DatosObv, ByVal services As ServiceProvider) As String
        Dim FilClie As New Filter
        FilClie.Add("IDCliente", FilterOperator.Equal, data.IDPrimaryKey)
        FilClie.Add("Entidad", FilterOperator.Equal, data.IDEntidad)
        Dim dtCliente As DataTable = New BE.DataEngine().Filter("frmClienteObservacion", FilClie)
        Dim StrResul As String = String.Empty
        If Not dtCliente Is Nothing AndAlso dtCliente.Rows.Count > 0 Then
            For Each rowObservacion As DataRow In dtCliente.Rows
                If Length(ObtenerObservacionesCliente) > 0 Then
                    StrResul &= vbNewLine
                End If
                StrResul &= rowObservacion("DescObservacion")
            Next
        End If

        Dim FilObv As New Filter
        FilObv.Add("Entidad", FilterOperator.Equal, data.IDEntidad)
        FilObv.Add("ImprimirSiempre", FilterOperator.NotEqual, 0)
        Dim dtObservacion As DataTable = New Observacion().Filter(FilObv)
        If Not dtObservacion Is Nothing AndAlso dtObservacion.Rows.Count > 0 Then
            For Each rowObservacion As DataRow In dtObservacion.Select
                If Length(ObtenerObservacionesCliente) > 0 Then StrResul &= vbNewLine
                StrResul &= rowObservacion("DescObservacion")
            Next
        End If
        Return StrResul
    End Function

    <Task()> Public Shared Function ObtenerObservacionesProveedor(ByVal data As DatosObv, ByVal services As ServiceProvider) As String
        Dim FilBE As New Filter
        FilBE.Add("IDProveedor", FilterOperator.Equal, data.IDPrimaryKey, FilterType.String)
        FilBE.Add("Entidad", FilterOperator.Equal, data.IDEntidad, FilterType.String)
        Dim DtCliente As DataTable = New DataEngine().Filter("frmProveedorObservacion", FilBE)
        Dim StrObv As String
        If Not DtCliente Is Nothing AndAlso DtCliente.Rows.Count > 0 Then
            For Each Dr As DataRow In DtCliente.Select
                If Length(StrObv) > 0 Then StrObv &= vbNewLine
                StrObv &= Dr("DescObservacion")
            Next
        End If

        Dim FilObv As New Filter
        FilObv.Add("Entidad", FilterOperator.Equal, data.IDEntidad, FilterType.String)
        FilObv.Add("ImprimirSiempre", FilterOperator.NotEqual, 0)
        Dim DtObv As DataTable = New Observacion().Filter(FilObv)
        If Not DtObv Is Nothing AndAlso DtObv.Rows.Count > 0 Then
            For Each DrObv As DataRow In DtObv.Select
                If Length(StrObv) > 0 Then
                    StrObv &= vbNewLine
                End If
                StrObv &= DrObv("DescObservacion")
            Next
        End If
        Return StrObv
    End Function

    <Task()> Public Shared Function ObtenerObservacionesPorDefecto(ByVal data As String, ByVal services As ServiceProvider) As String
        Dim FilObv As New Filter
        FilObv.Add("Entidad", FilterOperator.Equal, data)
        FilObv.Add("ImprimirSiempre", FilterOperator.NotEqual, 0)
        Dim dtObservacion As DataTable = New Observacion().Filter(FilObv)
        Dim StrResul As String = String.Empty
        If Not dtObservacion Is Nothing AndAlso dtObservacion.Rows.Count > 0 Then
            For Each rowObservacion As DataRow In dtObservacion.Select
                If Length(ObtenerObservacionesPorDefecto) > 0 Then
                    StrResul &= vbNewLine
                End If
                StrResul &= rowObservacion("DescObservacion")
            Next
        End If
        Return StrResul
    End Function

    <Task()> Public Shared Function ObtenerxEntidades(ByVal data As Object, ByVal service As ServiceProvider) As DataTable
        Return New BE.DataEngine().Filter("xEntity", "*", "", , , True)
    End Function

#End Region

End Class