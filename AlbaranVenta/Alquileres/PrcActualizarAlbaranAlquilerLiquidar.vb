'David Velasco 10/8/22
Public Class PrcActualizarAlbaranAlquilerLiquidar
    Inherits Process(Of dataPrcActualizarAlbaranAlquiler, ResultAlbaranAlquiler)

    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of dataPrcActualizarAlbaranAlquiler, ArrayList)(AddressOf PrepararInformacionProceso)
        Me.AddForEachTask(Of DocumentoAlbaranVenta)(AddressOf ActualizarDocumentoAlbaranAlquiler, OnExceptionBehaviour.NextLoop)
        Me.AddTask(Of Object, ResultAlbaranAlquiler)(AddressOf Resultado)
    End Sub

    <Task()> Public Shared Function PrepararInformacionProceso(ByVal data As dataPrcActualizarAlbaranAlquiler, ByVal services As ServiceProvider) As ArrayList
        Dim arDocsAlbaranAlq As ArrayList = AdminData.GetSessionData("__AlbAlqx__")
        If Not arDocsAlbaranAlq Is Nothing Then
            If data.RstAlbaranAlquiler.CreateData Is Nothing Then data.RstAlbaranAlquiler.CreateData = New LogProcess
            services.RegisterService(data, GetType(dataPrcActualizarAlbaranAlquiler))
        End If

        Return arDocsAlbaranAlq
    End Function

    <Task()> Public Shared Sub ActualizarDocumentoAlbaranAlquiler(ByVal doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        AdminData.BeginTx()
        ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ActualizarDatosCabecera, doc, services)
        ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.GrabarDocumento, doc, services)

        ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.ActualizacionAutomaticaStock2, doc, services)
        ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVentaAlquiler.CerrarOrdenesServicio, doc, services)

        Dim data As dataPrcActualizarAlbaranAlquiler = services.GetService(Of dataPrcActualizarAlbaranAlquiler)()

        ProcessServer.ExecuteTask(Of dataPrcActualizarAlbaranAlquiler)(AddressOf ProcesoAlbaranVentaAlquiler.ActualizarIncidencias, data, services)
        If data.Activos Is Nothing Then
            ProcessServer.ExecuteTask(Of DataTable)(AddressOf ProcesoAlbaranVentaAlquiler.ActualizarActivos, doc.dtLineas, services)
        Else
            ProcessServer.ExecuteTask(Of DataTable)(AddressOf ProcesoAlbaranVentaAlquiler.ActualizarActivos, data.Activos, services)
        End If
        ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVentaAlquiler.AddConductores, doc, services)
        ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVentaAlquiler.AddContadores, doc, services)
        ProcessServer.ExecuteTask(Of dataPrcActualizarAlbaranAlquiler)(AddressOf ProcesoAlbaranVentaAlquiler.AddHistoricoAvisos, data, services)
        ProcessServer.ExecuteTask(Of dataPrcActualizarAlbaranAlquiler)(AddressOf ProcesoAlbaranVentaAlquiler.AddObraMaterial, data, services)

        AdminData.CommitTx(True)
        AgregarAlbaranAlquilerResultado(doc, services)
    End Sub

    <Task()> Public Shared Sub ActualizarDatosCabecera(ByVal doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        Dim dataAlbaranVenta As dataPrcActualizarAlbaranAlquiler = services.GetService(Of dataPrcActualizarAlbaranAlquiler)()
        Dim AlbaranVenta As ResultAlbaranAlquiler = dataAlbaranVenta.RstAlbaranAlquiler
        Dim dv As New DataView(AlbaranVenta.PropuestaAlbaranes, Nothing, "IDAlbaran", DataViewRowState.CurrentRows)
        Dim idx As Integer = dv.Find(doc.HeaderRow("IDAlbaran"))
        If idx >= 0 Then
            doc.HeaderRow("IDContador") = dv(idx)("IDContador")
            doc.HeaderRow("FechaAlbaran") = dv(idx)("FechaAlbaran")
            doc.HeaderRow("IDEjercicio") = dv(idx)("IDEjercicio")
            doc.HeaderRow("Matricula") = dv(idx)("Matricula")
            Dim infoCounterValue As New Contador.DatosCounterValue(doc.HeaderRow("IDContador"), New AlbaranVentaCabecera, "NAlbaran", "FechaAlbaran", doc.HeaderRow("FechaAlbaran"))
            infoCounterValue.IDEjercicio = doc.HeaderRow("IDEjercicio") & String.Empty
            doc.HeaderRow("NAlbaran") = ProcessServer.ExecuteTask(Of Contador.DatosCounterValue, String)(AddressOf Contador.CounterValue, infoCounterValue, services)
        End If
    End Sub

    <Task()> Public Shared Sub AgregarAlbaranAlquilerResultado(ByVal data As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        Dim dataAlbaranAlquiler As dataPrcActualizarAlbaranAlquiler = services.GetService(Of dataPrcActualizarAlbaranAlquiler)()
        Dim AlbaranAlquiler As ResultAlbaranAlquiler = dataAlbaranAlquiler.RstAlbaranAlquiler

        ReDim Preserve AlbaranAlquiler.StockUpdateData(UBound(AlbaranAlquiler.StockUpdateData) + 1)
        AlbaranAlquiler.StockUpdateData = services.GetService(Of AlbaranLogProcess)().StockUpdateData

        ReDim Preserve AlbaranAlquiler.CreateData.CreatedElements(UBound(AlbaranAlquiler.CreateData.CreatedElements) + 1)
        AlbaranAlquiler.CreateData.CreatedElements(UBound(AlbaranAlquiler.CreateData.CreatedElements)) = New CreateElement
        AlbaranAlquiler.CreateData.CreatedElements(UBound(AlbaranAlquiler.CreateData.CreatedElements)).IDElement = data.HeaderRow("IDAlbaran")
        AlbaranAlquiler.CreateData.CreatedElements(UBound(AlbaranAlquiler.CreateData.CreatedElements)).NElement = data.HeaderRow("NAlbaran")
    End Sub

    <Task()> Public Shared Function Resultado(ByVal data As Object, ByVal services As ServiceProvider) As ResultAlbaranAlquiler
        Dim dataAlbaranesAlquiler As dataPrcActualizarAlbaranAlquiler = services.GetService(Of dataPrcActualizarAlbaranAlquiler)()
        Return dataAlbaranesAlquiler.RstAlbaranAlquiler
    End Function

    Protected Overrides Function OnException(ByVal exceptionArgs As Engine.BE.BusinessProcesses.ProcessExceptionArgs) As Engine.BE.BusinessProcesses.OnExceptionBehaviour
        ProcessServer.ExecuteTask(Of Boolean)(AddressOf Comunes.RollbackTransaction, True, New ServiceProvider())

        Dim alog As AlbaranLogProcess = exceptionArgs.Services.GetService(Of AlbaranLogProcess)()
        If alog.CreateData Is Nothing Then alog.CreateData = New LogProcess
        ReDim Preserve alog.CreateData.Errors(alog.CreateData.Errors.Length)
        If TypeOf exceptionArgs.TaskData Is DocumentoAlbaranVenta Then
            Dim alb As AlbCabVenta = CType(exceptionArgs.TaskData, DocumentoAlbaranVenta).Cabecera
            If TypeOf alb Is AlbCabVentaAlquiler Then
                alog.CreateData.Errors(alog.CreateData.Errors.Length - 1) = New ClassErrors(" Alquiler: " & alb.NOrigen, exceptionArgs.Exception.Message)
            End If
        Else
            alog.CreateData.Errors(alog.CreateData.Errors.Length - 1) = New ClassErrors(String.Empty, exceptionArgs.Exception.Message)
        End If

        Return MyBase.OnException(exceptionArgs)
    End Function

End Class