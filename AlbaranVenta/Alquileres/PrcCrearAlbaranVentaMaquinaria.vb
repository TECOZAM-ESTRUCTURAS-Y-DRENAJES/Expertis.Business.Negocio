Public Class PrcCrearAlbaranVentaMaquinaria
    Inherits Process(Of dataPrcAlbaranVentaMaquinaria, AlbaranLogProcess)

    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of dataPrcAlbaranVentaMaquinaria, AlbCabVentaObras)(AddressOf DatosIniciales)
        Me.AddTask(Of AlbCabVentaObras, DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.CrearDocumentoAlbaranVenta)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVentaAlquiler.AsignarValoresPredeterminadosGeneralesVentaMaquinaria)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVentaObras.AsignarDatosCliente)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVentaObras.AsignarCentroGestion)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVentaObras.AsignarPedidoCliente)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.AsignarDireccion)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVentaObras.AsignarTexto)
        'Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVentaAlquiler.AsignarAlmacenCabecera)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.AsignarContador)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoComunes.AsignarNumeroAlbaran)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoComercial.AsignarEjercicio)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoComunes.AsignarCambiosMoneda)

        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVentaAlquiler.CrearLineasVentaMaquinaria)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.CalcularBasesImponibles)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoComunes.TotalDocumento)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.TotalPesos)

        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.GrabarDocumento)

        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.ActualizacionAutomaticaStock)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ActualizarActivos)

        Me.AddTask(Of DocumentoAlbaranVenta, AlbaranLogProcess)(AddressOf Resultado)
    End Sub

    <Task()> Public Shared Function DatosIniciales(ByVal data As dataPrcAlbaranVentaMaquinaria, ByVal services As ServiceProvider) As AlbCabVentaMaquinaria
        Dim IDContador As String = data.IDContador
        If Len(IDContador) = 0 AndAlso Len(data.IDCentroGestion) > 0 Then
            Dim dtCG As DataTable = New CentroGestion().SelOnPrimaryKey(data.IDCentroGestion)
            If Not dtCG Is Nothing AndAlso dtCG.Rows.Count > 0 Then
                If Length(dtCG.Rows(0)("IDContadorAlbaranVenta")) > 0 Then IDContador = dtCG.Rows(0)("IDContadorAlbaranVenta")
            End If
        End If

        Dim OC As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraCabecera")
        Dim drObra As DataRow = OC.GetItemRow(data.IDObra)
        Dim IDFormaPago As String = drObra("IDFormaPago") & String.Empty
        Dim IDCondicionPago As String = drObra("IDCondicionPago") & String.Empty
        Dim IDCondicionEnvio As String = drObra("IDCondicionEnvio") & String.Empty
        Dim IDMoneda As String = drObra("IDMoneda") & String.Empty
        Dim IDDireccion As Integer = Nz(drObra("IDDireccion"), 0)

        Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
        Dim Cliente As ClienteInfo = Clientes.GetEntity(data.IDCliente)

        If Len(IDFormaPago) = 0 Then IDFormaPago = Cliente.FormaPago
        If Len(IDCondicionPago) = 0 Then IDCondicionPago = Cliente.CondicionPago
        If Len(IDCondicionEnvio) = 0 Then IDCondicionEnvio = Cliente.CondicionEnvio
        If Len(IDMoneda) = 0 Then IDMoneda = Cliente.Moneda
        Dim IDFormaEnvio As String = Cliente.FormaEnvio
        Dim IDModoTransporte As String = Cliente.IDModoTransporte
        If IDDireccion > 0 Then
            Dim infoDireccion As New ClienteDireccion.DataDirecEnvio(data.IDCliente, enumcdTipoDireccion.cdDireccionEnvio)
            Dim dtClienteDireccion As DataTable = ProcessServer.ExecuteTask(Of ClienteDireccion.DataDirecEnvio, DataTable)(AddressOf ClienteDireccion.ObtenerDireccionEnvio, infoDireccion, services)
            If Not dtClienteDireccion Is Nothing AndAlso dtClienteDireccion.Rows.Count > 0 Then
                IDDireccion = Nz(dtClienteDireccion.Rows(0)("IDDireccion"), 0)
            End If
        End If

        Dim dtHeader As New DataTable
        dtHeader.Columns.Add("IDContador", GetType(String))
        dtHeader.Columns.Add("IDObra", GetType(Integer))
        dtHeader.Columns.Add("NObra", GetType(String))
        dtHeader.Columns.Add("IDCliente", GetType(String))
        dtHeader.Columns.Add("FechaAlbaran", GetType(Date))
        dtHeader.Columns.Add("IDFormaPago", GetType(String))
        dtHeader.Columns.Add("IDCondicionPago", GetType(String))
        dtHeader.Columns.Add("IDFormaEnvio", GetType(String))
        dtHeader.Columns.Add("IDCondicionEnvio", GetType(String))
        dtHeader.Columns.Add("IDMoneda", GetType(String))
        dtHeader.Columns.Add("IDDireccion", GetType(Integer))
        dtHeader.Columns.Add("IDCentroGestion", GetType(String))
        dtHeader.Columns.Add("PedidoCliente", GetType(String))

        Dim drHeader As DataRow = dtHeader.NewRow
        drHeader("IDContador") = IDContador
        drHeader("IDObra") = data.IDObra
        drHeader("NObra") = data.NObra
        drHeader("IDCliente") = data.IDCliente
        drHeader("FechaAlbaran") = data.FechaVenta
        drHeader("IDFormaPago") = IDFormaPago
        drHeader("IDCondicionPago") = IDCondicionPago
        drHeader("IDFormaEnvio") = IDFormaEnvio
        drHeader("IDCondicionEnvio") = IDCondicionEnvio
        drHeader("IDMoneda") = IDMoneda
        drHeader("IDDireccion") = IDDireccion
        drHeader("IDCentroGestion") = data.IDCentroGestion
        'Dim moneda As MonedaInfo = New Moneda().ObtenerMoneda(IDMoneda, date.today)
        'If Not moneda Is Nothing Then
        '    drHeader("CambioA") = moneda.CambioA
        '    drHeader("CambioB") = moneda.CambioB
        'End If
        If Len(data.PedidoCliente) > 0 Then drHeader("PedidoCliente") = data.PedidoCliente

        Dim AVM As New AlbCabVentaMaquinaria(drHeader)

        Dim IDTipoAlbaran As String = New Parametro().TipoAlbaranPorDefecto
        services.RegisterService(New ProcessInfoAV(data.IDContador, IDTipoAlbaran, data.FechaVenta, enumTipoExpedicion.teObra))

        Dim TipoInfo As ProcesoAlbaranVenta.TipoAlbaranInfo = ProcessServer.ExecuteTask(Of String, ProcesoAlbaranVenta.TipoAlbaranInfo)(AddressOf ProcesoAlbaranVenta.TipoDeAlbaran, IDTipoAlbaran, services)
        AVM.IDTipoAlbaran = TipoInfo.IDTipo

        For Each drActivo As DataRow In data.dtActivos.Select
            data.IDActivo = drActivo("IDActivo")
            Dim drDatosLinea As DataRow = ProcessServer.ExecuteTask(Of dataPrcAlbaranVentaMaquinaria, DataRow)(AddressOf DatosInicialesLinea, data, services)

            ReDim Preserve AVM.LineasOrigen(AVM.LineasOrigen.Length)
            Dim linMaq As New AlbLinVentaMaquinaria(drDatosLinea)
            linMaq.IDActivo = drActivo("IDActivo")
            linMaq.Precio = Nz(drActivo("Precio"), 0)
            linMaq.Dto1 = Nz(drActivo("Dto1"), 0)
            linMaq.IDObra = data.IDObra
            linMaq.IDTrabajo = data.IDTrabajo
            linMaq.TipoFactAlquiler = data.TipoFactAlquiler

            AVM.LineasOrigen(AVM.LineasOrigen.Length - 1) = linMaq
        Next

        AVM.dtActivos = data.dtActivos

        Return AVM
    End Function

    <Task()> Public Shared Function DatosInicialesLinea(ByVal data As dataPrcAlbaranVentaMaquinaria, ByVal services As ServiceProvider) As DataRow
        Dim dr As DataRow = Nothing

        If data.IDTrabajo > 0 Then
            Dim f As New Filter
            f.Add(New NumberFilterItem("IDObra", data.IDObra))
            f.Add(New NumberFilterItem("IDTrabajo", data.IDTrabajo))
            f.Add(New StringFilterItem("Lote", data.IDActivo))

            Dim OM As BusinessHelper= BusinessHelper.CreateBusinessObject("ObraMaterial")
            Dim dt As DataTable = OM.Filter(f)
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
            Else
                dt = New BE.DataEngine().Filter("vAlquilerNegVentaMaquinariaDatosActivo", New StringFilterItem("IDActivo", data.IDActivo))
                If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                    dr = dt.Rows(0)
                End If
            End If
        End If

        Return dr
    End Function

    <Task()> Public Shared Sub ActualizarActivos(ByVal doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        Dim data As AlbCabVentaMaquinaria = doc.Cabecera
        data.dtActivos.Columns.Add("IDEstadoActivo", GetType(String))
        For Each drActivo As DataRow In data.dtActivos.Select
            drActivo("IDEstadoActivo") = NegocioGeneral.ESTADOACTIVO_VENDIDO
        Next
        ProcessServer.ExecuteTask(Of DataTable)(AddressOf ProcesoAlbaranVentaAlquiler.ActualizarActivos, data.dtActivos, services)
    End Sub

    <Task()> Public Shared Function Resultado(ByVal Doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider) As AlbaranLogProcess
        Dim alog As AlbaranLogProcess = services.GetService(Of AlbaranLogProcess)()
        If alog.CreateData Is Nothing Then alog.CreateData = New LogProcess
        ReDim Preserve alog.CreateData.CreatedElements(alog.CreateData.CreatedElements.Length)

        alog.CreateData.CreatedElements(alog.CreateData.CreatedElements.Length - 1) = New CreateElement
        alog.CreateData.CreatedElements(alog.CreateData.CreatedElements.Length - 1).IDElement = Doc.HeaderRow("IDAlbaran")
        alog.CreateData.CreatedElements(alog.CreateData.CreatedElements.Length - 1).NElement = Doc.HeaderRow("NAlbaran")
        Return alog
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
