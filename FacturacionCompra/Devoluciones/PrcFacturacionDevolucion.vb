Public Class PrcFacturacionDevolucion
    Inherits Process(Of DataPrcFacturacionDevolucion, CreateElement)

    'Lista de tareas a ejecutar
     Public Overrides Sub RegisterTasks()
        Me.AddTask(Of DataPrcFacturacionDevolucion, FraCabCompra)(AddressOf DatosIniciales)
        Me.AddTask(Of FraCabCompra, DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CrearDocumentoFactura)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarValoresPredeterminadosGenerales)
        'Me.AddTask(Of DocumentoFacturaCompra)(AddressOf AsignarValoresDevolucion)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoComunes.AsignarContador)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf AsignarFechaFactura)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf AsignarTipoAsiento)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarProveedorGrupo)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarDatosProveedor)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarDireccion)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarBanco)

        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoCompra.AsignarObservacionesCompra)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoComunes.AsignarCentroGestion)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarNumeroFactura)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoCompra.AsignarEjercicio)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarSuFactura)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoComunes.AsignarCambiosMoneda)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf NegocioGeneral.AsignarEnvio347Doc)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf NegocioGeneral.AsignarEnvio349Doc)

        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf CrearLineaDevolucion)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularImporteLineasFacturas)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf NegocioGeneral.CalcularImpuestos)
        'Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.FacturaAnaliticaAlbaran)

        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularBasesImponibles)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularTotales)

        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularVencimientos)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarClaveOperacion)
        'Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.ValidarDocumento)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf AsignarFacturaContabilizada)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.GrabarDocumento)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf AsignarFacturaDevoluciones)
        Me.AddTask(Of DocumentoFacturaCompra, CreateElement)(AddressOf Resultado)
    End Sub

    'Registar en el services la información que vamos a compartir durante el proceso, para evitar accesos innecesarios
    <Task()> Public Shared Function DatosIniciales(ByVal data As DataPrcFacturacionDevolucion, ByVal services As ServiceProvider) As FraCabCompra
        ProcessServer.ExecuteTask(Of DataPrcFacturacionDevolucion)(AddressOf ValidacionesPrevias, data, services)

        Dim TipoLineaDef As String = ProcessServer.ExecuteTask(Of Object, String)(AddressOf TipoLinea.TipoLineaPorDefecto, Nothing, services)
        services.RegisterService(New ProcessInfoFraDevolucion(data.IDContador, TipoLineaDef, data.FechaFactura, data.SuFactura, data.SuFechaFactura))

        services.RegisterService(data, GetType(DataPrcFacturacionDevolucion))

        Dim header As New BusinessData
        header("IDProveedor") = data.IDProveedor
        header("IDContador") = data.IDContador
        header("NFactura") = data.NFactura
        header("SuFactura") = data.SuFactura
        header("SuFechaFactura") = data.SuFechaFactura
        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
        header("IDMoneda") = Monedas.MonedaA.ID
        Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
        Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(data.IDProveedor)
        If Not ProvInfo Is Nothing AndAlso Length(ProvInfo.IDProveedor) > 0 Then
            header("IDFormaPago") = ProvInfo.IDFormaPago
            header("IDCondicionPago") = ProvInfo.IDCondicionPago
            header("IDCentroGestion") = ProvInfo.IDCentroGestion
        End If

        Return New FraCabCompra(header)
    End Function

    <Task()> Public Shared Sub ValidacionesPrevias(ByVal data As DataPrcFacturacionDevolucion, ByVal services As ServiceProvider)
        'ProcessServer.ExecuteTask(Of String)(AddressOf ProcesoComunes.ValidarContadorObligatorio, data.IDContador & String.Empty, services)
        ProcessServer.ExecuteTask(Of Date)(AddressOf ValidarFechaDeclaracion, data.FechaFactura, services)
        Dim datosValCont As New ProcesoComunes.DataValidarContadorEntidad(data.IDContador, GetType(FacturaCompraCabecera).Name)
        ProcessServer.ExecuteTask(Of ProcesoComunes.DataValidarContadorEntidad)(AddressOf ProcesoComunes.ValidarContadorEntidad, datosValCont, services)
    End Sub

    <Task()> Public Shared Sub ValidarFechaDeclaracion(ByVal FechaFactura As Date, ByVal services As ServiceProvider)
        If Nz(FechaFactura, cnMinDate) <> cnMinDate Then
            Dim Ejercicio As String = ProcessServer.ExecuteTask(Of Date, String)(AddressOf NegocioGeneral.EjercicioPredeterminado, FechaFactura, services)
            Dim Ej As BusinessHelper = BusinessHelper.CreateBusinessObject("EjercicioContable")
            Dim dtEjercicio As DataTable = Ej.SelOnPrimaryKey(Ejercicio)
            If dtEjercicio.Rows.Count > 0 Then
                Dim UltimaFechaDeclaracion As Date = Nz(dtEjercicio.Rows(0)("UltimaFechaDeclaracion"), cnMinDate)
                If UltimaFechaDeclaracion >= FechaFactura Then
                    ApplicationService.GenerateError("La Fecha de la factura debe ser posterior a la Fecha de Ultima Declaración del Ejercicio {0}.", Quoted(Ejercicio))
                End If
            End If
        End If
    End Sub

    '<Task()> Public Shared Sub AsignarValoresDevolucion(ByVal Doc As DocumentoFacturaCompra, ByVal services As ServiceProvider)
    '    ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf AsignarFechaFactura, Doc, services)
    'End Sub

    '<Task()> Public Shared Sub AsignarContador(ByVal Doc As DocumentoFacturaCompra, ByVal services As ServiceProvider)
    '    Dim InfoProcess As ProcessInfoFraDevolucion = services.GetService(Of ProcessInfoFraDevolucion)()
    '    If Length(InfoProcess.IDContador) > 0 Then Doc.HeaderRow("IDContador") = InfoProcess.IDContador
    'End Sub

    <Task()> Public Shared Sub AsignarFechaFactura(ByVal Doc As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        Dim InfoProcess As ProcessInfoFraDevolucion = services.GetService(Of ProcessInfoFraDevolucion)()
        'Doc.HeaderRow("FechaFactura") = InfoProcess.FechaFactura
        Dim dr As DataRow = Doc.HeaderRow
        dr = New FacturaCompraCabecera().ApplyBusinessRule("FechaFactura", InfoProcess.FechaFactura, dr, Nothing)
        If Length(InfoProcess.SuFechaFactura) > 0 Then Doc.HeaderRow("SuFechaFactura") = InfoProcess.SuFechaFactura
    End Sub

    <Task()> Public Shared Sub AsignarTipoAsiento(ByVal Doc As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        Doc.HeaderRow("IDTipoAsiento") = enumTipoAsiento.taBancoSinPago
    End Sub

    <Task()> Public Shared Sub AsignarFacturaContabilizada(ByVal Doc As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        Doc.HeaderRow("NoDescontabilizar") = True
        Doc.HeaderRow("Estado") = enumContabilizado.Contabilizado
    End Sub

    <Task()> Public Shared Sub CrearLineaDevolucion(ByVal Doc As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        Dim datosUsuario As DataPrcFacturacionDevolucion = services.GetService(Of DataPrcFacturacionDevolucion)()

        Dim lineas As DataTable = Doc.dtLineas
        Dim FCL As New FacturaCompraLinea
        If lineas Is Nothing Then
            lineas = FCL.AddNew
            Doc.Add(GetType(FacturaCompraLinea).Name, lineas)
        End If

        If datosUsuario.Precio <> 0 Then
            Dim context As New BusinessData(Doc.HeaderRow)
            Dim linea As DataRow = lineas.NewRow
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoFacturacionCompra.AsignarValoresPredeterminadosLinea, linea, services)

            linea("IDFactura") = Doc.HeaderRow("IDFactura")
            '       linea("NFactura") = Doc.HeaderRow("NFactura")
            linea("IDArticulo") = datosUsuario.IDArticulo
            linea = FCL.ApplyBusinessRule("IDArticulo", datosUsuario.IDArticulo, linea, context)
            Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
            linea("IDCentroGestion") = AppParams.CentroGestion
            linea = FCL.ApplyBusinessRule("IDTipoIva", datosUsuario.IDTipoIVA, linea, context)
            linea = FCL.ApplyBusinessRule("Cantidad", 1, linea, context)
            linea = FCL.ApplyBusinessRule("QInterna", 1, linea, context)
            linea("Dto1") = 0 : linea("Dto2") = 0 : linea("Dto3") = 0
            linea("Dto") = 0 : linea("DtoProntoPago") = 0
            linea = FCL.ApplyBusinessRule("Precio", datosUsuario.Precio, linea, context)
            linea = FCL.ApplyBusinessRule("CContable", datosUsuario.CContable, linea, context)
            lineas.Rows.Add(linea.ItemArray)
        End If
    End Sub

    <Task()> Public Shared Sub AsignarFacturaDevoluciones(ByVal Doc As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        'Doc.HeaderRow("IDFactura") '

        Dim datosUsuario As DataPrcFacturacionDevolucion = services.GetService(Of DataPrcFacturacionDevolucion)()
        If Not datosUsuario.IDDevoluciones Is Nothing AndAlso datosUsuario.IDDevoluciones.Length > 0 Then
            Dim IDDevolObject(datosUsuario.IDDevoluciones.Length) As Object

            datosUsuario.IDDevoluciones.CopyTo(IDDevolObject, 0)

            Dim dtCobrosDevol As DataTable = New CobroDevolucion().Filter(New InListFilterItem("IDDevolucion", IDDevolObject, FilterType.Numeric))
            If dtCobrosDevol.Rows.Count > 0 Then
                For Each drDevol As DataRow In dtCobrosDevol.Rows
                    drDevol("IDFacturaCompra") = Doc.HeaderRow("IDFactura")
                Next
            End If
            BusinessHelper.UpdateTable(dtCobrosDevol)
        End If
    End Sub

    <Task()> Public Shared Function Resultado(ByVal Doc As DocumentoFacturaCompra, ByVal services As ServiceProvider) As CreateElement
        Dim result As New CreateElement
        result.IDElement = Doc.HeaderRow("IDFactura")
        result.NElement = Doc.HeaderRow("NFactura")
        Return result
    End Function

End Class



