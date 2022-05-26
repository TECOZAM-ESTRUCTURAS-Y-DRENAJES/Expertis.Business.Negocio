Public Class PrcFacturacionNuevoGasto
    Inherits Process(Of DataPrcFacturacionNuevoGasto, CreateElement)

    'Lista de tareas a ejecutar
    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of DataPrcFacturacionNuevoGasto, FraCabCompra)(AddressOf DatosIniciales)
        Me.AddTask(Of FraCabCompra, DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CrearDocumentoFactura)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarValoresPredeterminadosGenerales)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf AsignarTipoAsiento)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf AsignarBancoPropio)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoComunes.AsignarContador)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarProveedorGrupo)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarDatosProveedor)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarDireccion)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarBanco)

        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoCompra.AsignarObservacionesCompra)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoComunes.AsignarCentroGestion)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarNumeroFactura)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoCompra.AsignarEjercicio)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarSuFechaFactura)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarSuFactura)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoComunes.AsignarCambiosMoneda)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf NegocioGeneral.AsignarEnvio347Doc)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf NegocioGeneral.AsignarEnvio349Doc)

        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf CrearLineaNuevoGasto)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularImporteLineasFacturas)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf NegocioGeneral.CalcularImpuestos)

        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf CalcularAnalitica)

        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularBasesImponibles)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularTotales)

        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularVencimientos)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarClaveOperacion)
        'Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.ValidarDocumento)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.GrabarDocumento)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.ActualizarObras)

        Me.AddTask(Of DocumentoFacturaCompra, CreateElement)(AddressOf ProcesoFacturacionCompra.Resultado)
    End Sub

    <Task()> Public Shared Function DatosIniciales(ByVal data As DataPrcFacturacionNuevoGasto, ByVal services As ServiceProvider) As FraCabCompraNuevoGasto
        Dim TipoLineaDef As String = ProcessServer.ExecuteTask(Of Object, String)(AddressOf TipoLinea.TipoLineaPorDefecto, Nothing, services)
        Dim IDContador As String
        If Length(data.IDContador) > 0 Then
            IDContador = data.IDContador
        Else
            IDContador = ProcessServer.ExecuteTask(Of ContadorEntidad, String)(AddressOf CentroGestion.GetContadorPredeterminadoCGestionUsuario, ContadorEntidad.FacturaCompra, services)
        End If

        services.RegisterService(New ProcessInfoFraDevolucion(IDContador, TipoLineaDef, data.FechaFactura, data.SuFactura, data.SuFechaFactura))

        Dim header As New BusinessData
        header("IDProveedor") = data.IDProveedor
        'header("IDContador") = IDContador
        'header("NFactura") = data.NFactura
        header("FechaFactura") = data.FechaFactura
        header("SuFactura") = data.SuFactura
        header("SuFechaFactura") = data.SuFechaFactura
        header("IDMoneda") = data.IDMoneda
        header("IDFormaPago") = data.IDFormaPago
        header("IDCondicionPago") = data.IDCondicionPago
        Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
        Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(data.IDProveedor)
        If Not ProvInfo Is Nothing AndAlso Length(ProvInfo.IDProveedor) > 0 Then
            header("IDCentroGestion") = ProvInfo.IDCentroGestion
        End If

        header("RazonSocial") = data.RazonSocial
        header("CIF") = data.CIF
        header("IDDiaPago") = data.IDDiaPago
        header("IDBancoPropio") = data.IDBancoPropio
        header("IDTipoAsiento") = data.IDTipoAsiento

        Dim ProcInfo As ProcessInfoFra = services.GetService(Of ProcessInfoFra)()
        ProcInfo.SuFechaFactura = data.SuFechaFactura
        Dim FCC As New FraCabCompraNuevoGasto(header)

        For Each lin As DataPrcFacturacionLineaNuevoGasto In data.Lineas
            Dim FCL As New FraLinCompraNuevoGasto(lin)
            FCC.lineas.Add(FCL)
        Next
        Return FCC
    End Function

    <Task()> Public Shared Sub AsignarTipoAsiento(ByVal Doc As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        Doc.HeaderRow("IDTipoAsiento") = CType(Doc.Cabecera, FraCabCompraNuevoGasto).IDTipoAsiento
    End Sub

    <Task()> Public Shared Sub AsignarBancoPropio(ByVal Doc As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        Doc.HeaderRow("IDBancoPropio") = CType(Doc.Cabecera, FraCabCompraNuevoGasto).IDBancoPropio
    End Sub
    '<Task()> Public Shared Sub AsignarContador(ByVal Doc As DocumentoFacturaCompra, ByVal services As ServiceProvider)
    '    Dim InfoProcess As ProcessInfoFraDevolucion = services.GetService(Of ProcessInfoFraDevolucion)()
    '    If Length(InfoProcess.IDContador) > 0 Then Doc.HeaderRow("IDContador") = InfoProcess.IDContador
    'End Sub

    <Task()> Public Shared Sub CrearLineaNuevoGasto(ByVal Doc As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        Dim lineas As DataTable = Doc.dtLineas
        Dim FCL As New FacturaCompraLinea
        If lineas Is Nothing Then
            lineas = FCL.AddNew
            Doc.Add(GetType(FacturaCompraLinea).Name, lineas)
        End If
        Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()

        For Each LineaFactura As FraLinCompraNuevoGasto In CType(Doc.Cabecera, FraCabCompraNuevoGasto).lineas
            If LineaFactura.Importe <> 0 Then
                Dim context As New BusinessData(Doc.HeaderRow)
                Dim linea As DataRow = lineas.NewRow
                ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoFacturacionCompra.AsignarValoresPredeterminadosLinea, linea, services)

                linea("IDFactura") = Doc.HeaderRow("IDFactura")
                'linea("NFactura") = Doc.HeaderRow("NFactura")

                linea = FCL.ApplyBusinessRule("IDArticulo", LineaFactura.IDArticulo, linea, context)
                If Length(LineaFactura.DescArticulo) > 0 Then linea("DescArticulo") = LineaFactura.DescArticulo
                If Length(LineaFactura.RefProveedor) > 0 Then linea("RefProveedor") = LineaFactura.RefProveedor
                If Length(LineaFactura.DescRefProveedor) > 0 Then linea("DescRefProveedor") = LineaFactura.DescRefProveedor


                linea("IDCentroGestion") = AppParams.CentroGestion
                linea = FCL.ApplyBusinessRule("IDTipoIva", LineaFactura.IDTipoIVA, linea, context)
                linea = FCL.ApplyBusinessRule("Cantidad", 1, linea, context)
                linea = FCL.ApplyBusinessRule("QInterna", 1, linea, context)
                linea = FCL.ApplyBusinessRule("Precio", LineaFactura.Importe, linea, context)
                linea = FCL.ApplyBusinessRule("CContable", LineaFactura.CContable, linea, context)
                If Length(LineaFactura.IDObra) > 0 Then linea("IDObra") = LineaFactura.IDObra
                If Length(LineaFactura.IDTrabajo) > 0 Then linea("IDTrabajo") = LineaFactura.IDTrabajo
                linea("Cantidad") = 1
                linea("QInterna") = 1
                linea("Precio") = LineaFactura.Importe
                linea("Importe") = linea("QInterna") * linea("Precio")
                linea("Dto1") = 0 : linea("Dto2") = 0 : linea("Dto3") = 0

                lineas.Rows.Add(linea.ItemArray)
            End If
        Next

    End Sub


    <Task()> Public Shared Sub CalcularAnalitica(ByVal doc As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        Dim AppParams As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
        If Not AppParams.Contabilidad OrElse Not AppParams.Analitica.AplicarAnalitica Then Exit Sub

        For Each LineaFactura As FraLinCompraNuevoGasto In CType(doc.Cabecera, FraCabCompraNuevoGasto).lineas
            If Not LineaFactura.Analitica Is Nothing Then
                For Each lineaAnalitica As DataRow In LineaFactura.Analitica.Rows
                    lineaAnalitica("IDLineaFactura") = doc.dtLineas.Rows(0)("IDLineaFactura")
                    doc.dtAnalitica.ImportRow(lineaAnalitica)
                Next
            End If
        Next
    End Sub

End Class
