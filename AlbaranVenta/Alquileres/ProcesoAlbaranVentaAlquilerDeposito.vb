Public Class ProcesoAlbaranVentaAlquilerDeposito

#Region " AgruparAlquiler "

    <Serializable()> _
    Public Class dataColAgrupacionAVAlquiler
        Public Lineas As DataTable
        Public TipoAgrupacion As enummcAgrupAlbaranObra

        Public Sub New()
        End Sub

        Public Sub New(ByVal Lineas As DataTable)
            Me.Lineas = Lineas
        End Sub
    End Class
    <Task()> Public Shared Function GetGroupColumnsAlquiler(ByVal data As dataColAgrupacionAVAlquiler, ByVal services As ServiceProvider) As DataColumn()
        Dim columns(data.TipoAgrupacion + 1) As DataColumn

        If data.TipoAgrupacion = enummcAgrupAlbaranObra.mcCliente Then
            columns(0) = data.Lineas.Columns("IDCliente")
            columns(1) = data.Lineas.Columns("IDAlmacenDeposito")
        Else
            columns(0) = data.Lineas.Columns("IDCliente")
            columns(1) = data.Lineas.Columns("IDObra")
            If data.TipoAgrupacion = enummcAgrupAlbaranObra.mcObraTrabajo Then
                columns(2) = data.Lineas.Columns("IDTrabajo")
                columns(3) = data.Lineas.Columns("IDAlmacenDeposito")
            Else
                columns(2) = data.Lineas.Columns("IDAlmacenDeposito")
            End If
        End If

        Dim info As ProcessInfoAV = services.GetService(Of ProcessInfoAV)()
        Dim TipoInfo As ProcesoAlbaranVenta.TipoAlbaranInfo = ProcessServer.ExecuteTask(Of String, ProcesoAlbaranVenta.TipoAlbaranInfo)(AddressOf ProcesoAlbaranVenta.TipoDeAlbaran, info.IDTipoAlbaran, services)
        If TipoInfo.Tipo = enumTipoAlbaran.RetornoAlquiler Or TipoInfo.Tipo = enumTipoAlbaran.Consumo Then
            ReDim Preserve columns(columns.Length)
            columns(columns.Length - 1) = data.Lineas.Columns("FechaAlquiler")
        End If

        Return columns
    End Function

    <Task()> Public Shared Function AgruparAlquiler(ByVal data As DataPrcAlbaranar, ByVal services As ServiceProvider) As AlbCabVentaAlquiler()
        Dim IDMaterial(-1) As Object
        Dim htLins As New Hashtable
        For Each AVInfo As CrearAlbaranVentaInfo In data.AlbVentaInfo
            If AVInfo.IDLinea > 0 Then 'IDLineaMaterial o IDLineaAlbaran
                ReDim Preserve IDMaterial(UBound(IDMaterial) + 1)
                IDMaterial(UBound(IDMaterial)) = AVInfo.IDLinea
            End If
            htLins.Add(AVInfo.IDLinea, AVInfo)
        Next

        If IDMaterial.Length > 0 Then
            Dim ProcInfo As ProcessInfoAV = services.GetService(Of ProcessInfoAV)()
            Dim cnViewName As String = "vAlquilerCIExpedicion"
            Dim strSelect As String = "IDLineaMaterial"
            Dim OrderBy As String = "IDObra DESC"
            Dim TipoAlbInfo As ProcesoAlbaranVenta.TipoAlbaranInfo = ProcesoAlbaranVenta.TipoDeAlbaran(ProcInfo.IDTipoAlbaran, services)
            If TipoAlbInfo.Tipo = enumTipoAlbaran.Consumo Or TipoAlbInfo.Tipo = enumTipoAlbaran.RetornoAlquiler Or ProcInfo.TipoExpedicion = enumTipoExpedicion.teAlquilerCambioMaquina Then
                If ProcInfo.TipoExpedicion = enumTipoExpedicion.teAlquilerCambioMaquina Then
                    cnViewName = "vAlquilerNegCambioMaquina"
                Else
                    cnViewName = "vAlquilerCIRetornos"
                End If
                If ProcInfo.TipoExpedicion <> enumTipoExpedicion.teAlquilerRetorno Then
                    strSelect = "IDLineaAlbaran"
                    OrderBy = OrderBy & ", FechaAlbaran DESC"
                End If
            End If

            Dim dtLineas As DataTable = New BE.DataEngine().Filter(cnViewName, New InListFilterItem(strSelect, IDMaterial, FilterType.Numeric), , OrderBy)
            If Not dtLineas.Columns.Contains("FechaAlquiler") Then
                dtLineas.Columns.Add("FechaAlquiler", GetType(Date))
            End If
            For Each drLinea As DataRow In dtLineas.Select()
                drLinea("FechaAlquiler") = DirectCast(htLins(drLinea(strSelect)), CrearAlbaranVentaInfo).FechaAlquiler
                'ALMACEN DE SALIDA DAVID V 3/3/22
                drLinea("IDAlmacen") = "T642"

            Next

            If Not dtLineas Is Nothing AndAlso dtLineas.Rows.Count > 0 Then
                Dim Agrup As New dataColAgrupacionAVAlquiler(dtLineas)
                Agrup.TipoAgrupacion = enummcAgrupAlbaranObra.mcCliente
                Dim ColsAgrupCliente As DataColumn() = ProcessServer.ExecuteTask(Of dataColAgrupacionAVAlquiler, DataColumn())(AddressOf GetGroupColumnsAlquiler, Agrup, services)

                Agrup.TipoAgrupacion = enummcAgrupAlbaranObra.mcObra
                Dim ColsAgrupObra As DataColumn() = ProcessServer.ExecuteTask(Of dataColAgrupacionAVAlquiler, DataColumn())(AddressOf GetGroupColumnsAlquiler, Agrup, services)

                Agrup.TipoAgrupacion = enummcAgrupAlbaranObra.mcObraTrabajo
                Dim ColsAgrupObraTrabajo As DataColumn() = ProcessServer.ExecuteTask(Of dataColAgrupacionAVAlquiler, DataColumn())(AddressOf GetGroupColumnsAlquiler, Agrup, services)

                Dim oGrprUser As New GroupUserAlquiler
                Dim groupers(2) As GroupHelper
                groupers(enummcAgrupAlbaranObra.mcCliente) = New GroupHelper(ColsAgrupCliente, oGrprUser)
                groupers(enummcAgrupAlbaranObra.mcObra) = New GroupHelper(ColsAgrupObra, oGrprUser)
                groupers(enummcAgrupAlbaranObra.mcObraTrabajo) = New GroupHelper(ColsAgrupObraTrabajo, oGrprUser)

                Dim sort As String = "IDCliente, IDObra, IDCentroGestion, IDLineaMaterial"
                If ProcInfo.TipoExpedicion = enumTipoExpedicion.teAlquilerRetorno Then
                    sort = "IDCliente, IDObra, IDDireccionOT, IDLineaAlbaran"
                End If
                For Each drLinea As DataRow In dtLineas.Select(Nothing, sort)
                    groupers(drLinea("AgrupAlbaranObra")).Group(drLinea)
                Next

                For Each alb As AlbCabVentaAlquiler In oGrprUser.Albs
                    For Each alblin As AlbLinVentaAlquiler In alb.LineasOrigen
                        alblin.QaServir = DirectCast(htLins(alblin.IDLineaOrigen), CrearAlbaranVentaInfo).Cantidad
                        alblin.Lotes = DirectCast(htLins(alblin.IDLineaOrigen), CrearAlbaranVentaInfo).Lotes
                        alblin.Series = DirectCast(htLins(alblin.IDLineaOrigen), CrearAlbaranVentaInfo).Series
                        alblin.FechaPrevistaRetorno = DirectCast(htLins(alblin.IDLineaOrigen), CrearAlbaranVentaInfo).FechaPrevistaRetorno
                        alblin.IDEstadoActivo = DirectCast(htLins(alblin.IDLineaOrigen), CrearAlbaranVentaInfo).IDEstadoActivo
                        alblin.HoraAlquiler = DirectCast(htLins(alblin.IDLineaOrigen), CrearAlbaranVentaInfo).HoraAlquiler
                    Next
                Next

                Return oGrprUser.Albs
            End If
        End If
    End Function

#End Region

#Region " Cabecera "

    <Task()> Public Shared Sub AsignarNumeroAlbaranAlquilerPropuesta(ByVal doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        If doc.HeaderRow.IsNull("NAlbaran") Then
            Dim counters As ProvisionalCounter = services.GetService(Of ProvisionalCounter)()
            doc.HeaderRow("NAlbaran") = counters.GetCounterValue(doc.HeaderRow("IDContador"))
        End If
    End Sub

    <Task()> Public Shared Sub AsignarAlmacenCabecera(ByVal doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        doc.HeaderRow("IDAlmacen") = CType(doc.Cabecera, AlbCabVentaAlquiler).IDAlmacen
        Dim ProcInfo As ProcessInfoAV = services.GetService(Of ProcessInfoAV)()
        Dim EsDeposito As Boolean = False

        If ProcInfo.Tipo = enumTipoAlbaran.RetornoAlquiler OrElse ProcInfo.Tipo = enumTipoAlbaran.Consumo OrElse ProcInfo.Tipo = enumTipoAlbaran.Deposito Then
            If ProcInfo.Tipo = enumTipoAlbaran.Consumo Then
                doc.HeaderRow("IDAlmacenDeposito") = CType(doc.Cabecera, AlbCabVentaAlquiler).IDAlmacen
            Else
                doc.HeaderRow("IDAlmacenDeposito") = CType(doc.Cabecera, AlbCabVentaAlquiler).IDAlmacenDeposito
            End If
            EsDeposito = (Length(doc.HeaderRow("IDAlmacenDeposito")) > 0 And CType(doc.Cabecera, AlbCabVentaAlquiler).Deposito)
        Else
            doc.HeaderRow("IDAlmacenDeposito") = CType(doc.Cabecera, AlbCabVentaAlquiler).IDAlmacenTransferencia
            EsDeposito = True
        End If
        If Length(doc.HeaderRow("IDAlmacen")) > 0 Then
            If EsDeposito Then
                If ProcInfo.Tipo <> enumTipoAlbaran.Deposito Then
                    doc.HeaderRow("IDAlmacen") = doc.HeaderRow("IDAlmacenDeposito")
                End If
            Else
                ApplicationService.GenerateError("El almacén | asignado a la dirección de envío del cliente | no es de depósito", Quoted(doc.HeaderRow("IDAlmacenDeposito")), Quoted(doc.HeaderRow("IDCliente")))
            End If
        Else
            ApplicationService.GenerateError("La dirección de envio asignada al Cliente | no tine un almacén asignado.", Quoted(doc.HeaderRow("IDCliente")))
        End If
    End Sub

    <Task()> Public Shared Sub AsignarEstado(ByVal doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        Dim ProcInfo As ProcessInfoAV = services.GetService(Of ProcessInfoAV)()
        If ProcInfo.Tipo = enumTipoAlbaran.RetornoAlquiler OrElse ProcInfo.Tipo = enumTipoAlbaran.Deposito Then
            doc.HeaderRow("Estado") = enumavcEstadoFactura.avcNoFacturable
        ElseIf ProcInfo.Tipo = enumTipoAlbaran.Consumo Then
            doc.HeaderRow("Estado") = enumavcEstadoFactura.avcNoFacturado
        End If
    End Sub

#End Region

#Region " Lineas "

    <Task()> Public Shared Sub AsignarDatosAlquiler(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        Dim ProcInfo As ProcessInfoAV = services.GetService(Of ProcessInfoAV)()
        If ProcInfo.Tipo = enumTipoAlbaran.RetornoAlquiler OrElse ProcInfo.Tipo = enumTipoAlbaran.Consumo Then
            data.Row("IDLineaAlbaranDeposito") = data.AlbLin.IDLineaOrigen
            data.Row("IDAlbaranDeposito") = data.Doc.Cabecera.IDOrigen
        Else
            data.Row("QPendienteDevolverAInicio") = data.Cantidad
        End If
        data.Row("IDAlmacen") = CType(data.AlbLin, AlbLinVentaAlquiler).IDAlmacen
    End Sub

    <Task()> Public Shared Sub AsignarFechaHoraAlquiler(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        data.Row("FechaAlquiler") = CType(data.Doc.Cabecera, AlbCabVentaAlquiler).Fecha
        data.Row("HoraAlquiler") = CType(data.AlbLin, AlbLinVentaAlquiler).HoraAlquiler
    End Sub

    <Task()> Public Shared Sub AsignarTipoFacturacionAlquiler(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        data.Row("TipoFactAlquiler") = data.Origen("TipoFactAlquiler")
    End Sub

    <Task()> Public Shared Sub AsignarConsumo(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        Dim ProcInfo As ProcessInfoAV = services.GetService(Of ProcessInfoAV)()
        data.Row("ConsumoAlquiler") = (ProcInfo.Tipo = enumTipoAlbaran.Consumo)
    End Sub

    <Task()> Public Shared Sub AsignarCContableConsumo(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        Dim ProcInfo As ProcessInfoAV = services.GetService(Of ProcessInfoAV)()
        If ProcInfo.Tipo = enumTipoAlbaran.Consumo Then
            Dim AppParams As ParametroAlquiler = services.GetService(Of ParametroAlquiler)()
            data.Row("CContable") = AppParams.CContableMaterialAlquiler
        End If
    End Sub

    <Task()> Public Shared Sub AsignarPreciosDtosImportes(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        Dim Precio As Double = 0
        Dim ProcInfo As ProcessInfoAV = services.GetService(Of ProcessInfoAV)()
        If ProcInfo.Tipo = enumTipoAlbaran.Consumo Then
            Precio = data.Origen("ValorReposicionA")
        Else
            Precio = data.Origen("PrecioPrevMatA")
        End If

        Precio = (Precio / data.Doc.CambioA) * Nz(data.Row("Factor"), 1)
        data.Row("Dto1") = data.Origen("Dto1")
        data.Row("Dto2") = data.Origen("Dto2")
        data.Row("Dto3") = data.Origen("Dto3")

        Dim AVL As New AlbaranVentaLinea
        Dim context As New BusinessData(data.Doc.HeaderRow)
        AVL.ApplyBusinessRule("Precio", Precio, data.Row, context)

        Dim ValAyB As New ValoresAyB(New DataRowPropertyAccessor(data.Row), data.Doc.IDMoneda, data.Doc.CambioA, data.Doc.CambioB)
        ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)
    End Sub

    <Task()> Public Shared Sub AsignarFechaPrevistaRetorno(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        If CType(data.AlbLin, AlbLinVentaAlquiler).FechaPrevistaRetorno <> cnMinDate Then
            data.Row("FechaPrevistaRetorno") = CType(data.AlbLin, AlbLinVentaAlquiler).FechaPrevistaRetorno
        End If
    End Sub

    <Task()> Public Shared Sub AsignarFacturable(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        If Nz(data.Origen("TipoFacturacion"), enumotTipoFacturacion.otfPorConceptos) = enumotTipoFacturacion.otfPorVencimientos Then
            data.Row("Facturable") = False
        Else
            data.Row("Facturable") = data.Origen("Facturable")
        End If
    End Sub

    <Task()> Public Shared Sub AsignarDatosNSerie(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        If Length(data.Origen("Lote")) > 0 Then
            data.Row("Lote") = data.Origen("Lote")
            Dim f As New Filter
            f.Add(New StringFilterItem("IDArticulo", data.Origen("IDMaterial")))
            f.Add(New StringFilterItem("NSerie", data.Origen("Lote")))

            Dim dtSerie As DataTable = New ArticuloNSerie().Filter(f)
            If Not dtSerie Is Nothing AndAlso dtSerie.Rows.Count > 0 Then
                Dim ProcInfo As ProcessInfoAV = services.GetService(Of ProcessInfoAV)()
                If ProcInfo.Tipo = enumTipoAlbaran.Deposito Then
                    If data.Origen("TipoFactAlquiler") = enumTipoFacturacionAlquiler.enumTFASinAlquiler Then
                        data.Row("IDEstadoActivo") = NegocioGeneral.ESTADOACTIVO_VENDIDO
                    Else
                        data.Row("IDEstadoActivo") = NegocioGeneral.ESTADOACTIVO_TRABAJANDO
                    End If
                ElseIf ProcInfo.Tipo = enumTipoAlbaran.RetornoAlquiler Then
                    data.Row("IDEstadoActivo") = CType(data.AlbLin, AlbLinVentaAlquiler).IDEstadoActivo
                Else
                    data.Row("IDEstadoActivo") = NegocioGeneral.ESTADOACTIVO_VENDIDO
                End If

                data.Row("IDEstadoActivoAnterior") = dtSerie.Rows(0)("IDEstadoActivo")
                data.Row("IDOperario") = dtSerie.Rows(0)("IDOperario")
                If dtSerie.Columns.Contains("Ubicacion") Then data.Row("Ubicacion") = dtSerie.Rows(0)("Ubicacion")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarFechaRetornoDiasMinimos(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        Dim ProcInfo As ProcessInfoAV = services.GetService(Of ProcessInfoAV)()
        If ProcInfo.Tipo = enumTipoAlbaran.Deposito Then
            Dim Parametros As ParametroAlquiler = services.GetService(Of ParametroAlquiler)()
            Dim dtmLimHora As Date = CDate(Format(CDate(Nz(Parametros.LimiteHoraAlquiler, 0)), "HH:mm:ss"))
            Dim datadiasMinimos As New GeneralAlquiler.dataFechaRetornoDiasMinimos(data.Row("FechaAlquiler"), data.Row("TipoFactAlquiler"), _
                                                                                   data.Row("IDArticulo"), Nz(data.Row("IDObra"), 0), _
                                                                                   Nz(data.Row("HoraAlquiler"), dtmLimHora), dtmLimHora)
            data.Row("FechaRetornoDiasMinimos") = ProcessServer.ExecuteTask(Of GeneralAlquiler.dataFechaRetornoDiasMinimos, Date)(AddressOf GeneralAlquiler.ObtenerFechaRetornoDiasMinimos, datadiasMinimos, services)
        End If
    End Sub

#End Region

#Region " Venta Maquinaria "

    <Task()> Public Shared Sub AsignarValoresPredeterminadosGeneralesVentaMaquinaria(ByVal doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        Dim TipoInfo As ProcesoAlbaranVenta.TipoAlbaranInfo = ProcessServer.ExecuteTask(Of String, ProcesoAlbaranVenta.TipoAlbaranInfo)(AddressOf ProcesoAlbaranVenta.TipoDeAlbaran, New Parametro().TipoAlbaranPorDefecto, services)
        doc.HeaderRow("IDTipoAlbaran") = TipoInfo.IDTipo

        ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoComunes.AsignarIdentificadorAlbaran, doc.HeaderRow, services)
    End Sub

#Region " Lineas Venta Maquinaria "

    <Task()> Public Shared Sub CrearLineasVentaMaquinaria(ByVal doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        Dim Cantidad As Double = 1
        For i As Integer = 0 To doc.Cabecera.LineasOrigen.Length - 1
            Dim AlbLin As AlbLinVentaMaquinaria = doc.Cabecera.LineasOrigen(i)

            Dim linea As DataRow = doc.dtLineas.NewRow
            linea("IDAlbaran") = doc.HeaderRow("IDAlbaran")

            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoAlbaranVenta.AsignarValoresPredeterminadosLinea, linea, services)
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarValoresPredeterminadosLineasVentaMaquinaria, linea, services)

            Dim linAlbPed As New DataLineasAVDesdeOrigen(linea, AlbLin.row, doc, AlbLin, Cantidad)

            If AlbLin.OrigenDatos = AlbLinVentaMaquinaria.enumOrigenDatosLineaVentaMaquinaria.ObraMaterial Then
                ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf AsignarDatosDesdeObraMaterial, linAlbPed, services)
            Else
                ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf AsignarDatosDesdeActivo, linAlbPed, services)
            End If
            ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf AsignarActivo, linAlbPed, services)
            ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf AsignarAlmacen, linAlbPed, services)
            ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf ProcesoAlbaranVentaObras.AsignarTipoIVA, linAlbPed, services)
            ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf ProcesoAlbaranVentaObras.AsignarCondicionesEconomicas, linAlbPed, services)
            ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf ProcesoAlbaranVentaObras.AsignarCentroGestionLinea, linAlbPed, services)
            ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf ProcesoAlbaranVentaObras.AsignarPedidoClienteDeCabecera, linAlbPed, services)
            ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf AsignarDatosEconomicos, linAlbPed, services)

            doc.dtLineas.Rows.Add(linea.ItemArray)
        Next
    End Sub

    <Task()> Public Shared Sub AsignarValoresPredeterminadosLineasVentaMaquinaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        data("IDEstadoActivo") = NegocioGeneral.ESTADOACTIVO_VENDIDO
        data("IDOperario") = ProcessServer.ExecuteTask(Of Object, String)(AddressOf Operario.ObtenerIDOperarioUsuario, Nothing, services)
        data("IDTipoLinea") = ProcessServer.ExecuteTask(Of Object, String)(AddressOf TipoLinea.TipoLineaPorDefecto, Nothing, services)
    End Sub

    <Task()> Public Shared Sub AsignarDatosDesdeObraMaterial(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        data.Row("IDArticulo") = data.Origen("IDMaterial")
        data.Row("DescArticulo") = data.Origen("DescMaterial")
        data.Row("IDLineaMaterial") = data.Origen("IDLineaMaterial")
        data.Row("IDLineaMaterialEstruAlq") = data.Origen("IDLineaMaterialEstruAlq")
        data.Row("IDMaterialEstruAlq") = data.Origen("IDMaterialEstruAlq")
        data.Row("IDUDMedida") = data.Origen("IDUDVenta")
        data.Row("IDUDInterna") = data.Origen("IDUDInterna")

        Dim dtArticulo As DataTable = New Articulo().SelOnPrimaryKey(data.Row("IDArticulo"))
        If Not dtArticulo Is Nothing AndAlso dtArticulo.Rows.Count > 0 Then
            data.Row("CContable") = Nz(dtArticulo.Rows(0)("CCVtaInmovilizado"), dtArticulo.Rows(0)("CCVenta"))
        End If
        If Length(data.Row("CContable")) = 0 Then data.Row("CContable") = data.Origen("CCVenta")

        data.Row("TipoFactAlquiler") = data.Origen("TipoFactAlquiler")
        data.Row("UDValoracion") = data.Origen("UDValoracion")
        If data.Origen("TipoFacturacion") = enumomTipoFacturacion.omPorVencimientos OrElse data.Origen("Facturable") Then
            data.Row("Facturable") = False
        Else
            data.Row("Facturable") = data.Origen("Facturable")
        End If
    End Sub

    <Task()> Public Shared Sub AsignarDatosDesdeActivo(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        data.Row("IDArticulo") = data.Origen("IDArticulo")
        data.Row("DescArticulo") = data.Origen("DescArticulo")
        data.Row("Regalo") = data.Origen("NoImprimirEnFactura")
        data.Row("IDUDMedida") = data.Origen("IDUDVenta")
        data.Row("IDUDInterna") = data.Origen("IDUDInterna")
        data.Row("Facturable") = True
        data.Row("CContable") = New Parametro().CContableMaterialAlquiler
    End Sub

    <Task()> Public Shared Sub AsignarActivo(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        data.Row("Lote") = CType(data.AlbLin, AlbLinVentaMaquinaria).IDActivo
    End Sub

    <Task()> Public Shared Sub AsignarAlmacen(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        Dim dt As DataTable = New ArticuloNSerie().SelOnPrimaryKey(data.Row("IDArticulo"), data.Row("Lote"))
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            data.Row("IDAlmacen") = dt.Rows(0)("IDAlmacen")
        End If
    End Sub

    <Task()> Public Shared Sub AsignarDatosEconomicos(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        data.Row("QInterna") = 1
        data.Row("Factor") = 1
        data.Row("QServida") = 1
        data.Row("Precio") = CType(data.AlbLin, AlbLinVentaMaquinaria).Precio
        data.Row("Dto1") = CType(data.AlbLin, AlbLinVentaMaquinaria).Dto1

        Dim AVL As New AlbaranVentaLinea
        Dim context As New BusinessData(data.Doc.HeaderRow)
        AVL.ApplyBusinessRule("Precio", data.Row("Precio"), data.Row, context)

        Dim ValAyB As New ValoresAyB(New DataRowPropertyAccessor(data.Row), data.Doc.IDMoneda, data.Doc.CambioA, data.Doc.CambioB)
        ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)
    End Sub

#End Region

#End Region

    <Task()> Public Shared Sub ActualizarIncidencias(ByVal data As dataPrcActualizarAlbaranAlquiler, ByVal services As ServiceProvider)
        BusinessHelper.UpdateTable(data.Incidencias)
    End Sub

    <Task()> Public Shared Sub ActualizarActivos(ByVal data As DataTable, ByVal services As ServiceProvider)
        If Not data Is Nothing Then
            'La actualización del estado de los activos se debe hacer después de la actualización de los stocks
            'No se puede actualizar directamente data.dtActivos al dar un error de concurrencia.
            Dim IDActivo As String = "IDActivo"
            If Not data.Columns.Contains("IDActivo") Then IDActivo = "Lote"

            Dim Values(-1) As Object
            For Each dr As DataRow In data.Rows
                ReDim Preserve Values(UBound(Values) + 1)
                Values(UBound(Values)) = dr(IDActivo)
            Next
            If Values.Length > 0 Then
                Dim ac As New Activo
                Dim dtActivosModif As DataTable = ac.Filter(New InListFilterItem("IDActivo", Values, FilterType.String))
                If Not IsNothing(dtActivosModif) AndAlso dtActivosModif.Rows.Count > 0 Then
                    For Each dr As DataRow In dtActivosModif.Select
                        data.DefaultView.RowFilter = IDActivo & "='" & dr("IDActivo") & "'"
                        dr("IDEstadoActivo") = data.DefaultView(0).Row("IDEstadoActivo")
                    Next
                    ac.Update(dtActivosModif)
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CerrarOrdenesServicio(ByVal doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        ''Cierre de O.S. con fianza introducida al hacer un albarán de retorno o de consumo
        Dim AppParamsAV As ParametroAlbaranVenta = services.GetService(Of ParametroAlbaranVenta)()

        Dim IDTipoAlbaran As String = doc.HeaderRow("IDTipoAlbaran")
        If IDTipoAlbaran = AppParamsAV.TipoAlbaranRetornoAlquiler Or IDTipoAlbaran = AppParamsAV.TipoAlbaranDeConsumo Or IDTipoAlbaran = AppParamsAV.TipoAlbaranDeDeposito Then
            Dim IDTrabajo(-1) As Object
            For Each drLinea As DataRow In doc.dtLineas.Rows
                If Length(drLinea("IDTrabajo")) > 0 Then
                    ReDim Preserve IDTrabajo(UBound(IDTrabajo) + 1)
                    IDTrabajo(UBound(IDTrabajo)) = drLinea("IDTrabajo")
                End If
            Next
            If IDTrabajo.Length > 0 Then
                Dim OT As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraTrabajo")
                Dim dtOT As DataTable = OT.Filter(New InListFilterItem("IDTrabajo", IDTrabajo, FilterType.Numeric))
                If Not dtOT Is Nothing AndAlso dtOT.Rows.Count > 0 Then
                    For Each drTrabajo As DataRow In dtOT.Rows
                        Dim f As New Filter
                        f.Add(New NumberFilterItem("IDTrabajo", drTrabajo("IDTrabajo")))
                        f.Add(New StringFilterItem("IDTipoAlbaran", AppParamsAV.TipoAlbaranDeDeposito()))
                        f.Add(New NumberFilterItem("QPendiente", FilterOperator.GreaterThanOrEqual, 1))
                        f.Add(New NumberFilterItem("TipoFactAlquiler", FilterOperator.GreaterThanOrEqual, enumTipoFacturacionAlquiler.enumTFASinAlquiler))

                        Dim dtRet As DataTable = New BE.DataEngine().Filter("vAlquilerCIRetornos", f, "", "IDObra DESC, FechaAlbaran DESC")
                        If Not IsNothing(dtRet) AndAlso dtRet.Rows.Count = 0 Then
                            drTrabajo("Estado") = enumotEstado.otTerminado
                            drTrabajo("FechaFin") = doc.HeaderRow("FechaAlbaran")
                        End If
                    Next

                    BusinessHelper.UpdateTable(dtOT)
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarOrdenesServicioPorBorrado(ByVal drLineaAlbaran As DataRow, ByVal services As ServiceProvider)
        If Length(drLineaAlbaran("IDTrabajo")) > 0 Then
            Dim Parametros As ParametroGeneral = services.GetService(Of ParametroGeneral)()
            If Parametros.AplicacionGestionAlquiler Then
                Dim OT As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraTrabajo")
                Dim dtTrabajo As DataTable = OT.SelOnPrimaryKey(drLineaAlbaran("IDTrabajo"))
                If dtTrabajo.Rows.Count > 0 Then
                    If dtTrabajo.Rows(0)("Estado") = enumotEstado.otTerminado Then
                        dtTrabajo.Rows(0)("Estado") = enumotEstado.otComenzado
                        dtTrabajo.Rows(0)("FechaFin") = DBNull.Value
                        BusinessHelper.UpdateTable(dtTrabajo)
                    End If
                End If
            End If
        End If
    End Sub

#Region " Generación Conductores Alquiler "

    <Serializable()> _
    Public Class dataAddConductores
        Public Conductores As DataTable
        Public LineasAlbaran() As Object    'Array de Integer
        Public IDAlbaran As Integer
        Public Retornos As Boolean

        Public Sub New(ByVal Conductores As DataTable, ByVal IDAlbaran As Integer, ByVal Retornos As Boolean)
            Me.Conductores = Conductores
            Me.IDAlbaran = IDAlbaran
            Me.Retornos = Retornos
        End Sub
    End Class
    <Task()> Public Shared Sub AddConductores(ByVal doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        Dim data As dataPrcActualizarAlbaranAlquiler = services.GetService(Of dataPrcActualizarAlbaranAlquiler)()

        Dim dataConductores As New dataAddConductores(data.Conductores, doc.HeaderRow("IDAlbaran"), data.Retornos)
        ProcessServer.ExecuteTask(Of dataAddConductores)(AddressOf ActualizarAlbaranEnConductores, dataConductores, services)
        ProcessServer.ExecuteTask(Of dataAddConductores)(AddressOf GenerarLineaConductorAlquiler, dataConductores, services)
    End Sub

    <Task()> Public Shared Sub ActualizarAlbaranEnConductores(ByVal data As dataAddConductores, ByVal services As ServiceProvider)
        Dim IDLineasAlbaran(-1) As Object
        If data.IDAlbaran > 0 Then
            Dim dtAVL As DataTable = New DataEngine().Filter("vAlquilerExpedicionArticulosConConductor", New NumberFilterItem("IDAlbaran", data.IDAlbaran))
            For Each drAVL As DataRow In dtAVL.Rows
                ReDim Preserve IDLineasAlbaran(IDLineasAlbaran.Length)
                IDLineasAlbaran(IDLineasAlbaran.Length - 1) = drAVL("IDLineaAlbaran")

                Dim dvConductores As New DataView(data.Conductores)
                dvConductores.RowFilter = String.Empty
                If Not data.Retornos Then
                    If Length(drAVL("IDLineaMaterial")) > 0 Then
                        dvConductores.RowFilter = "IDLineaMaterial= " & drAVL("IDLineaMaterial")
                    End If
                ElseIf Length(drAVL("IDLineaAlbaranDeposito")) > 0 Then
                    dvConductores.RowFilter = "IDLineaAlbaranOrigen= " & drAVL("IDLineaAlbaranDeposito")
                End If

                If dvConductores.RowFilter <> String.Empty AndAlso dvConductores.Count > 0 Then
                    For Each drv As DataRowView In dvConductores
                        drv("IDAlbaran") = drAVL("IDAlbaran")
                        drv("IDLineaAlbaran") = drAVL("IDLineaAlbaran")
                    Next
                End If

                dvConductores.RowFilter = String.Empty
            Next
        End If

        data.LineasAlbaran = IDLineasAlbaran
    End Sub

    <Task()> Public Shared Sub GenerarLineaConductorAlquiler(ByVal data As dataAddConductores, ByVal services As ServiceProvider)
        If Not data.LineasAlbaran Is Nothing AndAlso data.LineasAlbaran.Length > 0 AndAlso Not data.Conductores Is Nothing Then
            Dim AVL As New AlbaranVentaLinea
            Dim dtAVL As DataTable = AVL.Filter(New InListFilterItem("IDLineaAlbaran", data.LineasAlbaran, FilterType.Numeric))
            If Not dtAVL Is Nothing AndAlso dtAVL.Rows.Count > 0 Then
                Dim dtAVLNew As DataTable = AVL.AddNew

                For Each drAVL As DataRow In dtAVL.Rows
                    Dim drAVLNew As DataRow = dtAVLNew.NewRow
                    For Each dc As DataColumn In dtAVL.Columns
                        If dc.ColumnName <> "IDLineaMaterial" And dc.ColumnName <> "Lote" And dc.ColumnName <> "IDMovimiento" And dc.ColumnName <> "IDMovimientoEntrada" And dc.ColumnName <> "Ubicacion" Then
                            drAVLNew(dc.ColumnName) = drAVL(dc.ColumnName)
                        End If
                    Next

                    Dim drConductor() As DataRow = data.Conductores.Select("IDLineaAlbaran=" & drAVL("IDLineaAlbaran"))
                    If drConductor.Length > 0 Then
                        drAVLNew("IDLineaAlbaran") = AdminData.GetAutoNumeric
                        drAVLNew("IDALbaran") = drAVL("IDAlbaran")
                        drAVLNew("IDArticulo") = drConductor(0)("Conductor")

                        If Length(drAVL("Lote")) > 0 Then
                            Dim f As New Filter
                            f.Add(New StringFilterItem("IDArticulo", drConductor(0)("Conductor")))
                            Dim dtGestionStock As DataTable = New DataEngine().Filter("vNegCaractArticulo", f, "GestionStock")
                            If Not dtGestionStock Is Nothing AndAlso dtGestionStock.Rows.Count > 0 Then
                                If Not dtGestionStock.Rows(0)("GestionStock") Then
                                    drAVLNew("Lote") = drAVL("Lote")
                                    drAVLNew("EstadoStock") = enumavlEstadoStock.avlSinGestion
                                End If
                            End If
                        End If

                        drAVLNew("DescArticulo") = drConductor(0)("DescConductor")
                        drAVLNew("QServida") = drConductor(0)("Cantidad")
                        drAVLNew("Factor") = 1
                        drAVLNew("QInterna") = drConductor(0)("Cantidad")
                        Dim AppParamsConta As ParametroContabilidadVenta = services.GetService(Of ParametroContabilidadVenta)()
                        If AppParamsConta.Contabilidad Then drAVLNew("CContable") = drConductor(0)("CContable")

                        'Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                        'Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(drAVLNew("IDArticulo"))

                        drAVLNew("TipoFactAlquiler") = 0 ' ArtInfo.TipoFactAlquiler
                        If Length(drConductor(0)("IDUDMedida")) > 0 Then drAVLNew("IDUDMedida") = drConductor(0)("IDUDMedida")
                        If Length(drConductor(0)("IDUDInterna")) > 0 Then drAVLNew("IDUDInterna") = drConductor(0)("IDUDInterna")

                        Dim dataTarifa As New CalculoTarifaAlquiler.DataCalculoTarifaAlquiler(drAVLNew("IDObra"), drAVLNew("IDArticulo"), _
                                                                                              drConductor(0)("IDCliente"), drAVLNew("QServida"), Date.Today)
                        ProcessServer.ExecuteTask(Of CalculoTarifaAlquiler.DataCalculoTarifaAlquiler)(AddressOf CalculoTarifaAlquiler.TarifaAlquiler, dataTarifa, services)

                        drAVLNew("Precio") = dataTarifa.DatosTarifa.Precio
                        If Length(dataTarifa.DatosTarifa.IDMoneda) > 0 AndAlso drConductor(0)("IDMoneda") & String.Empty <> dataTarifa.DatosTarifa.IDMoneda Then
                            Dim datos As New DataCambioMoneda(New DataRowPropertyAccessor(drAVLNew), drConductor(0)("IDMoneda"), dataTarifa.DatosTarifa.IDMoneda, Date.Today)
                            ProcessServer.ExecuteTask(Of DataCambioMoneda)(AddressOf NegocioGeneral.CambioMoneda, datos, services)
                        End If
                        drAVLNew("Dto1") = dataTarifa.DatosTarifa.Dto1
                        drAVLNew("Dto2") = dataTarifa.DatosTarifa.Dto2
                        drAVLNew("Dto3") = dataTarifa.DatosTarifa.Dto3
                        drAVLNew("UDValoracion") = dataTarifa.DatosTarifa.UDValoracion
                        drAVLNew("SeguimientoTarifa") = dataTarifa.DatosTarifa.SeguimientoTarifa

                        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf NegocioGeneral.CalcularPrecioImporte, New DataRowPropertyAccessor(drAVLNew), services)
                        If drAVLNew("Precio") <> 0 Then
                            Dim ValAyB As New ValoresAyB(New DataRowPropertyAccessor(drAVLNew), drConductor(0)("IDMoneda"), drConductor(0)("CambioA"), drConductor(0)("CambioB"))
                            ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)
                        End If

                        dtAVLNew.Rows.Add(drAVLNew)
                    End If
                Next

                BusinessHelper.UpdateTable(dtAVLNew)
            End If
        End If
    End Sub

#End Region

#Region " Generación Contadores Alquiler "

    <Serializable()> _
    Public Class DataAddContadores
        Public Contadores As DataTable
        Public IDAlbaran As Integer
        Public Retornos As Boolean
        Public SalidaRetornos As Boolean

        Public Sub New(ByVal Contadores As DataTable, ByVal IDAlbaran As Integer, ByVal Retornos As Boolean, ByVal SalidaRetornos As Boolean)
            Me.Contadores = Contadores
            Me.IDAlbaran = IDAlbaran
            Me.Retornos = Retornos
            Me.SalidaRetornos = SalidaRetornos
        End Sub
    End Class
    <Task()> Public Shared Sub AddContadores(ByVal doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        Dim data As dataPrcActualizarAlbaranAlquiler = services.GetService(Of dataPrcActualizarAlbaranAlquiler)()
        If Not data.Contadores Is Nothing Then
            Dim dataContadores As New DataAddContadores(data.Contadores, doc.HeaderRow("IDAlbaran"), data.Retornos, data.SalidaRetornos)
            If dataContadores.Contadores.Columns.Contains("IDLineaMaterial") Then
                Dim OM As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraMaterial")

                Dim WHERE As String = New IsNullFilterItem("IDLineaMaterial").Compose(New AdoFilterComposer)
                For Each drContador As DataRow In dataContadores.Contadores.Select(WHERE)
                    Dim f As New Filter
                    f.Add(New NumberFilterItem("IDTrabajo", drContador("IDTrabajo")))
                    f.Add(New StringFilterItem("Lote", drContador("IDActivo")))
                    Dim dtObraMaterial As DataTable = OM.Filter(f)
                    If Not IsNothing(dtObraMaterial) AndAlso dtObraMaterial.Rows.Count > 0 Then
                        drContador("IDLineaMaterial") = dtObraMaterial.Rows(0)("IDLineaMaterial")
                    End If
                Next
            End If

            Dim IDLineasAlbaran() As Object = ProcessServer.ExecuteTask(Of DataAddContadores, Object())(AddressOf ActualizarAlbaranEnContador, dataContadores, services)
            If Not IDLineasAlbaran Is Nothing AndAlso IDLineasAlbaran.Length > 0 Then
                If dataContadores.SalidaRetornos Then
                    ProcessServer.ExecuteTask(Of DataAddContadores)(AddressOf ActualizarContadoresSalidasRetornos, dataContadores, services)
                Else
                    ProcessServer.ExecuteTask(Of DataTable)(AddressOf ActualizarContadoresDesdeAlquiler, dataContadores.Contadores, services)
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Function ActualizarAlbaranEnContador(ByVal data As DataAddContadores, ByVal services As ServiceProvider) As Object()
        Dim IDLineasAlbaran(-1) As Object
        If data.IDAlbaran > 0 Then
            'Se comprueba si alguno de los albarans generados tiene un artículo cuyo activo está asociado a un contador.
            Dim dtAVL As DataTable = New BE.DataEngine().Filter("vAlquilerCILecturaContador", New NumberFilterItem("IDAlbaran", data.IDAlbaran))
            If Not dtAVL Is Nothing AndAlso dtAVL.Rows.Count > 0 Then
                For Each drAVL As DataRow In dtAVL.Rows
                    ReDim Preserve IDLineasAlbaran(IDLineasAlbaran.Length)
                    IDLineasAlbaran(IDLineasAlbaran.Length - 1) = drAVL("IDLineaAlbaran")

                    Dim dvContadores As New DataView(data.Contadores)
                    dvContadores.RowFilter = String.Empty
                    If Not data.Retornos Then
                        If Length(drAVL("IDLineaMaterial")) > 0 AndAlso Length(drAVL("Lote")) > 0 Then
                            dvContadores.RowFilter = "IDLineaMaterial= " & drAVL("IDLineaMaterial")
                        End If
                    ElseIf Length(drAVL("IDLineaAlbaranDeposito")) > 0 Then
                        If data.SalidaRetornos Then
                            dvContadores.RowFilter = "IDLineaAlbaran= " & drAVL("IDLineaAlbaranDeposito")
                        Else
                            dvContadores.RowFilter = "IDLineaAlbaranOrigen= " & drAVL("IDLineaAlbaranDeposito")
                        End If
                    End If

                    If dvContadores.RowFilter <> String.Empty AndAlso dvContadores.Count > 0 Then
                        For Each drv As DataRowView In dvContadores
                            drv("IDAlbaran") = drAVL("IDAlbaran")
                            If Length(drAVL("IDLineaAlbaranDeposito")) > 0 Then
                                drv("IDLineaAlbaranDeposito") = drAVL("IDLineaAlbaranDeposito")
                            End If

                            drv("IDLineaAlbaran") = drAVL("IDLineaAlbaran")
                        Next
                    End If
                    dvContadores.RowFilter = String.Empty
                Next
            End If
        End If

        Return IDLineasAlbaran
    End Function

    <Task()> Public Shared Sub ActualizarContadoresSalidasRetornos(ByVal data As DataAddContadores, ByVal services As ServiceProvider)
        If Not IsNothing(data.Contadores) AndAlso data.Contadores.Rows.Count > 0 Then
            Dim dtFecha As Date
            Dim dtAVL, dtContHist, dtCont As DataTable
            Dim intIDLineaAlbaran, intIDLineaAlbaranRetorno, intUltimaMedida, intIDOT As Integer
            Dim strWhere, strNROT As String
            Dim strOperario As String = ProcessServer.ExecuteTask(Of Object, String)(AddressOf Operario.ObtenerIDOperarioUsuario, Nothing, services)

            Dim Contador As BusinessHelper
            Contador = BusinessHelper.CreateBusinessObject("PreventivoContador")
            Dim ContHist As BusinessHelper
            ContHist = BusinessHelper.CreateBusinessObject("PreventivoContadorHist")
            Dim OT As BusinessHelper
            OT = BusinessHelper.CreateBusinessObject("MntoOT")
            Dim AVL As New AlbaranVentaLinea

            Dim dtContHistNew As DataTable = ContHist.AddNew
            Dim dtContNew As DataTable = Contador.AddNew

            Dim dvContadores As New DataView(data.Contadores)
            dvContadores.Sort = "IDContadorPrev"
            For Each drv As DataRowView In dvContadores
                dtFecha = Date.Today
                dtAVL = AVL.SelOnPrimaryKey(drv("IDLineaAlbaran"))
                If Not dtAVL Is Nothing AndAlso dtAVL.Rows.Count > 0 Then
                    dtFecha = Nz(dtAVL.Rows(0)("FechaAlquiler"), Date.Today)
                End If

                If data.Retornos Then
                    intIDLineaAlbaranRetorno = drv("IDLineaAlbaran")
                    intIDLineaAlbaran = drv("IDLineaAlbaranDeposito")

                    strWhere = "IDLineaAlbaranRetorno=" & drv("IDLineaAlbaran")
                    Dim dtOT As DataTable = OT.Filter("IDOT,NROT", strWhere)
                    If Not dtOT Is Nothing AndAlso dtOT.Rows.Count > 0 Then
                        intIDOT = dtOT.Rows(0)("IDOT")
                        strNROT = dtOT.Rows(0)("NROT") & String.Empty
                    End If
                Else
                    intIDLineaAlbaran = drv("IDLineaAlbaran")
                End If

                strWhere = "IDContadorPrev ='" & drv("IDContadorPrev") & "'"
                dtContHist = ContHist.Filter(, strWhere, "IDHistoricoContador DESC")
                If Not dtContHist Is Nothing AndAlso dtContHist.Rows.Count > 0 Then
                    intUltimaMedida = dtContHist.Rows(0)("UltimaMedida")
                Else
                    intUltimaMedida = 0
                End If

                'CONTADOR HISTORICO
                Dim drContHistNew As DataRow = dtContHistNew.NewRow

                drContHistNew("IDContadorPrev") = drv("IDContadorPrev")
                If drv("Resetear") Then
                    If data.Retornos Then
                        drContHistNew("UltimaMedida") = Nz(drv("MedidaRetorno"), 0) + intUltimaMedida
                        drContHistNew("MedidaReseteo") = Nz(drv("MedidaRetorno"), 0)
                    Else
                        drContHistNew("UltimaMedida") = Nz(drv("MedidaSalida"), 0)
                        drContHistNew("MedidaReseteo") = Nz(drv("MedidaSalida"), 0) - intUltimaMedida
                    End If
                Else
                    If data.Retornos Then
                        drContHistNew("UltimaMedida") = Nz(drv("MedidaRetorno"), 0)
                        drContHistNew("MedidaReseteo") = Nz(drv("MedidaRetorno"), 0) - intUltimaMedida
                    Else
                        drContHistNew("UltimaMedida") = Nz(drv("MedidaSalida"), 0)
                        drContHistNew("MedidaReseteo") = Nz(drv("MedidaSalida"), 0) - intUltimaMedida
                    End If
                End If

                drContHistNew("IDOperario") = strOperario
                drContHistNew("Fecha") = dtFecha
                If Length(drv("IDLineaAlbaranDeposito")) > 0 Then
                    drContHistNew("IDLineaAlbaranRetorno") = intIDLineaAlbaranRetorno
                    drContHistNew("IDLineaAlbaran") = intIDLineaAlbaran
                Else
                    drContHistNew("IDLineaAlbaran") = intIDLineaAlbaran
                End If
                If intIDOT > 0 Then drContHistNew("IDOT") = intIDOT
                If Length(strNROT) > 0 Then drContHistNew("NROT") = strNROT

                dtContHistNew.Rows.Add(drContHistNew)

                dtCont = Contador.SelOnPrimaryKey(drv("IDContadorPrev"))
                If Not dtCont Is Nothing AndAlso dtCont.Rows.Count > 0 Then
                    ''CONTADOR 
                    Dim drContNew As DataRow = dtContNew.NewRow

                    For Each dc As DataColumn In dtCont.Columns
                        drContNew(dc.ColumnName) = dtCont.Rows(0)(dc.ColumnName)
                    Next
                    If data.Retornos Then
                        drContNew("UltimaMedida") = Nz(drv("MedidaRetorno"), 0)
                    Else
                        drContNew("UltimaMedida") = Nz(drv("MedidaSalida"), 0)
                    End If
                    drContNew("Fecha") = dtFecha

                    dtContNew.Rows.Add(drContNew)
                End If
            Next

            ContHist.Update(dtContHistNew)
            ProcessServer.ExecuteTask(Of DataTable)(AddressOf ActualizarPreventivoContador, dtContNew, services)
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarPreventivoContador(ByVal dtContNew As DataTable, ByVal services As ServiceProvider)
        Dim Contador As BusinessHelper = BusinessHelper.CreateBusinessObject("PreventivoContador")
        If Not IsNothing(dtContNew) AndAlso dtContNew.Rows.Count > 0 Then
            For Each dr As DataRow In dtContNew.Select
                Dim dtPC As DataTable = Contador.SelOnPrimaryKey(dr("IDContadorPrev"))
                If Not IsNothing(dtPC) AndAlso dtPC.Rows.Count > 0 Then
                    dtPC.Rows(0)("UltimaMedida") = dr("UltimaMedida")
                    dtPC.Rows(0)("fecha") = dr("fecha")
                    dtPC.Rows(0)("Pendiente") = dr("Pendiente")
                    Contador.Update(dtPC)
                End If
            Next
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarContadoresDesdeAlquiler(ByVal dtContadores As DataTable, ByVal services As ServiceProvider)
        If Not IsNothing(dtContadores) AndAlso dtContadores.Rows.Count > 0 Then
            Dim dtFecha As Date
            Dim dtCont As DataTable
            Dim intIDLineaAlbaran, intIDLineaAlbaranRetorno, intUltimaMedida, intIDOT As Integer
            Dim strNROT As String
            Dim IDOperario As String = ProcessServer.ExecuteTask(Of Object, String)(AddressOf Operario.ObtenerIDOperarioUsuario, Nothing, services)
            Dim Contador As BusinessHelper = BusinessHelper.CreateBusinessObject("PreventivoContador")
            Dim ContHist As BusinessHelper = BusinessHelper.CreateBusinessObject("PreventivoContadorHist")
            Dim OT As BusinessHelper = BusinessHelper.CreateBusinessObject("MntoOT")
            Dim AVL As New AlbaranVentaLinea

            Dim dtContHistNew As DataTable = ContHist.AddNew
            Dim dtContNew As DataTable = Contador.AddNew

            Dim dvContadores As New DataView(dtContadores)
            dvContadores.Sort = "IDContadorPrev"
            For Each drv As DataRowView In dvContadores
                dtFecha = Date.Today
                Dim dtAVL As DataTable = AVL.SelOnPrimaryKey(drv("IDLineaAlbaran"))
                If Not dtAVL Is Nothing AndAlso dtAVL.Rows.Count > 0 Then
                    dtFecha = Nz(dtAVL.Rows(0)("FechaAlquiler"), Date.Today)
                End If

                If Length(drv("IDLineaAlbaranDeposito")) > 0 Then
                    intIDLineaAlbaranRetorno = drv("IDLineaAlbaran")
                    intIDLineaAlbaran = drv("IdLineaAlbaranDeposito")

                    Dim dtOT As DataTable = OT.Filter(New NumberFilterItem("IDLineaAlbaranRetorno", drv("IDLineaAlbaran")), , "IDOT,NROT")
                    If Not dtOT Is Nothing AndAlso dtOT.Rows.Count > 0 Then
                        intIDOT = dtOT.Rows(0)("IDOT")
                        strNROT = dtOT.Rows(0)("NROT") & String.Empty
                    End If
                Else
                    intIDLineaAlbaran = drv("IDLineaAlbaran")
                End If

                Dim dtContHist As DataTable = ContHist.Filter(New StringFilterItem("IDContadorPrev", drv("IDContadorPrev")), "IDHistoricoContador DESC")
                If Not dtContHist Is Nothing AndAlso dtContHist.Rows.Count > 0 Then
                    intUltimaMedida = dtContHist.Rows(0)("UltimaMedida")
                Else
                    intUltimaMedida = 0
                End If

                ''CONTADOR HISTORICO
                Dim drContHistNew As DataRow = dtContHistNew.NewRow

                drContHistNew("IDContadorPrev") = drv("IDContadorPrev")
                Dim Col As String
                If drv("Resetear") Then
                    drContHistNew("UltimaMedida") = Nz(drv("NuevaMedida"), 0) + intUltimaMedida
                    drContHistNew("MedidaReseteo") = Nz(drv("NuevaMedida"), 0)
                    'drContHistNew("UltimaMedida") = Nz(drv("MedidaRetorno"), 0) + intUltimaMedida
                    'drContHistNew("MedidaReseteo") = Nz(drv("MedidaRetorno"), 0)
                Else
                    drContHistNew("UltimaMedida") = Nz(drv("NuevaMedida"), 0)
                    drContHistNew("MedidaReseteo") = Nz(drv("NuevaMedida"), 0) - intUltimaMedida
                    'drContHistNew("UltimaMedida") = Nz(drv("MedidaRetorno"), 0)
                    'drContHistNew("MedidaReseteo") = Nz(drv("MedidaRetorno"), 0) - intUltimaMedida
                End If

                drContHistNew("IDOperario") = IDOperario
                drContHistNew("Fecha") = dtFecha
                If Length(drv("IDLineaAlbaranDeposito")) > 0 Then
                    drContHistNew("IDLineaAlbaranRetorno") = intIDLineaAlbaranRetorno
                    drContHistNew("IDLineaAlbaran") = intIDLineaAlbaran
                Else
                    drContHistNew("IDLineaAlbaran") = intIDLineaAlbaran
                End If
                If intIDOT > 0 Then drContHistNew("IDOT") = intIDOT
                If Length(strNROT) > 0 Then drContHistNew("NROT") = strNROT

                dtContHistNew.Rows.Add(drContHistNew)

                dtCont = Contador.SelOnPrimaryKey(drv("IDContadorPrev"))
                If Not dtCont Is Nothing AndAlso dtCont.Rows.Count > 0 Then
                    dtCont.Rows(0)("UltimaMedida") = Nz(drv("NuevaMedida"), 0)
                    'dtCont.Rows(0)("UltimaMedida") = Nz(drv("MedidaRetorno"), 0)
                    dtCont.Rows(0)("Fecha") = dtFecha
                End If
                Contador.Update(dtCont)
            Next

            ContHist.Update(dtContHistNew)
        End If
    End Sub

#End Region

    <Task()> Public Shared Sub AddHistoricoAvisos(ByVal data As dataPrcActualizarAlbaranAlquiler, ByVal services As ServiceProvider)
        If Not data.Avisos Is Nothing AndAlso data.Avisos.Rows.Count > 0 Then
            Dim IDLineasAlbaran(-1) As Object
            For Each drAvisos As DataRow In data.Avisos.Rows
                ReDim Preserve IDLineasAlbaran(IDLineasAlbaran.Length)
                IDLineasAlbaran(IDLineasAlbaran.Length - 1) = drAvisos("IDLineaAlbaran")
            Next

            Dim AVL As New AlbaranVentaLinea
            Dim dtAVL As DataTable = AVL.Filter(New InListFilterItem("IDLineaAlbaran", IDLineasAlbaran, FilterType.Numeric))
            If Not IsNothing(dtAVL) AndAlso dtAVL.Rows.Count > 0 Then
                Dim HA As BusinessHelper = BusinessHelper.CreateBusinessObject("HistoricoAvisoRetorno")
                Dim dtHistoricoAvisosNew As DataTable = HA.AddNew

                For Each drAVL As DataRow In dtAVL.Rows
                    If Length(drAVL("ARNalbaranRecogida")) > 0 Or Length(drAVL("ARContactoObra")) > 0 Or Length(drAVL("ARTelefono")) > 0 Or Length(drAVL("ARRecogidoPor")) > 0 Or Length(drAVL("ARMatricula")) > 0 Or drAVL("ARQAvisoRetorno") > 0 Then
                        Dim drHistoricoAvisosNew As DataRow = dtHistoricoAvisosNew.NewRow

                        drHistoricoAvisosNew("ARNAlbaranRecogida") = drAVL("ARNAlbaranRecogida")
                        drHistoricoAvisosNew("ARContactoObra") = drAVL("ARContactoObra")
                        drHistoricoAvisosNew("ARTelefono") = drAVL("ARTelefono")
                        drHistoricoAvisosNew("ARRecogidoPor") = drAVL("ARRecogidoPor")
                        drHistoricoAvisosNew("ARMatricula") = drAVL("ARMatricula")
                        drHistoricoAvisosNew("ARQAvisoRetorno") = drAVL("ARQAvisoRetorno")
                        drHistoricoAvisosNew("IDLineaAlbaran") = drAVL("IDLineaAlbaran")
                        drHistoricoAvisosNew("ARFechaAvisoRecogida") = drAVL("ARFechaAvisoRecogida")
                        drHistoricoAvisosNew("ARIDContador") = drAVL("ARIDContador")
                        drHistoricoAvisosNew("ARFechaPrevistaRetorno") = drAVL("ARFechaPrevistaRetorno")
                        drHistoricoAvisosNew("ARTexto") = drAVL("ARTexto")

                        Dim dtAVLRetorno As DataTable = AVL.Filter(New NumberFilterItem("IDLineaAlbaranDeposito", drAVL("IDLineaAlbaran")))
                        If Not dtAVLRetorno Is Nothing AndAlso dtAVLRetorno.Rows.Count > 0 Then
                            drHistoricoAvisosNew("IDAlbaranRetorno") = dtAVLRetorno.Rows(0)("IDAlbaran")
                            drHistoricoAvisosNew("IdLineaAlbaranRetorno") = dtAVLRetorno.Rows(0)("IDLineaAlbaran")
                            drHistoricoAvisosNew("QAlbaranRetornada") = dtAVLRetorno.Rows(0)("QServida")
                            drHistoricoAvisosNew("FechaRetorno") = dtAVLRetorno.Rows(0)("FechaAlquiler")
                        End If

                        dtHistoricoAvisosNew.Rows.Add(drHistoricoAvisosNew)
                    End If

                    dtAVL.Rows(0)("ARNAlbaranRecogida") = System.DBNull.Value
                    dtAVL.Rows(0)("ARContactoObra") = System.DBNull.Value
                    dtAVL.Rows(0)("ARTelefono") = System.DBNull.Value
                    dtAVL.Rows(0)("ARRecogidoPor") = System.DBNull.Value
                    dtAVL.Rows(0)("ARMatricula") = System.DBNull.Value
                    dtAVL.Rows(0)("ARQAvisoRetorno") = 0
                    dtAVL.Rows(0)("ARFechaAvisoRecogida") = System.DBNull.Value
                    dtAVL.Rows(0)("ARFechaPrevistaRetorno") = System.DBNull.Value
                    dtAVL.Rows(0)("ARIDContador") = System.DBNull.Value
                    dtAVL.Rows(0)("ARTexto") = System.DBNull.Value
                Next

                HA.Update(dtHistoricoAvisosNew)
                AVL.Update(dtAVL)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AddObraMaterial(ByVal data As dataPrcActualizarAlbaranAlquiler, ByVal services As ServiceProvider)
        If Not data.ADDObraMaterial Is Nothing AndAlso data.ADDObraMaterial.Rows.Count > 0 Then
            Dim OM As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraMaterial")
            Dim dtNewOM As DataTable = OM.AddNew

            For Each drObraMaterial As DataRow In data.ADDObraMaterial.Rows
                Dim dtOM As DataTable = OM.SelOnPrimaryKey(drObraMaterial("IDLineaMaterial"))
                If Not IsNothing(dtOM) AndAlso dtOM.Rows.Count > 0 Then
                    Dim drNewOM As DataRow = dtNewOM.NewRow
                    For Each oCol As DataColumn In dtNewOM.Columns
                        If oCol.ColumnName <> "IDMaterial" And oCol.ColumnName <> "DescMaterial" And oCol.ColumnName <> "IDMaterialOrigen" And _
                           oCol.ColumnName <> "Lote" And oCol.ColumnName <> "Estado" And oCol.ColumnName <> "IDLineaMaterial" And _
                           oCol.ColumnName <> "QServida" And oCol.ColumnName <> "QRetornada" Then
                            drNewOM(oCol.ColumnName) = dtOM.Rows(0)(oCol.ColumnName)
                        ElseIf oCol.ColumnName = "IDMaterial" Then
                            drNewOM("IDMaterial") = drObraMaterial("IDMaterial")
                        ElseIf oCol.ColumnName = "DescMaterial" Then
                            drNewOM("DescMaterial") = drObraMaterial("DescArticulo")
                        ElseIf oCol.ColumnName = "Lote" Then
                            drNewOM("Lote") = drObraMaterial("Lote")
                        ElseIf oCol.ColumnName = "Estado" Then
                            drNewOM("Estado") = enumomEstado.omPendiente
                        End If
                    Next

                    drNewOM("IDLineaMaterial") = AdminData.GetAutoNumeric

                    Dim context As New BusinessData
                    context("IDProveedor") = drNewOM("IDProveedor")

                    Dim dtO As DataTable = New BE.DataEngine().Filter("tbObraCabecera", New NumberFilterItem("IDObra", drNewOM("IDObra")), "IDCliente")
                    If Not IsNothing(dtO) AndAlso dtO.Rows.Count > 0 Then
                        context("IDCliente") = dtO.Rows(0)("IDCliente")
                    End If

                    drNewOM = OM.ApplyBusinessRule("IDMaterial", drObraMaterial("IDMaterial"), drNewOM, context)

                    Dim AlmAlq As New Almacen.DataRecuperarAlmacenAlquiler(drObraMaterial("IDMaterial"), drObraMaterial("IDCentroGestion") & String.Empty)
                    Dim IDAlmacen As String = ProcessServer.ExecuteTask(Of Almacen.DataRecuperarAlmacenAlquiler, String)(AddressOf Almacen.RecuperaAlmacenAlquiler, AlmAlq, services)
                    If Len(IDAlmacen) > 0 Then drNewOM("IDAlmacen") = IDAlmacen

                    dtNewOM.Rows.Add(drNewOM)
                End If
            Next

            OM.Update(dtNewOM)
        End If
    End Sub

    <Task()> Public Shared Function Resultado(ByVal data As Object, ByVal services As ServiceProvider) As ResultAlbaranAlquiler
        'Elimina la información almacenada en memoria si previamente hemos cancelado el albarán
        AdminData.GetSessionData("__AlbAlqx__")
        'Guardamos la información del documento en memoria, para recuperarla cuando volvamos del preview de presentación
        AdminData.SetSessionData("__AlbAlqx__", services.GetService(Of ArrayList))
        Return services.GetService(Of ResultAlbaranAlquiler)()
    End Function

End Class
