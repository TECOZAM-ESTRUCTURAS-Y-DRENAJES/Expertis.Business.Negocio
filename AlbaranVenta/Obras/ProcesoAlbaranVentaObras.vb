Public Class ProcesoAlbaranVentaObras

#Region " Agrupación de Obras "

    <Serializable()> _
    Public Class DataColAgrupacionAVObras
        Public Lineas As DataTable
        Public TipoAgrupacion As enummcAgrupAlbaranObra
    End Class

    <Task()> Public Shared Function GetGroupColumnsObras(ByVal data As DataColAgrupacionAVObras, ByVal services As ServiceProvider) As DataColumn()
        Dim columns(data.TipoAgrupacion) As DataColumn
        If data.TipoAgrupacion = enummcAgrupAlbaranObra.mcCliente Then
            columns(0) = data.Lineas.Columns("IDCliente")
        Else
            columns(0) = data.Lineas.Columns("IDCliente")
            columns(1) = data.Lineas.Columns("IDObra")
            If data.TipoAgrupacion = enummcAgrupAlbaranObra.mcObraTrabajo Then
                columns(2) = data.Lineas.Columns("IDTrabajo")
            End If
        End If

        Return columns
    End Function

    <Task()> Public Shared Function AgruparObras(ByVal data As DataPrcAlbaranar, ByVal services As ServiceProvider) As AlbCabVentaObras()
        Dim IDMaterial(-1) As Object
        Dim htLins As New Hashtable
        For Each AVInfo As CrearAlbaranVentaInfo In data.AlbVentaInfo
            If AVInfo.IDLineaMaterial > 0 Then
                ReDim Preserve IDMaterial(IDMaterial.Length)
                IDMaterial(IDMaterial.Length - 1) = AVInfo.IDLineaMaterial
            End If
            htLins.Add(AVInfo.IDLineaMaterial, AVInfo)
        Next

        If IDMaterial.Length > 0 Then
            Dim ProcInfo As ProcessInfoAV = services.GetService(Of ProcessInfoAV)()
            Dim cnViewName As String = "vFrmMntoExpedicionObras"
            Dim strSelect As String = "IDLineaMaterial"
            Dim OrderBy As String = "IDObra DESC"

            Dim f As New Filter
            f.Add(New InListFilterItem(strSelect, IDMaterial, FilterType.Numeric))
            Dim dtLineas As DataTable = AdminData.GetData(cnViewName, f, , OrderBy)
            If Not dtLineas Is Nothing AndAlso dtLineas.Rows.Count > 0 Then
                Dim Agrup As New DataColAgrupacionAVObras
                Agrup.Lineas = dtLineas
                Agrup.TipoAgrupacion = enummcAgrupAlbaranObra.mcCliente
                Dim ColsAgrupCliente As DataColumn() = ProcessServer.ExecuteTask(Of DataColAgrupacionAVObras, DataColumn())(AddressOf GetGroupColumnsObras, Agrup, services)
                Agrup.TipoAgrupacion = enummcAgrupAlbaranObra.mcObra
                Dim ColsAgrupObra As DataColumn() = ProcessServer.ExecuteTask(Of DataColAgrupacionAVObras, DataColumn())(AddressOf GetGroupColumnsObras, Agrup, services)
                Agrup.TipoAgrupacion = enummcAgrupAlbaranObra.mcObraTrabajo
                Dim ColsAgrupObraTrabajo As DataColumn() = ProcessServer.ExecuteTask(Of DataColAgrupacionAVObras, DataColumn())(AddressOf GetGroupColumnsObras, Agrup, services)

                Dim oGrprUser As New GroupUserObras
                Dim groupers(2) As GroupHelper
                groupers(enummcAgrupAlbaranObra.mcCliente) = New GroupHelper(ColsAgrupCliente, oGrprUser)
                groupers(enummcAgrupAlbaranObra.mcObra) = New GroupHelper(ColsAgrupObra, oGrprUser)
                groupers(enummcAgrupAlbaranObra.mcObraTrabajo) = New GroupHelper(ColsAgrupObraTrabajo, oGrprUser)

                For Each rwLin As DataRow In dtLineas.Select(Nothing, "IDCliente,IDObra,IDCentroGestion,IDLineaMaterial")
                    groupers(rwLin("AgrupAlbaranObra")).Group(rwLin)
                Next

                For Each alb As AlbCabVentaObras In oGrprUser.Albs
                    For Each alblin As AlbLinVentaObras In alb.LineasOrigen
                        alblin.QaServir = DirectCast(htLins(alblin.IDLineaOrigen), CrearAlbaranVentaInfo).Cantidad
                        alblin.Lotes = DirectCast(htLins(alblin.IDLineaOrigen), CrearAlbaranVentaInfo).Lotes
                        alblin.Series = DirectCast(htLins(alblin.IDLineaOrigen), CrearAlbaranVentaInfo).Series
                    Next
                Next

                Return oGrprUser.Albs
            End If
        End If
    End Function

#End Region

#Region " Nueva Cabecera "

    <Task()> Public Shared Sub AsignarDatosCliente(ByVal alb As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        If Length(alb.HeaderRow("IDCliente")) = 0 Then Exit Sub
        ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.AsignarDatosCliente, alb, services)

        Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
        If alb.Cliente Is Nothing Then alb.Cliente = Clientes.GetEntity(alb.HeaderRow("IDCliente"))

        If alb.HeaderRow.IsNull("IDCentroGestion") Then alb.HeaderRow("IDCentroGestion") = alb.Cliente.CentroGestion
    End Sub

    <Task()> Public Shared Sub AsignarCentroGestion(ByVal alb As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        If Length(CType(alb.Cabecera, AlbCabVentaObras).IDCentroGestion) > 0 Then
            alb.HeaderRow("IDCentroGestion") = CType(alb.Cabecera, AlbCabVentaObras).IDCentroGestion
        End If
    End Sub

    <Task()> Public Shared Sub AsignarPedidoCliente(ByVal alb As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        If Length(CType(alb.Cabecera, AlbCabVentaObras).PedidoCliente) > 0 Then
            alb.HeaderRow("PedidoCliente") = CType(alb.Cabecera, AlbCabVentaObras).PedidoCliente
        End If
    End Sub

    <Task()> Public Shared Sub AsignarTexto(ByVal doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        If Length(CType(doc.Cabecera, AlbCabVentaObras).PedidoCliente) > 0 Then
            doc.HeaderRow("Texto") = CType(doc.Cabecera, AlbCabVentaObras).Texto
        End If
    End Sub

#End Region

#Region " Crear Lineas "

    <Task()> Public Shared Sub CrearLineasDesdeObras(ByVal oDocAlb As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        Dim dtOrigen As DataTable = ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta, DataTable)(AddressOf RecuperarDatosObras, oDocAlb, services)
        Dim OrdenLinea As Integer
        For Each lineaOrigen As DataRow In dtOrigen.Rows
            Dim alblin As Object

            For i As Integer = 0 To oDocAlb.Cabecera.LineasOrigen.Length - 1
                If lineaOrigen(oDocAlb.Cabecera.LineasOrigen(i).PrimaryKeyLinOrigen) = oDocAlb.Cabecera.LineasOrigen(i).IDLineaOrigen Then
                    alblin = oDocAlb.Cabecera.LineasOrigen(i)
                    Exit For
                ElseIf TypeOf oDocAlb.Cabecera.LineasOrigen(i) Is AlbLinVentaAlquiler AndAlso _
                     lineaOrigen(oDocAlb.Cabecera.LineasOrigen(i).PrimaryKeyLinOrigen) = CType(oDocAlb.Cabecera.LineasOrigen(i), AlbLinVentaAlquiler).IDLineaMaterial Then
                    alblin = CType(oDocAlb.Cabecera.LineasOrigen(i), AlbLinVentaAlquiler)
                    Exit For
                End If
            Next

            If Not alblin Is Nothing Then
                Dim dblCantidad As Double = alblin.QaServir
                If dblCantidad <> 0 Then
                    Dim NumLineasInsertar As Integer = 1
                    If alblin.QaServir > 1 AndAlso Not alblin.Series Is Nothing AndAlso alblin.Series.Rows.Count > 0 Then NumLineasInsertar = alblin.Series.Rows.Count
                    Dim linAlbPed As DataLineasAVDesdeOrigen
                    For i As Integer = NumLineasInsertar - 1 To 0 Step -1
                        Dim linea As DataRow = oDocAlb.dtLineas.NewRow
                        linea("IDAlbaran") = oDocAlb.HeaderRow("IDAlbaran")

                        ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoAlbaranVenta.AsignarValoresPredeterminadosLinea, linea, services)
                        linAlbPed = New DataLineasAVDesdeOrigen(linea, lineaOrigen, oDocAlb, alblin, dblCantidad)
                        If Not alblin.Series Is Nothing AndAlso alblin.Series.Rows.Count > 0 Then
                            linAlbPed = New DataLineasAVDesdeOrigen(linea, lineaOrigen, oDocAlb, alblin, dblCantidad, alblin.Series.Rows(i))
                        Else
                            linAlbPed = New DataLineasAVDesdeOrigen(linea, lineaOrigen, oDocAlb, alblin, dblCantidad)
                        End If

                        ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf AsignarPedidoClienteDeCabecera, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf AsignarDatosArticulo, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf AsignarDatosObra, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf AsignarCuenta, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf AsignarTipoIVA, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf AsignarCentroGestionLinea, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf AsignarCondicionesEconomicas, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf AsignarAlmacen, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf AsignarUnidadesCantidades, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf AsignarQPendienteDevolverAInicio, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf AsignarPreciosDtosImportes, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf AsignarTipoLinea, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf AsignarFacturable, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf AsignarNSerie, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf ProcesoAlbaranVenta.AsignarEstadoStock, linAlbPed, services)

                        If oDocAlb.Cabecera.Origen = enumOrigenAlbaranVenta.Alquiler Then
                            ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf ProcesoAlbaranVentaAlquiler.AsignarDatosAlquiler, linAlbPed, services)
                            ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf ProcesoAlbaranVentaAlquiler.AsignarFechaHoraAlquiler, linAlbPed, services)
                            ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf ProcesoAlbaranVentaAlquiler.AsignarFechaPrevistaRetorno, linAlbPed, services)
                            ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf ProcesoAlbaranVentaAlquiler.AsignarTipoFacturacionAlquiler, linAlbPed, services)
                            ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf ProcesoAlbaranVentaAlquiler.AsignarConsumo, linAlbPed, services)
                            ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf ProcesoAlbaranVentaAlquiler.AsignarCContableConsumo, linAlbPed, services)
                            ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf ProcesoAlbaranVentaAlquiler.AsignarPreciosDtosImportes, linAlbPed, services)
                            ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf ProcesoAlbaranVentaAlquiler.AsignarFacturable, linAlbPed, services)
                            ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf ProcesoAlbaranVentaAlquiler.AsignarDatosNSerie, linAlbPed, services)
                            ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf ProcesoAlbaranVentaAlquiler.AsignarFechaRetornoDiasMinimos, linAlbPed, services)
                        End If
                        OrdenLinea += 1
                        linea("IDOrdenLinea") = OrdenLinea

                        oDocAlb.dtLineas.Rows.Add(linea)
                    Next
                End If
            End If
        Next
    End Sub

    <Task()> Public Shared Sub AsignarNSerie(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        If Length(data.NSerie) > 0 Then data.Row("Lote") = data.NSerie
        If Length(data.IDEstadoActivo) > 0 Then data.Row("IDEstadoActivo") = data.IDEstadoActivo
        If Length(data.IDOperario) > 0 Then data.Row("IDOperario") = data.IDOperario
        If Length(data.Ubicacion) > 0 Then data.Row("Ubicacion") = data.Ubicacion
    End Sub

    <Task()> Public Shared Function RecuperarDatosObras(ByVal oDocAlb As DocumentoAlbaranVenta, ByVal services As ServiceProvider) As DataTable
        Dim albCabPed As AlbCabVenta = oDocAlb.Cabecera

        Dim FieldRow As String
        Dim ids(albCabPed.LineasOrigen.Length - 1) As Object
        For i As Integer = 0 To ids.Length - 1
            If Length(FieldRow) = 0 Then FieldRow = albCabPed.LineasOrigen(i).PrimaryKeyLinOrigen

            Dim ProcInfo As ProcessInfoAV = services.GetService(Of ProcessInfoAV)()
            If ProcInfo.IDTipoAlbaran = enumTipoAlbaran.RetornoAlquiler Or ProcInfo.IDTipoAlbaran = enumTipoAlbaran.Consumo Then
                ids(i) = CType(albCabPed.LineasOrigen(i), AlbLinVentaAlquiler).IDLineaMaterial
            Else

                ids(i) = albCabPed.LineasOrigen(i).IDLineaOrigen
            End If
        Next

        Dim oFltr As New Filter
        oFltr.Add(New InListFilterItem(FieldRow, ids, FilterType.Numeric))
        Return New BE.DataEngine().Filter("vNegDatosExpedicionObraMaterial", oFltr)
    End Function

    <Task()> Public Shared Sub AsignarDatosObra(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        data.Row(data.AlbLin.PrimaryKeyLinOrigen) = data.Origen(data.AlbLin.PrimaryKeyLinOrigen)
        data.Row(data.Doc.Cabecera.PrimaryKeyCabOrigen) = data.Origen(data.Doc.Cabecera.PrimaryKeyCabOrigen)
        data.Row("IDCentroGestion") = CType(data.AlbLin, AlbLinVentaObras).IDCentroGestion
        data.Row("IDTipoIva") = CType(data.AlbLin, AlbLinVentaObras).IDTipoIVA
        data.Row("IDTrabajo") = data.Origen("IDTrabajo")
    End Sub

    <Task()> Public Shared Sub AsignarDatosArticulo(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        Dim AVL As New AlbaranVentaLinea
        Dim context As New BusinessData(data.Doc.HeaderRow)
        AVL.ApplyBusinessRule("IDArticulo", data.Origen("IDMaterial"), data.Row, context)

        data.Row("DescArticulo") = data.Origen("DescMaterial")
    End Sub

    <Task()> Public Shared Sub AsignarTipoIVA(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        If Length(data.Row("IDTipoIVA")) = 0 Then
            Dim datIVA As New ProcesoComercial.DataIvaArticuloCliente(data.Doc.IDCliente, data.Row("IDArticulo"), Nz(data.Doc.HeaderRow("IDDireccion"), 0))
            data.Row("IDTipoIVA") = ProcessServer.ExecuteTask(Of ProcesoComercial.DataIvaArticuloCliente, String)(AddressOf ProcesoComercial.GetIva, datIVA, services)
        End If
    End Sub

    <Task()> Public Shared Sub AsignarCuenta(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        Dim AppParamsConta As ParametroContabilidadVenta = services.GetService(Of ParametroContabilidadVenta)()
        If AppParamsConta.Contabilidad Then data.Row("CContable") = data.Origen("CCVenta")
    End Sub

    <Task()> Public Shared Sub AsignarUnidadesCantidades(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        'data.Row("IDUdMedida") = data.Origen("IDUdMedida")
        data.Row("IDUdInterna") = data.Origen("IDUdInterna")
        data.Row("UdValoracion") = data.Origen("UdValoracion")
        data.Row("QInterna") = data.Cantidad
        Dim datFactor As New ArticuloUnidadAB.DatosFactorConversion(data.Row("IDArticulo"), data.Row("IDUdMedida"), data.Row("IDUdInterna"), True)
        data.Row("Factor") = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, datFactor, services)
        data.Row("QServida") = data.Row("QInterna") / data.Row("Factor")
    End Sub

    <Task()> Public Shared Sub AsignarPreciosDtosImportes(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        Dim Precio As Double : Dim Margen As Double
        If Nz(data.Origen("PrecioVentaA"), 0) = 0 Then
            Margen = IIf(Nz(data.Origen("MargenPrev"), 0) <> 0, 1 + Nz(data.Origen("MargenPrev"), 0) / 100, 1)
            Precio = Nz(data.Origen("PrecioPrevMatA"), 0) * Margen
        Else
            Precio = data.Origen("PrecioVentaA")
        End If
        Precio = (Precio / data.Doc.CambioA) * Nz(data.Row("Factor"), 1)

        data.Row("Dto1") = data.Origen("DtoVenta1")
        data.Row("Dto2") = data.Origen("DtoVenta2")
        data.Row("Dto3") = data.Origen("DtoVenta3")
        data.Row("Dto") = Nz(data.Doc.HeaderRow("DtoAlbaran"), 0)
        data.Row("DtoProntoPago") = Nz(data.Doc.HeaderRow("DtoProntoPago"), 0)

        Dim AVL As New AlbaranVentaLinea
        Dim context As New BusinessData(data.Doc.HeaderRow)
        AVL.ApplyBusinessRule("Precio", Precio, data.Row, context)

        Dim ValAyB As New ValoresAyB(New DataRowPropertyAccessor(data.Row), data.Doc.IDMoneda, data.Doc.CambioA, data.Doc.CambioB)
        ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)
    End Sub

    <Task()> Public Shared Sub AsignarAlmacen(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        Dim ProcInfo As ProcessInfoAV = services.GetService(Of ProcessInfoAV)()
        If ProcInfo.IDTipoAlbaran <> enumTipoAlbaran.RetornoAlquiler AndAlso ProcInfo.IDTipoAlbaran <> enumTipoAlbaran.Consumo Then
            data.Row("IDAlmacen") = data.Origen("IDAlmacen")
        End If
    End Sub

    <Task()> Public Shared Sub AsignarQPendienteDevolverAInicio(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        Dim ProcInfo As ProcessInfoAV = services.GetService(Of ProcessInfoAV)()
        If ProcInfo.IDTipoAlbaran <> enumTipoAlbaran.RetornoAlquiler AndAlso ProcInfo.IDTipoAlbaran <> enumTipoAlbaran.Consumo Then
            data.Row("QPendienteDevolverAInicio") = data.Cantidad
        End If
    End Sub

    '<Task()> public Shared Sub AsignarLote(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
    '    If Length(data.Origen("Lote")) > 0 Then
    '        data.Row("Lote") = data.Origen("Lote")
    '    End If
    'End Sub

    <Task()> Public Shared Sub AsignarTipoLinea(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        data.Row("IDTipoLinea") = ProcessServer.ExecuteTask(Of Object, String)(AddressOf TipoLinea.TipoLineaPorDefecto, Nothing, services)
    End Sub

    <Task()> Public Shared Sub AsignarFacturable(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        If Nz(data.Origen("TipoFacturacion"), enumomTipoFacturacion.omPorVencimientos) = enumomTipoFacturacion.omPorVencimientos Or Nz(data.Origen("TipoFacturacion"), enumomTipoFacturacion.omPorVencimientos) = enumomTipoFacturacion.omPorCantidad Then
            data.Row("Facturable") = False
        Else
            data.Row("Facturable") = data.Origen("Facturable")
        End If
    End Sub

    <Task()> Public Shared Sub AsignarCondicionesEconomicas(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        If Length(data.Doc.HeaderRow("IDCondicionPago")) > 0 Then
            data.Row("IDCondicionPago") = data.Doc.HeaderRow("IDCondicionPago")
        End If
        If Length(data.Doc.HeaderRow("IDFormaPago")) > 0 Then
            data.Row("IDFormaPago") = data.Doc.HeaderRow("IDFormaPago")
        End If
    End Sub

    <Task()> Public Shared Sub AsignarCentroGestionLinea(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        data.Row("IDCentroGestion") = data.Doc.HeaderRow("IDCentroGestion")
    End Sub

    <Task()> Public Shared Sub AsignarPedidoClienteDeCabecera(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        If Length(data.Doc.HeaderRow("PedidoCliente")) > 0 Then
            data.Row("PedidoCliente") = data.Doc.HeaderRow("PedidoCliente")
        End If
    End Sub

#End Region

#Region " Actualizar Obra Material desde Linea de Albarán "

    <Task()> Public Shared Sub ActualizarObraMaterialLineaPorBorrado(ByVal drLineaAlbaran As DataRow, ByVal services As ServiceProvider)
        If Length(drLineaAlbaran("IDLineaMaterial")) > 0 AndAlso drLineaAlbaran("TipoLineaAlbaran") <> enumavlTipoLineaAlbaran.avlComponente Then
            Dim dtCab As DataTable = New AlbaranVentaCabecera().SelOnPrimaryKey(drLineaAlbaran("IDAlbaran"))
            If dtCab.Rows.Count > 0 Then
                Dim datosAct As New DataActualizarObraMaterialLinea(dtCab.Rows(0)("IDTipoAlbaran"), drLineaAlbaran, True)
                ProcessServer.ExecuteTask(Of DataActualizarObraMaterialLinea)(AddressOf ActualizarObraMaterialLinea, datosAct, services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarObrasDesdeAlbaran(ByVal Doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        For Each drLineaAlbaran As DataRow In Doc.dtLineas.Rows
            If drLineaAlbaran.RowState <> DataRowState.Deleted AndAlso Length(drLineaAlbaran("IDLineaMaterial")) > 0 Then
                If drLineaAlbaran.RowState <> DataRowState.Modified OrElse Nz(drLineaAlbaran("QInterna"), 0) <> Nz(drLineaAlbaran("QInterna", DataRowVersion.Original), 0) Then
                    Dim datosAct As New DataActualizarObraMaterialLinea(Doc.HeaderRow("IDTipoAlbaran"), drLineaAlbaran)
                    ProcessServer.ExecuteTask(Of DataActualizarObraMaterialLinea)(AddressOf ActualizarObraMaterialLinea, datosAct, services)
                End If
            End If
        Next
    End Sub

    '<Task()> Public Shared Sub ActualizarObrasDesdeAlbaran(ByVal Doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
    '    If Doc Is Nothing OrElse Doc.dtLineas Is Nothing Then Exit Sub
    '    For Each lineaAlbaran As DataRow In Doc.dtLineas.Select(Nothing, "IDAlbaran,IDLineaAlbaran")
    '        Dim data As New DataActualizarObraMaterialLinea(Doc.HeaderRow("IDTipoAlbaran"), lineaAlbaran)
    '        ProcessServer.ExecuteTask(Of Object)(AddressOf ActualizarObraMaterialLinea, data, services)
    '    Next
    '    ProcessServer.ExecuteTask(Of Object)(AddressOf GrabarPedidos, Nothing, services)
    'End Sub

#Region " ActualizarObraMaterialLinea "

    Public Class DataActualizarObraMaterialLinea
        Public IDTipoAlbaran As String
        Public LineaAlbaran As DataRow
        Public LineaMaterial As DataRow
        Public Delete As Boolean

        Public Sub New(ByVal IDTipoAlbaran As String, ByVal LineaAlbaran As DataRow, Optional ByVal Delete As Boolean = False)
            Me.IDTipoAlbaran = IDTipoAlbaran
            Me.LineaAlbaran = LineaAlbaran
            Me.Delete = Delete
        End Sub
    End Class

    <Task()> Public Shared Sub ActualizarObraMaterialLinea(ByVal data As DataActualizarObraMaterialLinea, ByVal services As ServiceProvider)
        '// Si el albarán de consumo viene derivado de una generación automática de consumos de alquiler, la cantidad servida 
        '// no se tiene que modificarla ya que entonces se desvirtua la información referente a esta cantidad.  
        If Length(data.LineaAlbaran("IDLineaMaterial")) > 0 AndAlso data.LineaAlbaran("TipoLineaAlbaran") <> enumavlTipoLineaAlbaran.avlComponente Then
            Dim OM As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraMaterial")
            Dim dtLineaMaterial As DataTable = OM.SelOnPrimaryKey(data.LineaAlbaran("IDLineaMaterial"))
            If Not dtLineaMaterial Is Nothing AndAlso dtLineaMaterial.Rows.Count > 0 Then
                Dim blnEsConsumo, blnEsAlquiler, blnEsDeposito As Boolean
                If Length(data.IDTipoAlbaran) > 0 Then
                    Dim AppParamsAV As ParametroAlbaranVenta = services.GetService(Of ParametroAlbaranVenta)()
                    blnEsConsumo = (data.IDTipoAlbaran = AppParamsAV.TipoAlbaranDeConsumo())
                    blnEsDeposito = (data.IDTipoAlbaran = AppParamsAV.TipoAlbaranDeDeposito())
                    blnEsAlquiler = (data.IDTipoAlbaran = AppParamsAV.TipoAlbaranRetornoAlquiler())
                End If

                ''''//Si es una línea que viene generada de los consumos de alquiler, no ha actualizado para nada
                ''''//las líneas de obra material, por lo tanto en el borrado no tiene que hacer nada en esta tabla.
                data.LineaMaterial = dtLineaMaterial.Rows(0)
                If blnEsConsumo OrElse blnEsAlquiler Then
                    ProcessServer.ExecuteTask(Of DataActualizarObraMaterialLinea)(AddressOf ActualizarQRetornadaMaterial, data, services)
                    ProcessServer.ExecuteTask(Of DataActualizarObraMaterialLinea)(AddressOf ActualizarQRetornadaLineaAlbaranDeposito, data, services)
                Else
                    If data.LineaMaterial("Deposito") = 0 And blnEsDeposito Then
                        data.LineaMaterial("Deposito") = True
                    End If
                    If data.Delete Then data.LineaAlbaran("QInterna") = 0
                    ProcessServer.ExecuteTask(Of DataActualizarObraMaterialLinea)(AddressOf ActualizarQServidaMaterial, data, services)
                End If
                BusinessHelper.UpdateTable(dtLineaMaterial)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarQRetornadaLineaAlbaranDeposito(ByVal data As DataActualizarObraMaterialLinea, ByVal services As ServiceProvider)
        '//La segunda ajusta en un albarán de retorno/consumo las cantidades retornadas del albarán de deposito
        If data.LineaAlbaran("IDLineaAlbaranDeposito") <> 0 Then
            Dim dtAct As DataTable = New AlbaranVentaLinea().SelOnPrimaryKey(data.LineaAlbaran("IDLineaAlbaranDeposito"))
            If Not dtAct Is Nothing AndAlso dtAct.Rows.Count > 0 Then
                Dim f As New Filter
                f.Add(New NumberFilterItem("IDLineaAlbaranDeposito", data.LineaAlbaran("IDLineaAlbaranDeposito")))
                Dim dt As DataTable = New BE.DataEngine().Filter("tbAlbaranVentaLinea", f, "SUM(QInterna) AS Devuelto")
                If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                    If data.Delete Then
                        dtAct.Rows(0)("QRetornada") = Nz(dt.Rows(0)("Devuelto"), 0) - data.LineaAlbaran("QInterna") 'en la qfacturada se guarda la Qdevuelta
                    Else
                        dtAct.Rows(0)("QRetornada") = Nz(dt.Rows(0)("Devuelto"), 0) + data.LineaAlbaran("QInterna") 'en la qfacturada se guarda la Qdevuelta
                    End If
                    BusinessHelper.UpdateTable(dtAct)
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarQRetornadaMaterial(ByVal data As DataActualizarObraMaterialLinea, ByVal services As ServiceProvider)
        If data.LineaMaterial Is Nothing Then ApplicationService.GenerateError("Debe indicar la línea de ObraMaterial.")
        Dim f As New Filter
        f.Add(New IsNullFilterItem("IDLineaAlbaranDeposito", False))
        f.Add(New NumberFilterItem("IDLineaMaterial", data.LineaAlbaran("IDLineaMaterial")))

        Dim dt As DataTable = New BE.DataEngine().Filter("tbAlbaranVentaLinea", f, "SUM(QInterna) AS Devuelto")
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            If data.Delete Then
                data.LineaMaterial("QRetornada") = Nz(dt.Rows(0)("Devuelto"), 0) - data.LineaAlbaran("QInterna") 'en la qfacturada se guarda la Qdevuelta
            Else
                data.LineaMaterial("QRetornada") = Nz(dt.Rows(0)("Devuelto"), 0) 'en la qfacturada se guarda la Qdevuelta
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarQServidaMaterial(ByVal data As DataActualizarObraMaterialLinea, ByVal services As ServiceProvider)
        If data.LineaMaterial Is Nothing Then ApplicationService.GenerateError("Debe indicar la línea de ObraMaterial.")

        Dim OriginalQServida As Double
        Dim ProposedQServida As Double = Nz(data.LineaAlbaran("QInterna"), 0)
        If data.LineaAlbaran.RowState = DataRowState.Modified Then
            OriginalQServida = data.LineaAlbaran("QInterna", DataRowVersion.Original)
        End If

        Dim DiferenciaQServida As Double = ProposedQServida - OriginalQServida
        data.LineaMaterial("QServida") = data.LineaMaterial("QServida") + DiferenciaQServida
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf ActualizarEstadoMaterial, data.LineaMaterial, services)
    End Sub

    <Task()> Public Shared Sub ActualizarEstadoMaterial(ByVal drMaterial As DataRow, ByVal services As ServiceProvider)
        If drMaterial("QServida") >= drMaterial("QPrev") Then
            drMaterial("Estado") = enumomEstado.omServido
        ElseIf drMaterial("QServida") < drMaterial("QPrev") And drMaterial("QServida") > 0 Then
            drMaterial("Estado") = enumomEstado.omParcServido
        Else
            drMaterial("Estado") = enumomEstado.omPendiente
        End If
    End Sub

#End Region

#End Region

End Class
