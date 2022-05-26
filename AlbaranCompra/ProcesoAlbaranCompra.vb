Public Class ProcesoAlbaranCompra

    Public Shared _ACC As _AlbaranCompraCabecera
    Public Shared _ACL As _AlbaranCompraLinea
    Public Shared _PCL As _PedidoCompraLinea
    Public Shared _PCC As _PedidoCompraCabecera
    Public Shared _AAL As _ArticuloAlmacenLote
    Public Shared _ACLT As _AlbaranCompraLote

#Region " Validaciones "

    <Task()> Public Shared Sub ValidacionesContador(ByVal data As DataPrcAlbaranarPedCompra, ByVal services As ServiceProvider)
        'ProcessServer.ExecuteTask(Of String)(AddressOf ProcesoComunes.ValidarContadorObligatorio, data.IDContador & String.Empty, services)

        Dim datosValCont As New ProcesoComunes.DataValidarContadorEntidad(data.IDContador, GetType(AlbaranCompraCabecera).Name)
        ProcessServer.ExecuteTask(Of ProcesoComunes.DataValidarContadorEntidad)(AddressOf ProcesoComunes.ValidarContadorEntidad, datosValCont, services)
    End Sub


    <Task()> Public Shared Sub ValidarDocumento(ByVal Doc As DocumentoAlbaranCompra, ByVal services As ServiceProvider)
        Dim ACC As New AlbaranCompraCabecera
        ACC.Validate(Doc.HeaderRow.Table)

        '//Esto se valida fuera del Validate de las líneas por que necesita información de la cabecera
        ProcessServer.ExecuteTask(Of DocumentoAlbaranCompra)(AddressOf ValidarEstadoLineas, Doc, services)

        Dim ACL As New AlbaranCompraLinea
        ACL.Validate(Doc.dtLineas)
    End Sub


    <Task()> Public Shared Sub ValidarAlbaranFacturado(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data("Estado") = enumavcEstadoFactura.avcFacturado Then
            If data.RowState = DataRowState.Modified Then
                If Nz(data("Conductor")) <> Nz(data("Conductor", DataRowVersion.Original)) OrElse _
                    Nz(data("EmpresaTransp")) <> Nz(data("EmpresaTransp", DataRowVersion.Original)) OrElse _
                    Nz(data("CifTransportista")) <> Nz(data("CifTransportista", DataRowVersion.Original)) OrElse _
                    Nz(data("DNIConductor")) <> Nz(data("DNIConductor", DataRowVersion.Original)) OrElse _
                    Nz(data("IDFormaEnvio")) <> Nz(data("IDFormaEnvio", DataRowVersion.Original)) OrElse _
                    Nz(data("Matricula")) <> Nz(data("Matricula", DataRowVersion.Original)) OrElse _
                    Nz(data("Remolque")) <> Nz(data("Remolque", DataRowVersion.Original)) OrElse _
                    Nz(data("PesoBrutoManual")) <> Nz(data("PesoBrutoManual", DataRowVersion.Original)) OrElse _
                    Nz(data("PesoNetoManual")) <> Nz(data("PesoNetoManual", DataRowVersion.Original)) OrElse _
                    Nz(data("NBultos")) <> Nz(data("NBultos", DataRowVersion.Original)) OrElse _
                    Nz(data("IDCondicionEnvio")) <> Nz(data("IDCondicionEnvio", DataRowVersion.Original)) OrElse _
                    Nz(data("IDModoTransporte")) <> Nz(data("IDModoTransporte", DataRowVersion.Original)) Then
                Else
                    ApplicationService.GenerateError("El Albarán está Facturado.")
                End If
            Else
                ApplicationService.GenerateError("El Albarán está Facturado.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarCondicionesEconomicas(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Not Nz(data("Automatico"), False) Then
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf NegocioGeneral.ValidarFormaPagoObligatoria, data, services)
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf NegocioGeneral.ValidarCondicionPagoObligatoria, data, services)
        End If
    End Sub

    <Task()> Public Shared Sub ValidacionesContabilidad(ByVal data As DataRow, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf NegocioGeneral.ValidarEjercicioContableAlbaran, data, services)
    End Sub

    <Task()> Public Shared Sub ValidarSuNumAlbaranFecha(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim BlnChanged As Boolean = False
        Dim FilFind As New Filter
        If data.RowState = DataRowState.Added Then
            If Length(data("SuAlbaran")) > 0 Then
                BlnChanged = True
            End If
        ElseIf data.RowState = DataRowState.Modified Then
            If Nz(data("SuAlbaran"), String.Empty) <> Nz(data("SuAlbaran", DataRowVersion.Original), String.Empty) AndAlso Length(data("SuAlbaran")) > 0 Then
                FilFind.Add("IDAlbaran", FilterOperator.NotEqual, data("IDAlbaran"))
                BlnChanged = True
            End If
        End If
        If BlnChanged Then
            Dim strEjercicio As String = ProcessServer.ExecuteTask(Of Date, String)(AddressOf NegocioGeneral.EjercicioPredeterminado, data("FechaAlbaran"), services)
            FilFind.Add("SuAlbaran", FilterOperator.Equal, data("SuAlbaran"))
            FilFind.Add("IDEjercicio", FilterOperator.Equal, strEjercicio)
            FilFind.Add("IDProveedor", FilterOperator.Equal, data("IDProveedor"))
            Dim DtAlbCompCab As DataTable = New AlbaranCompraCabecera().Filter(FilFind)
            If Not DtAlbCompCab Is Nothing AndAlso DtAlbCompCab.Rows.Count > 0 Then
                ApplicationService.GenerateError("El número de su albarán | ya existe en la base de datos para el Ejercicio |.", data("SuAlbaran"), data("IDEjercicio"))
            End If
        End If
    End Sub

    '<Task()> Public Shared Sub ValidarNumeroAlbaran(ByVal data As DataRow, ByVal services As ServiceProvider)
    '    If data.RowState = DataRowState.Added Then
    '        Dim f As New Filter
    '        f.Add(New StringFilterItem("NAlbaran", data("NAlbaran")))
    '        If Length(data("IDContador")) > 0 Then
    '            f.Add(New StringFilterItem("IDContador", data("IDContador")))
    '        Else
    '            f.Add(New IsNullFilterItem("IDContador", True))
    '        End If

    '        Dim AppParamsConta As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
    '        If AppParamsConta.Contabilidad Then f.Add(New StringFilterItem("IDEjercicio", data("IDEjercicio")))
    '        Dim dtACC As DataTable = New AlbaranCompraCabecera().Filter(f)
    '        If Not dtACC Is Nothing AndAlso dtACC.Rows.Count > 0 Then
    '            If AppParamsConta.Contabilidad Then
    '                ApplicationService.GenerateError("El Albarán {0} ya existe para el Ejercicio {1}.", Quoted(data("NAlbaran")), Quoted(data("IDEjercicio")))
    '            Else
    '                ApplicationService.GenerateError("El Albarán {0} ya existe.", Quoted(data("NAlbaran")))
    '            End If
    '        End If
    '    End If
    'End Sub

    <Task()> Public Shared Sub ValidarEstadoLineas(ByVal Doc As DocumentoAlbaranCompra, ByVal services As ServiceProvider)
        'If Not Nz(Doc.HeaderRow(_ACC.Automatico), False) Then
        For Each dr As DataRow In Doc.dtLineas.Rows
            If dr.RowState = DataRowState.Modified Then
                If Not (dr(_ACL.EstadoFactura, DataRowVersion.Original) = enumaclEstadoFactura.aclNoFacturado) Then
                    If dr(_ACL.IDArticulo) <> dr(_ACL.IDArticulo, DataRowVersion.Original) OrElse dr(_ACL.QServida) <> dr(_ACL.QServida, DataRowVersion.Original) Then
                        If dr(_ACL.EstadoFactura, DataRowVersion.Original) = enumaclEstadoFactura.aclParcFacturado Then ApplicationService.GenerateError("No se puede modificar la línea de Albarán. Está parcialmente facturada.")
                        If dr(_ACL.EstadoFactura, DataRowVersion.Original) = enumaclEstadoFactura.aclFacturado Then ApplicationService.GenerateError("No se puede modificar la línea de Albarán. Está facturada.")
                    End If
                End If
            End If
        Next
        'End If
    End Sub

#End Region

#Region " Proceso Albaran "

#Region " Proceso Creación Albaranes desde Pedidos (PrcAlbaranarPedCompra) "
    <Serializable()> _
    Public Class DataInfoProceso
        Public IDContador As String
        Public FechaAlbaran As Date?
        Public IDTipoCompra As String

        Public Sub New(ByVal IDContador As String, ByVal FechaAlbaran As Date, ByVal IDTipoCompra As String)
            Me.IDContador = IDContador
            Me.FechaAlbaran = FechaAlbaran
            Me.IDTipoCompra = IDTipoCompra
        End Sub
    End Class

    '//Prepara información de entrada al proceso
    <Task()> Public Shared Sub DatosIniciales(ByVal data As DataPrcAlbaranarPedCompra, ByVal services As ServiceProvider)
        If data.Pedidos IsNot Nothing Then
            '//Prepara en el service lo que va a retornar el proceso, para tenerlo disponible en todo el proceso.
            services.RegisterService(New AlbaranLogProcess)

            '//Preparar información necesaria a lo largo del proceso
            Dim datos As New DataInfoProceso(data.IDContador, data.FechaAlbaran, data.IDTipoCompra)
            ProcessServer.ExecuteTask(Of DataInfoProceso)(AddressOf PrepararInformacionProceso, datos, services)
        End If
    End Sub

    '//Preparar información necesaria a lo largo del proceso
    <Task()> Public Shared Sub PrepararInformacionProceso(ByVal data As DataInfoProceso, ByVal services As ServiceProvider)
        If Length(data.IDTipoCompra) = 0 Then
            Dim ParamsAC As ParametroAlbaranCompra = services.GetService(Of ParametroAlbaranCompra)()
            data.IDTipoCompra = ParamsAC.TipoCompraNormal
        End If

        If data.FechaAlbaran Is Nothing OrElse data.FechaAlbaran = cnMinDate Then data.FechaAlbaran = Today

        services.RegisterService(New ProcessInfoAlbCompra(data.IDContador, data.IDTipoCompra, data.FechaAlbaran))
    End Sub

    '//Validaciones antes de empezar a Crear Documentos
    <Task()> Public Shared Sub ValidacionesPrevias(ByVal oResltAgrp() As AlbCabPedidoCompra, ByVal services As ServiceProvider)
        If oResltAgrp Is Nothing OrElse oResltAgrp.Length = 0 Then ApplicationService.GenerateError("No hay datos para crear el Albarán.")
    End Sub

#End Region

#Region " Agrupaciones "
    <Serializable()> _
    Public Class DataColAgrupacionAC
        Public Lineas As DataTable
        Public TipoAgrupacion As enummpAgrupAlbaran
    End Class

    <Task()> Public Shared Function AgruparPedidos(ByVal data As DataPrcAlbaranarPedCompra, ByVal services As ServiceProvider) As AlbCabPedidoCompra()
        If data.Pedidos Is Nothing OrElse data.Pedidos.Length = 0 Then ApplicationService.GenerateError("No hay líneas a procesar. Revise el estado de los Pedidos.")
        Dim IDPedidos() As CrearAlbaranCompraInfo = data.Pedidos

        Dim htLins As New Hashtable
        Dim ids(IDPedidos.Length - 1) As Object
        For i As Integer = 0 To IDPedidos.Length - 1
            ids(i) = IDPedidos(i).IDLinea
            htLins.Add(IDPedidos(i).IDLinea, IDPedidos(i))
        Next

        Const strViewName As String = "vNegCompraCrearAlbaran"

        Dim dtLineas As DataTable
        If ids.Length > 0 Then
            Dim oFltr As New Filter
            oFltr.Add(New InListFilterItem(_PCL.IDLineaPedido, ids, FilterType.Numeric))
            oFltr.Add(New NumberFilterItem(_PCL.Estado, FilterOperator.NotEqual, enumpclEstado.pclCerrado))
            'oFltr.Add(New NumberFilterItem(_PCL.Estado, FilterOperator.NotEqual, enumpclEstado.pclservido))
            dtLineas = New BE.DataEngine().Filter(strViewName, oFltr, "", "IDProveedor, IDMoneda, IDLineaPedido")
        End If

        If Not dtLineas Is Nothing AndAlso dtLineas.Rows.Count > 0 Then
            Dim Agrup As New DataColAgrupacionAC
            Agrup.Lineas = dtLineas
            Agrup.TipoAgrupacion = enummpAgrupAlbaran.mpPedido
            Dim ColsAgrupPedido As DataColumn() = ProcessServer.ExecuteTask(Of DataColAgrupacionAC, DataColumn())(AddressOf GetGroupColumns, Agrup, services)
            Agrup.TipoAgrupacion = enummpAgrupAlbaran.mpProveedor
            Dim ColsAgrupProv As DataColumn() = ProcessServer.ExecuteTask(Of DataColAgrupacionAC, DataColumn())(AddressOf GetGroupColumns, Agrup, services)

            Dim oGrprUser As New GroupUserPedidosCompra()
            Dim groupers(1) As GroupHelper
            groupers(enummpAgrupAlbaran.mpPedido) = New GroupHelper(ColsAgrupPedido, oGrprUser)
            groupers(enummpAgrupAlbaran.mpProveedor) = New GroupHelper(ColsAgrupProv, oGrprUser)

            For Each rwLin As DataRow In dtLineas.Rows
                '//Sólo crearemos líneas de albarán, si Cantidad>0 o si es una devolución.
                If (DirectCast(htLins(rwLin(_PCL.IDLineaPedido)), CrearAlbaranCompraInfo).Cantidad > 0 OrElse _
                   (DirectCast(htLins(rwLin(_PCL.IDLineaPedido)), CrearAlbaranCompraInfo).Cantidad < 0 AndAlso _
                    (rwLin(_PCL.Estado) = enumpclEstado.pclparcservido OrElse rwLin(_PCL.Estado) = enumpclEstado.pclservido))) OrElse _
                    (rwLin(_PCL.QPedida) < 0 AndAlso (rwLin(_PCL.Estado) = enumpclEstado.pclparcservido OrElse rwLin(_PCL.Estado) = enumpclEstado.pclservido)) Then
                    groupers(rwLin("AgrupAlbaran")).Group(rwLin)
                End If
            Next

            For Each alb As AlbCabPedidoCompra In oGrprUser.Albs
                For Each alblin As AlbLinPedidoCompra In alb.Lineas
                    alblin.QaRecibir = DirectCast(htLins(alblin.IDLineaPedido), CrearAlbaranCompraInfo).Cantidad
                    alblin.Cantidad = DirectCast(htLins(alblin.IDLineaPedido), CrearAlbaranCompraInfo).CantidadUD
                    If Length(DirectCast(htLins(alblin.IDLineaPedido), CrearAlbaranCompraInfo).Cantidad2) > 0 Then
                        alblin.Cantidad2 = DirectCast(htLins(alblin.IDLineaPedido), CrearAlbaranCompraInfo).Cantidad2
                    End If
                    alblin.FechaEntregaModificado = DirectCast(htLins(alblin.IDLineaPedido), CrearAlbaranCompraInfo).FechaEntregaModificado
                    alblin.Lotes = DirectCast(htLins(alblin.IDLineaPedido), CrearAlbaranCompraInfo).Lotes
                    alblin.Series = DirectCast(htLins(alblin.IDLineaPedido), CrearAlbaranCompraInfo).Series
                Next
            Next

            Return oGrprUser.Albs
        End If
    End Function

    <Task()> Public Shared Function GetGroupColumns(ByVal data As DataColAgrupacionAC, ByVal services As ServiceProvider) As DataColumn()
        Dim columns(3) As DataColumn
        columns(0) = data.Lineas.Columns("IDProveedor")
        columns(1) = data.Lineas.Columns("IDMoneda")
        columns(2) = data.Lineas.Columns("IDTipoCompra")
        columns(3) = data.Lineas.Columns("IDFormaEnvio")
        If data.TipoAgrupacion = enummpAgrupAlbaran.mpPedido Then
            ReDim Preserve columns(4)
            columns(4) = data.Lineas.Columns("IDPedido")
        End If

        Return columns
    End Function
#End Region

#Region " Creación del Documento Albarán (PrcCrearAlbPedCompra) "

#Region " Cabecera (Documento) "

    <Task()> Public Shared Function CrearDocumentoAlbaranCompra(ByVal alb As AlbCabCompra, ByVal services As ServiceProvider) As DocumentoAlbaranCompra
        Return New DocumentoAlbaranCompra(alb, services)
    End Function

    <Task()> Public Shared Sub AsignarValoresPredeterminadosGenerales(ByVal alb As DocumentoAlbaranCompra, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoComunes.AsignarIdentificadorAlbaran, alb.HeaderRow, services)
        Dim InfoProc As ProcessInfoAlbCompra = services.GetService(Of ProcessInfoAlbCompra)()
        If alb.HeaderRow.IsNull("FechaAlbaran") Or alb.HeaderRow("FechaAlbaran") = cnMinDate Then
            If InfoProc.FechaAlbaran <> cnMinDate Then
                alb.Fecha = InfoProc.FechaAlbaran
            Else
                ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoComunes.AsignarFechaAlbaran, alb.HeaderRow, services)
            End If
            Dim data As New DataEjercicio(New DataRowPropertyAccessor(alb.HeaderRow), alb.HeaderRow("FechaAlbaran"))
            ProcessServer.ExecuteTask(Of DataEjercicio)(AddressOf NegocioGeneral.AsignarEjercicioContable, data, services)
        End If

        If alb.HeaderRow.IsNull("Estado") Then alb.HeaderRow("Estado") = enumaccEstado.accNoFacturado
        If alb.HeaderRow.IsNull("Automatico") Then alb.HeaderRow("Automatico") = False
        If alb.HeaderRow.IsNull("IDTipoCompra") Then alb.HeaderRow("IDTipoCompra") = InfoProc.IDTipoCompra
    End Sub

    <Task()> Public Shared Sub AsignarDatosProveedor(ByVal alb As DocumentoAlbaranCompra, ByVal services As ServiceProvider)
        If alb.HeaderRow.IsNull("IDProveedor") Then Exit Sub
        ProcessServer.ExecuteTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoCompra.AsignarDatosProveedor, alb, services)

        Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
        If alb.Proveedor Is Nothing Then alb.Proveedor = Proveedores.GetEntity(alb.HeaderRow("IDProveedor"))

        If alb.HeaderRow.IsNull("IDFormaEnvio") Then alb.HeaderRow("IDFormaEnvio") = alb.Proveedor.IDFormaEnvio
        If alb.HeaderRow.IsNull("IDCondicionEnvio") Then alb.HeaderRow("IDCondicionEnvio") = alb.Proveedor.IDCondicionEnvio
        If alb.HeaderRow.IsNull("IDModoTransporte") Then alb.HeaderRow("IDModoTransporte") = alb.Proveedor.IDModoTransporte
        If alb.HeaderRow.IsNull("Dto") Then alb.HeaderRow("Dto") = alb.Proveedor.DtoComercial
    End Sub

    <Task()> Public Shared Sub AsignarDireccion(ByVal alb As DocumentoAlbaranCompra, ByVal services As ServiceProvider)
        If Not alb.HeaderRow.IsNull("IDDireccion") Then
            Dim StDatosDirec As New ProveedorDireccion.DataDirecDe
            StDatosDirec.IDDireccion = alb.HeaderRow("IDDireccion")
            StDatosDirec.TipoDireccion = enumpdTipoDireccion.pdDireccionPedido
            If Not ProcessServer.ExecuteTask(Of ProveedorDireccion.DataDirecDe, Boolean)(AddressOf ProveedorDireccion.EsDireccionDe, StDatosDirec, services) Then
                Dim DIR As New DataDireccionProv(enumpdTipoDireccion.pdDireccionPedido, "IDDireccion", New DataRowPropertyAccessor(alb.HeaderRow))
                ProcessServer.ExecuteTask(Of DataDireccionProv)(AddressOf ProcesoCompra.AsignarDireccionProveedor, DIR, services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarContador(ByVal alb As DocumentoAlbaranCompra, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoComunes.AsignarContador, alb, services)
        If Length(alb.HeaderRow("IDContador")) = 0 Then ApplicationService.GenerateError("Debe indicar un Contador para la entidad {0}.", Quoted(GetType(AlbaranCompraCabecera).Name))
    End Sub

    <Task()> Public Shared Sub AsignarCentroCoste(ByVal Doc As DocumentoAlbaranCompra, ByVal services As ServiceProvider)
        If Not Doc.Cabecera Is Nothing AndAlso TypeOf Doc.Cabecera Is AlbCabPedidoCompra Then
            If Doc.HeaderRow.IsNull("IDCentroCoste") Then Doc.HeaderRow("IDCentroCoste") = CType(Doc.Cabecera, AlbCabPedidoCompra).IDCentroCoste
        End If
    End Sub

#End Region

#Region " Lineas (Documento) "

    Public Class DataLineasDesdePedidoCompra
        Public Row As DataRow
        Public Pedido As DataRow
        Public Cantidad As Double
        Public NSerie As String
        Public Ubicacion As String
        Public IDEstadoActivo As String
        Public IDOperario As String

        Public Doc As DocumentoAlbaranCompra
        Public AlbLin As AlbLinPedidoCompra

        Public Sub New(ByVal Row As DataRow, ByVal pedido As DataRow, ByVal Doc As DocumentoAlbaranCompra, ByVal AlbLin As AlbLinPedidoCompra, Optional ByVal RowSerie As DataRow = Nothing)
            Me.Row = Row
            Me.Pedido = pedido
            Me.Doc = Doc
            Me.AlbLin = AlbLin
            If Not RowSerie Is Nothing Then
                Me.NSerie = RowSerie("NSerie")
                Me.Cantidad = IIf(AlbLin.QaRecibir > 0, 1, -1)
                Me.IDEstadoActivo = RowSerie("IDEstadoActivo")
                Me.IDOperario = RowSerie("IDOperario")
                If RowSerie.Table.Columns.Contains("Ubicacion") AndAlso Length(RowSerie("Ubicacion")) > 0 Then Me.Ubicacion = RowSerie("Ubicacion")
            End If
        End Sub
    End Class

    <Task()> Public Shared Sub AsignarDatosTransporte(ByVal docAlbaran As DocumentoAlbaranCompra, ByVal services As ServiceProvider)
        If Not docAlbaran.HeaderRow("IDFormaEnvio") Is Nothing Then
            Dim FE As New FormaEnvio
            Dim detalle As New FormaEnvioDetalle
            Dim filtro As New Filter
            If Not docAlbaran.HeaderRow("IDFormaEnvio") Is Nothing Then
                filtro.Add("IDFormaEnvio", docAlbaran.HeaderRow("IDFormaEnvio"))
                Dim dt As DataTable = FE.Filter(filtro)
                If dt.Rows.Count > 0 Then
                    docAlbaran.HeaderRow("EmpresaTransp") = Nz(dt.Rows(0)("DescFormaEnvio"), String.Empty)
                    filtro.Add("Predeterminado", True)
                    Dim dtDetalle As DataTable = detalle.Filter(filtro)
                    If dtDetalle.Rows.Count > 0 Then
                        docAlbaran.HeaderRow("Conductor") = Nz(dtDetalle.Rows(0)("Conductor"), String.Empty)
                        docAlbaran.HeaderRow("DNIConductor") = Nz(dtDetalle.Rows(0)("DNIConductor"), String.Empty)
                        docAlbaran.HeaderRow("Matricula") = Nz(dtDetalle.Rows(0)("Matricula"), String.Empty)
                        docAlbaran.HeaderRow("Remolque") = Nz(dtDetalle.Rows(0)("Remolque"), String.Empty)
                    End If

                    Dim fProv As New Filter
                    fProv.Add("IDProveedor", Nz(dt.Rows(0)("IDProveedor"), String.Empty))
                    Dim Prov As New Proveedor
                    Dim dtProv As DataTable = Prov.Filter(fProv)
                    If dtProv.Rows.Count > 0 Then
                        docAlbaran.HeaderRow("CifTransportista") = Nz(dtProv.Rows(0)("CifProveedor"), String.Empty)
                    End If
                End If
            End If
        End If
    End Sub


    <Task()> Public Shared Sub TotalPesos(ByVal Doc As DocumentoAlbaranCompra, ByVal services As ServiceProvider)
        If Not IsNothing(Doc.HeaderRow) Then
            Dim PesoNetoManual As Double = Nz(Doc.HeaderRow("PesoNetoManual"), 0)
            Dim PesoBrutoManual As Double = Nz(Doc.HeaderRow("PesoBrutoManual"), 0)

            'Completar los pesos
            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
            Dim PesoNetoAcum As Double = 0
            Dim PesoBrutoAcum As Double = 0
            For Each linea As DataRow In Doc.dtLineas.Rows
                Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(linea("IDArticulo"))
                PesoNetoAcum += (Nz(linea("QInterna"), 0) * ArtInfo.PesoNeto)
                PesoBrutoAcum += (Nz(linea("QInterna"), 0) * ArtInfo.PesoBruto)
            Next
            Doc.HeaderRow("PesoNeto") = Math.Abs(PesoNetoAcum)
            Doc.HeaderRow("PesoBruto") = Math.Abs(PesoBrutoAcum)
            If Nz(Doc.HeaderRow("PesoNetoManual"), 0) = 0 Then Doc.HeaderRow("PesoNetoManual") = Math.Abs(PesoNetoAcum)
            If Nz(Doc.HeaderRow("PesoBrutoManual"), 0) = 0 Then Doc.HeaderRow("PesoBrutoManual") = Math.Abs(PesoBrutoAcum)

            Dim PesoNeto As Double = Nz(Doc.HeaderRow("PesoNeto"), 0)
            Dim PesoBruto As Double = Nz(Doc.HeaderRow("PesoBruto"), 0)

            If PesoNetoManual = 0 Then PesoNetoManual = Nz(Doc.HeaderRow("PesoNetoManual"), 0)
            If PesoBrutoManual = 0 Then PesoBrutoManual = Nz(Doc.HeaderRow("PesoBrutoManual"), 0)

            Doc.HeaderRow("PesoNeto") = PesoNeto
            Doc.HeaderRow("PesoBruto") = PesoBruto
            Doc.HeaderRow("PesoNetoManual") = PesoNetoManual
            Doc.HeaderRow("PesoBrutoManual") = PesoBrutoManual
        End If
    End Sub

    <Task()> Public Shared Sub CrearLineasDesdePedido(ByVal docAlbaran As DocumentoAlbaranCompra, ByVal services As ServiceProvider)
        'Dim lineas As DataTable = docAlbaran.dtLineas
        'If lineas Is Nothing Then
        '    Dim oACL As New AlbaranCompraLinea
        '    lineas = oACL.AddNew
        '    docAlbaran.Add(GetType(AlbaranCompraLinea).Name, lineas)
        'End If

        Dim PCCabeceras As EntityInfoCache(Of PedidoCompraCabeceraInfo) = services.GetService(Of EntityInfoCache(Of PedidoCompraCabeceraInfo))()
        Dim dtPedido As DataTable = ProcessServer.ExecuteTask(Of DocumentoAlbaranCompra, DataTable)(AddressOf RecuperarDatosPedido, docAlbaran, services)

        For Each pedido As DataRow In dtPedido.Rows
            Dim alblin As AlbLinPedidoCompra = Nothing
            For i As Integer = 0 To CType(docAlbaran.Cabecera, AlbCabPedidoCompra).Lineas.Length - 1
                If pedido("IDLineaPedido") = CType(docAlbaran.Cabecera, AlbCabPedidoCompra).Lineas(i).IDLineaPedido Then
                    alblin = CType(docAlbaran.Cabecera, AlbCabPedidoCompra).Lineas(i)
                    Exit For
                End If
            Next

            If Not alblin Is Nothing Then
                Dim dblCantidad As Double
                If Double.IsNaN(alblin.QaRecibir) Then
                    dblCantidad = pedido("QPedida") - pedido("QServida")
                Else
                    dblCantidad = alblin.QaRecibir
                End If
                If dblCantidad <> 0 Then
                    Dim NumLineasInsertar As Integer = 1
                    If alblin.QaRecibir > 1 AndAlso Not alblin.Series Is Nothing AndAlso alblin.Series.Rows.Count > 0 Then NumLineasInsertar = alblin.Series.Rows.Count
                    Dim linAlbPed As DataLineasDesdePedidoCompra
                    For i As Integer = NumLineasInsertar - 1 To 0 Step -1
                        Dim linea As DataRow = docAlbaran.dtLineas.NewRow

                        If Not alblin.Series Is Nothing AndAlso alblin.Series.Rows.Count > 0 Then
                            linAlbPed = New DataLineasDesdePedidoCompra(linea, pedido, docAlbaran, alblin, alblin.Series.Rows(i))
                        Else
                            linAlbPed = New DataLineasDesdePedidoCompra(linea, pedido, docAlbaran, alblin)
                        End If

                        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarValoresPredeterminadosLinea, linea, services)
                        ProcessServer.ExecuteTask(Of DataLineasDesdePedidoCompra)(AddressOf AsignarIDAlbaran, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasDesdePedidoCompra)(AddressOf AsignarAlmacen, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasDesdePedidoCompra)(AddressOf AsignarFechaEntregaModificado, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasDesdePedidoCompra)(AddressOf AsignarDatosOrigen, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasDesdePedidoCompra)(AddressOf AsignarDatosArticulo, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasDesdePedidoCompra)(AddressOf AsignarCuenta, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasDesdePedidoCompra)(AddressOf AsignarDatosEconomicos, linAlbPed, services)
                        If alblin.Series Is Nothing OrElse alblin.Series.Rows.Count = 0 Then linAlbPed.Cantidad = dblCantidad
                        ProcessServer.ExecuteTask(Of DataLineasDesdePedidoCompra)(AddressOf AsignarCantidades, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasDesdePedidoCompra)(AddressOf AsignarUnidadesCantidadesSegundaUnidad, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasDesdePedidoCompra)(AddressOf AsignarPrecios, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasDesdePedidoCompra)(AddressOf AsignarTexto, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasDesdePedidoCompra)(AddressOf AsignarDatosOT, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasDesdePedidoCompra)(AddressOf AsignarDatosObras, linAlbPed, services)
                        'ProcessServer.ExecuteTask(Of DataLineasDesdePedidoCompra)(AddressOf AsignarDatosActivos, linAlbPed, services)     'PENDIENTE
                        'ProcessServer.ExecuteTask(Of DataLineasDesdePedidoCompra)(AddressOf AsignarOperario, linAlbPed, services)         'PENDIENTE
                        ProcessServer.ExecuteTask(Of DataLineasDesdePedidoCompra)(AddressOf AsignarNSerie, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasDesdePedidoCompra)(AddressOf AsignarDatosAlquiler, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasDesdePedidoCompra)(AddressOf AsignarControlCalidad, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasDesdePedidoCompra)(AddressOf AsignarLineasAdicionales, linAlbPed, services)

                        docAlbaran.dtLineas.Rows.Add(linea.ItemArray)
                    Next
                End If
            End If
        Next
    End Sub

    <Task()> Public Shared Function RecuperarDatosPedido(ByVal docAlbaran As DocumentoAlbaranCompra, ByVal services As ServiceProvider) As DataTable
        Dim albCabPed As AlbCabPedidoCompra = docAlbaran.Cabecera

        Dim ids(albCabPed.Lineas.Length - 1) As Object
        For i As Integer = 0 To ids.Length - 1
            ids(i) = albCabPed.Lineas(i).IDLineaPedido
        Next

        Dim oFltr As New Filter
        oFltr.Add(New InListFilterItem("IDlineaPedido", ids, FilterType.Numeric))
        oFltr.Add(New NumberFilterItem("Estado", FilterOperator.NotEqual, enumpclEstado.pclCerrado))
        oFltr.Add(New NumberFilterItem("TipoLineaCompra", FilterOperator.NotEqual, enumaclTipoLineaAlbaran.aclComponente))

        Return New PedidoCompraLinea().Filter(oFltr)
    End Function

    <Task()> Public Shared Sub AsignarValoresPredeterminadosLinea(ByVal row As DataRow, ByVal services As ServiceProvider)
        row("IdLineaAlbaran") = AdminData.GetAutoNumeric
        row("EstadoFactura") = enumaclEstadoFactura.aclNoFacturado
    End Sub
    <Task()> Public Shared Sub AsignarIDAlbaran(ByVal data As DataLineasDesdePedidoCompra, ByVal services As ServiceProvider)
        data.Row(_ACL.IDAlbaran) = data.Doc.HeaderRow("IDAlbaran")
    End Sub
    <Task()> Public Shared Sub AsignarAlmacen(ByVal data As DataLineasDesdePedidoCompra, ByVal services As ServiceProvider)
        data.Row(_ACL.IDAlmacen) = data.Pedido(_PCL.IDAlmacen)
    End Sub
    <Task()> Public Shared Sub AsignarFechaEntregaModificado(ByVal data As DataLineasDesdePedidoCompra, ByVal services As ServiceProvider)
        If CDate(data.AlbLin.FechaEntregaModificado) <> cnMinDate Then
            data.Row(_ACL.FechaEntregaModificado) = data.AlbLin.FechaEntregaModificado
        Else
            data.Row(_ACL.FechaEntregaModificado) = Today
        End If
    End Sub
    <Task()> Public Shared Sub AsignarDatosOrigen(ByVal data As DataLineasDesdePedidoCompra, ByVal services As ServiceProvider)
        data.Row(_ACL.IDPedido) = data.Pedido(_PCL.IDPedido)
        data.Row(_ACL.IDLineaPedido) = data.Pedido(_PCL.IDLineaPedido)
        data.Row(_ACL.IdCentroGestion) = data.Pedido(_PCL.IDCentroGestion)

        If Length(data.Pedido(_PCL.IdContrato)) > 0 Then
            data.Row(_ACL.IdContrato) = data.Pedido(_PCL.IdContrato)
        End If
        If Length(data.Pedido(_PCL.IdLineaContrato)) > 0 Then
            data.Row(_ACL.IdLineaContrato) = data.Pedido(_PCL.IdLineaContrato)
        End If

        data.Row(_ACL.TipoLineaAlbaran) = data.Pedido(_PCL.TipoLineaCompra)
        data.Row(_ACL.SeguimientoTarifa) = data.Pedido(_PCL.SeguimientoTarifa)
    End Sub
    <Task()> Public Shared Sub AsignarDatosArticulo(ByVal data As DataLineasDesdePedidoCompra, ByVal services As ServiceProvider)
        data.Row(_ACL.IDArticulo) = data.Pedido(_PCL.IDArticulo)
        data.Row(_ACL.DescArticulo) = data.Pedido(_PCL.DescArticulo)
        data.Row(_ACL.RefProveedor) = data.Pedido(_PCL.RefProveedor)
        data.Row(_ACL.DescRefProveedor) = data.Pedido(_PCL.DescRefProveedor)
        'El estado del stock de la línea depende de si el artículo tiene gestión de stock o no
        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(GetType(EntityInfoCache(Of ArticuloInfo)))
        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.Row(_ACL.IDArticulo))
        data.Row(_ACL.EstadoStock) = enumaclEstadoStock.aclNoActualizado
        If Not ArtInfo.GestionStock Then
            data.Row(_ACL.EstadoStock) = enumaclEstadoStock.aclSinGestion
        End If
    End Sub
    <Task()> Public Shared Sub AsignarCuenta(ByVal data As DataLineasDesdePedidoCompra, ByVal services As ServiceProvider)
        Dim AppParamsConta As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
        If AppParamsConta.Contabilidad Then data.Row(_ACL.CContable) = data.Pedido(_PCL.CContable)
    End Sub
    <Task()> Public Shared Sub AsignarDatosEconomicos(ByVal data As DataLineasDesdePedidoCompra, ByVal services As ServiceProvider)
        Dim PCCabeceras As EntityInfoCache(Of PedidoCompraCabeceraInfo) = services.GetService(Of EntityInfoCache(Of PedidoCompraCabeceraInfo))()
        Dim CabPedidoInfo As PedidoCompraCabeceraInfo = PCCabeceras.GetEntity(data.Pedido(_PCL.IDPedido))
        If Length(CabPedidoInfo.IDCondicionPago) > 0 Then
            data.Row(_ACL.IDCondicionPago) = CabPedidoInfo.IDCondicionPago
        Else
            ApplicationService.GenerateError("La Condición Pago no existe.")
        End If
        If Length(CabPedidoInfo.IDFormaPago) > 0 Then
            data.Row(_ACL.IDFormaPago) = CabPedidoInfo.IDFormaPago
        Else
            ApplicationService.GenerateError("La Forma Pago no existe.")
        End If
        data.Row(_ACL.Dto) = CabPedidoInfo.Dto
        data.Row(_ACL.IDTipoIva) = data.Pedido(_PCL.IDTipoIva)
    End Sub
    <Task()> Public Shared Sub AsignarCantidades(ByVal data As DataLineasDesdePedidoCompra, ByVal services As ServiceProvider)
        data.Row(_ACL.IDUdMedida) = data.Pedido(_PCL.IDUdMedida)
        data.Row(_ACL.UdValoracion) = data.Pedido(_PCL.UdValoracion)
        data.Row(_ACL.QServida) = data.Cantidad
        'data.Row(_ACL.Factor) = data.Pedido(_PCL.Factor)
        'data.Row(_ACL.QInterna) = data.Row(_ACL.QServida) * data.Pedido(_PCL.Factor)
        'If data.Pedido(_PCL.QPedida) <> 0 Then
        '    If data.Row(_ACL.QInterna) <> (data.Row(_ACL.QServida) * (data.Pedido(_PCL.QInterna) / data.Pedido(_PCL.QPedida))) Then
        '        data.Row(_ACL.QInterna) = data.Row(_ACL.QServida) * (data.Pedido(_PCL.QInterna) / data.Pedido(_PCL.QPedida))
        '    End If
        'End If
        If data.Row(_ACL.QServida) <> 0 Then
            If Length(data.NSerie) > 0 Then
                data.Row(_ACL.Factor) = data.Pedido(_PCL.Factor)
                data.Row(_ACL.QInterna) = data.Cantidad * data.Pedido(_PCL.Factor)
            Else
                data.Row(_ACL.QInterna) = data.AlbLin.Cantidad
                data.Row(_ACL.Factor) = data.Row(_ACL.QInterna) / data.Row(_ACL.QServida)
            End If
        Else
            data.Row(_ACL.QInterna) = 0
            data.Row(_ACL.Factor) = 0
        End If
        data.Row(_ACL.IDUdInterna) = data.Pedido(_PCL.IDUdInterna)
    End Sub

    <Task()> Public Shared Sub AsignarUnidadesCantidadesSegundaUnidad(ByVal data As DataLineasDesdePedidoCompra, ByVal services As ServiceProvider)
        If ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.Row("IDArticulo"), services) Then
            If TypeOf data.AlbLin Is AlbLinPedidoCompra Then
                If Length(CType(data.AlbLin, AlbLinPedidoCompra).Cantidad2) = 0 Then
                    ApplicationService.GenerateError("El Articulo {0} se gestiona con Doble Unidad. Debe indicar la misma.", Quoted(data.Row("IDArticulo")))
                Else
                    data.Row("QInterna2") = CType(data.AlbLin, AlbLinPedidoCompra).Cantidad2
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarPrecios(ByVal data As DataLineasDesdePedidoCompra, ByVal services As ServiceProvider)
        data.Row(_ACL.Precio) = data.Pedido(_PCL.Precio)
        data.Row(_ACL.Dto1) = data.Pedido(_PCL.Dto1)
        data.Row(_ACL.Dto2) = data.Pedido(_PCL.Dto2)
        data.Row(_ACL.Dto3) = data.Pedido(_PCL.Dto3)
        data.Row(_ACL.Dto) = data.Pedido(_PCL.Dto)
        data.Row(_ACL.DtoProntoPago) = data.Pedido(_PCL.DtoProntoPago)
    End Sub
    <Task()> Public Shared Sub AsignarTexto(ByVal data As DataLineasDesdePedidoCompra, ByVal services As ServiceProvider)
        data.Row(_ACL.Texto) = data.Pedido(_PCL.Texto)
    End Sub
    <Task()> Public Shared Sub AsignarDatosOT(ByVal data As DataLineasDesdePedidoCompra, ByVal services As ServiceProvider)
        data.Row(_ACL.IDMntoOTPrev) = data.Pedido(_PCL.IDMntoOTPrev)
        data.Row(_ACL.IdOrdenLinea) = data.Pedido(_PCL.IdOrdenLinea)
        data.Row(_ACL.IDOrdenRuta) = data.Pedido(_PCL.IDOrdenRuta)
    End Sub
    <Task()> Public Shared Sub AsignarDatosObras(ByVal data As DataLineasDesdePedidoCompra, ByVal services As ServiceProvider)
        data.Row(_ACL.IDTrabajo) = data.Pedido(_PCL.IDTrabajo)
        data.Row(_ACL.IDObra) = data.Pedido(_PCL.IDObra)
        data.Row(_ACL.IDLineaPadre) = data.Pedido(_PCL.IDLineaMaterial)
        data.Row(_ACL.TipoGastoObra) = enumfclTipoGastoObra.enumfclMaterial
    End Sub

    <Task()> Public Shared Sub AsignarNSerie(ByVal data As DataLineasDesdePedidoCompra, ByVal services As ServiceProvider)
        Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
        If AppParams.GestionNumeroSerieConActivos AndAlso Length(data.Row(_ACL.IDActivoAImputar)) > 0 Then
            data.Row(_ACL.IDActivoAImputar) = data.Pedido(_PCL.IDActivoAImputar)
        End If

        If data.Row(_ACL.TipoLineaAlbaran) <> CInt(enumaclTipoLineaAlbaran.aclComponente) Then
            data.Row(_ACL.Lote) = data.NSerie
        Else
            data.Row(_ACL.Lote) = Nothing
        End If
        If Length(data.IDEstadoActivo) > 0 Then data.Row(_ACL.IDEstadoActivo) = data.IDEstadoActivo
        If Length(data.IDOperario) > 0 Then data.Row(_ACL.IDOperario) = data.IDOperario
        If Length(data.Ubicacion) > 0 Then data.Row(_ACL.Ubicacion) = data.Ubicacion
    End Sub
    <Task()> Public Shared Sub AsignarDatosAlquiler(ByVal data As DataLineasDesdePedidoCompra, ByVal services As ServiceProvider)
        If data.Row.Table.Columns.Contains(_ACL.QTiempo) Then
            data.Row(_ACL.QTiempo) = data.Pedido(_PCL.QTiempo)
            ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf NegocioGeneral.CalcularPrecioImporte, New DataRowPropertyAccessor(data.Row), services)
        End If
    End Sub
    <Task()> Public Shared Sub AsignarControlCalidad(ByVal data As DataLineasDesdePedidoCompra, ByVal services As ServiceProvider)
        Dim AppParams As ParametroAlbaranCompra = services.GetService(Of ParametroAlbaranCompra)()
        If Not AppParams.ExpertisSAAS Then
            Dim datControlCal As New DataControlCalidad(data.Row("IDArticulo"), data.Doc.HeaderRow("IDProveedor"))
            data.Row(_ACL.ControlCalidad) = ProcessServer.ExecuteTask(Of DataControlCalidad, Boolean)(AddressOf ControlCalidad, datControlCal, services)
        End If
    End Sub
    <Task()> Public Shared Sub AsignarLineasAdicionales(ByVal data As DataLineasDesdePedidoCompra, ByVal services As ServiceProvider)
        '// Introducido por la incorporacion de la gestion de subcontrataciones.
        'El campo 'TipoLineaAlbaran' viene dado en principio por el tipo de la linea de pedido.
        'La linea del pedido puede ser:
        '1.Normal
        '2.Sucontratacion
        '3.Componente
        'Si en principio el tipo es 'Normal', hay que  comprobar si el artículo es kit, en cuyo caso habria que insertar las líneas de sus componentes.
        'Si es de tipo 'Subcontratacion' hay que comprobar si tiene lineas de pedido 'componentes' para añadirlas en el albaran.
        'Si es de tipo 'Componente' no hay que hacer nada
        'La función devuelve los componentes como lineas de albaran listas para ser actualizadas.


        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.Row("IDArticulo"))

        Dim componentes As DataTable
        'Dim acl As New AlbaranCompraLinea
        If data.Row(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclSubcontratacion Then
            componentes = ProcessServer.ExecuteTask(Of DataRow, DataTable)(AddressOf ComponentesDePedido, data.Row, services)
        ElseIf ArtInfo.KitVenta AndAlso data.Row(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclNormal Then
            componentes = ProcessServer.ExecuteTask(Of DataRow, DataTable)(AddressOf ComponentesDePrimerNivel, data.Row, services)
        End If
        If Not componentes Is Nothing AndAlso componentes.Rows.Count > 0 Then
            If data.Row(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclNormal Then
                data.Row(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclKit
            End If
            For Each componente As DataRow In componentes.Rows
                ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf NegocioGeneral.CalcularPrecioImporte, New DataRowPropertyAccessor(data.Row), services)
                Dim ValAyB As New ValoresAyB(New DataRowPropertyAccessor(data.Row), data.Doc.IDMoneda, data.Doc.CambioA, data.Doc.CambioB)
                ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)

                data.Doc.dtLineas.Rows.Add(componente.ItemArray)
            Next
        End If
    End Sub

#Region " CALIDAD "

    <Serializable()> _
    Public Class DataControlCalidad
        Public IDArticulo As String
        Public IDProveedor As String

        Public Sub New(ByVal IDArticulo As String, ByVal IDProveedor As String)
            Me.IDArticulo = IDArticulo
            Me.IDProveedor = IDProveedor
        End Sub
    End Class

    <Task()> Public Shared Function ControlCalidad(ByVal data As DataControlCalidad, ByVal services As ServiceProvider) As Boolean
        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(GetType(EntityInfoCache(Of ArticuloInfo)))
        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.IDArticulo)

        Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
        Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(data.IDProveedor)

        If Not IsNothing(ArtInfo) Then
            Select Case ArtInfo.ControlRecepcion
                Case enumControlRecepcion.crNunca
                    Return False
                Case enumControlRecepcion.crSiempre
                    Return True
                Case enumControlRecepcion.crProcesoCalidad
                    If Not IsNothing(ProvInfo) Then
                        If Len(ProvInfo.IDProveedor) > 0 Then
                            If ProvInfo.Homologable Then
                                Dim BlnExisteArticuloProveedor As Boolean
                                Dim FilArtProv As New Filter
                                FilArtProv.Add("IDArticulo", FilterOperator.Equal, ArtInfo.IDArticulo, FilterType.String)
                                FilArtProv.Add("IDProveedor", FilterOperator.Equal, ProvInfo.IDProveedor, FilterType.String)
                                Dim DtArtProv As DataTable = New ArticuloProveedor().Filter(FilArtProv)
                                If Not DtArtProv Is Nothing Then
                                    BlnExisteArticuloProveedor = (DtArtProv.Rows.Count > 0)
                                End If
                                If BlnExisteArticuloProveedor Then
                                    If DtArtProv.Rows(0)("ControlCalidad") Then
                                        If ProvInfo.CalidadConcertada Then
                                            Return False
                                        Else
                                            Dim cal As New DataCalidad
                                            cal.IDCalificacion = ProvInfo.IDCalificacion
                                            cal.IDProveedor = ProvInfo.IDProveedor
                                            cal.IDArticulo = ArtInfo.IDArticulo
                                            Return ProcessServer.ExecuteTask(Of DataCalidad, Boolean)(AddressOf Calidad, cal, services)
                                        End If
                                    Else
                                        Return False
                                    End If
                                Else
                                    If ProvInfo.CalidadConcertada Then
                                        Return False
                                    Else
                                        Dim cal As New DataCalidad
                                        cal.IDCalificacion = ProvInfo.IDCalificacion
                                        cal.IDProveedor = ProvInfo.IDProveedor
                                        cal.IDArticulo = ArtInfo.IDArticulo
                                        Return ProcessServer.ExecuteTask(Of DataCalidad, Boolean)(AddressOf Calidad, cal, services)
                                    End If
                                End If
                            Else
                                Return False
                            End If
                        End If
                    End If
            End Select
        End If
    End Function

    <Serializable()> _
    Public Class DataCalidad
        Public IDCalificacion As String
        Public IDProveedor As String
        Public IDArticulo As String
    End Class

    <Task()> Public Shared Sub GestionCalidadArticulo(ByVal Doc As DocumentoAlbaranCompra, ByVal services As ServiceProvider)
        If Not Doc Is Nothing AndAlso Not Doc.dtLineas Is Nothing AndAlso Doc.dtLineas.Rows.Count > 0 Then
            'Dim datos As New DataCalidad
            'datos.IDProveedor = Doc.IDProveedor
            For Each lineaAlbaran As DataRow In Doc.dtLineas.Select
                If lineaAlbaran.RowState = DataRowState.Added Then
                    'datos.IDArticulo = lineaAlbaran("IDArticulo")
                    If lineaAlbaran(_ACL.QServida) >= 0 Then  'No se tienen en cuenta las devoluciones
                        Dim datos As New DataControlCalidad(lineaAlbaran("IDArticulo"), Doc.IDProveedor)
                        lineaAlbaran(_ACL.ControlCalidad) = ProcessServer.ExecuteTask(Of DataControlCalidad, Boolean)(AddressOf ControlCalidad, datos, services)
                    End If
                End If
            Next
        End If
    End Sub

    <Task()> Public Shared Function Calidad(ByVal data As DataCalidad, ByVal services As ServiceProvider) As Boolean
        'Circuito de Calidad
        Dim CALCONTROL, CALPERIODO, CALPUNTOS As Integer
        Dim CALULTRECP, CALDEM As Integer
        Dim IntDemeritosCalculados As Integer
        Dim IntFrecuencia As Integer, IntNRecepciones As Integer
        Dim ClsParam As New Parametro
        Dim BlnHayLineasDeAlbaran As Boolean

        Dim DtAux As DataTable = ClsParam.SelOnPrimaryKey("CALPERIODO")
        If Not DtAux Is Nothing Then If DtAux.Rows.Count > 0 Then CALPERIODO = DtAux.Rows(0)("Valor")
        DtAux = ClsParam.SelOnPrimaryKey("CALCONTROL")
        If Not DtAux Is Nothing Then If DtAux.Rows.Count > 0 Then CALCONTROL = DtAux.Rows(0)("Valor")
        DtAux = ClsParam.SelOnPrimaryKey("CALDEM")
        If Not DtAux Is Nothing Then If DtAux.Rows.Count > 0 Then CALDEM = DtAux.Rows(0)("Valor")

        'Establecemos el periodo de tiempo a evaluar
        'Evaluación contínua

        Dim DteFechaDesde As Date = Today.Date.AddDays(-CALPERIODO)
        Dim FilCal As New Filter
        FilCal.Add("FechaRecepcion", FilterOperator.GreaterThan, DteFechaDesde, FilterType.DateTime)
        FilCal.Add("IDProveedor", FilterOperator.Equal, data.IDProveedor, FilterType.String)
        DtAux = New BE.DataEngine().Filter("vNegControlCalidadRecepcion", FilCal, "*", "FechaRecepcion DESC")
        If Not DtAux Is Nothing AndAlso DtAux.Rows.Count > 0 Then
            IntNRecepciones = DtAux.Rows.Count
            For Each DrAux As DataRow In DtAux.Select
                If Length(DrAux("Demerito")) > 0 Then
                    IntDemeritosCalculados += DrAux("Demerito")
                End If
            Next
            IntDemeritosCalculados = IntDemeritosCalculados / IntNRecepciones
        Else
            IntDemeritosCalculados = CALDEM
        End If

        'Comprobamos si la puntuación obtenida en el tiempo evaluado es superior a la parametrizada en calcontrol
        CALPUNTOS = 100 + CALDEM
        'Para controlar si la fórmula es con 100 o con 101
        If CALCONTROL > CALPUNTOS - IntDemeritosCalculados Then
            Return True
        Else
            Dim BEDataEngine As New BE.DataEngine
            If Length(data.IDCalificacion) > 0 Then
                DtAux = BEDataEngine.Filter("tbMaestroCalificacion", New StringFilterItem("IDCalificacion", data.IDCalificacion))
                If Not DtAux Is Nothing AndAlso DtAux.Rows.Count > 0 Then
                    IntFrecuencia = Nz(DtAux.Rows(0)("ControlesTrasDemerito"), 0)
                End If
            End If
            If IntFrecuencia <> 0 Then
                CALULTRECP = IntFrecuencia
            Else
                DtAux = ClsParam.SelOnPrimaryKey("CALULTRECP")
                If Not DtAux Is Nothing Then If DtAux.Rows.Count > 0 Then CALULTRECP = DtAux.Rows(0)("Valor")
            End If
            Dim StrSelect As String = "DISTINCT TOP " & CALULTRECP & " * "
            Dim f As New Filter
            f.Add(New StringFilterItem("IDProveedor", data.IDProveedor))
            f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
            f.Add(New NumberFilterItem("TipoLineaAlbaran", FilterOperator.NotEqual, enumaclTipoLineaAlbaran.aclComponente))

            Dim StrOrderBy As String = "FechaAlbaran DESC, IDAlbaran DESC"
            DtAux = BEDataEngine.Filter("vNegControlCalidadAlbaran", f, StrSelect, StrOrderBy)
            If Not DtAux Is Nothing Then BlnHayLineasDeAlbaran = (DtAux.Rows.Count > 0)
            If BlnHayLineasDeAlbaran Then
                Dim DrsAux() As DataRow = DtAux.Select("ControlCalidad=1")
                If DrsAux.Length > 0 Then
                    Dim fRecep As New Filter(FilterUnionOperator.Or)
                    For Each DrRecep As DataRow In DtAux.Select
                        If Length(DrRecep("IDRecepcion")) > 0 Then
                            fRecep.Add("IDRecepcion", DrRecep("IDRecepcion"))
                        End If
                    Next
                    If fRecep.Count > 0 Then
                        IntDemeritosCalculados = CALDEM
                        IntNRecepciones = 0
                        DtAux = BEDataEngine.Filter("vNegControlCalidadRecepcion", fRecep, , "FechaRecepcion DESC")
                        If Not DtAux Is Nothing AndAlso DtAux.Rows.Count > 0 Then
                            IntDemeritosCalculados = 0
                            IntNRecepciones = DtAux.Rows.Count
                            For Each DrDem As DataRow In DtAux.Select
                                If Length(DrDem("Demerito")) > 0 Then
                                    IntDemeritosCalculados += DrDem("Demerito")
                                End If
                            Next
                            If IntNRecepciones > 0 And IntDemeritosCalculados > 0 Then
                                IntDemeritosCalculados = IntDemeritosCalculados / IntNRecepciones
                            End If
                        End If
                        Return (IntDemeritosCalculados <> CALDEM)
                    Else
                        Return True
                    End If
                Else
                    Return True
                End If
            Else
                Return True
            End If
        End If
    End Function


#End Region

#Region " Componentes "

    <Task()> Public Shared Function ComponentesDePedido(ByVal lineaAlbaran As DataRow, ByVal services As ServiceProvider) As DataTable
        Dim Componentes As DataTable = New AlbaranCompraLinea().AddNew

        Dim f As New Filter(FilterUnionOperator.Or)
        f.Add(New NumberFilterItem(_PCL.IDLineaPedido, FilterOperator.Equal, lineaAlbaran(_ACL.IDLineaPedido)))
        f.Add(New NumberFilterItem(_PCL.IDLineaPadre, FilterOperator.Equal, lineaAlbaran(_ACL.IDLineaPedido)))
        Dim PCL As New PedidoCompraLinea
        Dim Pedidos As DataTable = PCL.Filter(f)

        If Not Pedidos Is Nothing AndAlso Pedidos.Rows.Count > 0 Then
            f.Clear()
            f.Add(New NumberFilterItem(_PCL.IDLineaPedido, FilterOperator.Equal, lineaAlbaran(_ACL.IDLineaPedido)))
            Dim WhereLineaPedido As String = f.Compose(New AdoFilterComposer)
            Dim aux() As DataRow = Pedidos.Select(WhereLineaPedido)
            If aux.Length > 0 Then
                Dim lineaPadrePedido As DataRow = aux(0)
                Dim QPedida As Double = lineaPadrePedido(_PCL.QPedida)
                Dim UdsAB As New ArticuloUnidadAB
                Dim articulo As New Negocio.Articulo
                Dim factor As Double

                If QPedida <> 0 Then
                    For Each pedido As DataRow In Pedidos.Rows
                        If Length(pedido(_PCL.IDLineaPadre)) > 0 Then
                            factor = pedido(_PCL.QPedida) / QPedida

                            Dim newrow As DataRow = Componentes.NewRow
                            newrow(_ACL.IDLineaAlbaran) = AdminData.GetAutoNumeric
                            newrow(_ACL.IDAlbaran) = lineaAlbaran(_ACL.IDAlbaran)
                            newrow(_ACL.IDPedido) = lineaAlbaran(_ACL.IDPedido)
                            newrow(_ACL.IDLineaPedido) = pedido(_PCL.IDLineaPedido)
                            newrow(_ACL.IDArticulo) = pedido(_PCL.IDArticulo)
                            newrow(_ACL.DescArticulo) = pedido(_PCL.DescArticulo)
                            'los componentes de subcontratación cogen cada uno su almacén
                            newrow(_ACL.IDAlmacen) = pedido(_PCL.IDAlmacen)
                            newrow(_ACL.IDFormaPago) = lineaAlbaran(_ACL.IDFormaPago)
                            newrow(_ACL.IDCondicionPago) = lineaAlbaran(_ACL.IDCondicionPago)
                            newrow(_ACL.IDTipoIva) = lineaAlbaran(_ACL.IDTipoIva)
                            If Length(pedido(_PCL.IDUdMedida)) > 0 Then
                                newrow(_ACL.IDUdMedida) = pedido(_PCL.IDUdMedida)
                            End If
                            newrow(_ACL.IDUdInterna) = pedido(_PCL.IDUdInterna)
                            newrow(_ACL.QServida) = factor * lineaAlbaran(_ACL.QServida)
                            newrow(_ACL.Factor) = pedido(_PCL.Factor) 'UdsAB.FactorDeConversion(newrow(_ACL.IDArticulo), newrow(_ACL.IDUdMedida) & String.Empty, newrow(_ACL.IDUdInterna))
                            If newrow(_ACL.Factor) <= 0 Then
                                newrow(_ACL.Factor) = 1
                            End If
                            newrow(_ACL.QInterna) = newrow(_ACL.QServida) * newrow(_ACL.Factor)

                            newrow(_ACL.UdValoracion) = IIf(pedido(_PCL.UdValoracion) > 0, pedido(_PCL.UdValoracion), 1)
                            newrow(_ACL.Lote) = Nothing
                            newrow(_ACL.CContable) = lineaAlbaran(_ACL.CContable)
                            newrow(_ACL.Precio) = 0
                            newrow(_ACL.PrecioA) = 0
                            newrow(_ACL.PrecioB) = 0
                            newrow(_ACL.Importe) = 0
                            newrow(_ACL.ImporteA) = 0
                            newrow(_ACL.ImporteB) = 0
                            newrow(_ACL.EstadoFactura) = enumaclEstadoFactura.aclNoFacturado

                            '//El estado del stock de la línea depende de si el artículo tiene gestión de stock o no

                            Dim CaracteristicaArticulo As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf articulo.CaracteristicasArticulo, newrow(_ACL.IDArticulo), services)
                            If Not CaracteristicaArticulo Is Nothing AndAlso CaracteristicaArticulo.Rows.Count > 0 Then
                                If CaracteristicaArticulo.Rows(0)("GestionStock") Then
                                    newrow(_ACL.EstadoStock) = enumaclEstadoStock.aclNoActualizado
                                Else
                                    newrow(_ACL.EstadoStock) = enumaclEstadoStock.aclSinGestion
                                End If
                            Else
                                newrow(_ACL.EstadoStock) = enumaclEstadoStock.aclNoActualizado
                            End If

                            newrow(_ACL.IDOrdenRuta) = lineaAlbaran(_ACL.IDOrdenRuta)
                            newrow(_ACL.IDLineaPadre) = lineaAlbaran(_ACL.IDLineaAlbaran)
                            newrow(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclComponente
                            newrow(_ACL.Dto) = 0
                            newrow(_ACL.DtoProntoPago) = 0

                            Componentes.Rows.Add(newrow)
                        End If
                    Next
                End If
            End If
        End If

        Return Componentes
    End Function

    <Task()> Public Shared Function ComponentesDePrimerNivel(ByVal lineaAlbaran As DataRow, ByVal services As ServiceProvider) As DataTable
        Dim Kits As DataTable = New AlbaranCompraLinea().AddNew

        'las lineas de tipo kit entran en principio como lineas normales
        Dim f As New Filter
        f.Add(New StringFilterItem(_ACL.IDArticulo, FilterOperator.Equal, lineaAlbaran(_ACL.IDArticulo)))
        Dim Componentes As DataTable = New BE.DataEngine().Filter("vNegArticuloComponentesPrimerNivel", f)
        If Not Componentes Is Nothing AndAlso Componentes.Rows.Count > 0 Then

            Dim UdsAB As New ArticuloUnidadAB
            Dim articulo As New Negocio.Articulo
            For Each componente As DataRow In Componentes.Rows
                Dim newrow As DataRow = Kits.NewRow

                newrow(_ACL.IDLineaAlbaran) = AdminData.GetAutoNumeric
                newrow(_ACL.IDAlbaran) = lineaAlbaran(_ACL.IDAlbaran)
                newrow(_ACL.IDPedido) = lineaAlbaran(_ACL.IDPedido)
                newrow(_ACL.IDLineaPedido) = lineaAlbaran(_ACL.IDLineaPedido)
                newrow(_ACL.IDArticulo) = componente("IDComponente")
                newrow(_ACL.DescArticulo) = componente("DescComponente")
                'los kits cogen el almacén del padre



                newrow(_ACL.IDAlmacen) = lineaAlbaran(_ACL.IDAlmacen)
                newrow(_ACL.IDFormaPago) = lineaAlbaran(_ACL.IDFormaPago)
                newrow(_ACL.IDCondicionPago) = lineaAlbaran(_ACL.IDCondicionPago)
                newrow(_ACL.IDTipoIva) = lineaAlbaran(_ACL.IDTipoIva)
                If Length(componente("IDUdCompra")) > 0 Then
                    newrow(_ACL.IDUdMedida) = componente("IDUdCompra")
                End If
                newrow(_ACL.IDUdInterna) = componente("IDUdInterna")

                newrow(_ACL.QInterna) = lineaAlbaran(_ACL.QInterna) * componente("Cantidad")
                Dim StDatos As New ArticuloUnidadAB.DatosFactorConversion
                StDatos.IDArticulo = newrow(_ACL.IDArticulo)
                StDatos.IDUdMedidaA = newrow(_ACL.IDUdMedida)
                StDatos.IDUdMedidaB = newrow(_ACL.IDUdInterna)
                newrow(_ACL.Factor) = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, StDatos, services)
                If newrow(_ACL.Factor) <= 0 Then
                    newrow(_ACL.Factor) = 1
                End If
                newrow(_ACL.QServida) = newrow(_ACL.QInterna) / newrow(_ACL.Factor)

                newrow(_ACL.UdValoracion) = IIf(componente("UdValoracion") > 0, componente("UdValoracion"), 1)
                newrow(_ACL.Lote) = lineaAlbaran(_ACL.Lote)
                newrow(_ACL.CContable) = lineaAlbaran(_ACL.CContable)
                'No se valora la entrada
                newrow(_ACL.Precio) = 0
                newrow(_ACL.PrecioA) = 0
                newrow(_ACL.PrecioB) = 0
                newrow(_ACL.Importe) = 0
                newrow(_ACL.ImporteA) = 0
                newrow(_ACL.ImporteB) = 0
                newrow(_ACL.EstadoFactura) = enumaclEstadoFactura.aclNoFacturado

                'El estado del stock de la línea depende de si el artículo tiene gestión de stock o no
                Dim CaracteristicaArticulo As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf articulo.CaracteristicasArticulo, newrow(_ACL.IDArticulo), services)
                If Not CaracteristicaArticulo Is Nothing AndAlso CaracteristicaArticulo.Rows.Count > 0 Then
                    If CaracteristicaArticulo.Rows(0)("GestionStock") AndAlso componente("GestionStock") = False Then
                        newrow(_ACL.EstadoStock) = enumaclEstadoStock.aclNoActualizado
                    Else
                        newrow(_ACL.EstadoStock) = enumaclEstadoStock.aclSinGestion
                    End If
                Else
                    newrow(_ACL.EstadoStock) = enumaclEstadoStock.aclNoActualizado
                End If

                newrow(_ACL.IDOrdenRuta) = lineaAlbaran(_ACL.IDOrdenRuta)
                newrow(_ACL.IDLineaPadre) = lineaAlbaran(_ACL.IDLineaAlbaran)
                newrow(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclComponente

                Kits.Rows.Add(newrow)
            Next
        End If

        Return Kits
    End Function

    <Task()> Public Shared Function ComponentesDePrimerNivelDeSubcontratacion(ByVal lineaAlbaran As DataRow, ByVal services As ServiceProvider) As DataTable
        Dim newData As DataTable = New AlbaranCompraLinea().AddNew
        Dim f As New Filter
        f.Add(New StringFilterItem(_ACL.IDArticulo, FilterOperator.Equal, lineaAlbaran(_ACL.IDArticulo)))

        Dim Componentes As DataTable = New BE.DataEngine().Filter("vNegArticuloCompPrimerNivelSubcontratacion", f)
        If Not Componentes Is Nothing AndAlso Componentes.Rows.Count > 0 Then
            lineaAlbaran(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclSubcontratacion

            Dim UDS As New ArticuloUnidadAB
            Dim Articulo As New Negocio.Articulo
            For Each componente As DataRow In Componentes.Rows
                Dim newrow As DataRow = newData.NewRow

                newrow(_ACL.IDLineaAlbaran) = AdminData.GetAutoNumeric
                newrow(_ACL.IDAlbaran) = lineaAlbaran(_ACL.IDAlbaran)
                newrow(_ACL.IDPedido) = lineaAlbaran(_ACL.IDPedido)
                newrow(_ACL.IDLineaPedido) = lineaAlbaran(_ACL.IDLineaPedido)
                newrow(_ACL.IDArticulo) = componente("IDComponente")
                newrow(_ACL.DescArticulo) = componente("DescComponente")

                Dim data As New BusinessRuleData("IDAlmacen", lineaAlbaran(_ACL.IDAlmacen), New DataRowPropertyAccessor(newrow), Nothing)
                ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf NegocioGeneral.AsignarArticuloAlmacen, data, services)
                newrow(_ACL.IDAlmacen) = data.Current("IDAlmacen")

                newrow(_ACL.IDFormaPago) = lineaAlbaran(_ACL.IDFormaPago)
                newrow(_ACL.IDCondicionPago) = lineaAlbaran(_ACL.IDCondicionPago)
                newrow(_ACL.IDTipoIva) = lineaAlbaran(_ACL.IDTipoIva)
                newrow(_ACL.CContable) = lineaAlbaran(_ACL.CContable)
                If Length(componente("IdUdCompra")) > 0 Then
                    newrow(_ACL.IDUdMedida) = componente("IdUdCompra")
                End If
                newrow(_ACL.IDUdInterna) = componente("IDUdInterna")

                newrow(_ACL.QInterna) = lineaAlbaran(_ACL.QInterna) * componente("Cantidad")
                Dim StDatos As New ArticuloUnidadAB.DatosFactorConversion
                StDatos.IDArticulo = newrow(_ACL.IDArticulo)
                StDatos.IDUdMedidaA = newrow(_ACL.IDUdMedida)
                StDatos.IDUdMedidaB = newrow(_ACL.IDUdInterna)
                newrow(_ACL.Factor) = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, StDatos, services)
                If newrow(_ACL.Factor) <= 0 Then
                    newrow(_ACL.Factor) = 1
                End If
                newrow(_ACL.QServida) = newrow(_ACL.QInterna) / newrow(_ACL.Factor)
                newrow(_ACL.UdValoracion) = IIf(componente("UdValoracion") > 0, componente("UdValoracion"), 1)

                newrow(_ACL.IDLineaPadre) = lineaAlbaran(_ACL.IDLineaAlbaran)
                newrow(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclComponente
                'El estado del stock de la línea depende de si el artículo tiene gestión de stock o no
                Dim CaracteristicaArticulo As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf Articulo.CaracteristicasArticulo, newrow(_ACL.IDArticulo), services)
                If Not CaracteristicaArticulo Is Nothing AndAlso CaracteristicaArticulo.Rows.Count > 0 Then
                    If CaracteristicaArticulo.Rows(0)("GestionStock") Then
                        newrow(_ACL.EstadoStock) = enumaclEstadoStock.aclNoActualizado
                    Else
                        newrow(_ACL.EstadoStock) = enumaclEstadoStock.aclSinGestion
                    End If
                Else
                    newrow(_ACL.EstadoStock) = enumaclEstadoStock.aclNoActualizado
                End If
                newData.Rows.Add(newrow)
            Next
        End If

        Return newData
    End Function

#End Region

#End Region

#Region " Otras Entidades (Documento) "

    '<Task()> Public Shared Sub CalcularAlbaranCompraPrecio(ByVal docAlbaran As DocumentoAlbaranCompra, ByVal services As ServiceProvider)
    '    If Not docAlbaran.dtLineas Is Nothing AndAlso docAlbaran.dtLineas.Rows.Count > 0 Then
    '        Dim lineas As DataTable = docAlbaran.dtPrecios
    '        If lineas Is Nothing Then
    '            Dim oACP As New AlbaranCompraPrecio
    '            lineas = oACP.AddNew
    '            docAlbaran.Add(GetType(AlbaranCompraPrecio).Name, lineas)
    '        End If

    '        Dim PCP As New PedidoCompraPrecio
    '        Dim PCL As New PedidoCompraLinea
    '        For Each drACL As DataRow In docAlbaran.dtLineas.Rows
    '            Dim f As New Filter
    '            f.Add(New NumberFilterItem("IDLineaPedido", drACL("IDLineaPedido")))
    '            Dim dtPCP As DataTable = PCP.Filter(f)
    '            If Not dtPCP Is Nothing AndAlso dtPCP.Rows.Count > 0 Then
    '                Dim dblQPedida As Double : Dim dblQServida As Double
    '                Dim dtPCL As DataTable = PCL.Filter(f)
    '                If Not dtPCL Is Nothing AndAlso dtPCL.Rows.Count > 0 Then
    '                    dblQPedida = Nz(dtPCL.Rows(0)("QPedida"), 0)
    '                End If

    '                For Each drPCP As DataRow In dtPCP.Rows
    '                    Dim linea As DataRow = lineas.NewRow

    '                    linea("IDLineaAlbaranPrecio") = AdminData.GetAutoNumeric
    '                    linea("IDLineaAlbaran") = drACL("IDLineaAlbaran")
    '                    linea("IDArticulo") = drPCP("IDArticulo")
    '                    linea("DescArticulo") = drPCP("DescArticulo")
    '                    linea("Porcentaje") = drPCP("Porcentaje")
    '                    linea("Importe") = drPCP("Importe")
    '                    If dblQPedida <> dblQServida Then
    '                        If dblQPedida <> 0 Then
    '                            linea("Porcentaje") = linea("Porcentaje") * dblQServida / dblQPedida
    '                        Else
    '                            linea("Porcentaje") = 0
    '                        End If
    '                        linea("Importe") = linea("Importe") * IIf(linea("Porcentaje") <> 0, (linea("Porcentaje") / 100), 1)
    '                    End If


    '                    Dim lineaIProperty As New ValoresAyB(New DataRowPropertyAccessor(linea), docAlbaran.IDMoneda, docAlbaran.CambioA, docAlbaran.CambioB)
    '                    ProcessServer.ExecuteTask(Of ValoresAyB)(AddressOf General.MantenimientoValoresAyB, lineaIProperty, services)

    '                    If Length(drPCP("IDLineaPedidoHija")) > 0 Then
    '                        Dim dv As New DataView(docAlbaran.dtLineas.Copy)
    '                        dv.RowFilter = "IDLineaPedido=" & drPCP("IDLineaPedidoHija")
    '                        If dv.Count > 0 Then
    '                            linea("IDLineaAlbaranHija") = dv(0)("IDLineaAlbaran")
    '                        End If
    '                    End If

    '                    lineas.Rows.Add(linea.ItemArray)
    '                Next
    '            End If
    '        Next

    '    End If
    'End Sub
    <Task()> Public Shared Sub CalcularAlbaranCompraGastos(ByVal docAlbaran As DocumentoAlbaranCompra, ByVal services As ServiceProvider)
        If Not docAlbaran.dtLineas Is Nothing AndAlso docAlbaran.dtLineas.Rows.Count > 0 Then
            Dim f As New Filter : Dim ACP As New AlbaranCompraPrecio
            For Each lineaAlbaran As DataRow In docAlbaran.dtLineas.Rows
                Select Case lineaAlbaran.RowState
                    Case DataRowState.Modified
                        If lineaAlbaran("Importe") <> lineaAlbaran("Importe", DataRowVersion.Original) Then
                            f.Clear()
                            f.Add(New NumberFilterItem("IDLineaAlbaranHija", lineaAlbaran("IDLineaAlbaran")))
                            Dim dtPrecios As DataTable = ACP.Filter(f)
                            If Not dtPrecios Is Nothing AndAlso dtPrecios.Rows.Count > 0 Then
                                For Each lineaGasto As DataRow In dtPrecios.Rows
                                    If Nz(lineaGasto("Porcentaje"), 0) > 0 Then lineaGasto("Importe") = lineaAlbaran("Importe") * (lineaGasto("Porcentaje") / 100)
                                    lineaGasto("Importe") = xRound(lineaGasto("Importe"), docAlbaran.Moneda.NDecimalesImporte)
                                    lineaGasto("ImporteA") = xRound(lineaGasto("Importe") * docAlbaran.Moneda.CambioA, docAlbaran.Moneda.NDecimalesImporte)
                                    lineaGasto("ImporteB") = xRound(lineaGasto("Importe") * docAlbaran.Moneda.CambioB, docAlbaran.Moneda.NDecimalesImporte)
                                Next
                            End If
                            ACP.Update(dtPrecios)
                            'Dim ctx As New ContextDocumentLineas(docAlbaran, lineaAlbaran)
                            'ProcessServer.ExecuteTask(Of ContextDocumentLineas)(AddressOf ProcesoAlbaranCompra.CorregirMovimiento, ctx, services)
                        End If
                End Select
            Next
        End If

        'If Not docAlbaran.dtPrecios Is Nothing AndAlso docAlbaran.dtPrecios.Rows.Count > 0 Then
        '    For Each lineaGasto As DataRow In docAlbaran.dtPrecios.Select(Nothing, Nothing, DataViewRowState.Added)
        '        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AlbaranCompraPrecio.AsignarIdentificador, lineaGasto, services)
        '    Next
        'End If
    End Sub

    <Task()> Public Shared Sub AsignarAlbaranCompraLotes(ByVal docAlbaran As DocumentoAlbaranCompra, ByVal services As ServiceProvider)
        If Not docAlbaran.dtLineas Is Nothing Then
            Dim lineas As DataTable = docAlbaran.dtLineas
            If lineas Is Nothing Then
                Dim oACP As New AlbaranCompraLinea
                lineas = oACP.AddNew
                docAlbaran.Add(GetType(AlbaranCompraLinea).Name, lineas)
            End If

            Dim lotes As DataTable = docAlbaran.dtLote
            If lotes Is Nothing Then
                Dim oACLo As New AlbaranCompraLote
                lotes = oACLo.AddNew
                docAlbaran.Add(GetType(AlbaranCompraLote).Name, lotes)
            End If

            For Each drACL As DataRow In docAlbaran.dtLineas.Rows
                Dim alblin As AlbLinPedidoCompra = Nothing
                For i As Integer = 0 To CType(docAlbaran.Cabecera, AlbCabPedidoCompra).Lineas.Length - 1
                    If drACL("IDLineaPedido") = CType(docAlbaran.Cabecera, AlbCabPedidoCompra).Lineas(i).IDLineaPedido Then
                        alblin = CType(docAlbaran.Cabecera, AlbCabPedidoCompra).Lineas(i)
                        Exit For
                    End If
                Next

                If Not alblin Is Nothing AndAlso Not alblin.Lotes Is Nothing AndAlso alblin.Lotes.Rows.Count > 0 Then
                    Dim SegundaUnidad As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, drACL("IDArticulo"), services)
                    For Each lt As DataRow In alblin.Lotes.Rows
                        If lt(_AAL.Cantidad) <> 0 Then
                            Dim lote As DataRow = lotes.NewRow

                            lote(_ACLT.IDLineaLote) = AdminData.GetAutoNumeric
                            lote(_ACLT.IDLineaAlbaran) = drACL(_ACL.IDLineaAlbaran)
                            lote(_ACLT.Lote) = lt(_AAL.Lote)
                            lote(_ACLT.Ubicacion) = lt(_AAL.Ubicacion)
                            lote(_ACLT.QInterna) = lt(_AAL.Cantidad)
                            lote(_ACLT.SeriePrecinta) = lt(_AAL.SeriePrecinta)
                            lote(_ACLT.NDesdePrecinta) = lt(_AAL.NDesdePrecinta)
                            lote(_ACLT.NHastaPrecinta) = lt(_AAL.NHastaPrecinta)
                            If SegundaUnidad Then lote(_ACLT.QInterna2) = lt(_AAL.Cantidad2)
                            'TODO VER como rellenar observaciones (añadir campo en AlbaranLineaLote)
                            'lote(_ACLT.Observaciones) = lt(_AAL.Observaciones)
                            'lote(_ACLT.FechaCaducidad) = lt(_AAL.FechaCaducidad)
                            lotes.Rows.Add(lote.ItemArray)
                        End If
                    Next
                End If

            Next

        End If
    End Sub


#End Region

#End Region

    <Task()> Public Shared Sub GrabarDocumento(ByVal Doc As DocumentoAlbaranCompra, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DocumentoAlbaranCompra)(AddressOf Business.General.Comunes.BeginTransaction, Nothing, services)
        'ProcessServer.ExecuteTask(Of DocumentoAlbaranCompra)(AddressOf ValidarDocumento, Doc, services)
        ProcessServer.ExecuteTask(Of DocumentoAlbaranCompra)(AddressOf Business.General.Comunes.UpdateDocument, Doc, services)
        ProcessServer.ExecuteTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.ValoracionSuministro, Doc, services)
        ' ProcessServer.ExecuteTask(Of DataTable)(AddressOf Business.General.Comunes.UpdateEntityDataTable, Doc.dtValoracion, services)
        ProcessServer.ExecuteTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.ActualizarPedidoDesdeAlbaran, Doc, services)
        ProcessServer.ExecuteTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.ActualizarProgramaDesdeAlbaran, Doc, services)
        ProcessServer.ExecuteTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.ActualizarObras, Doc, services)
        For Each linea As DataRow In Doc.dtLineas.Rows
            Dim StData As New DataPrepararArtUltCompra(Doc, linea)
            ProcessServer.ExecuteTask(Of DataPrepararArtUltCompra)(AddressOf PrepararArticuloUltimaCompra, StData, services)
        Next
        'Dim datos As DataCorregirMovimientosStockPrecios
        'datos.Doc = Doc
        'datos.Precios = dtPrecios
        'ProcessServer.ExecuteTask(Of DataCorregirMovimientosStockPrecios)(AddressOf CorregirMovimientosStockPrecios, datos, services)

        'TODO Produccion
        'Dim OrdenRuta(-1) As DataTable
        'OrdenRuta = PrepararOrdenRuta(Doc.dtACL)
        'AdminData.SetData(OrdenRuta)
        'TODO PENDIENTE
        'Dim DtProgCL As DataTable = PrepararProgramaCompraLinea(LineasPedido)
        'If Not DtProgCL Is Nothing AndAlso DtProgCL.Rows.Count > 0 Then
        '    AdminData.SetData(DtProgCL)
        'End If
        ProcessServer.ExecuteTask(Of Boolean)(AddressOf Business.General.Comunes.CommitTransaction, True, services)
    End Sub

    <Serializable()> _
    Public Class DataCorregirMovimientosStockPrecios
        Public Doc As DocumentoAlbaranCompra
        Public Precios As DataTable
    End Class
    <Task()> Public Shared Sub CorregirMovimientosStockPrecios(ByVal data As DataCorregirMovimientosStockPrecios, ByVal services As ServiceProvider)
        Dim ACL As New AlbaranCompraLinea
        For Each dr As DataRow In data.Precios.Rows
            Dim drACL As DataRow = ACL.GetItemRow(dr("IDLineaAlbaran"))
            If dr.RowState = DataRowState.Added Then
                Dim ctx As New DataDocRow(data.Doc, drACL)
                ProcessServer.ExecuteTask(Of DataDocRow)(AddressOf ProcesoAlbaranCompra.CorregirMovimiento, ctx, services)
            Else
                If dr("Importe") <> dr("Importe", DataRowVersion.Original) Then
                    Dim ctx As New DataDocRow(data.Doc, drACL)
                    ProcessServer.ExecuteTask(Of DataDocRow)(AddressOf ProcesoAlbaranCompra.CorregirMovimiento, ctx, services)
                End If
            End If
        Next
    End Sub

#End Region

#Region " Analítica "

    <Task()> Public Shared Sub CopiarAnalitica(ByVal Doc As DocumentoAlbaranCompra, ByVal services As ServiceProvider)
        If Doc.HeaderRow("Estado") <> enumaccEstado.accFacturado Then
            Dim AppParams As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
            If Not AppParams.Contabilidad OrElse Not AppParams.Analitica.AplicarAnalitica Then Exit Sub

            Dim IDOrigen(-1) As Object
            Dim f As New Filter
            f.Add(New IsNullFilterItem("IDLineaPedido", False))
            Dim WhereNotNullLineaPedido As String = f.Compose(New AdoFilterComposer)
            For Each linea As DataRow In Doc.dtLineas.Select(WhereNotNullLineaPedido)
                ReDim Preserve IDOrigen(IDOrigen.Length)
                IDOrigen(IDOrigen.Length - 1) = linea("IDLineaPedido")
            Next
            If IDOrigen.Length > 0 Then
                Dim dtAnaliticaOrigen As DataTable = New PedidoCompraAnalitica().Filter(New InListFilterItem("IDLineaPedido", IDOrigen, FilterType.Numeric))
                Dim datosCopia As New NegocioGeneral.DataCopiarAnalitica(dtAnaliticaOrigen, Doc)
                ProcessServer.ExecuteTask(Of NegocioGeneral.DataCopiarAnalitica)(AddressOf NegocioGeneral.CopiarAnalitica, datosCopia, services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CalcularAnalitica(ByVal Doc As DocumentoAlbaranCompra, ByVal services As ServiceProvider)
        If Doc.HeaderRow("Estado") <> enumaccEstado.accFacturado Then
            ProcessServer.ExecuteTask(Of DocumentoAlbaranCompra)(AddressOf NegocioGeneral.CalcularAnalitica, Doc, services)
        End If
    End Sub

#End Region

#Region " Calcular Albarán "

    <Task()> Public Shared Sub CalcularBasesImponibles(ByVal Doc As DocumentoAlbaranCompra, ByVal services As ServiceProvider)
        If Not IsNothing(Doc.dtLineas) Then
            Dim desglose() As DataBaseImponible = ProcessServer.ExecuteTask(Of DocumentoAlbaranCompra, DataBaseImponible())(AddressOf ProcesoComunes.DesglosarImporte, Doc, services)
            Dim datosCalculo As New ProcesoComunes.DataCalculoTotalesCab
            datosCalculo.Doc = Doc
            datosCalculo.BasesImponibles = desglose
            ProcessServer.ExecuteTask(Of ProcesoComunes.DataCalculoTotalesCab)(AddressOf CalcularTotalesCabecera, datosCalculo, services)
        End If
    End Sub

    <Task()> Public Shared Sub CalcularTotalesCabecera(ByVal data As ProcesoComunes.DataCalculoTotalesCab, ByVal services As ServiceProvider)
        Dim BaseImponibleTotal As Double = 0 : Dim ImporteLineas As Double = 0
        Dim ImporteIVATotal As Double = 0 : Dim ImporteRETotal As Double = 0
        If Not IsNothing(data.Doc.HeaderRow) Then
            If Not IsNothing(data.BasesImponibles) AndAlso data.BasesImponibles.Length > 0 Then
                Dim CondsPago As EntityInfoCache(Of CondicionPagoInfo) = services.GetService(Of EntityInfoCache(Of CondicionPagoInfo))()
                Dim CondPagoInfo As CondicionPagoInfo = CondsPago.GetEntity(data.Doc.HeaderRow("IDCondicionPago"))
                If Not IsNothing(CondPagoInfo) Then
                    data.Doc.HeaderRow("DtoProntoPago") = CondPagoInfo.DtoProntoPago
                    data.Doc.HeaderRow("RecFinan") = CondPagoInfo.RecFinan
                End If

                Dim TiposIVA As EntityInfoCache(Of TipoIvaInfo) = services.GetService(Of EntityInfoCache(Of TipoIvaInfo))()
                For Each BI As DataBaseImponible In data.BasesImponibles
                    ImporteLineas = ImporteLineas + BI.BaseImponible
                    If Length(BI.IDTipoIva) > 0 Then
                        Dim factor As Double = 0
                        BaseImponibleTotal = BaseImponibleTotal + BI.BaseImponible
                        ' HistoricoTipoIVA
                        Dim TIVAInfo As TipoIvaInfo = TiposIVA.GetEntity(BI.IDTipoIva, data.Doc.Fecha)

                        If Length(TIVAInfo.IDTipoIVA) > 0 Then
                            '//valor por defecto
                            factor = TIVAInfo.Factor

                            '//Para los ivas especiales que no se repercuten
                            '//If TIVAInfo.SinRepercutir Then factor = TIVAInfo.IVASinRepercutir
                        End If

                        Dim Base As Double = BI.BaseImponible
                        ImporteIVATotal = ImporteIVATotal + Base * factor / 100
                        Dim AppParamsCompra As ParametroCompra = services.GetService(Of ParametroCompra)()
                        If AppParamsCompra.EmpresaConRecargoEquivalencia Then
                            ImporteRETotal = ImporteRETotal + Base * TIVAInfo.IVARE / 100
                        End If
                    End If
                Next
            End If
        End If

        Dim ImpLineasNormales As Double = 0
        If Nz(data.Doc.HeaderRow("RecFinan"), 0) > 0 Then
            Dim ImpLineasEspeciales As Double = 0
            If Not data.Doc.dtLineas Is Nothing AndAlso data.Doc.dtLineas.Rows.Count > 0 Then
                Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                For Each linea As DataRow In data.Doc.dtLineas.Rows
                    Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(linea("IDArticulo"))
                    If ArtInfo.Especial Then
                        ImpLineasEspeciales += Nz(linea("Importe"), 0)
                    End If
                Next
            End If
            ImpLineasNormales = ImporteLineas - ImpLineasEspeciales

            Dim ValAyBImpLin As New ValoresAyB(ImpLineasNormales, data.Doc.IDMoneda, data.Doc.CambioA, data.Doc.CambioB)
            Dim fImpLineasNormales As fImporte = ProcessServer.ExecuteTask(Of ValoresAyB, fImporte)(AddressOf NegocioGeneral.MantenimientoValoresImporteAyB, ValAyBImpLin, services)

            data.Doc.HeaderRow("ImpRecFinan") = xRound(fImpLineasNormales.Importe * data.Doc.HeaderRow("RecFinan") / 100, data.Doc.Moneda.NDecimalesImporte)
            data.Doc.HeaderRow("ImpRecFinanA") = xRound(fImpLineasNormales.ImporteA * data.Doc.HeaderRow("RecFinan") / 100, data.Doc.MonedaA.NDecimalesImporte)
            data.Doc.HeaderRow("ImpRecFinanB") = xRound(fImpLineasNormales.ImporteB * data.Doc.HeaderRow("RecFinan") / 100, data.Doc.MonedaB.NDecimalesImporte)
        End If
        data.Doc.HeaderRow("BaseImponible") = BaseImponibleTotal
        data.Doc.HeaderRow("ImpIVA") = ImporteIVATotal
        data.Doc.HeaderRow("ImpRE") = ImporteRETotal
        data.Doc.HeaderRow("Importe") = ImporteLineas

        Dim ValAyB As New ValoresAyB(New DataRowPropertyAccessor(data.Doc.HeaderRow), data.Doc.IDMoneda, data.Doc.CambioA, data.Doc.CambioB)
        ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)
    End Sub

#End Region

#Region " ValoracionSuministro "

    <Task()> Public Shared Sub ValoracionSuministro(ByVal Doc As DocumentoAlbaranCompra, ByVal services As ServiceProvider)
        If Not Doc.dtLineas Is Nothing AndAlso Doc.dtLineas.Rows.Count > 0 Then
            Dim dtCriterio As DataTable = BusinessHelper.CreateBusinessObject("CriterioValoracion").Filter
            Dim Cal As BusinessHelper = BusinessHelper.CreateBusinessObject(AdminData.GetEntityInfo("AlbaranCompraValoracion"))
            Dim dtValoracion As DataTable = Doc.dtValoracion
            For Each dr As DataRow In Doc.dtLineas.Rows
                ' Inicialización de variables
                Dim blnHayPedido As Boolean = False
                Dim dblQServida As Double = 0
                Dim dblQPedida As Double = 0
                Dim dteFechaPedido As Date = System.DateTime.MinValue
                Dim blnGestionPorLotes As Boolean = False

                '///Obtener Fecha del Albaran y cantidad de la LINEA del albaran
                If ProcessServer.ExecuteTask(Of DataRow, Boolean)(AddressOf CrearValoracion, dr, services) Then
                    If Length(dr("QInterna")) > 0 Then dblQServida = dr("QInterna")

                    '///Obtener Fecha entrega de la LINEA de pedido y la QPedida
                    If Length(dr("IdLineaPedido")) > 0 Then
                        If dr("IdLineaPedido") > 0 Then
                            Dim dtAux As DataTable = New PedidoCompraLinea().Filter(New NumberFilterItem("IDLineaPedido", dr("IdLineaPedido")))
                            If Not dtAux Is Nothing AndAlso dtAux.Rows.Count > 0 Then
                                blnHayPedido = True

                                ' dblQPedida = dtAux.Rows(0)("QPedida")
                                If dr(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclNormal Then
                                    dblQPedida = dtAux.Rows(0)("QInterna") '* dtAux.Rows(0)("Factor")
                                ElseIf dr(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclComponente Then
                                    Dim dvOrigen As DataView = New DataView(Doc.dtLineas, Nothing, _ACL.IDLineaAlbaran, DataViewRowState.CurrentRows)
                                    Dim idx As Integer = dvOrigen.Find(dr(_ACL.IDLineaPadre))
                                    If idx >= 0 AndAlso dvOrigen(idx)(_ACL.QInterna) <> 0 Then
                                        dblQPedida = (dtAux.Rows(0)("QInterna") / dvOrigen(idx)(_ACL.QInterna)) * dr(_ACL.QInterna)
                                    Else
                                        dblQPedida = 0
                                    End If
                                End If

                                If Length(dr("FechaEntregaModificado")) > 0 AndAlso dr("FechaEntregaModificado") <> cnMinDate Then
                                    dteFechaPedido = dr("FechaEntregaModificado")
                                Else
                                    dteFechaPedido = dtAux.Rows(0)("FechaEntrega")
                                End If
                            End If
                        End If
                    Else
                        blnHayPedido = True
                        dblQPedida = dblQServida
                        dteFechaPedido = Doc.Fecha
                    End If

                    If Not dtCriterio Is Nothing AndAlso dtCriterio.Rows.Count > 0 Then
                        Dim CacheCriterioLinea As Hashtable
                        For Each drCriterio As DataRow In dtCriterio.Rows
                            Dim blnHayDesglose As Boolean : Dim dtDesglose As DataTable
                            Dim blnValoresPorDefecto As Boolean = False
                            Dim dblDiferencia As Double = 0
                            If drCriterio("TipoCriterio") = enumCriteriosCalidad.CCCantidad Or drCriterio("TipoCriterio") = enumCriteriosCalidad.ccCantidadPorcentaje Or drCriterio("TipoCriterio") = enumCriteriosCalidad.CCFechas Then
                                If IsNothing(CacheCriterioLinea) Then CacheCriterioLinea = New Hashtable
                                If CacheCriterioLinea.ContainsKey(drCriterio("IDCriterio")) Then
                                    dtDesglose = CacheCriterioLinea(drCriterio("IDCriterio"))
                                Else
                                    dtDesglose = BusinessHelper.CreateBusinessObject("CriterioValoracionLinea").Filter(New StringFilterItem("IDCriterio", drCriterio("IDCriterio")))
                                    CacheCriterioLinea(drCriterio("IDCriterio")) = dtDesglose
                                End If
                                blnHayDesglose = (Not dtDesglose Is Nothing AndAlso dtDesglose.Rows.Count > 0)
                            End If
                            Dim rw As DataRow = Nothing
                            If dr.RowState = DataRowState.Added Then
                                rw = dtValoracion.NewRow
                                rw("IDValoracion") = AdminData.GetAutoNumeric
                                rw("IDLineaAlbaran") = dr("IDLineaAlbaran")
                                rw("IDCriterio") = drCriterio("IDCriterio")
                                rw("IDDemerito") = drCriterio("IDDemeritoCorrecto")
                                rw("Correcto") = True
                                dtValoracion.Rows.Add(rw)

                            ElseIf dr.RowState = DataRowState.Modified OrElse dr.RowState = DataRowState.Unchanged Then
                                ' Modificando línea de albarán, o bien, para el caso en el que se haya modificado
                                ' algo de la cabecera del albarán, o del Pedido de Compra.
                                'FwnCalidad = New BusinessHelper("AlbaranCompraValoracion")
                                Dim objFilter As New Filter
                                objFilter.Add(New NumberFilterItem("IDLineaAlbaran", dr("IDLineaAlbaran")))
                                objFilter.Add(New StringFilterItem("IDCriterio", drCriterio("IDCriterio")))

                                dtValoracion.DefaultView.RowFilter = objFilter.Compose(New AdoFilterComposer)
                                If dtValoracion.DefaultView.Count > 0 Then rw = dtValoracion.DefaultView(0).Row
                            End If

                            Select Case CType(drCriterio("TipoCriterio"), enumCriteriosCalidad)
                                '///Criterios que SI tienen desglose
                                Case enumCriteriosCalidad.CCCantidad, enumCriteriosCalidad.ccCantidadPorcentaje, enumCriteriosCalidad.CCFechas
                                    blnValoresPorDefecto = True
                                    If blnHayPedido Then
                                        If blnHayDesglose Then
                                            blnValoresPorDefecto = False
                                            Select Case CType(drCriterio("TipoCriterio"), enumCriteriosCalidad)
                                                Case enumCriteriosCalidad.CCCantidad
                                                    dblDiferencia = dblQPedida - dblQServida
                                                Case enumCriteriosCalidad.ccCantidadPorcentaje
                                                    If dblQPedida <> 0 Then
                                                        dblDiferencia = 100 * (1 - (dblQServida / dblQPedida))
                                                    Else
                                                        blnValoresPorDefecto = True
                                                    End If
                                                Case enumCriteriosCalidad.CCFechas
                                                    dblDiferencia = DateDiff(Microsoft.VisualBasic.DateInterval.Day, dteFechaPedido, Doc.Fecha, FirstDayOfWeek.System, FirstWeekOfYear.System)
                                            End Select
                                        End If
                                    End If
                                    '///Criterios que NO tienen desglose
                                Case enumCriteriosCalidad.ccCorrectoIncorrecto, enumCriteriosCalidad.ccControlCalidad, enumCriteriosCalidad.ccReactividad
                                    blnValoresPorDefecto = True
                            End Select

                            If Not dtValoracion Is Nothing AndAlso dtValoracion.Rows.Count > 0 Then
                                If Not rw Is Nothing Then
                                    If Not blnValoresPorDefecto Then
                                        rw("Diferencia") = dblDiferencia
                                        If Not AreEquals(rw("Diferencia"), 0) Then
                                            rw("IDDemerito") = ProcessServer.ExecuteTask(Of DataBuscarEnIntervalo, Double)(AddressOf BuscarEnIntervalo, New DataBuscarEnIntervalo(rw("Diferencia"), dtDesglose), services)
                                            rw("Correcto") = (rw("IDDemerito") = drCriterio("IDDemeritoCorrecto"))
                                        Else
                                            rw("IDDemerito") = drCriterio("IDDemeritoCorrecto")
                                            rw("Correcto") = True
                                        End If
                                    End If
                                End If
                            End If

                            dtValoracion.DefaultView.RowFilter = Nothing
                        Next
                    End If
                End If
            Next
        End If
    End Sub

    <Task()> Public Shared Function CrearValoracion(ByVal data As DataRow, ByVal services As ServiceProvider) As Boolean
        If (data("TipoLineaAlbaran") = enumaclTipoLineaAlbaran.aclComponente And Length(data("IDOrdenRuta")) <> 0) Then
            Return False
        End If
        If (data("TipoLineaAlbaran") = enumaclTipoLineaAlbaran.aclKit) Then
            Return False
        End If
        Return True
    End Function

    <Serializable()> _
    Public Class DataBuscarEnIntervalo
        Public Valor As Double
        Public Intervalos As DataTable

        Public Sub New(ByVal Valor As Double, ByVal Intervalos As DataTable)
            Me.Valor = Valor
            Me.Intervalos = Intervalos
        End Sub
    End Class
    <Task()> Public Shared Function BuscarEnIntervalo(ByVal data As DataBuscarEnIntervalo, ByVal services As ServiceProvider) As Double
        If Not data.Intervalos Is Nothing AndAlso data.Intervalos.Rows.Count > 0 Then
            Dim DvInter As New DataView(data.Intervalos)
            DvInter.Sort = "Hasta DESC"
            If data.Valor >= DvInter(0)("Hasta") Then
                Return DvInter(0)("IDDemerito")
            Else
                Dim dblUmbralSup As Double
                Dim dblValorUmbralSup As Double
                dblUmbralSup = DvInter(0)("Hasta")
                dblValorUmbralSup = DvInter(0)("IDDemerito")
                For Each Dv As DataRowView In DvInter
                    If data.Valor > Dv("Hasta") Then
                        If data.Valor <= dblUmbralSup Then
                            Return dblValorUmbralSup
                        End If
                    Else
                        dblUmbralSup = Dv("Hasta")
                        dblValorUmbralSup = Dv("IDDemerito")
                    End If
                Next
            End If
            Return DvInter(DvInter.Count - 1)("IDDemerito")
        End If
    End Function

#End Region

#Region " PrepararArticuloUltimaCompra "

    <Serializable()> _
     Public Class DataPrepararArtUltCompra
        Public DrLinea As DataRow
        Public Doc As DocumentoAlbaranCompra

        Public Sub New()
        End Sub
        Public Sub New(ByVal Doc As DocumentoAlbaranCompra, ByVal DrLinea As DataRow)
            Me.Doc = Doc
            Me.DrLinea = DrLinea
        End Sub
    End Class

    <Task()> Public Shared Sub PrepararArticulosUltimaCompra(ByVal Doc As DocumentoAlbaranCompra, ByVal services As ServiceProvider)
        If Not Doc Is Nothing AndAlso Not Doc.dtLineas Is Nothing AndAlso Doc.dtLineas.Rows.Count > 0 Then
            For Each linea As DataRow In Doc.dtLineas.Rows
                Dim StData As New DataPrepararArtUltCompra(Doc, linea)
                ProcessServer.ExecuteTask(Of DataPrepararArtUltCompra)(AddressOf PrepararArticuloUltimaCompra, StData, services)
            Next
        End If
    End Sub

    <Task()> Public Shared Sub PrepararArticuloUltimaCompra(ByVal data As DataPrepararArtUltCompra, ByVal services As ServiceProvider)
        Dim rwArt As DataRow = New Articulo().GetItemRow(data.DrLinea("IDArticulo"))
        Dim dtAlbaUltimaFecha As DataTable = New BE.DataEngine().Filter("vAlbaranCompraFecha", New FilterItem("IDArticulo", rwArt("IDArticulo")), "TOP 1 IDAlbaran,FechaAlbaran,TipoLineaAlbaran,IDOrdenRuta,IDProveedor,QInterna,ImporteA,ImporteB", "FechaAlbaran DESC, IDAlbaran DESC")
        Dim BlnEstUltimo As Boolean = False
        If Not IsNothing(dtAlbaUltimaFecha) AndAlso dtAlbaUltimaFecha.Rows.Count Then
            If dtAlbaUltimaFecha.Rows(0)("IDAlbaran") = data.Doc.HeaderRow("IDAlbaran") Then
                BlnEstUltimo = False
            Else
                If dtAlbaUltimaFecha.Rows(0)("FechaAlbaran") > data.Doc.HeaderRow("FechaAlbaran") Then
                    BlnEstUltimo = True
                Else : BlnEstUltimo = False
                End If
            End If
        Else : BlnEstUltimo = False
        End If
        Select Case BlnEstUltimo
            Case True
                If dtAlbaUltimaFecha.Rows(0)("TipoLineaAlbaran") = enumaclTipoLineaAlbaran.aclNormal OrElse _
                  (dtAlbaUltimaFecha.Rows(0)("TipoLineaAlbaran") = enumaclTipoLineaAlbaran.aclSubcontratacion AndAlso _
                  Length(dtAlbaUltimaFecha.Rows(0)("IDOrdenRuta")) > 0) OrElse _
                  dtAlbaUltimaFecha.Rows(0)("TipoLineaAlbaran") = enumaclTipoLineaAlbaran.aclSubcontratacion OrElse _
                  dtAlbaUltimaFecha.Rows(0)("TipoLineaAlbaran") = enumaclTipoLineaAlbaran.aclKit Then
                    If dtAlbaUltimaFecha.Rows(0)("QInterna") > 0 Then
                        rwArt("FechaUltimaCompra") = dtAlbaUltimaFecha.Rows(0)("FechaAlbaran")
                        rwArt("IdProveedorUltimaCompra") = dtAlbaUltimaFecha.Rows(0)("IDProveedor")
                        rwArt("PrecioUltimaCompraA") = dtAlbaUltimaFecha.Rows(0)("ImporteA") / dtAlbaUltimaFecha.Rows(0)("QInterna")
                        rwArt("PrecioUltimaCompraB") = dtAlbaUltimaFecha.Rows(0)("ImporteB") / dtAlbaUltimaFecha.Rows(0)("QInterna")
                        BE.BusinessHelper.UpdateTable(rwArt.Table)
                    End If
                End If
            Case False
                If data.DrLinea("TipoLineaAlbaran") = enumaclTipoLineaAlbaran.aclNormal OrElse _
                  (data.DrLinea("TipoLineaAlbaran") = enumaclTipoLineaAlbaran.aclSubcontratacion AndAlso _
                  Length(data.DrLinea("IDOrdenRuta")) > 0) OrElse _
                  data.DrLinea("TipoLineaAlbaran") = enumaclTipoLineaAlbaran.aclSubcontratacion OrElse _
                  data.DrLinea("TipoLineaAlbaran") = enumaclTipoLineaAlbaran.aclKit Then
                    If data.DrLinea("QInterna") > 0 Then
                        rwArt("FechaUltimaCompra") = data.Doc.HeaderRow("FechaAlbaran")
                        rwArt("IdProveedorUltimaCompra") = data.Doc.HeaderRow("IDProveedor")
                        rwArt("PrecioUltimaCompraA") = data.DrLinea("ImporteA") / data.DrLinea("QInterna")
                        rwArt("PrecioUltimaCompraB") = data.DrLinea("ImporteB") / data.DrLinea("QInterna")
                        BE.BusinessHelper.UpdateTable(rwArt.Table)
                    End If
                End If
        End Select
    End Sub


#End Region

#Region " Actualización Estados de Albaranes "

    <Task()> Public Shared Sub ActualizarEstadoLineas(ByVal data As DocumentoAlbaranCompra, ByVal services As ServiceProvider)
        Dim f As New Filter
        f.Add(New NumberFilterItem("TipoLineaAlbaran", FilterOperator.NotEqual, enumaclTipoLineaAlbaran.aclComponente))
        Dim WhereNoComponentes As String = f.Compose(New AdoFilterComposer)
        For Each linea As DataRow In data.dtLineas.Select(WhereNoComponentes)
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarEstadoLinea, linea, services)

            '//Actualizamos el EstadoFactura a sus lineas hijas (componentes)
            If linea("TipoLineaAlbaran") = enumaclTipoLineaAlbaran.aclKit OrElse linea("TipoLineaAlbaran") = enumaclTipoLineaAlbaran.aclSubcontratacion Then
                Dim fLineasHijas As New Filter
                fLineasHijas.Add(New NumberFilterItem("IDLineaPadre", linea("IDLineaAlbaran")))
                fLineasHijas.Add(New NumberFilterItem("TipoLineaAlbaran", enumaclTipoLineaAlbaran.aclComponente))
                Dim WhereLineasHijas As String = fLineasHijas.Compose(New AdoFilterComposer)
                For Each componente As DataRow In data.dtLineas.Select(WhereLineasHijas)
                    componente("EstadoFactura") = linea("EstadoFactura")
                Next
            End If
        Next
    End Sub

    <Task()> Public Shared Sub AsignarEstadoLinea(ByVal linea As DataRow, ByVal services As ServiceProvider)
        Dim QFacturada As Double = Nz(linea("QFacturada"), 0)
        Dim QServida As Double = Nz(linea("QServida"), 0)

        If QFacturada = 0 Then
            linea("EstadoFactura") = enumaclEstadoFactura.aclNoFacturado
        Else
            If Math.Abs(QFacturada) >= Math.Abs(QServida) Then
                linea("EstadoFactura") = enumaclEstadoFactura.aclFacturado
            ElseIf Math.Abs(QFacturada) < Math.Abs(QServida) Then
                linea("EstadoFactura") = enumaclEstadoFactura.aclParcFacturado
            End If
        End If

    End Sub

    <Task()> Public Shared Sub ActualizarEstadoAlbaran(ByVal Doc As DocumentoAlbaranCompra, ByVal services As ServiceProvider)
        If Doc Is Nothing Then Exit Sub
        If Not IsNothing(Doc.dtLineas) Then
            Dim EstadoFacturaAnt As Integer = -1
            Dim NumEstados As Integer
            For Each lineaAlbaran As DataRow In Doc.dtLineas.Select(Nothing, "EstadoFactura")
                If EstadoFacturaAnt = -1 Then
                    EstadoFacturaAnt = lineaAlbaran("EstadoFactura")
                    NumEstados = 1
                End If
                If EstadoFacturaAnt <> lineaAlbaran("EstadoFactura") Then
                    NumEstados = NumEstados + 1
                    Exit For
                End If
            Next

            If NumEstados = 0 Then
                Doc.HeaderRow("Estado") = enumaccEstado.accNoFacturado
            ElseIf NumEstados = 1 Then
                Doc.HeaderRow("Estado") = Doc.dtLineas.Rows(0)("EstadoFactura")
            Else
                Doc.HeaderRow("Estado") = enumaccEstado.accParcFacturado
            End If
        End If
    End Sub

    <Task()> Public Shared Sub GestionArticulosKit(ByVal Doc As DocumentoAlbaranCompra, ByVal services As ServiceProvider)
        '//Gestion de articulo Kit, subcontratacion o viruta                            
        If Doc Is Nothing Then Exit Sub
        If Not Doc.dtLineas Is Nothing Then
            Dim Componentes As DataTable
            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
            For Each lineaAlbaran As DataRow In Doc.dtLineas.Select
                Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(lineaAlbaran("IDArticulo"))
                If ArtInfo.KitVenta Then
                    Select Case lineaAlbaran.RowState
                        Case DataRowState.Added
                            If IsDBNull(lineaAlbaran(_ACL.IdLineaContratoSub)) Then
                                'No es viruta
                                If Not (lineaAlbaran(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclComponente) Then
                                    Componentes = ProcessServer.ExecuteTask(Of DataRow, DataTable)(AddressOf ComponentesDePrimerNivelDeSubcontratacion, lineaAlbaran, services)
                                    If Not Componentes Is Nothing AndAlso Componentes.Rows.Count > 0 Then
                                        lineaAlbaran(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclSubcontratacion
                                    Else
                                        '//Comprobar si articulo es Kit
                                        Componentes = ProcessServer.ExecuteTask(Of DataRow, DataTable)(AddressOf ComponentesDePrimerNivel, lineaAlbaran, services)
                                        If Not Componentes Is Nothing AndAlso Componentes.Rows.Count > 0 Then
                                            lineaAlbaran(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclKit
                                        Else
                                            lineaAlbaran(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclNormal
                                        End If
                                    End If
                                End If
                            End If
                        Case DataRowState.Modified
                            If (lineaAlbaran(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclKit) OrElse _
                                (lineaAlbaran(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclSubcontratacion) Then
                                Dim datos As New DataDocRow(Doc, lineaAlbaran)
                                ProcessServer.ExecuteTask(Of DataDocRow)(AddressOf ActualizarComponentes, datos, services)
                            End If
                    End Select
                    If Not Componentes Is Nothing Then
                        For Each componente As DataRow In Componentes.Select
                            Doc.dtLineas.ImportRow(componente)
                        Next
                    End If
                ElseIf (lineaAlbaran.RowState = DataRowState.Modified) AndAlso (lineaAlbaran(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclSubcontratacion) Then
                    Dim datos As New DataDocRow(Doc, lineaAlbaran)
                    ProcessServer.ExecuteTask(Of DataDocRow)(AddressOf ActualizarComponentes, datos, services)
                End If
            Next
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarComponentes(ByVal data As DataDocRow, ByVal services As ServiceProvider)
        Dim f As New Filter
        f.Add(New NumberFilterItem(_ACL.IDAlbaran, data.Row(_ACL.IDAlbaran)))
        f.Add(New NumberFilterItem(_ACL.IDLineaPadre, data.Row(_ACL.IDLineaAlbaran)))
        f.Add(New NumberFilterItem(_ACL.TipoLineaAlbaran, enumaclTipoLineaAlbaran.aclComponente))
        Dim WhereComponentes As String = f.Compose(New AdoFilterComposer)
        Dim Componentes() As DataRow = CType(data.Doc, DocumentCabLin).dtLineas.Select(WhereComponentes)
        If Not Componentes Is Nothing AndAlso Componentes.Length > 0 Then
            If Nz(data.Row(_ACL.QInterna, DataRowVersion.Original), 0) <> 0 Then
                Dim factorVariacion As Double = data.Row(_ACL.QInterna) / data.Row(_ACL.QInterna, DataRowVersion.Original)
                For Each componente As DataRow In Componentes
                    componente(_ACL.QServida) = componente(_ACL.QServida) * factorVariacion
                    componente(_ACL.QInterna) = componente(_ACL.QInterna) * factorVariacion
                Next
            End If
        End If
    End Sub

#End Region

#Region " Actualización de Pedidos "

    <Task()> Public Shared Sub ActualizarPedidoDesdeAlbaran(ByVal Doc As DocumentoAlbaranCompra, ByVal services As ServiceProvider)
        If Doc Is Nothing OrElse Doc.dtLineas Is Nothing Then Exit Sub
        For Each lineaAlbaran As DataRow In Doc.dtLineas.Select(Nothing, "IDAlbaran,IDLineaAlbaran")
            'If Length(lineaAlbaran("Lote")) > 0 Then
            ProcessServer.ExecuteTask(Of Object)(AddressOf ActualizarLineaPedido, lineaAlbaran, services)
            'End If

        Next
        ProcessServer.ExecuteTask(Of Object)(AddressOf GrabarPedidos, Nothing, services)
    End Sub

    <Task()> Public Shared Sub ActualizarLineaPedido(ByVal lineaAlbaran As DataRow, ByVal services As ServiceProvider)
        If Length(lineaAlbaran("IDPedido")) > 0 AndAlso Length(lineaAlbaran("IDLineaPedido")) > 0 Then
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ActualizarQServidaLineaPedido, lineaAlbaran, services)
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarQServidaLineaPedido(ByVal lineaAlbaran As DataRow, ByVal services As ServiceProvider)
        If lineaAlbaran.RowState <> DataRowState.Modified OrElse lineaAlbaran("QServida") <> lineaAlbaran("QServida", DataRowVersion.Original) _
           OrElse lineaAlbaran("QRechazada") <> lineaAlbaran("QRechazada", DataRowVersion.Original) Then
            Dim Albaranes As DocumentInfoCache(Of DocumentoAlbaranCompra) = services.GetService(Of DocumentInfoCache(Of DocumentoAlbaranCompra))()
            Dim DocAlb As DocumentoAlbaranCompra = Albaranes.GetDocument(lineaAlbaran("IDAlbaran"))
            Dim Pedidos As DocumentInfoCache(Of DocumentoPedidoCompra) = services.GetService(Of DocumentInfoCache(Of DocumentoPedidoCompra))()
            Dim DocPed As DocumentoPedidoCompra = Pedidos.GetDocument(lineaAlbaran("IDPedido"))
            Dim OriginalQServida As Double
            Dim ProposedQServida As Double = Nz(lineaAlbaran("QServida"), 0) + Nz(lineaAlbaran("QRechazada"), 0)
            If lineaAlbaran.RowState = DataRowState.Modified Then
                OriginalQServida = lineaAlbaran("QServida", DataRowVersion.Original) + Nz(lineaAlbaran("QRechazada", DataRowVersion.Original), 0)
            End If

            DocPed.SetQServida(lineaAlbaran("IDLineaPedido"), ProposedQServida - OriginalQServida, services)
        End If
    End Sub

    <Task()> Public Shared Sub GrabarPedidos(ByVal data As Object, ByVal services As ServiceProvider)
        AdminData.BeginTx()

        Dim Pedidos As DocumentInfoCache(Of DocumentoPedidoCompra) = services.GetService(Of DocumentInfoCache(Of DocumentoPedidoCompra))()

        For Each key As Integer In Pedidos.Keys
            Dim DocPed As DocumentoPedidoCompra = Pedidos.GetDocument(key)
            DocPed.SetData()
        Next

    End Sub

#End Region

#Region " Actualización de Programas "

    Public Class DataActualizarProgramaLinea
        Public IDLineaPrograma As Integer
        Public LineaAlbaran As DataRow
        Public DeletingRow As Boolean


        Public Sub New(ByVal IDLineaPrograma As Integer, ByVal LineaAlbaran As DataRow, Optional ByVal DeletingRow As Boolean = False)
            Me.IDLineaPrograma = IDLineaPrograma
            Me.LineaAlbaran = LineaAlbaran
            Me.DeletingRow = DeletingRow
        End Sub
    End Class

    <Task()> Public Shared Sub ActualizarProgramaDesdeAlbaran(ByVal Doc As DocumentoAlbaranCompra, ByVal services As ServiceProvider)
        If Doc Is Nothing OrElse Doc.dtLineas Is Nothing Then Exit Sub
        For Each lineaAlbaran As DataRow In Doc.dtLineas.Select(Nothing, "IDAlbaran,IDLineaAlbaran")
            ProcessServer.ExecuteTask(Of Object)(AddressOf ActualizarLineaPrograma, lineaAlbaran, services)
        Next
    End Sub

    <Task()> Public Shared Sub ActualizarLineaPrograma(ByVal lineaAlbaran As DataRow, ByVal services As ServiceProvider)
        If Length(lineaAlbaran("IDPedido")) > 0 AndAlso Length(lineaAlbaran("IDLineaPedido")) > 0 Then
            If lineaAlbaran.RowState <> DataRowState.Modified OrElse lineaAlbaran("QServida") <> lineaAlbaran("QServida", DataRowVersion.Original) Then
                Dim pl As New PedidoCompraLinea
                Dim DtPedido As DataTable = pl.SelOnPrimaryKey(lineaAlbaran("IDLineaPedido"))
                For Each lineaPedido As DataRow In DtPedido.Select
                    If Length(lineaPedido("IDPrograma")) > 0 AndAlso Length(lineaPedido("IDLineaPrograma")) > 0 Then
                        Dim datosActProg As New DataActualizarProgramaLinea(lineaPedido("IDLineaPrograma"), lineaAlbaran)
                        ProcessServer.ExecuteTask(Of DataActualizarProgramaLinea)(AddressOf ProcesoAlbaranCompra.ActualizarProgramaLinea, datosActProg, services)
                    End If
                Next
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarProgramaLinea(ByVal data As DataActualizarProgramaLinea, ByVal services As ServiceProvider)
        Dim CambioQServida As Boolean

        If data.LineaAlbaran.RowState = DataRowState.Modified Then
            CambioQServida = (data.LineaAlbaran("QServida") <> data.LineaAlbaran("QServida", DataRowVersion.Original))
        End If
        If data.LineaAlbaran.RowState = DataRowState.Added OrElse CambioQServida OrElse data.DeletingRow Then
            Dim pl As New ProgramaCompraLinea
            Dim Programa As DataTable = pl.SelOnPrimaryKey(data.IDLineaPrograma)
            If Not IsNothing(Programa) AndAlso Programa.Rows.Count > 0 Then
                If data.DeletingRow Then
                    Programa.Rows(0)("QServida") -= data.LineaAlbaran("QServida", DataRowVersion.Original)
                Else
                    Dim dblQModificada As Integer
                    If data.LineaAlbaran.RowState = DataRowState.Modified Then
                        dblQModificada = data.LineaAlbaran("QServida", DataRowVersion.Original)
                    End If
                    Programa.Rows(0)("QServida") = Nz(Programa.Rows(0)("QServida"), 0) + (data.LineaAlbaran("QServida") - dblQModificada)
                End If
                BusinessHelper.UpdateTable(Programa)
            End If
        End If
    End Sub

#End Region



#Region "Actualización de Bodegas"

    <Task()> Public Shared Sub ActualizarDAAARCBodegas(ByVal data As DocumentoAlbaranCompra, ByVal services As ServiceProvider)
        If New Parametro().GestionBodegas Then
            If data.HeaderRow.RowState = DataRowState.Modified AndAlso _
            (Nz(data.HeaderRow("NDaa"), String.Empty) <> Nz(data.HeaderRow("NDaa", DataRowVersion.Original), String.Empty) OrElse _
             Nz(data.HeaderRow("AadReferenceCode"), String.Empty) <> Nz(data.HeaderRow("AadReferenceCode", DataRowVersion.Original), String.Empty)) Then
                For Each DrLinea As DataRow In data.dtLineas.Select()
                    Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(GetType(EntityInfoCache(Of ArticuloInfo)))
                    Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(DrLinea(_ACL.IDArticulo))
                    Dim DtAlbCompraLote As DataTable = New AlbaranCompraLote().Filter(New FilterItem("IDLineaAlbaran", FilterOperator.Equal, DrLinea(_ACL.IDLineaAlbaran)))
                    If Not DtAlbCompraLote Is Nothing AndAlso DtAlbCompraLote.Rows.Count > 0 Then
                        Dim datIStock As New ProcesoStocks.DataCreateIStockClass(ArtInfo.EnsambladoStock, ArtInfo.ClaseStock)
                        Dim IStockClassBdg As IStock = ProcessServer.ExecuteTask(Of ProcesoStocks.DataCreateIStockClass, IStock)(AddressOf ProcesoStocks.CreateIStockClass, datIStock, services)
                        For Each DrLote As DataRow In DtAlbCompraLote.Select
                            If Length(DrLote("NEntrada")) > 0 Then
                                Dim StData As New DataActDAAARCEntVino(DrLote("NEntrada"), Nz(data.HeaderRow("NDaa"), String.Empty), Nz(data.HeaderRow("AadReferenceCode"), String.Empty))
                                IStockClassBdg.ActualizarDAAARCEntradaVino(StData)
                            End If
                        Next
                    End If
                Next
            End If
        End If
    End Sub

#End Region

#Region " Actualización Obras "

    <Task()> Public Shared Sub ActualizarObras(ByVal doc As DocumentoAlbaranCompra, ByVal services As ServiceProvider)
        If Not doc Is Nothing AndAlso Not doc.HeaderRow Is Nothing AndAlso Not doc.dtLineas Is Nothing Then
            For Each drLinea As DataRow In doc.dtLineas.Select
                Dim GeneradoControl As Boolean = ProcessServer.ExecuteTask(Of DataRow, Boolean)(AddressOf ActualizacionControlObras.AlbaranGeneradoControl, drLinea, services)
                If GeneradoControl Then
                    If (drLinea.RowState = DataRowState.Added AndAlso Length(drLinea("IDObra")) > 0) OrElse _
                      (drLinea.RowState = DataRowState.Modified AndAlso (drLinea("ImporteA") <> drLinea("ImporteA", DataRowVersion.Original) _
                      OrElse Nz(drLinea("TipoGastoObra")) <> Nz(drLinea("TipoGastoObra", DataRowVersion.Original)) _
                      OrElse Nz(drLinea("IDObra")) <> Nz(drLinea("IDObra", DataRowVersion.Original)) _
                      OrElse Nz(drLinea("IDTrabajo")) <> Nz(drLinea("IDTrabajo", DataRowVersion.Original)) _
                      OrElse Nz(drLinea("IDLineaPadre")) <> Nz(drLinea("IDLineaPadre", DataRowVersion.Original)))) _
                      OrElse (doc.HeaderRow.RowState = DataRowState.Modified AndAlso doc.HeaderRow("FechaAlbaran") <> doc.HeaderRow("FechaAlbaran", DataRowVersion.Original)) Then
                        Dim info As New ActualizacionControlObras.dataControlObras(drLinea, doc.HeaderRow("FechaAlbaran"), ActualizacionControlObras.enumOrigen.Albaran)
                        Select Case CType(Nz(drLinea("TipoGastoObra"), enumfclTipoGastoObra.enumfclMaterial), enumfclTipoGastoObra)
                            Case enumfclTipoGastoObra.enumfclMaterial
                                ProcessServer.ExecuteTask(Of ActualizacionControlObras.dataControlObras)(AddressOf ActualizacionControlObras.ActualizarObraMaterialControl, info, services)
                            Case enumfclTipoGastoObra.enumfclGastos
                                ProcessServer.ExecuteTask(Of ActualizacionControlObras.dataControlObras)(AddressOf ActualizacionControlObras.ActualizarObraGastoControl, info, services)
                            Case enumfclTipoGastoObra.enumfclVarios
                                ProcessServer.ExecuteTask(Of ActualizacionControlObras.dataControlObras)(AddressOf ActualizacionControlObras.ActualizarObraVariosControl, info, services)
                        End Select
                    End If
                End If
            Next
        End If
    End Sub

#End Region

#Region " Update "

    <Task()> Public Shared Function CrearDocumento(ByVal data As UpdatePackage, ByVal services As ServiceProvider) As DocumentoAlbaranCompra
        Return New DocumentoAlbaranCompra(data)
    End Function

#End Region

#Region " STOCKS "

#Region " Actualizar Stock "


    <Task()> Public Shared Sub DetalleActualizacionStocks(ByVal data As Object, ByVal services As ServiceProvider)
        Dim AppParamsAlb As ParametroAlbaranCompra = services.GetService(Of ParametroAlbaranCompra)()
        If Not AppParamsAlb.ActualizacionAutomaticaStock Then Exit Sub

        Dim alog As AlbaranLogProcess = services.GetService(Of AlbaranLogProcess)()
        System.Runtime.Remoting.Messaging.CallContext.SetData(GetType(AlbaranLogProcess).Name, alog)
    End Sub

    <Task()> Public Shared Sub ActualizacionAutomaticaStock(ByVal Doc As DocumentoAlbaranCompra, ByVal services As ServiceProvider)
        Dim AppParamsAlb As ParametroAlbaranCompra = services.GetService(Of ParametroAlbaranCompra)()
        If Not AppParamsAlb.ActualizacionAutomaticaStock Then Exit Sub

        'comprobar las lineas que se estén actualizando el precio y que contengan números de serie pero que esté dado de baja porque se ha vendido,
        'para solo así actualizar esas lineas su precio en los movimientos y no realizar nada mas

        '//El UpdateDocument hace que se mantengan los estados de los registros. Recuperamos los datos de nuevo.
        Doc = New DocumentoAlbaranCompra(Doc.HeaderRow("IDAlbaran"))
        Dim AppParamsCompra As ParametroCompra = services.GetService(Of ParametroCompra)()
        If (Doc.HeaderRow("IDTipoCompra") = AppParamsCompra.TipoCompraSubcontratacion) AndAlso Not Doc.Cabecera Is Nothing Then
            '//Si es de subcontratación y venimos desde un proceso de generación
            Exit Sub
        End If

        '//El UpdateDocument hace que se mantengan los estados de los registros. Los cambiamos para poder hacer la actualización.
        'For Each key As String In Doc.Keys
        '    Doc.Item(key).AcceptChanges()
        'Next

        '//Terminamos de grabar el AC antes de empezar con la actualización de Stocks
        AdminData.CommitTx(True)
        Dim ActStock As New ProcesoStocks.DataActualizarStockLineas(Doc)
        Dim stockUD() As StockUpdateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarStockLineas, StockUpdateData())(AddressOf ActualizarStockLineas, ActStock, services)
        If Not stockUD Is Nothing AndAlso stockUD.Length > 0 Then
            Dim alog As AlbaranLogProcess = services.GetService(Of AlbaranLogProcess)()
            If Not alog Is Nothing Then
                For Each data As StockUpdateData In stockUD
                    ReDim Preserve alog.StockUpdateData(UBound(alog.StockUpdateData) + 1)
                    alog.StockUpdateData(UBound(alog.StockUpdateData)) = data
                Next
            End If
        End If
    End Sub


    <Task()> Public Shared Function ActualizarStockLineas(ByVal data As ProcesoStocks.DataActualizarStockLineas, ByVal services As ServiceProvider) As StockUpdateData()
        Dim stkUptData As StockUpdateData
        Dim updateDataArray(-1) As StockUpdateData
        Dim ofc As BusinessHelper = BusinessHelper.CreateBusinessObject("OFControl")
        Dim [or] As BusinessHelper = BusinessHelper.CreateBusinessObject("OrdenRuta")
        Dim oe As BusinessHelper = BusinessHelper.CreateBusinessObject("OrdenEstructura")

        Dim OperarioGenerico As String = New Parametro().OperarioGenerico()
        Dim dtLinea As DataTable = data.DocumentoAlbaran.dtLineas.Clone
        'Dim LotesLinea As DataTable
        Dim actStockAlb As New ProcesoStocks.DataActualizarStockAlbaranTx
        Dim f As New Filter(FilterUnionOperator.Or)
        If Not data.IDLineasAlbaran Is Nothing AndAlso data.IDLineasAlbaran.Length > 0 Then
            For Each IDLinea As Integer In data.IDLineasAlbaran
                f.Add(New NumberFilterItem("IDLineaAlbaran", IDLinea))
            Next
        End If
        Dim strWhere As String = String.Empty
        If f.Count > 0 Then
            strWhere = f.Compose(New AdoFilterComposer)
        End If

        Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
        Dim IStockClass As IStockInventarioPermanente
        If AppParams.GestionInventarioPermanente Then
            Dim datIStock As New ProcesoStocks.DataCreateIStockClass(InventariosPermanentes.ENSAMBLADO_INV_PERMANENTE_STOCKS, InventariosPermanentes.CLASE_INV_PERMANENTE_STOCKS)
            IStockClass = ProcessServer.ExecuteTask(Of ProcesoStocks.DataCreateIStockClass, IStockInventarioPermanente)(AddressOf ProcesoStocks.CreateIStockInventarios, datIStock, services)
        End If

        Dim ACP As New AlbaranCompraPrecio
        For Each lineaAlbaran As DataRow In data.DocumentoAlbaran.dtLineas.Select(strWhere)
            AdminData.BeginTx()
            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(GetType(EntityInfoCache(Of ArticuloInfo)))
            Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(lineaAlbaran(_ACL.IDArticulo))
            Dim dtTipo As DataTable = New TipoArticulo().SelOnPrimaryKey(ArtInfo.IDTipo)
            If Length(dtTipo.Rows(0)("EnsambladoStock")) = 0 AndAlso Length(dtTipo.Rows(0)("ClaseStock")) = 0 Then
                If lineaAlbaran(_ACL.EstadoStock) = enumaclEstadoStock.aclNoActualizado Then
                    Dim fLinAlb As New Filter
                    fLinAlb.Add(New NumberFilterItem(_ACLT.IDLineaAlbaran, lineaAlbaran(_ACL.IDLineaAlbaran)))
                    If lineaAlbaran(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclNormal OrElse _
                       lineaAlbaran(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclKit OrElse _
                       lineaAlbaran(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclRealquiler OrElse _
                       (lineaAlbaran(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclSubcontratacion AndAlso IsDBNull(lineaAlbaran(_ACL.IDOrdenRuta))) OrElse _
                       (lineaAlbaran(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclComponente AndAlso IsDBNull(lineaAlbaran(_ACL.IDOrdenRuta))) Then
                        '//Linea NORMAL, o de tipo KIT, o SUBCONTRATACION MANUAL(que NO proviene de una OF)
                        If ArtInfo.GestionStock Then
                            'actStockAlb.IDCliente = data.DocumentoAlbaran.HeaderRow("IDCliente")
                            actStockAlb.IDAlbaran = data.DocumentoAlbaran.HeaderRow("IDAlbaran")
                            actStockAlb.NAlbaran = data.DocumentoAlbaran.HeaderRow("NAlbaran")
                            actStockAlb.FechaAlbaran = data.DocumentoAlbaran.HeaderRow("FechaAlbaran")
                            actStockAlb.NumeroMovimiento = Nz(data.DocumentoAlbaran.HeaderRow("NMovimiento"), 0)
                            actStockAlb.LineaAlbaran = lineaAlbaran
                            actStockAlb.LineasAlbaran = data.DocumentoAlbaran.dtLineas
                            actStockAlb.ImporteExtraA = 0
                            actStockAlb.ImporteExtraB = 0
                            actStockAlb.Circuito = Circuito.Compras

                            actStockAlb.LotesLineaAlbaran = CType(data.DocumentoAlbaran, DocumentoAlbaranCompra).dtLote.Clone
                            Dim Importes As DataTable = ACP.Filter(fLinAlb) 'CType(data.DocumentoAlbaran, DocumentoAlbaranCompra).dtPrecios.Select(fLinAlb.Compose(New AdoFilterComposer))
                            If Not Importes Is Nothing Then
                                For Each importe As DataRow In Importes.Rows
                                    actStockAlb.ImporteExtraA = actStockAlb.ImporteExtraA + Nz(importe("ImporteA"), 0)
                                    actStockAlb.ImporteExtraB = actStockAlb.ImporteExtraB + Nz(importe("ImporteB"), 0)
                                Next
                            End If

                            If ArtInfo.GestionStockPorLotes Then  ''//CON GESTION POR LOTES
                                '//Obtenemos los Lotes de la línea
                                Dim WhereLotesLinea As String = fLinAlb.Compose(New AdoFilterComposer)
                                For Each lineaLote As DataRow In CType(data.DocumentoAlbaran, DocumentoAlbaranCompra).dtLote.Select(WhereLotesLinea)
                                    actStockAlb.LotesLineaAlbaran.ImportRow(lineaLote)
                                Next

                                If actStockAlb.LotesLineaAlbaran.Rows.Count > 0 Then
                                    Dim uda() As StockUpdateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarStockAlbaranTx, StockUpdateData())(AddressOf ProcesoStocks.ActualizarStockAlbaranTx, actStockAlb, services)
                                    If Not uda Is Nothing AndAlso uda.Length > 0 AndAlso uda(0).NumeroMovimiento > 0 AndAlso uda(0).Estado = EstadoStock.Actualizado Then
                                        If Nz(data.DocumentoAlbaran.HeaderRow("NMovimiento"), 0) = 0 Then data.DocumentoAlbaran.HeaderRow("NMovimiento") = uda(0).NumeroMovimiento
                                        stkUptData = uda(0)
                                    End If
                                    ArrayManager.Copy(uda, updateDataArray)
                                Else
                                    ReDim Preserve updateDataArray(updateDataArray.Length)
                                    Dim das As New ProcesoStocks.DataLogActualizarStock("El lote es obligatorio.", lineaAlbaran("IDArticulo"), lineaAlbaran("IDAlmacen"))
                                    updateDataArray(updateDataArray.Length - 1) = ProcessServer.ExecuteTask(Of ProcesoStocks.DataLogActualizarStock, StockUpdateData)(AddressOf ProcesoStocks.LogActualizarStock, das, services)
                                End If
                            Else ''//SIN GESTION POR LOTES
                                Dim uda() As StockUpdateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarStockAlbaranTx, StockUpdateData())(AddressOf ProcesoStocks.ActualizarStockAlbaranTx, actStockAlb, services)
                                If Not uda Is Nothing AndAlso uda.Length > 0 AndAlso uda(0).NumeroMovimiento > 0 AndAlso uda(0).Estado = EstadoStock.Actualizado Then
                                    If Nz(data.DocumentoAlbaran.HeaderRow("NMovimiento"), 0) = 0 Then data.DocumentoAlbaran.HeaderRow("NMovimiento") = uda(0).NumeroMovimiento
                                    stkUptData = uda(0)
                                End If
                                ArrayManager.Copy(uda, updateDataArray)
                            End If
                        End If

                        dtLinea.ImportRow(lineaAlbaran)
                        BE.BusinessHelper.UpdateTable(data.DocumentoAlbaran.HeaderRow.Table)
                        BE.BusinessHelper.UpdateTable(dtLinea)
                        BE.BusinessHelper.UpdateTable(actStockAlb.LotesLineaAlbaran)
                    ElseIf lineaAlbaran(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclSubcontratacion And Nz(lineaAlbaran(_ACL.IDOrdenRuta), 0) <> 0 Then
                        '//Lineas de SUBCONTRATACION QUE PROVIENEN DE UNA ORDEN DE FABRICACION.
                        '//Las lineas de albaran se actualizaran automaticamente, independientemente
                        '//de que la actualizacion del stock se haga correctamente. Los movimientos 
                        '//pendientes de actualizar, si existen, se gestionaran desde el programa
                        '//'Movimientos de Stock asociados a la Orden' (tbOFControlEstructura).

                        Dim ParteTrabajo As DataTable
                        Dim produccionLog As ControlProduccionUpdateData

                        '//parte de trabajo (registro de tbOFControl)
                        If IsNumeric(lineaAlbaran("IDOFControl")) Then
                            ParteTrabajo = ofc.SelOnPrimaryKey(lineaAlbaran("IDOFControl"))
                        Else
                            Dim Cabecera As DataRow = data.DocumentoAlbaran.HeaderRow
                            If Not Cabecera Is Nothing Then
                                Dim FechaActualizacion As Date = Cabecera("FechaAlbaran")

                                Dim operacion As DataRow
                                If IsNumeric(lineaAlbaran("IDOrdenRuta")) Then
                                    operacion = [or].GetItemRow(lineaAlbaran("IDOrdenRuta"))
                                    ParteTrabajo = ofc.AddNewForm()
                                    Dim parte As DataRow = ParteTrabajo.Rows(0)
                                    Dim context As New BusinessData
                                    context("IDAlbaran") = lineaAlbaran("IDAlbaran")
                                    parte = ofc.ApplyBusinessRule("IDOrden", operacion("IDOrden"), parte, context)
                                    parte("FechaInicio") = FechaActualizacion
                                    parte("FechaFin") = FechaActualizacion
                                    parte("IDOperario") = OperarioGenerico
                                    parte("IDOrdenRuta") = operacion("IDOrdenRuta")
                                    parte("Secuencia") = operacion("Secuencia")
                                    parte = ofc.ApplyBusinessRule("Secuencia", parte("Secuencia"), parte)
                                    parte("QBuenaUdProduccion") = lineaAlbaran("QServida")
                                    parte("QRechazadaUdProduccion") = 0 'Nz(lineaAlbaran("QRechazada"), 0)
                                    parte("QDudosaUdProduccion") = 0
                                    parte("FactorProduccion") = lineaAlbaran("Factor")
                                    parte = ofc.ApplyBusinessRule("QBuenaUdProduccion", parte("QBuenaUdProduccion"), parte)
                                    parte = ofc.ApplyBusinessRule("QRechazadaUdProduccion", parte("QRechazadaUdProduccion"), parte)
                                    parte = ofc.ApplyBusinessRule("QDudosaUdProduccion", parte("QDudosaUdProduccion"), parte)

                                    lineaAlbaran("IDOFControl") = parte("IDOFControl")
                                    lineaAlbaran("EstadoStock") = enumaclEstadoStock.aclActualizado
                                End If
                            End If
                        End If

                        '//Obtener las lineas componentes de la linea de subcontratacion actual
                        Dim fLineaPadre As New Filter
                        fLineaPadre.Add(New NumberFilterItem("IDLineaPadre", lineaAlbaran(_ACL.IDLineaAlbaran)))
                        Dim WhereLineaPadre As String = fLineaPadre.Compose(New AdoFilterComposer)
                        Dim componentes() As DataRow = data.DocumentoAlbaran.dtLineas.Select(WhereLineaPadre)
                        If Not componentes Is Nothing AndAlso componentes.Length > 0 Then
                            For Each componente As DataRow In componentes
                                componente("IDOFControl") = lineaAlbaran("IDOFControl")
                                componente("EstadoStock") = enumaclEstadoStock.aclActualizado
                            Next
                        End If

                        produccionLog = CType(ofc, IControlProduccion).ControlProduccion(ParteTrabajo)
                        If Nz(lineaAlbaran("QRechazada")) <> 0 Then
                            Dim ClsOFR As BusinessHelper = BusinessHelper.CreateBusinessObject("OFControlRechazo")
                            Dim ClsAlbRech As BusinessHelper = BusinessHelper.CreateBusinessObject("AlbaranCompraRechazo")
                            Dim DtAlbRech As DataTable = ClsAlbRech.Filter(New FilterItem("IDAlbaranCompraLinea", FilterOperator.Equal, lineaAlbaran("IDLineaAlbaran")))
                            If Not DtAlbRech Is Nothing AndAlso DtAlbRech.Rows.Count > 0 Then
                                Dim DtNew As DataTable = ClsOFR.AddNew()
                                For Each DrRech As DataRow In DtAlbRech.Select
                                    Dim FilRechazo As New Filter
                                    FilRechazo.Add("IDOFControl", FilterOperator.Equal, ParteTrabajo.Rows(0)("IDOFControl"))
                                    FilRechazo.Add("IDCausaRechazo", FilterOperator.Equal, DrRech("IDCausaRechazo"))
                                    FilRechazo.Add("QRechazadaUDProduccion", FilterOperator.Equal, DrRech("QRechazada"))
                                    Dim DtParteRechazo As DataTable = ClsOFR.Filter(FilRechazo)
                                    ClsOFR.Delete(DtParteRechazo)

                                    Dim DrNew As DataRow = DtNew.NewRow
                                    DrNew("IDOFControl") = ParteTrabajo.Rows(0)("IDOFControl")
                                    DrNew("IDCausaRechazo") = DrRech("IDCausaRechazo")
                                    Dim DblParte As Double = 1
                                    If Nz(ParteTrabajo.Rows(0)("FactorProduccion"), 0) = 0 Then
                                        DblParte = 1
                                    Else : DblParte = ParteTrabajo.Rows(0)("FactorProduccion")
                                    End If
                                    DrNew("FactorProduccion") = DblParte
                                    DrNew("QRechazadaUDProduccion") = DrRech("QRechazada")
                                    DrNew("QRechazada") = DrNew("QRechazadaUDProduccion") * DblParte
                                    DtNew.Rows.Add(DrNew)
                                Next
                                ClsOFR.Update(DtNew)
                            End If
                        End If
                        If Not produccionLog Is Nothing Then
                            ArrayManager.Copy(produccionLog.Entradas, updateDataArray)
                            ArrayManager.Copy(produccionLog.Salidas, updateDataArray)
                        End If
                        dtLinea.ImportRow(lineaAlbaran)
                        BE.BusinessHelper.UpdateTable(dtLinea)
                        'AdminData.SetData(componentes)
                    End If
                End If
            Else
                If Not ArtInfo.GestionStock OrElse Not ArtInfo.GestionStockPorLotes Then
                    ApplicationService.GenerateError("El artículo de bodega debe llevar gestión de Stock y gestión por lotes.")
                End If
                If lineaAlbaran(_ACL.EstadoStock) = enumaclEstadoStock.aclNoActualizado Then
                    Dim DtAlbCompraLote As DataTable = New AlbaranCompraLote().Filter(New FilterItem("IDLineaAlbaran", FilterOperator.Equal, lineaAlbaran(_ACL.IDLineaAlbaran)))
                    If Not DtAlbCompraLote Is Nothing AndAlso DtAlbCompraLote.Rows.Count > 0 Then
                        Dim datIStock As New ProcesoStocks.DataCreateIStockClass(ArtInfo.EnsambladoStock, ArtInfo.ClaseStock)
                        Dim IStockClassBdg As IStock = ProcessServer.ExecuteTask(Of ProcesoStocks.DataCreateIStockClass, IStock)(AddressOf ProcesoStocks.CreateIStockClass, datIStock, services)
                        For Each DrLote As DataRow In DtAlbCompraLote.Select
                            Dim Precio As Double = 0

                            If lineaAlbaran(_ACL.QInterna) <> 0 Then
                                Precio = lineaAlbaran(_ACL.ImporteA) / lineaAlbaran(_ACL.QInterna)
                            End If

                            Dim StData As New DataAltEntVino(lineaAlbaran(_ACL.IDArticulo), data.DocumentoAlbaran.HeaderRow("IDProveedor"), Precio, DrLote("QInterna"), DrLote("Lote"), data.DocumentoAlbaran.HeaderRow("FechaAlbaran"), Nz(data.DocumentoAlbaran.HeaderRow("NDaa"), String.Empty), Nz(data.DocumentoAlbaran.HeaderRow("AadReferenceCode"), String.Empty))
                            Dim NEntrada As Integer = IStockClassBdg.AltaEntradaVino(StData)
                            DrLote("NEntrada") = NEntrada
                        Next
                        BusinessHelper.UpdateTable(DtAlbCompraLote)

                        lineaAlbaran("EstadoStock") = enumaclEstadoStock.aclActualizado
                        dtLinea.ImportRow(lineaAlbaran)
                        BE.BusinessHelper.UpdateTable(dtLinea)
                    Else

                    End If
                End If
            End If
            AdminData.CommitTx(True)

            '//Una vez actualizada la linea, vemos si hay que contabilizarla
            If AppParams.GestionInventarioPermanente Then
                If lineaAlbaran(_ACL.EstadoStock) = enumaclEstadoStock.aclActualizado Then
                    If Not IStockClass Is Nothing Then
                        If (lineaAlbaran.RowState = DataRowState.Modified AndAlso _
                           (Nz(lineaAlbaran(_ACL.EstadoStock), -1) <> Nz(lineaAlbaran(_ACL.EstadoStock, DataRowVersion.Original), -1)) OrElse _
                            Nz(lineaAlbaran("Contabilizado"), enumContabilizado.NoContabilizado) = enumContabilizado.NoContabilizado) AndAlso _
                            lineaAlbaran("EstadoFactura") = enumaclEstadoFactura.aclNoFacturado AndAlso _
                            Nz(lineaAlbaran("EstadoFactura"), -1) = Nz(lineaAlbaran("EstadoFactura", DataRowVersion.Original), -1) AndAlso _
                            (Nz(lineaAlbaran("TipoLineaAlbaran"), -1) <> enumaclTipoLineaAlbaran.aclComponente) Then

                            Try
                                IStockClass.SincronizarContaAlbaranCompra(lineaAlbaran("IDLineaAlbaran"), lineaAlbaran("Contabilizado"), services)
                            Catch ex As Exception
                                If Not stkUptData Is Nothing Then
                                    stkUptData.Estado = EstadoStock.NoActualizado
                                    stkUptData.Log = ex.Message
                                    stkUptData.Detalle = ex.Message
                                Else
                                    Dim alog As AlbaranLogProcess = services.GetService(Of AlbaranLogProcess)()
                                    If alog.CreateData Is Nothing Then alog.CreateData = New LogProcess
                                    ReDim Preserve alog.CreateData.Errors(alog.CreateData.Errors.Length)

                                    alog.CreateData.Errors(alog.CreateData.Errors.Length - 1) = New ClassErrors(lineaAlbaran("IDArticulo"), ex.Message)
                                End If
                            End Try
                        End If
                    End If
                End If
            End If
        Next

        Return updateDataArray
    End Function


    '<Task()> Public Shared Function ActualizarStock(ByVal LineasAlbaran As DataTable, ByVal services As ServiceProvider) As StockUpdateData()
    '    If Not IsNothing(LineasAlbaran) Then
    '        For Each drLineaAlbaran As DataRow In LineasAlbaran.Rows
    '            ActualizarModifStock(drLineaAlbaran, services)
    '        Next
    '    End If
    'End Function

    '<Task()> Public Shared Sub ActualizarModifStock(ByVal dr As DataRow, ByVal services As ServiceProvider)
    '    '//Gestion stock
    '    If dr.RowState = DataRowState.Modified Then
    '        If dr(_ACL.QServida, DataRowVersion.Original) <> dr(_ACL.QServida) _
    '        Or dr(_ACL.QInterna, DataRowVersion.Original) <> dr(_ACL.QInterna) _
    '        Or dr(_ACL.ImporteA, DataRowVersion.Original) <> dr(_ACL.ImporteA) _
    '        Or dr(_ACL.ImporteB, DataRowVersion.Original) <> dr(_ACL.ImporteB) Then
    '            If Nz(dr(_ACL.IDOrdenRuta), 0) = 0 Then
    '                Dim Componentes As DataTable
    '                Dim ACL As New AlbaranCompraLinea
    '                If dr(_ACL.QServida, DataRowVersion.Original) <> dr(_ACL.QServida) And (dr(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclKit Or dr(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclSubcontratacion) Then
    '                    Componentes = ACL.ActualizarComponentes(dr)
    '                End If
    '                Dim updateData As StockUpdateData
    '                updateData = ACL.CorregirMovimiento(dr)
    '                If dr(_ACL.EstadoStock) <> EstadoStock.NoActualizado Then
    '                    If Not Componentes Is Nothing AndAlso Componentes.Rows.Count > 0 Then
    '                        ACL.CorregirMovimiento(Componentes)
    '                    End If
    '                Else
    '                    If Not IsNothing(updateData.Log) Then
    '                        Throw New Exception(updateData.Log)
    '                    End If
    '                End If

    '            Else
    '                If dr(_ACL.QServida, DataRowVersion.Original) <> dr(_ACL.QServida) _
    '                Or dr(_ACL.QInterna, DataRowVersion.Original) <> dr(_ACL.QInterna) Then
    '                    If dr(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclSubcontratacion Then
    '                        If Nz(dr(_ACL.IDOFControl), 0) <> 0 Then
    '                            Dim ofc As BusinessHelper = BusinessHelper.CreateBusinessObject("OFControl")
    '                            Dim parteTrabajo As DataTable = ofc.SelOnPrimaryKey(dr(_ACL.IDOFControl))
    '                            If parteTrabajo.Rows.Count > 0 Then
    '                                parteTrabajo.Rows(0)("QBuena") = dr(_ACL.QInterna)
    '                                CType(ofc, IControlProduccion).ControlProduccion(parteTrabajo)
    '                            End If
    '                        End If
    '                    End If
    '                End If
    '            End If
    '        End If
    '    End If
    'End Sub


    'Public Function ActualizarStock(ByVal LineasAlbaran As DataRow()) As StockUpdateData()
    '    Dim updateDataArray(-1) As StockUpdateData
    '    Dim acc As New AlbaranCompraCabecera
    '    Dim aclt As New AlbaranCompraLote
    '    Dim ofc As BusinessHelper = BusinessHelper.CreateBusinessObject("OFControl")
    '    Dim [or] As BusinessHelper = BusinessHelper.CreateBusinessObject("OrdenRuta")
    '    Dim oe As BusinessHelper = BusinessHelper.CreateBusinessObject("OrdenEstructura")

    '    Dim OperarioGenerico As String = New Parametro().OperarioGenerico()

    '    For Each dr As DataRow In LineasAlbaran
    '        Me.BeginTx()
    '        Dim Linea As DataTable = Me.SelOnPrimaryKey(dr(_ACL.IDLineaAlbaran))
    '        If Linea.Rows.Count > 0 Then
    '            Dim lineaAlbaran As DataRow = Linea.Rows(0)
    '            If lineaAlbaran(_ACL.EstadoStock) = enumaclEstadoStock.aclNoActualizado Then
    '                If lineaAlbaran(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclNormal _
    '                Or lineaAlbaran(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclKit _
    '                Or lineaAlbaran(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclRealquiler _
    '                Or (lineaAlbaran(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclSubcontratacion And IsDBNull(lineaAlbaran(_ACL.IDOrdenRuta))) _
    '                Or (lineaAlbaran(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclComponente And IsDBNull(lineaAlbaran(_ACL.IDOrdenRuta))) Then
    '                    '//Linea NORMAL, o de tipo KIT, o SUBCONTRATACION MANUAL(que NO proviene de una OF)

    '                    Dim Lotes(-1) As DataTable
    '                    Dim Cabecera As DataTable = acc.SelOnPrimaryKey(lineaAlbaran(_ACL.IDAlbaran))
    '                    If Cabecera.Rows.Count > 0 Then
    '                        Dim updateData() As StockUpdateData
    '                        Dim lote As DataTable
    '                        Dim f As New Filter
    '                        f.Add(New NumberFilterItem(_ACLT.IDLineaAlbaran, lineaAlbaran(_ACL.IDLineaAlbaran)))
    '                        lote = aclt.Filter(f)
    '                        If lote.Rows.Count > 0 Then
    '                            updateData = Me.ActualizarStock(Cabecera.Rows(0), lineaAlbaran, lote)
    '                            ArrayManager.Copy(lote, Lotes)
    '                        Else
    '                            updateData = Me.ActualizarStock(Cabecera.Rows(0), lineaAlbaran)
    '                        End If
    '                        ArrayManager.Copy(updateData, updateDataArray)

    '                        AdminData.SetData(Cabecera)
    '                        AdminData.SetData(Linea)
    '                        AdminData.SetData(Lotes)
    '                    End If

    '                ElseIf lineaAlbaran(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclSubcontratacion And Nz(lineaAlbaran(_ACL.IDOrdenRuta), 0) <> 0 Then
    '                    '//Lineas de SUBCONTRATACION QUE PROVIENEN DE UNA ORDEN DE FABRICACION.
    '                    '//Las lineas de albaran se actualizaran automaticamente, independientemente
    '                    '//de que la actualizacion del stock se haga correctamente. Los movimientos 
    '                    '//pendientes de actualizar, si existen, se gestionaran desde el programa
    '                    '//'Movimientos de Stock asociados a la Orden' (tbOFControlEstructura).

    '                    Dim ParteTrabajo As DataTable
    '                    Dim produccionLog As ControlProduccionUpdateData

    '                    '//parte de trabajo (registro de tbOFControl)
    '                    If IsNumeric(lineaAlbaran("IDOFControl")) Then
    '                        ParteTrabajo = ofc.SelOnPrimaryKey(lineaAlbaran("IDOFControl"))
    '                    Else
    '                        Dim Cabecera As DataTable = acc.SelOnPrimaryKey(lineaAlbaran(_ACL.IDAlbaran))
    '                        If Cabecera.Rows.Count > 0 Then
    '                            Dim FechaActualizacion As Date = Cabecera.Rows(0)("FechaAlbaran")

    '                            Dim operacion As DataRow
    '                            If IsNumeric(lineaAlbaran("IDOrdenRuta")) Then
    '                                operacion = [or].GetItemRow(lineaAlbaran("IDOrdenRuta"))
    '                                ParteTrabajo = ofc.AddNewForm()
    '                                Dim parte As DataRow = ParteTrabajo.Rows(0)
    '                                parte = ofc.ApplyBusinessRule("IDOrden", operacion("IDOrden"), parte)
    '                                parte("FechaInicio") = FechaActualizacion
    '                                parte("FechaFin") = FechaActualizacion
    '                                parte("IDOperario") = OperarioGenerico
    '                                parte("IDOrdenRuta") = operacion("IDOrdenRuta")
    '                                parte("Secuencia") = operacion("Secuencia")
    '                                parte = ofc.ApplyBusinessRule("Secuencia", parte("Secuencia"), parte)
    '                                parte("QBuenaUdProduccion") = lineaAlbaran("QServida")
    '                                parte("QRechazadaUdProduccion") = 0
    '                                parte("QDudosaUdProduccion") = 0
    '                                parte = ofc.ApplyBusinessRule("QBuenaUdProduccion", parte("QBuenaUdProduccion"), parte)
    '                                parte = ofc.ApplyBusinessRule("QRechazadaUdProduccion", parte("QRechazadaUdProduccion"), parte)
    '                                parte = ofc.ApplyBusinessRule("QDudosaUdProduccion", parte("QDudosaUdProduccion"), parte)

    '                                lineaAlbaran("IDOFControl") = parte("IDOFControl")
    '                                lineaAlbaran("EstadoStock") = enumaclEstadoStock.aclActualizado
    '                            End If
    '                        End If
    '                    End If

    '                    '//Obtener las lineas componentes de la linea de subcontratacion actual
    '                    Dim f As New Filter
    '                    f.Add(New NumberFilterItem("IDLineaPadre", lineaAlbaran(_ACL.IDLineaAlbaran)))
    '                    Dim componentes As DataTable = Me.Filter(f)
    '                    If componentes.Rows.Count > 0 Then
    '                        For Each componente As DataRow In componentes.Rows
    '                            componente("IDOFControl") = lineaAlbaran("IDOFControl")
    '                            componente("EstadoStock") = enumaclEstadoStock.aclActualizado
    '                        Next
    '                    End If

    '                    produccionLog = CType(ofc, IControlProduccion).ControlProduccion(ParteTrabajo)
    '                    If Not produccionLog Is Nothing Then
    '                        ArrayManager.Copy(produccionLog.Entradas, updateDataArray)
    '                        ArrayManager.Copy(produccionLog.Salidas, updateDataArray)
    '                    End If

    '                    AdminData.SetData(Linea)
    '                    AdminData.SetData(componentes)
    '                End If
    '            End If
    '        End If
    '        Me.CommitTx()
    '    Next

    '    Return updateDataArray
    'End Function

#End Region

#Region " Corregir Movimientos "
    <Task()> Public Shared Sub CorregirMovimientos(ByVal Doc As DocumentoAlbaranCompra, ByVal services As ServiceProvider)
        Dim alog As AlbaranLogProcess = services.GetService(Of AlbaranLogProcess)()
        Dim aStockUpdateData(-1) As StockUpdateData
        aStockUpdateData = ProcessServer.ExecuteTask(Of DocumentoAlbaranCompra, StockUpdateData())(AddressOf CorreccionMovimientosCambiosCabecera, Doc, services)
        If Not aStockUpdateData Is Nothing Then ArrayManager.Copy(aStockUpdateData, alog.StockUpdateData)
        aStockUpdateData = ProcessServer.ExecuteTask(Of DocumentoAlbaranCompra, StockUpdateData())(AddressOf CorreccionMovimientosCambiosLineas, Doc, services)
        If Not aStockUpdateData Is Nothing Then ArrayManager.Copy(aStockUpdateData, alog.StockUpdateData)
    End Sub

    '<Task()> Public Shared Function CorregirMovimientos(ByVal Doc As DocumentoAlbaranCompra, ByVal services As ServiceProvider) As StockUpdateData()
    '    Dim returnData(-1) As StockUpdateData

    '    Dim aStockUpdateData(-1) As StockUpdateData
    '    aStockUpdateData = ProcessServer.ExecuteTask(Of DocumentoAlbaranCompra, StockUpdateData())(AddressOf CorreccionMovimientosCambiosCabecera, Doc, services)
    '    ArrayManager.Copy(aStockUpdateData, returnData)
    '    aStockUpdateData = ProcessServer.ExecuteTask(Of DocumentoAlbaranCompra, StockUpdateData())(AddressOf CorreccionMovimientosCambiosLineas, Doc, services)
    '    ArrayManager.Copy(aStockUpdateData, returnData)

    '    Return returnData
    'End Function

    <Task()> Public Shared Function CorreccionMovimientosCambiosCabecera(ByVal Doc As DocumentoAlbaranCompra, ByVal services As ServiceProvider) As StockUpdateData()
        Dim returnData(-1) As StockUpdateData
        If Doc.HeaderRow.RowState = DataRowState.Modified Then
            If Doc.HeaderRow("FechaAlbaran", DataRowVersion.Original) <> Doc.HeaderRow("FechaAlbaran") Then
                If Doc.HeaderRow("FechaAlbaran") <> DateTime.MinValue Then
                    Dim FechaDocumento As Date = Doc.HeaderRow("FechaAlbaran")
                    Dim f As New Filter
                    f.Add(New NumberFilterItem("IDAlbaran", Doc.HeaderRow("IDAlbaran")))
                    f.Add(New NumberFilterItem("EstadoStock", enumaclEstadoStock.aclActualizado))
                    Dim WhereStockActualizado As String = f.Compose(New AdoFilterComposer)
                    Dim lineasAlbaran() As DataRow = Doc.dtLineas.Select(WhereStockActualizado)
                    If Not lineasAlbaran Is Nothing Then
                        Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
                        If AppParams.GestionInventarioPermanente Then
                            Dim LineasContabilizadas As List(Of DataRow) = (From c In Doc.dtLineas _
                                                                            Where Not c.IsNull("Contabilizado") AndAlso _
                                                                            c("Contabilizado") <> CInt(enumContabilizado.NoContabilizado)).ToList()
                            If LineasContabilizadas.Count > 0 Then
                                ApplicationService.GenerateError("Existen lineas contabilizadas, no es posible realizar la corrección de los movimientos.")
                            End If
                        End If

                        For Each lineaAlbaran As DataRow In lineasAlbaran
                            f.Clear()
                            f.Add(New NumberFilterItem("IDLineaAlbaran", lineaAlbaran("IDLineaAlbaran")))
                            Dim WhereLotesLinea As String = f.Compose(New AdoFilterComposer)
                            Dim lotes() As DataRow = Doc.dtLote.Select(WhereLotesLinea)
                            If lotes.Length > 0 Then
                                For Each lote As DataRow In lotes
                                    Dim IDLineaMovimiento As Integer = 0
                                    ''//Movimiento de salida
                                    'If IsNumeric(lote("IDMovimientoSalida")) Then IDLineaMovimiento = lote("IDMovimientoSalida")

                                    '//Movimiento de entrada (si existe)
                                    If IsNumeric(lote("IDMovimientoEntrada")) Then IDLineaMovimiento = lote("IDMovimientoEntrada")

                                    If IDLineaMovimiento <> 0 Then
                                        Dim datCorrMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, IDLineaMovimiento, FechaDocumento, CStr(Doc.HeaderRow("NAlbaran")))
                                        Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datCorrMovto, services)
                                        If Not updateData Is Nothing AndAlso updateData.Estado = EstadoStock.NoActualizado Then
                                            ApplicationService.GenerateError(updateData.Detalle)
                                        End If
                                    End If

                                    If Length(lote("NEntrada")) > 0 Then
                                        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(GetType(EntityInfoCache(Of ArticuloInfo)))
                                        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(lineaAlbaran(_ACL.IDArticulo))

                                        Dim datIStock As New ProcesoStocks.DataCreateIStockClass(ArtInfo.EnsambladoStock, ArtInfo.ClaseStock)
                                        Dim IStockClassBdg As IStock = ProcessServer.ExecuteTask(Of ProcesoStocks.DataCreateIStockClass, IStock)(AddressOf ProcesoStocks.CreateIStockClass, datIStock, services)
                                        Dim StData As New DataFechaEntVino(lote("NEntrada"), FechaDocumento)
                                        IStockClassBdg.ActualizarFechaEntradaVino(StData)
                                    End If
                                Next
                            Else
                                Dim IDLineaMovimiento As Integer = 0
                                '//Movimiento en tbAlbaranCompraLinea
                                If IsNumeric(lineaAlbaran("IDMovimiento")) Then IDLineaMovimiento = lineaAlbaran("IDMovimiento")

                                If IDLineaMovimiento <> 0 Then
                                    Dim datCorrMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, IDLineaMovimiento, FechaDocumento, CStr(Doc.HeaderRow("NAlbaran")))
                                    Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datCorrMovto, services)
                                    If Not updateData Is Nothing AndAlso updateData.Estado = EstadoStock.NoActualizado Then
                                        ApplicationService.GenerateError(updateData.Detalle)
                                    End If
                                End If

                            End If
                        Next
                    End If
                End If
            End If
        End If
        Return returnData
    End Function

    <Task()> Public Shared Function CorreccionMovimientosCambiosLineas(ByVal Doc As DocumentoAlbaranCompra, ByVal services As ServiceProvider) As StockUpdateData()
        Dim aStockUpdateData(-1) As StockUpdateData
        'Dim f As New Filter
        'f.Add(New NumberFilterItem("TipoLineaAlbaran", FilterOperator.NotEqual, enumaclTipoLineaAlbaran.aclComponente))

        Dim AppParams As ParametroStocks = services.GetService(Of ParametroStocks)()
        Dim LineasModificadas As List(Of DataRow) = (From c In Doc.dtLineas Where c.RowState = DataRowState.Modified Select c).ToList
        If LineasModificadas Is Nothing OrElse LineasModificadas.Count = 0 Then
            Exit Function
        End If
        For Each lineaAlbaran As DataRow In LineasModificadas 'Select(f.Compose(New AdoFilterComposer))
            If lineaAlbaran.RowState = DataRowState.Modified Then
                If lineaAlbaran(_ACL.QServida, DataRowVersion.Original) <> lineaAlbaran(_ACL.QServida) OrElse _
                   lineaAlbaran(_ACL.QInterna, DataRowVersion.Original) <> lineaAlbaran(_ACL.QInterna) OrElse _
                   lineaAlbaran(_ACL.ImporteA, DataRowVersion.Original) <> lineaAlbaran(_ACL.ImporteA) OrElse _
                   lineaAlbaran(_ACL.ImporteB, DataRowVersion.Original) <> lineaAlbaran(_ACL.ImporteB) OrElse _
                  (lineaAlbaran.Table.Columns.Contains("QInterna2") AndAlso Nz(lineaAlbaran("QInterna2", DataRowVersion.Original), 0) <> Nz(lineaAlbaran("QInterna2"), 0)) Then
                    If Nz(lineaAlbaran(_ACL.IDOrdenRuta), 0) = 0 Or lineaAlbaran(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclComponente Then
                        Dim BlnFind As Boolean = False
                        If Nz(lineaAlbaran(_ACL.PrecioA), 0) <> Nz(lineaAlbaran(_ACL.PrecioA, DataRowVersion.Original), 0) AndAlso Length(lineaAlbaran("Lote")) > 0 Then
                            Dim ClsArtSerie As BusinessHelper = BusinessHelper.CreateBusinessObject("ArticuloNSerie")
                            Dim FilArtSerie As New Filter
                            FilArtSerie.Add("IDArticulo", FilterOperator.Equal, lineaAlbaran("IDArticulo"))
                            FilArtSerie.Add("NSerie", FilterOperator.Equal, lineaAlbaran("Lote"))
                            Dim DtArtNSerie As DataTable = ClsArtSerie.Filter(FilArtSerie, , "IDEstadoActivo")
                            If Not DtArtNSerie Is Nothing AndAlso DtArtNSerie.Rows.Count > 0 Then
                                Dim ClsEstadoActivo As BusinessHelper = BusinessHelper.CreateBusinessObject("MntoEstadoActivo")
                                Dim DtEstadoActivo As DataTable = ClsEstadoActivo.SelOnPrimaryKey(DtArtNSerie.Rows(0)("IDEstadoActivo"))
                                If Not DtEstadoActivo Is Nothing AndAlso DtEstadoActivo.Rows.Count > 0 Then
                                    If DtEstadoActivo.Rows(0)("Baja") Then
                                        Dim DtHist As DataTable = New BE.DataEngine().Filter("tbHistoricoMovimiento", New FilterItem("IDLineaMovimiento", FilterOperator.Equal, lineaAlbaran("IDMovimiento")))
                                        If Not DtHist Is Nothing AndAlso DtHist.Rows.Count > 0 Then
                                            For Each DrHist As DataRow In DtHist.Select
                                                DrHist("PrecioA") = lineaAlbaran("PrecioA")
                                            Next
                                            DtHist.TableName = "Stock"
                                            BusinessHelper.UpdateTable(DtHist)
                                            BlnFind = True
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        If Not BlnFind Then
                            Dim ctx As New DataDocRow(Doc, lineaAlbaran)
                            Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of DataDocRow, StockUpdateData)(AddressOf CorregirMovimiento, ctx, services)
                            'If Not updateData Is Nothing Then
                            '    ReDim Preserve aStockUpdateData(aStockUpdateData.Length)
                            '    aStockUpdateData(aStockUpdateData.Length - 1) = updateData
                            'End If
                            If Not updateData Is Nothing AndAlso updateData.Estado = EstadoStock.NoActualizado Then
                                ArrayManager.Copy(updateData, aStockUpdateData)
                            End If

                            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                            Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(lineaAlbaran(_ACL.IDArticulo))

                            Dim DtAlbCompraLote As DataTable = New AlbaranCompraLote().Filter(New FilterItem("IDLineaAlbaran", FilterOperator.Equal, lineaAlbaran(_ACL.IDLineaAlbaran)))
                            If Not DtAlbCompraLote Is Nothing AndAlso DtAlbCompraLote.Rows.Count > 0 Then
                                Dim datIStock As New ProcesoStocks.DataCreateIStockClass(ArtInfo.EnsambladoStock, ArtInfo.ClaseStock)
                                Dim IStockClassBdg As IStock = ProcessServer.ExecuteTask(Of ProcesoStocks.DataCreateIStockClass, IStock)(AddressOf ProcesoStocks.CreateIStockClass, datIStock, services)
                                For Each DrLote As DataRow In DtAlbCompraLote.Select
                                    If Length(DrLote("NEntrada")) > 0 Then
                                        Dim StData As New DataPrecioEntVino(DrLote("NEntrada"), lineaAlbaran("Precio"))
                                        IStockClassBdg.ActualizarPrecioEntradaVino(StData)
                                    End If
                                Next
                            End If
                        End If
                    Else
                        If lineaAlbaran(_ACL.QServida, DataRowVersion.Original) <> lineaAlbaran(_ACL.QServida) OrElse _
                           lineaAlbaran(_ACL.QInterna, DataRowVersion.Original) <> lineaAlbaran(_ACL.QInterna) OrElse _
                          (lineaAlbaran.Table.Columns.Contains("QInterna2") AndAlso Nz(lineaAlbaran("QInterna2", DataRowVersion.Original), 0) <> Nz(lineaAlbaran("QInterna2"), 0)) Then
                            If lineaAlbaran(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclSubcontratacion Then
                                If Nz(lineaAlbaran(_ACL.IDOFControl), 0) <> 0 Then
                                    Dim ofc As BusinessHelper = BusinessHelper.CreateBusinessObject("OFControl")
                                    Dim parteTrabajo As DataTable = ofc.SelOnPrimaryKey(lineaAlbaran(_ACL.IDOFControl))
                                    If parteTrabajo.Rows.Count > 0 Then
                                        parteTrabajo.Rows(0)("QBuena") = lineaAlbaran(_ACL.QInterna)
                                        parteTrabajo.Rows(0)("QBuenaUDProduccion") = lineaAlbaran(_ACL.QInterna)
                                        parteTrabajo.Rows(0)("QRechazada") = Nz(lineaAlbaran("QRechazada"), 0)
                                        parteTrabajo.Rows(0)("QRechazadaUDProduccion") = Nz(lineaAlbaran("QRechazada"), 0)
                                        CType(ofc, IControlProduccion).ControlProduccion(parteTrabajo)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next
        Return aStockUpdateData
    End Function

    <Task()> Public Shared Function CorregirMovimiento(ByVal ctx As DataDocRow, ByVal services As ServiceProvider) As StockUpdateData
        '//Lineas de albaran de tipo subcontratacion se actualizan desde el control de la produccion.
        Dim lineaAlbaran As DataRow = ctx.Row
        If Nz(lineaAlbaran(_ACL.IDOrdenRuta), 0) = 0 Then
            Dim f As New Filter
            f.Add(New NumberFilterItem(_ACLT.IDLineaAlbaran, lineaAlbaran(_ACL.IDLineaAlbaran)))

            Dim updateData As StockUpdateData
            Dim Cantidad As Double
            Dim PrecioA As Double
            Dim PrecioB As Double

            '//Importes extras
            Dim ImporteExtraA As Double
            Dim ImporteExtraB As Double
            Dim Importes As DataTable = New AlbaranCompraPrecio().Filter(New NumberFilterItem("IDLineaAlbaran", lineaAlbaran(_ACL.IDLineaAlbaran))) 'CType(ctx.Doc, DocumentoAlbaranCompra).dtPrecios
            For Each importe As DataRow In Importes.Rows
                ImporteExtraA += importe("ImporteA")
                ImporteExtraB += importe("ImporteB")
            Next

            Cantidad = lineaAlbaran(_ACL.QInterna)
            If Cantidad <> 0 Then
                Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
                Dim monedaA As MonedaInfo = Monedas.MonedaA
                Dim monedaB As MonedaInfo = Monedas.MonedaB
                PrecioA = xRound((ImporteExtraA / Cantidad) + (lineaAlbaran(_ACL.PrecioA) / lineaAlbaran(_ACL.Factor) / lineaAlbaran(_ACL.UdValoracion) * (1 - lineaAlbaran(_ACL.Dto1) / 100) * (1 - lineaAlbaran(_ACL.Dto2) / 100) * (1 - lineaAlbaran(_ACL.Dto3) / 100) * (1 - lineaAlbaran(_ACL.Dto) / 100) * (1 - lineaAlbaran(_ACL.DtoProntoPago) / 100)), monedaA.NDecimalesPrecio)
                PrecioB = xRound((ImporteExtraB / Cantidad) + (lineaAlbaran(_ACL.PrecioB) / lineaAlbaran(_ACL.Factor) / lineaAlbaran(_ACL.UdValoracion) * (1 - lineaAlbaran(_ACL.Dto1) / 100) * (1 - lineaAlbaran(_ACL.Dto2) / 100) * (1 - lineaAlbaran(_ACL.Dto3) / 100) * (1 - lineaAlbaran(_ACL.Dto) / 100) * (1 - lineaAlbaran(_ACL.DtoProntoPago) / 100)), monedaB.NDecimalesPrecio)
            End If

            Dim SegundaUnidad As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, lineaAlbaran("IDArticulo"), services)
            ProcessServer.ExecuteTask(Of Object)(AddressOf Comunes.BeginTransaction, Nothing, services)
            Dim WhereLotesLinea As String = f.Compose(New AdoFilterComposer)
            Dim lote() As DataRow = CType(ctx.Doc, DocumentoAlbaranCompra).dtLote.Select(WhereLotesLinea)
            If Not lote Is Nothing AndAlso lote.Length > 0 Then
                For Each dr As DataRow In lote
                    If Not dr.IsNull(_ACLT.IDMovimientoEntrada) Then
                        '//Correccion movimiento de entrada
                        Dim datCorrMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, dr(_ACLT.IDMovimientoEntrada), PrecioA, PrecioB)
                        updateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datCorrMovto, services)

                        If updateData Is Nothing OrElse updateData.Estado <> EstadoStock.Actualizado Then
                            If updateData.Estado = EstadoStock.NoActualizado AndAlso Length(lineaAlbaran(_ACL.IDMovimiento)) > 0 Then
                                Dim actMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Eliminar, lineaAlbaran(_ACL.IDMovimiento))
                                ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento)(AddressOf ProcesoStocks.ActualizarMovimiento, actMovto, services)
                                Dim act As New ProcesoStocks.DataActualizarLineas(updateData, lineaAlbaran)
                                ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarLineas)(AddressOf ProcesoStocks.ActualizarLineasAC, act, services)
                            End If
                            Return updateData
                        End If
                    End If
                Next
            Else
                If Not lineaAlbaran.IsNull(_ACL.IDMovimiento) Then
                    If Length(lineaAlbaran(_ACL.Lote)) > 0 Then
                        Dim datCorrMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, lineaAlbaran(_ACL.IDMovimiento), PrecioA, PrecioB)
                        If SegundaUnidad AndAlso Length(lineaAlbaran("QInterna2")) > 0 Then
                            datCorrMovto.Cantidad2 = CDbl(lineaAlbaran("QInterna2"))
                            datCorrMovto.CorrectContext.CorreccionEnCantidad2 = True
                        End If
                        updateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datCorrMovto, services)
                    Else
                        Dim datCorrMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, lineaAlbaran(_ACL.IDMovimiento), PrecioA, PrecioB)
                        datCorrMovto.CorrectContext.CorreccionEnCantidad = lineaAlbaran.RowState = DataRowState.Added OrElse _
                                                                          (lineaAlbaran.RowState = DataRowState.Modified AndAlso Nz(lineaAlbaran(_ACL.QInterna), 0) <> Nz(lineaAlbaran(_ACL.QInterna, DataRowVersion.Original), 0))
                        If datCorrMovto.CorrectContext.CorreccionEnCantidad Then
                            datCorrMovto.Cantidad = CDbl(Nz(lineaAlbaran(_ACL.QInterna), 0))
                        End If

                        If SegundaUnidad AndAlso Length(lineaAlbaran("QInterna2")) > 0 Then
                            datCorrMovto.CorrectContext.CorreccionEnCantidad2 = lineaAlbaran.RowState = DataRowState.Added OrElse _
                                                                         (lineaAlbaran.RowState = DataRowState.Modified AndAlso Nz(lineaAlbaran("QInterna2"), 0) <> Nz(lineaAlbaran("QInterna2", DataRowVersion.Original), 0))
                            If datCorrMovto.CorrectContext.CorreccionEnCantidad2 Then
                                datCorrMovto.Cantidad2 = CDbl(lineaAlbaran("QInterna2"))
                            End If
                        End If
                        updateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datCorrMovto, services)
                    End If

                    If updateData Is Nothing OrElse updateData.Estado <> EstadoStock.Actualizado Then
                        If updateData.Estado = EstadoStock.NoActualizado AndAlso Length(lineaAlbaran(_ACL.IDMovimiento)) > 0 Then
                            Dim actMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Eliminar, lineaAlbaran(_ACL.IDMovimiento))
                            ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, actMovto, services)
                            Dim act As New ProcesoStocks.DataActualizarLineas(updateData, lineaAlbaran)
                            ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarLineas)(AddressOf ProcesoStocks.ActualizarLineasAC, act, services)
                        End If
                        Return updateData
                    End If
                End If
            End If
            ProcessServer.ExecuteTask(Of Boolean)(AddressOf Comunes.CommitTransaction, False, services)

            Return updateData
        End If
    End Function


#End Region

#End Region

#Region " Gastos - AlbaranCompraPrecio "

    '//Copia para pasar los representates desde pedido a albarán y de este a la factura.
    <Task()> Public Shared Sub CopiarGastos(ByVal Doc As DocumentoAlbaranCompra, ByVal services As ServiceProvider)
        Dim PKField As String

        For Each dr As DataRow In Doc.dtLineas.Rows
            If Length(dr("IDLineaPedido")) > 0 Then

                Dim dtGastosOrigen As DataTable : Dim f As New Filter
                Select Case Doc.EntidadLineas
                    Case GetType(AlbaranCompraLinea).Name
                        PKField = "IDLineaAlbaranPrecio"
                        f.Add(New NumberFilterItem("IDLineaPedido", dr("IDLineaPedido")))
                        dtGastosOrigen = New PedidoCompraPrecio().Filter(f)
                End Select

                If Not dtGastosOrigen Is Nothing Then
                    For Each drGasto As DataRow In dtGastosOrigen.Select
                        Dim IDLineaPedidoHija As Integer = drGasto("IDLineaPedidoHija")
                        Dim adr As DataRow() = Doc.dtLineas.Select("IDLineaPedido=" & IDLineaPedidoHija)
                        If Not adr Is Nothing AndAlso adr.Length > 0 Then
                            Dim drNewGasto As DataRow = Doc.dtPrecios.NewRow
                            drNewGasto(PKField) = AdminData.GetAutoNumeric
                            drNewGasto("IDLineaAlbaran") = dr("IDLineaAlbaran")
                            drNewGasto("IDLineaAlbaranHija") = adr(0)("IDLineaAlbaran")
                            drNewGasto("IDArticulo") = drGasto("IDArticulo")
                            drNewGasto("DescArticulo") = drGasto("DescArticulo")
                            drNewGasto("Porcentaje") = drGasto("Porcentaje")
                            drNewGasto("Importe") = drGasto("Importe")

                            Dim ValAyB As New ValoresAyB(New DataRowPropertyAccessor(drNewGasto), Doc.IDMoneda, Doc.CambioA, Doc.CambioB)
                            ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)

                            Doc.dtPrecios.Rows.Add(drNewGasto)
                        End If
                    Next
                End If
            End If
        Next
    End Sub

#End Region

#Region " Inventarios Permanentes "

    <Serializable()> _
    Public Class DataGetLineasDescontabilizar
        Public IDLineasAlbaran() As Object

        Public ApuntesAlbaran As DataTable

        Public Sub New(ByVal IDLineasAlbaran() As Object)
            Me.IDLineasAlbaran = IDLineasAlbaran
        End Sub
    End Class
    <Task()> Public Shared Function GetLineasDescontabilizar(ByVal data As DataGetLineasDescontabilizar, ByVal services As ServiceProvider) As DataGetLineasDescontabilizar
        Dim f As New Filter

        Dim fLineasAlbaran As New Filter
        fLineasAlbaran.Add(New InListFilterItem("IDLineaAlbaran", data.IDLineasAlbaran, FilterType.Numeric))
        f.Add(fLineasAlbaran)

        Dim fTipoApunte As New Filter
        fTipoApunte.Add(New NumberFilterItem("IDTipoApunte", enumDiarioTipoApunte.AlbaranCompra))
        f.Add(fTipoApunte)

        f.Add(New NumberFilterItem("Contabilizado", FilterOperator.NotEqual, enumContabilizado.NoContabilizado))
        f.Add(New NumberFilterItem("EstadoFactura", enumaclEstadoFactura.aclNoFacturado))
        data.ApuntesAlbaran = New BE.DataEngine().Filter("NegDescontabilizarAC", f)

        Return data
    End Function

#End Region

End Class

Public Interface IEntradaVino
    Function SincronizarSalida(ByVal NumeroMovimiento As Integer, ByVal data As StockData) As StockUpdateData

End Interface
