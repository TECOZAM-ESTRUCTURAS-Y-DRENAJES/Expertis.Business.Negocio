Public Class ProcesoAlbaranVenta

    Friend Shared _AVC As _AlbaranVentaCabecera
    Friend Shared _AVL As _AlbaranVentaLinea
    Friend Shared _PVL As _PedidoVentaLinea
    Friend Shared _PVC As _PedidoVentaCabecera
    Friend Shared _AA As _ArticuloAlmacen
    Friend Shared _AAL As _ArticuloAlmacenLote
    Friend Shared _AVLT As _AlbaranVentaLote

    <Task()> Public Shared Function GetDocumento(ByVal IDAlbaran As Integer, ByVal services As ServiceProvider) As DocumentoAlbaranVenta
        Return New DocumentoAlbaranVenta(IDAlbaran)
    End Function

    <Task()> Public Shared Function CrearDocumento(ByVal data As UpdatePackage, ByVal services As ServiceProvider) As DocumentoAlbaranVenta
        Return New DocumentoAlbaranVenta(data)
    End Function

#Region " Validaciones "

    <Task()> Public Shared Sub ValidarDocumento(ByVal Doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        Dim AVC As New AlbaranVentaCabecera
        AVC.Validate(Doc.HeaderRow.Table)

        '//Esto se valida fuera del Validate de las líneas por que necesita información de la cabecera
        ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ValidarEstadoLineas, Doc, services)

        Dim AVL As New AlbaranVentaLinea
        AVL.Validate(Doc.dtLineas)

        Dim AVR As New AlbaranVentaRepresentante
        AVR.Validate(Doc.dtVentaRepresentante)

    End Sub

    <Task()> Public Shared Sub ValidarFacturasPorCondicionEnvio(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDCondicionEnvio")) > 0 Then
            Dim CondEnvio As EntityInfoCache(Of CondicionEnvioInfo) = services.GetService(Of EntityInfoCache(Of CondicionEnvioInfo))()
            Dim CondEnvioInfo As CondicionEnvioInfo = CondEnvio.GetEntity(data("IDCondicionEnvio"))
            If Length(data("IDFacturaPortes")) > 0 AndAlso Not CondEnvioInfo.FacturaPortes Then
                ApplicationService.GenerateError("La Condición de Envío introducida no permite factura de portes.")
            End If
            If Length(data("IDFacturaDespacho")) > 0 AndAlso Not CondEnvioInfo.FacturaDespacho Then
                ApplicationService.GenerateError("La Condición de Envío introducida no permite factura de despacho.")
            End If
            If Length(data("IDFacturaOtros")) > 0 AndAlso Not CondEnvioInfo.FacturaOtros Then
                ApplicationService.GenerateError("La Condición de Envío introducida no permite factura otros.")
            End If
        Else
            data("IDFacturaPortes") = System.DBNull.Value
            data("IDFacturaDespacho") = System.DBNull.Value
            data("IDFacturaOtros") = System.DBNull.Value
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
    '        Dim dtAVC As DataTable = New AlbaranVentaCabecera().Filter(f)
    '        If Not dtAVC Is Nothing AndAlso dtAVC.Rows.Count > 0 Then
    '            If AppParamsConta.Contabilidad Then
    '                ApplicationService.GenerateError("El Albarán {0} ya existe para el Ejercicio {1}.", Quoted(data("NAlbaran")), Quoted(data("IDEjercicio")))
    '            Else
    '                ApplicationService.GenerateError("El Albarán {0} ya existe.", Quoted(data("NAlbaran")))
    '            End If
    '        End If
    '    End If
    'End Sub


    <Task()> Public Shared Sub ValidarAlbaranFacturado(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data("Estado") = enumavcEstadoFactura.avcFacturado Then
            If data.RowState = DataRowState.Modified Then

                Dim campos_permitidos As New List(Of String)
                campos_permitidos.Add("Conductor")
                campos_permitidos.Add("EmpresaTransp")
                campos_permitidos.Add("CifTransportista")
                campos_permitidos.Add("DNIConductor")
                campos_permitidos.Add("IDFormaEnvio")
                campos_permitidos.Add("Matricula")
                campos_permitidos.Add("Precinto")
                campos_permitidos.Add("NContenedor")
                campos_permitidos.Add("Remolque")
                campos_permitidos.Add("Vehiculo")
                campos_permitidos.Add("Departamento")
                campos_permitidos.Add("PesoBrutoManual")
                campos_permitidos.Add("PesoNetoManual")
                campos_permitidos.Add("NBultos")
                campos_permitidos.Add("NContenedores")
                campos_permitidos.Add("IDCondicionEnvio")
                campos_permitidos.Add("IDModoTransporte")
                campos_permitidos.Add("Sucursal")

                Dim campos_modificados As New List(Of String)

                For Each c As DataColumn In data.Table.Columns
                    If Nz(data(c)) <> Nz(data(c, DataRowVersion.Original)) Then
                        campos_modificados.Add(c.ColumnName)
                    End If
                Next

                Dim control As List(Of String) = (From c As String In campos_modificados Where Not campos_permitidos.Contains(c)).ToList()
                If control.Count > 0 Then
                    ApplicationService.GenerateError("El albarán está Facturado.")
                End If
            Else
                ApplicationService.GenerateError("El albarán está Facturado.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarArticuloContenedor(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Modified Then
            If Length(data(_AVL.IDEntradaContenedor)) > 0 AndAlso Length(data(_AVL.IDSalidaContenedor)) > 0 Then
                If data(_AVL.IDArticuloContenedor) & String.Empty <> data(_AVL.IDArticuloContenedor, DataRowVersion.Original) & String.Empty Then
                    ApplicationService.GenerateError("No se puede modificar el Articulo Contenedor, se ha actualizado ya el stock del contenedor.")
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarArticuloBloqueado(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDArticulo")) <> 0 Then
            Dim Cabecera As DataTable = New AlbaranVentaCabecera().SelOnPrimaryKey(data("IDAlbaran"))
            If Not Cabecera Is Nothing AndAlso Cabecera.Rows.Count > 0 Then
                Dim StDatos As New Cliente.DataBloqArtClie
                StDatos.IDArticulo = data("IDArticulo") : StDatos.IDCliente = Cabecera.Rows(0)("IDCliente")
                If ProcessServer.ExecuteTask(Of Cliente.DataBloqArtClie, Boolean)(AddressOf Cliente.ComprobarBloqueoArticuloCliente, StDatos, services) Then
                    ApplicationService.GenerateError("El Artículo está bloqueado para este Cliente.")
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarEstadoLineas(ByVal Doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        'If Not Nz(Doc.HeaderRow(_AVC.Automatico), False) Then
        For Each dr As DataRow In Doc.dtLineas.Rows
            If dr.RowState = DataRowState.Modified Then
                If Not (dr(_AVL.EstadoFactura, DataRowVersion.Original) = enumavlEstadoFactura.avlNoFacturado) Then
                    If dr(_AVL.IDArticulo) <> dr(_AVL.IDArticulo, DataRowVersion.Original) OrElse dr(_AVL.QServida) <> dr(_AVL.QServida, DataRowVersion.Original) Then
                        If dr(_AVL.EstadoFactura, DataRowVersion.Original) = enumavlEstadoFactura.avlParcFacturado Then ApplicationService.GenerateError("No se puede modificar la línea de Albarán. Está parcialmente facturada.")
                        If dr(_AVL.EstadoFactura, DataRowVersion.Original) = enumavlEstadoFactura.avlFacturado Then ApplicationService.GenerateError("No se puede modificar la línea de Albarán. Está facturada.")
                    End If
                End If
            End If
        Next
        'End If
    End Sub

#End Region

#Region " Tipo de Albarán "

    <Serializable()> _
    Public Class TipoAlbaranInfo
        Public IDTipo As String
        Public Tipo As enumTipoAlbaran
        Public EsFacturable As Boolean
    End Class

    <Task()> Public Shared Function TipoDeAlbaran(ByVal IDTipoAlbaran As String, ByVal services As ServiceProvider) As TipoAlbaranInfo
        Dim info As New TipoAlbaranInfo
        If Len(IDTipoAlbaran) > 0 Then
            info.IDTipo = IDTipoAlbaran

            Dim f As New Filter
            f.Add(New StringFilterItem("Entidad", FilterOperator.Equal, "TipoAlbaran"))
            f.Add(New StringFilterItem("CampoEntidad", FilterOperator.Equal, "IDTipoAlbaran"))
            Dim TiposParametrizados As DataTable = New Parametro().Filter(f)
            If Not TiposParametrizados Is Nothing AndAlso TiposParametrizados.Rows.Count > 0 Then
                f.Clear()
                f.Add(New StringFilterItem("Valor", IDTipoAlbaran))
                TiposParametrizados.DefaultView.RowFilter = f.Compose(New AdoFilterComposer)
                If TiposParametrizados.DefaultView.Count > 0 Then
                    Dim ID As String = TiposParametrizados.DefaultView(0).Row("IDParametro")
                    Select Case ID
                        Case Parametro.cgFwAlbaranDefault         '"TIPO_ALB"
                            info.Tipo = enumTipoAlbaran.Normal
                        Case Parametro.cgFwAlbaranServicioDefault '"TIPO_ALB_S"
                            info.Tipo = enumTipoAlbaran.Servicio
                        Case Parametro.cgFwAlbaranDeDeposito      '"TIPO_ALB_D"
                            info.Tipo = enumTipoAlbaran.Deposito
                        Case Parametro.cgFwAlbaranDeConsumo       '"TIPO_ALB_C"
                            info.Tipo = enumTipoAlbaran.Consumo
                        Case Parametro.cgFwAlbaranRetornoAlquiler '"TIPO_ALB_A"
                            info.Tipo = enumTipoAlbaran.RetornoAlquiler
                        Case Parametro.cgFwAlbaranDeDevolucion    '"TIPOALB_DV"
                            info.Tipo = enumTipoAlbaran.DevolucionGeneral
                        Case Parametro.cgFwAlbaranDeIntercambio
                            info.Tipo = enumTipoAlbaran.Intercambio
                        Case Parametro.cnTIPO_ALB_DISTRIBUIDOR    '"TIPOALB_ED"
                            info.Tipo = enumTipoAlbaran.ExpedDistribuidor
                        Case Parametro.cnTIPO_ALB_ABONO_DISTRIBUIDOR     '"TIPOALB_AD"
                            info.Tipo = enumTipoAlbaran.AbonoDistribuidor
                        Case Else
                            info.Tipo = enumTipoAlbaran.NoEsDeSistema
                    End Select
                Else
                    info.Tipo = enumTipoAlbaran.NoEsDeSistema
                End If
            End If
        Else
            Dim AppParams As ParametroAlbaranVenta = services.GetService(Of ParametroAlbaranVenta)()
            IDTipoAlbaran = AppParams.TipoAlbaranPorDefecto
            If Len(IDTipoAlbaran) = 0 Then
                ApplicationService.GenerateError("El parámetro del sistema 'TIPO_ALB' no está correctamente configurado.")
            Else
                info.IDTipo = IDTipoAlbaran
                info.Tipo = enumTipoAlbaran.Normal
            End If
        End If

        Dim tipo As New Negocio.TipoAlbaran
        Dim dt As DataTable = tipo.SelOnPrimaryKey(IDTipoAlbaran)
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            info.EsFacturable = dt.Rows(0)("Facturable")
        End If

        Return info
    End Function

    <Task()> Public Shared Function ValidarTipoAlbaran(ByVal IDTipoAlbaran As String, ByVal services As ServiceProvider) As BusinessEnum.enumTipoAlbaran
        Dim valorTipoAlbaran As enumTipoAlbaran
        If String.IsNullOrEmpty(IDTipoAlbaran) Then
            valorTipoAlbaran = valorTipoAlbaran.Normal
        Else
            valorTipoAlbaran = ProcessServer.ExecuteTask(Of String, ProcesoAlbaranVenta.TipoAlbaranInfo)(AddressOf ProcesoAlbaranVenta.TipoDeAlbaran, IDTipoAlbaran, services).Tipo
        End If
        services.RegisterService(valorTipoAlbaran, GetType(BusinessEnum.enumTipoAlbaran))
        Return valorTipoAlbaran
    End Function
#End Region

#Region " Proceso Albaran "

#Region " Preparación de Información de Proceso y Validaciones "

    '//Prepara información de entrada al proceso
    <Task()> Public Shared Sub DatosIniciales(ByVal data As DataPrcAlbaranar, ByVal services As ServiceProvider)
        If data.AlbVentaInfo IsNot Nothing Then
            '//Prepara en el service lo que va a retornar el proceso, para tenerlo disponible en todo el proceso.
            services.RegisterService(New AlbaranLogProcess)

            '//Preparar información necaesaria a lo largo del proceso
            ProcessServer.ExecuteTask(Of DataPrcAlbaranar)(AddressOf ProcesoAlbaranVenta.PrepararInformacionProceso, data, services)
        End If
    End Sub
    'David V 3/3/22
    <Task()> Public Shared Sub DatosIniciales2(ByVal data As DataPrcAlbaranarDeposito, ByVal services As ServiceProvider)
        If data.AlbVentaInfo IsNot Nothing Then
            '//Prepara en el service lo que va a retornar el proceso, para tenerlo disponible en todo el proceso.
            services.RegisterService(New AlbaranLogProcess)

            '//Preparar información necaesaria a lo largo del proceso
            ProcessServer.ExecuteTask(Of DataPrcAlbaranarDeposito)(AddressOf ProcesoAlbaranVenta.PrepararInformacionProceso2, data, services)
        End If
    End Sub
    '//Preparar información necesaria a lo largo del proceso 
    <Task()> Public Shared Sub PrepararInformacionProceso2(ByVal data As DataPrcAlbaranarDeposito, ByVal services As ServiceProvider)
        Dim TipoInfo As TipoAlbaranInfo = ProcessServer.ExecuteTask(Of String, TipoAlbaranInfo)(AddressOf ProcesoAlbaranVenta.TipoDeAlbaran, data.IDTipoAlbaran & String.Empty, services)
        If Length(data.IDTipoAlbaran) = 0 Then
            data.IDTipoAlbaran = TipoInfo.IDTipo
        End If
        If data.FechaAlbaran = cnMinDate Then data.FechaAlbaran = Today

        services.RegisterService(New ProcessInfoAV(data.IDContador, data.IDTipoAlbaran, data.FechaAlbaran, data.TipoExpedicion))
    End Sub
    'David V 3/3/22
    '//Preparar información necesaria a lo largo del proceso 
    <Task()> Public Shared Sub PrepararInformacionProceso(ByVal data As DataPrcAlbaranar, ByVal services As ServiceProvider)
        Dim TipoInfo As TipoAlbaranInfo = ProcessServer.ExecuteTask(Of String, TipoAlbaranInfo)(AddressOf ProcesoAlbaranVenta.TipoDeAlbaran, data.IDTipoAlbaran & String.Empty, services)
        If Length(data.IDTipoAlbaran) = 0 Then
            data.IDTipoAlbaran = TipoInfo.IDTipo
        End If
        If data.FechaAlbaran = cnMinDate Then data.FechaAlbaran = Today

        services.RegisterService(New ProcessInfoAV(data.IDContador, data.IDTipoAlbaran, data.FechaAlbaran, data.TipoExpedicion))
    End Sub

    <Task()> Public Shared Function ValidacionesPreviasDesdeObras(ByVal oResltAgrp() As AlbCabVentaObras, ByVal services As ServiceProvider) As AlbCabVentaObras()
        Return ProcessServer.ExecuteTask(Of AlbCabVenta(), AlbCabVenta())(AddressOf ValidacionesPrevias, oResltAgrp, services)
    End Function

    <Task()> Public Shared Function ValidacionesPreviasDesdePedido(ByVal oResltAgrp() As AlbCabVentaPedido, ByVal services As ServiceProvider) As AlbCabVentaPedido()
        Return ProcessServer.ExecuteTask(Of AlbCabVenta(), AlbCabVenta())(AddressOf ValidacionesPrevias, oResltAgrp, services)
    End Function

    '//Validaciones antes de empezar a Crear Documentos
    <Task()> Public Shared Function ValidacionesPrevias(ByVal oResltAgrp() As AlbCabVenta, ByVal services As ServiceProvider) As AlbCabVenta()
        Dim ProcInfo As ProcessInfoAV = services.GetService(Of ProcessInfoAV)()
        Dim TipoInfo As TipoAlbaranInfo = TipoDeAlbaran(ProcInfo.IDTipoAlbaran, services)
        If TipoInfo.Tipo = enumTipoAlbaran.Desconocido Then
            ApplicationService.GenerateError("El tipo de albarán no es válido.")
        End If

        '//Excluir los pedidos de intercambios
        If TypeOf oResltAgrp Is AlbCabVentaPedido() Then
            Dim aux As New Collections.Generic.List(Of AlbCabVentaPedido)
            For Each pedido As AlbCabVentaPedido In oResltAgrp
                If TipoInfo.Tipo = enumTipoAlbaran.Intercambio Then
                    '//Solo seleccionar los pedidos de intercambio
                    If pedido.Intercambio Then
                        aux.Add(pedido)
                    End If
                Else
                    '//Excluir los pedidos de intercambio
                    If Not pedido.Intercambio Then
                        aux.Add(pedido)
                    End If
                End If
            Next
            oResltAgrp = aux.ToArray()
        End If

        If oResltAgrp Is Nothing OrElse oResltAgrp.Length = 0 Then
            ApplicationService.GenerateError("No hay datos para crear el tipo de albarán seleccionado.")
        End If
        Return oResltAgrp
    End Function

    <Task()> Public Shared Sub ValidacionesContador(ByVal data As DataPrcAlbaranar, ByVal services As ServiceProvider)
        'ProcessServer.ExecuteTask(Of String)(AddressOf ProcesoComunes.ValidarContadorObligatorio, data.IDContador & String.Empty, services)

        Dim datosValCont As New ProcesoComunes.DataValidarContadorEntidad(data.IDContador, GetType(AlbaranVentaCabecera).Name)
        ProcessServer.ExecuteTask(Of ProcesoComunes.DataValidarContadorEntidad)(AddressOf ProcesoComunes.ValidarContadorEntidad, datosValCont, services)
    End Sub

#End Region

#Region " Creación del Documento Albarán (PrcCrearAlbPed) "

#Region " Cabecera (Documento) "

    <Task()> Public Shared Function CrearDocumentoAlbaranVenta(ByVal Alb As AlbCabVenta, ByVal services As ServiceProvider) As DocumentoAlbaranVenta
        Return New DocumentoAlbaranVenta(Alb, services)
    End Function

    <Task()> Public Shared Sub AsignarValoresPredeterminadosGenerales(ByVal alb As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoComunes.AsignarIdentificadorAlbaran, alb.HeaderRow, services)

        Dim InfoProc As ProcessInfoAV = services.GetService(Of ProcessInfoAV)()
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
        'If alb.HeaderRow.IsNull("ResponsableExpedicion") Then alb.HeaderRow("ResponsableExpedicion") = InfoProc.ResponsableExpedicion
        If alb.HeaderRow.IsNull("IDTipoAlbaran") Then alb.HeaderRow("IDTipoAlbaran") = InfoProc.IDTipoAlbaran
        If alb.HeaderRow.IsNull("Aparcado") Then alb.HeaderRow("Aparcado") = False

        ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf AsignarDAA, alb, services)
    End Sub

    <Task()> Public Shared Sub AsignarDAA(ByVal alb As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
        If AppParams.GestionBodegas Then
            If alb.HeaderRow.IsNull("IDDaa") AndAlso Not alb.Cabecera.IDDAA Is Nothing AndAlso Not CType(alb.Cabecera.IDDAA, Guid).Equals(Guid.Empty) Then
                alb.HeaderRow("IDDaa") = alb.Cabecera.IDDAA
                alb.HeaderRow("NDAA") = alb.Cabecera.NDAA
                alb.HeaderRow("IDDAABaseDatos") = alb.Cabecera.IDDAABaseDatos
                alb.HeaderRow("AadReferenceCode") = alb.Cabecera.AadReferenceCode
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarDatosCliente(ByVal alb As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        If alb.HeaderRow.IsNull("IDCliente") Then Exit Sub
        ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoComercial.AsignarDatosCliente, alb, services)

        Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
        If alb.Cliente Is Nothing Then alb.Cliente = Clientes.GetEntity(alb.HeaderRow("IDCliente"))

        If alb.Cliente.Bloqueado Then ApplicationService.GenerateError("El Cliente está bloqueado.")
        If alb.HeaderRow.IsNull("IDFormaEnvio") Then alb.HeaderRow("IDFormaEnvio") = alb.Cliente.FormaEnvio
        If alb.HeaderRow.IsNull("IDCondicionEnvio") Then alb.HeaderRow("IDCondicionEnvio") = alb.Cliente.CondicionEnvio
        If alb.HeaderRow.IsNull("IDModoTransporte") Then alb.HeaderRow("IDModoTransporte") = alb.Cliente.IDModoTransporte
        If alb.HeaderRow.IsNull("DtoAlbaran") Then alb.HeaderRow("DtoAlbaran") = alb.Cliente.DtoComercial
        If alb.HeaderRow.IsNull("IDBancoPropio") Then alb.HeaderRow("IDBancoPropio") = alb.Cliente.IDBancoPropio

        If TypeOf alb.Cabecera Is AlbCabVentaPedido AndAlso Length(CType(alb.Cabecera, AlbCabVentaPedido).IDClienteDistribuidor) > 0 AndAlso alb.HeaderRow.IsNull("IDClienteDistribuidor") Then
            alb.HeaderRow("IDClienteDistribuidor") = CType(alb.Cabecera, AlbCabVentaPedido).IDClienteDistribuidor
        End If
      
    End Sub

    <Task()> Public Shared Sub AsignarDatosEDI(ByVal alb As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        If TypeOf alb.Cabecera Is AlbCabVentaPedido AndAlso Length(CType(alb.Cabecera, AlbCabVentaPedido).Muelle) > 0 AndAlso alb.HeaderRow.IsNull("muelle") Then
            alb.HeaderRow("muelle") = CType(alb.Cabecera, AlbCabVentaPedido).Muelle
        End If

        If TypeOf alb.Cabecera Is AlbCabVentaPedido AndAlso Length(CType(alb.Cabecera, AlbCabVentaPedido).Muelle) > 0 AndAlso alb.HeaderRow.IsNull("PuntoDescarga") Then
            alb.HeaderRow("PuntoDescarga") = CType(alb.Cabecera, AlbCabVentaPedido).PuntoDescarga
        End If
    End Sub

    <Task()> Public Shared Sub AsignarDatosTransporte(ByVal alb As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        If Not alb.HeaderRow("IDFormaEnvio") Is Nothing Then
            Dim FE As New FormaEnvio
            Dim detalle As New FormaEnvioDetalle
            Dim filtro As New Filter
            If Not alb.HeaderRow("IDFormaEnvio") Is Nothing Then
                filtro.Add("IDFormaEnvio", alb.HeaderRow("IDFormaEnvio"))
                Dim dt As DataTable = FE.Filter(filtro)
                If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                    alb.HeaderRow("EmpresaTransp") = Nz(dt.Rows(0)("DescFormaEnvio"), String.Empty)
                    filtro.Add("Predeterminado", True)
                    Dim dtDetalle As DataTable = detalle.Filter(filtro)
                    If Not dtDetalle Is Nothing AndAlso dtDetalle.Rows.Count > 0 Then
                        alb.HeaderRow("Conductor") = Nz(dtDetalle.Rows(0)("Conductor"), String.Empty)
                        alb.HeaderRow("DNIConductor") = Nz(dtDetalle.Rows(0)("DNIConductor"), String.Empty)
                        alb.HeaderRow("Matricula") = Nz(dtDetalle.Rows(0)("Matricula"), String.Empty)
                        alb.HeaderRow("Remolque") = Nz(dtDetalle.Rows(0)("Remolque"), String.Empty)
                    End If
                    Dim fProv As New Filter
                    If dt.Rows.Count > 0 Then
                        fProv.Add("IDProveedor", dt.Rows(0)("IDProveedor"))
                        Dim Prov As New Proveedor
                        Dim dtProv As DataTable = Prov.Filter(fProv)
                        If dtProv.Rows.Count > 0 Then
                            alb.HeaderRow("CifTransportista") = Nz(dtProv.Rows(0)("CifProveedor"), String.Empty)
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarDireccion(ByVal alb As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        If Not alb.HeaderRow.IsNull("IDDireccion") Then
            Dim StDatosDirec As New ClienteDireccion.DataDirecDe(alb.HeaderRow("IDDireccion"), enumcdTipoDireccion.cdDireccionEnvio)
            If Not ProcessServer.ExecuteTask(Of ClienteDireccion.DataDirecDe, Boolean)(AddressOf ClienteDireccion.EsDireccionDe, StDatosDirec, services) Then
                Dim DIR As New DataDireccionClte(enumcdTipoDireccion.cdDireccionEnvio, "IDDireccion", New DataRowPropertyAccessor(alb.HeaderRow))
                ProcessServer.ExecuteTask(Of DataDireccionClte)(AddressOf ProcesoComercial.AsignarDireccionCliente, DIR, services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarBanco(ByVal alb As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        If alb.HeaderRow.IsNull("IDClienteBanco") Then
            Dim IDBanco As Integer = New ClienteBanco().GetBancoPredeterminado(alb.HeaderRow("IDCliente"), services)
            If IDBanco > 0 Then
                alb.HeaderRow("IDClienteBanco") = IDBanco
            Else
                alb.HeaderRow("IDClienteBanco") = System.DBNull.Value
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarAlmacen(ByVal alb As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        Dim InfoProc As ProcessInfoAV = services.GetService(Of ProcessInfoAV)()
        Dim ParamsAV As ParametroAlbaranVenta = services.GetService(Of ParametroAlbaranVenta)()
        Dim strAlmDeposito As String : Dim blnDeposito As Boolean
        'TODO (REVISAR GESTIÓN DE DEVOLUCIONES) If InfoProc.IDTipoAlbaran = AlbaranVentaCabecera.enumTipoAlbaran.DevolucionCalidad AndAlso TipoOrigenDevolucion.Tipo = AlbaranVentaCabecera.enumTipoAlbaran.Deposito Then
        If InfoProc.IDTipoAlbaran = ParamsAV.TipoAlbaranDeDeposito AndAlso TypeOf alb.Cabecera Is AlbCabVentaPedido Then
            strAlmDeposito = CType(alb.Cabecera, AlbCabVentaPedido).IDAlmacenDeposito 'alb.HeaderRow("IDAlmacenDeposito") & String.Empty
            blnDeposito = True
        Else
            Dim f As New Filter
            f.Add(New NumberFilterItem("IDDireccion", alb.HeaderRow("IDDireccion")))
            f.Add(New BooleanFilterItem("Envio", True))
            Dim dtAlmDep As DataTable = AdminData.GetData("vNegPedidoClienteAlmacenDeposito", f, "TOP 1 IDAlmacen,Deposito")
            If Not dtAlmDep Is Nothing AndAlso dtAlmDep.Rows.Count > 0 Then
                strAlmDeposito = dtAlmDep.Rows(0)("IDAlmacen") & String.Empty
                blnDeposito = (Len(strAlmDeposito) > 0 And Nz(dtAlmDep.Rows(0)("Deposito"), False))
            End If
        End If

        If InfoProc.IDTipoAlbaran = ParamsAV.TipoAlbaranDeDeposito Or InfoProc.IDTipoAlbaran = ParamsAV.TipoAlbaranDeConsumo Then ' TODO Or (InfoProc.IDTipoAlbaran = AlbaranVentaCabecera.enumTipoAlbaran.DevolucionCalidad And TipoOrigenDevolucion.Tipo = AlbaranVentaCabecera.enumTipoAlbaran.Deposito) Then
            If Len(strAlmDeposito) > 0 Then
                alb.HeaderRow("IDAlmacenDeposito") = strAlmDeposito
                If Not blnDeposito Then
                    ApplicationService.GenerateError("El almacén {0} asignado a la dirección de envío del cliente {1} no es de depósito.", Quoted(strAlmDeposito), Quoted(alb.HeaderRow("IdCliente")))
                End If
            Else
                ApplicationService.GenerateError("La dirección de envío asignada al cliente {0} en el pedido {1} no tiene asociado un almacén.", Quoted(alb.HeaderRow("IdCliente")), Quoted(CType(alb.Cabecera, AlbCabVentaPedido).NOrigen))
            End If
        Else
            If blnDeposito Then
                ApplicationService.GenerateError("El almacén de envío {0} asociado al cliente {1} es de depósito.", Quoted(strAlmDeposito), Quoted(alb.HeaderRow("IdCliente")))
            End If
        End If
        If InfoProc.IDTipoAlbaran = ParamsAV.TipoAlbaranDeConsumo Then alb.HeaderRow("IDAlmacen") = strAlmDeposito
        If Len(strAlmDeposito) > 0 Then alb.HeaderRow("IDAlmacenDeposito") = strAlmDeposito

        ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoComunes.AsignarAlmacen, alb, services)

        If alb.HeaderRow.IsNull("IDAlmacen") Then ApplicationService.GenerateError("Debe haber un Almacén predeterminado.")
    End Sub

    <Task()> Public Shared Sub AsignarContador(ByVal alb As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoComunes.AsignarContador, alb, services)
        If Length(alb.HeaderRow("IDContador")) = 0 Then ApplicationService.GenerateError("Debe indicar un Contador para la entidad {0}.", Quoted(GetType(AlbaranVentaCabecera).Name))
    End Sub

    <Task()> Public Shared Sub TotalBultos(ByVal Doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        If Not IsNothing(Doc.dtLineas) AndAlso Doc.dtLineas.Rows.Count > 0 Then
            Dim NBultos As Integer = 0
            For Each linea As DataRow In Doc.dtLineas.Select
                Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(linea("IDArticulo"))
                If Not IsNothing(ArtInfo) Then
                    If Not ArtInfo.Especial Then
                        If Doc.HeaderRow.Table.Columns.Contains("NBultos") Then NBultos += linea(_AVL.QServida)
                    End If
                End If
            Next
        End If
    End Sub

    <Task()> Public Shared Sub TotalPesos(ByVal Doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        If Not IsNothing(Doc.HeaderRow) Then
            Dim QEtiEmbalaje As Integer
            Dim QEtiContenedor As Integer

            If Not IsNothing(Doc.dtLineas) AndAlso Doc.dtLineas.Rows.Count Then
                QEtiEmbalaje = Nz(Doc.dtLineas.Compute("SUM(QEtiEmbalaje)", Nothing), 0)
                QEtiContenedor = Nz(Doc.dtLineas.Compute("SUM(QEtiContenedor)", Nothing), 0)
            End If
            Dim PesoNetoManual As Double = Nz(Doc.HeaderRow("PesoNetoManual"), 0)
            Dim PesoBrutoManual As Double = Nz(Doc.HeaderRow("PesoBrutoManual"), 0)

            'Completar los pesos
            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
            Dim PesoNetoAcum As Double = 0
            Dim PesoBrutoAcum As Double = 0
            For Each linea As DataRow In Doc.dtLineas.Rows
                If linea("TipoLineaAlbaran") <> enumavlTipoLineaAlbaran.avlComponente Then
                    Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(linea("IDArticulo"))
                    PesoNetoAcum += (Nz(linea("QInterna"), 0) * ArtInfo.PesoNeto)
                    PesoBrutoAcum += (Nz(linea("QInterna"), 0) * ArtInfo.PesoBruto)
                End If
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

            'If Nz(Doc.HeaderRow("NBultos")) = 0 Then
            '    If QEtiEmbalaje <> 0 Then
            Doc.HeaderRow("NBultos") = Math.Abs(QEtiEmbalaje)
            'End If
            '    End If
            'If Nz(Doc.HeaderRow("NContenedores")) = 0 Then
            Doc.HeaderRow("NContenedores") = Math.Abs(QEtiContenedor)
            'End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarEstadoAlbaran(ByVal Doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        If Doc Is Nothing Then Exit Sub
        If Not IsNothing(Doc.dtLineas) Then
            Dim EstadosDiferentes As List(Of Object) = (From c In Doc.dtLineas _
                                                        Where Not c.IsNull("EstadoFactura") AndAlso _
                                                              Not c.IsNull("Facturable") AndAlso c("Facturable") = True AndAlso _
                                                              Not c.IsNull("TipoLineaAlbaran") AndAlso c("TipoLineaAlbaran") <> enumavlTipoLineaAlbaran.avlComponente _
                                                        Select c("EstadoFactura") Distinct).ToList
            If Not EstadosDiferentes Is Nothing AndAlso EstadosDiferentes.Count > 0 Then
                Select Case EstadosDiferentes.Count
                    Case 1
                        Doc.HeaderRow("Estado") = EstadosDiferentes(0)
                    Case Else
                        Doc.HeaderRow("Estado") = enumavcEstadoFactura.avcParcFacturado
                End Select
            Else
                Doc.HeaderRow("Estado") = enumavcEstadoFactura.avcNoFacturado
            End If
        End If
    End Sub

#End Region

#Region " Lineas (Documento) "

    <Task()> Public Shared Sub AsignarValoresPredeterminadosLinea(ByVal row As DataRow, ByVal services As ServiceProvider)
        row("IdLineaAlbaran") = AdminData.GetAutoNumeric
        row("TipoLineaAlbaran") = enumavlTipoLineaAlbaran.avlNormal
        row("EstadoStock") = enumavlEstadoStock.avlNoActualizado
        row("Dto1") = 0
        row("Dto2") = 0
        row("Dto3") = 0
        row("UdValoracion") = 1
        row("Factor") = 1
        row("QInterna") = 0
        row("QServida") = 0
        row("QFacturada") = 0
        row("Regalo") = False
        row("EstadoFactura") = enumavlEstadoFactura.avlNoFacturado

        Dim ProcInfo As ProcessInfoAV = services.GetService(Of ProcessInfoAV)()
        Dim TipoAlbInfo As TipoAlbaranInfo = ProcessServer.ExecuteTask(Of String, TipoAlbaranInfo)(AddressOf TipoDeAlbaran, ProcInfo.IDTipoAlbaran, services)

        row("Facturable") = TipoAlbInfo.EsFacturable
    End Sub

    <Task()> Public Shared Sub AsignarEstadoStock(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.Row("IDArticulo"))
        'El estado del stock de la línea depende de si el artículo tiene gestión de stock o no
        data.Row("EstadoStock") = enumavlEstadoStock.avlNoActualizado
        If Not ArtInfo.GestionStock Then
            data.Row("EstadoStock") = enumavlEstadoStock.avlSinGestion
        End If
        If Not data.Doc Is Nothing AndAlso Not data.Doc.HeaderRow Is Nothing AndAlso Length(data.Doc.HeaderRow("IDTipoAlbaran")) > 0 Then
            Dim TipoAlbInfo As ProcesoAlbaranVenta.TipoAlbaranInfo = ProcessServer.ExecuteTask(Of String, ProcesoAlbaranVenta.TipoAlbaranInfo)(AddressOf ProcesoAlbaranVenta.TipoDeAlbaran, data.Doc.HeaderRow("IDTipoAlbaran"), services)
            If TipoAlbInfo.Tipo = enumTipoAlbaran.ExpedDistribuidor OrElse TipoAlbInfo.Tipo = enumTipoAlbaran.AbonoDistribuidor Then
                data.Row("EstadoStock") = CInt(enumavlEstadoStock.avlSinGestion)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarDireccionFacturaEnLineas(ByVal oDocAlb As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        If oDocAlb.HeaderRow.RowState = DataRowState.Modified AndAlso Nz(oDocAlb.HeaderRow(_AVC.IDDireccionFra), 0) <> Nz(oDocAlb.HeaderRow(_AVC.IDDireccionFra, DataRowVersion.Original), 0) Then
            For Each linea As DataRow In oDocAlb.dtLineas.Rows
                linea(_AVL.IDDireccionFra) = oDocAlb.HeaderRow(_AVC.IDDireccionFra)
            Next
        Else
            For Each linea As DataRow In oDocAlb.dtLineas.Select(Nothing, Nothing, DataViewRowState.Added)
                linea(_AVL.IDDireccionFra) = oDocAlb.HeaderRow(_AVC.IDDireccionFra)
            Next
        End If
    End Sub

    <Task()> Public Shared Sub AsignarCondicionesEnLineas(ByVal oDocAlb As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        If oDocAlb.HeaderRow.RowState = DataRowState.Modified AndAlso Nz(oDocAlb.HeaderRow(_AVC.IDFormaPago), 0) <> Nz(oDocAlb.HeaderRow(_AVC.IDFormaPago, DataRowVersion.Original), 0) Then
            For Each linea As DataRow In oDocAlb.dtLineas.Rows
                linea(_AVL.IDFormaPago) = oDocAlb.HeaderRow(_AVC.IDFormaPago)
            Next
        Else
            For Each linea As DataRow In oDocAlb.dtLineas.Select(Nothing, Nothing, DataViewRowState.Added)
                linea(_AVL.IDFormaPago) = oDocAlb.HeaderRow(_AVC.IDFormaPago)
            Next
        End If

        If oDocAlb.HeaderRow.RowState = DataRowState.Modified AndAlso Nz(oDocAlb.HeaderRow(_AVC.IDCondicionPago), 0) <> Nz(oDocAlb.HeaderRow(_AVC.IDCondicionPago, DataRowVersion.Original), 0) Then
            For Each linea As DataRow In oDocAlb.dtLineas.Rows
                linea(_AVL.IDCondicionPago) = oDocAlb.HeaderRow(_AVC.IDCondicionPago)
            Next
        Else
            For Each linea As DataRow In oDocAlb.dtLineas.Select(Nothing, Nothing, DataViewRowState.Added)
                linea(_AVL.IDCondicionPago) = oDocAlb.HeaderRow(_AVC.IDCondicionPago)
            Next
        End If
    End Sub

    <Task()> Public Shared Sub AsignarClienteBancoEnLineas(ByVal oDocAlb As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        If oDocAlb.HeaderRow.RowState = DataRowState.Modified AndAlso oDocAlb.HeaderRow(_AVC.IDClienteBanco) & String.Empty <> oDocAlb.HeaderRow(_AVC.IDClienteBanco, DataRowVersion.Original) & String.Empty Then
            For Each linea As DataRow In oDocAlb.dtLineas.Rows
                linea(_AVL.IDClienteBanco) = oDocAlb.HeaderRow(_AVC.IDClienteBanco)
            Next
        Else
            For Each linea As DataRow In oDocAlb.dtLineas.Select(Nothing, Nothing, DataViewRowState.Added)
                linea(_AVL.IDClienteBanco) = oDocAlb.HeaderRow(_AVC.IDClienteBanco)
            Next
        End If
    End Sub

    <Task()> Public Shared Sub AsignarEstadoFacturaEnLineas(ByVal oDocAlb As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        If oDocAlb.HeaderRow.RowState = DataRowState.Modified AndAlso Nz(oDocAlb.HeaderRow(_AVC.Estado), -1) <> Nz(oDocAlb.HeaderRow(_AVC.Estado, DataRowVersion.Original), -1) Then
            For Each linea As DataRow In oDocAlb.dtLineas.Rows
                linea(_AVL.EstadoFactura) = oDocAlb.HeaderRow(_AVC.Estado)
            Next
        Else
            For Each linea As DataRow In oDocAlb.dtLineas.Select(Nothing, Nothing, DataViewRowState.Added)
                linea(_AVL.EstadoFactura) = oDocAlb.HeaderRow(_AVC.Estado)
            Next
        End If
    End Sub

#End Region

#End Region

    <Task()> Public Shared Sub GrabarDocumento(ByVal Doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf Business.General.Comunes.BeginTransaction, Nothing, services)
        'ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ValidarDocumento, Doc, services)
        ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf Business.General.Comunes.UpdateDocument, Doc, services)
        ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVentaPedidos.ActualizarPedidoDesdeAlbaran, Doc, services)
        ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVentaPedidos.ActualizarAlbaranClteDesdeAlbaran, Doc, services)
        ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVentaObras.ActualizarObrasDesdeAlbaran, Doc, services)
        ProcessServer.ExecuteTask(Of Boolean)(AddressOf Business.General.Comunes.CommitTransaction, True, services)
        ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ActualizarOrigenDAA, Doc, services)
    End Sub

    <Task()> Public Shared Sub ActualizarOrigenDAA(ByVal Doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        AdminData.BeginTx()

        Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
        If AppParams.GestionBodegas AndAlso Not IsDBNull(Doc.HeaderRow("IDDAA")) AndAlso Not CType(Doc.HeaderRow("IDDAA"), Guid).Equals(Guid.Empty) Then
            Dim DAAOrig As BusinessHelper = BusinessHelper.CreateBusinessObject("BdgDAAOrigenes")
            Dim dtBBDD As DataTable = AdminData.GetUserDataBases

            Dim BEDataEngine As New BE.DataEngine
            Dim currentBBDD As Guid = AdminData.GetConnectionInfo.IDDataBase
            For Each drEmpresa As DataRow In dtBBDD.Rows
                AdminData.SetCurrentConnection(drEmpresa("IDBaseDatos"))
                Dim fDaa As New Filter
                fDaa.Add(New GuidFilterItem("IDDAA", Doc.HeaderRow("IDDAA")))
                fDaa.Add(New GuidFilterItem("IDBaseDatos", currentBBDD))
                fDaa.Add(New IsNullFilterItem("IDPedido", False))
                Dim dtOrigen As DataTable = DAAOrig.Filter(fDaa)
                If dtOrigen.Rows.Count > 0 Then
                    dtOrigen.Rows(0)("IDAlbaran") = Doc.HeaderRow("IDAlbaran")
                    dtOrigen.Rows(0)("NAlbaran") = Doc.HeaderRow("NAlbaran")
                    DAAOrig.Update(dtOrigen)

                    Dim datActBdg As New DataReconstruirDAALineas
                    datActBdg.IDDaa = Doc.HeaderRow("IDDAA")
                    datActBdg.IDBaseDatosOrigen = currentBBDD
                    datActBdg.IDBaseDatosDAA = drEmpresa("IDBaseDatos")
                    datActBdg.IDAlbaran = Doc.HeaderRow("IDAlbaran")
                    datActBdg.dtCabeceraAlbaran = Doc.HeaderRow.Table

                    Dim BdgEMCS As IEMCS = ProcessServer.ExecuteTask(Of Object, IEMCS)(AddressOf Comunes.CreateEMCSGeneral, Nothing, services)
                    BdgEMCS.ReconstruirDAALineas(datActBdg, services)

                    Exit For
                End If
            Next
            AdminData.SetCurrentConnection(currentBBDD)
        End If
    End Sub

    <Task()> Public Shared Sub AñadirAResultado(ByVal oDocAlb As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        Dim alog As AlbaranLogProcess = services.GetService(Of AlbaranLogProcess)()
        If alog.CreateData Is Nothing Then alog.CreateData = New LogProcess
        ReDim Preserve alog.CreateData.CreatedElements(UBound(alog.CreateData.CreatedElements) + 1)
        alog.CreateData.CreatedElements(UBound(alog.CreateData.CreatedElements)) = New CreateElement
        alog.CreateData.CreatedElements(UBound(alog.CreateData.CreatedElements)).IDElement = oDocAlb.HeaderRow("IDAlbaran")
        alog.CreateData.CreatedElements(UBound(alog.CreateData.CreatedElements)).NElement = oDocAlb.HeaderRow("NAlbaran")
    End Sub

#End Region

#Region " Calcular albarán "

    <Task()> Public Shared Sub CalcularBasesImponibles(ByVal Doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        If Not IsNothing(Doc.dtLineas) Then
            Dim desglose() As DataBaseImponible = ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta, DataBaseImponible())(AddressOf ProcesoComunes.DesglosarImporte, Doc, services)
            Dim datosCalculo As New ProcesoComunes.DataCalculoTotalesCab
            datosCalculo.Doc = Doc
            datosCalculo.BasesImponibles = desglose
            ProcessServer.ExecuteTask(Of ProcesoComunes.DataCalculoTotalesCab)(AddressOf CalcularTotalesCabecera, datosCalculo, services)
            ' ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.CorregirMovimientos, Doc, services)
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
                Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
                Dim ClteInfo As ClienteInfo = Clientes.GetEntity(data.Doc.HeaderRow("IDCliente"))
                For Each BI As DataBaseImponible In data.BasesImponibles
                    ImporteLineas = ImporteLineas + BI.BaseImponible
                    If Length(BI.IDTipoIva) > 0 Then
                        Dim factor As Double = 0

                        ' HistoricoTipoIVA
                        Dim TIVAInfo As TipoIvaInfo = TiposIVA.GetEntity(BI.IDTipoIva, data.Doc.Fecha)

                        If Length(TIVAInfo.IDTipoIVA) > 0 Then
                            '//valor por defecto
                            factor = TIVAInfo.Factor

                            '//Para los ivas especiales que no se repercuten
                            If TIVAInfo.SinRepercutir Then factor = TIVAInfo.IVASinRepercutir
                        End If

                        Dim Base As Double
                        If BI.ImporteIVA <> 0 Then
                            ImporteIVATotal = ImporteIVATotal + xRound(BI.ImporteIVA - BI.ImporteIVA * 100 / (100 + factor), data.Doc.Moneda.NDecimalesImporte)
                            Base = xRound(BI.ImporteIVA - xRound(BI.ImporteIVA - BI.ImporteIVA * 100 / (100 + factor), data.Doc.Moneda.NDecimalesImporte), data.Doc.Moneda.NDecimalesImporte)
                            BaseImponibleTotal = BaseImponibleTotal + Base
                            ImporteLineas = BaseImponibleTotal
                        Else
                            BaseImponibleTotal = BaseImponibleTotal + BI.BaseImponible
                            Base = BI.BaseImponible
                            ImporteIVATotal = ImporteIVATotal + xRound(Base * factor / 100, data.Doc.Moneda.NDecimalesImporte)
                        End If

                        If ClteInfo.TieneRE Then
                            ImporteRETotal = ImporteRETotal + xRound(Base * TIVAInfo.IVARE / 100, data.Doc.Moneda.NDecimalesImporte)
                        End If
                    End If
                Next
            End If
        End If
        data.Doc.HeaderRow("BaseImponible") = BaseImponibleTotal
        data.Doc.HeaderRow("ImpIVA") = ImporteIVATotal
        data.Doc.HeaderRow("ImpRE") = ImporteRETotal
        data.Doc.HeaderRow("ImpAlbaran") = ImporteLineas

        If Nz(data.Doc.HeaderRow("RecFinan"), 0) > 0 Then
            Dim Total As Double = 0
            Total = BaseImponibleTotal + ImporteIVATotal + ImporteRETotal
            data.Doc.HeaderRow("ImpRecFinan") = xRound(Total * data.Doc.HeaderRow("RecFinan") / 100, data.Doc.Moneda.NDecimalesImporte)
        Else
            data.Doc.HeaderRow("ImpRecFinan") = 0
        End If

        Dim ValAyB As New ValoresAyB(New DataRowPropertyAccessor(data.Doc.HeaderRow), data.Doc.IDMoneda, data.Doc.CambioA, data.Doc.CambioB)
        ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)
    End Sub
#End Region

#Region " Actualizar Stock "

    <Task()> Public Shared Sub DetalleActualizacionStocks(ByVal data As Object, ByVal services As ServiceProvider)
        Dim AppParamsAlb As ParametroAlbaranVenta = services.GetService(Of ParametroAlbaranVenta)()
        If Not AppParamsAlb.ActualizacionAutomaticaStock Then Exit Sub

        Dim alog As AlbaranLogProcess = services.GetService(Of AlbaranLogProcess)()
        System.Runtime.Remoting.Messaging.CallContext.SetData(GetType(AlbaranLogProcess).Name, alog)
    End Sub

    <Task()> Public Shared Sub ActualizacionAutomaticaStock(ByVal Doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        Dim actualizar As Boolean
        Dim AppParamsAlb As ParametroAlbaranVenta = services.GetService(Of ParametroAlbaranVenta)()
        If Not AppParamsAlb Is Nothing Then
            If Not Doc.HeaderRow.IsNull("IDTipoAlbaran") AndAlso Doc.HeaderRow("IDTipoAlbaran").ToString() = AppParamsAlb.TipoAlbaranDeIntercambio Then
                actualizar = False
            Else
                actualizar = AppParamsAlb.ActualizacionAutomaticaStock
            End If
        End If
        If actualizar Then
            '//Necesitamos recuperar el Doc de nuevo por el Estado de la Cabecera del Albarán. El UpdateDocument hace que se mantengan los estados de los registros. 
            Doc = New DocumentoAlbaranVenta(Doc.HeaderRow("IDAlbaran"))

            '//Terminamos de grabar el AV antes de empezar con la actualización de Stocks
            AdminData.CommitTx(True)

            '//Actualización de Stock de Lineas de Albarén
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
        End If
    End Sub
    <Task()> Public Shared Sub ActualizacionAutomaticaStock2(ByVal Doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        Dim actualizar As Boolean
        Dim AppParamsAlb As ParametroAlbaranVenta = services.GetService(Of ParametroAlbaranVenta)()
        If Not AppParamsAlb Is Nothing Then
            If Not Doc.HeaderRow.IsNull("IDTipoAlbaran") AndAlso Doc.HeaderRow("IDTipoAlbaran").ToString() = AppParamsAlb.TipoAlbaranDeIntercambio Then
                actualizar = False
            Else
                actualizar = AppParamsAlb.ActualizacionAutomaticaStock
            End If
        End If
        If actualizar Then
            '//Necesitamos recuperar el Doc de nuevo por el Estado de la Cabecera del Albarán. El UpdateDocument hace que se mantengan los estados de los registros. 
            Doc = New DocumentoAlbaranVenta(Doc.HeaderRow("IDAlbaran"))

            '//Terminamos de grabar el AV antes de empezar con la actualización de Stocks
            AdminData.CommitTx(True)

            '//Actualización de Stock de Lineas de Albarén
            Dim ActStock As New ProcesoStocks.DataActualizarStockLineas(Doc)
            Dim stockUD() As StockUpdateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarStockLineas, StockUpdateData())(AddressOf ActualizarStockLineas2, ActStock, services)
            If Not stockUD Is Nothing AndAlso stockUD.Length > 0 Then
                Dim alog As AlbaranLogProcess = services.GetService(Of AlbaranLogProcess)()
                If Not alog Is Nothing Then
                    For Each data As StockUpdateData In stockUD
                        ReDim Preserve alog.StockUpdateData(UBound(alog.StockUpdateData) + 1)
                        alog.StockUpdateData(UBound(alog.StockUpdateData)) = data
                    Next
                End If
            End If
        End If
    End Sub


    <Task()> Public Shared Function ActualizarStockLineas(ByVal data As ProcesoStocks.DataActualizarStockLineas, ByVal services As ServiceProvider) As StockUpdateData()
        Dim stkUptData As StockUpdateData
        Dim updateDataArray(-1) As StockUpdateData
        Dim OperarioGenerico As String = New Parametro().OperarioGenerico()
        Dim dtLinea As DataTable = data.DocumentoAlbaran.dtLineas.Clone
        Dim actStockAlb As New ProcesoStocks.DataActualizarStockAlbaranTx
        Dim f As New Filter(FilterUnionOperator.Or)
        If Not data.IDLineasAlbaran Is Nothing AndAlso data.IDLineasAlbaran.Length > 0 Then
            For Each IDLinea As Integer In data.IDLineasAlbaran
                f.Add(New NumberFilterItem("IDLineaAlbaran", IDLinea))
            Next
        End If
        Dim strWhere As String
        If f.Count > 0 Then
            strWhere = f.Compose(New AdoFilterComposer)
        End If

        Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
        Dim IStockClass As IStockInventarioPermanente
        If AppParams.GestionInventarioPermanente Then
            Dim datIStock As New ProcesoStocks.DataCreateIStockClass(InventariosPermanentes.ENSAMBLADO_INV_PERMANENTE_STOCKS, InventariosPermanentes.CLASE_INV_PERMANENTE_STOCKS)
            IStockClass = ProcessServer.ExecuteTask(Of ProcesoStocks.DataCreateIStockClass, IStockInventarioPermanente)(AddressOf ProcesoStocks.CreateIStockInventarios, datIStock, services)
        End If

        For Each lineaAlbaran As DataRow In data.DocumentoAlbaran.dtLineas.Select(strWhere)
            AdminData.BeginTx()
            If lineaAlbaran(_AVL.EstadoStock) = enumavlEstadoStock.avlNoActualizado Then
                Dim blnActualizarStockLinea As Boolean = True
                actStockAlb.IDCliente = data.DocumentoAlbaran.HeaderRow("IDCliente")
                actStockAlb.IDAlbaran = data.DocumentoAlbaran.HeaderRow("IDAlbaran")
                actStockAlb.NAlbaran = data.DocumentoAlbaran.HeaderRow("NAlbaran")
                actStockAlb.FechaAlbaran = data.DocumentoAlbaran.HeaderRow("FechaAlbaran")
                actStockAlb.NumeroMovimiento = Nz(data.DocumentoAlbaran.HeaderRow("NMovimiento"), 0)
                actStockAlb.IDTipoAlbaran = data.DocumentoAlbaran.HeaderRow("IDTipoAlbaran") & String.Empty
                actStockAlb.IDAlmacenDeposito = data.DocumentoAlbaran.HeaderRow("IDAlmacenDeposito") & String.Empty
                actStockAlb.LineaAlbaran = lineaAlbaran
                actStockAlb.Circuito = Circuito.Ventas
                actStockAlb.LotesLineaAlbaran = CType(data.DocumentoAlbaran, DocumentoAlbaranVenta).dtLote.Clone

                Dim valorTipoAlbaran As enumTipoAlbaran = ProcessServer.ExecuteTask(Of String, enumTipoAlbaran)(AddressOf ProcesoAlbaranVenta.ValidarTipoAlbaran, actStockAlb.IDTipoAlbaran, services)

                '//Actualizar Stock Contenedor
                If Length(lineaAlbaran("IDArticuloContenedor")) > 0 Then
                    Dim uda() As StockUpdateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarStockAlbaranTx, StockUpdateData())(AddressOf ActualizarStockContenedores, actStockAlb, services)
                    If Not uda Is Nothing Then
                        For Each stkud As StockUpdateData In uda
                            If stkud.Estado <> EstadoStock.Actualizado Then
                                blnActualizarStockLinea = False
                            End If
                        Next
                    End If
                    ArrayManager.Copy(uda, updateDataArray)
                End If
                If blnActualizarStockLinea Then
                    '//Linea NORMAL, o de tipo KIT, o SUBCONTRATACION MANUAL(que NO proviene de una OF)
                    Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                    Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(lineaAlbaran("IDArticulo"))
                    If ArtInfo.GestionStock Then
                        If ArtInfo.GestionStockPorLotes Then  ''//CON GESTION POR LOTES
                            '//Obtenemos los Lotes de la línea

                            Dim fLinAlb As New Filter
                            fLinAlb.Add(New NumberFilterItem(_AVLT.IDLineaAlbaran, lineaAlbaran(_AVL.IDLineaAlbaran)))
                            Dim WhereLineaAlbaran As String = fLinAlb.Compose(New AdoFilterComposer)
                            For Each lineaLote As DataRow In CType(data.DocumentoAlbaran, DocumentoAlbaranVenta).dtLote.Select(WhereLineaAlbaran)
                                actStockAlb.LotesLineaAlbaran.ImportRow(lineaLote)
                            Next

                            If actStockAlb.LotesLineaAlbaran.Rows.Count > 0 Then
                                Dim uda() As StockUpdateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarStockAlbaranTx, StockUpdateData())(AddressOf ProcesoStocks.ActualizarStockAlbaranTx, actStockAlb, services)
                                If Not uda Is Nothing AndAlso uda.Length > 0 AndAlso uda(0).NumeroMovimiento > 0 AndAlso uda(0).Estado = EstadoStock.Actualizado Then
                                    If Nz(data.DocumentoAlbaran.HeaderRow("NMovimiento"), 0) = 0 Then data.DocumentoAlbaran.HeaderRow("NMovimiento") = uda(0).NumeroMovimiento
                                End If
                                If Not uda Is Nothing AndAlso uda.Length > 0 Then stkUptData = uda(0)
                                ArrayManager.Copy(uda, updateDataArray)
                            Else
                                Dim das As New ProcesoStocks.DataLogActualizarStock("El lote es obligatorio.", lineaAlbaran("IDArticulo"), lineaAlbaran("IDAlmacen"))
                                Dim uda As StockUpdateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataLogActualizarStock, StockUpdateData)(AddressOf ProcesoStocks.LogActualizarStock, das, services)
                                If Not uda Is Nothing Then stkUptData = uda
                                ArrayManager.Copy(uda, updateDataArray)
                            End If
                        Else ''//SIN GESTION POR LOTES
                            Dim uda() As StockUpdateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarStockAlbaranTx, StockUpdateData())(AddressOf ProcesoStocks.ActualizarStockAlbaranTx, actStockAlb, services)
                            If Not uda Is Nothing AndAlso uda.Length > 0 AndAlso uda(0).NumeroMovimiento > 0 AndAlso uda(0).Estado = EstadoStock.Actualizado Then
                                If Nz(data.DocumentoAlbaran.HeaderRow("NMovimiento"), 0) = 0 Then data.DocumentoAlbaran.HeaderRow("NMovimiento") = uda(0).NumeroMovimiento
                            End If
                            If Not uda Is Nothing AndAlso uda.Length > 0 Then stkUptData = uda(0)
                            ArrayManager.Copy(uda, updateDataArray)
                        End If
                    End If

                    dtLinea.ImportRow(lineaAlbaran)
                    BusinessHelper.UpdateTable(data.DocumentoAlbaran.HeaderRow.Table)
                    BusinessHelper.UpdateTable(dtLinea)
                    BusinessHelper.UpdateTable(actStockAlb.LotesLineaAlbaran)

                    If valorTipoAlbaran = enumTipoAlbaran.Intercambio AndAlso lineaAlbaran("EstadoStock") = enumavlEstadoStock.avlActualizado Then
                        Dim Pedidos As DocumentInfoCache(Of DocumentoPedidoVenta) = services.GetService(Of DocumentInfoCache(Of DocumentoPedidoVenta))()
                        If Pedidos.Keys.Count > 0 Then
                            BusinessHelper.UpdateTable(Pedidos.GetDocument(lineaAlbaran("IDPedido")).dtLineas)
                        End If
                    End If
                End If
            End If
            AdminData.CommitTx(True)

            '//Una vez actualizada la linea, vemos si hay que contabilizarla
            If AppParams.GestionInventarioPermanente Then
                '//Si es un Albarán de Trasferencia no gestionamos el inventario permanente.
                Dim ParamsAV As ParametroAlbaranVenta = services.GetService(Of ParametroAlbaranVenta)()
                If data.DocumentoAlbaran.HeaderRow("IDTipoAlbaran") = ParamsAV.TipoAlbaranDeDeposito Then Exit Function

                If lineaAlbaran(_AVL.EstadoStock) = enumavlEstadoStock.avlActualizado Then
                    If Not IStockClass Is Nothing Then
                        If (lineaAlbaran.RowState = DataRowState.Modified AndAlso _
                           (Nz(lineaAlbaran(_AVL.EstadoStock), -1) <> Nz(lineaAlbaran(_AVL.EstadoStock, DataRowVersion.Original), -1)) OrElse _
                            Nz(lineaAlbaran("Contabilizado"), enumContabilizado.NoContabilizado) = enumContabilizado.NoContabilizado) AndAlso _
                            lineaAlbaran("EstadoFactura") = enumavlEstadoFactura.avlNoFacturado AndAlso _
                            Nz(lineaAlbaran("EstadoFactura"), -1) = Nz(lineaAlbaran("EstadoFactura", DataRowVersion.Original), -1) Then
                            Try
                                IStockClass.SincronizarContaAlbaranVenta(lineaAlbaran("IDLineaAlbaran"), lineaAlbaran("Contabilizado"), services)
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

    'David Velasco 10/8/22
    <Task()> Public Shared Function ActualizarStockLineas2(ByVal data As ProcesoStocks.DataActualizarStockLineas, ByVal services As ServiceProvider) As StockUpdateData()
        Dim stkUptData As StockUpdateData
        Dim updateDataArray(-1) As StockUpdateData
        Dim OperarioGenerico As String = New Parametro().OperarioGenerico()
        Dim dtLinea As DataTable = data.DocumentoAlbaran.dtLineas.Clone
        Dim actStockAlb As New ProcesoStocks.DataActualizarStockAlbaranTx
        Dim f As New Filter(FilterUnionOperator.Or)
        If Not data.IDLineasAlbaran Is Nothing AndAlso data.IDLineasAlbaran.Length > 0 Then
            For Each IDLinea As Integer In data.IDLineasAlbaran
                f.Add(New NumberFilterItem("IDLineaAlbaran", IDLinea))
            Next
        End If
        Dim strWhere As String
        If f.Count > 0 Then
            strWhere = f.Compose(New AdoFilterComposer)
        End If

        Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
        Dim IStockClass As IStockInventarioPermanente
        If AppParams.GestionInventarioPermanente Then
            Dim datIStock As New ProcesoStocks.DataCreateIStockClass(InventariosPermanentes.ENSAMBLADO_INV_PERMANENTE_STOCKS, InventariosPermanentes.CLASE_INV_PERMANENTE_STOCKS)
            IStockClass = ProcessServer.ExecuteTask(Of ProcesoStocks.DataCreateIStockClass, IStockInventarioPermanente)(AddressOf ProcesoStocks.CreateIStockInventarios, datIStock, services)
        End If

        For Each lineaAlbaran As DataRow In data.DocumentoAlbaran.dtLineas.Select(strWhere)
            AdminData.BeginTx()
            If lineaAlbaran(_AVL.EstadoStock) = enumavlEstadoStock.avlNoActualizado Then
                Dim blnActualizarStockLinea As Boolean = True
                actStockAlb.IDCliente = data.DocumentoAlbaran.HeaderRow("IDCliente")
                actStockAlb.IDAlbaran = data.DocumentoAlbaran.HeaderRow("IDAlbaran")
                actStockAlb.NAlbaran = data.DocumentoAlbaran.HeaderRow("NAlbaran")
                actStockAlb.FechaAlbaran = data.DocumentoAlbaran.HeaderRow("FechaAlbaran")
                actStockAlb.NumeroMovimiento = Nz(data.DocumentoAlbaran.HeaderRow("NMovimiento"), 0)
                actStockAlb.IDTipoAlbaran = data.DocumentoAlbaran.HeaderRow("IDTipoAlbaran") & String.Empty
                actStockAlb.IDAlmacenDeposito = data.DocumentoAlbaran.HeaderRow("IDAlmacenDeposito") & String.Empty
                actStockAlb.LineaAlbaran = lineaAlbaran
                actStockAlb.Circuito = Circuito.Ventas
                actStockAlb.LotesLineaAlbaran = CType(data.DocumentoAlbaran, DocumentoAlbaranVenta).dtLote.Clone

                Dim valorTipoAlbaran As enumTipoAlbaran = ProcessServer.ExecuteTask(Of String, enumTipoAlbaran)(AddressOf ProcesoAlbaranVenta.ValidarTipoAlbaran, actStockAlb.IDTipoAlbaran, services)

                '//Actualizar Stock Contenedor
                If Length(lineaAlbaran("IDArticuloContenedor")) > 0 Then
                    Dim uda() As StockUpdateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarStockAlbaranTx, StockUpdateData())(AddressOf ActualizarStockContenedores, actStockAlb, services)
                    If Not uda Is Nothing Then
                        For Each stkud As StockUpdateData In uda
                            If stkud.Estado <> EstadoStock.Actualizado Then
                                blnActualizarStockLinea = False
                            End If
                        Next
                    End If
                    ArrayManager.Copy(uda, updateDataArray)
                End If
                If blnActualizarStockLinea Then
                    '//Linea NORMAL, o de tipo KIT, o SUBCONTRATACION MANUAL(que NO proviene de una OF)
                    Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                    Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(lineaAlbaran("IDArticulo"))
                    If ArtInfo.GestionStock Then
                        If ArtInfo.GestionStockPorLotes Then  ''//CON GESTION POR LOTES
                            '//Obtenemos los Lotes de la línea

                            Dim fLinAlb As New Filter
                            fLinAlb.Add(New NumberFilterItem(_AVLT.IDLineaAlbaran, lineaAlbaran(_AVL.IDLineaAlbaran)))
                            Dim WhereLineaAlbaran As String = fLinAlb.Compose(New AdoFilterComposer)
                            For Each lineaLote As DataRow In CType(data.DocumentoAlbaran, DocumentoAlbaranVenta).dtLote.Select(WhereLineaAlbaran)
                                actStockAlb.LotesLineaAlbaran.ImportRow(lineaLote)
                            Next

                            If actStockAlb.LotesLineaAlbaran.Rows.Count > 0 Then
                                Dim uda() As StockUpdateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarStockAlbaranTx, StockUpdateData())(AddressOf ProcesoStocks.ActualizarStockAlbaranTx, actStockAlb, services)
                                If Not uda Is Nothing AndAlso uda.Length > 0 AndAlso uda(0).NumeroMovimiento > 0 AndAlso uda(0).Estado = EstadoStock.Actualizado Then
                                    If Nz(data.DocumentoAlbaran.HeaderRow("NMovimiento"), 0) = 0 Then data.DocumentoAlbaran.HeaderRow("NMovimiento") = uda(0).NumeroMovimiento
                                End If
                                If Not uda Is Nothing AndAlso uda.Length > 0 Then stkUptData = uda(0)
                                ArrayManager.Copy(uda, updateDataArray)
                            Else
                                Dim das As New ProcesoStocks.DataLogActualizarStock("El lote es obligatorio.", lineaAlbaran("IDArticulo"), lineaAlbaran("IDAlmacen"))
                                Dim uda As StockUpdateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataLogActualizarStock, StockUpdateData)(AddressOf ProcesoStocks.LogActualizarStock, das, services)
                                If Not uda Is Nothing Then stkUptData = uda
                                ArrayManager.Copy(uda, updateDataArray)
                            End If
                        Else ''//SIN GESTION POR LOTES
                            Dim uda() As StockUpdateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarStockAlbaranTx, StockUpdateData())(AddressOf ProcesoStocks.ActualizarStockAlbaranTx2, actStockAlb, services)
                            If Not uda Is Nothing AndAlso uda.Length > 0 AndAlso uda(0).NumeroMovimiento > 0 AndAlso uda(0).Estado = EstadoStock.Actualizado Then
                                If Nz(data.DocumentoAlbaran.HeaderRow("NMovimiento"), 0) = 0 Then data.DocumentoAlbaran.HeaderRow("NMovimiento") = uda(0).NumeroMovimiento
                            End If
                            If Not uda Is Nothing AndAlso uda.Length > 0 Then stkUptData = uda(0)
                            ArrayManager.Copy(uda, updateDataArray)
                        End If
                    End If

                    dtLinea.ImportRow(lineaAlbaran)
                    BusinessHelper.UpdateTable(data.DocumentoAlbaran.HeaderRow.Table)
                    BusinessHelper.UpdateTable(dtLinea)
                    BusinessHelper.UpdateTable(actStockAlb.LotesLineaAlbaran)

                    If valorTipoAlbaran = enumTipoAlbaran.Intercambio AndAlso lineaAlbaran("EstadoStock") = enumavlEstadoStock.avlActualizado Then
                        Dim Pedidos As DocumentInfoCache(Of DocumentoPedidoVenta) = services.GetService(Of DocumentInfoCache(Of DocumentoPedidoVenta))()
                        If Pedidos.Keys.Count > 0 Then
                            BusinessHelper.UpdateTable(Pedidos.GetDocument(lineaAlbaran("IDPedido")).dtLineas)
                        End If
                    End If
                End If
            End If
            AdminData.CommitTx(True)

            '//Una vez actualizada la linea, vemos si hay que contabilizarla
            If AppParams.GestionInventarioPermanente Then
                '//Si es un Albarán de Trasferencia no gestionamos el inventario permanente.
                Dim ParamsAV As ParametroAlbaranVenta = services.GetService(Of ParametroAlbaranVenta)()
                If data.DocumentoAlbaran.HeaderRow("IDTipoAlbaran") = ParamsAV.TipoAlbaranDeDeposito Then Exit Function

                If lineaAlbaran(_AVL.EstadoStock) = enumavlEstadoStock.avlActualizado Then
                    If Not IStockClass Is Nothing Then
                        If (lineaAlbaran.RowState = DataRowState.Modified AndAlso _
                           (Nz(lineaAlbaran(_AVL.EstadoStock), -1) <> Nz(lineaAlbaran(_AVL.EstadoStock, DataRowVersion.Original), -1)) OrElse _
                            Nz(lineaAlbaran("Contabilizado"), enumContabilizado.NoContabilizado) = enumContabilizado.NoContabilizado) AndAlso _
                            lineaAlbaran("EstadoFactura") = enumavlEstadoFactura.avlNoFacturado AndAlso _
                            Nz(lineaAlbaran("EstadoFactura"), -1) = Nz(lineaAlbaran("EstadoFactura", DataRowVersion.Original), -1) Then
                            Try
                                IStockClass.SincronizarContaAlbaranVenta(lineaAlbaran("IDLineaAlbaran"), lineaAlbaran("Contabilizado"), services)
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


    <Task()> Public Shared Function ActualizarStockContenedores(ByVal data As ProcesoStocks.DataActualizarStockAlbaranTx, ByVal services As ServiceProvider) As StockUpdateData()
        Dim updateDataArray(-1) As StockUpdateData

        Dim updateSalidaContenedor As StockUpdateData
        Dim updateEntradaContenedor As StockUpdateData

        If Length(data.LineaAlbaran(_AVL.IDArticuloContenedor)) > 0 Then
            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
            Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.LineaAlbaran(_AVL.IDArticuloContenedor))
            If ArtInfo.GestionStock Then
                If data.NumeroMovimiento Is Nothing OrElse data.NumeroMovimiento = 0 Then
                    data.NumeroMovimiento = ProcessServer.ExecuteTask(Of Object, Integer)(AddressOf ProcesoStocks.NuevoNumeroMovimiento, Nothing, services)
                End If
                Dim Fecha As Date = data.FechaAlbaran
                Dim ArticuloContenedor As String = data.LineaAlbaran(_AVL.IDArticuloContenedor)
                Dim AlmacenPredeterminado As String

                Dim AlmacenContenedor As String
                Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
                Dim ClteInfo As ClienteInfo = Clientes.GetEntity(data.IDCliente)
                If Not ClteInfo Is Nothing AndAlso Length(ClteInfo.IDCliente) > 0 Then
                    If Length(ClteInfo.IDAlmacenContenedor) > 0 Then
                        AlmacenContenedor = ClteInfo.IDAlmacenContenedor
                    End If
                End If
                If Len(AlmacenContenedor) = 0 Then
                    Dim das As New ProcesoStocks.DataLogActualizarStock("El cliente " & Quoted(data.IDCliente) & " no tiene asignado ningún almacén contenedor.", ArticuloContenedor)
                    Dim stkud As StockUpdateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataLogActualizarStock, StockUpdateData)(AddressOf ProcesoStocks.LogActualizarStock, das, services)
                    If Not stkud Is Nothing Then
                        ReDim Preserve updateDataArray(updateDataArray.Length)
                        updateDataArray(updateDataArray.Length - 1) = stkud
                    End If
                    Return updateDataArray
                Else
                    Dim QContenedor As Double = Nz(data.LineaAlbaran(_AVL.QEtiContenedor), 0)
                    Dim f As New Filter
                    f.Add(New StringFilterItem(_AA.IDArticulo, FilterOperator.Equal, ArticuloContenedor))
                    f.Add(New BooleanFilterItem(_AA.Predeterminado, FilterOperator.Equal, True))
                    Dim almacen As DataTable = New Negocio.ArticuloAlmacen().Filter(f)
                    If Not almacen Is Nothing AndAlso almacen.Rows.Count > 0 Then
                        AlmacenPredeterminado = almacen.Rows(0)(_AA.IDAlmacen)
                    End If
                    If Len(AlmacenPredeterminado) = 0 Then
                        Dim das As New ProcesoStocks.DataLogActualizarStock("El articulo " & Quoted(ArticuloContenedor) & " no tiene asignado ningún almacén predeterminado.", ArticuloContenedor)
                        Dim stkud As StockUpdateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataLogActualizarStock, StockUpdateData)(AddressOf ProcesoStocks.LogActualizarStock, das, services)
                        If Not stkud Is Nothing Then
                            ReDim Preserve updateDataArray(updateDataArray.Length)
                            updateDataArray(updateDataArray.Length - 1) = stkud
                        End If

                        Return updateDataArray
                    Else
                        ProcessServer.ExecuteTask(Of Object)(AddressOf Business.General.Comunes.BeginTransaction, Nothing, services)
                        '//Salida de contenedor
                        If Length(data.LineaAlbaran("IDSalidaContenedor")) = 0 Then
                            Dim salidaContenedor As New StockData
                            salidaContenedor.TipoMovimiento = enumTipoMovimiento.tmSalTransferencia
                            salidaContenedor.Articulo = ArticuloContenedor
                            salidaContenedor.Almacen = AlmacenPredeterminado
                            salidaContenedor.Cantidad = QContenedor
                            salidaContenedor.Documento = data.NAlbaran
                            salidaContenedor.Texto = data.LineaAlbaran(_AVL.Texto) & String.Empty
                            salidaContenedor.FechaDocumento = Fecha
                            If IsNumeric(data.LineaAlbaran(_AVL.IDObra)) Then
                                salidaContenedor.Obra = data.LineaAlbaran(_AVL.IDObra)
                            End If

                            Dim datosSalida As New DataNumeroMovimientoSinc(data.NumeroMovimiento, salidaContenedor)
                            updateSalidaContenedor = ProcessServer.ExecuteTask(Of DataNumeroMovimientoSinc, StockUpdateData)(AddressOf ProcesoStocks.Salida, datosSalida, services)

                            If updateSalidaContenedor.Estado = EstadoStock.Actualizado Then
                                '//Entrada en el almacen del cliente
                                Dim entradaContenedor As New StockData
                                entradaContenedor.TipoMovimiento = enumTipoMovimiento.tmEntTransferencia
                                entradaContenedor.Articulo = ArticuloContenedor
                                entradaContenedor.Almacen = AlmacenContenedor
                                entradaContenedor.Cantidad = QContenedor
                                entradaContenedor.Documento = data.NAlbaran
                                entradaContenedor.Texto = data.LineaAlbaran(_AVL.Texto) & String.Empty
                                entradaContenedor.FechaDocumento = Fecha
                                If IsNumeric(data.LineaAlbaran(_AVL.IDObra)) Then
                                    entradaContenedor.Obra = data.LineaAlbaran(_AVL.IDObra)
                                End If

                                Dim datosEntrada As New DataNumeroMovimientoSinc(data.NumeroMovimiento, entradaContenedor)
                                updateEntradaContenedor = ProcessServer.ExecuteTask(Of DataNumeroMovimientoSinc, StockUpdateData)(AddressOf ProcesoStocks.Entrada, datosEntrada, services)

                                If updateEntradaContenedor.Estado = EstadoStock.Actualizado Then
                                    data.LineaAlbaran(_AVL.IDEntradaContenedor) = updateEntradaContenedor.IDLineaMovimiento
                                    data.LineaAlbaran(_AVL.IDSalidaContenedor) = updateSalidaContenedor.IDLineaMovimiento
                                Else
                                    ProcessServer.ExecuteTask(Of Boolean)(AddressOf Business.General.Comunes.RollbackTransaction, True, services)
                                    data.LineaAlbaran(_AVL.EstadoStock) = enumaclEstadoStock.aclNoActualizado
                                End If
                                '///Fin entrada en el almacen del cliente
                            Else
                                data.LineaAlbaran(_AVL.EstadoStock) = enumaclEstadoStock.aclNoActualizado
                            End If
                            '//Fin Salida de contenedor
                            If Not updateSalidaContenedor Is Nothing Then
                                ReDim Preserve updateDataArray(updateDataArray.Length)
                                updateDataArray(updateDataArray.Length - 1) = updateSalidaContenedor
                            End If

                            If Not updateEntradaContenedor Is Nothing Then
                                ReDim Preserve updateDataArray(updateDataArray.Length)
                                updateDataArray(updateDataArray.Length - 1) = updateEntradaContenedor
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Return updateDataArray
    End Function

    <Task()> Public Shared Sub AsignarAlbaranVentaLotes(ByVal docAlbaran As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        If Not docAlbaran.dtLineas Is Nothing Then
            Dim lineas As DataTable = docAlbaran.dtLineas
            If lineas Is Nothing Then
                Dim oAVL As New AlbaranVentaLinea
                lineas = oAVL.AddNew
                docAlbaran.Add(GetType(AlbaranVentaLinea).Name, lineas)
            End If

            Dim lotes As DataTable = docAlbaran.dtLote
            If lotes Is Nothing Then
                Dim oAVLo As New AlbaranVentaLote
                lotes = oAVLo.AddNew
                docAlbaran.Add(GetType(AlbaranVentaLote).Name, lotes)
            End If
            Dim BlnSeguimiento As Boolean = New Parametro().BodegaSeguimiento
            Dim Seguimiento As DataTable = docAlbaran.dtSeguimiento
            If BlnSeguimiento Then
                If Seguimiento Is Nothing Then
                    Dim oAVSeg As New AlbaranVentaSeguimiento
                    Seguimiento = oAVSeg.AddNew
                    docAlbaran.Add(GetType(AlbaranVentaSeguimiento).Name, Seguimiento)
                End If
            End If

            For Each drACL As DataRow In docAlbaran.dtLineas.Rows
                Dim alblin As AlbLinVenta = Nothing
                For i As Integer = 0 To docAlbaran.Cabecera.LineasOrigen.Length - 1
                    If drACL(docAlbaran.Cabecera.LineasOrigen(i).PrimaryKeyLinOrigen) = docAlbaran.Cabecera.LineasOrigen(i).IDLineaOrigen Then
                        alblin = docAlbaran.Cabecera.LineasOrigen(i)
                        Exit For
                    End If
                Next

                If Not alblin Is Nothing AndAlso Not alblin.Lotes Is Nothing AndAlso alblin.Lotes.Rows.Count > 0 Then
                    Dim SegundaUnidad As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, drACL("IDArticulo"), services)

                    Dim datIStock As New ProcesoStocks.DataCreateIStockClass("Expertis.Business.Bodega.dll", "Solmicro.Expertis.Business.Bodega.BdgStock")
                    Dim IStockClass As IStock = ProcessServer.ExecuteTask(Of ProcesoStocks.DataCreateIStockClass, IStock)(AddressOf ProcesoStocks.CreateIStockClass, datIStock, services)
                    For Each lt As DataRow In alblin.Lotes.Rows
                        If lt(_AAL.Cantidad) <> 0 Then
                            Dim lote As DataRow = lotes.NewRow

                            lote("IDLineaLote") = AdminData.GetAutoNumeric
                            lote("IDLineaAlbaran") = drACL("IDLineaAlbaran")
                            lote("Lote") = lt("Lote")
                            lote("Ubicacion") = lt("Ubicacion")
                            lote("QInterna") = lt("Cantidad")
                            If SegundaUnidad Then lote("QInterna2") = lt("Cantidad2")
                            'TODO VER como rellenar observaciones (añadir campo en AlbaranLineaLote)
                            'lote("Observaciones") = lt("Observaciones")
                            'lote("FechaCaducidad") = lt("FechaCaducidad")
                            If Not alblin.ArtCompatibles Is Nothing AndAlso Length(alblin.ArtCompatibles.IDDeposito) > 0 AndAlso Length(alblin.ArtCompatibles.IDTipoOperacion) > 0 Then
                                If Not IStockClass Is Nothing Then
                                    Dim datLote As New DataPrepararArtCompatiblesLote(alblin.ArtCompatibles, drACL("IDArticulo"), drACL("IDAlmacen"), lote("Lote"), lote("Ubicacion"))
                                    Dim LotesComp As DataArtCompatiblesExp = ProcessServer.ExecuteTask(Of DataPrepararArtCompatiblesLote, DataArtCompatiblesExp)(AddressOf PrepararArticulosCompatiblesLote, datLote, services)
                                    If Not LotesComp Is Nothing Then
                                        Dim datCrearOp As New DataArtCompatiblesExp(alblin.ArtCompatibles.IDLineaPedido, LotesComp.dtArtCompatibles, alblin.ArtCompatibles.IDDeposito, alblin.ArtCompatibles.IDTipoOperacion, Nz(docAlbaran.HeaderRow("FechaAlbaran"), Now))
                                        lote("IDOperacion") = IStockClass.CrearOperacionArticulosCompatibles(datCrearOp)
                                    End If
                                End If
                            End If
                            lotes.Rows.Add(lote.ItemArray)

                            If BlnSeguimiento AndAlso (Not alblin.Seguimiento Is Nothing AndAlso alblin.Seguimiento.Rows.Count > 0) Then
                                For Each Seg As DataRow In alblin.Seguimiento.Select("Lote = '" & lote("Lote") & "'")
                                    Dim SegNew As DataRow = Seguimiento.NewRow
                                    SegNew("IDLineaSeguimiento") = AdminData.GetAutoNumeric
                                    SegNew("IDLineaLote") = lote("IDLineaLote")
                                    SegNew("NDesde") = Seg("NDesde")
                                    SegNew("NHasta") = Seg("NHasta")
                                    SegNew("NPallet") = Seg("NPallet")
                                    SegNew("Cantidad") = Seg("Cantidad")
                                    Seguimiento.Rows.Add(SegNew)
                                Next
                            End If
                        End If
                    Next
                End If
            Next
        End If
    End Sub


    <Serializable()> _
    Public Class DataPrepararArtCompatiblesLote
        Public ArtCompatibles As DataArtCompatiblesExp
        Public IDArticulo As String
        Public IDAlmacen As String
        Public Lote As String
        Public Ubicacion As String

        Public Sub New(ByVal ArtCompatibles As DataArtCompatiblesExp, ByVal IDArticulo As String, ByVal IDAlmacen As String, ByVal Lote As String, ByVal Ubicacion As String)
            Me.ArtCompatibles = ArtCompatibles
            Me.IDArticulo = IDArticulo
            Me.IDAlmacen = IDAlmacen
            Me.Lote = Lote
            Me.Ubicacion = Ubicacion
        End Sub
    End Class
    <Task()> Public Shared Function PrepararArticulosCompatiblesLote(ByVal data As DataPrepararArtCompatiblesLote, ByVal services As ServiceProvider) As DataArtCompatiblesExp
        If data.ArtCompatibles Is Nothing Then Exit Function

        Dim dtArtCompatiblesLote As DataTable
        If Not data.ArtCompatibles.dtArtCompatibles Is Nothing AndAlso data.ArtCompatibles.dtArtCompatibles.Rows.Count > 0 Then
            Dim ArtCompatiblesLote As List(Of DataRow) = (From c In data.ArtCompatibles.dtArtCompatibles _
                                                            Where Not c.IsNull("Lote") AndAlso c("Lote") = data.Lote AndAlso _
                                                                  Not c.IsNull("Ubicacion") AndAlso c("Ubicacion") = data.Ubicacion).ToList()
            If Not ArtCompatiblesLote Is Nothing AndAlso ArtCompatiblesLote.Count > 0 Then
                dtArtCompatiblesLote = ArtCompatiblesLote.CopyToDataTable
            End If
        End If

        Return New DataArtCompatiblesExp(data.ArtCompatibles.IDLineaPedido, dtArtCompatiblesLote)
    End Function

    <Serializable()> _
    Public Class DataAddMntoOT
        Public IDEstadoActivo As String
        Public IDCliente As String
        Public LineaAlbaran As DataRow

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDEstadoActivo As String, ByVal IDCliente As String, ByVal LineaAlbaran As DataRow)
            Me.IDEstadoActivo = IDEstadoActivo
            Me.IDCliente = IDCliente
            Me.LineaAlbaran = LineaAlbaran
        End Sub
    End Class
    <Task()> Public Shared Sub AddMntoOT(ByVal data As DataAddMntoOT, ByVal services As ServiceProvider)
        Dim ofc As BusinessHelper
        ofc = BusinessHelper.CreateBusinessObject("MntoEstadoActivo")
        Dim dtEA As DataTable = ofc.SelOnPrimaryKey(data.IDEstadoActivo)
        If Not IsNothing(dtEA) AndAlso dtEA.Rows.Count > 0 Then
            If Nz(dtEA.Rows(0)("GeneraOT"), False) Then
                Dim dtOT As New DataTable
                dtOT.Columns.Add("FechaSolicitud", GetType(Date))
                dtOT.Columns.Add("HoraSolicitud", GetType(Date))
                dtOT.Columns.Add("IDActivo", GetType(String))
                dtOT.Columns.Add("IDCliente", GetType(String))
                dtOT.Columns.Add("IDObra", GetType(Integer))
                dtOT.Columns.Add("IDTrabajo", GetType(Integer))
                dtOT.Columns.Add("IDAlbaranRetorno", GetType(Integer))
                dtOT.Columns.Add("IDLineaAlbaranRetorno", GetType(Integer))

                Dim drOT As DataRow = dtOT.NewRow

                drOT("FechaSolicitud") = Nz(data.LineaAlbaran(_AVL.FechaAlquiler), Date.Today)
                drOT("HoraSolicitud") = Nz(data.LineaAlbaran(_AVL.HoraAlquiler), Date.Today)
                drOT("IDActivo") = data.LineaAlbaran(_AVL.Lote) & String.Empty
                drOT("IDCliente") = data.IDCliente
                If Length(data.LineaAlbaran(_AVL.IDObra)) > 0 Then drOT("IDObra") = data.LineaAlbaran(_AVL.IDObra)
                If Length(data.LineaAlbaran(_AVL.IDTrabajo)) > 0 Then drOT("IDTrabajo") = data.LineaAlbaran(_AVL.IDTrabajo)
                If Length(data.LineaAlbaran(_AVL.IDAlbaran)) > 0 Then drOT("IDAlbaranRetorno") = data.LineaAlbaran(_AVL.IDAlbaran)
                If Length(data.LineaAlbaran(_AVL.IDLineaAlbaran)) > 0 Then drOT("IDLineaAlbaranRetorno") = data.LineaAlbaran(_AVL.IDLineaAlbaran)

                dtOT.Rows.Add(drOT)

                ofc = BusinessHelper.CreateBusinessObject("MntoOT")
                ProcessServer.ExecuteTask(Of DataTable)(AddressOf CType(ofc, IControlOT).GenerarNuevaOTDEsdeRetornos, dtOT, services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambiarNDocumentoMovimientos(ByVal Doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        If Length(Doc.HeaderRow("IDTPV")) > 0 AndAlso Doc.HeaderRow.RowState = DataRowState.Modified Then
            If Doc.HeaderRow("NAlbaran") & String.Empty <> Doc.HeaderRow("NAlbaran", DataRowVersion.Original) & String.Empty Then

                Dim IDLineaMovimiento As List(Of Object) = (From c In Doc.dtLineas Where Not c.IsNull("IDMovimiento") Select c("IDMovimiento") Distinct).ToList
                If IDLineaMovimiento Is Nothing OrElse IDLineaMovimiento.Count = 0 Then Exit Sub

                Dim fMovto As New Filter
                fMovto.Add(New InListFilterItem("IDLineaMovimiento", IDLineaMovimiento.ToArray, FilterType.Numeric))
                fMovto.Add(New NumberFilterItem("IDDocumento", Doc.HeaderRow("IDAlbaran")))
                Dim fTipoMovimiento As New Filter(FilterUnionOperator.Or)
                fTipoMovimiento.Add(New NumberFilterItem("IDTipoMovimiento", enumTipoMovimiento.tmCorreccion))
                fTipoMovimiento.Add(New NumberFilterItem("IDTipoMovimiento", enumTipoMovimiento.tmSalAlbaranVenta))
                fMovto.Add(fTipoMovimiento)
                Dim dtHistoricoMovto As DataTable = New BE.DataEngine().Filter("tbHistoricoMovimiento", fMovto)
                dtHistoricoMovto.TableName = "Stock"
                If dtHistoricoMovto.Rows.Count > 0 Then
                    For Each dr As DataRow In dtHistoricoMovto.Select("Documento <> " & Quoted(Doc.HeaderRow("NAlbaran")))
                        dr("Documento") = Doc.HeaderRow("NAlbaran")
                    Next
                End If
                BusinessHelper.UpdateTable(dtHistoricoMovto)

            End If
        End If
    End Sub

#End Region

#Region " Delete "

    <Task()> Public Shared Sub ValidarAlbaranArqueoCaja(ByVal DocHeaderRow As DataRow, ByVal services As ServiceProvider)
        If Length(DocHeaderRow("Arqueo")) > 0 AndAlso DocHeaderRow("Arqueo") Then
            ApplicationService.GenerateError("No se permite eliminar un albarán Arqueado.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarAlbaranContado(ByVal DocHeaderRow As DataRow, ByVal services As ServiceProvider)
        Dim ff As New Filter
        ff.Add("IDAlbaran", FilterOperator.Equal, DocHeaderRow("IDAlbaran"))
        Dim dt As DataTable = AdminData.GetData("tbAlbaranVentaFormaPago", ff)
        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then ApplicationService.GenerateError("No se puede borrar el Albarán. Hay datos relacionados con la forma de pago.")
    End Sub
    <Task()> Public Shared Sub ActualizarAlbaranesMultiEmpresa(ByVal DocHeaderRow As DataRow, ByVal services As ServiceProvider)
        Dim AppParamsVenta As ParametroVenta = services.GetService(Of ParametroVenta)()
        If Not AppParamsVenta.ComprasEmpresasGrupo Then Exit Sub

        Dim IDAlbaran As Integer = DocHeaderRow("IDAlbaran")

        '//control albaranes multiempresa
        Dim grp As New GRPAlbaranVentaCompraLinea
        Dim control As DataTable = grp.TrazaAVPrincipal(IDAlbaran)
        If Not control Is Nothing AndAlso control.Rows.Count > 0 Then
            For Each dr As DataRow In control.Rows
                dr("IDAVPrincipal") = DBNull.Value
                dr("NAVPrincipal") = DBNull.Value
                dr("IDLineaAVPrincipal") = DBNull.Value
            Next
            BusinessHelper.UpdateTable(control)
        Else
            control = grp.TrazaAVSecundaria(IDAlbaran)
            If Not control Is Nothing AndAlso control.Rows.Count > 0 Then
                Dim databaseBak As String = AdminData.GetSessionDataBase()
                Try
                    Dim aclEstandar As New AlbaranCompraLinea
                    Dim avlEstandar As New AlbaranVentaLinea
                    For Each dr As DataRow In control.Rows
                        '//borrado de la traza
                        grp.Delete(dr("IDAVLinea"))
                        '//borrado en cascada de la recepcion y expedicion de la BBDD principal
                        AdminData.SetSessionDataBase(dr("IDBDPrincipal"))
                        If Length(dr("IDLineaACPrincipal")) > 0 Then
                            Dim dtACL As DataTable = aclEstandar.SelOnPrimaryKey(dr("IDLineaACPrincipal"))
                            aclEstandar.Delete(dtACL)
                        End If
                        If Length(dr("IDLineaAVPrincipal")) > 0 Then
                            Dim dtAVL As DataTable = avlEstandar.SelOnPrimaryKey(dr("IDLineaAVPrincipal"))
                            avlEstandar.Delete(dtAVL)
                        End If
                    Next

                    Dim accEstandar As New AlbaranCompraCabecera
                    Dim avcEstandar As New AlbaranVentaCabecera
                    Dim IDAlbaranVenta, IDAlbaranCompra As Integer
                    For Each dr As DataRow In control.Rows
                        AdminData.SetSessionDataBase(dr("IDBDPrincipal"))
                        If Length(dr("IDACPrincipal")) > 0 Then
                            If IDAlbaranCompra <> dr("IDACPrincipal") Then
                                IDAlbaranCompra = dr("IDACPrincipal")
                                Dim dtACC As DataTable = accEstandar.SelOnPrimaryKey(dr("IDACPrincipal"))
                                accEstandar.Delete(dtACC)
                            End If
                        End If
                        If Length(dr("IDAVPrincipal")) > 0 Then
                            If IDAlbaranVenta <> dr("IDAVPrincipal") Then
                                IDAlbaranVenta = dr("IDAVPrincipal")
                                Dim dtAVC As DataTable = avcEstandar.SelOnPrimaryKey(dr("IDAVPrincipal"))
                                avcEstandar.Delete(dtAVC)
                            End If
                        End If
                    Next
                Catch ex As Exception
                    AdminData.RollBackTx(True)
                    Throw ex
                Finally
                    AdminData.SetSessionDataBase(databaseBak)
                End Try
            End If
        End If
    End Sub

#End Region

#Region " Update "

    <Task()> Public Shared Sub TratarTipoAlbaran(ByVal Doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        Dim strTipo As String = Doc.HeaderRow("IdTipoAlbaran") & String.Empty

        Dim eTipoAlbaran As enumTipoAlbaran = ProcessServer.ExecuteTask(Of Object, TipoAlbaranInfo)(AddressOf TipoDeAlbaran, strTipo, services).Tipo
        If Length(Doc.HeaderRow("IdTipoAlbaran")) = 0 Then
            If Len(strTipo) > 0 Then Doc.HeaderRow("IdTipoAlbaran") = strTipo
        End If

        If eTipoAlbaran = enumTipoAlbaran.Desconocido Then
            Dim AppParamsTes As ParametroTesoreria = services.GetService(Of ParametroTesoreria)()
            Doc.HeaderRow("IdTipoAlbaran") = AppParamsTes.TipoAlbaranPorDefecto
            If Length(Doc.HeaderRow("IdTipoAlbaran")) = 0 Then
                ApplicationService.GenerateError("El parámetro {0} no existe o no está correctamente configurado.", Quoted(Parametro.cgFwAlbaranDefault))
            End If
            eTipoAlbaran = enumTipoAlbaran.Normal
        End If

        Select Case eTipoAlbaran
            Case enumTipoAlbaran.Servicio
                If Length(Doc.HeaderRow("IDProveedorServicio")) = 0 Then ApplicationService.GenerateError("Falta indicar proveedor de servicios para el albarán de servicios.")
                Doc.HeaderRow("IDAlmacenDeposito") = System.DBNull.Value
            Case enumTipoAlbaran.Deposito, enumTipoAlbaran.Consumo ', avcRetornoAlquiler
                If Length(Doc.HeaderRow("IDAlmacenDeposito")) = 0 Then
                    ApplicationService.GenerateError("El almacén de depósito no existe o está vacío.")
                Else
                    Dim Almacenes As EntityInfoCache(Of AlmacenInfo) = services.GetService(Of EntityInfoCache(Of AlmacenInfo))()
                    Dim AlmInfo As AlmacenInfo = Almacenes.GetEntity(Doc.HeaderRow("IDAlmacenDeposito"))
                    If Not AlmInfo.Deposito Then ApplicationService.GenerateError("El almacén {0} no es de depósito", Quoted(Doc.HeaderRow("IDAlmacenDeposito")))
                End If
                Doc.HeaderRow("IDProveedorServicio") = System.DBNull.Value
        End Select
    End Sub

    <Task()> Public Shared Sub ResponsableExpedicion(ByVal Doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        If Length(Doc.HeaderRow("ResponsableExpedicion")) = 0 Then
            Dim strOperario As String = ProcessServer.ExecuteTask(Of Object, String)(AddressOf Operario.ObtenerIDOperarioUsuario, Nothing, services)
            Dim AppParamsGen As ParametroGeneral = services.GetService(Of ParametroGeneral)()
            If Length(strOperario) = 0 Then strOperario = AppParamsGen.OperarioGenerico
            If Length(strOperario) > 0 Then
                Doc.HeaderRow("ResponsableExpedicion") = strOperario
            Else
                ApplicationService.GenerateError("El parámetro de Operario Genérico no existe o no está configurado correctamente.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarEstadoLineas(ByVal data As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        For Each linea As DataRow In data.dtLineas.Rows
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarEstadoLinea, linea, services)
        Next
    End Sub

    <Task()> Public Shared Sub AsignarEstadoLinea(ByVal linea As DataRow, ByVal services As ServiceProvider)
        Dim QFacturada As Double = Nz(linea("QFacturada"), 0)
        Dim QServida As Double = Nz(linea("QServida"), 0)

        If QFacturada = 0 Then
            linea("EstadoFactura") = enumavlEstadoFactura.avlNoFacturado
        Else
            If Math.Abs(QFacturada) >= Math.Abs(QServida) Then
                linea("EstadoFactura") = enumavlEstadoFactura.avlFacturado
            ElseIf Math.Abs(QFacturada) < Math.Abs(QServida) Then
                linea("EstadoFactura") = enumavlEstadoFactura.avlParcFacturado
            End If
        End If
    End Sub

#End Region

#Region " Componentes "

    <Task()> Public Shared Sub GestionArticulosFantasma(ByVal Doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        '//Gestion de articulo Kit o viruta                
        If Doc Is Nothing Then Exit Sub
        If Not Doc.dtLineas Is Nothing Then
            Dim Componentes As DataTable
            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
            Dim IntPos As Integer = 1
            For Each lineaAlbaran As DataRow In Doc.dtLineas.Select("", "IDOrdenLinea")
                Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(lineaAlbaran("IDArticulo"))
                If ArtInfo.Fantasma AndAlso Not ArtInfo.KitVenta Then
                    Select Case lineaAlbaran.RowState
                        Case DataRowState.Added
                            Componentes = ProcessServer.ExecuteTask(Of DataRow, DataTable)(AddressOf ComponentesDePrimerNivel, lineaAlbaran, services)
                            If Componentes Is Nothing OrElse Componentes.Rows.Count = 0 Then
                                lineaAlbaran(_AVL.TipoLineaAlbaran) = enumavlTipoLineaAlbaran.avlNormal
                            Else
                                lineaAlbaran(_AVL.TipoLineaAlbaran) = enumavlTipoLineaAlbaran.avlFantasma
                            End If
                        Case DataRowState.Modified
                            If (lineaAlbaran(_AVL.TipoLineaAlbaran) = enumavlTipoLineaAlbaran.avlFantasma) Then
                                Dim datos As New DataDocRow(Doc, lineaAlbaran)
                                ProcessServer.ExecuteTask(Of DataDocRow)(AddressOf ActualizarComponentes, datos, services)
                            End If
                    End Select

                    If Not Componentes Is Nothing AndAlso Componentes.Rows.Count > 0 Then
                        For Each componente As DataRow In Componentes.Select
                            Doc.dtLineas.ImportRow(componente)
                        Next
                        IntPos = Componentes.Rows(Componentes.Rows.Count - 1)("IDOrdenLinea") + 1

                        Componentes.Rows.Clear()
                    Else
                        lineaAlbaran(_AVL.IDOrdenLinea) = IntPos
                        IntPos += 1
                    End If
                End If
            Next
        End If
    End Sub
    <Task()> Public Shared Sub GestionArticulosKit(ByVal Doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        '//Gestion de articulo Kit o viruta                
        If Doc Is Nothing Then Exit Sub
        If Not Doc.dtLineas Is Nothing Then
            Dim Componentes As DataTable
            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
            Dim IntPos As Integer = 1
            For Each lineaAlbaran As DataRow In Doc.dtLineas.Select("", "IDOrdenLinea")
                Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(lineaAlbaran("IDArticulo"))
                Dim ConsumoKit As Boolean = False
                If ArtInfo.KitVenta AndAlso Not ArtInfo.Fantasma Then
                    Select Case lineaAlbaran.RowState
                        Case DataRowState.Added
                            'No es viruta
                            If Not (lineaAlbaran(_AVL.TipoLineaAlbaran) = enumavlTipoLineaAlbaran.avlComponente) Then
                                '//Comprobar si articulo es Kit
                                lineaAlbaran(_AVL.IDOrdenLinea) = IntPos
                                Componentes = ProcessServer.ExecuteTask(Of DataRow, DataTable)(AddressOf ComponentesDePrimerNivel, lineaAlbaran, services)
                                If Not Componentes Is Nothing AndAlso Componentes.Rows.Count > 0 Then
                                    lineaAlbaran(_AVL.TipoLineaAlbaran) = enumavlTipoLineaAlbaran.avlKit
                                    'si es un alta y es un kit y el albaran es de consumo: el almacen de los componentes debe ser el indicado en la linea del kit.
                                    If Doc.HeaderRow("IDTipoAlbaran") = enumTipoAlbaran.Consumo Then
                                        ConsumoKit = True
                                    End If
                                Else
                                    lineaAlbaran(_AVL.TipoLineaAlbaran) = enumavlTipoLineaAlbaran.avlNormal
                                End If
                            End If
                        Case DataRowState.Modified
                            If (lineaAlbaran(_AVL.TipoLineaAlbaran) = enumavlTipoLineaAlbaran.avlKit) Then
                                Dim datos As New DataDocRow(Doc, lineaAlbaran)
                                ProcessServer.ExecuteTask(Of DataDocRow)(AddressOf ActualizarComponentes, datos, services)
                            End If
                    End Select
                    If Not Componentes Is Nothing AndAlso Componentes.Rows.Count > 0 Then
                        For Each componente As DataRow In Componentes.Select
                            If ConsumoKit Then
                                componente("IDAlmacen") = lineaAlbaran(_AVL.IDAlmacen)
                            End If
                            Doc.dtLineas.ImportRow(componente)
                        Next
                        IntPos = Componentes.Rows(Componentes.Rows.Count - 1)("IDOrdenLinea") + 1
                        Componentes.Rows.Clear()
                    Else
                        lineaAlbaran(_AVL.IDOrdenLinea) = IntPos
                        IntPos += 1
                    End If
                Else
                    lineaAlbaran(_AVL.IDOrdenLinea) = IntPos
                    IntPos += 1
                End If
            Next
        End If
    End Sub

    <Task()> Public Shared Function ComponentesDePrimerNivel(ByVal lineaalbaran As DataRow, ByVal services As ServiceProvider) As DataTable
        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(lineaalbaran("IDArticulo"))
        Dim Kits As DataTable = New AlbaranVentaLinea().AddNew
        Dim f As New Filter
        f.Add(New StringFilterItem("IDArticulo", lineaalbaran("IDArticulo")))
        Dim Componentes As DataTable = AdminData.GetData("vNegArticuloComponentesPrimerNivel", f)
        If Not Componentes Is Nothing AndAlso Componentes.Rows.Count > 0 Then
            Dim UdsAB As New ArticuloUnidadAB
            Dim articulo As New Negocio.Articulo
            Dim IntPos As Integer = lineaalbaran("IDOrdenLinea") + 1
            Dim StrTarifaKit As String = New Parametro().TarifaComponenteKit
            Dim DtTarArt As DataTable
            If Length(StrTarifaKit) > 0 Then DtTarArt = New TarifaArticulo().Filter(New FilterItem("IDTarifa", FilterOperator.Equal, StrTarifaKit))
            For Each componente As DataRow In Componentes.Rows
                Dim newrow As DataRow = Kits.NewRow
                newrow(_AVL.IDLineaAlbaran) = AdminData.GetAutoNumeric
                newrow(_AVL.IDAlbaran) = lineaalbaran(_AVL.IDAlbaran)
                newrow(_AVL.IDPedido) = lineaalbaran(_AVL.IDPedido)
                newrow(_AVL.IDLineaPedido) = lineaalbaran(_AVL.IDLineaPedido)
                newrow(_AVL.IDArticulo) = componente("IDComponente")
                newrow(_AVL.DescArticulo) = componente("DescComponente")

                Dim AppParamsGen As ParametroGeneral = services.GetService(Of ParametroGeneral)()
                If AppParamsGen.AlmacenCentroGestionActivo Then
                    newrow(_AVL.IDAlmacen) = lineaalbaran(_AVL.IDAlmacen)
                Else
                    Dim data As New BusinessRuleData("IDAlmacen", lineaalbaran(_AVL.IDAlmacen), New DataRowPropertyAccessor(newrow), Nothing)
                    ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf NegocioGeneral.AsignarArticuloAlmacen, data, services)
                    newrow(_AVL.IDAlmacen) = data.Current("IDAlmacen")
                End If
                newrow(_AVL.IDFormaPago) = lineaalbaran(_AVL.IDFormaPago)
                newrow(_AVL.IDCondicionPago) = lineaalbaran(_AVL.IDCondicionPago)
                newrow(_AVL.IDTipoIva) = lineaalbaran(_AVL.IDTipoIva)
                If Length(componente("IDUdVenta")) > 0 Then
                    newrow(_AVL.IDUdMedida) = componente("IDUdVenta")
                End If
                newrow(_AVL.IDUdInterna) = componente("IDUdInterna")
                newrow(_AVL.QInterna) = lineaalbaran(_AVL.QInterna) * componente("Cantidad")
                Dim StDatos As New ArticuloUnidadAB.DatosFactorConversion
                StDatos.IDArticulo = newrow(_AVL.IDArticulo)
                StDatos.IDUdMedidaA = newrow(_AVL.IDUdMedida)
                StDatos.IDUdMedidaB = newrow(_AVL.IDUdInterna)
                newrow(_AVL.Factor) = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, StDatos, services)
                If newrow(_AVL.Factor) <= 0 Then
                    newrow(_AVL.Factor) = 1
                End If
                newrow(_AVL.QServida) = newrow(_AVL.QInterna) / newrow(_AVL.Factor)
                newrow(_AVL.QServidaUd) = componente("Cantidad")
                newrow(_AVL.UdValoracion) = IIf(componente("UdValoracion") > 0, componente("UdValoracion"), 1)
                newrow(_AVL.Lote) = lineaalbaran(_AVL.Lote)
                newrow(_AVL.CContable) = lineaalbaran(_AVL.CContable)
                newrow(_AVL.Facturable) = False
                newrow(_AVL.IDTipoLinea) = lineaalbaran(_AVL.IDTipoLinea)

                If componente.Table.Columns.Contains("Fantasma") AndAlso componente("Fantasma") Then '//Esta condición se refiere al padre
                    newrow(_AVL.Precio) = 0
                Else
                    If Not DtTarArt Is Nothing AndAlso DtTarArt.Rows.Count > 0 Then
                        Dim DrFind() As DataRow = DtTarArt.Select("IDArticulo = '" & componente("IDComponente") & "'")
                        If DrFind.Length > 0 Then
                            Dim DblTotalPrecio As Double = 0
                            For Each DrComp As DataRow In Componentes.Select
                                Dim DrFindComp() As DataRow = DtTarArt.Select("IDArticulo = '" & DrComp("IDComponente") & "'")
                                If DrFindComp.Length > 0 Then DblTotalPrecio += DrFindComp(0)("Precio") * DrComp("Cantidad")
                            Next
                            Dim DblPorcen As Double = ((DrFind(0)("Precio") * componente("Cantidad")) * 100) / DblTotalPrecio
                            newrow(_AVL.Precio) = ((DblPorcen * (lineaalbaran(_AVL.Importe) / lineaalbaran(_AVL.QServida))) / 100) / componente("Cantidad")
                        Else : newrow(_AVL.Precio) = 0
                        End If
                    Else : newrow(_AVL.Precio) = 0
                    End If
                End If

                'No se valora
                newrow(_AVL.PrecioA) = newrow(_AVL.Precio)
                newrow(_AVL.PrecioB) = 0
                newrow(_AVL.Importe) = 0
                newrow(_AVL.ImporteA) = 0
                newrow(_AVL.ImporteB) = 0
                newrow(_AVL.EstadoFactura) = enumavlEstadoFactura.avlNoFacturado

                'El estado del stock de la línea depende de si el artículo tiene gestión de stock o no

                Dim CaracteristicaArticulo As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf articulo.CaracteristicasArticulo, newrow(_AVL.IDArticulo), services)
                If Not CaracteristicaArticulo Is Nothing AndAlso CaracteristicaArticulo.Rows.Count > 0 Then
                    If CaracteristicaArticulo.Rows(0)("GestionStock") AndAlso componente("GestionStock") = False Then
                        newrow(_AVL.EstadoStock) = enumavlEstadoStock.avlNoActualizado
                    Else
                        newrow(_AVL.EstadoStock) = enumavlEstadoStock.avlSinGestion
                    End If
                Else
                    newrow(_AVL.EstadoStock) = enumavlEstadoStock.avlNoActualizado
                End If

                If ArtInfo.Fantasma AndAlso ArtInfo.Fabrica Then
                    newrow(_AVL.EstadoStock) = enumavlEstadoStock.avlSinGestion
                End If


                newrow(_AVL.IDLineaPadre) = lineaalbaran(_AVL.IDLineaAlbaran)
                newrow(_AVL.TipoLineaAlbaran) = enumavlTipoLineaAlbaran.avlComponente
                newrow(_AVL.IDLineaMaterial) = lineaalbaran(_AVL.IDLineaMaterial)
                newrow(_AVL.Regalo) = False
                newrow(_AVL.IDOrdenLinea) = IntPos

                IntPos += 1
                Kits.Rows.Add(newrow)
            Next
        End If
        Return Kits
    End Function

    <Task()> Public Shared Sub ActualizarComponentes(ByVal data As DataDocRow, ByVal services As ServiceProvider)
        'Dim f As New Filter
        'f.Add(New NumberFilterItem(_AVL.IDAlbaran, data.Row(_AVL.IDAlbaran)))
        'f.Add(New NumberFilterItem(_AVL.IDLineaPadre, data.Row(_AVL.IDLineaAlbaran)))

        'Dim fTipoLineas As New Filter(FilterUnionOperator.Or)
        'fTipoLineas.Add(New NumberFilterItem(_AVL.TipoLineaAlbaran, enumavlTipoLineaAlbaran.avlFantasma))
        'fTipoLineas.Add(New NumberFilterItem(_AVL.TipoLineaAlbaran, enumavlTipoLineaAlbaran.avlComponente))
        'f.Add(fTipoLineas)

        Dim f As New Filter
        f = (New NumberFilterItem(_AVL.IDAlbaran, data.Row(_AVL.IDAlbaran)) And _
             New NumberFilterItem(_AVL.IDLineaPadre, data.Row(_AVL.IDLineaAlbaran)) And _
             ((New NumberFilterItem(_AVL.TipoLineaAlbaran, enumavlTipoLineaAlbaran.avlFantasma) Or New NumberFilterItem(_AVL.TipoLineaAlbaran, enumavlTipoLineaAlbaran.avlComponente))))

        Dim WhereComponentes As String = f.Compose(New AdoFilterComposer)
        Dim Componentes() As DataRow = CType(data.Doc, DocumentCabLin).dtLineas.Select(WhereComponentes)
        If Not Componentes Is Nothing AndAlso Componentes.Length > 0 Then
            If Nz(data.Row(_AVL.QInterna, DataRowVersion.Original), 0) <> 0 Then
                Dim factorVariacion As Double = data.Row(_AVL.QInterna) / data.Row(_AVL.QInterna, DataRowVersion.Original)
                For Each componente As DataRow In Componentes
                    componente(_AVL.QServida) = componente(_AVL.QServida) * factorVariacion
                    componente(_AVL.QInterna) = componente(_AVL.QInterna) * factorVariacion
                Next
            End If
        End If
    End Sub

#End Region

#Region " Corregir Movimientos "

    <Task()> Public Shared Sub CorregirMovimientos(ByVal Doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        ' Dim returnData(-1) As StockUpdateData
        Dim alog As AlbaranLogProcess = services.GetService(Of AlbaranLogProcess)()
        Dim aStockUpdateData(-1) As StockUpdateData
        aStockUpdateData = ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta, StockUpdateData())(AddressOf CorreccionMovimientosCambiosCabecera, Doc, services)
        If Not aStockUpdateData Is Nothing Then ArrayManager.Copy(aStockUpdateData, alog.StockUpdateData)
        aStockUpdateData = ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta, StockUpdateData())(AddressOf CorreccionMovimientosCambiosLineas, Doc, services)
        If Not aStockUpdateData Is Nothing Then ArrayManager.Copy(aStockUpdateData, alog.StockUpdateData)

        ' Return returnData
    End Sub

    <Task()> Public Shared Function CorreccionMovimientosCambiosCabecera(ByVal Doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider) As StockUpdateData()
        Dim returnData(-1) As StockUpdateData
        If Doc.HeaderRow.RowState = DataRowState.Modified Then
            If Doc.HeaderRow("FechaAlbaran", DataRowVersion.Original) <> Doc.HeaderRow("FechaAlbaran") Then
                If Doc.HeaderRow("FechaAlbaran") <> DateTime.MinValue Then
                    Dim FechaDocumento As Date = Doc.HeaderRow("FechaAlbaran")
                    Dim f As New Filter
                    f.Add(New NumberFilterItem("IDAlbaran", Doc.HeaderRow("IDAlbaran")))
                    f.Add(New NumberFilterItem("EstadoStock", enumavlEstadoStock.avlActualizado))
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
                            Dim WhereLineaAlbaran As String = f.Compose(New AdoFilterComposer)
                            Dim lotes() As DataRow = Doc.dtLote.Select(WhereLineaAlbaran)
                            If lotes.Length > 0 Then
                                For Each lote As DataRow In lotes
                                    Dim IDLineaMovimiento As Integer = 0
                                    Dim IDLineaMovimientoEntrada As Integer = 0
                                    '//Movimiento de salida
                                    If IsNumeric(lote("IDMovimientoSalida")) Then IDLineaMovimiento = lote("IDMovimientoSalida")

                                    '//Movimiento de entrada (si existe)
                                    If IsNumeric(lote("IDMovimientoEntrada")) Then IDLineaMovimientoEntrada = lote("IDMovimientoEntrada")

                                    If IDLineaMovimiento <> 0 Then
                                        Dim datCorrMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, IDLineaMovimiento, FechaDocumento, CStr(Doc.HeaderRow("NAlbaran")))
                                        Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datCorrMovto, services)
                                        'ArrayManager.Copy(updateData, returnData)
                                        If Not updateData Is Nothing AndAlso updateData.Estado = EstadoStock.NoActualizado Then
                                            ' ArrayManager.Copy(updateData, returnData)
                                            ApplicationService.GenerateError(updateData.Detalle)
                                        End If
                                    End If
                                    If IDLineaMovimientoEntrada <> 0 Then
                                        Dim datCorrMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, IDLineaMovimientoEntrada, FechaDocumento, CStr(Doc.HeaderRow("NAlbaran")))
                                        Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datCorrMovto, services)
                                        'ArrayManager.Copy(updateData, returnData)
                                        If Not updateData Is Nothing AndAlso updateData.Estado = EstadoStock.NoActualizado Then
                                            ' ArrayManager.Copy(updateData, returnData)
                                            ApplicationService.GenerateError(updateData.Detalle)
                                        End If
                                    End If
                                Next
                            Else
                                Dim IDLineaMovimiento As Integer = 0
                                Dim IDLineaMovimientoEntrada As Integer = 0
                                '//Movimiento de salida
                                If IsNumeric(lineaAlbaran("IDMovimiento")) Then IDLineaMovimiento = lineaAlbaran("IDMovimiento")
                                '//Movimiento de entrada (si existe)
                                If IsNumeric(lineaAlbaran("IDMovimientoEntrada")) Then IDLineaMovimientoEntrada = lineaAlbaran("IDMovimientoEntrada")
                                If IDLineaMovimiento <> 0 Then
                                    Dim datCorrMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, IDLineaMovimiento, FechaDocumento, CStr(Doc.HeaderRow("NAlbaran")))
                                    Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datCorrMovto, services)
                                    If Not updateData Is Nothing AndAlso updateData.Estado = EstadoStock.NoActualizado Then
                                        ' ArrayManager.Copy(updateData, returnData)
                                        ApplicationService.GenerateError(updateData.Detalle)
                                    End If
                                End If
                                If IDLineaMovimientoEntrada <> 0 Then
                                    Dim datCorrMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, IDLineaMovimientoEntrada, FechaDocumento, CStr(Doc.HeaderRow("NAlbaran")))
                                    Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datCorrMovto, services)
                                    If Not updateData Is Nothing AndAlso updateData.Estado = EstadoStock.NoActualizado Then
                                        ' ArrayManager.Copy(updateData, returnData)
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

    <Task()> Public Shared Function CorreccionMovimientosCambiosLineas(ByVal Doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider) As StockUpdateData()

        Dim aStockUpdateData(-1) As StockUpdateData
        'Dim f As New Filter
        'f.Add(New NumberFilterItem("TipoLineaAlbaran", FilterOperator.NotEqual, enumavlTipoLineaAlbaran.avlComponente))

        Dim LineasModificadas As List(Of DataRow) = (From c In Doc.dtLineas Where c.RowState = DataRowState.Modified AndAlso _
                                                      Not c.IsNull("TipoLineaAlbaran") AndAlso _
                                                      c("TipoLineaAlbaran") <> enumavlTipoLineaAlbaran.avlComponente Select c).ToList
        If LineasModificadas Is Nothing OrElse LineasModificadas.Count = 0 Then
            Exit Function
        End If
     
        If ProcessServer.ExecuteTask(Of DocumentCabLin, Boolean)(AddressOf ProcesoComunes.AlbaranEnPeriodoCerradoDoc, Doc, services) Then
            Exit Function
        End If

        For Each lineaAlbaran As DataRow In LineasModificadas
            If lineaAlbaran.RowState = DataRowState.Modified Then
                If lineaAlbaran("QServida", DataRowVersion.Original) <> lineaAlbaran("QServida") OrElse _
                   lineaAlbaran("QInterna", DataRowVersion.Original) <> lineaAlbaran("QInterna") OrElse _
                   lineaAlbaran("ImporteA", DataRowVersion.Original) <> lineaAlbaran("ImporteA") OrElse _
                   lineaAlbaran("ImporteB", DataRowVersion.Original) <> lineaAlbaran("ImporteB") OrElse _
                   lineaAlbaran("Precio", DataRowVersion.Original) <> lineaAlbaran("Precio") OrElse _
                   lineaAlbaran("QEtiContenedor", DataRowVersion.Original) <> lineaAlbaran("QEtiContenedor") OrElse _
                   (lineaAlbaran.Table.Columns.Contains("QInterna2") AndAlso Nz(lineaAlbaran("QInterna2", DataRowVersion.Original), 0) <> Nz(lineaAlbaran("QInterna2"), 0)) Then

                    If lineaAlbaran("Precio") <> lineaAlbaran("Precio", DataRowVersion.Original) AndAlso _
                       lineaAlbaran("QServida") = lineaAlbaran("QServida", DataRowVersion.Original) AndAlso _
                       lineaAlbaran("QInterna") = lineaAlbaran("QInterna", DataRowVersion.Original) AndAlso _
                       lineaAlbaran("QEtiContenedor") = lineaAlbaran("QEtiContenedor", DataRowVersion.Original) AndAlso _
                       (lineaAlbaran.Table.Columns.Contains("QInterna2") AndAlso Nz(lineaAlbaran("QInterna2", DataRowVersion.Original), 0) = Nz(lineaAlbaran("QInterna2"), 0)) Then

                        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(lineaAlbaran("IDArticulo"))
                        If ArtInfo.RecalcularValoracion = CInt(enumtaValoracionSalidas.taMantenerPrecio) Then
                            Dim ctx As New DataDocRow(Doc, lineaAlbaran)
                            Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of DataDocRow, StockUpdateData)(AddressOf CorregirMovimiento, ctx, services)
                            If updateData Is Nothing Then
                                If lineaAlbaran("EstadoStock") = EstadoStock.Actualizado OrElse lineaAlbaran("EstadoStock") = EstadoStock.SinGestion Then
                                    Dim fComponentes As New Filter
                                    fComponentes.Add(New NumberFilterItem("TipoLineaAlbaran", enumavlTipoLineaAlbaran.avlComponente))
                                    fComponentes.Add(New NumberFilterItem("IDLineaPadre", lineaAlbaran("IDLineaAlbaran")))
                                    Dim WhereComponentes As String = fComponentes.Compose(New AdoFilterComposer)
                                    For Each componente As DataRow In Doc.dtLineas.Select(WhereComponentes)
                                        ctx = New DataDocRow(Doc, componente)
                                        updateData = ProcessServer.ExecuteTask(Of DataDocRow, StockUpdateData)(AddressOf CorregirMovimiento, ctx, services)
                                        'ReDim Preserve aStockUpdateData(aStockUpdateData.Length)
                                        'aStockUpdateData(aStockUpdateData.Length - 1) = updateData
                                        If Not updateData Is Nothing AndAlso updateData.Estado = EstadoStock.NoActualizado Then
                                            ArrayManager.Copy(updateData, aStockUpdateData)
                                        End If
                                    Next
                                End If
                            Else
                                'ReDim Preserve aStockUpdateData(aStockUpdateData.Length)
                                'aStockUpdateData(aStockUpdateData.Length - 1) = updateData
                                'ElseIf updateData.Estado = EstadoStock.NoActualizado Then
                                'Throw New Exception(updateData.Detalle)
                                If Not updateData Is Nothing AndAlso updateData.Estado = EstadoStock.NoActualizado Then
                                    ArrayManager.Copy(updateData, aStockUpdateData)
                                End If
                            End If
                        End If
                    Else
                        Dim ctx As New DataDocRow(Doc, lineaAlbaran)
                        Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of DataDocRow, StockUpdateData)(AddressOf CorregirMovimiento, ctx, services)
                        If updateData Is Nothing Then
                            If lineaAlbaran("EstadoStock") = EstadoStock.Actualizado Or lineaAlbaran("EstadoStock") = EstadoStock.SinGestion Then
                                Dim fComponentes As New Filter
                                fComponentes.Add(New NumberFilterItem("TipoLineaAlbaran", enumavlTipoLineaAlbaran.avlComponente))
                                fComponentes.Add(New NumberFilterItem("IDLineaPadre", lineaAlbaran("IDLineaAlbaran")))
                                Dim WhereComponentes As String = fComponentes.Compose(New AdoFilterComposer)
                                For Each componente As DataRow In Doc.dtLineas.Select(WhereComponentes)
                                    ctx = New DataDocRow(Doc, componente)
                                    updateData = ProcessServer.ExecuteTask(Of DataDocRow, StockUpdateData)(AddressOf CorregirMovimiento, ctx, services)
                                    'ReDim Preserve aStockUpdateData(aStockUpdateData.Length)
                                    'aStockUpdateData(aStockUpdateData.Length - 1) = updateData
                                    If Not updateData Is Nothing AndAlso updateData.Estado = EstadoStock.NoActualizado Then
                                        ArrayManager.Copy(updateData, aStockUpdateData)
                                    End If
                                Next
                            End If
                            'ElseIf updateData.Estado = EstadoStock.NoActualizado Then
                            'Throw New Exception(updateData.Detalle)
                        Else
                            'ReDim Preserve aStockUpdateData(aStockUpdateData.Length)
                            'aStockUpdateData(aStockUpdateData.Length - 1) = updateData
                            If Not updateData Is Nothing AndAlso updateData.Estado = EstadoStock.NoActualizado Then
                                ArrayManager.Copy(updateData, aStockUpdateData)
                            End If
                        End If
                    End If
                End If
            End If
        Next

        Return aStockUpdateData
    End Function

    <Task()> Public Shared Function CorregirMovimiento(ByVal ctx As DataDocRow, ByVal services As ServiceProvider) As StockUpdateData
        Dim Cantidad As Double : Dim updateData As StockUpdateData
        Dim lineaAlbaran As DataRow = ctx.Row

        Dim SegundaUnidad As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, lineaAlbaran("IDArticulo"), services)
        ProcessServer.ExecuteTask(Of Object)(AddressOf Comunes.BeginTransaction, Nothing, services)
        If lineaAlbaran("QEtiContenedor", DataRowVersion.Original) <> lineaAlbaran("QEtiContenedor") Then
            Cantidad = Nz(lineaAlbaran("QEtiContenedor"), 0)
            If IsNumeric(lineaAlbaran("IDSalidaContenedor")) Then
                '//Correccion movimiento de salida de contenedor
                Dim datCorrMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, lineaAlbaran("IDSalidaContenedor"), Cantidad, CStr(ctx.Doc.HeaderRow("NAlbaran")))
                updateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datCorrMovto, services)
                If Not updateData Is Nothing AndAlso updateData.Estado = EstadoStock.Actualizado Then
                    lineaAlbaran("IDSalidaContenedor") = updateData.IDLineaMovimiento
                Else
                    If updateData.Estado = EstadoStock.NoActualizado AndAlso Length(lineaAlbaran("IDSalidaContenedor")) > 0 Then
                        Dim actMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Eliminar, lineaAlbaran("IDSalidaContenedor"))
                        ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento)(AddressOf ProcesoStocks.ActualizarMovimiento, actMovto, services)
                        Dim act As New ProcesoStocks.DataActualizarLineas(updateData, lineaAlbaran)
                        ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarLineas)(AddressOf ProcesoStocks.ActualizarLineasAV, act, services)
                    End If
                    Return updateData
                End If
            End If

            If IsNumeric(lineaAlbaran("IDEntradaContenedor")) Then
                '//Correccion movimiento de entrada de contenedor
                Dim datCorrMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, lineaAlbaran("IDEntradaContenedor"), Cantidad, CStr(ctx.Doc.HeaderRow("NAlbaran")))
                updateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datCorrMovto, services)
                If Not updateData Is Nothing AndAlso updateData.Estado = EstadoStock.Actualizado Then
                    lineaAlbaran("IDEntradaContenedor") = updateData.IDLineaMovimiento
                Else
                    If updateData.Estado = EstadoStock.NoActualizado AndAlso Length(lineaAlbaran("IDEntradaContenedor")) > 0 Then
                        Dim actMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Eliminar, lineaAlbaran("IDEntradaContenedor"))
                        ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento)(AddressOf ProcesoStocks.ActualizarMovimiento, actMovto, services)
                        Dim act As New ProcesoStocks.DataActualizarLineas(updateData, lineaAlbaran)
                        ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarLineas)(AddressOf ProcesoStocks.ActualizarLineasAV, act, services)
                    End If
                    Return updateData
                End If
            End If
        End If


        Dim PrecioA As Double : Dim PrecioB As Double
        Cantidad = lineaAlbaran("QInterna")
        If (lineaAlbaran("Factor") <> 0 And lineaAlbaran("UdValoracion") <> 0) Then
            Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
            Dim monedaA As MonedaInfo = Monedas.MonedaA
            Dim monedaB As MonedaInfo = Monedas.MonedaB
            PrecioA = xRound(lineaAlbaran("PrecioA") / lineaAlbaran("Factor") / lineaAlbaran("UdValoracion") * (1 - lineaAlbaran("Dto1") / 100) * (1 - lineaAlbaran("Dto2") / 100) * (1 - lineaAlbaran("Dto3") / 100) * (1 - lineaAlbaran("Dto") / 100) * (1 - lineaAlbaran("DtoProntoPago") / 100), monedaA.NDecimalesPrecio)
            PrecioB = xRound(lineaAlbaran("PrecioB") / lineaAlbaran("Factor") / lineaAlbaran("UdValoracion") * (1 - lineaAlbaran("Dto1") / 100) * (1 - lineaAlbaran("Dto2") / 100) * (1 - lineaAlbaran("Dto3") / 100) * (1 - lineaAlbaran("Dto") / 100) * (1 - lineaAlbaran("DtoProntoPago") / 100), monedaB.NDecimalesPrecio)
        End If

        Dim f As New Filter
        f.Add(New NumberFilterItem("IDLineaAlbaran", lineaAlbaran("IDLineaAlbaran")))
        Dim WhereLotesLinea As String = f.Compose(New AdoFilterComposer)
        Dim lote() As DataRow = CType(ctx.Doc, DocumentoAlbaranVenta).dtLote.Select(WhereLotesLinea)
        If lote.Length > 0 Then
            '//Corregir todos los movimientos asociados a los lotes (solo se corrigen si hay cambio en precio-importe)
            If lineaAlbaran("ImporteA", DataRowVersion.Original) <> lineaAlbaran("ImporteA") OrElse _
               lineaAlbaran("ImporteB", DataRowVersion.Original) <> lineaAlbaran("ImporteB") Then
                For Each dr As DataRow In lote
                    If Not dr.IsNull("IDMovimientoSalida") Then
                        '//Correccion movimiento de salida
                        Dim datCorrMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, dr("IDMovimientoSalida"), PrecioA, PrecioB, CStr(ctx.Doc.HeaderRow("NAlbaran")))
                        updateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datCorrMovto, services)
                        If updateData Is Nothing OrElse updateData.Estado <> EstadoStock.Actualizado Then
                            If updateData.Estado = EstadoStock.NoActualizado AndAlso Length(lineaAlbaran("IDMovimientoSalida")) > 0 Then
                                Dim actMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Eliminar, lineaAlbaran("IDMovimientoSalida"))
                                ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento)(AddressOf ProcesoStocks.ActualizarMovimiento, actMovto, services)
                                Dim act As New ProcesoStocks.DataActualizarLineas(updateData, lineaAlbaran)
                                ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarLineas)(AddressOf ProcesoStocks.ActualizarLineasAV, act, services)
                            End If
                            Return updateData
                        End If
                    End If

                    If Not dr.IsNull("IDMovimientoEntrada") Then
                        '//Correccion movimiento de entrada
                        Dim datCorrMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, dr("IDMovimientoEntrada"), PrecioA, PrecioB, CStr(ctx.Doc.HeaderRow("NAlbaran")))
                        updateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datCorrMovto, services)
                        If updateData Is Nothing OrElse updateData.Estado <> EstadoStock.Actualizado Then
                            If updateData.Estado = EstadoStock.NoActualizado AndAlso Length(lineaAlbaran("IDMovimientoEntrada")) > 0 Then
                                Dim actMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Eliminar, lineaAlbaran("IDMovimientoEntrada"))
                                ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento)(AddressOf ProcesoStocks.ActualizarMovimiento, actMovto, services)
                                Dim act As New ProcesoStocks.DataActualizarLineas(updateData, lineaAlbaran)
                                ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarLineas)(AddressOf ProcesoStocks.ActualizarLineasAV, act, services)
                            End If
                            Return updateData
                        End If
                    End If
                Next
            End If
        Else
            If Not lineaAlbaran.IsNull("IDMovimiento") Then
                '//Correccion movimiento de salida
                Dim datCorrMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, lineaAlbaran("IDMovimiento"), PrecioA, PrecioB, CStr(ctx.Doc.HeaderRow("NAlbaran")))
                datCorrMovto.CorrectContext.CorreccionEnCantidad = lineaAlbaran.RowState = DataRowState.Added OrElse _
                                                                  (lineaAlbaran.RowState = DataRowState.Modified AndAlso Nz(lineaAlbaran("QInterna"), 0) <> Nz(lineaAlbaran("QInterna", DataRowVersion.Original), 0))
                If datCorrMovto.CorrectContext.CorreccionEnCantidad Then
                    datCorrMovto.Cantidad = CDbl(Nz(lineaAlbaran("QInterna"), 0))
                End If

                If SegundaUnidad AndAlso Length(lineaAlbaran("QInterna2")) > 0 Then
                    datCorrMovto.CorrectContext.CorreccionEnCantidad2 = lineaAlbaran.RowState = DataRowState.Added OrElse _
                                                                        (lineaAlbaran.RowState = DataRowState.Modified AndAlso Nz(lineaAlbaran("QInterna2"), 0) <> Nz(lineaAlbaran("QInterna2", DataRowVersion.Original), 0))
                    If datCorrMovto.CorrectContext.CorreccionEnCantidad2 Then
                        datCorrMovto.Cantidad2 = CDbl(lineaAlbaran("QInterna2"))
                    End If
                End If
                updateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datCorrMovto, services)
                If Not updateData Is Nothing AndAlso updateData.Estado = EstadoStock.Actualizado Then
                    lineaAlbaran("IDMovimiento") = updateData.IDLineaMovimiento
                Else
                    If updateData.Estado = EstadoStock.NoActualizado AndAlso Length(lineaAlbaran("IDMovimiento")) > 0 Then
                        Dim actMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Eliminar, lineaAlbaran("IDMovimiento"))
                        ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento)(AddressOf ProcesoStocks.ActualizarMovimiento, actMovto, services)
                        Dim act As New ProcesoStocks.DataActualizarLineas(updateData, lineaAlbaran)
                        ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarLineas)(AddressOf ProcesoStocks.ActualizarLineasAV, act, services)
                    End If
                    Return updateData
                End If
            End If

            If Not lineaAlbaran.IsNull("IDMovimientoEntrada") Then
                '//Correccion movimiento de entrada
                Dim datCorrMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, lineaAlbaran("IDMovimientoEntrada"), PrecioA, PrecioB, CStr(ctx.Doc.HeaderRow("NAlbaran")))
                datCorrMovto.CorrectContext.CorreccionEnCantidad = lineaAlbaran.RowState = DataRowState.Added OrElse _
                                                                 (lineaAlbaran.RowState = DataRowState.Modified AndAlso Nz(lineaAlbaran("QInterna"), 0) <> Nz(lineaAlbaran("QInterna", DataRowVersion.Original), 0))
                If datCorrMovto.CorrectContext.CorreccionEnCantidad Then
                    datCorrMovto.Cantidad = CDbl(Nz(lineaAlbaran("QInterna"), 0))
                End If

                If SegundaUnidad AndAlso Length(lineaAlbaran("QInterna2")) > 0 Then
                    datCorrMovto.CorrectContext.CorreccionEnCantidad2 = lineaAlbaran.RowState = DataRowState.Added OrElse _
                                                                        (lineaAlbaran.RowState = DataRowState.Modified AndAlso Nz(lineaAlbaran("QInterna2"), 0) <> Nz(lineaAlbaran("QInterna2", DataRowVersion.Original), 0))
                    If datCorrMovto.CorrectContext.CorreccionEnCantidad2 Then
                        datCorrMovto.Cantidad2 = CDbl(lineaAlbaran("QInterna2"))
                    End If
                End If
                updateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datCorrMovto, services)
                If Not updateData Is Nothing AndAlso updateData.Estado = EstadoStock.Actualizado Then
                    lineaAlbaran("IDMovimientoEntrada") = updateData.IDLineaMovimiento
                Else
                    If updateData.Estado = EstadoStock.NoActualizado AndAlso Length(lineaAlbaran("IDMovimientoEntrada")) > 0 Then
                        Dim actMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Eliminar, lineaAlbaran("IDMovimientoEntrada"))
                        ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento)(AddressOf ProcesoStocks.ActualizarMovimiento, actMovto, services)
                        Dim act As New ProcesoStocks.DataActualizarLineas(updateData, lineaAlbaran)
                        ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarLineas)(AddressOf ProcesoStocks.ActualizarLineasAV, act, services)
                    End If
                    Return updateData
                End If
            End If
        End If

        ProcessServer.ExecuteTask(Of Boolean)(AddressOf Comunes.CommitTransaction, False, services)
        Return updateData
    End Function


#End Region

#Region " REGALOS "

#Region " Promociones "

    <Task()> Public Shared Sub TratarPromocionesLineas(ByVal Doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
        If Not AppParams.GestionPromocionesComerciales Then Exit Sub

        If Not Nz(Doc.HeaderRow("Ticket"), False) AndAlso Length(Doc.HeaderRow("IDTPV")) = 0 Then
            For Each linea As DataRow In Doc.dtLineas.Select
                If linea("Regalo") = 0 Then
                    If Nz(linea("IDLineaPedido"), 0) <> 0 AndAlso Nz(linea("IDPromocionLinea"), 0) <> 0 Then
                        If linea.RowState = DataRowState.Modified AndAlso _
                          (linea("QServida") <> linea("QServida", DataRowVersion.Original) OrElse linea("IdArticulo") <> linea("IdArticulo", DataRowVersion.Original)) Then
                            ProcessServer.ExecuteTask(Of Integer)(AddressOf PromocionLinea.ActualizarLineaPromocionQ, linea("IDPromocionLinea"), services)
                        End If
                    End If

                    '10. Quitamos la información anterior
                    If linea.RowState = DataRowState.Modified Then
                        If Nz(linea("IDPromocionLinea", DataRowVersion.Original), 0) <> 0 Then
                            If (linea("QServida") <> linea("QServida", DataRowVersion.Original) OrElse linea("IdArticulo") <> linea("IdArticulo", DataRowVersion.Original)) Then
                                Dim Dt As DataTable = Doc.dtLineas.Clone
                                Dt.ImportRow(linea)
                                Dim datActPromo As New PromocionLinea.DatosActuaLinPromoDr(Dt, True, Doc)
                                ProcessServer.ExecuteTask(Of PromocionLinea.DatosActuaLinPromoDr)(AddressOf PromocionLinea.ActualizarLineaPromocion, datActPromo, services)
                            End If
                        End If
                    End If

                    '20. Insertamos la nueva información
                    If Nz(linea("IDPromocionLinea"), 0) <> 0 Then
                        If linea.RowState = DataRowState.Added OrElse _
                          (linea("QServida") <> linea("QServida", DataRowVersion.Original) OrElse linea("IdArticulo") <> linea("IdArticulo", DataRowVersion.Original)) Then
                            Dim pl As New PromocionLinea
                            Dim dtPromLinea As DataTable = pl.SelOnPrimaryKey(linea("IDPromocionLinea"))
                            If Not IsNothing(dtPromLinea) AndAlso dtPromLinea.Rows.Count > 0 Then
                                If linea("QServida") >= dtPromLinea.Rows(0)("QMinPedido") Then
                                    Dim Dt As DataTable = Doc.dtLineas.Clone
                                    Dt.ImportRow(linea)
                                    Dim datActPromo As New PromocionLinea.DatosActuaLinPromoDr(Dt, False, Doc)
                                    ProcessServer.ExecuteTask(Of PromocionLinea.DatosActuaLinPromoDr)(AddressOf PromocionLinea.ActualizarLineaPromocion, datActPromo, services)

                                    Dim datosRegalo As New DataNuevaLineaRegalo(Doc, linea, dtPromLinea.Rows(0))
                                    ProcessServer.ExecuteTask(Of DataNuevaLineaRegalo)(AddressOf NuevaLineaRegalo, datosRegalo, services)
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        End If
    End Sub

    Public Class DataNuevaLineaRegalo
        Public Doc As DocumentoAlbaranVenta
        Public Row As DataRow
        Public RowPromocion As DataRow
        Public ActualizarPromo As Boolean

        Public Sub New(ByVal Doc As DocumentoAlbaranVenta, ByVal Row As DataRow, ByVal RowPromocion As DataRow, Optional ByVal ActualizarPromo As Boolean = True)
            Me.Doc = Doc
            Me.Row = Row
            Me.RowPromocion = RowPromocion
            Me.ActualizarPromo = ActualizarPromo
        End Sub
    End Class

    <Task()> Public Shared Sub NuevaLineaRegalo(ByVal data As DataNuevaLineaRegalo, ByVal services As ServiceProvider)
        If Not IsNothing(data.Row) AndAlso Length(data.Row("IDPromocionLinea")) > 0 Then
            Dim dblQServida As Double
            If data.Row("QServida") > data.RowPromocion("QMaxPedido") Then
                dblQServida = data.RowPromocion("QMaxPedido")
            Else
                dblQServida = data.Row("QServida")
            End If

            Dim f As New Filter
            f.Add(New NumberFilterItem("IDPromocionLinea", data.Row("IDPromocionLinea")))
            f.Add(New StringFilterItem("IDArticulo", data.Row("IDArticulo")))

            Dim dtArticuloRegalo As DataTable = AdminData.GetData("vNegPromocionArticulosRegaloAlbaran", f)
            If Not IsNothing(dtArticuloRegalo) AndAlso dtArticuloRegalo.Rows.Count > 0 Then
                Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
                Dim strAlmacenPred As String = AppParams.Almacen
                Dim AVL As New AlbaranVentaLinea
                Dim strTipoLinea As String = ProcessServer.ExecuteTask(Of Object, String)(AddressOf TipoLinea.TipoLineaRegalo, Nothing, services)
                Dim intOrden As Integer = Nz(data.Doc.dtLineas.Compute("MAX(IDOrdenLinea)", Nothing), 0)

                Dim context As New BusinessData(data.Doc.HeaderRow)
                f.Clear()
                f.Add(New NumberFilterItem("IDAlbaran", data.Row("IDAlbaran")))
                For Each drArticuloRegalo As DataRow In dtArticuloRegalo.Rows
                    'Nuevo registro
                    Dim drAVL As DataRow = data.Doc.dtLineas.NewRow
                    drAVL("IDLineaAlbaran") = AdminData.GetAutoNumeric
                    drAVL("IDTipoLinea") = strTipoLinea
                    drAVL("IDAlbaran") = data.Doc.HeaderRow("IDAlbaran")
                    drAVL("IDFormaPago") = data.Doc.HeaderRow("IDFormaPago")
                    drAVL("IDCondicionPago") = data.Doc.HeaderRow("IDCondicionPago")
                    drAVL("IDCentroGestion") = data.Doc.HeaderRow("IDCentroGestion")

                    drAVL = AVL.ApplyBusinessRule("IDArticulo", drArticuloRegalo("IDArticuloRegalo"), drAVL, context)
                    context("Fecha") = data.Doc.HeaderRow("FechaAlbaran")
                    drAVL("Regalo") = True

                    'En el campo Cantidad guardamos la Cantidad indicada con el ArticuloRegalo
                    drAVL("QServida") = Fix((dblQServida / drArticuloRegalo("QPedida"))) * drArticuloRegalo("QRegalo")
                    If drAVL("QServida") = 0 Then
                        drAVL("QServida") = drArticuloRegalo("QRegalo")
                    End If

                    'Se incrementa el IDOrden para cada linea de regalo generada
                    intOrden = intOrden + 1
                    drAVL("IDOrdenLinea") = intOrden

                    drAVL = AVL.ApplyBusinessRule("QServida", drAVL("QServida"), drAVL, context)
                    drAVL("IDPromocion") = data.Row("IDPromocion")
                    drAVL("IDPromocionLinea") = data.Row("IDPromocionLinea")

                    data.Doc.dtLineas.Rows.Add(drAVL)
                Next

                If data.ActualizarPromo AndAlso Length(data.Row("IDLineaPedido")) = 0 Then
                    'Actualización QPromocionada
                    Dim PL As New PromocionLinea
                    Dim drPromocionLinea As DataRow = PL.GetItemRow(data.Row("IDPromocionLinea"))
                    drPromocionLinea("QPromocionada") = drPromocionLinea("QPromocionada") + dblQServida
                    BusinessHelper.UpdateTable(drPromocionLinea.Table)
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarQLineasPromociones(ByVal Doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
        If Not AppParams.GestionPromocionesComerciales Then Exit Sub

        If Not Nz(Doc.HeaderRow("Ticket"), False) AndAlso Length(Doc.HeaderRow("IDTPV")) = 0 Then
            For Each linea As DataRow In Doc.dtLineas.Select
                If Not Nz(linea("Regalo"), False) Then
                    '10. Quitamos la información anterior
                    If linea.RowState = DataRowState.Modified Then
                        If Nz(linea("IDPromocionLinea", DataRowVersion.Original), 0) <> 0 Then
                            If (linea("QServida") <> linea("QServida", DataRowVersion.Original) OrElse linea("IdArticulo") <> linea("IdArticulo", DataRowVersion.Original)) Then
                                ProcessServer.ExecuteTask(Of Integer)(AddressOf PromocionLinea.ActualizarLineaPromocionQ, linea("IDPromocionLinea", DataRowVersion.Original), services)
                            End If
                        End If
                    End If

                    '30. Actualizamos en función de la QServida al cerrar la línea.
                    If linea.RowState = DataRowState.Modified Then
                        If Nz(linea("IDPromocionLinea"), 0) <> 0 AndAlso Nz(linea("QServida"), 0) <> Nz(linea("QServida", DataRowVersion.Original), 0) Then
                            ProcessServer.ExecuteTask(Of Integer)(AddressOf PromocionLinea.ActualizarLineaPromocionQ, linea("IDPromocionLinea"), services)
                        End If
                    ElseIf linea.RowState = DataRowState.Added Then
                        If Nz(linea("IDPromocionLinea"), 0) <> 0 Then
                            ProcessServer.ExecuteTask(Of Integer)(AddressOf PromocionLinea.ActualizarLineaPromocionQ, linea("IDPromocionLinea"), services)
                        End If
                    End If
                End If
            Next
        End If
    End Sub

#End Region

#Region " Regalos Fidelización "

    <Serializable()> _
    Public Class DataRegalosFidelizacion
        Public IDAlbaran As Integer
        Public dtArticulosRegalo As DataTable

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDAlbaran As Integer, ByVal dtArticulosRegalo As DataTable)
            Me.IDAlbaran = IDAlbaran
            Me.dtArticulosRegalo = dtArticulosRegalo
        End Sub
    End Class

    <Task()> Public Shared Sub ADDRegalosFidelizacion(ByVal data As DataRegalosFidelizacion, ByVal services As ServiceProvider)
        If data.IDAlbaran > 0 Then
            Dim dtArtTratar As DataTable = data.dtArticulosRegalo.Clone
            For Each dr As DataRow In data.dtArticulosRegalo.Select("Cantidad > 0")
                Dim drNew As DataRow = dtArtTratar.NewRow
                drNew.ItemArray = dr.ItemArray
                dtArtTratar.Rows.Add(drNew)
            Next
            If Not IsNothing(dtArtTratar) AndAlso dtArtTratar.Rows.Count > 0 Then
                Dim datosRegalo As New DataRegalosFidelizacion(data.IDAlbaran, dtArtTratar)
                ProcessServer.ExecuteTask(Of DataRegalosFidelizacion)(AddressOf ADDLineaRegaloFidelizacionAlbaran, datosRegalo, services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ADDLineaRegaloFidelizacionAlbaran(ByVal data As DataRegalosFidelizacion, ByVal services As ServiceProvider)
        Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
        Dim strAlmacenPred As String = AppParams.Almacen
        Dim strTipoLinea As String = ProcessServer.ExecuteTask(Of Object, String)(AddressOf TipoLinea.TipoLineaRegalo, Nothing, services)

        Dim drAVC As DataRow = New AlbaranVentaCabecera().GetItemRow(data.IDAlbaran)
        Dim AVL As New AlbaranVentaLinea
        Dim dtAVL As DataTable = AVL.Filter(New NoRowsFilterItem)
        For Each drArticulo As DataRow In data.dtArticulosRegalo.Rows
            Dim dblQServida As Double = drArticulo("Cantidad")
            Dim f As New Filter
            'Nuevo registro
            Dim drAVL As DataRow = dtAVL.NewRow
            drAVL("IDLineaAlbaran") = AdminData.GetAutoNumeric
            drAVL("IDTipoLinea") = strTipoLinea
            drAVL("IDAlbaran") = drAVC("IDAlbaran")
            drAVL("IDFormaPago") = drAVC("IDFormaPago")
            drAVL("IDCondicionPago") = drAVC("IDCondicionPago")
            drAVL("IDCentroGestion") = drAVC("IDCentroGestion")
            Dim drPA_AVL As New DataRowPropertyAccessor(drAVL)
            Dim context As New BusinessData(drAVC)
            AVL.ApplyBusinessRule("IDArticulo", drArticulo("IDArticulo"), drPA_AVL, context)
            drAVL("Regalo") = True
            'En el campo Cantidad guardamos la Cantidad indicada con el ArticuloRegalo
            AVL.ApplyBusinessRule("QServida", dblQServida, drPA_AVL, context)
            drAVL("Precio") = 0 : drAVL("PrecioA") = 0 : drAVL("PrecioB") = 0 : drAVL("PVP") = 0
            drAVL("Importe") = 0 : drAVL("ImporteA") = 0 : drAVL("ImporteB") = 0
            drAVL("ImportePVP") = 0 : drAVL("ImportePVPA") = 0 : drAVL("ImportePVPB") = 0
            dtAVL.Rows.Add(drAVL)
        Next
        If Not IsNothing(dtAVL) AndAlso dtAVL.Rows.Count > 0 Then
            Dim DtParam As DataTable = New Parametro().SelOnPrimaryKey("PUNTOS_IMP")
            If Not DtParam Is Nothing AndAlso DtParam.Rows.Count > 0 Then
                If DtParam.Rows(0)("Valor") > 0 Then
                    If Length(drAVC("IDTarjetaFidelizacion")) > 0 Then
                        Dim StData As New DataActuaPuntos(dtAVL, drAVC("IDTarjetaFidelizacion"), DataViewRowState.Added)
                        ProcessServer.ExecuteTask(Of DataActuaPuntos)(AddressOf ActualizarPuntos, StData, services)
                        AVL.Update(dtAVL)
                        Dim DocCabAV As New DocumentoAlbaranVenta(data.IDAlbaran)
                        ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.ActualizacionAutomaticaStock, DocCabAV, services)
                    End If
                End If
            End If
        End If
    End Sub

#End Region

#End Region

#Region " PUNTOS MARKETING "

    <Task()> Public Shared Sub ActualizarPuntosMarketing(ByVal data As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        Dim DtParam As DataTable = New Parametro().SelOnPrimaryKey("PUNTOS_IMP")
        If Not DtParam Is Nothing AndAlso DtParam.Rows.Count > 0 Then
            If DtParam.Rows(0)("Valor") > 0 Then
                If Not data.dtLineas Is Nothing AndAlso data.dtLineas.Rows.Count > 0 Then
                    Dim DtAdded As DataTable = data.dtLineas.GetChanges(DataRowState.Added)
                    If Not DtAdded Is Nothing AndAlso DtAdded.Rows.Count > 0 AndAlso Length(data.HeaderRow("IDTarjetaFidelizacion")) > 0 Then
                        Dim StData As New DataActuaPuntos(data.dtLineas, data.HeaderRow("IDTarjetaFidelizacion"), DataViewRowState.Added)
                        ProcessServer.ExecuteTask(Of DataActuaPuntos)(AddressOf ActualizarPuntos, StData, services)
                    End If
                    Dim DtModified As DataTable = data.dtLineas.GetChanges(DataRowState.Modified)
                    If Not DtModified Is Nothing AndAlso DtModified.Rows.Count > 0 AndAlso Length(data.HeaderRow("IDTarjetaFidelizacion")) > 0 Then
                        Dim StData As New DataActuaPuntos(data.dtLineas, data.HeaderRow("IDTarjetaFidelizacion"), DataViewRowState.ModifiedCurrent)
                        ProcessServer.ExecuteTask(Of DataActuaPuntos)(AddressOf ActualizarPuntos, StData, services)
                    End If
                End If
            End If
        End If
    End Sub

    '<Task()> Public Shared Sub BorradoPuntosMarketing(ByVal data As DataRow, ByVal services As ServiceProvider)
    '    If CBool(data("Regalo")) = True Then
    '        Dim DtPuntos As DataTable = New BE.DataEngine().Filter("tbAlbaranVentaLineaPuntos", New FilterItem("IDLineaAlbaran", FilterOperator.Equal, data("IDLineaAlbaran")))
    '        If Not DtPuntos Is Nothing AndAlso DtPuntos.Rows.Count > 0 Then
    '            For Each Dr As DataRow In DtPuntos.Select("", "IDLineaAlbaranPuntos")
    '                Dim DtAlbLin As DataTable = New AlbaranVentaLinea().SelOnPrimaryKey(Dr("IDLineaAlbaranPuntos"))
    '                If Not DtAlbLin Is Nothing AndAlso DtAlbLin.Rows.Count > 0 Then
    '                    DtAlbLin.Rows(0)("PuntosUtilizados") -= Dr("Puntos")
    '                    BusinessHelper.UpdateTable(DtAlbLin)
    '                End If
    '            Next
    '        End If
    '    End If
    'End Sub

    <Serializable()> _
    Public Class DataActuaPuntos
        Public DtDatos As DataTable
        Public IDTarjetaFidelizacion As String
        Public Modo As DataViewRowState

        Public Sub New()
        End Sub

        Public Sub New(ByVal DtDatos As DataTable, ByVal IDTarjetaFidelizacion As String, ByVal Modo As DataViewRowState)
            Me.DtDatos = DtDatos
            Me.IDTarjetaFidelizacion = IDTarjetaFidelizacion
            Me.Modo = Modo
        End Sub
    End Class

    <Task()> Public Shared Sub ActualizarPuntos(ByVal data As DataActuaPuntos, ByVal services As ServiceProvider)
        For Each Dr As DataRow In data.DtDatos.Select("", "", data.Modo)
            'If Dr.RowState = DataRowState.Modified Then ProcessServer.ExecuteTask(Of DataRow)(AddressOf BorradoPuntosMarketing, Dr, services)
            Select Case CBool(Dr("Regalo"))
                Case False
                    Dim DtArt As DataTable = New Articulo().SelOnPrimaryKey(Dr("IDArticulo"))
                    Dim IntPuntos As Integer = 0
                    If Not DtArt Is Nothing AndAlso DtArt.Rows.Count > 0 Then
                        IntPuntos = DtArt.Rows(0)("PuntosMarketing")
                    End If
                    Dr("PuntosMarketing") = IntPuntos * Dr("QServida")
                Case True
                    Dim DtArt As DataTable = New Articulo().SelOnPrimaryKey(Dr("IDArticulo"))
                    Dim IntPuntos As Integer = 0
                    If Not DtArt Is Nothing AndAlso DtArt.Rows.Count > 0 Then
                        IntPuntos = DtArt.Rows(0)("ValorPuntosMarketing")
                    End If
                    Dr("PuntosUtilizados") = IntPuntos * Dr("QServida")
                    'Dim DtPuntos As DataTable = New BE.DataEngine().Filter("vFrmMntoAlbaranVentaPuntos", New FilterItem("IDTarjetaFidelizacion", FilterOperator.Equal, data.IDTarjetaFidelizacion), , "IDLineaAlbaran")
                    'If Not DtPuntos Is Nothing AndAlso DtPuntos.Rows.Count > 0 Then
                    '    For Each DrP As DataRow In DtPuntos.Select()
                    '        Dim IntPuntosPend As Integer = (DrP("PuntosMarketing") - IntPuntos) 'Dr("PuntosUtilizados"))
                    '        Dim IntPuntosAsig As Integer = 0
                    '        If IntPuntos > 0 Then
                    '            If IntPuntos > IntPuntosPend Then
                    '                IntPuntosAsig = IntPuntosPend
                    '            Else : IntPuntosAsig = IntPuntos
                    '            End If
                    'Actualizar lineas de albaran con puntos
                    'Dim StrSQLUpdate As String = "UPDATE tbAlbaranVentaLinea "
                    'StrSQLUpdate &= "SET PuntosUtilizados = PuntosUtilizados + " & IntPuntos & " "
                    'StrSQLUpdate &= "WHERE IDLineaAlbaran = " & Dr("IDLineaAlbaran")
                    'AdminData.Execute(StrSQLUpdate)

                    'Insertar nuevo registro en albaranventalineapuntos
                    'Dim FilPuntos As New Filter
                    'FilPuntos.Add("IDLineaAlbaran", FilterOperator.Equal, Dr("IDLineaAlbaran"))
                    'FilPuntos.Add("IDLineaAlbaranPuntos", FilterOperator.Equal, DrP("IDLineaAlbaran"))
                    'Dim DtSelPuntos As DataTable = New BE.DataEngine().Filter("tbAlbaranVentaLineaPuntos", FilPuntos)
                    'Dim StrSQLInsert As String
                    'If Not DtSelPuntos Is Nothing AndAlso DtSelPuntos.Rows.Count > 0 Then
                    '    StrSQLInsert = "UPDATE tbAlbaranVentaLineaPuntos "
                    '    StrSQLInsert &= "SET Puntos = " & IntPuntosAsig & " "
                    '    StrSQLInsert &= "WHERE IDLineaAlbaran = " & Dr("IDLineaAlbaran") & " AND IDLineaAlbaranPuntos = " & DrP("IDLineaAlbaran")
                    'Else
                    '    StrSQLInsert = "INSERT INTO tbAlbaranVentaLineaPuntos "
                    '    StrSQLInsert &= "(IDLineaAlbaran, IDLineaAlbaranPuntos, Puntos) "
                    '    StrSQLInsert &= "VALUES (" & Dr("IDLineaAlbaran") & ", " & DrP("IDLineaAlbaran") & ", " & IntPuntosAsig & ")"
                    'End If
                    'AdminData.Execute(StrSQLInsert)
                    'IntPuntos -= IntPuntosAsig
                    'Else : Exit For
                    'End If
                    '        Next
                    'End If
            End Select
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
        fTipoApunte.Add(New NumberFilterItem("IDTipoApunte", CInt(enumDiarioTipoApunte.AlbaranVenta)))
        f.Add(fTipoApunte)

        f.Add(New NumberFilterItem("Contabilizado", FilterOperator.NotEqual, CInt(enumContabilizado.NoContabilizado)))
        f.Add(New NumberFilterItem("EstadoFactura", CInt(enumContabilizado.NoContabilizado)))
        data.ApuntesAlbaran = New BE.DataEngine().Filter("NegDescontabilizarAV", f)

        Return data
    End Function

#End Region

#Region " Grandes Distribuidores.Canal 9 "

    <Task()> Public Shared Sub AsignarDatosAlbaranOrigen(ByVal doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        Dim AVC As New AlbaranVentaCabecera
        Dim AVL As New AlbaranVentaLinea

        Dim p As New Parametro
        Dim IDTipoAlbaran As String = p.TipoAlbaranAbonoDistribuidor()
        If Len(IDTipoAlbaran) = 0 Then
            ApplicationService.GenerateError("El parámetro 'TIPOALB_AD' no existe o no está correctamente configurado.")
        End If

        Dim ProcInfo As ProcessInfoAVDistrib = services.GetService(Of ProcessInfoAVDistrib)()

        Dim docAlbOrigen As New DocumentoAlbaranVenta(doc.Cabecera.IDOrigen)
        doc.HeaderRow.ItemArray = docAlbOrigen.HeaderRow.ItemArray
        doc.HeaderRow("IDAlbaran") = AdminData.GetAutoNumeric
        doc.HeaderRow("Estado") = enumavcEstadoFactura.avcNoFacturado
        doc.HeaderRow("IDTipoAlbaran") = IDTipoAlbaran
        If Length(ProcInfo.IDContador) > 0 Then
            doc.HeaderRow("IDContador") = ProcInfo.IDContador
        End If
        If Nz(ProcInfo.FechaAlbaran, cnMinDate) <> cnMinDate Then
            doc.HeaderRow("FechaAlbaran") = ProcInfo.FechaAlbaran
        End If

        AVC.ApplyBusinessRule("IDCliente", docAlbOrigen.HeaderRow("IDClienteDistribuidor"), doc.HeaderRow, Nothing)
        doc.HeaderRow("IDClienteDistribuidor") = System.DBNull.Value

        Dim IDTarifaAbono As String
        Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
        Dim ClteInfo As ClienteInfo = Clientes.GetEntity(doc.HeaderRow("IDCliente"))
        If Not ClteInfo Is Nothing Then
            IDTarifaAbono = ClteInfo.IDTarifaAbono
        End If


        Dim context As New BusinessData(doc.HeaderRow)
        context("Fecha") = doc.HeaderRow("FechaAlbaran")
        For Each drLineaOrigen As DataRow In docAlbOrigen.dtLineas.Rows
            Dim IDLineaAlbaranOLD As Integer = drLineaOrigen("IDLineaAlbaran")
            Dim drLinea As DataRow = doc.dtLineas.NewRow
            Dim IDLineaAlbaranNEW As Integer = AdminData.GetAutoNumeric
            drLinea.ItemArray = drLineaOrigen.ItemArray
            drLinea("IDLineaAlbaran") = IDLineaAlbaranNEW
            drLinea("IDAlbaran") = doc.HeaderRow("IDAlbaran")
            drLinea("EstadoFactura") = enumavlEstadoFactura.avlNoFacturado
            drLinea("EstadoStock") = enumavlEstadoStock.avlSinGestion
            drLinea("Precio") = 0
            drLinea("IDPedido") = System.DBNull.Value
            drLinea("IDLineaPedido") = System.DBNull.Value

            If Length(IDTarifaAbono) > 0 Then drLinea("IDTarifa") = IDTarifaAbono
            '//Introducimos la cantidad en positivo para que me traiga los precios del Distribuidor, 
            '/pero luego le ponemos la cantidad en negativo y el importe
            AVL.ApplyBusinessRule("QServida", drLineaOrigen("QServida"), drLinea, context)
            drLinea("QServida") = -1 * drLineaOrigen("QServida")
            drLinea("Importe") = -1 * drLinea("Importe")
            doc.dtLineas.Rows.Add(drLinea)

            '//Recalcular
            For Each drRepresOrigen As DataRow In docAlbOrigen.dtVentaRepresentante.Select("IDLineaAlbaran=" & IDLineaAlbaranOLD, Nothing)
                Dim drRepres As DataRow = doc.dtVentaRepresentante.NewRow
                drRepres.ItemArray = drRepresOrigen.ItemArray
                drRepres("IDLineaAlbaran") = IDLineaAlbaranNEW

                doc.dtVentaRepresentante.Rows.Add(drRepres)
            Next

            '//Recalcular
            For Each drAnaliticaOrigen As DataRow In docAlbOrigen.dtAnalitica.Select("IDLineaAlbaran=" & IDLineaAlbaranOLD, Nothing)
                Dim drAnalitica As DataRow = doc.dtAnalitica.NewRow
                drAnalitica.ItemArray = drAnaliticaOrigen.ItemArray
                drAnalitica("IDLineaAlbaran") = IDLineaAlbaranNEW

                doc.dtAnalitica.Rows.Add(drAnalitica)
            Next

            '//Hay que recalcular los importes de las analiticas , representantes y totales cabecera
            Dim ctx As New DataDocRow(doc, drLinea)
            ProcessServer.ExecuteTask(Of DataDocRow)(AddressOf ProcesoComercial.ActualizarRepresentantes, ctx, services)
            ProcessServer.ExecuteTask(Of DataDocRow)(AddressOf NegocioGeneral.ActualizarAnalitica, ctx, services)
        Next


        Dim datGastos As New ProcesoAlbaranVenta.DataAddGastoDistribuidor(docAlbOrigen, doc)
        ProcessServer.ExecuteTask(Of ProcesoAlbaranVenta.DataAddGastoDistribuidor)(AddressOf ProcesoAlbaranVenta.AddGastoDistribuidor, datGastos, services)

    End Sub


    <Serializable()> _
    Public Class DataAddGastoDistribuidor
        Public DocOrigen As DocumentoAlbaranVenta
        Public Doc As DocumentoAlbaranVenta

        Public Sub New(ByVal DocOrigen As DocumentoAlbaranVenta, ByVal Doc As DocumentoAlbaranVenta)
            Me.DocOrigen = DocOrigen
            Me.Doc = Doc
        End Sub
    End Class
    <Task()> Public Shared Sub AddGastoDistribuidor(ByVal data As DataAddGastoDistribuidor, ByVal services As ServiceProvider)
        Dim doc As DocumentoAlbaranVenta = data.Doc
        Dim docOrigen As DocumentoAlbaranVenta = data.DocOrigen

        If Length(docOrigen.HeaderRow("IDClienteDistribuidor")) > 0 AndAlso Length(doc.HeaderRow("IDTipoAlbaran")) > 0 Then
            Dim TipoAlbInfo As ProcesoAlbaranVenta.TipoAlbaranInfo = ProcessServer.ExecuteTask(Of String, ProcesoAlbaranVenta.TipoAlbaranInfo)(AddressOf ProcesoAlbaranVenta.TipoDeAlbaran, doc.HeaderRow("IDTipoAlbaran"), services)
            If TipoAlbInfo.Tipo = enumTipoAlbaran.AbonoDistribuidor Then
                Dim f As New Filter
                f.Add(New StringFilterItem("IDCliente", docOrigen.HeaderRow("IDCliente")))
                f.Add(New StringFilterItem("IDClienteDistribuidor", doc.HeaderRow("IDCliente")))

                Dim fPeriodo As New Filter
                fPeriodo.Add(New DateFilterItem("FechaDesde", FilterOperator.LessThanOrEqual, doc.HeaderRow("FechaAlbaran")))
                fPeriodo.Add(New DateFilterItem("FechaHasta", FilterOperator.GreaterThanOrEqual, doc.HeaderRow("FechaAlbaran")))
                Dim fFechas As New Filter(FilterUnionOperator.Or)
                fFechas.Add(New IsNullFilterItem("FechaDesde", True))
                fFechas.Add(New IsNullFilterItem("FechaHasta", True))
                fFechas.Add(fPeriodo)
                f.Add(fFechas)

                Dim dtRegClteDistrib As DataTable = New ClienteDistribuidor().Filter(f)
                If dtRegClteDistrib.Rows.Count > 0 Then
                    Dim AVL As New AlbaranVentaLinea
                    Dim context As New BusinessData(doc.HeaderRow)

                    Dim PorcGasto As Double = dtRegClteDistrib.Rows(0)("PorcentajeLogistica")
                    Dim ImporteUD As Double = dtRegClteDistrib.Rows(0)("ImporteUD")

                    data.Doc.HeaderRow("PorcentajeLogistica") = PorcGasto
                    data.Doc.HeaderRow("ImporteUD") = ImporteUD

                End If

            End If
        End If

    End Sub


#End Region

End Class

Public Class LineasAlbaranEliminadas
    Public IDLineas As Hashtable

    Public Sub New()
        IDLineas = New Hashtable
    End Sub
End Class


