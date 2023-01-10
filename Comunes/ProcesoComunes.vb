Public Class DataCopiaDatos
    Public Origen As DataRow
    Public Destino As DataRow

    Public Sub New(ByVal Origen As DataRow, ByVal Destino As DataRow)
        Me.Origen = Origen
        Me.Destino = Destino
    End Sub
End Class

Public Class ProcesoComunes

    <Serializable()> _
    Public Class DataCalculoTotalesCab
        Public BasesImponibles() As DataBaseImponible
        Public Doc As DocumentCabLin

        Public Sub New()
        End Sub

        Public Sub New(ByVal BasesImponibles() As DataBaseImponible, ByVal Doc As DocumentCabLin)
            Me.BasesImponibles = BasesImponibles
            Me.Doc = Doc
        End Sub
    End Class

#Region "Contador"

    <Task()> Public Shared Sub AsignarContador(ByVal Doc As DocumentCabLin, ByVal services As ServiceProvider)
        If Doc.HeaderRow.IsNull("IDContador") Then
            Dim Info As ProcessInfo = services.GetService(Of ProcessInfo)()
            If Len(Info.IDContador) > 0 Then
                Doc.HeaderRow("IDContador") = Info.IDContador
            Else
                Dim TipoContador As CentroGestion.ContadorEntidad
                Select Case Doc.EntidadCabecera
                    Case GetType(PedidoVentaCabecera).Name
                        TipoContador = CentroGestion.ContadorEntidad.PedidoVenta
                    Case GetType(AlbaranVentaCabecera).Name
                        TipoContador = CentroGestion.ContadorEntidad.AlbaranVenta
                    Case GetType(FacturaVentaCabecera).Name
                        TipoContador = CentroGestion.ContadorEntidad.FacturaVenta
                    Case GetType(PedidoCompraCabecera).Name
                        TipoContador = CentroGestion.ContadorEntidad.PedidoCompra
                    Case GetType(AlbaranCompraCabecera).Name
                        TipoContador = CentroGestion.ContadorEntidad.AlbaranCompra
                    Case GetType(FacturaCompraCabecera).Name
                        TipoContador = CentroGestion.ContadorEntidad.FacturaCompra
                End Select
                Dim o As New CentroEntidad
                o.CentroGestion = Doc.HeaderRow("IDCentroGestion") & String.Empty
                o.ContadorEntidad = TipoContador
                Doc.HeaderRow("IDContador") = ProcessServer.ExecuteTask(Of CentroEntidad, String)(AddressOf CentroGestion.GetContadorPredeterminado, o, services)
                Info.IDContadorEntidad = Doc.HeaderRow("IDContador")
            End If
        End If
    End Sub
    'David Velasco 27/7/22 
    <Task()> Public Shared Sub AsignarContadorPiso(ByVal Doc As DocumentCabLin, ByVal services As ServiceProvider)

        Dim Info As ProcessInfo = services.GetService(Of ProcessInfo)()
        Doc.HeaderRow("IDContador") = Info.IDContador
        Dim TipoContador As CentroGestion.ContadorEntidad
        TipoContador = CentroGestion.ContadorEntidad.FacturaCompra
        Dim o As New CentroEntidad
        o.CentroGestion = Doc.HeaderRow("IDCentroGestion") & String.Empty
        o.ContadorEntidad = TipoContador
        Doc.HeaderRow("IDContador") = "FCVV23"
        Info.IDContadorEntidad = Doc.HeaderRow("IDContador")

    End Sub
#End Region

#Region " Totales Documentos "

    <Task()> Public Shared Function DesglosarImporte(ByVal Doc As DocumentCabLin, ByVal services As ServiceProvider) As DataBaseImponible()
        If Not IsNothing(Doc.dtLineas) AndAlso Doc.dtLineas.Rows.Count > 0 Then
            Dim rslt(-1) As DataBaseImponible

            For Each linea As DataRow In Doc.dtLineas.Select(Nothing, "IDTipoIVA")
                If Length(linea("IDTipoIva")) > 0 Then
                    Dim IDTipoIVA As String = linea("IDTipoIva")

                    '//se busca el objeto BaseImponible adecuado
                    Dim bi As DataBaseImponible = Nothing
                    For i As Integer = 0 To rslt.Length - 1
                        If rslt(i).IDTipoIva = IDTipoIVA Then bi = rslt(i)
                    Next
                    If bi Is Nothing Then
                        ReDim Preserve rslt(rslt.Length)
                        bi = New DataBaseImponible(IDTipoIVA)
                        rslt(rslt.Length - 1) = bi
                    End If
                    '//

                    Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                    Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(linea("IDArticulo"))
                    If ArtInfo.PorcenIVANoDeducible <> 0 Then
                        bi.PorcenIVANoDeducible = ArtInfo.PorcenIVANoDeducible
                        bi.ImporteIVANoDeducible = bi.ImporteIVANoDeducible + linea("Importe") * (bi.PorcenIVANoDeducible / 100)
                        bi.ImporteIVANoDeducibleA = bi.ImporteIVANoDeducibleA + linea("Importea") * (bi.PorcenIVANoDeducible / 100)
                        bi.ImporteIVANoDeducibleB = bi.ImporteIVANoDeducibleB + linea("Importeb") * (bi.PorcenIVANoDeducible / 100)

                    End If
                    bi.BaseImponible = bi.BaseImponible + linea("Importe")
                    bi.BaseImponibleA = bi.BaseImponibleA + Nz(linea("ImporteA"), 0)
                    bi.BaseImponibleB = bi.BaseImponibleB + Nz(linea("ImporteB"), 0)
                    If Doc.dtLineas.Columns.Contains("ImportePVP") AndAlso Nz(linea("ImportePVP"), 0) <> 0 Then
                        bi.ImporteIVA = bi.ImporteIVA + linea("ImportePVP")
                        bi.ImporteIVAA = bi.ImporteIVAA + (Nz(linea("ImportePVPA"), 0))
                        bi.ImporteIVAB = bi.ImporteIVAA + (Nz(linea("ImportePVPB"), 0))
                    End If
                End If
            Next
            Return rslt
        End If
    End Function

    <Task()> Public Shared Sub TotalDocumento(ByVal Doc As DocumentCabLin, ByVal services As ServiceProvider)
        If Not IsNothing(Doc.HeaderRow) Then
            'Se calcula el total del documento sumando todos los recargos a la Base Imponible.
            'RetencionIRPF

            Dim Total As Double = Nz(Doc.HeaderRow("BaseImponible"), 0) + Nz(Doc.HeaderRow("ImpIVA"), 0) + Nz(Doc.HeaderRow("ImpRE"), 0)
            If Doc.HeaderRow.Table.Columns.Contains("ImpRecFinan") Then Total += Nz(Doc.HeaderRow("ImpRecFinan"), 0)
            Dim TotalA As Double = Nz(Doc.HeaderRow("BaseImponibleA"), 0) + Nz(Doc.HeaderRow("ImpIVAA"), 0) + Nz(Doc.HeaderRow("ImpREA"), 0)
            If Doc.HeaderRow.Table.Columns.Contains("ImpRecFinanA") Then TotalA += Nz(Doc.HeaderRow("ImpRecFinanA"), 0)
            Dim TotalB As Double = Nz(Doc.HeaderRow("BaseImponibleB"), 0) + Nz(Doc.HeaderRow("ImpIVAB"), 0) + Nz(Doc.HeaderRow("ImpREB"), 0)
            If Doc.HeaderRow.Table.Columns.Contains("ImpRecFinanB") Then TotalB += Nz(Doc.HeaderRow("ImpRecFinanB"), 0)
            Doc.HeaderRow("ImpTotal") = Total
            Doc.HeaderRow("ImpTotalA") = TotalA
            Doc.HeaderRow("ImpTotalB") = TotalB

        End If
    End Sub

#End Region

#Region "Importes"

    '//Este método se utlizará únicamente desde donde el campo correspondiente se llame Cantidad.
    <Task()> Public Shared Sub CalcularImporteLineas(ByVal doc As DocumentCabLin, ByVal services As ServiceProvider)
        For Each linea As DataRow In doc.dtLineas.Rows
            If linea.RowState <> DataRowState.Deleted Then
                Dim bdLinea As New DataRowPropertyAccessor(linea)
                bdLinea("IDMoneda") = doc.HeaderRow("IDMoneda") & String.Empty
                ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf NegocioGeneral.CalcularPrecioImporte, bdLinea, services)
                Dim lineaIProperty As New ValoresAyB(bdLinea, doc.IDMoneda, doc.CambioA, doc.CambioB)
                ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, lineaIProperty, services)
            End If
        Next
    End Sub



#End Region
#Region "Cambios Moneda"

    <Task()> Public Shared Sub CambioMoneda(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("Fecha")) = 0 Then data.Current("Fecha") = Today
        If Length(data.Current("IDMoneda")) > 0 Then
            Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
            Dim MonInfo As MonedaInfo = Monedas.GetMoneda(data.Current("IDMoneda"), data.Current("Fecha"))
            data.Current("CambioA") = MonInfo.CambioA
            data.Current("CambioB") = MonInfo.CambioB
        Else
            data.Current("CambioA") = System.DBNull.Value
            data.Current("CambioB") = System.DBNull.Value
        End If
    End Sub

    <Task()> Public Shared Sub AsignarCambiosMoneda(ByVal Doc As DocumentCabLin, ByVal services As ServiceProvider)
        Doc.HeaderRow("CambioA") = Doc.Moneda.CambioA
        Doc.HeaderRow("CambioB") = Doc.Moneda.CambioB
        ProcessServer.ExecuteTask(Of DocumentCabLin)(AddressOf ProcesoComunes.ActualizarCambiosMoneda, Doc, services)
    End Sub

    <Task()> Public Shared Sub ActualizarCambiosMoneda(ByVal Doc As DocumentCabLin, ByVal services As ServiceProvider)
        If Doc.HeaderRow.RowState = DataRowState.Modified Then
            If Doc.HeaderRow("IDMoneda") <> Nz(Doc.HeaderRow("IDMoneda", DataRowVersion.Original), Nothing) OrElse _
               Doc.HeaderRow("CambioA") <> Nz(Doc.HeaderRow("CambioA", DataRowVersion.Original), Nothing) OrElse _
               Doc.HeaderRow("CambioB") <> Nz(Doc.HeaderRow("CambioB", DataRowVersion.Original), Nothing) Then

                Dim Lineas As DataTable = Doc.dtLineas
                If Not IsNothing(Lineas) AndAlso Lineas.Rows.Count Then
                    Dim context As New BusinessData(Doc.HeaderRow)
                    Dim L As BusinessHelper = BusinessHelper.CreateBusinessObject(Doc.EntidadLineas)
                    For Each row As DataRow In Lineas.Rows
                        If Doc.HeaderRow("IDMoneda") <> Nz(Doc.HeaderRow("IDMoneda", DataRowVersion.Original), Nothing) Then
                            Dim datos As New DataCambioMoneda(New DataRowPropertyAccessor(row), Doc.HeaderRow("IDMoneda", DataRowVersion.Original), Doc.HeaderRow("IDMoneda"), Doc.Fecha)
                            ProcessServer.ExecuteTask(Of DataCambioMoneda)(AddressOf NegocioGeneral.CambioMoneda, datos, services)
                        End If
                        L.ApplyBusinessRule("Precio", row("Precio"), row, context)
                    Next
                End If
            End If
        End If
    End Sub

#End Region

#Region " Métodos comunes de asignación de campos "

    <Task()> Public Shared Sub AsignarCentroGestion(ByVal Doc As DocumentCabLin, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf NegocioGeneral.AsignarCentroGestion, Doc.HeaderRow, services)
    End Sub

    <Task()> Public Shared Sub AsignarAlmacen(ByVal Doc As DocumentCabLin, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf NegocioGeneral.AsignarAlmacen, Doc.HeaderRow, services)
    End Sub

    <Task()> Public Shared Sub AsignarIdentificadorPedido(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.IsNull("IDPedido") Then data("IDPedido") = AdminData.GetAutoNumeric
    End Sub

    <Task()> Public Shared Sub AsignarIdentificadorAlbaran(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.IsNull("IDAlbaran") Then data("IDAlbaran") = AdminData.GetAutoNumeric
    End Sub

    <Task()> Public Shared Sub AsignarIdentificadorFactura(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.IsNull("IDFactura") Then data("IDFactura") = AdminData.GetAutoNumeric
    End Sub

    <Task()> Public Shared Sub AsignarFechaPedido(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.IsNull("FechaPedido") OrElse data("FechaPedido") = cnMinDate Then data("FechaPedido") = Date.Today
    End Sub

    <Task()> Public Shared Sub AsignarFechaEntrega(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider)
        If IsDBNull(data("FechaEntrega")) Then data("FechaEntrega") = data("FechaPedido")
    End Sub

    <Task()> Public Shared Sub AsignarFechaAlbaran(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.IsNull("FechaAlbaran") Then data("FechaAlbaran") = Date.Today
    End Sub

    <Task()> Public Shared Sub AsignarFechaFactura(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.IsNull("FechaFactura") Then data("FechaFactura") = Date.Today
    End Sub

    <Task()> Public Shared Sub AsignarFechaParaDeclaracion(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.IsNull("FechaDeclaracionManual") Then data("FechaDeclaracionManual") = False
        If data.IsNull("FechaParaDeclaracion") Then
            data("FechaParaDeclaracion") = Date.Today
            Dim ipData As New DataRowPropertyAccessor(data)
            ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ProcesoFacturacionCompra.FechaParaDeclaracionPorProveedor, ipData, services)
            ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ProcesoFacturacionVenta.FechaParaDeclaracionComoProveedor, ipData, services)
        End If
    End Sub

    <Task()> Public Shared Sub AsignarEjercicioContablePedido(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider)
        Dim dataEj As New DataEjercicio(data, Nz(data("FechaPedido"), cnMinDate))
        ProcessServer.ExecuteTask(Of DataEjercicio)(AddressOf NegocioGeneral.AsignarEjercicioContable, dataEj, services)
        If Length(data("FechaPedido")) = 0 Then data("IDEjercicio") = System.DBNull.Value
    End Sub

    <Task()> Public Shared Sub AsignarEjercicioContableAlbaran(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider)
        Dim dataEj As New DataEjercicio(data, Nz(data("FechaAlbaran"), cnMinDate))
        ProcessServer.ExecuteTask(Of DataEjercicio)(AddressOf NegocioGeneral.AsignarEjercicioContable, dataEj, services)
        If Length(data("FechaAlbaran")) = 0 Then data("IDEjercicio") = System.DBNull.Value
    End Sub

    <Task()> Public Shared Sub AsignarEjercicioContableFactura(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider)
        Dim dataEj As New DataEjercicio(data, Nz(data("FechaFactura"), cnMinDate))
        ProcessServer.ExecuteTask(Of DataEjercicio)(AddressOf NegocioGeneral.AsignarEjercicioContable, dataEj, services)
        If Length(data("FechaFactura")) = 0 Then data("IDEjercicio") = System.DBNull.Value
    End Sub

    <Task()> Public Shared Sub AsignarNumeroPedido(ByVal Doc As DocumentCabLin, ByVal services As ServiceProvider)
        If Doc.HeaderRow.RowState = DataRowState.Added Then
            If Not IsDBNull(Doc.HeaderRow("IDContador")) Then
                Dim Business As BusinessHelper = BusinessHelper.CreateBusinessObject(Doc.EntidadCabecera)
                Dim StDatos As New Contador.DatosCounterValue
                StDatos.IDCounter = Doc.HeaderRow("IDContador")
                StDatos.TargetClass = Business
                StDatos.TargetField = "NPedido"
                StDatos.DateField = "FechaPedido"
                StDatos.DateValue = Doc.HeaderRow("FechaPedido")
                StDatos.IDEjercicio = Doc.HeaderRow("IDEjercicio") & String.Empty
                Doc.HeaderRow("NPedido") = ProcessServer.ExecuteTask(Of Contador.DatosCounterValue, String)(AddressOf Contador.CounterValue, StDatos, services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarNumeroAlbaran(ByVal Doc As DocumentCabLin, ByVal services As ServiceProvider)
        If Doc.HeaderRow.RowState = DataRowState.Added Then
            If Not IsDBNull(Doc.HeaderRow("IDContador")) Then
                Dim Business As BusinessHelper = BusinessHelper.CreateBusinessObject(Doc.EntidadCabecera)
                Dim StDatos As New Contador.DatosCounterValue
                StDatos.IDCounter = Doc.HeaderRow("IDContador")
                StDatos.TargetClass = Business
                StDatos.TargetField = "NAlbaran"
                StDatos.DateField = "FechaAlbaran"
                StDatos.DateValue = Doc.HeaderRow("FechaAlbaran")
                StDatos.IDEjercicio = Doc.HeaderRow("IDEjercicio") & String.Empty
                Doc.HeaderRow("NAlbaran") = ProcessServer.ExecuteTask(Of Contador.DatosCounterValue, String)(AddressOf Contador.CounterValue, StDatos, services)
            End If
        ElseIf Doc.HeaderRow.RowState = DataRowState.Modified Then
            '//En el TPV tenemos la opción de cambiar de contador en un Ticket cuando se está modificando.
            If Doc.HeaderRow.Table.Columns.Contains("IDTPV") AndAlso Length(Doc.HeaderRow("IDTPV")) > 0 AndAlso Doc.HeaderRow("IDContador", DataRowVersion.Original) & String.Empty <> Doc.HeaderRow("IDContador") & String.Empty Then
                Dim Business As BusinessHelper = BusinessHelper.CreateBusinessObject(Doc.EntidadCabecera)
                Dim StDatos As New Contador.DatosCounterValue
                StDatos.IDCounter = Doc.HeaderRow("IDContador")
                StDatos.TargetClass = Business
                StDatos.TargetField = "NAlbaran"
                StDatos.DateField = "FechaAlbaran"
                StDatos.DateValue = Doc.HeaderRow("FechaAlbaran")
                StDatos.IDEjercicio = Doc.HeaderRow("IDEjercicio") & String.Empty
                Doc.HeaderRow("NAlbaran") = ProcessServer.ExecuteTask(Of Contador.DatosCounterValue, String)(AddressOf Contador.CounterValue, StDatos, services)

                ProcessServer.ExecuteTask(Of DocumentCabLin)(AddressOf ProcesoAlbaranVenta.CambiarNDocumentoMovimientos, Doc, services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarNumeroFactura(ByVal Doc As DocumentCabLin, ByVal services As ServiceProvider)
        If Doc.HeaderRow.RowState = DataRowState.Added Then
            If Not IsDBNull(Doc.HeaderRow("IDContador")) Then
                Dim Business As BusinessHelper = BusinessHelper.CreateBusinessObject(Doc.EntidadCabecera)
                Dim StDatos As New Contador.DatosCounterValue
                StDatos.IDCounter = Doc.HeaderRow("IDContador")
                StDatos.TargetClass = Business
                StDatos.TargetField = "NFactura"
                StDatos.DateField = "FechaFactura"
                StDatos.DateValue = Doc.HeaderRow("FechaFactura")
                StDatos.IDEjercicio = Doc.HeaderRow("IDEjercicio") & String.Empty
                Doc.HeaderRow("NFactura") = ProcessServer.ExecuteTask(Of Contador.DatosCounterValue, String)(AddressOf Contador.CounterValue, StDatos, services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub DescuentosCeroCabeceraFactura(ByVal Doc As DocumentCabLin, ByVal services As ServiceProvider)
        Doc.HeaderRow("DtoFactura") = 0
        Doc.HeaderRow("DtoProntoPago") = 0
    End Sub

    <Task()> Public Shared Sub DescuentosCeroLineas(ByVal Doc As DocumentCabLin, ByVal services As ServiceProvider)
        For Each linea As DataRow In Doc.dtLineas.Rows
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf DescuentosCeroLinea, linea, services)
        Next
    End Sub

    <Task()> Public Shared Sub DescuentosCeroLinea(ByVal data As DataRow, ByVal services As ServiceProvider)
        data("Dto1") = 0
        data("Dto2") = 0
        data("Dto3") = 0
        data("Dto") = 0
        data("DtoProntoPago") = 0
    End Sub


#End Region

#Region " Retorno de resultados "

    <Task()> Public Shared Function ResultadoAlbaran(ByVal data As Object, ByVal services As ServiceProvider) As AlbaranLogProcess
        Return services.GetService(Of AlbaranLogProcess)()
    End Function

    <Task()> Public Shared Function ResultadoLogProcess(ByVal data As Object, ByVal services As ServiceProvider) As LogProcess
        Return services.GetService(Of LogProcess)()
    End Function

#End Region

#Region " Validaciones de borrado comunes "

    <Task()> Public Shared Sub ValidarDelRegistroSistema(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Nz(data("Sistema"), False) Then
            ApplicationService.GenerateError("No se puede eliminar el registro seleccionado. Es un registo de sistema.")
        End If
    End Sub

#End Region

    <Task()> Public Shared Sub CalcularImporteLineasAlbaran(ByVal doc As DocumentCabLin, ByVal services As ServiceProvider)
        ' ProcessServer.ExecuteTask(Of DocumentCabLin)(AddressOf ProcesoComercial.RecuperarTiposIVADireccionEnvio, doc, services)
        '//NO utilizar el CalcularImporteLineas de ProcesoComunes. Hay que pasar la cantidad a la QServida y viceversa.
        For Each linea As DataRow In doc.dtLineas.Rows
            If linea.RowState <> DataRowState.Unchanged AndAlso linea.RowState <> DataRowState.Deleted Then
                If linea("TipoLineaAlbaran") <> enumavlTipoLineaAlbaran.avlComponente Then
                    Dim ILinea As IPropertyAccessor = New DataRowPropertyAccessor(linea)

                    ' Dim ILinea As New DataRowPropertyAccessor(linea)
                    ILinea("Cantidad") = linea("QServida")
                    ILinea("IDMoneda") = doc.IDMoneda

                    ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf NegocioGeneral.CalcularPrecioImporte, ILinea, services)
                    Dim lineaIProperty As New ValoresAyB(ILinea, doc.IDMoneda, doc.CambioA, doc.CambioB)
                    ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, lineaIProperty, services)
                End If
            End If
        Next
    End Sub

#Region " Business Rules comunes "

    <Task()> Public Shared Sub ValidarValorNumerico(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Not IsNumeric(data.Value) Then
            ApplicationService.GenerateError("Campo no numérico.")
        End If
    End Sub

    <Task()> Public Shared Sub CambioFechaAlbaran(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim FechaOld As Date?
        If data.ColumnName = "Fecha" Then
            '//Hay que ponerlo en los dos campos indicados en el Synonimous.
            If Nz(data.Current("FechaAlbaran"), cnMinDate) <> cnMinDate Then FechaOld = data.Current("FechaAlbaran")
            data.Current(data.ColumnName) = data.Value
            data.Current("FechaAlbaran") = data.Value
        End If
        If Length(data.Current("FechaAlbaran")) > 0 Then
            Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
            If AppParams.GestionInventarioPermanente Then
                Dim datIStock As New ProcesoStocks.DataCreateIStockClass(InventariosPermanentes.ENSAMBLADO_INV_PERMANENTE_STOCKS, InventariosPermanentes.CLASE_INV_PERMANENTE_STOCKS)
                Dim IStockClass As IStockInventarioPermanente = ProcessServer.ExecuteTask(Of ProcesoStocks.DataCreateIStockClass, IStockInventarioPermanente)(AddressOf ProcesoStocks.CreateIStockInventarios, datIStock, services)
                If Not IStockClass Is Nothing Then
                    If Not FechaOld Is Nothing Then IStockClass.ValidarPeriodoCerrado(FechaOld, services)
                    If Not data.Current("FechaAlbaran") Is Nothing Then IStockClass.ValidarPeriodoCerrado(data.Current("FechaAlbaran"), services)
                End If
            End If

            ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ProcesoComunes.AsignarEjercicioContableAlbaran, data.Current, services)
            ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoComunes.CambioMoneda, data, services)
        Else
            Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
            If AppParams.GestionInventarioPermanente Then
                Dim datIStock As New ProcesoStocks.DataCreateIStockClass(InventariosPermanentes.ENSAMBLADO_INV_PERMANENTE_STOCKS, InventariosPermanentes.CLASE_INV_PERMANENTE_STOCKS)
                Dim IStockClass As IStockInventarioPermanente = ProcessServer.ExecuteTask(Of ProcesoStocks.DataCreateIStockClass, IStockInventarioPermanente)(AddressOf ProcesoStocks.CreateIStockInventarios, datIStock, services)
                If Not IStockClass Is Nothing Then
                    If Not FechaOld Is Nothing Then IStockClass.ValidarPeriodoCerrado(FechaOld, services)
                End If
            End If
            data.Current("IDEjercicio") = DBNull.Value
        End If
    End Sub

    <Task()> Public Shared Sub CambioCentroGestion(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If New Parametro().AlmacenCentroGestionActivo() AndAlso Length(data.Current("IDCentroGestion")) > 0 Then
            Dim f As New Filter
            f.Add("IDCentroGestion", FilterOperator.Equal, data.Current("IDCentroGestion"))
            f.Add("Principal", FilterOperator.Equal, True)
            f.Add("Activo", FilterOperator.Equal, True)
            Dim dt As DataTable = New Almacen().Filter(f)
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                data.Current("IDAlmacen") = dt.Rows(0)("IDAlmacen")
            Else
                data.Current("IDAlmacen") = New Parametro().AlmacenPredeterminado()
            End If
        Else
            data.Current("IDAlmacen") = New Parametro().AlmacenPredeterminado()
        End If
    End Sub

    <Task()> Public Shared Sub CambioAlmacen(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDAlmacen")) > 0 Then
            Dim dt As DataTable = New Almacen().SelOnPrimaryKey(data.Current("IDAlmacen"))
            If IsNothing(dt) OrElse dt.Rows.Count = 0 Then
                ApplicationService.GenerateError("El Almacen | no existe.", Quoted(data.Current("IDAlmacen")))
            End If
        End If
    End Sub
    '//CambioCondicionPago para las líneas
    <Task()> Public Shared Sub CambioCondicionPagoLineas(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        ' If IsDBNull(data.Current("IDLineaPedido")) Then
        Dim CPOld As String = Nz(data.Current("IDCondicionPago"), String.Empty)
        data.Current(data.ColumnName) = data.Value
        Dim CPNew As String = Nz(data.Current("IDCondicionPago"), String.Empty)
        If CPOld <> CPNew Then

            Dim CondicionesPago As EntityInfoCache(Of CondicionPagoInfo) = services.GetService(Of EntityInfoCache(Of CondicionPagoInfo))()
            Dim CPagoInfoOld As CondicionPagoInfo = CondicionesPago.GetEntity(CPOld)
            Dim CPagoInfoNew As CondicionPagoInfo = CondicionesPago.GetEntity(CPNew)

            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
            Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.Current("IDArticulo"))
            If Not ArtInfo.Especial Then
                Dim dblDtoProntoPagoRecFinanOld As Double
                If Not CPagoInfoOld Is Nothing Then
                    If CPagoInfoOld.DtoProntoPago <> 0 Then
                        dblDtoProntoPagoRecFinanOld = CPagoInfoOld.DtoProntoPago
                        'ElseIf CPagoInfoOld.RecFinan <> 0 Then
                        '    dblDtoProntoPagoRecFinanOld = -1 * CPagoInfoOld.RecFinan
                    Else
                        dblDtoProntoPagoRecFinanOld = 0
                    End If
                End If
                If dblDtoProntoPagoRecFinanOld = Nz(data.Current("DtoProntoPago"), 0) Then
                    '//El DtoProntoPago de las líneas era el de la condición de pago anterior. Si se hubiese manipulado, lo mantenemos.
                    Dim dblDtoProntoPagoRecFinanNew As Double
                    If CPagoInfoNew.DtoProntoPago <> 0 Then
                        dblDtoProntoPagoRecFinanNew = CPagoInfoNew.DtoProntoPago
                        'ElseIf CPagoInfoNew.RecFinan <> 0 Then
                        '    dblDtoProntoPagoRecFinanNew = -1 * CPagoInfoNew.RecFinan
                    Else
                        dblDtoProntoPagoRecFinanNew = 0
                    End If
                    data.Current("DtoProntoPago") = dblDtoProntoPagoRecFinanNew
                    ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf NegocioGeneral.CambioPrecio, data, services)
                End If
            End If
        End If
        'End If
    End Sub
#End Region

#Region " Validaciones Contador "

    <Task()> Public Shared Sub ValidarContadorObligatorio(ByVal IDContador As String, ByVal services As ServiceProvider)
        If Length(IDContador) = 0 Then
            ApplicationService.GenerateError("El Contador es un dato obligatorio.")
        End If
    End Sub

    <Serializable()> _
    Public Class DataValidarContadorEntidad
        Public IDContador As String
        Public Entidad As String

        Public Sub New(ByVal IDContador As String, ByVal Entidad As String)
            Me.IDContador = IDContador
            Me.Entidad = Entidad
        End Sub
    End Class

    <Task()> Public Shared Sub ValidarContadorEntidad(ByVal data As DataValidarContadorEntidad, ByVal services As ServiceProvider)
        If Length(data.IDContador) > 0 Then
            Dim f As New Filter
            f.Add(New StringFilterItem("Entidad", FilterOperator.Equal, data.Entidad))
            f.Add(New StringFilterItem("IDContador", FilterOperator.Equal, data.IDContador))
            Dim dtC As DataTable = New EntidadContador().Filter(f)
            If IsNothing(dtC) OrElse dtC.Rows.Count = 0 Then
                ApplicationService.GenerateError("El contador {0} no está definido para la entidad {1}.", Quoted(data.IDContador), Quoted(data.Entidad))
            End If
        End If
    End Sub

#End Region

#Region " Tratamiento resultado de procesos "

    <Task()> Public Shared Sub AddFacturaCreadaResultadoFacturacion(ByVal Doc As DocumentCabLin, ByVal services As ServiceProvider)
        Dim fra As ResultFacturacion = services.GetService(Of ResultFacturacion)()
        ReDim Preserve fra.Log.CreatedElements(fra.Log.CreatedElements.Length)
        fra.Log.CreatedElements(fra.Log.CreatedElements.Length - 1) = New CreateElement
        fra.Log.CreatedElements(fra.Log.CreatedElements.Length - 1).IDElement = Doc.HeaderRow("IDFactura")
        fra.Log.CreatedElements(fra.Log.CreatedElements.Length - 1).NElement = Doc.HeaderRow("NFactura")
    End Sub

    <Task()> Public Shared Function GetResultadoFacturacion(ByVal data As Object, ByVal services As ServiceProvider) As ResultFacturacion
        Dim rslt As ResultFacturacion = services.GetService(Of ResultFacturacion)()
        Return rslt
    End Function

#End Region

#Region " Facturaciones desde Obras "

    <Task()> Public Shared Sub AsignarRetencionPorGarantiaObra(ByVal Doc As DocumentCabLin, ByVal services As ServiceProvider)
        If (Doc.HeaderRow.RowState = DataRowState.Added AndAlso Length(Doc.HeaderRow("IDObra")) > 0) OrElse _
           (Doc.HeaderRow.RowState = DataRowState.Modified AndAlso Length(Doc.HeaderRow("IDObra")) > 0 AndAlso Nz(Doc.HeaderRow("IDObra"), 0) <> Nz(Doc.HeaderRow("IDObra", DataRowVersion.Original), 0)) Then

            Dim TipoRetencion As enumTipoRetencion?
            Dim FechaRetencion As Date?
            Dim PorcentajeRet As Double
            'Dim ImporteRet As Double?
            Dim dtOrigenRetencion As DataTable
            If TypeOf Doc Is DocumentoFacturaCompra AndAlso Length(Doc.HeaderRow("IDProveedor")) > 0 Then
                Dim OP As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraProveedor")
                Dim f As New Filter
                f.Add(New NumberFilterItem("IDObra", Doc.HeaderRow("IDObra")))
                f.Add(New StringFilterItem("IDProveedor", Doc.HeaderRow("IDProveedor")))
                f.Add(New IsNullFilterItem("TipoRetencion", False))
                f.Add(New NumberFilterItem("Retencion", FilterOperator.NotEqual, 0))
                dtOrigenRetencion = OP.Filter(f)
            ElseIf TypeOf Doc Is DocumentoFacturaVenta Then
                Dim OC As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraCabecera")
                Dim f As New Filter
                f.Add(New NumberFilterItem("IDObra", Doc.HeaderRow("IDObra")))
                f.Add(New IsNullFilterItem("TipoRetencion", False))
                f.Add(New NumberFilterItem("Retencion", FilterOperator.NotEqual, 0))
                dtOrigenRetencion = OC.Filter(f)
            End If


            If Not dtOrigenRetencion Is Nothing AndAlso dtOrigenRetencion.Rows.Count > 0 Then
                If Length(dtOrigenRetencion.Rows(0)("TipoRetencion")) > 0 Then TipoRetencion = CInt(dtOrigenRetencion.Rows(0)("TipoRetencion"))
                If Length(dtOrigenRetencion.Rows(0)("Retencion")) > 0 AndAlso Length(dtOrigenRetencion.Rows(0)("Impuestos")) > 0 AndAlso dtOrigenRetencion.Rows(0)("Impuestos") = TipoRetencionImpuestos.DespuesImpuestos Then
                    PorcentajeRet = dtOrigenRetencion.Rows(0)("Retencion")
                End If
                If Length(dtOrigenRetencion.Rows(0)("FechaRetencion")) > 0 Then
                    FechaRetencion = dtOrigenRetencion.Rows(0)("FechaRetencion")
                ElseIf Length(dtOrigenRetencion.Rows(0)("Periodo")) > 0 AndAlso Length(dtOrigenRetencion.Rows(0)("TipoPeriodo")) > 0 Then
                    Dim periodo As DateInterval
                    Select Case dtOrigenRetencion.Rows(0)("TipoPeriodo")
                        Case enumcpPeriodo.cpDia
                            periodo = DateInterval.Day
                        Case enumcpPeriodo.cpSemana
                            periodo = DateInterval.WeekOfYear
                        Case enumcpPeriodo.cpMes
                            periodo = DateInterval.Month
                        Case enumcpPeriodo.cpAño
                            periodo = DateInterval.Year
                    End Select
                    FechaRetencion = DateAdd(periodo, dtOrigenRetencion.Rows(0)("Periodo"), Doc.Fecha)
                End If

                Dim EntidadCabecera As BusinessHelper = BusinessHelper.CreateBusinessObject(Doc.EntidadCabecera)
                If Not TipoRetencion Is Nothing Then EntidadCabecera.ApplyBusinessRule("TipoRetencion", TipoRetencion, Doc.HeaderRow)
                EntidadCabecera.ApplyBusinessRule("Retencion", PorcentajeRet, Doc.HeaderRow)
                If Not FechaRetencion Is Nothing Then
                    EntidadCabecera.ApplyBusinessRule("FechaRetencion", FechaRetencion, Doc.HeaderRow)
                Else
                    ApplicationService.GenerateError("No se ha indicado una fecha para la retención, o bien, un período para calcularla.")
                End If
            End If

        End If
    End Sub


#End Region

#Region " Cambios de Empresa "

    <Task()> Public Shared Sub EstablecerEmpresaSecundaria(ByVal data As Object, ByVal services As ServiceProvider)
        AdminData.CommitTx(True)
        Dim BDInfo As DataBasesDatosMultiempresa = services.GetService(Of DataBasesDatosMultiempresa)()
        '   If Length(BDInfo.IDBaseDatosSecundaria) > 0 Then AdminData.SetSessionConnection(BDInfo.BaseDatosSecundaria)
        If Length(BDInfo.IDBaseDatosSecundaria) > 0 Then AdminData.SetCurrentConnection(BDInfo.IDBaseDatosSecundaria)
    End Sub

    <Task()> Public Shared Sub EstablecerEmpresaPrincipal(ByVal data As Object, ByVal services As ServiceProvider)
        AdminData.CommitTx(True)
        Dim BDInfo As DataBasesDatosMultiempresa = services.GetService(Of DataBasesDatosMultiempresa)()
        ' If Length(BDInfo.IDBaseDatosPrincipal) > 0 Then AdminData.SetSessionConnection(BDInfo.BaseDatosPrincipal)
        If Length(BDInfo.IDBaseDatosPrincipal) > 0 Then AdminData.SetCurrentConnection(BDInfo.IDBaseDatosPrincipal)
    End Sub

    <Task()> Public Shared Sub GetDescripcionBasesDatosMultiempresa(ByVal data As DataBasesDatosMultiempresa, ByVal services As ServiceProvider)
        If Not data Is Nothing AndAlso Length(data.IDBaseDatosPrincipal) > 0 AndAlso Length(data.IDBaseDatosSecundaria) > 0 Then
            Dim dt As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf General.DatosSistema.GetDataBases, Nothing, services)
            Dim f As New Filter
            f.Add(New GuidFilterItem("IDBaseDatos", data.IDBaseDatosPrincipal))
            Dim Filtro As String = f.Compose(New AdoFilterComposer)
            Dim adr() As DataRow = dt.Select(Filtro)
            If Not adr Is Nothing AndAlso adr.Length > 0 Then
                data.DescBaseDatosPrincipal = adr(0)("DescBaseDatos")
                data.BaseDatosPrincipal = adr(0)("BaseDatos")
            End If

            f.Clear()
            f.Add(New GuidFilterItem("IDBaseDatos", data.IDBaseDatosSecundaria))
            Filtro = f.Compose(New AdoFilterComposer)
            adr = dt.Select(Filtro)
            If Not adr Is Nothing AndAlso adr.Length > 0 Then
                data.DescBaseDatosSecundaria = adr(0)("DescBaseDatos")
                data.BaseDatosSecundaria = adr(0)("BaseDatos")
            End If
        End If
    End Sub


#End Region

#Region " Gestión de Doble Unidad "

    <Task()> Public Shared Function AplicarSegundaUnidad(ByVal IDArticulo As String, ByVal services As ServiceProvider) As Boolean
        AplicarSegundaUnidad = False
        Dim AppParams As ParametroStocks = services.GetService(Of ParametroStocks)()
        If AppParams.GestionDobleUnidad Then
            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
            Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(IDArticulo)
            AplicarSegundaUnidad = (ArtInfo.GestionStock AndAlso ArtInfo.TieneSegundaUnidad AndAlso Not ArtInfo.GestionPorNumeroSerie AndAlso Not ArtInfo.KitVenta)
        End If
    End Function

    <Task()> Public Shared Sub ValidarFactorDobleUnidad(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Current("IDArticulo")) > 0 Then
            If ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf AplicarSegundaUnidad, data.Current("IDArticulo"), services) Then
                Dim IDUDInterna As String
                If data.Current.ContainsKey("IDUdInterna") AndAlso Length(data.Current("IDUdInterna")) > 0 Then
                    IDUDInterna = data.Current("IDUdInterna")
                End If

                Dim datFactor As New Articulo.DataFactorDobleUnidad(data.Current("IDArticulo"), (data.ColumnName = "QInterna" OrElse data.ColumnName = "Cantidad" OrElse data.ColumnName = "IDUdMedida"), (data.ColumnName = "QInterna2"), IDUDInterna)
                Dim Factor As Double = ProcessServer.ExecuteTask(Of Articulo.DataFactorDobleUnidad, Double)(AddressOf Articulo.FactorDobleUnidad, datFactor, services)
                If Factor <> 0 Then
                    data.Current("QInterna2") = Nz(data.Current("QInterna"), 0) * Factor
                Else
                    If data.Current("IDUdMedida") & String.Empty = data.Current("IDUDInterna2") & String.Empty Then
                        data.Current("QInterna2") = data.Current("Cantidad")
                    End If
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioSegundaUnidad(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDArticulo")) > 0 Then
            If ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf AplicarSegundaUnidad, data.Current("IDArticulo"), services) Then
                Dim IDUDInterna As String
                If data.Current.ContainsKey("IDUdInterna") AndAlso Length(data.Current("IDUdInterna")) > 0 Then
                    IDUDInterna = data.Current("IDUdInterna")
                End If
                Dim datFactor As New Articulo.DataFactorDobleUnidad(data.Current("IDArticulo"), (data.ColumnName = "QInterna"), (data.ColumnName = "QInterna2"), IDUDInterna)
                Dim Factor As Double = ProcessServer.ExecuteTask(Of Articulo.DataFactorDobleUnidad, Double)(AddressOf Articulo.FactorDobleUnidad, datFactor, services)
                If Factor <> 0 Then
                    data.Current("QInterna") = Nz(data.Current("QInterna2"), 0) * Factor
                    data.Current("CambioQInterna") = True
                    If data.Current.ContainsKey("IDUdMedida") AndAlso data.Current.ContainsKey("IDUdInterna") Then
                        If data.Current("IDUdMedida") = data.Current("IDUDInterna") Then
                            data.Current("Cantidad") = data.Current("QInterna")
                        End If
                    End If
                    If Nz(data.Current("Cantidad"), 0) <> 0 Then
                        data.Current("Factor") = data.Current("QInterna") / data.Current("Cantidad")
                    End If
                End If
            End If
        End If
    End Sub


#End Region
#Region " Gestión cambio base imponible para ajustar facturas "
    <Task()> Public Shared Sub CalcularIVA(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim TiposIVA As EntityInfoCache(Of TipoIvaInfo) = services.GetService(Of EntityInfoCache(Of TipoIvaInfo))()
        ' HistoricoTipoIVA
        Dim TIVAInfo As TipoIvaInfo = TiposIVA.GetEntity(data.Current("IDTipoIva"), data.Context("FechaFactura"))

        'valor por defecto
        data.Current(data.ColumnName) = data.Value
        data.Current("ImpIVA") = data.Current("BaseImponible") * TIVAInfo.Factor / 100
        If data.Context.ContainsKey("IDMoneda") AndAlso Length(data.Context("IDMoneda")) > 0 Then
            Dim ValAyB As New ValoresAyB(data.Current, data.Context("IDMoneda"), data.Context("CambioA"), data.Context("CambioB"))
            ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)
        End If

    End Sub
#End Region

    <Serializable()> _
    Public Class DataFechasCierreValidadas
        Public FechasCerradas As New Dictionary(Of Date, Boolean)
    End Class

    <Task()> Public Shared Function AlbaranEnPeriodoCerrado(ByVal FechaAlbaran As Date, ByVal services As ServiceProvider) As Boolean
        Dim FechasValidadas As DataFechasCierreValidadas = services.GetService(Of DataFechasCierreValidadas)()
        If FechasValidadas.FechasCerradas.ContainsKey(FechaAlbaran) Then
            Return FechasValidadas.FechasCerradas(FechaAlbaran)
        Else
            Dim dt As DataTable = AdminData.GetData("vNegCierreInventarioFechaUltimoCierre", , "TOP 1 FechaHasta", "FechaHasta DESC, FechaDesde DESC")
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                If FechaAlbaran <= dt.Rows(0)("FechaHasta") Then
                    FechasValidadas.FechasCerradas(FechaAlbaran) = True
                    Return True
                End If
            Else
                FechasValidadas.FechasCerradas(FechaAlbaran) = False
            End If
        End If
    End Function

    <Task()> Public Shared Function AlbaranEnPeriodoCerradoDoc(ByVal Doc As DocumentCabLin, ByVal services As ServiceProvider) As Boolean
        AlbaranEnPeriodoCerradoDoc = False
        Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
        If Not AppParams.ActualizarPrecioAlbaranPeriodoCerrado Then
            Dim LineasModificadasActualizadas As List(Of DataRow) = (From c In Doc.dtLineas _
                                                                     Where c.RowState = DataRowState.Modified AndAlso _
                                                                          Not c.IsNull("EstadoStock") AndAlso _
                                                                          c("EstadoStock") = EstadoStock.Actualizado Select c).ToList
            If Not LineasModificadasActualizadas Is Nothing AndAlso LineasModificadasActualizadas.Count > 0 Then

                If ProcessServer.ExecuteTask(Of Date, Boolean)(AddressOf ProcesoComunes.AlbaranEnPeriodoCerrado, Doc.HeaderRow("FechaAlbaran"), services) Then
                    AlbaranEnPeriodoCerradoDoc = True
                    ApplicationService.GenerateError("El período está cerrado. No se permiten movimientos nuevos con fecha documento anterior a la fecha del ultimo período cerrado.")
                End If
            End If

        End If
    End Function


    '#Region " Validaciones de IVA "

    '    <Task()> Public Shared Sub ValidarDocumentoIdentificativoFV(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider)
    '        If Length(data("IDCliente")) = 0 Then Exit Sub

    '        Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
    '        Dim ClteInfo As ClienteInfo = Clientes.GetEntity(data("IDCliente"))
    '        If Length(data("IDPais")) > 0 AndAlso Length(data("CifCliente")) > 0 AndAlso Length(ClteInfo.CifCliente) > 0 AndAlso data("CifCliente") <> ClteInfo.CifCliente Then
    '            Dim info As New DataDocIdentificacion(data("CifCliente"), data("IDPais"), ClteInfo.TipoDocIdentidad)
    '            info = ProcessServer.ExecuteTask(Of DataDocIdentificacion, DataDocIdentificacion)(AddressOf Comunes.ValidarDocumentoIdentificativo, info, services)
    '            If Not info.EsCorrecto Then
    '                ApplicationService.GenerateError("El Documento introducido no es un '|'. Intoduzca uno correcto o cambie de tipo de documento", info.TipoDocumento)
    '            End If
    '        End If
    '    End Sub

    '    <Task()> Public Shared Sub ValidarDocumentoIdentificativoFC(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider)
    '        If Length(data("IDProveedor")) = 0 Then Exit Sub

    '        Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
    '        Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(data("IDProveedor"))

    '        If Length(data("IDPais")) > 0 AndAlso Length(data("CifProveedor")) > 0 AndAlso Length(ProvInfo.CifProveedor) > 0 AndAlso data("CifProveedor") <> ProvInfo.CifProveedor Then
    '            Dim info As New DataDocIdentificacion(data("CifProveedor"), data("IDPais"), data("TipoDocIdentidad"))
    '            info = ProcessServer.ExecuteTask(Of DataDocIdentificacion, DataDocIdentificacion)(AddressOf Comunes.ValidarDocumentoIdentificativo, info, services)
    '            If Not info.EsCorrecto Then
    '                ApplicationService.GenerateError("El Documento introducido no es un '|'. Intoduzca uno correcto o cambie de tipo de documento", info.TipoDocumento)
    '            End If
    '        End If
    '    End Sub

    '    <Task()> Public Shared Sub ValidarIVASDocFV(ByVal Doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)
    '        Dim ValidarFacturaIntracomunitaria As Boolean

    '        Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
    '        Dim ClteInfo As ClienteInfo = Clientes.GetEntity(Doc.HeaderRow("IDCliente"))
    '        If ClteInfo.TipoDocIdentidad = enumTipoDocIdent.NIFOperadorIntra Then
    '            ValidarFacturaIntracomunitaria = True
    '        End If
    '        Dim datVal As New DataValidarIVASDoc(Doc, Circuito.Ventas, ValidarFacturaIntracomunitaria)
    '        ProcessServer.ExecuteTask(Of DataValidarIVASDoc)(AddressOf ValidarIVASDoc, datVal, services)
    '    End Sub

    '    <Task()> Public Shared Sub ValidarIVASDocFC(ByVal doc As DocumentoFacturaCompra, ByVal services As ServiceProvider)
    '        Dim ValidarFacturaIntracomunitaria As Boolean

    '        Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
    '        ValidarFacturaIntracomunitaria = AppParams.OperadorVIES
    '        Dim datVal As New DataValidarIVASDoc(doc, Circuito.Compras, ValidarFacturaIntracomunitaria)
    '        ProcessServer.ExecuteTask(Of DataValidarIVASDoc)(AddressOf ValidarIVASDoc, datVal, services)
    '    End Sub

    '    Public Class DataValidarIVASDoc
    '        Public Doc As DocumentCabLin
    '        Public Circuito As Circuito
    '        Public ValidarFacturaIntracomunitaria As Boolean

    '        Public Sub New(ByVal Doc As DocumentCabLin, ByVal Circuito As Circuito, ByVal ValidarFacturaIntracomunitaria As Boolean)
    '            Me.Doc = Doc
    '            Me.Circuito = Circuito
    '            Me.ValidarFacturaIntracomunitaria = ValidarFacturaIntracomunitaria
    '        End Sub
    '    End Class
    '    <Task()> Public Shared Sub ValidarIVASDoc(ByVal data As DataValidarIVASDoc, ByVal services As ServiceProvider)
    '        If data.Doc Is Nothing OrElse data.Circuito < 0 Then Exit Sub
    '        Dim IDPais As String = data.Doc.HeaderRow("IDPais") & String.Empty
    '        If Length(IDPais) > 0 Then
    '            Dim dtPais As DataTable = New Pais().SelOnPrimaryKey(IDPais)
    '            If dtPais.Rows.Count > 0 Then
    '                Dim NFactura As String = data.Doc.HeaderRow("NFactura")
    '                Dim FraServicios As Boolean = Nz(data.Doc.HeaderRow("Servicios349"), False)


    '                Dim dtIVAS As DataTable = New TipoIva().Filter
    '                Dim IVASExportacion As List(Of DataRow) = (From tIVA In dtIVAS Where tIVA.RowState <> DataRowState.Deleted AndAlso Not tIVA.IsNull("Tipo") AndAlso tIVA("Tipo") = enumTipoIVA.OperacionesExportacion Select tIVA).ToList
    '                Dim IVASImportacion As List(Of DataRow) = (From tIVA In dtIVAS Where tIVA.RowState <> DataRowState.Deleted AndAlso Not tIVA.IsNull("Tipo") AndAlso (tIVA("Tipo") = enumTipoIVA.OperacionesImportacionesBienesInversion OrElse _
    '                                                                                     tIVA("Tipo") = enumTipoIVA.OperacionesImportacionesCorrientes) _
    '                                                           Select tIVA).ToList
    '                Dim IVASIntracomunitarios As List(Of DataRow) = (From tIVA In dtIVAS Where Not tIVA.RowState <> DataRowState.Deleted AndAlso tIVA.IsNull("Tipo") AndAlso (tIVA("Tipo") = enumTipoIVA.OperacionesIntracomunitariasCorrientes OrElse _
    '                                                                                    tIVA("Tipo") = enumTipoIVA.OperacionesIntracomunitariasInversion) _
    '                                                                Select tIVA).ToList
    '                Dim IVASInteriores As List(Of DataRow) = (From tIVA In dtIVAS Where tIVA.RowState <> DataRowState.Deleted AndAlso Not tIVA.IsNull("Tipo") AndAlso (tIVA("Tipo") = enumTipoIVA.OperacionesInterioresCorrientes OrElse _
    '                                                                                    tIVA("Tipo") = enumTipoIVA.OperacionesInterioresBienesInversion) _
    '                                                          Select tIVA).ToList
    '                Dim IVASInterioresCorrientes As List(Of DataRow) = (From tIVA In dtIVAS Where tIVA.RowState <> DataRowState.Deleted AndAlso Not tIVA.IsNull("Tipo") AndAlso (tIVA("Tipo") = enumTipoIVA.OperacionesInterioresCorrientes) _
    '                                                          Select tIVA).ToList
    '                Dim IVASInterioresBienesInversion As List(Of DataRow) = (From tIVA In dtIVAS Where tIVA.RowState <> DataRowState.Deleted AndAlso Not tIVA.IsNull("Tipo") AndAlso (tIVA("Tipo") = enumTipoIVA.OperacionesInterioresBienesInversion) _
    '                                                          Select tIVA).ToList
    '                Dim IVASSujetoPasivo As List(Of DataRow) = (From tIVA In dtIVAS Where tIVA.RowState <> DataRowState.Deleted AndAlso Not tIVA.IsNull("Tipo") AndAlso (tIVA("Tipo") = enumTipoIVA.OperacionesInversionSujetoPasivo) _
    '                                                         Select tIVA).ToList

    '                Dim IVASNoDeducibles As List(Of DataRow) = (From tIVA In dtIVAS Where tIVA.RowState <> DataRowState.Deleted AndAlso Not tIVA.IsNull("Tipo") AndAlso (tIVA("Tipo") = enumTipoIVA.OperacionesNoDeducibles) _
    '                                                        Select tIVA).ToList


    '                Dim Extranjero As Boolean = Nz(dtPais.Rows(0)("Extranjero"), False)
    '                Dim CEE As Boolean = Nz(dtPais.Rows(0)("CEE"), False)
    '                Dim CCM As Boolean = Nz(dtPais.Rows(0)("CanariasCeutaMelilla"), False)

    '                Dim dtBasesImponibles As DataTable
    '                Select Case data.Circuito
    '                    Case Circuito.Ventas
    '                        dtBasesImponibles = CType(data.Doc, DocumentoFacturaVenta).dtFVBI
    '                    Case Circuito.Compras
    '                        dtBasesImponibles = CType(data.Doc, DocumentoFacturaCompra).dtFCBI
    '                End Select

    '                If Not dtBasesImponibles Is Nothing Then
    '                    Dim lstBasesImponibles As List(Of DataRow) = (From b In dtBasesImponibles Where b.RowState <> DataRowState.Deleted Select b).ToList
    '                    Dim strIVAsSExportacion As String = String.Empty : Dim MensajeExportacion As String = String.Empty
    '                    Dim strIVAsImportacion As String = String.Empty : Dim MensajeImportacion As String = String.Empty
    '                    Dim strIVAsIntracomunitarios As String = String.Empty : Dim MensajeIntracomunitarios As String = String.Empty
    '                    Dim strIVAsInteriores As String = String.Empty : Dim MensajeInteriores As String = String.Empty
    '                    Dim strIVAsSujetoPasivo As String = String.Empty : Dim MensajeSujetoPasivo As String = String.Empty
    '                    Dim strIVAsNoDeducibles As String = String.Empty : Dim MensajeNoDeducibles As String = String.Empty


    '                    If Not IVASExportacion Is Nothing Then
    '                        Dim ExistenExportacion As List(Of DataRow) = (From tIVA In IVASExportacion Join fIVA In lstBasesImponibles On UCase(tIVA("IDTipoIVA")) Equals UCase(fIVA("IDTipoIVA")) Select fIVA).ToList
    '                        If Not ExistenExportacion Is Nothing AndAlso ExistenExportacion.Count > 0 Then
    '                            Dim IVAS As List(Of String) = (From iva In ExistenExportacion Select CStr(iva("IDTipoIVA")) Distinct).ToList
    '                            strIVAsSExportacion = Strings.Join(IVAS.ToArray, ",")
    '                            MensajeExportacion = "La Factura {0} tiene IVAs de Tipo Exportación y no cumple los requisitos para ello.{1}Tipos de IVA: {2}"
    '                        End If
    '                    End If
    '                    If Not IVASImportacion Is Nothing Then
    '                        Dim ExistenImportacion As List(Of DataRow) = (From tIVA In IVASImportacion Join fIVA In lstBasesImponibles On UCase(tIVA("IDTipoIVA")) Equals UCase(fIVA("IDTipoIVA")) Select fIVA).ToList
    '                        If Not ExistenImportacion Is Nothing AndAlso ExistenImportacion.Count > 0 Then
    '                            Dim IVAS As List(Of String) = (From iva In ExistenImportacion Select CStr(iva("IDTipoIVA")) Distinct).ToList
    '                            strIVAsImportacion = Strings.Join(IVAS.ToArray, ",")
    '                            MensajeImportacion = "La Factura {0} tiene IVAs de Tipo Importación y no cumple los requisitos para ello.{1}Tipos de IVA: {2}"
    '                        End If
    '                    End If
    '                    If Not IVASIntracomunitarios Is Nothing Then
    '                        Dim ExistenIntracomunitarios As List(Of DataRow) = (From tIVA In IVASIntracomunitarios Join fIVA In lstBasesImponibles On UCase(tIVA("IDTipoIVA")) Equals UCase(fIVA("IDTipoIVA")) Select fIVA).ToList
    '                        If Not ExistenIntracomunitarios Is Nothing AndAlso ExistenIntracomunitarios.Count > 0 Then
    '                            Dim IVAS As List(Of String) = (From iva In ExistenIntracomunitarios Select CStr(iva("IDTipoIVA")) Distinct).ToList
    '                            strIVAsIntracomunitarios = Strings.Join(IVAS.ToArray, ",")
    '                            MensajeIntracomunitarios = "La Factura {0} tiene IVAs de Tipo Intracomunitario y no cumple los requisitos para ello.{1}Tipos de IVA: {2}"
    '                        End If
    '                    End If
    '                    If Not IVASInteriores Is Nothing Then
    '                        Dim ExistenInteriores As List(Of DataRow) = (From tIVA In IVASInteriores Join fIVA In lstBasesImponibles On UCase(tIVA("IDTipoIVA")) Equals UCase(fIVA("IDTipoIVA")) Select fIVA).ToList
    '                        If Not ExistenInteriores Is Nothing AndAlso ExistenInteriores.Count > 0 Then
    '                            Dim IVAS As List(Of String) = (From iva In ExistenInteriores Select CStr(iva("IDTipoIVA")) Distinct).ToList
    '                            strIVAsInteriores = Strings.Join(IVAS.ToArray, ",")
    '                            MensajeInteriores = "La Factura {0} tiene IVAs de Tipo Interiores y no cumple los requisitos para ello.{1}Tipos de IVA: {2}"
    '                        End If
    '                    End If
    '                    If Not IVASSujetoPasivo Is Nothing Then
    '                        Dim ExistenSujetoPasivo As List(Of DataRow) = (From tIVA In IVASSujetoPasivo Join fIVA In lstBasesImponibles On UCase(tIVA("IDTipoIVA")) Equals UCase(fIVA("IDTipoIVA")) Select fIVA).ToList
    '                        If Not ExistenSujetoPasivo Is Nothing AndAlso ExistenSujetoPasivo.Count > 0 Then
    '                            Dim IVAS As List(Of String) = (From iva In ExistenSujetoPasivo Select CStr(iva("IDTipoIVA")) Distinct).ToList
    '                            strIVAsSujetoPasivo = Strings.Join(IVAS.ToArray, ",")
    '                            MensajeSujetoPasivo = "La Factura {0} tiene IVAs de Tipo Sujeto Pasivo y no cumple los requisitos para ello.{1}Tipos de IVA: {2}"
    '                        End If
    '                    End If
    '                    If Not IVASNoDeducibles Is Nothing Then
    '                        Dim ExistenNoDeducibles As List(Of DataRow) = (From tIVA In IVASNoDeducibles Join fIVA In lstBasesImponibles On UCase(tIVA("IDTipoIVA")) Equals UCase(fIVA("IDTipoIVA")) Select fIVA).ToList
    '                        If Not ExistenNoDeducibles Is Nothing AndAlso ExistenNoDeducibles.Count > 0 Then
    '                            Dim IVAS As List(Of String) = (From iva In ExistenNoDeducibles Select CStr(iva("IDTipoIVA")) Distinct).ToList
    '                            strIVAsNoDeducibles = Strings.Join(IVAS.ToArray, ",")
    '                            MensajeNoDeducibles = "La Factura {0} tiene IVAs de No Deducibles y no cumple los requisitos para ello.{1}Tipos de IVA: {2}"
    '                        End If
    '                    End If


    '                    If (Not Extranjero) AndAlso CEE AndAlso (Not CCM) Then
    '                        If Length(strIVAsImportacion) > 0 Then ApplicationService.GenerateError(MensajeImportacion, Quoted(NFactura), vbNewLine, strIVAsImportacion)
    '                        If Length(strIVAsIntracomunitarios) > 0 Then ApplicationService.GenerateError(MensajeIntracomunitarios, Quoted(NFactura), vbNewLine, strIVAsIntracomunitarios)
    '                        Select Case data.Circuito
    '                            Case Circuito.Ventas
    '                                If Length(strIVAsNoDeducibles) > 0 Then ApplicationService.GenerateError(MensajeNoDeducibles, Quoted(NFactura), vbNewLine, strIVAsNoDeducibles)
    '                                If Length(strIVAsSExportacion) > 0 Then ApplicationService.GenerateError(MensajeExportacion, Quoted(NFactura), vbNewLine, strIVAsSExportacion)
    '                            Case Circuito.Compras
    '                                If Length(strIVAsSExportacion) > 0 Then
    '                                    If Not IVASExportacion Is Nothing Then
    '                                        Dim ExistenExportacion As List(Of DataRow) = (From tIVA In IVASExportacion Join fIVA In lstBasesImponibles On UCase(tIVA("IDTipoIVA")) Equals UCase(fIVA("IDTipoIVA")) Where Nz(fIVA("BaseImponible"), 0) <> 0 Select fIVA).ToList
    '                                        If Not ExistenExportacion Is Nothing AndAlso ExistenExportacion.Count > 0 Then
    '                                            Dim IVAS As List(Of String) = (From iva In ExistenExportacion Select CStr(iva("IDTipoIVA")) Distinct).ToList
    '                                            strIVAsSExportacion = Strings.Join(IVAS.ToArray, ",")
    '                                            MensajeExportacion = "La Factura {0} tiene IVAs de Tipo Exportación y no cumple los requisitos para ello.{1}Tipos de IVA: {2}"
    '                                            ApplicationService.GenerateError(MensajeExportacion, Quoted(NFactura), vbNewLine, strIVAsSExportacion)
    '                                        End If
    '                                    End If
    '                                End If
    '                        End Select
    '                    ElseIf (Not Extranjero) AndAlso CEE AndAlso CCM Then
    '                        Select Case data.Circuito
    '                            Case Circuito.Ventas
    '                                If Length(strIVAsImportacion) > 0 Then ApplicationService.GenerateError(MensajeImportacion, Quoted(NFactura), vbNewLine, strIVAsImportacion)
    '                                If Length(strIVAsNoDeducibles) > 0 Then ApplicationService.GenerateError(MensajeNoDeducibles, Quoted(NFactura), vbNewLine, strIVAsNoDeducibles)
    '                                If Length(strIVAsSujetoPasivo) > 0 Then ApplicationService.GenerateError(MensajeSujetoPasivo, Quoted(NFactura), vbNewLine, strIVAsSujetoPasivo)

    '                                If Length(strIVAsInteriores) > 0 Then
    '                                    If Not FraServicios Then
    '                                        ApplicationService.GenerateError(MensajeInteriores, Quoted(NFactura), vbNewLine, strIVAsInteriores)
    '                                    Else
    '                                        If Not IVASInterioresBienesInversion Is Nothing Then
    '                                            Dim ExistenInterioresBienesInversion As List(Of DataRow) = (From tIVA In IVASInterioresBienesInversion Join fIVA In lstBasesImponibles On UCase(tIVA("IDTipoIVA")) Equals UCase(fIVA("IDTipoIVA")) Select fIVA).ToList
    '                                            If Not ExistenInterioresBienesInversion Is Nothing AndAlso ExistenInterioresBienesInversion.Count > 0 Then
    '                                                Dim IVAS As List(Of String) = (From iva In ExistenInterioresBienesInversion Select CStr(iva("IDTipoIVA")) Distinct).ToList
    '                                                strIVAsInteriores = Strings.Join(IVAS.ToArray, ",")
    '                                                MensajeInteriores = "La Factura {0} tiene IVAs de Tipo Interiores de Bienes de Inversión y no cumple los requisitos para ello.{1}Tipos de IVA: {2}"
    '                                                ApplicationService.GenerateError(MensajeInteriores, Quoted(NFactura), vbNewLine, strIVAsInteriores)
    '                                            End If
    '                                        End If
    '                                    End If
    '                                End If
    '                            Case Circuito.Compras
    '                                If Length(strIVAsSExportacion) > 0 Then ApplicationService.GenerateError(MensajeExportacion, Quoted(NFactura), vbNewLine, strIVAsSExportacion)
    '                                If Length(strIVAsInteriores) > 0 Then ApplicationService.GenerateError(MensajeInteriores, Quoted(NFactura), vbNewLine, strIVAsInteriores)
    '                                If Length(strIVAsNoDeducibles) > 0 Then ApplicationService.GenerateError(MensajeNoDeducibles, Quoted(NFactura), vbNewLine, strIVAsNoDeducibles)
    '                                If Length(strIVAsSujetoPasivo) > 0 AndAlso Not FraServicios Then ApplicationService.GenerateError(MensajeSujetoPasivo, Quoted(NFactura), vbNewLine, strIVAsSujetoPasivo)
    '                        End Select
    '                        If Length(strIVAsIntracomunitarios) > 0 Then ApplicationService.GenerateError(MensajeIntracomunitarios, Quoted(NFactura), vbNewLine, strIVAsIntracomunitarios)
    '                        'If Length(strIVAsImportacion) > 0 Then ApplicationService.GenerateError(MensajeImportacion, Quoted(NFactura), vbNewLine, strIVAsImportacion)

    '                    ElseIf Extranjero AndAlso CEE AndAlso (Not CCM) Then

    '                        If Length(strIVAsSExportacion) > 0 Then ApplicationService.GenerateError(MensajeExportacion, Quoted(NFactura), vbNewLine, strIVAsSExportacion)
    '                        If Length(strIVAsImportacion) > 0 Then ApplicationService.GenerateError(MensajeImportacion, Quoted(NFactura), vbNewLine, strIVAsImportacion)

    '                        If data.ValidarFacturaIntracomunitaria Then
    '                            If Length(strIVAsInteriores) > 0 Then ApplicationService.GenerateError(MensajeInteriores, Quoted(NFactura), vbNewLine, strIVAsInteriores)
    '                            If Length(strIVAsSujetoPasivo) > 0 Then ApplicationService.GenerateError(MensajeSujetoPasivo, Quoted(NFactura), vbNewLine, strIVAsSujetoPasivo)
    '                        Else
    '                            If Length(strIVAsIntracomunitarios) > 0 Then ApplicationService.GenerateError(MensajeIntracomunitarios, Quoted(NFactura), vbNewLine, strIVAsIntracomunitarios)
    '                        End If
    '                    ElseIf Extranjero AndAlso (Not CEE) AndAlso (Not CCM) Then
    '                        Select Case data.Circuito
    '                            Case Circuito.Ventas
    '                                If Length(strIVAsImportacion) > 0 Then ApplicationService.GenerateError(MensajeImportacion, Quoted(NFactura), vbNewLine, strIVAsImportacion)
    '                                If Length(strIVAsNoDeducibles) > 0 Then ApplicationService.GenerateError(MensajeNoDeducibles, Quoted(NFactura), vbNewLine, strIVAsNoDeducibles)
    '                                If Length(strIVAsSujetoPasivo) > 0 Then ApplicationService.GenerateError(MensajeSujetoPasivo, Quoted(NFactura), vbNewLine, strIVAsSujetoPasivo)
    '                            Case Circuito.Compras
    '                                If Length(strIVAsNoDeducibles) > 0 Then ApplicationService.GenerateError(MensajeNoDeducibles, Quoted(NFactura), vbNewLine, strIVAsNoDeducibles)
    '                                If Length(strIVAsSExportacion) > 0 Then ApplicationService.GenerateError(MensajeExportacion, Quoted(NFactura), vbNewLine, strIVAsSExportacion)
    '                                If Length(strIVAsSujetoPasivo) > 0 AndAlso Not FraServicios Then ApplicationService.GenerateError(MensajeSujetoPasivo, Quoted(NFactura), vbNewLine, strIVAsSujetoPasivo)
    '                        End Select
    '                        If Length(strIVAsInteriores) > 0 Then ApplicationService.GenerateError(MensajeInteriores, Quoted(NFactura), vbNewLine, strIVAsInteriores)
    '                        If Length(strIVAsIntracomunitarios) > 0 Then ApplicationService.GenerateError(MensajeIntracomunitarios, Quoted(NFactura), vbNewLine, strIVAsIntracomunitarios)
    '                    End If
    '                End If
    '            End If
    '        End If
    '    End Sub

    '#End Region

End Class

