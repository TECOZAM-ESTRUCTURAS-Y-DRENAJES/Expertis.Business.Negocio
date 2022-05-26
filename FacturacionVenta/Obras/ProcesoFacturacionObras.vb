Public Class ProcesoFacturacionObras

#Region " Agrupaciones "

    <Task()> Public Shared Function CreateGroupHelper(ByVal data As DataGetGroupColumns, ByVal services As ServiceProvider) As GroupHelper
        Dim ColsGroupSinAgrupar() As DataColumn = ProcessServer.ExecuteTask(Of DataGetGroupColumns, DataColumn())(AddressOf ProcesoFacturacionObras.GetGroupColumns, data, services)
        Return New GroupHelper(ColsGroupSinAgrupar, data.GrprUser)
    End Function

    <Task()> Public Shared Function CreateGroupHelperObraPromo(ByVal data As DataGetGroupColumns, ByVal services As ServiceProvider) As GroupHelper
        Dim ColsGroupSinAgrupar() As DataColumn = ProcessServer.ExecuteTask(Of DataGetGroupColumns, DataColumn())(AddressOf ProcesoFacturacionObras.GetGroupColumnsObraPromo, data, services)
        Return New GroupHelper(ColsGroupSinAgrupar, data.GrprUser)
    End Function

    <Task()> Public Shared Function AgruparObras(ByVal data As DataPrcFacturacionObras, ByVal services As ServiceProvider) As FraCabObra()
        Dim oResult(-1) As FraCabObra
        Select Case data.TipoFacturacion
            Case enumTipoFactura.tfObra
                oResult = ProcessServer.ExecuteTask(Of DataPrcFacturacionObras, FraCabObra())(AddressOf AgruparVencimientos, data, services)
            Case enumTipoFactura.tfCertificacion
                oResult = ProcessServer.ExecuteTask(Of DataPrcFacturacionObras, FraCabObraCertificacion())(AddressOf AgruparVencimientosCertificacion, data, services)
            Case enumTipoFactura.tfPromocionObra, enumTipoFactura.tfPromocionObraFinal
                oResult = ProcessServer.ExecuteTask(Of DataPrcFacturacionObras, FraCabObraPromo())(AddressOf AgruparVencimientosObraPromo, data, services)
        End Select
        Return oResult
    End Function

    <Task()> Public Shared Function AgruparVencimientos(ByVal data As DataPrcFacturacionObras, ByVal services As ServiceProvider) As FraCabObra()
        Const cnViewName As String = "vNegFacturacionVencimientosObra"
        If data.IDVencimiento.Length > 0 Then
            Dim oFltr As New Filter
            oFltr.Add(New InListFilterItem("IDVencimiento", data.IDVencimiento, FilterType.Numeric))
            oFltr.Add(New BooleanFilterItem("Facturado", False))

            Dim dtLineas As DataTable = New BE.DataEngine().Filter(cnViewName, oFltr)

            Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
            Dim strCentroGestion As String = AppParams.CentroGestion
            Dim strCondicionPago As String = AppParams.CondicionPago
            Dim oGrprUsr As New GroupUserObraVencimiento(strCentroGestion, strCondicionPago)

            Dim oGrpr(enummcAgrupFacturaObra.mcObraPedidoClte) As GroupHelper
            Dim dataColGroups As New DataGetGroupColumns(dtLineas, enummcAgrupFacturaObra.mcSinAgrupar, oGrprUsr)
            oGrpr(enummcAgrupFacturaObra.mcSinAgrupar) = ProcessServer.ExecuteTask(Of DataGetGroupColumns, GroupHelper)(AddressOf ProcesoFacturacionObras.CreateGroupHelper, dataColGroups, services)
            dataColGroups = New DataGetGroupColumns(dtLineas, enummcAgrupFacturaObra.mcCliente, oGrprUsr)
            oGrpr(enummcAgrupFacturaObra.mcCliente) = ProcessServer.ExecuteTask(Of DataGetGroupColumns, GroupHelper)(AddressOf ProcesoFacturacionObras.CreateGroupHelper, dataColGroups, services)
            dataColGroups = New DataGetGroupColumns(dtLineas, enummcAgrupFacturaObra.mcObra, oGrprUsr)
            oGrpr(enummcAgrupFacturaObra.mcObra) = ProcessServer.ExecuteTask(Of DataGetGroupColumns, GroupHelper)(AddressOf ProcesoFacturacionObras.CreateGroupHelper, dataColGroups, services)
            dataColGroups = New DataGetGroupColumns(dtLineas, enummcAgrupFacturaObra.mcObraTrabajo, oGrprUsr)
            oGrpr(enummcAgrupFacturaObra.mcObraTrabajo) = ProcessServer.ExecuteTask(Of DataGetGroupColumns, GroupHelper)(AddressOf ProcesoFacturacionObras.CreateGroupHelper, dataColGroups, services)
            dataColGroups = New DataGetGroupColumns(dtLineas, enummcAgrupFacturaObra.mcObraPedidoClte, oGrprUsr)
            oGrpr(enummcAgrupFacturaObra.mcObraPedidoClte) = ProcessServer.ExecuteTask(Of DataGetGroupColumns, GroupHelper)(AddressOf ProcesoFacturacionObras.CreateGroupHelper, dataColGroups, services)

            For Each rwLin As DataRow In dtLineas.Select("", "IDVencimiento")
                If data.TipoAgrupacion = -1 Then
                    oGrpr(rwLin("AgrupFacturaObra")).Group(rwLin)
                Else
                    oGrpr(data.TipoAgrupacion).Group(rwLin)
                End If
            Next

            Return oGrprUsr.Fras
        End If
    End Function

    <Task()> Public Shared Function AgruparVencimientosCertificacion(ByVal data As DataPrcFacturacionObras, ByVal services As ServiceProvider) As FraCabObraCertificacion()
        Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
        Dim strCentroGestion As String = AppParams.CentroGestion

        Dim FilView As New Filter
        FilView.Add(New InListFilterItem("IDCertificacion", data.IDVencimiento, FilterType.Numeric))
        Dim DtLineas As DataTable = New BE.DataEngine().Filter("vFrmCIObraFacturacionCertificacion", FilView)

        Dim oGrprUsr As New GroupUserObraCertificacion(strCentroGestion)

        Dim oGrpr(enummcAgrupFacturaObra.mcObraPedidoClte) As GroupHelper
        'Dim dataColGroups As New DataGetGroupColumns(DtLineas, enummcAgrupFacturaObra.mcSinAgrupar, oGrprUsr)
        'oGrpr(enummcAgrupFacturaObra.mcSinAgrupar) = ProcessServer.ExecuteTask(Of DataGetGroupColumns, GroupHelper)(AddressOf ProcesoFacturacionObras.CreateGroupHelper, dataColGroups, services)
        'dataColGroups = New DataGetGroupColumns(DtLineas, enummcAgrupFacturaObra.mcCliente, oGrprUsr)
        'oGrpr(enummcAgrupFacturaObra.mcCliente) = ProcessServer.ExecuteTask(Of DataGetGroupColumns, GroupHelper)(AddressOf ProcesoFacturacionObras.CreateGroupHelper, dataColGroups, services)
        Dim dataColGroups As New DataGetGroupColumns(DtLineas, enummcAgrupFacturaObra.mcObra, oGrprUsr)
        oGrpr(0) = ProcessServer.ExecuteTask(Of DataGetGroupColumns, GroupHelper)(AddressOf ProcesoFacturacionObras.CreateGroupHelper, dataColGroups, services)
        'dataColGroups = New DataGetGroupColumns(DtLineas, enummcAgrupFacturaObra.mcObraTrabajo, oGrprUsr)
        'oGrpr(enummcAgrupFacturaObra.mcObraTrabajo) = ProcessServer.ExecuteTask(Of DataGetGroupColumns, GroupHelper)(AddressOf ProcesoFacturacionObras.CreateGroupHelper, dataColGroups, services)
        'dataColGroups = New DataGetGroupColumns(DtLineas, enummcAgrupFacturaObra.mcObraPedidoClte, oGrprUsr)
        'oGrpr(enummcAgrupFacturaObra.mcObraPedidoClte) = ProcessServer.ExecuteTask(Of DataGetGroupColumns, GroupHelper)(AddressOf ProcesoFacturacionObras.CreateGroupHelper, dataColGroups, services)

        For Each rwLin As DataRow In DtLineas.Select
            oGrpr(0).Group(rwLin)
        Next

        Return oGrprUsr.fras
    End Function

    <Task()> Public Shared Function AgruparVencimientosObraPromo(ByVal data As DataPrcFacturacionObras, ByVal services As ServiceProvider) As FraCabObraPromo()
        Const cnViewName As String = "vNegFacturacionVencimientosObraPromo"

        Dim oFltr As New Filter
        oFltr.Add(New InListFilterItem("IDLocalVencimiento", data.IDVencimiento, FilterType.Numeric))
        oFltr.Add(New BooleanFilterItem("Facturado", False))
        Dim dtLineas As DataTable = New BE.DataEngine().Filter(cnViewName, oFltr)

        Dim dv As DataView = dtLineas.DefaultView
        dv.RowFilter = "TipoFactura= " & enumfvcTipoFactura.fvcFinal
        Dim TipoFraPromo As enumfvcTipoFactura
        If dv.Count > 0 Then
            TipoFraPromo = enumfvcTipoFactura.fvcFinal
        Else
            TipoFraPromo = dtLineas.Rows(0)("TipoFactura")
        End If

        Dim oGrpr(0) As GroupHelper
        Dim oGrprUser As New GroupUserObraPromo(TipoFraPromo)

        Dim dataColGroups As New DataGetGroupColumns(dtLineas, data.TipoAgrupacion, oGrprUser)
        oGrpr(0) = ProcessServer.ExecuteTask(Of DataGetGroupColumns, GroupHelper)(AddressOf ProcesoFacturacionObras.CreateGroupHelperObraPromo, dataColGroups, services)

        For Each oRw As DataRow In dtLineas.Select(Nothing, "IDLocalVencimiento")
            oGrpr(0).Group(oRw)
        Next

        Return oGrprUser.fras
    End Function

    Public Class DataGetGroupColumns
        Public Datos As DataTable
        Public Agrupacion As enummcAgrupFacturaObra
        Public GrprUser As IGroupUser

        Public Sub New(ByVal Datos As DataTable, ByVal Agrupacion As enummcAgrupFacturaObra, ByVal GrprUser As IGroupUser)
            Me.Datos = Datos
            Me.Agrupacion = Agrupacion
            Me.GrprUser = GrprUser
        End Sub
    End Class
    <Task()> Public Shared Function GetGroupColumns(ByVal data As DataGetGroupColumns, ByVal services As ServiceProvider) As DataColumn()
        Dim columns(-1) As DataColumn
        If data.Agrupacion <> enummcAgrupFacturaObra.mcObraPedidoClte Then
            ReDim Preserve columns(columns.Length)
            columns(columns.Length - 1) = data.Datos.Columns("IDCliente")
        End If
        ReDim Preserve columns(columns.Length)
        columns(columns.Length - 1) = data.Datos.Columns("IDFormaPago")

        ReDim Preserve columns(columns.Length)
        columns(columns.Length - 1) = data.Datos.Columns("IDCondicionPago")

        ReDim Preserve columns(columns.Length)
        columns(columns.Length - 1) = data.Datos.Columns("IDDiaPago")

        ReDim Preserve columns(columns.Length)
        columns(columns.Length - 1) = data.Datos.Columns("IDMoneda")

        ReDim Preserve columns(columns.Length)
        columns(columns.Length - 1) = data.Datos.Columns("FechaVencimiento")

        Select Case data.Agrupacion
            Case enummcAgrupFacturaObra.mcObra
                ReDim Preserve columns(columns.Length)
                columns(columns.Length - 1) = data.Datos.Columns("IDObra")
                ReDim Preserve columns(columns.Length)
                columns(columns.Length - 1) = data.Datos.Columns("SeguroCambio")
            Case enummcAgrupFacturaObra.mcObraTrabajo
                ReDim Preserve columns(columns.Length)
                columns(columns.Length - 1) = data.Datos.Columns("IDObra")
                ReDim Preserve columns(columns.Length)
                columns(columns.Length - 1) = data.Datos.Columns("IDTrabajo")
                ReDim Preserve columns(columns.Length)
                columns(columns.Length - 1) = data.Datos.Columns("SeguroCambio")
                ReDim Preserve columns(columns.Length)
            Case enummcAgrupFacturaObra.mcObraPedidoClte
                ReDim Preserve columns(columns.Length)
                columns(columns.Length - 1) = data.Datos.Columns("IDObra")
                ReDim Preserve columns(columns.Length)
                'columns(columns.Length - 1) = data.Datos.Columns("DescTrabajo")
                columns(columns.Length - 1) = data.Datos.Columns("PedidoCliente")
            Case enummcAgrupFacturaObra.mcSinAgrupar
                Return New DataColumn() {data.Datos.Columns("IDVencimiento")}
        End Select

        Return columns
    End Function

    <Task()> Public Shared Function GetGroupColumnsObraPromo(ByVal data As DataGetGroupColumns, ByVal services As ServiceProvider) As DataColumn()
        If data.Agrupacion = enumTipoFacturacionPromociones.PorCliente Then
            Return New DataColumn() {data.Datos.Columns("IDCliente"), data.Datos.Columns("IDMoneda"), data.Datos.Columns("IDFormaPago"), data.Datos.Columns("IDCondicionPago")}
        ElseIf data.Agrupacion = enumTipoFacturacionPromociones.PorLocal Then
            Return New DataColumn() {data.Datos.Columns("IDLocal")}
        End If
    End Function

#End Region

#Region " Cabecera "

    <Task()> Public Shared Sub AsignarDireccion(ByVal doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If TypeOf doc.Cabecera Is FraCabObraCertificacion Then
            ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarDireccion, doc, services)
        Else
            If CType(doc.Cabecera, FraCabVencimiento).TipoMnto = enumTipoObra.tpalquiler Then
                Dim dataDireccion As New ClienteDireccion.DataDirecEnvio(doc.HeaderRow("IDCliente"), enumcdTipoDireccion.cdDireccionEnvio)
                Dim dtDireccion As DataTable = ProcessServer.ExecuteTask(Of ClienteDireccion.DataDirecEnvio, DataTable)(AddressOf ClienteDireccion.ObtenerDireccionEnvio, dataDireccion, services)
                If Not IsNothing(dtDireccion) AndAlso dtDireccion.Rows.Count > 0 Then
                    doc.HeaderRow("IDDireccion") = dtDireccion.Rows(0)("IDDireccion")
                End If
            Else
                ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarDireccion, doc, services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarDiaPago(ByVal doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If Not doc.Cabecera Is Nothing AndAlso TypeOf doc.Cabecera Is FraCabObra Then
            doc.HeaderRow("IDDiaPago") = CType(doc.Cabecera, FraCabObra).IDDiaPago
        End If
    End Sub

    <Task()> Public Shared Sub AsignarTipoFactura(ByVal doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If TypeOf doc.Cabecera Is FraCabObraPromo Then
            doc.HeaderRow("TipoFactura") = CType(doc.Cabecera, FraCabObraPromo).TipoFactura
        End If
    End Sub

    <Task()> Public Shared Sub AsignarCambiosMoneda(ByVal Doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If TypeOf Doc.Cabecera Is FraCabObra Then
            If CType(Doc.Cabecera, FraCabObra).SeguroCambio Then
                Doc.HeaderRow("CambioA") = CType(Doc.Cabecera, FraCabObra).CambioA
                Doc.HeaderRow("CambioB") = CType(Doc.Cabecera, FraCabObra).CambioB
            Else
                ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoComunes.AsignarCambiosMoneda, Doc, services)
            End If
        Else
            ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoComunes.AsignarCambiosMoneda, Doc, services)
        End If
    End Sub

#End Region

#Region " Creación de Lineas de Facturas de Obras "

#Region " Creación de Lineas desde Hitos (Vencimientos) "

    <Task()> Public Shared Sub CrearLineasDesdeObras(ByVal docfactura As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        Dim InfoObras As ProcessInfoFraObras = services.GetService(Of ProcessInfoFraObras)()
        Select Case InfoObras.TipoFacturacion
            Case enumTipoFactura.tfObra
                ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf CrearLineasDesdeVencimientos, docfactura, services)
            Case enumTipoFactura.tfCertificacion
                ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf CrearLineasDesdeCertificacion, docfactura, services)
            Case enumTipoFactura.tfPromocionObra, enumTipoFactura.tfPromocionObraFinal
                ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf CrearLineaDesdeFacturaObraPromo, docfactura, services)
        End Select
    End Sub

    Public Class DataAsignarValoresGenerales
        Inherits DataDocRowOrigen
        Public TipoLineaPorDefecto As String

        Public Sub New(ByVal data As DataDocRowOrigen, ByVal TipoLineaPorDefecto As String)
            MyBase.New(data.Doc, data.RowOrigen, data.RowDestino)
            Me.TipoLineaPorDefecto = TipoLineaPorDefecto
        End Sub
    End Class
    Public Class DataAsignarValoresObraVencimientos
        Inherits DataAsignarValoresGenerales

        Public FraLin As FraLinVencimiento

        Public Sub New(ByVal data As DataDocRowOrigen, ByVal TipoLineaPorDefecto As String, ByVal FraLin As FraLinVencimiento)
            MyBase.New(data, TipoLineaPorDefecto)
            Me.FraLin = FraLin
        End Sub
    End Class

    <Task()> Public Shared Sub CrearLineasDesdeVencimientos(ByVal docFactura As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If Not docFactura Is Nothing Then
            Dim TipoLineaPredet As String = ProcessServer.ExecuteTask(Of Object, String)(AddressOf TipoLinea.TipoLineaPorDefecto, Nothing, services)

            Dim ids(CType(docFactura.Cabecera, FraCabVencimiento).Lineas.Length - 1) As Object
            For i As Integer = 0 To CType(docFactura.Cabecera, FraCabVencimiento).Lineas.Length - 1
                ids(i) = CType(CType(docFactura.Cabecera, FraCabVencimiento).Lineas(i), FraLinVencimiento).IdVencimiento
            Next
            Dim oFltr As New Filter
            oFltr.Add(New InListFilterItem("IDVencimiento", ids, FilterType.Numeric))
            oFltr.Add(New BooleanFilterItem("Facturado", False))

            Dim ObraTbjoFrcn As BusinessHelper = BusinessHelper.CreateBusinessObject(AdminData.GetEntityInfo("ObraTrabajoFacturacion"))
            Dim dtVencimiento As DataTable = ObraTbjoFrcn.Filter(oFltr, "IDTrabajo, IdVencimiento")

            If dtVencimiento.Rows.Count > 0 Then
                Dim fvl As New FacturaVentaLinea
                Dim intIDOrdenLinea As Integer = 0
                For Each drVencimiento As DataRow In dtVencimiento.Rows
                    Dim FraLin As FraLinVencimiento
                    For j As Integer = 0 To CType(docFactura.Cabecera, FraCabVencimiento).Lineas.Length - 1
                        If CType(CType(docFactura.Cabecera, FraCabVencimiento).Lineas(j), FraLinVencimiento).IdVencimiento = drVencimiento("IDVencimiento") Then
                            FraLin = CType(CType(docFactura.Cabecera, FraCabVencimiento).Lineas(j), FraLinVencimiento)
                        End If
                    Next

                    Dim drlinea As DataRow = fvl.AddNewForm.Rows(0)
                    Dim datosLin As New DataDocRowOrigen(docFactura, drVencimiento, drlinea)
                    Dim ValGen As New DataAsignarValoresObraVencimientos(datosLin, TipoLineaPredet, FraLin)

                    ProcessServer.ExecuteTask(Of DataAsignarValoresGenerales)(AddressOf AsignarValoresGenerales, ValGen, services)
                    ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarVencimientoOrigen, datosLin, services)
                    ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarDatosArticuloVencimiento, datosLin, services)
                    ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarTipoIVA, datosLin, services)
                    ProcessServer.ExecuteTask(Of DataAsignarValoresObraVencimientos)(AddressOf AsignarPedidoCliente, ValGen, services)
                    ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarCentroGestion, datosLin, services)
                    ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarDatosTrabajo, datosLin, services)
                    ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarCContableCliente, datosLin, services)
                    ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarDatosObra, datosLin, services)
                    ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarDatosAlbaranVentaOrigen, datosLin, services)
                    ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarDatosAlquiler, datosLin, services)
                    ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarDatosMantenimiento, datosLin, services)
                    ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarLote, datosLin, services)
                    ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarSeguimiento, datosLin, services)
                    ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarConceptos, datosLin, services)
                    ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarCantidadesPrecioImporteVencimiento, datosLin, services)
                    ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarDescuentos, datosLin, services)

                    intIDOrdenLinea += 1
                    drlinea("IDOrdenLinea") = intIDOrdenLinea

                    docFactura.dtLineas.Rows.Add(drlinea.ItemArray)
                Next

                'ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf CalcularSegurosTasasAlquiler, docFactura, services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarValoresGenerales(ByVal data As DataAsignarValoresGenerales, ByVal services As ServiceProvider)
        data.RowDestino("IDLineaFactura") = AdminData.GetAutoNumeric
        data.RowDestino("IDFactura") = data.Doc.HeaderRow("IDFactura")
        data.RowDestino("IDTipoLinea") = data.TipoLineaPorDefecto
    End Sub

    <Task()> Public Shared Sub AsignarVencimientoOrigen(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        data.RowDestino("IDVencimiento") = data.RowOrigen("IDVencimiento")
    End Sub

    <Task()> Public Shared Sub AsignarDatosArticuloVencimiento(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        data.RowDestino("IDArticulo") = data.RowOrigen("IDArticulo")
        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.RowOrigen("IDArticulo"))
        If Length(data.RowOrigen("DescVencimiento")) = 0 Then
            data.RowDestino("DescArticulo") = ArtInfo.DescArticulo
        Else
            data.RowDestino("DescArticulo") = data.RowOrigen("DescVencimiento")
        End If
        data.RowDestino("CContable") = data.RowOrigen("IDCContable")
        data.RowDestino("IDUDMedida") = ArtInfo.IDUDVenta
        data.RowDestino("IDUDInterna") = ArtInfo.IDUDInterna
    End Sub

    <Task()> Public Shared Sub AsignarTipoIVA(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        data.RowDestino("IDTipoIva") = data.RowOrigen("IDTipoIva")
    End Sub

    <Task()> Public Shared Sub AsignarPedidoCliente(ByVal data As DataAsignarValoresObraVencimientos, ByVal services As ServiceProvider)
        If Length(data.FraLin.NumeroPedido) > 0 Then data.RowDestino("PedidoCliente") = data.FraLin.NumeroPedido
    End Sub

    <Task()> Public Shared Sub AsignarCentroGestion(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        data.RowDestino("IDCentroGestion") = data.Doc.Cabecera.IDCentroGestion
    End Sub

    <Task()> Public Shared Sub AsignarDatosTrabajo(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        If Length(data.RowOrigen("IDLineaMaterial")) = 0 AndAlso Length(data.RowOrigen("IDTrabajo")) > 0 Then  '//Asignar Datos Trabajo
            Dim OT As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraTrabajo")
            Dim dr As DataRow = OT.GetItemRow(data.RowOrigen("IDTrabajo"))
            If Length(data.RowDestino("CContable")) = 0 Then data.RowDestino("CContable") = dr("CContable")
            If Length(dr("NumeroPedido")) > 0 Then
                data.RowDestino("PedidoCliente") = dr("NumeroPedido")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarCContableCliente(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        If Length(data.RowDestino("CContable")) = 0 Then
            Dim Ctx As New BusinessData(data.Doc.HeaderRow)
            Dim datos As New BusinessRuleData("CContable", data.RowDestino("CContable"), New DataRowPropertyAccessor(data.RowDestino), Ctx)
            ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoComercial.TratarCContable, datos, services)
            data.RowDestino("CContable") = datos.Current("CContable")
        End If
    End Sub

    <Task()> Public Shared Sub AsignarDatosObra(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        data.RowDestino("IDObra") = data.RowOrigen("IDObra")
        data.RowDestino("IDTrabajo") = data.RowOrigen("IDTrabajo")
        data.RowDestino("IDLineaMaterial") = data.RowOrigen("IDLineaMaterial")
        data.RowDestino("IDLineaMod") = data.RowOrigen("IDLineaMod")
        data.RowDestino("IDLineaCentro") = data.RowOrigen("IDLineaCentro")
        data.RowDestino("IDLineaGasto") = data.RowOrigen("IDLineaGasto")
        data.RowDestino("IDLineaVarios") = data.RowOrigen("IDLineaVarios")
    End Sub

    <Task()> Public Shared Sub AsignarDatosAlbaranVentaOrigen(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        data.RowDestino("IDALbaran") = data.RowOrigen("IDAlbaranVentaOrigen")
        data.RowDestino("IDLineaAlbaran") = data.RowOrigen("IdLineaAlbaranVentaOrigen")
    End Sub

    <Task()> Public Shared Sub AsignarDatosAlquiler(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        data.RowDestino("IDLineaAlbaranRetorno") = data.RowOrigen("IDLineaAlbaranRetorno")
        data.RowDestino("IDAlbaranRetorno") = data.RowOrigen("IDAlbaranRetorno")
        data.RowDestino("FechaDesdeAlquiler") = data.RowOrigen("FechaDesdeAlquiler")
        data.RowDestino("FechaHastaAlquiler") = data.RowOrigen("FechaHastaAlquiler")

        ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarTipoFacturaAlquiler, data, services)
    End Sub

    <Task()> Public Shared Sub AsignarTipoFacturaAlquiler(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        data.RowDestino("TipoFactAlquiler") = data.RowOrigen("TipoFactAlquiler")
    End Sub

    <Task()> Public Shared Sub AsignarDatosMantenimiento(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        data.RowDestino("IDMntoOTControl") = data.RowOrigen("IDMntoOTControl")
        data.RowDestino("IdOT") = data.RowOrigen("IdOT")
    End Sub

    <Task()> Public Shared Sub AsignarLote(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        data.RowDestino("Lote") = data.RowOrigen("Lote")
    End Sub

    <Task()> Public Shared Sub AsignarSeguimiento(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        data.RowDestino("Texto") = data.RowOrigen("Seguimiento")
    End Sub

    <Task()> Public Shared Sub AsignarConceptos(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        data.RowDestino("IDConcepto") = data.RowOrigen("IDConcepto")
        data.RowDestino("DescConcepto") = data.RowOrigen("DescConcepto")
    End Sub

    <Task()> Public Shared Sub AsignarCantidadesPrecioImporteVencimiento(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        '//Cantidades, precio e Importe
        data.RowDestino("Importe") = data.RowOrigen("ImpVencimiento")
        data.RowDestino("UdValoracion") = 1
        Dim StDatos As New ArticuloUnidadAB.DatosFactorConversion(data.RowOrigen("IDArticulo"), data.RowDestino("IDUDMedida"), data.RowDestino("IDUDInterna"))
        data.RowDestino("Factor") = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, StDatos, services)
        If data.RowOrigen("TipoFactura") = enumotfTipoFactura.otfAlquiler Then
            data.RowDestino("Cantidad") = Nz(data.RowOrigen("QAFacturar"), 0)
            data.RowDestino("Precio") = Nz(data.RowOrigen("PrecioVto"), 0)
            data.RowDestino("QTiempo") = data.RowOrigen("QTiempo")
            data.RowDestino("QTiempoIncidencias") = Nz(data.RowOrigen("QTiempoIncidencias"), 0)
        Else
            Dim Parametros As ParametroGeneral = services.GetService(Of ParametroGeneral)()
            If data.RowOrigen("QInterna") = 0 Or Parametros.AplicacionGestionAlquiler Then
                data.RowDestino("Cantidad") = 1
                data.RowDestino("Precio") = data.RowDestino("Importe")
            Else
                data.RowDestino("Cantidad") = Nz(data.RowOrigen("QInterna"), 1)
                data.RowDestino("Precio") = Nz(data.RowOrigen("PrecioVto"), 0)
            End If

            data.RowDestino("QTiempo") = 1
            data.RowDestino("QTiempoIncidencias") = 0
        End If
        data.RowDestino("QInterna") = data.RowDestino("Factor") * data.RowDestino("Cantidad")
    End Sub

    <Task()> Public Shared Sub AsignarDescuentos(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        data.RowDestino("Dto1") = data.RowOrigen("Dto1")
        data.RowDestino("Dto2") = data.RowOrigen("Dto2")
        data.RowDestino("Dto3") = data.RowOrigen("Dto3")
        ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarDescuentosCabecera, data, services)
    End Sub

    <Task()> Public Shared Sub AsignarDescuentosCabecera(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        data.RowDestino("Dto") = data.Doc.HeaderRow("DtoFactura")
        data.RowDestino("DtoProntoPago") = data.Doc.HeaderRow("DtoProntoPago")
    End Sub

    <Task()> Public Shared Sub NuevaLineaFacturaObraEntregaCuenta(ByVal Doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If (Doc.HeaderRow.RowState = DataRowState.Added AndAlso Length(Doc.HeaderRow("IDObra")) > 0) OrElse _
          (Doc.HeaderRow.RowState = DataRowState.Modified AndAlso Length(Doc.HeaderRow("IDObra")) > 0 AndAlso Nz(Doc.HeaderRow("IDObra"), 0) <> Nz(Doc.HeaderRow("IDObra", DataRowVersion.Original), 0)) Then
            '//Si tenemos retención por Garantía de una Obra
            If Length(Doc.HeaderRow("TipoRetencion")) > 0 AndAlso Length(Doc.HeaderRow("Retencion")) > 0 AndAlso Length(Doc.HeaderRow("FechaRetencion")) > 0 Then
                If Not IsNothing(Doc.dtLineas) AndAlso Doc.dtLineas.Rows.Count > 0 Then
                    '//Creamos una Entrega nueva de Tipo Retención.
                    Dim EC As New EntregasACuenta
                    Dim dtEntregas As DataTable = EC.AddNew
                    Dim datNuevaEntrega As New EntregasACuenta.DatosNuevaEntrega(Doc.HeaderRow.Table, Doc.dtLineas, dtEntregas, Circuito.Ventas)
                    Dim drNuevaEntrega As DataRow = ProcessServer.ExecuteTask(Of EntregasACuenta.DatosNuevaEntrega, DataRow)(AddressOf EntregasACuenta.NuevaEntregaTipoRetencionFacturaObra, datNuevaEntrega, services)
                    If Not drNuevaEntrega Is Nothing Then
                        EC.Update(datNuevaEntrega.DtEntregas)
                        '//Creamos una nueva línea de factura de tipo Retención
                        Dim datFactEntrCta As New EntregasACuenta.DataFacturaVentaEntregas(Doc, dtEntregas)
                        ProcessServer.ExecuteTask(Of EntregasACuenta.DataFacturaVentaEntregas)(AddressOf EntregasACuenta.AddEntregasTipoFacturaVentas, datFactEntrCta, services)
                    End If
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CalcularSegurosTasasAlquiler(ByVal docFactura As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        Dim ProcInfo As ProcessInfoFraObras = services.GetService(Of ProcessInfoFraObras)()
        If ProcInfo.CalculoSeguros Then
            ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf CalculoSegurosTasas.AñadirSegurosFacturacion, docFactura, services)
        End If
        ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf CalculoSegurosTasas.AñadirTasasDeResiduos, docFactura, services)
    End Sub

#End Region

#Region " Creación de Lineas desde Certificaciones "

    Public Class DataAsignarValoresObraCertificacion
        Inherits DataAsignarValoresGenerales

        Public FraLin As FraLinObraCertificacion

        Public Sub New(ByVal data As DataDocRowOrigen, ByVal TipoLineaPorDefecto As String, ByVal FraLin As FraLinObraCertificacion)
            MyBase.New(data, TipoLineaPorDefecto)
            Me.FraLin = FraLin
        End Sub
    End Class

    <Task()> Public Shared Sub CrearLineasDesdeCertificacion(ByVal docFactura As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        Dim ids(CType(docFactura.Cabecera, FraCabObraCertificacion).Lineas.Length - 1) As Object
        For i As Integer = 0 To CType(docFactura.Cabecera, FraCabObraCertificacion).Lineas.Length - 1
            ids(i) = CType(docFactura.Cabecera, FraCabObraCertificacion).Lineas(i).IDTrabajo
        Next

        Dim ObraTbjo As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraTrabajo")
        Dim dtCertificacion As DataTable = ObraTbjo.Filter(New InListFilterItem("IDTrabajo", ids, FilterType.Numeric), "IDTrabajo")
        If dtCertificacion.Rows.Count > 0 Then
            Dim IDObras(-1) As Integer
            Dim intIDOrdenLinea As Integer = 0
            Dim fvl As New FacturaVentaLinea
            Dim Lineas As DataTable = fvl.AddNew
            Dim LineasAux As DataTable = fvl.AddNew
            Dim TipoLineaPredet As String = ProcessServer.ExecuteTask(Of Object, String)(AddressOf TipoLinea.TipoLineaPorDefecto, Nothing, services)
            For Each drCertificacion As DataRow In dtCertificacion.Rows
                Dim FraLin As FraLinObraCertificacion
                For j As Integer = 0 To CType(docFactura.Cabecera, FraCabObraCertificacion).Lineas.Length - 1
                    If CType(docFactura.Cabecera, FraCabObraCertificacion).Lineas(j).IDTrabajo = drCertificacion("IDTrabajo") Then
                        FraLin = CType(docFactura.Cabecera, FraCabObraCertificacion).Lineas(j)
                    End If
                Next

                If Not FraLin Is Nothing Then
                    Dim drlinea As DataRow = fvl.AddNewForm.Rows(0)
                    Dim datosLin As New DataDocRowOrigen(docFactura, drCertificacion, drlinea)
                    Dim ValGen As New DataAsignarValoresObraCertificacion(datosLin, TipoLineaPredet, FraLin)
                    ProcessServer.ExecuteTask(Of DataAsignarValoresObraCertificacion)(AddressOf AsignarValoresGenerales, ValGen, services)
                    ProcessServer.ExecuteTask(Of DataAsignarValoresObraCertificacion)(AddressOf AsignarValoresObraCertificacion, ValGen, services)
                    ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarDescripcionArticulo, datosLin, services)
                    ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarCContable, datosLin, services)
                    ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarUnidades, datosLin, services)
                    ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarTipoFacturaAlquiler, datosLin, services)
                    ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarDescuentosCabecera, datosLin, services)
                    ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarCantidadesPrecioImporteCertificacion, datosLin, services)

                    intIDOrdenLinea = intIDOrdenLinea + 1
                    drlinea("IDOrdenLinea") = intIDOrdenLinea

                    docFactura.dtLineas.Rows.Add(drlinea.ItemArray)

                    If Array.IndexOf(IDObras, FraLin.IDObra) < 0 Then
                        ReDim Preserve IDObras(IDObras.Length)
                        IDObras(IDObras.Length - 1) = FraLin.IDObra
                    End If
                End If
            Next

            Dim datAdic As New DataDatosAdicionales(docFactura, IDObras)
            ProcessServer.ExecuteTask(Of DataDatosAdicionales)(AddressOf NuevaLineaFacturaObraCertificacionDatosAdicionales, datAdic, services)
        End If
    End Sub

    <Task()> Public Shared Sub AsignarValoresObraCertificacion(ByVal data As DataAsignarValoresObraCertificacion, ByVal services As ServiceProvider)
        data.RowDestino("IDObra") = data.RowOrigen("IDObra")
        data.RowDestino("IDTrabajo") = data.RowOrigen("IDTrabajo")
        data.RowDestino("PedidoCliente") = CType(data.FraLin, FraLinObraCertificacion).NumeroPedido
        data.RowDestino("IDCertificacion") = data.FraLin.IDCertificacion
        data.RowDestino("IDTipoIva") = data.FraLin.IDTipoIva
        data.RowDestino("IDCentroGestion") = data.FraLin.IDCentroGestion
        data.RowDestino("Cantidad") = data.FraLin.QCertificada

    End Sub

    <Task()> Public Shared Sub AsignarCContable(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        'Si el trabajo tiene CContable se ha de mantener, sino se tiene que calcular
        'que CContable se ha de utilizar, teniendo en cuenta el Articulo y el Cliente
        If Length(data.RowOrigen("CContable")) > 0 Then
            data.RowDestino("CContable") = data.RowOrigen("CContable")
        Else
            ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarCContableCliente, data, services)
        End If
    End Sub

    <Task()> Public Shared Sub AsignarUnidades(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        data.RowDestino("IDUDMedida") = data.RowOrigen("IDUdMedida")
        data.RowDestino("IDUDInterna") = data.RowOrigen("IDUdMedida")
        data.RowDestino("UDValoracion") = 1
    End Sub

    <Task()> Public Shared Sub AsignarDescripcionArticulo(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        data.RowDestino("IDArticulo") = data.RowOrigen("IDArticulo")
        If Length(data.RowOrigen("DescTrabajo")) = 0 Then
            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
            Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.RowOrigen("IDArticulo"))
            data.RowDestino("DescArticulo") = ArtInfo.DescArticulo
        Else
            data.RowDestino("DescArticulo") = data.RowOrigen("DescTrabajo")
        End If
    End Sub

    <Task()> Public Shared Sub AsignarCantidadesPrecioImporteCertificacion(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        Dim dblPrecio As Double = Nz(data.RowOrigen("ImpPrevTrabajoVentaA"), 0) / Nz(data.Doc.CambioA, 1)
        Dim fvl As New FacturaVentaLinea
        fvl.ApplyBusinessRule("Precio", dblPrecio, data.RowDestino, New BusinessData(data.Doc.HeaderRow))
        'drlinea("Importe") = drlinea("Cantidad") * drlinea("Precio")
        data.RowDestino("PrecioCosteA") = 0
        data.RowDestino("PrecioCosteB") = 0
        ' dblTotalLineas = dblTotalLineas + data.RowDestino("Importe")
        data.RowDestino("Factor") = 1
        data.RowDestino("QInterna") = data.RowDestino("Factor") * data.RowDestino("Cantidad")
    End Sub

    Public Class DataDatosAdicionales
        Public Doc As DocumentoFacturaVenta
        Public IDObras() As Integer

        Public Sub New(ByVal Doc As DocumentoFacturaVenta, ByVal IDObras() As Integer)
            Me.Doc = Doc
            Me.IDObras = IDObras
        End Sub
    End Class
    <Task()> Public Shared Sub NuevaLineaFacturaObraCertificacionDatosAdicionales(ByVal data As DataDatosAdicionales, ByVal services As ServiceProvider)
        If Not IsNothing(data.Doc.dtLineas) AndAlso data.Doc.dtLineas.Rows.Count > 0 Then
            Dim fImpNotNull As New Filter
            fImpNotNull.Add(New IsNullFilterItem("Importe", False))
            Dim TotalLineasNormales As Double = data.Doc.dtLineas.Compute("SUM(Importe)", fImpNotNull.Compose(New AdoFilterComposer))
            Dim fvl As New FacturaVentaLinea
            Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
            Dim IDArticuloFactProy As String = AppParams.ArticuloFacturacionProyectos
            Dim strTipoLinea As String = ProcessServer.ExecuteTask(Of Object, String)(AddressOf TipoLinea.TipoLineaPorDefecto, Nothing, services)

            Dim IDObras(data.IDObras.Length - 1) As Object
            data.IDObras.CopyTo(IDObras, 0)

            Dim fFilterOr As New Filter(FilterUnionOperator.Or)
            Dim f As New Filter
            f.Add(New InListFilterItem("IDObra", IDObras, FilterType.Numeric))
            fFilterOr.Add(New NumberFilterItem("GastosGenerales", FilterOperator.GreaterThan, 0))
            fFilterOr.Add(New NumberFilterItem("BeneficioIndustrial", FilterOperator.GreaterThan, 0))
            fFilterOr.Add(New NumberFilterItem("CoefBaja", FilterOperator.GreaterThan, 0))
            f.Add(fFilterOr)
            Dim Obra As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraCabecera")
            Dim dtObras As DataTable = Obra.Filter(f)
            For Each drObra As DataRow In dtObras.Rows
                Dim drLinea As DataRow = data.Doc.dtLineas.NewRow
                Dim intIDOrdenLinea = data.Doc.dtLineas.Rows.Count + 1
                For i As Integer = 0 To 2
                    '//Se tienen que generar 3 líneas como máximo (Si GastosGenerales,  
                    '//BeneficioIndustrial o CoefBaja es cero no se generará línea)
                    drLinea("IDLineaFactura") = AdminData.GetAutoNumeric
                    drLinea("IDFactura") = data.Doc.HeaderRow("IDFactura")
                    drLinea("IDOrdenLinea") = intIDOrdenLinea
                    drLinea("PedidoCliente") = drObra("NumeroPedido")
                    drLinea("UDValoracion") = 1
                    drLinea("Cantidad") = 1
                    drLinea("QInterna") = 1
                    drLinea("Factor") = 1
                    drLinea("IDTipoLinea") = strTipoLinea
                    drLinea("Regalo") = False
                    drLinea("IDObra") = drObra("IDObra")
                    Dim datosLin As New DataDocRowOrigen(data.Doc, data.Doc.HeaderRow, drLinea)
                    ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarCentroGestion, datosLin, services)
                    ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarDescuentosCabecera, datosLin, services)

                    Dim Context As New BusinessData(data.Doc.HeaderRow)
                    drLinea = fvl.ApplyBusinessRule("IDArticulo", IDArticuloFactProy, drLinea, Context)

                    Dim blnCancel As Double = False
                    Select Case i
                        Case 0
                            If drObra("GastosGenerales") > 0 Then
                                drLinea("Precio") = TotalLineasNormales * drObra("GastosGenerales") / 100
                                drLinea("DescArticulo") = "Gastos Generales"
                            Else
                                blnCancel = True
                            End If
                        Case 1
                            If drObra("BeneficioIndustrial") > 0 Then
                                drLinea("Precio") = TotalLineasNormales * drObra("BeneficioIndustrial") / 100
                                drLinea("DescArticulo") = "Beneficio Industrial"
                            Else
                                blnCancel = True
                            End If
                        Case 2
                            If drObra("CoefBaja") > 0 Then
                                drLinea("Precio") = TotalLineasNormales * -drObra("CoefBaja") / 100
                                drLinea("DescArticulo") = "Coeficiente Baja"
                            Else
                                blnCancel = True
                            End If
                    End Select

                    If Not blnCancel Then
                        intIDOrdenLinea = intIDOrdenLinea + 1
                        data.Doc.dtLineas.Rows.Add(drLinea.ItemArray)
                    End If
                Next
            Next

        End If
    End Sub

#End Region

#Region " Crear Lineas desde Obra Promo "

    <Task()> Public Shared Sub CrearLineaDesdeFacturaObraPromo(ByVal doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If doc.HeaderRow("TipoFactura") = enumfvcTipoFactura.fvcFinal Then
            ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf NuevaLineaFacturaObraPromoFinal, doc, services)
        Else
            ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf NuevaLineaFacturaObraPromoAnticipo, doc, services)
        End If
    End Sub

    <Task()> Public Shared Sub NuevaLineaFacturaObraPromoAnticipo(ByVal doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        Dim IDLocalVencimiento(CType(doc.Cabecera, FraCabObraPromo).Lineas.Length - 1) As Object
        For i As Integer = 0 To CType(doc.Cabecera, FraCabObraPromo).Lineas.Length - 1
            IDLocalVencimiento(i) = CType(doc.Cabecera, FraCabObraPromo).Lineas(i).IDLocalVencimiento
        Next

        Dim IDMonedaA As String = New Parametro().MonedasInternasPredeterminadas.strMonedaA

        Dim f As New Filter
        f.Add(New InListFilterItem("IDLocalVencimiento", IDLocalVencimiento, FilterType.Numeric))
        f.Add(New BooleanFilterItem("Facturado", False))

        Dim dtVencimiento As DataTable = BusinessHelper.CreateBusinessObject("ObraPromoLocalVencimiento").Filter(f)
        If Not dtVencimiento Is Nothing AndAlso dtVencimiento.Rows.Count > 0 Then
            Dim IDOrdenLinea As Integer = 0
            Dim IDTipoLinea As String = ProcessServer.ExecuteTask(Of Object, String)(AddressOf TipoLinea.TipoLineaPorDefecto, Nothing, services)
            For Each drVencimiento As DataRow In dtVencimiento.Rows
                Dim infoFraLin As FraLinObraPromo
                For j As Integer = 0 To CType(doc.Cabecera, FraCabObraPromo).Lineas.Length - 1
                    If CType(doc.Cabecera, FraCabObraPromo).Lineas(j).IDLocalVencimiento = drVencimiento("IDLocalVencimiento") Then
                        infoFraLin = CType(doc.Cabecera, FraCabObraPromo).Lineas(j)
                    End If
                Next
                If Not infoFraLin Is Nothing Then
                    Dim Obras As EntityInfoCache(Of ObraCabeceraInfo) = services.GetService(Of EntityInfoCache(Of ObraCabeceraInfo))()
                    Dim Obra As ObraCabeceraInfo = Obras.GetEntity(drVencimiento("IDObra"))

                    Dim drlinea As DataRow = New FacturaVentaLinea().AddNewForm.Rows(0)
                    drlinea("IDLineaFactura") = AdminData.GetAutoNumeric
                    drlinea("IDFactura") = doc.HeaderRow("IDFactura")
                    drlinea("IDLocalVencimiento") = drVencimiento("IDLocalVencimiento")
                    IDOrdenLinea += 1
                    drlinea("IDOrdenLinea") = IDOrdenLinea
                    drlinea("IDArticulo") = drVencimiento("IDArticulo")

                    Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                    Dim Articulo As ArticuloInfo = Articulos.GetEntity(drVencimiento("IDArticulo"))

                    Dim TextoLineaFactura As String = infoFraLin.Descripcion2
                    If Length(TextoLineaFactura) > 0 Then
                        Dim strTextoLinea As String = infoFraLin.Descripcion3 & " " & TextoLineaFactura & " en " & infoFraLin.DireccionObra
                        drlinea("DescArticulo") = strTextoLinea & " correspondiente al vencimiento " & drVencimiento("FechaVencimiento")
                    ElseIf Length(drVencimiento("DescVencimiento")) = 0 Then
                        drlinea("DescArticulo") = Articulo.DescArticulo
                    Else
                        drlinea("DescArticulo") = drVencimiento("DescVencimiento")
                    End If

                    If Length(CType(doc.Cabecera, FraCabObraPromo).CCAnticipo) = 0 Then
                        If ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf Cliente.Nacional, drVencimiento("IDCliente"), services) Then
                            drlinea("CContable") = Articulo.CCVenta
                        Else
                            drlinea("CContable") = Articulo.CCExport
                        End If
                    Else
                        drlinea("CContable") = CType(doc.Cabecera, FraCabObraPromo).CCAnticipo
                    End If
                    drlinea("IDTipoIva") = drVencimiento("IDTipoIva")
                    drlinea("IDCentroGestion") = Obra.IDCentroGestion
                    drlinea("Cantidad") = 1
                    drlinea("UDValoracion") = 1
                    drlinea("IDUDMedida") = Articulo.IDUDVenta
                    drlinea("IDUDInterna") = Articulo.IDUDInterna

                    Dim dataFactor As New ArticuloUnidadAB.DatosFactorConversion(drVencimiento("IDArticulo"), drlinea("IDUDMedida") & String.Empty, drlinea("IDUDInterna"))
                    drlinea("Factor") = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, dataFactor, services)
                    drlinea("QInterna") = drlinea("Factor") * drlinea("Cantidad")
                    drlinea("IDTipoLinea") = IDTipoLinea
                    drlinea("IDObra") = drVencimiento("IDObra")
                    drlinea("Dto") = doc.HeaderRow("DtoFactura")
                    drlinea("DtoProntoPago") = doc.HeaderRow("DtoProntoPago")
                    If Obra.CambioA > 0 Then
                        drlinea("Precio") = Nz(drVencimiento("ImpVencimientoA"), 0) / Obra.CambioA
                    End If
                    If IDMonedaA <> doc.HeaderRow("IDMoneda") Then
                        Dim BR As New BusinessRuleData("IDMoneda", IDMonedaA, New DataRowPropertyAccessor(drlinea), Nothing)
                        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoComunes.CambioMoneda, BR, services)
                    End If

                    Dim dr As IPropertyAccessor = New DataRowPropertyAccessor(drlinea)
                    Dim dataInfoMoneda As New ValoresAyB(dr, Obra.IDMoneda, Obra.CambioA, Obra.CambioB)
                    ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, dataInfoMoneda, services)

                    doc.dtLineas.Rows.Add(drlinea.ItemArray)
                End If
            Next
        End If
    End Sub

    <Task()> Public Shared Sub NuevaLineaFacturaObraPromoFinal(ByVal doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        Dim IDLocal(CType(doc.Cabecera, FraCabObraPromo).Lineas.Length - 1) As Object
        For i As Integer = 0 To CType(doc.Cabecera, FraCabObraPromo).Lineas.Length - 1
            IDLocal(i) = CType(doc.Cabecera, FraCabObraPromo).Lineas(i).IDLocal
        Next

        Dim f As New Filter
        f.Add(New InListFilterItem("IDLocal", IDLocal, FilterType.Numeric))
        Dim dtLocal As DataTable = BusinessHelper.CreateBusinessObject("ObraPromoLocal").Filter(f)
        If dtLocal.Rows.Count > 0 Then
            Dim IDTipoLinea As String = ProcessServer.ExecuteTask(Of Object, String)(AddressOf TipoLinea.TipoLineaPorDefecto, Nothing, services)
            For Each drLocal As DataRow In dtLocal.Rows
                Dim Obras As EntityInfoCache(Of ObraCabeceraInfo) = services.GetService(Of EntityInfoCache(Of ObraCabeceraInfo))()
                Dim Obra As ObraCabeceraInfo = Obras.GetEntity(drLocal("IDObra"))

                Dim infoFraLin As FraLinObraPromo
                For j As Integer = 0 To CType(doc.Cabecera, FraCabObraPromo).Lineas.Length - 1
                    If CType(doc.Cabecera, FraCabObraPromo).Lineas(j).IDLocal = drLocal("IDLocal") Then
                        infoFraLin = CType(doc.Cabecera, FraCabObraPromo).Lineas(j)
                    End If
                Next

                Dim k As Integer = 0
                Dim intTotalAnticipos As Integer = 0
                Dim dtAnticipo As DataTable = New BE.DataEngine().Filter("vNegAnticiposPorLocal", New NumberFilterItem("IDLocal", drLocal("IDLocal")))
                If dtAnticipo.Rows.Count > 0 Then
                    intTotalAnticipos = dtAnticipo.Rows.Count
                End If
                Dim blnTieneVivienda As Boolean

                For IDOrdenLinea As Integer = 1 To 2 + intTotalAnticipos
                    If Not infoFraLin Is Nothing Then
                        Dim drlinea As DataRow = New FacturaVentaLinea().AddNewForm.Rows(0)
                        drlinea("IDLineaFactura") = AdminData.GetAutoNumeric
                        drlinea("IDFactura") = doc.HeaderRow("IDFactura")
                        drlinea("IDOrdenLinea") = IDOrdenLinea
                        drlinea("IDArticulo") = infoFraLin.IDArticulo

                        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                        Dim Articulo As ArticuloInfo = Articulos.GetEntity(drlinea("IDArticulo"))

                        Dim Nacional As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf Cliente.Nacional, drLocal("IDCliente"), services)

                        Dim DescArticulo As String = String.Empty
                        Dim CContable As String = String.Empty
                        Dim PrecioA As Double = 0
                        Select Case IDOrdenLinea
                            Case 1 'Importe total de la Venta
                                DescArticulo = infoFraLin.Descripcion4 & " " & infoFraLin.Descripcion2 & " en " & infoFraLin.DireccionObra
                                CContable = IIf(Nacional, Articulo.CCVenta, Articulo.CCExport)
                                If drLocal("PrecioGarajeA") = 0 Then
                                    PrecioA = drLocal("ImpVentaA")
                                Else
                                    PrecioA = drLocal("ImpVentaA") - drLocal("PrecioGarajeA")
                                End If
                                drlinea("IDLocalVencimiento") = infoFraLin.IDLocalVencimiento
                                If PrecioA > 0 Then blnTieneVivienda = True
                            Case 2 'Importe Garaje
                                If drLocal("PrecioGarajeA") > 0 Then
                                    Dim TextoLineaFacturaGaraje As String = "Garaje nº " & infoFraLin.NumeroGaraje
                                    If Length(infoFraLin.Edificio) > 0 Then
                                        TextoLineaFacturaGaraje = TextoLineaFacturaGaraje & " del Edificio " & infoFraLin.Edificio
                                    End If
                                    If Length(infoFraLin.DireccionObra) > 0 Then
                                        TextoLineaFacturaGaraje = TextoLineaFacturaGaraje & " en " & infoFraLin.DireccionObra
                                    End If
                                    DescArticulo = infoFraLin.Descripcion4 & " " & TextoLineaFacturaGaraje

                                    CContable = IIf(Nacional, Articulo.CCVenta, Articulo.CCExport)
                                    PrecioA = drLocal("PrecioGarajeA")
                                    If Not blnTieneVivienda Then
                                        drlinea("IDLocalVencimiento") = infoFraLin.IDLocalVencimiento
                                    End If
                                End If
                            Case Else 'Anticipos
                                DescArticulo = "Importe entregado a cuenta"
                                PrecioA = Nz(dtAnticipo.Rows(k)("ImpAnticipoA"), 0) * -1
                                CContable = dtAnticipo.Rows(k)("CCAnticipo") & String.Empty
                                If Len(CContable) = 0 Then CContable = doc.HeaderRow("CCAnticipo")
                                k += 1
                        End Select

                        If PrecioA <> 0 Then
                            drlinea("DescArticulo") = DescArticulo
                            drlinea("CContable") = CContable
                            drlinea("IDTipoIva") = infoFraLin.IDTipoIva
                            drlinea("IDCentroGestion") = Obra.IDCentroGestion
                            If Obra.CambioA > 0 Then
                                drlinea("Precio") = PrecioA / Obra.CambioA
                            End If
                            drlinea("IDUDMedida") = Articulo.IDUDVenta
                            drlinea("IDUDInterna") = Articulo.IDUDInterna
                            drlinea("UDValoracion") = 1
                            drlinea("IDTipoLinea") = IDTipoLinea
                            drlinea("IDObra") = drLocal("IDObra")
                            drlinea("Cantidad") = 1
                            drlinea("Dto") = doc.HeaderRow("DtoFactura")
                            drlinea("DtoProntoPago") = doc.HeaderRow("DtoProntoPago")

                            Dim BR As New ArticuloUnidadAB.DatosFactorConversion(drlinea("IDArticulo"), drlinea("IDUDMedida") & String.Empty, drlinea("IDUDInterna"))
                            drlinea("Factor") = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, BR, services)
                            drlinea("QInterna") = drlinea("Factor") * drlinea("Cantidad")

                            doc.dtLineas.Rows.Add(drlinea.ItemArray)
                        End If
                    End If
                Next
            Next
        End If
    End Sub

#End Region

    '//Copiar Analítica de Obra
    <Task()> Public Shared Sub CopiarAnalitica(ByVal Doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If Doc.HeaderRow("Estado") <> enumfvcEstado.fvcContabilizado Then
            Dim AppParams As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
            If Not AppParams.Contabilidad OrElse Not AppParams.Analitica.AplicarAnalitica Then Exit Sub

            Dim f As New Filter(FilterUnionOperator.Or)
            f.Add(New IsNullFilterItem("IDTrabajo", False))
            f.Add(New IsNullFilterItem("IDObra", False))
            Dim WhereNotNullTrabajo As String = f.Compose(New AdoFilterComposer)
            For Each linea As DataRow In Doc.dtLineas.Select(WhereNotNullTrabajo)
                Dim data As New DataDocRow(Doc, linea)
                ProcessServer.ExecuteTask(Of DataDocRow)(AddressOf NegocioGeneral.NuevaAnalitica, data, services)
            Next
        End If
    End Sub

#End Region

#Region " Funciones Públicas "

    <Task()> Public Shared Function Resultado(ByVal data As Object, ByVal services As ServiceProvider) As ResultFacturacion
        'Elimina la información almacenada en memoria si previamente hemos cancelado la facturación
        AdminData.GetSessionData("__frax__")
        'Guardamos la información del documento en memoria, para recuperarla cuando volvamos del preview de presentación
        AdminData.SetSessionData("__frax__", services.GetService(Of ArrayList))
        Return services.GetService(Of ResultFacturacion)()
    End Function

    <Task()> Public Shared Sub Ordenar(ByVal data As FraCabObra(), ByVal services As ServiceProvider)
        If data IsNot Nothing Then Array.Sort(data, New OrdenFacturasObras)
    End Sub

#End Region

#Region " Actualización Obras "

    Public Class DataActualizarConceptosObras
        Public Doc As DocumentoFacturaVenta
        Public Deleting As Boolean

        Public Sub New(ByVal Doc As DocumentoFacturaVenta, Optional ByVal Deleting As Boolean = False)
            Me.Doc = Doc
            Me.Deleting = Deleting
        End Sub
    End Class

    Public Class DataActualizarRowConceptosObras
        Public Row As DataRow
        Public Deleting As Boolean

        Public Sub New(ByVal Row As DataRow, Optional ByVal Deleting As Boolean = False)
            Me.Row = Row
            Me.Deleting = Deleting
        End Sub
    End Class
    <Task()> Public Shared Sub ActualizarConceptosObras(ByVal data As DataActualizarConceptosObras, ByVal services As ServiceProvider)
        If Not data.Doc.dtLineas Is Nothing AndAlso data.Doc.dtLineas.Rows.Count > 0 Then
            Dim f As New Filter(FilterUnionOperator.Or)
            f.Add(New IsNullFilterItem("IDCertificacion", False))
            f.Add(New IsNullFilterItem("IDLineaMaterial", False))
            f.Add(New IsNullFilterItem("IDLineaMod", False))
            f.Add(New IsNullFilterItem("IDLineaCentro", False))
            f.Add(New IsNullFilterItem("IDLineaGasto", False))
            f.Add(New IsNullFilterItem("IDLineaVarios", False))
            f.Add(New IsNullFilterItem("IDTrabajo", False))
            f.Add(New IsNullFilterItem("IDObra", False))
            Dim WhereLineasObras As String = f.Compose(New AdoFilterComposer)
            Dim adrActObras() As DataRow = data.Doc.dtLineas.Select(WhereLineasObras, "IDObra")
            If Not adrActObras Is Nothing AndAlso adrActObras.Length > 0 Then
                For Each dr As DataRow In adrActObras
                    Dim datAct As New DataActualizarRowConceptosObras(dr, data.Deleting)
                    ProcessServer.ExecuteTask(Of DataActualizarRowConceptosObras)(AddressOf ActualizarConceptosObrasPorLinea, datAct, services)
                Next
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarObras(ByVal Doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If Not Doc Is Nothing AndAlso Not Doc.dtLineas Is Nothing AndAlso Doc.dtLineas.Rows.Count > 0 Then
            For Each linea As DataRow In Doc.dtLineas.Select
                Dim Modified As Boolean = False
                If linea.RowState = DataRowState.Modified AndAlso (Nz(linea("ImporteA")) <> Nz(linea("ImporteA", DataRowVersion.Original)) OrElse _
                       Nz(linea("IDObra")) <> Nz(linea("IDObra", DataRowVersion.Original)) OrElse Nz(linea("IDTrabajo")) <> Nz(linea("IDTrabajo", DataRowVersion.Original))) Then
                    If (Nz(linea("IDTrabajo", DataRowVersion.Original), 0) > 0 AndAlso Length(linea("IDTrabajo")) = 0) OrElse _
                        (Length(linea("IDTrabajo", DataRowVersion.Original)) > 0 AndAlso Length(linea("IDTrabajo")) > 0 AndAlso linea("IDTrabajo", DataRowVersion.Original) <> linea("IDTrabajo")) Then
                        Dim OT As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraTrabajo")
                        Dim dtOT As DataTable = OT.SelOnPrimaryKey(linea("IDTrabajo", DataRowVersion.Original))
                        If Not IsNothing(dtOT) AndAlso dtOT.Rows.Count > 0 Then
                            dtOT.Rows(0)("ImpFactTrabajoA") = Nz(dtOT.Rows(0)("ImpFactTrabajoA"), 0) - linea("ImporteA")
                            OT.Update(dtOT)
                        End If
                    End If
                    If (Nz(linea("IDObra", DataRowVersion.Original), 0) > 0 AndAlso Length(linea("IDObra")) = 0) OrElse _
                        (Length(linea("IDObra", DataRowVersion.Original)) > 0 AndAlso Length(linea("IDObra")) > 0 AndAlso linea("IDObra", DataRowVersion.Original) <> linea("IDObra")) Then
                        Dim OC As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraCabecera")
                        Dim dtOC As DataTable = OC.SelOnPrimaryKey(linea("IDObra", DataRowVersion.Original))
                        If Not IsNothing(dtOC) AndAlso dtOC.Rows.Count > 0 Then
                            dtOC.Rows(0)("ImpFactA") = Nz(dtOC.Rows(0)("ImpFactA"), 0) - linea("ImporteA")
                            OC.Update(dtOC)
                        End If
                    End If
                    Modified = True
                End If

                If linea.RowState = DataRowState.Added OrElse Modified Then
                    Dim datActObras As New DataActualizarRowConceptosObras(linea)
                    ProcessServer.ExecuteTask(Of DataActualizarRowConceptosObras)(AddressOf ProcesoFacturacionObras.ActualizarConceptosObrasPorLinea, datActObras, services)
                End If
            Next
        End If
    End Sub

    <Serializable()> _
    Public Class DatosActuaObraTrabFact
        Public Dt As DataTable
        Public NFactura As String
    End Class

    <Task()> Public Shared Sub ActualizarObraTrabajoFacturacion(ByVal doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If Not doc.dtLineas Is Nothing AndAlso doc.dtLineas.Rows.Count > 0 Then
            Dim strIN As String
            Dim dtObraTrabajoFacturacion As DataTable
            Dim dtOTF(-1) As DataTable
            Dim fVto As New Filter
            fVto.Add(New IsNullFilterItem("IDVencimiento", False))
            Dim WhereNotNullVencimiento As String = fVto.Compose(New AdoFilterComposer)
            Dim adrLineasConVto() As DataRow = doc.dtLineas.Select(WhereNotNullVencimiento)
            If Not adrLineasConVto Is Nothing AndAlso adrLineasConVto.Length > 0 Then
                For Each dr As DataRow In adrLineasConVto
                    If Length(dr("IDVencimiento")) > 0 Then
                        If InStr(strIN, CStr(dr("IDVencimiento")), CompareMethod.Text) = 0 Then
                            If Len(strIN) > 0 Then strIN = strIN & ","
                            strIN = strIN & dr("IDVencimiento")

                            Dim Obra As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraTrabajoFacturacion")
                            Dim DtHito As DataTable = Obra.Filter(New FilterItem("IDVencimiento", FilterOperator.Equal, dr("IDVencimiento")), "IDTrabajo, IDVencimiento")
                            If Not DtHito Is Nothing AndAlso DtHito.Rows.Count > 0 Then
                                For Each Facturacion As DataRow In DtHito.Select
                                    Facturacion("Facturado") = True
                                    Facturacion("IDFactura") = dr("IDFactura")
                                    Facturacion("NFactura") = doc.HeaderRow("NFactura")
                                Next
                            End If

                            ReDim Preserve dtOTF(UBound(dtOTF) + 1) : dtOTF(UBound(dtOTF)) = DtHito
                            If Not DtHito Is Nothing Then DtHito.Dispose()
                        End If
                    End If
                Next
                Dim pa As New UpdatePackage
                For Each dt As DataTable In dtOTF
                    pa.Add(dt)
                Next
                BusinessHelper.UpdatePackage(pa)
            End If

        End If
    End Sub

#Region " Promotoras "

    <Task()> Public Shared Sub ActualizaObraPromoLocalVencimiento(ByVal doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If Not doc.dtLineas Is Nothing AndAlso doc.dtLineas.Rows.Count > 0 Then
            Dim Where As String = New IsNullFilterItem("IDLocalVencimiento", False).Compose(New AdoFilterComposer)
            Dim IDLocalVencimientoTratado As New Hashtable
            For Each drLinea As DataRow In doc.dtLineas.Select(Where)
                If Not IDLocalVencimientoTratado.ContainsKey(drLinea("IDLocalVencimiento")) Then
                    IDLocalVencimientoTratado(drLinea("IDLocalVencimiento")) = drLinea("IDLocalVencimiento")

                    Dim f As New Filter
                    f.Add(New NumberFilterItem("IDLocalVencimiento", drLinea("IDLocalVencimiento")))
                    Dim dtVtos As DataTable = BusinessHelper.CreateBusinessObject("ObraPromoLocalVencimiento").Filter(f)
                    If Not dtVtos Is Nothing AndAlso dtVtos.Rows.Count > 0 Then
                        For Each drVtos As DataRow In dtVtos.Rows
                            drVtos("Facturado") = True
                            drVtos("IDFactura") = drLinea("IDFactura")
                            drVtos("NFactura") = doc.HeaderRow("NFactura")
                            If drVtos("CobroGenerado") AndAlso Nz(drVtos("IDCobro"), 0) > 0 Then
                                Dim dataCobro As New dataActualizarCobro(drVtos("IDCobro"), drLinea("IDFactura"), doc.HeaderRow("NFactura"))
                                ProcessServer.ExecuteTask(Of dataActualizarCobro)(AddressOf ActualizarCobro, dataCobro, services)
                            End If
                        Next
                        BusinessHelper.UpdateTable(dtVtos)
                    End If
                End If
            Next
        End If
    End Sub

    <Task()> Public Shared Sub ActualizaObraPromoLocalVencimientoFinal(ByVal doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If Not doc.dtLineas Is Nothing AndAlso doc.dtLineas.Rows.Count > 0 Then
            Dim Where As String = New IsNullFilterItem("IDLocalVencimiento", False).Compose(New AdoFilterComposer)
            Dim IDLocalVencimientoTratado As New Hashtable
            For Each drLinea As DataRow In doc.dtLineas.Select(Where)
                If Not IDLocalVencimientoTratado.ContainsKey(drLinea("IDLocalVencimiento")) Then
                    IDLocalVencimientoTratado(drLinea("IDLocalVencimiento")) = drLinea("IDLocalVencimiento")

                    Dim PLV As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraPromoLocalVencimiento")
                    Dim f As New Filter
                    f.Add(New NumberFilterItem("IDLocalVencimiento", drLinea("IDLocalVencimiento")))
                    Dim dtVtos As DataTable = PLV.Filter(f)
                    If Not dtVtos Is Nothing AndAlso dtVtos.Rows.Count > 0 Then
                        f.Clear()
                        f.Add(New NumberFilterItem("IDLocal", dtVtos.Rows(0)("IDLocal")))
                        f.Add(New BooleanFilterItem("Facturado", False))
                        dtVtos = PLV.Filter(f)
                        If Not dtVtos Is Nothing AndAlso dtVtos.Rows.Count > 0 Then
                            For Each drVtos As DataRow In dtVtos.Rows
                                drVtos("Facturado") = True
                                drVtos("IDFactura") = drLinea("IDFactura")
                                drVtos("NFactura") = doc.HeaderRow("NFactura")
                                drVtos("TipoFactura") = enumoptvTipoFactura.optvFinal
                                If drVtos("CobroGenerado") AndAlso Nz(drVtos("IDCobro"), 0) > 0 Then
                                    Dim dataCobro As New dataActualizarCobro(drVtos("IDCobro"), drLinea("IDFactura"), doc.HeaderRow("NFactura"))
                                    ProcessServer.ExecuteTask(Of dataActualizarCobro)(AddressOf ActualizarCobro, dataCobro, services)
                                End If
                            Next
                            BusinessHelper.UpdateTable(dtVtos)
                        End If
                    End If
                End If
            Next
        End If
    End Sub

#End Region

    <Serializable()> _
    Public Class dataActualizarCobro
        Public IDCobro As Integer
        Public IDFactura As Integer
        Public NFactura As String

        Public Sub New(ByVal IDCobro As Integer, ByVal IDFactura As Integer, ByVal NFactura As String)
            Me.IDCobro = IDCobro
            Me.IDFactura = IDFactura
            Me.NFactura = NFactura
        End Sub
    End Class
    <Task()> Public Shared Sub ActualizarCobro(ByVal data As dataActualizarCobro, ByVal services As ServiceProvider)
        If data.IDCobro > 0 Then
            Dim dt As DataTable = New Cobro().SelOnPrimaryKey(data.IDCobro)
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                dt.Rows(0)("IDFactura") = data.IDFactura
                dt.Rows(0)("NFactura") = data.NFactura
                BusinessHelper.UpdateTable(dt)
            End If
        End If
    End Sub

#Region " ActualizarConceptosObrasPorLinea "

    <Task()> Public Shared Sub ActualizarConceptosObrasPorLinea(ByVal data As DataActualizarRowConceptosObras, ByVal services As ServiceProvider)
        If Length(data.Row("IDCertificacion")) > 0 Then
            ProcessServer.ExecuteTask(Of DataActualizarRowConceptosObras)(AddressOf ActualizarObraCertificaciones, data, services)
        ElseIf Length(data.Row("IDLineaMaterial")) > 0 Then
            ProcessServer.ExecuteTask(Of DataActualizarRowConceptosObras)(AddressOf ActualizarObraMateriales, data, services)
        ElseIf Length(data.Row("IDLineaMod")) > 0 Then
            ProcessServer.ExecuteTask(Of DataActualizarRowConceptosObras)(AddressOf ActualizarObraMod, data, services)
        ElseIf Length(data.Row("IDLineaCentro")) > 0 Then
            ProcessServer.ExecuteTask(Of DataActualizarRowConceptosObras)(AddressOf ActualizarObraCentro, data, services)
        ElseIf Length(data.Row("IDLineaGasto")) > 0 Then
            ProcessServer.ExecuteTask(Of DataActualizarRowConceptosObras)(AddressOf ActualizarObraGasto, data, services)
        ElseIf Length(data.Row("IDLineaVarios")) > 0 Then
            ProcessServer.ExecuteTask(Of DataActualizarRowConceptosObras)(AddressOf ActualizarObraVarios, data, services)
        End If
        If Length(data.Row("IDTrabajo")) > 0 Then ProcessServer.ExecuteTask(Of DataActualizarRowConceptosObras)(AddressOf ActualizarObraTrabajo, data, services)
        If Length(data.Row("IDObra")) > 0 Then ProcessServer.ExecuteTask(Of DataActualizarRowConceptosObras)(AddressOf ActualizarObraCabecera, data, services)
    End Sub

    <Task()> Public Shared Sub ActualizarObraCertificaciones(ByVal data As DataActualizarRowConceptosObras, ByVal services As ServiceProvider)
        If Length(data.Row("IDCertificacion")) > 0 Then
            Dim OC As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraCertificacion")
            Dim drOC As DataRow = OC.GetItemRow(data.Row("IDCertificacion"))
            If data.Deleting Then
                drOC("Estado") = enumEstadoCertificacion.ecAceptada
            Else
                drOC("Estado") = enumEstadoCertificacion.ecFacturado
            End If
            BusinessHelper.UpdateTable(drOC.Table)
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarObraMateriales(ByVal data As DataActualizarRowConceptosObras, ByVal services As ServiceProvider)
        If Length(data.Row("IDLineaMaterial")) > 0 Then
            Dim OM As BusinessHelper = BusinessHelper.CreateBusinessObject(AdminData.GetEntityInfo("ObraMaterial"))
            Dim dtOM As DataTable = OM.Filter(New NumberFilterItem("IDLineaMaterial", data.Row("IDLineaMaterial")))
            If Not IsNothing(dtOM) AndAlso dtOM.Rows.Count > 0 Then
                Dim dblImporteAOLD As Double
                If data.Row.RowState = DataRowState.Modified Then
                    If data.Row("ImporteA") <> Nz(data.Row("ImporteA", DataRowVersion.Original), 0) Then
                        dblImporteAOLD = Nz((data.Row("ImporteA", DataRowVersion.Original)), 0)
                    End If
                End If
                Dim intFactor As Integer = IIf(data.Deleting, -1, 1)
                dtOM.Rows(0)("ImpFactMatA") = Nz(dtOM.Rows(0)("ImpFactMatA"), 0) + (data.Row("ImporteA") * intFactor) - dblImporteAOLD
                OM.Update(New UpdatePackage(dtOM), services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarObraMod(ByVal data As DataActualizarRowConceptosObras, ByVal services As ServiceProvider)
        If Length(data.Row("IDLineaMod")) > 0 Then
            Dim OM As BusinessHelper = BusinessHelper.CreateBusinessObject(AdminData.GetEntityInfo("ObraMod"))
            Dim dtOM As DataTable = OM.Filter(New NumberFilterItem("IDLineaMod", data.Row("IDLineaMod")))
            If Not IsNothing(dtOM) AndAlso dtOM.Rows.Count > 0 Then
                Dim dblImporteAOLD As Double
                If data.Row.RowState = DataRowState.Modified Then
                    If data.Row("ImporteA") <> Nz(data.Row("ImporteA", DataRowVersion.Original), 0) Then
                        dblImporteAOLD = Nz(data.Row("ImporteA", DataRowVersion.Original), 0)
                    End If
                End If

                Dim intFactor As Integer = IIf(data.Deleting, -1, 1)
                dtOM.Rows(0)("ImpFactModA") = Nz(dtOM.Rows(0)("ImpFactModA"), 0) + (data.Row("ImporteA") * intFactor) - dblImporteAOLD
                OM.Update(New UpdatePackage(dtOM), services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarObraCentro(ByVal data As DataActualizarRowConceptosObras, ByVal services As ServiceProvider)
        If Length(data.Row("IDLineaCentro")) > 0 Then
            Dim OC As BusinessHelper = BusinessHelper.CreateBusinessObject(AdminData.GetEntityInfo("ObraCentro"))
            Dim dtOC As DataTable = OC.Filter(New NumberFilterItem("IDLineaCentro", data.Row("IDLineaCentro")))
            If Not IsNothing(dtOC) AndAlso dtOC.Rows.Count > 0 Then
                Dim dblImporteAOLD As Double
                If data.Row.RowState = DataRowState.Modified Then
                    If data.Row("ImporteA") <> Nz(data.Row("ImporteA", DataRowVersion.Original), 0) Then
                        dblImporteAOLD = Nz((data.Row("ImporteA", DataRowVersion.Original)), 0)
                    End If
                End If
                Dim intFactor As Integer = IIf(data.Deleting, -1, 1)
                dtOC.Rows(0)("ImpFactCentroA") = Nz(dtOC.Rows(0)("ImpFactCentroA"), 0) + (data.Row("ImporteA") * intFactor) - dblImporteAOLD
                OC.Update(New UpdatePackage(dtOC), services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarObraGasto(ByVal data As DataActualizarRowConceptosObras, ByVal services As ServiceProvider)
        If Length(data.Row("IDLineaGasto")) > 0 Then
            Dim OG As BusinessHelper = BusinessHelper.CreateBusinessObject(AdminData.GetEntityInfo("ObraGasto"))
            Dim dtOG As DataTable = OG.Filter(New NumberFilterItem("IDLineaGasto", data.Row("IDLineaGasto")))
            If Not IsNothing(dtOG) AndAlso dtOG.Rows.Count > 0 Then
                Dim dblImporteAOLD As Double
                If data.Row.RowState = DataRowState.Modified Then
                    If data.Row("ImporteA") <> Nz(data.Row("ImporteA", DataRowVersion.Original), 0) Then
                        dblImporteAOLD = Nz((data.Row("ImporteA", DataRowVersion.Original)), 0)
                    End If
                End If
                Dim intFactor As Integer = IIf(data.Deleting, -1, 1)
                dtOG.Rows(0)("ImpFactGastoA") = Nz(dtOG.Rows(0)("ImpFactGastoA"), 0) + (data.Row("ImporteA") * intFactor) - dblImporteAOLD
                OG.Update(New UpdatePackage(dtOG), services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarObraVarios(ByVal data As DataActualizarRowConceptosObras, ByVal services As ServiceProvider)
        If Length(data.Row("IDLineaVarios")) > 0 Then
            Dim OV As BusinessHelper = BusinessHelper.CreateBusinessObject(AdminData.GetEntityInfo("ObraVarios"))
            Dim dtOV As DataTable = OV.Filter(New NumberFilterItem("IDLineaVarios", data.Row("IDLineaVarios")))
            If Not IsNothing(dtOV) AndAlso dtOV.Rows.Count > 0 Then
                Dim dblImporteAOLD As Double
                If data.Row.RowState = DataRowState.Modified Then
                    If data.Row("ImporteA") <> Nz(data.Row("ImporteA", DataRowVersion.Original), 0) Then
                        dblImporteAOLD = Nz((data.Row("ImporteA", DataRowVersion.Original)), 0)
                    End If
                End If
                Dim intFactor As Integer = IIf(data.Deleting, -1, 1)
                dtOV.Rows(0)("ImpFactVariosA") = Nz(dtOV.Rows(0)("ImpFactVariosA"), 0) + (data.Row("ImporteA") * intFactor) - dblImporteAOLD
                OV.Update(New UpdatePackage(dtOV), services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarObraTrabajo(ByVal data As DataActualizarRowConceptosObras, ByVal services As ServiceProvider)
        If Length(data.Row("IDTrabajo")) > 0 Then
            Dim OT As BusinessHelper = BusinessHelper.CreateBusinessObject(AdminData.GetEntityInfo("ObraTrabajo"))
            Dim dtOT As DataTable = OT.Filter(New NumberFilterItem("IDTrabajo", data.Row("IDTrabajo")))
            If Not IsNothing(dtOT) AndAlso dtOT.Rows.Count > 0 Then
                Dim dblImporteAOLD, dblCantidadOLD As Double
                If data.Row.RowState = DataRowState.Modified Then
                    If data.Row("ImporteA") <> Nz(data.Row("ImporteA", DataRowVersion.Original), 0) Or data.Row("Cantidad") <> Nz(data.Row("Cantidad", DataRowVersion.Original), 0) Then
                        dblImporteAOLD = Nz((data.Row("ImporteA", DataRowVersion.Original)), 0)
                        dblCantidadOLD = Nz((data.Row("Cantidad", DataRowVersion.Original)), 0)
                    End If
                ElseIf data.Deleting Then
                    dblImporteAOLD = data.Row("ImporteA")
                    dblCantidadOLD = data.Row("Cantidad")
                End If
                Dim intFactor As Integer = IIf(data.Deleting, 0, 1)
                If dtOT.Rows(0)("TipoFacturacion") = enumotTipoFacturacion.otfPorVencimientos Then
                    dtOT.Rows(0)("ImpFactTrabajoA") = Nz(dtOT.Rows(0)("ImpFactTrabajoA"), 0) + (data.Row("ImporteA") * intFactor) - dblImporteAOLD
                    OT.Update(dtOT)
                ElseIf dtOT.Rows(0)("TipoFacturacion") = enumotTipoFacturacion.otfPorUdsObra Then
                    '//Certificaciones
                    dtOT.Rows(0)("QFact") = dtOT.Rows(0)("QFact") + (data.Row("Cantidad") * intFactor) - dblCantidadOLD
                    dtOT.Rows(0)("ImpFactTrabajoA") = dtOT.Rows(0)("ImpFactTrabajoA") + (data.Row("ImporteA") * intFactor) - dblImporteAOLD
                    OT.Update(dtOT)
                ElseIf dtOT.Rows(0)("TipoFacturacion") = enumotTipoFacturacion.otfPorConceptos Then
                    If Length(data.Row("IDLineaMaterial")) > 0 Then
                        dtOT.Rows(0)("ImpFactMatA") = Nz(dtOT.Rows(0)("ImpFactMatA"), 0) + (data.Row("ImporteA") * intFactor) - dblImporteAOLD
                    ElseIf Length(data.Row("IDLineaMOD")) > 0 Then
                        dtOT.Rows(0)("ImpFactModA") = Nz(dtOT.Rows(0)("ImpFactModA"), 0) + (data.Row("ImporteA") * intFactor) - dblImporteAOLD
                    ElseIf Length(data.Row("IDLineaCentro")) > 0 Then
                        dtOT.Rows(0)("ImpFactCentrosA") = Nz(dtOT.Rows(0)("ImpFactCentrosA"), 0) + (data.Row("ImporteA") * intFactor) - dblImporteAOLD
                    ElseIf Length(data.Row("IDLineaGasto")) > 0 Then
                        dtOT.Rows(0)("ImpFactGastosA") = Nz(dtOT.Rows(0)("ImpFactGastosA"), 0) + (data.Row("ImporteA") * intFactor) - dblImporteAOLD
                    ElseIf Length(data.Row("IDLineaVarios")) > 0 Then
                        dtOT.Rows(0)("ImpFactVariosA") = Nz(dtOT.Rows(0)("ImpFactVariosA"), 0) + (data.Row("ImporteA") * intFactor) - dblImporteAOLD
                        'Else
                        'dtOT.Rows(0)("ImpFactTrabajoA") = dtOT.Rows(0)("ImpFactTrabajoA") + (data.Row("ImporteA") * intFactor) - dblImporteAOLD
                    End If
                    dtOT.Rows(0)("ImpFactTrabajoA") = dtOT.Rows(0)("ImpFactTrabajoA") + (data.Row("ImporteA") * intFactor) - dblImporteAOLD

                    OT.Update(New UpdatePackage(dtOT), services)
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarObraCabecera(ByVal data As DataActualizarRowConceptosObras, ByVal services As ServiceProvider)
        If Length(data.Row("IDObra")) > 0 Then
            Dim OC As BusinessHelper = BusinessHelper.CreateBusinessObject(AdminData.GetEntityInfo("ObraCabecera"))
            Dim dtOC As DataTable = OC.Filter(New NumberFilterItem("IDObra", data.Row("IDObra")))
            If Not IsNothing(dtOC) AndAlso dtOC.Rows.Count > 0 Then
                Dim dblImporteAOLD As Double
                If data.Row.RowState = DataRowState.Modified Or data.Deleting Then
                    If data.Row.RowState = DataRowState.Modified Then
                        If data.Row("IDObra") = Nz(data.Row("IDObra", DataRowVersion.Original), 0) Or data.Row("ImporteA") <> Nz(data.Row("ImporteA", DataRowVersion.Original), 0) Or data.Row("Cantidad") <> Nz(data.Row("Cantidad", DataRowVersion.Original), 0) Then
                            dblImporteAOLD = Nz((data.Row("ImporteA", DataRowVersion.Original)), 0)
                        End If
                    ElseIf data.Deleting Then
                        dblImporteAOLD = data.Row("ImporteA")
                    End If
                End If
                Dim intFactor As Integer = IIf(data.Deleting, 0, 1)
                dtOC.Rows(0)("ImpFactA") = Nz(dtOC.Rows(0)("ImpFactA"), 0) + (Nz(data.Row("ImporteA"), 0) * intFactor) - dblImporteAOLD
                dtOC.Rows(0)("ImpQFactA") = dtOC.Rows(0)("ImpFactA")

                OC.Update(New UpdatePackage(dtOC), services)
            End If
        End If
    End Sub

#End Region

#End Region

    <Task()> Public Shared Sub RecalcularDireccion(ByVal docfactura As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        Dim f As New Filter
        f.Add("IDCliente", FilterOperator.Equal, docfactura.HeaderRow("IdCliente"))
        f.Add("IDObra", FilterOperator.Equal, docfactura.HeaderRow("IdObra"))
        f.Add("DireccionFactura", FilterOperator.Equal, 1)
        Dim dtClienteDir As DataTable = New ClienteDireccion().Filter(f)
        If Not IsNothing(dtClienteDir) AndAlso dtClienteDir.Rows.Count > 0 Then
            docfactura.HeaderRow("IDDireccion") = dtClienteDir.Rows(0)("IDDireccion")
        Else
            Dim StDatos As New ClienteDireccion.DataDirecEnvio(docfactura.HeaderRow("IDCliente"), enumcdTipoDireccion.cdDireccionFactura)
            Dim dtClDir As DataTable = ProcessServer.ExecuteTask(Of ClienteDireccion.DataDirecEnvio, DataTable)(AddressOf ClienteDireccion.ObtenerDireccionEnvio, StDatos, services)
            If Not IsNothing(dtClDir) AndAlso dtClDir.Rows.Count > 0 Then docfactura.HeaderRow("IDDireccion") = dtClDir.Rows(0)("IDDireccion")
        End If
    End Sub

End Class