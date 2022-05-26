Public Class ProcesoFacturacionVenta

    <Task()> Public Shared Function CrearDocumento(ByVal data As UpdatePackage, ByVal services As ServiceProvider) As DocumentoFacturaVenta
        Return New DocumentoFacturaVenta(data)
    End Function

#Region "Agrupaciones facturacion"

#Region "Agrupaciones facturacion  -  Origen OTs "

    Public Class DataGetGroupColumnsOT
        Public Table As DataTable
        Public Agrupacion As enummcAgrupOT

        Public Sub New(ByVal Table As DataTable, ByVal Agrupacion As enummcAgrupOT)
            Me.Table = Table
            Me.Agrupacion = Agrupacion
        End Sub
    End Class
    <Task()> Public Shared Function GetGroupColumnsOT(ByVal data As DataGetGroupColumnsOT, ByVal services As ServiceProvider) As DataColumn()
        Dim columns(4) As DataColumn
        columns(0) = data.Table.Columns("IDCliente")
        columns(1) = data.Table.Columns("IDFormaPago")
        columns(2) = data.Table.Columns("IDCondicionPago")
        columns(3) = data.Table.Columns("IDDiaPago")
        columns(4) = data.Table.Columns("IDMoneda")
        If data.Agrupacion = enummcAgrupOT.OT Then
            ReDim Preserve columns(5)
            columns(5) = data.Table.Columns("IDOT")
        End If
        Return columns
    End Function

    <Task()> Public Shared Function AgruparOTs(ByVal data As DataPrcFacturacionOTs, ByVal services As ServiceProvider) As FraCabMnto()
        Dim IDMntoOTControlCopy(data.IDMntoOTControl.Length - 1) As Object
        data.IDMntoOTControl.CopyTo(IDMntoOTControlCopy, 0)
        Dim dtControlOT As DataTable = New BE.DataEngine().Filter("CIFacturacionVentaOT", New InListFilterItem("IDMntoOTControl", IDMntoOTControlCopy, FilterType.Numeric))
        If dtControlOT.Rows.Count > 0 Then
            Dim oGrprUser As New GroupUserFVMantenimiento
            Dim grpOT As DataColumn() = ProcessServer.ExecuteTask(Of DataGetGroupColumnsOT, DataColumn())(AddressOf GetGroupColumnsOT, New DataGetGroupColumnsOT(dtControlOT, enummcAgrupOT.OT), services)
            Dim grpClte As DataColumn() = ProcessServer.ExecuteTask(Of DataGetGroupColumnsOT, DataColumn())(AddressOf GetGroupColumnsOT, New DataGetGroupColumnsOT(dtControlOT, enummcAgrupOT.Cliente), services)

            Dim groupers(1) As GroupHelper
            groupers(enummcAgrupOT.OT) = New GroupHelper(grpOT, oGrprUser)
            groupers(enummcAgrupOT.Cliente) = New GroupHelper(grpClte, oGrprUser)

            For Each rwLin As DataRow In dtControlOT.Rows
                groupers(rwLin("AgrupOT")).Group(rwLin)
            Next

            If Not data.FechaFactura Is Nothing Then
                For Each fra As FraCabMnto In oGrprUser.Fras
                    fra.Fecha = data.FechaFactura
                Next
            End If
            Return oGrprUser.Fras
        Else
            ApplicationService.GenerateError("No hay datos a Facturar. Revise sus OTs.")
        End If
    End Function

    <Task()> Public Shared Function GetDatosFacturacionOT(ByVal f As Filter, ByVal services As ServiceProvider) As DataTable
        Dim AppParams As Parametro = services.GetService(Of Parametro)()

        Dim dtControl As DataTable = New BE.DataEngine().Filter("CIFacturacionVentaOT", f)
        If Not dtControl Is Nothing AndAlso dtControl.Rows.Count > 0 Then
            For Each control As DataRow In dtControl.Rows
                Select Case control("Tipo")
                    Case enumOTTipoLineasControl.OTMaterial
                        control("PrecioA") = Nz(control("PrecioMatA"), 0)
                    Case enumOTTipoLineasControl.OTMod
                        Dim Precio As Double = 0
                        control("IDArticulo") = AppParams.ArticuloFacturacionMOD
                        control("DescArticulo") = Format(Nz(control("TiempoOperario"), 0), "###0.##") & " horas de Mano de Obra "
                        If Nz(control("TiempoOperario"), 0) <> 0 Then Precio = Nz(control("TiempoOperario"), 0) * Nz(control("TasaModA"), 0)
                        If Nz(control("TiempoParada"), 0) <> 0 Then
                            Precio += (Nz(control("TiempoParada"), 0) * Nz(control("CosteParadaA"), 0))
                            control("DescArticulo") &= "y " & Format(Nz(control("TiempoParada"), 0), "###0.##") & " horas de Tiempo de Parada"
                        End If
                        control("PrecioA") = Precio
                        control("PrecioVentaA") = Nz(control("ImporteVentaA"), 0)
                    Case enumOTTipoLineasControl.OTContrata
                        Dim ImporteA As Double   '//ImporteA en el control sólo tiene la Mano de Obra
                        Dim Precio As Double = Nz(control("PrecioMatA"), 0)
                        ImporteA += Nz(control("PrecioMatA"), 0)
                        control("IDArticulo") = AppParams.ArticuloFacturacionContrata
                        control("DescArticulo") = "Contratas:"
                        control("DescArticulo") = Format(Nz(control("TiempoOperario"), 0), "###0.##") & " horas de Mano de Obra "
                        If Nz(control("TiempoOperario"), 0) <> 0 Then
                            Precio += Nz(control("TiempoOperario"), 0) * Nz(control("TasaModA"), 0)
                            ImporteA += Nz(control("TiempoOperario"), 0) * Nz(control("TasaModA"), 0)
                        End If
                        If Nz(control("TiempoParada"), 0) <> 0 Then
                            Precio += (Nz(control("TiempoParada"), 0) * Nz(control("CosteParadaA"), 0))
                            ImporteA += (Nz(control("TiempoParada"), 0) * Nz(control("CosteParadaA"), 0))
                            control("DescArticulo") &= "y " & Format(Nz(control("TiempoParada"), 0), "###0.##") & " horas de Tiempo de Parada"
                        End If
                        Precio += Nz(control("TasaOtrosA"), 0)
                        ImporteA += Nz(control("TasaOtrosA"), 0)
                        If Nz(control("TasaOtrosA"), 0) <> 0 Then control("DescArticulo") &= " + Gastos"

                        control("PrecioA") = Precio
                        control("ImporteA") = ImporteA
                    Case enumOTTipoLineasControl.OTGastos
                        control("IDArticulo") = AppParams.ArticuloFacturacionGastos
                        control("DescArticulo") = "Gastos"

                        control("PrecioA") = Nz(control("TasaOtrosA"), 0)
                End Select

            Next

            dtControl.AcceptChanges()
            Return dtControl
        End If
    End Function

    Public Class DataActualizarRowControlOT
        Public Row As DataRow
        Public Deleting As Boolean

        Public Sub New(ByVal Row As DataRow, Optional ByVal Deleting As Boolean = False)
            Me.Row = Row
            Me.Deleting = Deleting
        End Sub
    End Class

    <Task()> Public Shared Function ActualizarControlOT(ByVal data As DataActualizarRowControlOT, ByVal services As ServiceProvider) As DataTable
        If Length(data.Row("IDMntoOTControl")) > 0 Then
            Dim OTControl As BusinessHelper = BusinessHelper.CreateBusinessObject("MntoOTControlLinea")
            Dim dtOTControl As DataTable = OTControl.SelOnPrimaryKey(data.Row("IDMntoOTControl"))
            If data.Deleting Then
                dtOTControl.Rows(0)("Facturado") = False
                '      dtOTControl.Rows(0)("IDLineaFactura") = System.DBNull.Value
            Else
                dtOTControl.Rows(0)("Facturado") = True
                '      dtOTControl.Rows(0)("IDLineaFactura") = data.Row("IDLineaFactura")
            End If
            BusinessHelper.UpdateTable(dtOTControl)
        End If
    End Function

    <Task()> Public Shared Sub ActualizarOTs(ByVal Doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If Doc.HeaderRow("Estado") = enumfvcEstado.fvcNoContabilizado Then
            If Not Doc Is Nothing AndAlso Not Doc.dtLineas Is Nothing AndAlso Doc.dtLineas.Rows.Count > 0 Then
                Dim f As New Filter(FilterUnionOperator.Or)
                f.Add(New IsNullFilterItem("IDMntoOTControl", False))
                Dim Where As String = f.Compose(New AdoFilterComposer)
                For Each linea As DataRow In Doc.dtLineas.Select(Where)
                    If linea.RowState = DataRowState.Added OrElse linea.RowState = DataRowState.Modified Then
                        Dim datActOTs As New DataActualizarRowControlOT(linea)
                        ProcessServer.ExecuteTask(Of DataActualizarRowControlOT)(AddressOf ProcesoFacturacionVenta.ActualizarControlOT, datActOTs, services)
                    End If
                Next
            End If
        End If
    End Sub

#End Region

    <Task()> Public Shared Function AgruparAlbaranes(ByVal data As DataPrcFacturacionGeneral, ByVal services As ServiceProvider) As FraCabAlbaran()

        Dim dtLineas As DataTable

        'se seleccionan todas las lineas de albaran no facturadas

        Dim strViewName As String = "vNegComercialCrearFactura"

        If data.IDAlbaranes.Length > 0 Then
            Dim values(data.IDAlbaranes.Length - 1) As Object
            data.IDAlbaranes.CopyTo(values, 0)
            Dim oFltr As New Filter
            oFltr.Add(New InListFilterItem("IDAlbaran", values, FilterType.Numeric))
            oFltr.Add(New NumberFilterItem("EstadoFactura", FilterOperator.NotEqual, enumavlEstadoFactura.avlFacturado))
            dtLineas = New BE.DataEngine().Filter(strViewName, oFltr)
        End If

        If Not dtLineas Is Nothing AndAlso dtLineas.Rows.Count > 0 Then
            Dim p As New Parametro
            Dim fvcFecha As enumfvcFechaAlbaran = p.FacturacionFechaAlbaran()
            Dim oGrprUser As New GroupUserAlbaranes(data.DteFechaFactura, fvcFecha)
            Dim strCondicionPago As String = p.CondicionPago

            Dim grpAlb As DataColumn() = ProcessServer.ExecuteTask(Of DataGetGroupColumns, DataColumn())(AddressOf GetGroupColumns, New DataGetGroupColumns(dtLineas, enummcAgrupFactura.mcAlbaran), services)
            Dim grpClte As DataColumn() = ProcessServer.ExecuteTask(Of DataGetGroupColumns, DataColumn())(AddressOf GetGroupColumns, New DataGetGroupColumns(dtLineas, enummcAgrupFactura.mcCliente), services)

            Dim groupers(1) As GroupHelper
            groupers(enummcAgrupFactura.mcAlbaran) = New GroupHelper(grpAlb, oGrprUser)
            groupers(enummcAgrupFactura.mcCliente) = New GroupHelper(grpClte, oGrprUser)

            For Each rwLin As DataRow In dtLineas.Rows
                groupers(rwLin("AgrupFactura")).Group(rwLin)
            Next
            Return oGrprUser.Fras
        Else
            ApplicationService.GenerateError("No hay datos a Facturar. Revise sus Albaranes.")
        End If

    End Function


    <Task()> Public Shared Function AgruparAlbaranesAutoFra(ByVal data As DataPrcAutofacturacion, ByVal services As ServiceProvider) As FraCabAlbaran()

        Dim dtLineas As DataTable

        'se seleccionan todas las lineas de albaran no facturadas
        Dim htLins As New Hashtable
        Dim ids(data.IDAlbaranes.Length - 1) As Object
        For i As Integer = 0 To data.IDAlbaranes.Length - 1
            ids(i) = data.IDAlbaranes(i).IDLineaAlbaran
            htLins.Add(data.IDAlbaranes(i).IDLineaAlbaran, data.IDAlbaranes(i))
        Next

        Dim strViewName As String = "vNegComercialAutoFacturacion"

        If ids.Length > 0 Then
            Dim oFltr As New Filter
            oFltr.Add(New InListFilterItem("IDLineaAlbaran", ids, FilterType.Numeric))
            oFltr.Add(New NumberFilterItem("EstadoFactura", FilterOperator.NotEqual, enumavlEstadoFactura.avlFacturado))
            dtLineas = New BE.DataEngine().Filter(strViewName, oFltr)
        End If

        If Not dtLineas Is Nothing AndAlso dtLineas.Rows.Count > 0 Then
            Dim p As New Parametro
            Dim fvcFecha As enumfvcFechaAlbaran = p.FacturacionFechaAlbaran()
            Dim strCondicionPago As String = p.CondicionPago

            Dim oGrprUser As New GroupUserAlbaranes(data.DteFechaFactura, fvcFecha)

            Dim grpAlb As DataColumn() = ProcessServer.ExecuteTask(Of DataGetGroupColumns, DataColumn())(AddressOf GetGroupColumns, New DataGetGroupColumns(dtLineas, enummcAgrupFactura.mcAlbaran), services)
            Dim grpClte As DataColumn() = ProcessServer.ExecuteTask(Of DataGetGroupColumns, DataColumn())(AddressOf GetGroupColumns, New DataGetGroupColumns(dtLineas, enummcAgrupFactura.mcCliente), services)

            Dim groupers(1) As GroupHelper
            groupers(enummcAgrupFactura.mcAlbaran) = New GroupHelper(grpAlb, oGrprUser)
            groupers(enummcAgrupFactura.mcCliente) = New GroupHelper(grpClte, oGrprUser)

            For Each rwLin As DataRow In dtLineas.Rows
                groupers(rwLin("AgrupFactura")).Group(rwLin)
            Next

            For Each fra As FraCabAlbaran In oGrprUser.Fras
                For Each fralin As FraLinAlbaran In fra.Lineas
                    fralin.QaFacturar = DirectCast(htLins(fralin.IDLineaAlbaran), DataAutoFact).QFacturar
                    fralin.QIntAFacturar = DirectCast(htLins(fralin.IDLineaAlbaran), DataAutoFact).QIntAFacturar
                Next
            Next

            Return oGrprUser.Fras
        Else
            ApplicationService.GenerateError("No hay datos a Facturar. Revise sus Albaranes.")
        End If

    End Function

    Public Class DataGetGroupColumns
        Public Table As DataTable
        Public Agrupacion As enummcAgrupFactura

        Public Sub New(ByVal Table As DataTable, ByVal Agrupacion As enummcAgrupFactura)
            Me.Table = Table
            Me.Agrupacion = Agrupacion
        End Sub
    End Class
    <Task()> Public Shared Function GetGroupColumns(ByVal data As DataGetGroupColumns, ByVal services As ServiceProvider) As DataColumn()
        Dim columns(8) As DataColumn
        columns(0) = data.Table.Columns("IDCliente")
        columns(1) = data.Table.Columns("IDFormaPago")
        columns(2) = data.Table.Columns("IDCondicionPago")
        columns(3) = data.Table.Columns("IDBancoPropio") '// banco propio??
        columns(4) = data.Table.Columns("IdMoneda")
        columns(5) = data.Table.Columns("EDI")
        'columns(6) = table.Columns("DtoAlbaran")
        columns(6) = data.Table.Columns("IDDireccion")
        columns(7) = data.Table.Columns("IDDireccionFra")
        columns(8) = data.Table.Columns("IDClienteBanco")
        If data.Agrupacion = enummcAgrupFactura.mcAlbaran Then
            ReDim Preserve columns(9)
            columns(9) = data.Table.Columns("IDAlbaran")
        End If
        Return columns
    End Function

#End Region
#Region "Ordenar Facturas"
    'Ordena las facturas teniendo en cuenta las fechas de los albaranes
    <Task()> Public Shared Sub Ordenar(ByVal data As FraCabAlbaran(), ByVal services As ServiceProvider)
        If data IsNot Nothing Then Array.Sort(data, New OrdenFacturas)
    End Sub
#End Region
#Region "Proceso Facturación "
    <Task()> Public Shared Function CrearDocumentoFactura(ByVal fra As FraCab, ByVal services As ServiceProvider) As DocumentoFacturaVenta
        Return New DocumentoFacturaVenta(fra, services)
    End Function

    <Task()> Public Shared Sub AsignarValoresPredeterminadosGenerales(ByVal fra As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        Dim fraRow As DataRow = fra.HeaderRow
        If fraRow.IsNull("IDFactura") Then fraRow("IDFactura") = AdminData.GetAutoNumeric
        If fraRow.IsNull("FechaFactura") Then fraRow("FechaFactura") = Date.Today
        If fraRow.IsNull("FechaDeclaracionManual") Then fraRow("FechaDeclaracionManual") = False
        If fraRow.IsNull("FechaParaDeclaracion") Then
            fraRow("FechaParaDeclaracion") = fraRow("FechaFactura")
            ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ProcesoFacturacionVenta.FechaParaDeclaracionComoProveedor, New DataRowPropertyAccessor(fraRow), services)
        End If

        If fraRow.IsNull("Estado") Then fraRow("Estado") = enumfvcEstado.fvcNoContabilizado
        If fraRow.IsNull("IVAManual") Then fraRow("IVAManual") = False
        If fraRow.IsNull("VencimientosManuales") Then fraRow("VencimientosManuales") = False
        If fraRow.IsNull("Enviar347") Then fraRow("Enviar347") = False
        If fraRow.IsNull("Tipofactura") Then fraRow("Tipofactura") = enumfvcTipoFactura.fvcNormal
    End Sub

    <Task()> Public Shared Sub AsignarClienteGrupo(ByVal fra As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
        If fra.Cliente Is Nothing Then fra.Cliente = Clientes.GetEntity(fra.HeaderRow("IDCliente"))
        If Len(fra.Cliente.GrupoCliente) > 0 And fra.Cliente.GrupoFactura Then
            fra.HeaderRow("IDClienteInicial") = fra.HeaderRow("IDCliente")
            fra.HeaderRow("IDCliente") = fra.Cliente.GrupoCliente
            fra.Cliente = Clientes.GetEntity(fra.HeaderRow("IDCliente"))
            fra.HeaderRow("IDClienteBanco") = System.DBNull.Value
        End If
    End Sub

    <Task()> Public Shared Sub AsignarDatosCliente(ByVal fra As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        'AsignarDatosCliente(New DataRowPropertyAccessor(fra.HeaderRow), fra.Info)
        Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
        If fra.Cliente Is Nothing Then fra.Cliente = Clientes.GetEntity(fra.HeaderRow("IDCliente"))
        If fra.HeaderRow.IsNull("IDClienteInicial") Then fra.HeaderRow("IDClienteInicial") = fra.Cliente.IDCliente
        If fra.HeaderRow.IsNull("CifCliente") Then fra.HeaderRow("CifCliente") = fra.Cliente.CifCliente
        If fra.HeaderRow.IsNull("RazonSocial") Then fra.HeaderRow("RazonSocial") = fra.Cliente.RazonSocial
        If fra.HeaderRow.IsNull("Direccion") Then fra.HeaderRow("Direccion") = fra.Cliente.Direccion
        If fra.HeaderRow.IsNull("CodPostal") Then fra.HeaderRow("CodPostal") = fra.Cliente.CodPostal
        If fra.HeaderRow.IsNull("Poblacion") Then fra.HeaderRow("Poblacion") = fra.Cliente.Poblacion
        If fra.HeaderRow.IsNull("Provincia") Then fra.HeaderRow("Provincia") = fra.Cliente.Provincia
        If fra.HeaderRow.IsNull("IDPais") Then fra.HeaderRow("IDPais") = fra.Cliente.Pais
        'If fra.HeaderRow.IsNull("Telefono") Then fra.HeaderRow("Telefono") = fra.Cliente.Telefono
        'If fra.HeaderRow.IsNull("Fax") Then fra.HeaderRow("Fax") = fra.Cliente.Fax
        If fra.HeaderRow.IsNull("IDTipoAsiento") Then fra.HeaderRow("IDTipoAsiento") = fra.Cliente.IDTipoAsiento
        If fra.HeaderRow.IsNull("RetencionIRPF") Then fra.HeaderRow("RetencionIRPF") = fra.Cliente.RetencionIRPF
        ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoComercial.AsignarDatosCliente, fra, services)
        If fra.HeaderRow.IsNull("IDDiaPago") Then fra.HeaderRow("IDDiaPago") = fra.Cliente.DiaPago
        If fra.HeaderRow.IsNull("DtoFactura") Then fra.HeaderRow("DtoFactura") = fra.Cliente.DtoComercial
        If fra.HeaderRow.IsNull("IDBancoPropio") Then fra.HeaderRow("IDBancoPropio") = fra.Cliente.IDBancoPropio

        If Length(fra.HeaderRow("IdCondicionPago")) Then
            If fra.HeaderRow.IsNull("DtoProntoPago") Or fra.HeaderRow.IsNull("RecFinan") Then
                Dim rw As DataRow = New CondicionPago().GetItemRow(fra.HeaderRow("IdCondicionPago"))
                fra.HeaderRow("DtoProntoPago") = rw("DtoProntoPago")
                fra.HeaderRow("RecFinan") = rw("RecFinan")
            End If
        Else
            fra.HeaderRow("DtoProntoPago") = 0
            fra.HeaderRow("RecFinan") = 0
        End If
        If fra.HeaderRow.IsNull("IDProveedor") Then fra.HeaderRow("IDProveedor") = fra.Cliente.IDProveedor
        If fra.HeaderRow.IsNull("IDOperario") Then fra.HeaderRow("IDOperario") = fra.Cliente.IDOperario
    End Sub

    <Task()> Public Shared Sub AsignarDireccion(ByVal fra As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        'Primero se comprueba si nos está llegando una dirección de factura específica desde el albarán
        'Segundo se comprueba si la dirección de envío del albaráne s también dirección de factura
        'Tercero si no se han cumplido los pasos anteriores se utiliza la predetermianda del cliente de factura

        '1º Dirección establecida en el albarán
        If fra.Cabecera.IDDireccionFra <> 0 Then
            Dim StDatosDirec As New ClienteDireccion.DataDirecDe(fra.Cabecera.IDDireccionFra, enumcdTipoDireccion.cdDireccionFactura)
            If ProcessServer.ExecuteTask(Of ClienteDireccion.DataDirecDe, Boolean)(AddressOf ClienteDireccion.EsDireccionDe, StDatosDirec, services) Then
                fra.HeaderRow("IDDireccion") = fra.Cabecera.IDDireccion
                Dim DtClieDirec As DataTable = New ClienteDireccion().SelOnPrimaryKey(fra.HeaderRow("IDDireccion"))
                fra.HeaderRow("IDOficinaContable") = Nz(DtClieDirec.Rows(0)("IDOficinaContable"), String.Empty)
                fra.HeaderRow("IDOrganoGestor") = Nz(DtClieDirec.Rows(0)("IDOrganoGestor"), String.Empty)
                fra.HeaderRow("IDUnidadTramitadora") = Nz(DtClieDirec.Rows(0)("IDUnidadTramitadora"), String.Empty)
                Exit Sub
            End If
        End If

        '2º Dirección de envío que a la vez es dirección de factura
        If Not fra.HeaderRow.IsNull("IDDireccion") And fra.HeaderRow("IDCliente") = fra.HeaderRow("IDClienteInicial") Then
            Dim StDatosDirec As New ClienteDireccion.DataDirecDe(fra.HeaderRow("IDDireccion"), enumcdTipoDireccion.cdDireccionFactura)
            If ProcessServer.ExecuteTask(Of ClienteDireccion.DataDirecDe, Boolean)(AddressOf ClienteDireccion.EsDireccionDe, StDatosDirec, services) Then
                Dim DtClieDirec As DataTable = New ClienteDireccion().SelOnPrimaryKey(fra.HeaderRow("IDDireccion"))
                fra.HeaderRow("IDOficinaContable") = Nz(DtClieDirec.Rows(0)("IDOficinaContable"), String.Empty)
                fra.HeaderRow("IDOrganoGestor") = Nz(DtClieDirec.Rows(0)("IDOrganoGestor"), String.Empty)
                fra.HeaderRow("IDUnidadTramitadora") = Nz(DtClieDirec.Rows(0)("IDUnidadTramitadora"), String.Empty)
                Exit Sub
            End If
        End If

        '3º Dirección de factura predeterminada del cliente
        Dim StDatosDirecEnv As New ClienteDireccion.DataDirecEnvio(fra.HeaderRow("IDCliente"), enumcdTipoDireccion.cdDireccionFactura)
        Dim DtDirec As DataTable = ProcessServer.ExecuteTask(Of ClienteDireccion.DataDirecEnvio, DataTable)(AddressOf ClienteDireccion.ObtenerDireccionEnvio, StDatosDirecEnv, services)
        fra.HeaderRow("IDDireccion") = DtDirec.Rows(0)("IDDireccion")
        fra.HeaderRow("IDOficinaContable") = Nz(DtDirec.Rows(0)("IDOficinaContable"), String.Empty)
        fra.HeaderRow("IDOrganoGestor") = Nz(DtDirec.Rows(0)("IDOrganoGestor"), String.Empty)
        fra.HeaderRow("IDUnidadTramitadora") = Nz(DtDirec.Rows(0)("IDUnidadTramitadora"), String.Empty)
    End Sub

    <Task()> Public Shared Sub AsignarDatosFiscales(ByVal fra As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If fra.HeaderRow("Estado") = enumfvcEstado.fvcNoContabilizado Then
            Dim oCD As ClienteDireccion = New ClienteDireccion
            If Not fra.Cabecera Is Nothing AndAlso fra.Cabecera.IDDireccionFra <> 0 Then
                fra.HeaderRow("IDDireccion") = fra.Cabecera.IDDireccionFra
            End If
            Dim drDireccion As DataRow = oCD.GetItemRow(fra.HeaderRow("IDDireccion"))
            If drDireccion.Table.Columns.Contains("DomicilioFiscal") AndAlso Nz(drDireccion("DomicilioFiscal"), False) Then
                fra.HeaderRow("CifCliente") = drDireccion("CifCliente")
                fra.HeaderRow("RazonSocial") = drDireccion("RazonSocial")
                fra.HeaderRow("Direccion") = drDireccion("Direccion")
                fra.HeaderRow("CodPostal") = drDireccion("CodPostal")
                fra.HeaderRow("Poblacion") = drDireccion("Poblacion")
                fra.HeaderRow("Provincia") = drDireccion("Provincia")
                fra.HeaderRow("IDPais") = drDireccion("IDPais")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarBanco(ByVal fra As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If fra.HeaderRow("Estado") = enumfvcEstado.fvcNoContabilizado Then
            If fra.HeaderRow.IsNull("IDClienteBanco") Then
                Dim IDBanco As Integer = ProcessServer.ExecuteTask(Of String, Integer)(AddressOf New ClienteBanco().GetBancoPredeterminado, fra.HeaderRow("IDCliente"), services)
                If IDBanco > 0 Then
                    fra.HeaderRow("IDClienteBanco") = IDBanco
                Else
                    fra.HeaderRow("IDClienteBanco") = System.DBNull.Value
                End If
            End If

            ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf AsignarBancoPropio, fra, services)
        End If
    End Sub

    <Task()> Public Shared Sub AsignarBancoPropio(ByVal fra As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If Not fra.Cabecera Is Nothing AndAlso Length(fra.Cabecera.IDBancoPropio) > 0 Then
            fra.HeaderRow("IDBancoPropio") = fra.Cabecera.IDBancoPropio
        End If
    End Sub

    <Task()> Public Shared Sub AsignarContador(ByVal fra As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If fra.HeaderRow.IsNull("IDContador") Then
            Dim Info As ProcessInfo = services.GetService(Of ProcessInfo)()
            If Length(Info.IDContador) > 0 Then
                fra.HeaderRow("IDContador") = Info.IDContador
            Else
                If Length(fra.Cliente.IDContadorCargo) > 0 Then
                    fra.HeaderRow("IDContador") = fra.Cliente.IDContadorCargo
                Else
                    If Length(Info.IDContadorEntidad) > 0 Then
                        fra.HeaderRow("IDContador") = Info.IDContadorEntidad
                    Else
                        ProcessServer.ExecuteTask(Of DocumentCabLin)(AddressOf ProcesoComunes.AsignarContador, fra, services)
                    End If
                End If
            End If
        End If

        Dim counters As ProvisionalCounter = services.GetService(Of ProvisionalCounter)()
        fra.HeaderRow("NFactura") = counters.GetCounterValue(fra.HeaderRow("IDContador"))
        fra.AIva = counters.GetCounter(fra.HeaderRow("IDContador"))("AIva") ' New Contador().GetItemRow(fra.HeaderRow("IDContador"))("AIva")
    End Sub


    <Task()> Public Shared Sub AsignarNumeroFactura(ByVal fra As DocumentoFacturaVenta, ByVal services As ServiceProvider)

        If fra.HeaderRow.RowState = DataRowState.Added Then
            If Not IsDBNull(fra.HeaderRow("IDContador")) Then
                Dim StDatos As New Contador.DatosCounterValue
                StDatos.IDCounter = fra.HeaderRow("IDContador")
                StDatos.TargetClass = New FacturaVentaCabecera
                StDatos.TargetField = "NFactura"
                StDatos.DateField = "FechaFactura"
                StDatos.DateValue = fra.HeaderRow("FechaFactura")
                StDatos.IDEjercicio = fra.HeaderRow("IDEjercicio") & String.Empty
                fra.HeaderRow("NFactura") = ProcessServer.ExecuteTask(Of Contador.DatosCounterValue, String)(AddressOf Contador.CounterValue, StDatos, services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarNumeroFacturaPropuesta(ByVal fra As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        Dim AppParams As ParametroVenta = services.GetService(Of ParametroVenta)()
        Dim TPVFactura As Boolean = AppParams.TPVFactura
        If TypeOf fra.Cabecera Is FraCabAlbaran AndAlso Length(CType(fra.Cabecera, FraCabAlbaran).IDTPV) > 0 AndAlso (TPVFactura OrElse CType(fra.Cabecera, FraCabAlbaran).AgrupFactura = enummcAgrupFactura.mcAlbaran) Then
            fra.HeaderRow("NFactura") = CType(fra.Cabecera, FraCabAlbaran).NAlbaran
        Else
            If fra.HeaderRow.IsNull("NFactura") Then
                Dim counters As ProvisionalCounter = services.GetService(Of ProvisionalCounter)()
                fra.HeaderRow("NFactura") = counters.GetCounterValue(fra.HeaderRow("IDContador"))
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarClaveOperacion(ByVal fra As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If fra.HeaderRow("Estado") = enumfvcEstado.fvcNoContabilizado Then
            If Not IsDBNull(fra.HeaderRow("IDContador")) Then
                Dim Contadores As EntityInfoCache(Of ContadorInfo) = services.GetService(Of EntityInfoCache(Of ContadorInfo))()
                Dim ContInfo As ContadorInfo = Contadores.GetEntity(fra.HeaderRow("IDContador"))
                If Length(ContInfo.IDTipoComprobante) > 0 AndAlso Length(ContInfo.ClaveOperacion) > 0 Then
                    '//Le asignamos la clave de operación del Tipo de Comprobante asociado al contador.
                    fra.HeaderRow("ClaveOperacion") = ContInfo.ClaveOperacion
                    Exit Sub
                End If
            End If

            If Not fra.dtFVBI Is Nothing AndAlso fra.dtFVBI.Rows.Count > 0 Then
                Dim fLineaBI As New Filter
                fLineaBI.Add(New NumberFilterItem("BaseImponible", FilterOperator.NotEqual, 0))
                Dim WhereLineaBI As String = fLineaBI.Compose(New AdoFilterComposer)
                Dim adr() As DataRow = fra.dtFVBI.Select(WhereLineaBI, Nothing, DataViewRowState.CurrentRows)
                If adr.Length > 1 Then
                    fra.HeaderRow("ClaveOperacion") = ClaveOperacion.FacturaVariosTiposImpositivos
                Else
                    'Si hay un solo tipo impositivo y la clave de operación es varios Tipos se resetea a nulo. 
                    'En otro caso se respeta la clave de operación que viene.
                    If Not fra.HeaderRow.IsNull("ClaveOperacion") AndAlso (fra.HeaderRow("ClaveOperacion") = ClaveOperacion.FacturaVariosTiposImpositivos Or fra.HeaderRow("ClaveOperacion") = ClaveOperacion.InversionSujetoPasivo Or fra.HeaderRow("ClaveOperacion") = ClaveOperacion.FacturaRectificativa Or fra.HeaderRow("ClaveOperacion") = ClaveOperacion.Tickets Or fra.HeaderRow("ClaveOperacion") = ClaveOperacion.ResumenTicket) Then
                        fra.HeaderRow("ClaveOperacion") = System.DBNull.Value
                    End If
                    'Si hay un solo tipo impositivo y es ISP la clave operación será InversionSujetoPasivo
                    If adr.Length = 1 AndAlso Length(adr(0)("IDtipoIVA")) > 0 Then
                        Dim TipoIVA As New TipoIva
                        Dim dtTipoIVA As DataTable = TipoIVA.SelOnPrimaryKey(adr(0)("IDtipoIVA"))
                        If Not dtTipoIVA Is Nothing AndAlso dtTipoIVA.Rows.Count > 0 Then
                            If dtTipoIVA(0)("SinRepercutir") = True AndAlso Nz(dtTipoIVA(0)("IVASinRepercutir"), 0) <> 0 Then
                                fra.HeaderRow("ClaveOperacion") = ClaveOperacion.InversionSujetoPasivo
                            End If
                        End If
                    End If
                End If
            End If

            If Length(fra.HeaderRow("IDFacturaRectificada")) > 0 Then
                fra.HeaderRow("ClaveOperacion") = ClaveOperacion.FacturaRectificativa
            End If

            If Not fra.dtLineas Is Nothing AndAlso fra.dtLineas.Rows.Count > 0 Then
                Dim IDAlbaranes(-1) As Object : Dim IDAlbaranAnt As Integer
                Dim f As New Filter
                f.Add(New IsNullFilterItem("IDAlbaran", False))
                Dim WhereNotNullAlbaran As String = f.Compose(New AdoFilterComposer)
                For Each drLineaFV As DataRow In fra.dtLineas.Select(WhereNotNullAlbaran, "IDAlbaran")
                    If IDAlbaranAnt <> drLineaFV("IDAlbaran") Then
                        IDAlbaranAnt = drLineaFV("IDAlbaran")
                        ReDim Preserve IDAlbaranes(UBound(IDAlbaranes) + 1)
                        IDAlbaranes(UBound(IDAlbaranes)) = drLineaFV("IDAlbaran")
                    End If
                Next

                If IDAlbaranes.Length > 0 Then
                    f.Clear()
                    f.Add(New InListFilterItem("IDAlbaran", IDAlbaranes, FilterType.Numeric))
                    f.Add(New BooleanFilterItem("Ticket", True))
                    Dim dtAVC As DataTable = New AlbaranVentaCabecera().Filter(f)
                    If Not dtAVC Is Nothing AndAlso dtAVC.Rows.Count > 0 Then
                        If IDAlbaranes.Length = 1 Then
                            fra.HeaderRow("ClaveOperacion") = ClaveOperacion.Tickets
                        Else
                            fra.HeaderRow("ClaveOperacion") = ClaveOperacion.ResumenTicket
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarDescuentoFactura(ByVal doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If doc.HeaderRow("TipoFactura") = enumTipoFactura.tfNormal Then doc.HeaderRow("DtoFactura") = doc.Cabecera.Dto
    End Sub


    <Task()> Public Shared Sub CrearLineasDesdeAlbaran(ByVal docFactura As DocumentoFacturaVenta, ByVal services As ServiceProvider)

        Dim dtAlbaran As DataTable = ProcessServer.ExecuteTask(Of DocumentoFacturaVenta, DataTable)(AddressOf RecuperarDatosAlbaran, docFactura, services)
        Dim lineas As DataTable = docFactura.dtLineas
        If lineas Is Nothing Then
            Dim oFVL As New FacturaVentaLinea
            lineas = oFVL.AddNew
            docFactura.Add(GetType(FacturaVentaLinea).Name, lineas)
        End If

        Dim fraCabAlb As FraCabAlbaran = docFactura.Cabecera
        For Each albaran As DataRow In dtAlbaran.Rows
            Dim fralin As FraLinAlbaran = Nothing
            For i As Integer = 0 To fraCabAlb.Lineas.Length - 1
                If albaran("IDLineaAlbaran") = fraCabAlb.Lineas(i).IDLineaAlbaran Then
                    fralin = fraCabAlb.Lineas(i)
                    Exit For
                End If
            Next
            If Not fralin Is Nothing Then
                Dim dblCantidad As Double
                If Double.IsNaN(fralin.QaFacturar) Then
                    dblCantidad = albaran("QServida") - albaran("QFacturada")
                Else
                    dblCantidad = fralin.QaFacturar
                End If

                If dblCantidad <> 0 Then
                    Dim linea As DataRow = lineas.NewRow

                    ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarValoresPredeterminadosLinea, linea, services)

                    linea("IDFactura") = docFactura.HeaderRow("IDFactura")
                    linea("IDPedido") = albaran("IDPedido")
                    linea("IDLineaPedido") = albaran("IDLineaPedido")
                    linea("PedidoCliente") = albaran("PedidoCliente")
                    linea("IdLineaAlbaran") = albaran("IdLineaAlbaran")
                    linea("IDAlbaran") = albaran("IDAlbaran")
                    linea("IDArticulo") = albaran("IDArticulo")
                    linea("DescArticulo") = albaran("DescArticulo")
                    linea("RefCliente") = albaran("RefCliente")
                    linea("DescRefCliente") = albaran("DescRefCliente")
                    linea("Revision") = albaran("Revision")
                    linea("lote") = albaran("lote")
                    linea("IDTipoIva") = albaran("IDTipoIva")
                    linea("PVP") = albaran("PVP")
                    linea("ImportePVP") = albaran("ImportePVP")
                    If Length(albaran("IDCentroGestion")) > 0 Then
                        linea("IDCentroGestion") = albaran("IDCentroGestion")
                    Else
                        Dim drCabecera As DataRow = New AlbaranVentaCabecera().GetItemRow(albaran("IDAlbaran"))
                        linea("IDCentroGestion") = drCabecera("IDCentroGestion")
                    End If
                    linea("Cantidad") = dblCantidad
                    '''
                    'linea("Factor") = albaran("Factor")
                    'linea("QInterna") = dblCantidad * albaran("Factor")
                    If TypeOf fralin Is FraLinAlbaran AndAlso Not Double.IsNaN(fralin.QIntAFacturar) Then
                        '//Autofacturación
                        If linea("Cantidad") <> 0 Then
                            linea("QInterna") = CType(fralin, FraLinAlbaran).QIntAFacturar
                            linea("Factor") = linea("QInterna") / linea("Cantidad")
                        Else
                            linea("QInterna") = 0
                            linea("Factor") = 0
                        End If
                    Else
                        '//Facturación general
                        linea("Factor") = albaran("Factor")
                        linea("QInterna") = dblCantidad * albaran("Factor")
                    End If

                    ''''
                    If ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, linea("IDArticulo"), services) AndAlso Length(albaran("IdLineaAlbaran")) > 0 Then
                        If linea.Table.Columns.Contains("QInterna2") Then
                            Dim AVL As New AlbaranVentaLinea
                            Dim dtAVL As DataTable = AVL.SelOnPrimaryKey(albaran("IdLineaAlbaran"))
                            If dtAVL.Rows.Count > 0 AndAlso dtAVL.Columns.Contains("QInterna2") AndAlso Length(dtAVL.Rows(0)("QInterna2")) > 0 Then
                                linea("QInterna2") = dtAVL.Rows(0)("QInterna2")
                            End If
                        End If
                    End If
                    linea("IDUDInterna") = albaran("IDUDInterna")
                    linea("IDUDMedida") = albaran("IDUDMedida")
                    linea("UdValoracion") = albaran("UdValoracion")
                    linea("Precio") = albaran("Precio")
                    linea("Regalo") = albaran("Regalo")
                    linea("IDPromocionLinea") = albaran("IDPromocionLinea")
                    linea("IDPromocion") = albaran("IDPromocion")
                    linea("IDTarifa") = albaran("IDTarifa")

                    linea("Dto1") = albaran("Dto1")
                    linea("Dto2") = albaran("Dto2")
                    linea("Dto3") = albaran("Dto3")
                    linea("Dto") = albaran("Dto")
                    linea("DtoProntoPago") = albaran("DtoProntoPago")

                    linea("PrecioCosteA") = albaran("PrecioCosteA")
                    linea("PrecioCosteB") = albaran("PrecioCosteB")
                    linea("IDOrdenLinea") = albaran("IDOrdenLinea")

                    linea("IDLineaOfertaDetalle") = albaran("IDLineaOfertaDetalle")
                    linea("IDTipoLinea") = albaran("IDTipoLinea")
                    linea("SeguimientoTarifa") = albaran("SeguimientoTarifa")
                    linea("CContable") = albaran("CContable")
                    linea("IDTrabajo") = albaran("IDTrabajo")
                    linea("IDObra") = albaran("IDObra")
                    linea("IdLineaMaterial") = albaran("IdLineaMaterial")
                    linea("Texto") = albaran("Texto")

                    lineas.Rows.Add(linea)

                End If
            End If
        Next
    End Sub
    <Task()> Public Shared Sub CalcularImporteLineasFacturas(ByVal docFactura As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If docFactura.HeaderRow("Estado") = enumfvcEstado.fvcNoContabilizado Then
            ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoComunes.CalcularImporteLineas, docFactura, services)
        End If
    End Sub
    <Task()> Public Shared Sub CalcularRepresentantes(ByVal docFactura As DocumentoFacturaVenta, ByVal services As ServiceProvider)

        If docFactura.HeaderRow("Estado") = enumfvcEstado.fvcNoContabilizado Then
            ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoComercial.CalcularRepresentantes, docFactura, services)
        End If

    End Sub
    <Task()> Public Shared Sub CalcularAnalitica(ByVal docFactura As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If docFactura.HeaderRow("Estado") = enumfvcEstado.fvcNoContabilizado Then
            ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf NegocioGeneral.CalcularAnalitica, docFactura, services)
        End If

    End Sub
    <Task()> Public Shared Function RecuperarDatosAlbaran(ByVal docFactura As DocumentoFacturaVenta, ByVal services As ServiceProvider) As DataTable

        Dim fvl As New FacturaVentaLinea
        Dim avl As New AlbaranVentaLinea


        Dim fraCabAlb As FraCabAlbaran = docFactura.Cabecera

        Dim ids(fraCabAlb.Lineas.Length - 1) As Object
        For i As Integer = 0 To ids.Length - 1
            ids(i) = fraCabAlb.Lineas(i).IDLineaAlbaran
        Next

        Dim oFltr As New Filter
        oFltr.Add(New InListFilterItem("IDlineaAlbaran", ids, FilterType.Numeric))
        oFltr.Add(New NumberFilterItem("EstadoFactura", FilterOperator.NotEqual, enumavlEstadoFactura.avlFacturado))

        Return avl.Filter(oFltr)

    End Function

    <Task()> Public Shared Sub AsignarValoresPredeterminadosLinea(ByVal row As DataRow, ByVal services As ServiceProvider)
        row("IdLineaFactura") = AdminData.GetAutoNumeric
        row("Dto1") = 0
        row("Dto2") = 0
        row("Dto3") = 0
        row("Dto") = 0
        row("DtoProntoPago") = 0
        row("UdValoracion") = 1
        row("Factor") = 1
        Dim Info As ProcessInfoFra = services.GetService(Of ProcessInfoFra)()
        row("IDTipoLinea") = Info.TipoLineaDef
        row("QInterna") = 0
        row("Cantidad") = 0
        row("Regalo") = False
    End Sub
#End Region
#Region "Resultado a mostrar en la pantalla intermedia"
    '  Guardamos la información para visualizar la pantalla intermedia de facturas a generar
    <Task()> Public Shared Function Resultado(ByVal data As Object, ByVal services As ServiceProvider) As ResultFacturacion
        Dim InfoFra As ProcessInfoFra = services.GetService(Of ProcessInfoFra)()
        If InfoFra.ConPropuesta Then
            'Elimina la información almacenada en memoria si previamente hemos cancelado la facturación
            AdminData.GetSessionData("__frax__")
            'Guardamos la información del documento en memoria, para recuperarla cuando volvamos del preview de presentación
            AdminData.SetSessionData("__frax__", services.GetService(Of ArrayList))
        End If

        Return services.GetService(Of ResultFacturacion)()

    End Function
#End Region

#Region "Analítica y Representantes"

    <Task()> Public Shared Sub CopiarAnalitica(ByVal Doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If Doc.HeaderRow("Estado") <> enumfvcEstado.fvcContabilizado Then
            Dim AppParams As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
            If Not AppParams.Contabilidad OrElse Not AppParams.Analitica.AplicarAnalitica Then Exit Sub

            Dim IDOrigen(-1) As Object
            Dim f As New Filter
            f.Add(New IsNullFilterItem("IDLineaAlbaran", False))
            Dim WhereNotNullLineaAlbaran As String = f.Compose(New AdoFilterComposer)
            For Each linea As DataRow In Doc.dtLineas.Select(WhereNotNullLineaAlbaran)
                ReDim Preserve IDOrigen(IDOrigen.Length)
                IDOrigen(IDOrigen.Length - 1) = linea("IDLineaAlbaran")
            Next
            If Not IDOrigen Is Nothing AndAlso IDOrigen.Length > 0 Then
                Dim dtAnaliticaOrigen As DataTable = New AlbaranVentaAnalitica().Filter(New InListFilterItem("IDLineaAlbaran", IDOrigen, FilterType.Numeric))
                Dim datosCopia As New NegocioGeneral.DataCopiarAnalitica(dtAnaliticaOrigen, Doc)
                ProcessServer.ExecuteTask(Of NegocioGeneral.DataCopiarAnalitica)(AddressOf NegocioGeneral.CopiarAnalitica, datosCopia, services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CopiarRepresentantes(ByVal oDocFra As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoComercial.CopiarRepresentantes, oDocFra, services)
    End Sub

#End Region
#Region "Calcular factura"
    <Task()> Public Shared Sub CalcularBasesImponibles(ByVal Doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If Doc.HeaderRow("Estado") = enumfvcEstado.fvcNoContabilizado Then
            If Not IsNothing(Doc.dtLineas) Then
                Dim desglose() As DataBaseImponible = ProcessServer.ExecuteTask(Of DocumentoFacturaVenta, DataBaseImponible())(AddressOf ProcesoComunes.DesglosarImporte, Doc, services)


                Dim datosCalculo As New ProcesoComunes.DataCalculoTotalesCab(desglose, Doc)
                ProcessServer.ExecuteTask(Of ProcesoComunes.DataCalculoTotalesCab)(AddressOf CalcularTotalesCabecera, datosCalculo, services)
            End If
        End If
    End Sub


    <Task()> Public Shared Sub CalcularTotalesCabecera(ByVal data As ProcesoComunes.DataCalculoTotalesCab, ByVal services As ServiceProvider)
        If Not IsNothing(data.Doc.HeaderRow) Then
            'Dim fbi As New FacturaVentaBaseImponible
            Dim Bases As DataTable = CType(data.Doc, DocumentoFacturaVenta).dtFVBI
            Bases.DefaultView.Sort = "IDTipoIva"
            Dim IvaManual As Boolean = Nz(data.Doc.HeaderRow("IVAManual"), False)

            If Not IvaManual Then
                Dim notDeleted() As DataRow = Bases.Select(Nothing, Nothing, DataViewRowState.Added Or DataViewRowState.ModifiedCurrent Or DataViewRowState.Unchanged)
                For Each r As DataRow In notDeleted
                    r.Delete()
                Next
            End If

            Dim ImporteLineas As Double
            Dim ImporteLineasPVP As Double
            Dim BaseImponibleTotal As Double
            Dim ImporteIVATotal As Double
            Dim ImporteRETotal As Double
            Dim ImporteIntrastatTotal As Double

            Dim ImporteLineasA As Double
            Dim ImporteLineasPVPA As Double
            Dim BaseImponibleTotalA As Double
            Dim ImporteIVATotalA As Double
            Dim ImporteRETotalA As Double
            Dim ImporteIntrastatTotalA As Double

            Dim ImporteLineasB As Double
            Dim ImporteLineasPVPB As Double
            Dim BaseImponibleTotalB As Double
            Dim ImporteIVATotalB As Double
            Dim ImporteRETotalB As Double
            Dim ImporteIntrastatTotalB As Double


            If Not IsNothing(data.BasesImponibles) AndAlso data.BasesImponibles.Length > 0 Then
                Dim AppParams As ParametroFacturaVenta = services.GetService(Of ParametroFacturaVenta)()
                Dim contador As String = AppParams.ContadorAutofactura
                Dim ContadorAutofactura As Boolean

                If Len(contador) > 0 AndAlso Not IsDBNull(data.Doc.HeaderRow("IDContador")) Then
                    ContadorAutofactura = (contador = data.Doc.HeaderRow("IDContador"))
                End If

                Dim IVA As New TipoIva

                Dim AplicarRE As Boolean = CType(data.Doc, DocumentoFacturaVenta).Cliente.TieneRE

                For Each bi As DataBaseImponible In data.BasesImponibles
                    ImporteLineas = ImporteLineas + bi.BaseImponible
                    ImporteLineasA = ImporteLineasA + bi.BaseImponibleA
                    ImporteLineasB = ImporteLineasB + bi.BaseImponibleB

                    If Length(bi.IDTipoIva) > 0 Then
                        Dim Descuento As Double = 0
                        Dim Base As Double = 0
                        Dim BaseA As Double = 0
                        Dim BaseB As Double = 0
                        Dim factor As Double = 0
                        Dim lineaBase As DataRow = Nothing

                        Base = bi.BaseImponible
                        BaseA = bi.BaseImponibleA
                        BaseB = bi.BaseImponibleB

                        Dim AddNew As Boolean = False

                        Dim idx As Integer = Bases.DefaultView.Find(bi.IDTipoIva.Trim)
                        If idx >= 0 Then
                            lineaBase = Bases.DefaultView(idx).Row
                        Else
                            AddNew = True
                        End If

                        If AddNew Then
                            lineaBase = Bases.NewRow
                            lineaBase("IDBaseImponible") = AdminData.GetAutoNumeric
                            lineaBase("IDFactura") = data.Doc.HeaderRow("IDFactura")
                            lineaBase("IDTipoIva") = bi.IDTipoIva
                        End If
                        'Gestión cambio base imponible para ajustar facturas 
                        If Not IvaManual Or (IvaManual And AddNew) Then
                            lineaBase("BaseImponible") = Base
                            lineaBase("BaseImponibleA") = BaseA
                            lineaBase("BaseImponibleB") = BaseB
                        End If
                        If data.Doc.AIva Then
                            Dim TiposIVA As EntityInfoCache(Of TipoIvaInfo) = services.GetService(Of EntityInfoCache(Of TipoIvaInfo))()
                            ' HistoricoTipoIVA
                            Dim TIVAInfo As TipoIvaInfo = TiposIVA.GetEntity(lineaBase("IDTipoIva"), data.Doc.Fecha)

                            If AddNew Then
                                'valor por defecto
                                factor = TIVAInfo.Factor
                                If TIVAInfo.SinRepercutir And ContadorAutofactura Then
                                    'Nuevo para los ivas especiales que no se repercuten
                                    factor = TIVAInfo.IVASinRepercutir
                                End If
                            End If
                            If Not IvaManual Or (IvaManual And AddNew) Then
                                If bi.ImporteIVA <> 0 Then
                                    lineaBase("ImpIVA") = xRound(bi.ImporteIVA - bi.ImporteIVA * 100 / (100 + factor), data.Doc.Moneda.NDecimalesImporte)
                                    lineaBase("ImpIVAA") = xRound(bi.ImporteIVAA - bi.ImporteIVAA * 100 / (100 + factor), data.Doc.Moneda.NDecimalesImporte)
                                    lineaBase("ImpIVAB") = xRound(bi.ImporteIVAB - bi.ImporteIVAB * 100 / (100 + factor), data.Doc.Moneda.NDecimalesImporte)


                                    lineaBase("BaseImponible") = bi.ImporteIVA - xRound(bi.ImporteIVA - bi.ImporteIVA * 100 / (100 + factor), data.Doc.Moneda.NDecimalesImporte)
                                    lineaBase("BaseImponibleA") = bi.ImporteIVAA - xRound(bi.ImporteIVAA - bi.ImporteIVAA * 100 / (100 + factor), data.Doc.Moneda.NDecimalesImporte)
                                    lineaBase("BaseImponibleB") = bi.ImporteIVAB - xRound(bi.ImporteIVAB - bi.ImporteIVAB * 100 / (100 + factor), data.Doc.Moneda.NDecimalesImporte)

                                    ImporteLineasPVP = ImporteLineasPVP + lineaBase("BaseImponible")
                                    ImporteLineasPVPA = ImporteLineasPVPA + lineaBase("BaseImponibleA")
                                    ImporteLineasPVPB = ImporteLineasPVPB + lineaBase("BaseImponibleB")

                                    Base = lineaBase("BaseImponible")
                                    BaseA = lineaBase("BaseImponibleA")
                                    BaseB = lineaBase("BaseImponibleB")
                                Else
                                    lineaBase("ImpIVA") = xRound(Base * factor / 100, data.Doc.Moneda.NDecimalesImporte)
                                    lineaBase("ImpIVAA") = xRound(BaseA * factor / 100, data.Doc.MonedaA.NDecimalesImporte)
                                    lineaBase("ImpIVAB") = xRound(BaseB * factor / 100, data.Doc.MonedaB.NDecimalesImporte)
                                End If

                                If AplicarRE Then
                                    lineaBase("ImpRE") = xRound(Base * TIVAInfo.IVARE / 100, data.Doc.Moneda.NDecimalesImporte)
                                    lineaBase("ImpREA") = xRound(BaseA * TIVAInfo.IVARE / 100, data.Doc.MonedaA.NDecimalesImporte)
                                    lineaBase("ImpREB") = xRound(BaseB * TIVAInfo.IVARE / 100, data.Doc.MonedaB.NDecimalesImporte)
                                Else
                                    lineaBase("ImpRE") = 0
                                    lineaBase("ImpREA") = 0
                                    lineaBase("ImpREB") = 0
                                End If

                                lineaBase("ImpIntrastat") = xRound(Base * TIVAInfo.IVAIntrastat / 100, data.Doc.Moneda.NDecimalesImporte)
                                lineaBase("ImpIntrastatA") = xRound(BaseA * TIVAInfo.IVAIntrastat / 100, data.Doc.MonedaA.NDecimalesImporte)
                                lineaBase("ImpIntrastatB") = xRound(BaseB * TIVAInfo.IVAIntrastat / 100, data.Doc.MonedaB.NDecimalesImporte)
                            ElseIf IvaManual AndAlso Not AddNew Then
                                lineaBase("ImpIVA") = xRound(lineaBase("ImpIVA"), data.Doc.Moneda.NDecimalesImporte)
                                lineaBase("ImpIVAA") = xRound(lineaBase("ImpIVA") * data.Doc.CambioA, data.Doc.MonedaA.NDecimalesImporte)
                                lineaBase("ImpIVAB") = xRound(lineaBase("ImpIVA") * data.Doc.CambioB, data.Doc.MonedaB.NDecimalesImporte)

                                'Gestión cambio base imponible para ajustar facturas 
                                lineaBase("BaseImponible") = xRound(lineaBase("BaseImponible"), data.Doc.Moneda.NDecimalesImporte)
                                lineaBase("BaseImponibleA") = xRound(lineaBase("BaseImponible") * data.Doc.CambioA, data.Doc.MonedaA.NDecimalesImporte)
                                lineaBase("BaseImponibleB") = xRound(lineaBase("BaseImponible") * data.Doc.CambioB, data.Doc.MonedaB.NDecimalesImporte)
                            End If
                        End If


                        BaseImponibleTotal = BaseImponibleTotal + Nz(lineaBase("BaseImponible"), 0)
                        BaseImponibleTotalA = BaseImponibleTotalA + Nz(lineaBase("BaseImponibleA"), 0)
                        BaseImponibleTotalB = BaseImponibleTotalB + Nz(lineaBase("BaseImponibleB"), 0)

                        ImporteIVATotal = ImporteIVATotal + Nz(lineaBase("ImpIVA"), 0)
                        ImporteIVATotalA = ImporteIVATotalA + Nz(lineaBase("ImpIVAA"), 0)
                        ImporteIVATotalB = ImporteIVATotalB + Nz(lineaBase("ImpIVAB"), 0)

                        ImporteRETotal = ImporteRETotal + Nz(lineaBase("ImpRE"), 0)
                        ImporteRETotalA = ImporteRETotalA + Nz(lineaBase("ImpREA"), 0)
                        ImporteRETotalB = ImporteRETotalB + Nz(lineaBase("ImpREB"), 0)

                        ImporteIntrastatTotal = ImporteIntrastatTotal + Nz(lineaBase("ImpIntrastat"), 0)
                        ImporteIntrastatTotalA = ImporteIntrastatTotalA + Nz(lineaBase("ImpIntrastatA"), 0)
                        ImporteIntrastatTotalB = ImporteIntrastatTotalB + Nz(lineaBase("ImpIntrastatB"), 0)

                        If AddNew Then
                            Bases.Rows.Add(lineaBase)
                        End If
                    End If
                Next
            End If
            data.Doc.HeaderRow("BaseImponible") = BaseImponibleTotal
            data.Doc.HeaderRow("BaseImponibleA") = BaseImponibleTotalA
            data.Doc.HeaderRow("BaseImponibleB") = BaseImponibleTotalB

            data.Doc.HeaderRow("ImpIVA") = ImporteIVATotal
            data.Doc.HeaderRow("ImpIVAA") = ImporteIVATotalA
            data.Doc.HeaderRow("ImpIVAB") = ImporteIVATotalB

            data.Doc.HeaderRow("ImpRE") = ImporteRETotal
            data.Doc.HeaderRow("ImpREA") = ImporteRETotalA
            data.Doc.HeaderRow("ImpREB") = ImporteRETotalB

            data.Doc.HeaderRow("ImpIntrastat") = ImporteIntrastatTotal
            data.Doc.HeaderRow("ImpIntrastatA") = ImporteIntrastatTotalA
            data.Doc.HeaderRow("ImpIntrastatB") = ImporteIntrastatTotalB

            If ImporteLineasPVP <> 0 Then
                ImporteLineas = ImporteLineasPVP
                ImporteLineasA = ImporteLineasPVPA
                ImporteLineasB = ImporteLineasPVPB
            End If

            data.Doc.HeaderRow("ImpLineas") = ImporteLineas
            data.Doc.HeaderRow("ImpLineasA") = ImporteLineasA
            data.Doc.HeaderRow("ImpLineasB") = ImporteLineasB

            Dim ImporteImpuestosTotal As Double = 0
            Dim ImporteImpuestosTotalA As Double = 0
            Dim ImporteImpuestosTotalB As Double = 0
            Dim dtImpuestos As DataTable = CType(data.Doc, DocumentoFacturaVenta).dtImpuestos
            If dtImpuestos.Rows.Count > 0 Then
                ImporteImpuestosTotal += Nz(dtImpuestos.Compute("SUM(Importe)", Nothing), 0)
                ImporteImpuestosTotalA += Nz(dtImpuestos.Compute("SUM(ImporteA)", Nothing), 0)
                ImporteImpuestosTotalB += Nz(dtImpuestos.Compute("SUM(ImporteB)", Nothing), 0)
            End If

            data.Doc.HeaderRow("ImpImpuestos") = ImporteImpuestosTotal
            data.Doc.HeaderRow("ImpImpuestosA") = ImporteImpuestosTotalA
            data.Doc.HeaderRow("ImpImpuestosB") = ImporteImpuestosTotalB
        End If
    End Sub

    <Task()> Public Shared Sub CalcularTotales(ByVal oCabFra As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If oCabFra.HeaderRow("Estado") = enumfvcEstado.fvcNoContabilizado Then
            If Not IsNothing(oCabFra) Then
                Dim factura As DataRow = oCabFra.HeaderRow
                Try
                    Dim RecFinan As Double
                    Dim RecFinanA As Double
                    Dim RecFinanB As Double
                    'Se calcula el total de la factura sumando todos los recargos a la Base Imponible.
                    'Se comprueba si RecFinan>0, en caso afirmativo re aplica el recargo (Se suma a la factura porque es un recargo)
                    'RetencionIRPF

                    Dim Total As Double = factura("BaseImponible") + factura("ImpIVA") + factura("ImpRE") + factura("ImpImpuestos")
                    Dim TotalA As Double = factura("BaseImponibleA") + factura("ImpIVAA") + factura("ImpREA") + factura("ImpImpuestosA")
                    Dim TotalB As Double = factura("BaseImponibleB") + factura("ImpIVAB") + factura("ImpREB") + factura("ImpImpuestosB")

                    If factura("RecFinan") > 0 Then
                        RecFinan = xRound(Total * factura("RecFinan") / 100, oCabFra.Moneda.NDecimalesImporte)
                        RecFinanA = xRound(TotalA * factura("RecFinan") / 100, oCabFra.MonedaA.NDecimalesImporte)
                        RecFinanB = xRound(TotalB * factura("RecFinan") / 100, oCabFra.MonedaB.NDecimalesImporte)

                        Total = Total + RecFinan
                        TotalA = TotalA + RecFinanA
                        TotalB = TotalB + RecFinanB


                        factura("ImpRecFinan") = RecFinan
                        factura("ImpRecFinanA") = RecFinanA
                        factura("ImpRecFinanB") = RecFinanB
                    Else
                        factura("ImpRecFinan") = 0
                        factura("ImpRecFinanA") = 0
                        factura("ImpRecFinanB") = 0
                    End If

                    If Length(factura("IDContador")) > 0 Then
                        If ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf NegocioGeneral.ContadorB, factura("IDContador"), services) Then factura("RetencionIRPF") = 0
                    Else
                        'Para importaciones de datos que vienen sin contador y necesitamos mantener el nfactura
                        ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoComunes.AsignarContador, oCabFra, services)
                    End If

                    ' Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()

                    If Nz(factura("RetencionIRPF"), 0) <> 0 Then
                        If factura.Table.Columns.Contains("RetencionManual") Then
                            If Not Nz(factura("RetencionManual"), False) Then
                                Dim fBasesRetencion As fImporte = ProcessServer.ExecuteTask(Of DocumentoFacturaVenta, fImporte)(AddressOf GetBasesRetencionLineas, oCabFra, services)
                                factura("BaseRetencion") = fBasesRetencion.Importe
                                factura("BaseRetencionA") = fBasesRetencion.ImporteA
                                factura("BaseRetencionB") = fBasesRetencion.ImporteB
                            Else
                                factura("BaseRetencion") = xRound(Nz(factura("BaseRetencion"), 0), oCabFra.Moneda.NDecimalesImporte)
                                factura("BaseRetencionA") = xRound(Nz(factura("BaseRetencionA"), 0), oCabFra.MonedaA.NDecimalesImporte)
                                factura("BaseRetencionB") = xRound(Nz(factura("BaseRetencionB"), 0), oCabFra.MonedaB.NDecimalesImporte)
                            End If
                            factura("ImpRetencion") = xRound(Nz(factura("BaseRetencion"), 0) * factura("RetencionIRPF") / 100, oCabFra.Moneda.NDecimalesImporte)
                            factura("ImpRetencionA") = xRound(Nz(factura("BaseRetencionA"), 0) * factura("RetencionIRPF") / 100, oCabFra.MonedaA.NDecimalesImporte)
                            factura("ImpRetencionB") = xRound(Nz(factura("BaseRetencionB"), 0) * factura("RetencionIRPF") / 100, oCabFra.MonedaB.NDecimalesImporte)
                        Else
                            Dim fBasesRetencion As fImporte = ProcessServer.ExecuteTask(Of DocumentoFacturaVenta, fImporte)(AddressOf GetBasesRetencionLineas, oCabFra, services)

                            factura("ImpRetencion") = xRound(fBasesRetencion.Importe * factura("RetencionIRPF") / 100, oCabFra.Moneda.NDecimalesImporte)
                            factura("ImpRetencionA") = xRound(fBasesRetencion.ImporteA * factura("RetencionIRPF") / 100, oCabFra.MonedaA.NDecimalesImporte)
                            factura("ImpRetencionB") = xRound(fBasesRetencion.ImporteB * factura("RetencionIRPF") / 100, oCabFra.MonedaB.NDecimalesImporte)
                        End If
                    Else
                        factura("ImpRetencion") = 0
                        factura("ImpRetencionA") = 0
                        factura("ImpRetencionB") = 0
                    End If

                    ' Inicio Para cuando nos retiene parte del cobro hasta la fecha de garantía
                    If Nz(factura("Retencion"), 0) > 0 Then
                        If Nz(factura("TipoRetencion"), 0) = enumTipoRetencion.troSobreBI Then
                            factura("ImpRetencionGar") = xRound(factura("BaseImponible") * factura("Retencion") / 100, oCabFra.Moneda.NDecimalesImporte)
                            factura("ImpRetencionGarA") = xRound(factura("BaseImponibleA") * factura("Retencion") / 100, oCabFra.MonedaA.NDecimalesImporte)
                            factura("ImpRetencionGarB") = xRound(factura("BaseImponibleB") * factura("Retencion") / 100, oCabFra.MonedaB.NDecimalesImporte)
                        Else
                            factura("ImpRetencionGar") = xRound((factura("BaseImponible") + factura("ImpIVA") + factura("ImpRE") + Nz(factura("ImpImpuestos"), 0)) * factura("Retencion") / 100, oCabFra.Moneda.NDecimalesImporte)
                            factura("ImpRetencionGarA") = xRound((factura("BaseImponibleA") + factura("ImpIVAA") + factura("ImpREA") + Nz(factura("ImpImpuestosA"), 0)) * factura("Retencion") / 100, oCabFra.MonedaA.NDecimalesImporte)
                            factura("ImpRetencionGarB") = xRound((factura("BaseImponibleB") + factura("ImpIVAB") + factura("ImpREB") + Nz(factura("ImpImpuestosB"), 0)) * factura("Retencion") / 100, oCabFra.MonedaB.NDecimalesImporte)
                        End If
                    Else
                        factura("ImpRetencionGar") = 0
                        factura("ImpRetencionGarA") = 0
                        factura("ImpRetencionGarB") = 0
                    End If

                    ' Fin Para cuando nos retiene parte del cobro hasta la fecha de garantía

                    factura("ImpTotal") = factura("BaseImponible") + factura("ImpIVA") + factura("ImpRE") + factura("ImpRecFinan") + Nz(factura("ImpImpuestos"), 0)
                    factura("ImpTotalA") = factura("BaseImponibleA") + factura("ImpIVAA") + factura("ImpREA") + factura("ImpRecFinanA") + Nz(factura("ImpImpuestosA"), 0)
                    factura("ImpTotalB") = factura("BaseImponibleB") + factura("ImpIVAB") + factura("ImpREB") + factura("ImpRecFinanB") + Nz(factura("ImpImpuestosB"), 0)

                Catch ex As Exception
                    ApplicationService.GenerateError("Compruebe los valores de los campos relacionados con el cálculo del Total de la Factura.|Nº FACTURA: ||Base Imponible, Recargo financiero, RetencionIRPF,.....", vbNewLine, factura("NFactura"), vbNewLine)
                End Try
            End If
        End If
    End Sub

    <Task()> Public Shared Function GetBasesRetencionLineas(ByVal Doc As DocumentoFacturaVenta, ByVal services As ServiceProvider) As fImporte
        Dim fImp As New fImporte
        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        For Each linea As DataRow In Doc.dtLineas.Rows
            Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(linea("IDArticulo"))
            If ArtInfo.RetencionIRPF Then
                fImp.Importe += Nz(linea("Importe"), 0)
                fImp.ImporteA += Nz(linea("ImporteA"), 0)
                fImp.ImporteB += Nz(linea("ImporteB"), 0)
            End If
        Next
        Return fImp
    End Function

    <Task()> Public Shared Sub CalcularPuntoVerde(ByVal oDocFra As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If oDocFra.HeaderRow("Estado") = enumfvcEstado.fvcNoContabilizado Then
            Dim factura As DataRow = oDocFra.HeaderRow
            Dim lineas As DataTable = oDocFra(GetType(FacturaVentaLinea).Name)

            Dim AcumPuntoVerde As Double
            Try
                Dim AppParams As ParametroFacturaVenta = services.GetService(Of ParametroFacturaVenta)()
                If AppParams.CalcularPuntoVerde Then
                    If Not IsNothing(lineas) Then
                        For Each drl As DataRow In lineas.Rows
                            If drl.RowState <> DataRowState.Deleted Then
                                AcumPuntoVerde = AcumPuntoVerde + (ProcessServer.ExecuteTask(Of String, Double)(AddressOf PuntoVerde, drl("IDArticulo"), services) * drl("Cantidad"))
                            End If
                        Next
                    End If
                    factura("ImpPuntoVerdeA") = AcumPuntoVerde

                    If factura("CambioA") > 0 Then
                        factura("ImpPuntoVerde") = AcumPuntoVerde / factura("CambioA")
                    End If
                    If factura("CambioB") > 0 Then
                        factura("ImpPuntoVerdeB") = factura("ImpPuntoVerde") * factura("CambioB")
                    End If
                End If
            Catch ex As Exception
                ApplicationService.GenerateError("Compruebe los valores de los campos relacionados con el Punto Verde.|Nº FACTURA: ||Artículo, Cantidad, Cambios de Monedas, ....", vbNewLine, factura("NFactura"), vbNewLine)
            End Try
        End If
    End Sub

    <Task()> Public Shared Function PuntoVerde(ByVal strIDArticulo As String, ByVal services As ServiceProvider) As Double
        Dim cmmComando As Common.DbCommand = AdminData.GetCommand
        cmmComando.CommandType = CommandType.StoredProcedure
        cmmComando.CommandText = "sp_PuntoVerdeExplosion"

        Dim prmParam As Common.DbParameter = cmmComando.CreateParameter
        cmmComando.Parameters.Add(prmParam)
        prmParam.ParameterName = "@pArticulo"
        prmParam.Value = strIDArticulo

        Dim prmPuntoVerde As Common.DbParameter = cmmComando.CreateParameter
        cmmComando.Parameters.Add(prmPuntoVerde)
        prmPuntoVerde.ParameterName = "@PuntoVerde"
        prmPuntoVerde.DbType = DbType.Double
        prmPuntoVerde.Direction = ParameterDirection.Output

        AdminData.Execute(cmmComando)

        Dim dblPuntoVerde As Double = Nz(prmPuntoVerde.Value, 0)

        Return dblPuntoVerde
    End Function

#End Region
#Region "Cálculo de Vencimientos "

    <Task()> Public Shared Sub CalcularVencimientos(ByVal FraCab As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If FraCab.HeaderRow("Estado") = enumfvcEstado.fvcNoContabilizado Then
            If Not FraCab Is Nothing AndAlso Not FraCab.HeaderRow Is Nothing Then
                If Not FraCab.HeaderRow("VencimientosManuales") Then
                    If ProcessServer.ExecuteTask(Of DocumentoFacturaVenta, Boolean)(AddressOf VieneDeTPV, FraCab, services) Then
                        ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf CobrosTPV, FraCab, services)
                    Else
                        ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf CobrosAutomaticos, FraCab, services)
                    End If
                Else
                    '//Si se han modificado las fechas de los vencimientos, tendremos que actualizar la cabecera.
                    Dim dtVtosModif As DataTable = FraCab.dtCobros.GetChanges
                    If Not dtVtosModif Is Nothing AndAlso dtVtosModif.Rows.Count > 0 AndAlso FraCab.dtCobros.Rows.Count > 0 Then
                        Dim FechaVto As Date = FraCab.dtCobros.Compute("MIN(FechaVencimiento)", Nothing)
                        FraCab.HeaderRow("FechaVencimiento") = FechaVto
                    End If
                End If
            End If
        End If
    End Sub
    <Task()> Public Shared Function VieneDeTPV(ByVal FraCab As DocumentoFacturaVenta, ByVal services As ServiceProvider) As Boolean

        If FraCab.dtLineas.Rows.Count > 0 AndAlso FraCab.dtLineas.Rows(0).RowState <> DataRowState.Deleted AndAlso Length(FraCab.dtLineas.Rows(0)("IDAlbaran")) > 0 Then
            Dim f As New Filter
            f.Add(New NumberFilterItem("IDAlbaran", FraCab.dtLineas.Rows(0)("IDAlbaran")))
            f.Add(New BooleanFilterItem("SegunCliente", False))
            Dim dtAlbaranFormaPago As DataTable = New BE.DataEngine().Filter("vNegTPVAlbaranFormaPago", f)
            If dtAlbaranFormaPago.Rows.Count > 0 Then
                VieneDeTPV = True
            End If
        End If
    End Function

    Public Class DataNuevosCobros
        Public Doc As DocumentoFacturaVenta
        Public DireccionCobro As Integer
        Public TipoCobro As Integer
        Public FechaVencimiento As Date
        Public ImporteVencimiento As fImporte
        Public ImporteRecFinanciero As fImporte

        Public Sub New(ByVal Doc As DocumentoFacturaVenta, ByVal DireccionCobro As Integer, ByVal TipoCobro As Integer)
            Me.Doc = Doc
            Me.DireccionCobro = DireccionCobro
            Me.TipoCobro = TipoCobro
        End Sub
    End Class
    <Task()> Public Shared Sub CobrosTPV(ByVal oDocFactura As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If oDocFactura.HeaderRow("IDTipoAsiento") = enumTipoAsientoFV.taClienteSinCobro Then Exit Sub

        Dim SumaImporteVencimiento As Double : Dim fSumaImporteVencimiento As New fImporte
        Dim SumaImporteVencimientoA As Double
        Dim SumaImporteVencimientoB As Double

        Dim SumaImporteRecFinanciero As Double : Dim fSumaImporteRecFinanciero As New fImporte
        Dim SumaImporteRecFinancieroA As Double
        Dim SumaImporteRecFinancieroB As Double
        '//Se borran todos los vencimientos existentes para la factura.
        For Each cobro As DataRow In oDocFactura.dtCobros.Rows
            '//Borramos los vencimientos que no provienen de Entregas a Cuenta.
            If Length(cobro("IDEntrega")) = 0 Then cobro.Delete()
        Next

        'Dim primero As Boolean = True
        Dim fImporteVencimiento As New fImporte : Dim fImporteRecFinanciero As New fImporte
        Dim Direccion As Integer = ProcessServer.ExecuteTask(Of DocumentoFacturaVenta, Integer)(AddressOf DireccionCobro, oDocFactura, services)
        Dim TCobro As Integer = ProcessServer.ExecuteTask(Of DocumentoFacturaVenta, Integer)(AddressOf TipoCobro, oDocFactura, services)

        Dim IDAlbaranes(-1) As Object
        Dim IDAlbaranANT As Integer
        For Each drLinea As DataRow In oDocFactura.dtLineas.Select("", "IDAlbaran")
            If IDAlbaranANT <> Nz(drLinea("IDAlbaran"), 0) Then
                IDAlbaranANT = drLinea("IDAlbaran")
                ReDim Preserve IDAlbaranes(UBound(IDAlbaranes) + 1)
                IDAlbaranes(UBound(IDAlbaranes)) = drLinea("IDAlbaran")
            End If
        Next
        If IDAlbaranes.Length = 0 Then Exit Sub
        Dim f As New Filter
        f.Add(New InListFilterItem("IDAlbaran", IDAlbaranes, FilterType.Numeric))
        f.Add(New BooleanFilterItem("SegunCliente", False))
        Dim dtAlbaranFormaPago As DataTable = New BE.DataEngine().Filter("vNegTPVAlbaranFormaPago", f)

        If Not dtAlbaranFormaPago Is Nothing AndAlso dtAlbaranFormaPago.Rows.Count > 0 Then
            Dim IDFormaPagoANT As String ', IDBancoPropioANT As String
            For Each drAlbaranFormaPago As DataRow In dtAlbaranFormaPago.Rows
                If IDFormaPagoANT <> drAlbaranFormaPago("IDFormaPago") Then 'And IDBancoPropioANT <> drAlbaranFormaPago("IDBancoPropio") Then
                    IDFormaPagoANT = drAlbaranFormaPago("IDFormaPago")
                    'IDBancoPropioANT = drAlbaranFormaPago("IDBancoPropio")

                    Dim drCobro As DataRow = oDocFactura.dtCobros.NewRow
                    drCobro("IDCobro") = AdminData.GetAutoNumeric
                    drCobro("IDFactura") = oDocFactura.HeaderRow("IDFactura")
                    drCobro("NFactura") = oDocFactura.HeaderRow("NFactura")
                    drCobro("IDCliente") = oDocFactura.HeaderRow("IDCliente")
                    drCobro("CContable") = oDocFactura.Cliente.CCCliente
                    drCobro("FechaVencimiento") = oDocFactura.HeaderRow("FechaFactura")
                    drCobro("FechaVencimientoFactura") = oDocFactura.HeaderRow("FechaFactura")
                    drCobro("IDDireccion") = Direccion
                    drCobro("Titulo") = Nz(oDocFactura.HeaderRow("RazonSocial"), oDocFactura.Cliente.RazonSocial)
                    drCobro("IDClienteBanco") = oDocFactura.HeaderRow("IDClienteBanco")
                    drCobro("IDMoneda") = oDocFactura.HeaderRow("IDMoneda")
                    drCobro("CambioA") = oDocFactura.HeaderRow("CambioA")
                    drCobro("CambioB") = oDocFactura.HeaderRow("CambioB")
                    drCobro("IDTipoCobro") = TCobro
                    drCobro("IDFormaPago") = drAlbaranFormaPago("IDFormaPago")
                    drCobro("IDBancoPropio") = drAlbaranFormaPago("IDBancoPropio")


                    If drCobro.Table.Columns.Contains("IDMandato") Then
                        drCobro("IDMandato") = oDocFactura.HeaderRow("IDMandato")
                    End If


                    Dim fCompute As New Filter
                    fCompute.Add(New StringFilterItem("IDFormaPago", drAlbaranFormaPago("IDFormaPago")))
                    fCompute.Add(New StringFilterItem("IDBancoPropio", drAlbaranFormaPago("IDBancoPropio")))
                    drCobro("ImpVencimiento") = dtAlbaranFormaPago.Compute("SUM(ImporteVencimiento)", fCompute.Compose(New AdoFilterComposer))
                    Dim ValAyB As New ValoresAyB(CDbl(Nz(drCobro("ImpVencimiento"), 0)), drCobro("IDMoneda") & String.Empty, Nz(drCobro("CambioA"), 1), Nz(drCobro("CambioB"), 1))
                    Dim ImportesAB As fImporte = ProcessServer.ExecuteTask(Of ValoresAyB, fImporte)(AddressOf NegocioGeneral.MantenimientoValoresImporteAyB, ValAyB, services)
                    drCobro("ImpVencimientoA") = ImportesAB.ImporteA
                    drCobro("ImpVencimientoB") = ImportesAB.ImporteB

                    fImporteVencimiento.Importe = drCobro("ImpVencimiento")
                    fImporteVencimiento.ImporteA = drCobro("ImpVencimientoA")
                    fImporteVencimiento.ImporteB = drCobro("ImpVencimientoB")

                    SumaImporteVencimiento = SumaImporteVencimiento + drCobro("ImpVencimiento")
                    SumaImporteVencimientoA = SumaImporteVencimientoA + drCobro("ImpVencimientoA")
                    SumaImporteVencimientoB = SumaImporteVencimientoB + drCobro("ImpVencimientoB")

                    'Comentamos esta linea para que se asigne automaticamente el TipoAsiente del cliente
                    ''Cambiado porque cuando viene de TPV todo se cobra en situación de contado si no son condiciones del cliente
                    'oDocFactura.HeaderRow("IdTipoAsiento") = enumTipoAsientoFV.taClienteBancoCobroC
                    Dim datEstSit As New DataAsignarEstadoSituacionFV(oDocFactura.HeaderRow("IdTipoAsiento"), drCobro)
                    ProcessServer.ExecuteTask(Of DataAsignarEstadoSituacionFV)(AddressOf AsignarEstadoSituacion, datEstSit, services)

                    oDocFactura.dtCobros.Rows.Add(drCobro.ItemArray)
                End If

            Next
        End If

        fSumaImporteVencimiento.Importe = SumaImporteVencimiento
        fSumaImporteVencimiento.ImporteA = SumaImporteVencimientoA
        fSumaImporteVencimiento.ImporteB = SumaImporteVencimientoB

        Dim fImporteFacturaFinal As New fImporte

        '///Ajuste de los vencimientos
        Dim datAjuste As New DataAjusteVencimientosFV(oDocFactura, fSumaImporteVencimiento, fSumaImporteRecFinanciero, fImporteFacturaFinal)
        ProcessServer.ExecuteTask(Of DataAjusteVencimientosFV)(AddressOf AjusteVencimientos, datAjuste, services)

        If oDocFactura.dtCobros.Rows.Count > 0 Then
            Dim FechaVto As Date = oDocFactura.dtCobros.Compute("MIN(FechaVencimiento)", Nothing)
            oDocFactura.HeaderRow("FechaVencimiento") = FechaVto
        End If
    End Sub
    <Task()> Public Shared Sub CobrosAutomaticos(ByVal oDocFactura As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If oDocFactura.HeaderRow("IDTipoAsiento") = enumTipoAsientoFV.taClienteSinCobro Then Exit Sub

        Dim SumaImporteVencimiento As Double : Dim fSumaImporteVencimiento As New fImporte
        Dim SumaImporteVencimientoA As Double
        Dim SumaImporteVencimientoB As Double

        Dim SumaImporteRecFinanciero As Double : Dim fSumaImporteRecFinanciero As New fImporte
        Dim SumaImporteRecFinancieroA As Double
        Dim SumaImporteRecFinancieroB As Double

        Dim fImporteEntregasACta As New fImporte
        '//Se borran todos los vencimientos existentes para la factura.
        For Each cobro As DataRow In oDocFactura.dtCobros.Rows
            '//Borramos los vencimientos que no provienen de Entregas a Cuenta.
            If Length(cobro("IDEntrega")) = 0 Then
                cobro.Delete()
            Else
                fImporteEntregasACta.Importe = fImporteEntregasACta.Importe + Nz(cobro("ImpVencimiento"), 0)
            End If
        Next

        Dim ValAyB As New ValoresAyB(fImporteEntregasACta.Importe, oDocFactura.IDMoneda, oDocFactura.CambioA, oDocFactura.CambioB)
        fImporteEntregasACta = ProcessServer.ExecuteTask(Of ValoresAyB, fImporte)(AddressOf NegocioGeneral.MantenimientoValoresImporteAyB, ValAyB, services)


        'Dim primero As Boolean = True
        Dim fImporteVencimiento As New fImporte : Dim fImporteRecFinanciero As New fImporte
        Dim Direccion As Integer = ProcessServer.ExecuteTask(Of DocumentoFacturaVenta, Integer)(AddressOf DireccionCobro, oDocFactura, services)
        Dim TCobro As Integer = ProcessServer.ExecuteTask(Of DocumentoFacturaVenta, Integer)(AddressOf TipoCobro, oDocFactura, services)
        Dim DiaPago As String = oDocFactura.HeaderRow("IDDiaPago") & String.Empty
        Dim datNuevosCobros As New DataNuevosCobros(oDocFactura, Direccion, TCobro)
        Dim fImporteFacturaFinal As fImporte = ProcessServer.ExecuteTask(Of DocumentoFacturaVenta, fImporte)(AddressOf ImporteFacturaFinal, oDocFactura, services)
        'fImporteFacturaFinal.Importe -= fImporteEntregasACta.Importe
        'fImporteFacturaFinal.ImporteA -= fImporteEntregasACta.ImporteA
        'fImporteFacturaFinal.ImporteB -= fImporteEntregasACta.ImporteB

        '//Se recorren las líneas de condicion de pago para crear un cobro por cada una de ellas.
        Dim primero As Boolean = True
        Dim condiciones As DataTable = New CondicionPagoLinea().Filter(New StringFilterItem("IdCondicionPago", oDocFactura.HeaderRow("IDCondicionPago")))
        For Each condicion As DataRow In condiciones.Rows
            Dim ImporteVencimiento As Double = 0
            Dim ImporteVencimientoA As Double = 0
            Dim ImporteVencimientoB As Double = 0

            Dim ImporteRecFinanciero As Double = 0
            Dim ImporteRecFinancieroA As Double = 0
            Dim ImporteRecFinancieroB As Double = 0

            '//Guardamos la fecha de vencimiento más lejana para la cabecera.
            Dim dataVto As New NegocioGeneral.dataCalculoFechaVencimiento(oDocFactura.HeaderRow("FechaFactura"), condicion("Periodo"), condicion("TipoPeriodo"), DiaPago, True, oDocFactura.HeaderRow("IdCliente"))
            ProcessServer.ExecuteTask(Of NegocioGeneral.dataCalculoFechaVencimiento)(AddressOf NegocioGeneral.CalcularFechaVencimiento, dataVto, services)
            Dim FechaVencimiento As Date = dataVto.FechaVencimiento
            If primero Then
                primero = False
                oDocFactura.HeaderRow("FechaVencimiento") = FechaVencimiento
            Else
                If oDocFactura.HeaderRow("FechaVencimiento") > FechaVencimiento Then
                    oDocFactura.HeaderRow("FechaVencimiento") = FechaVencimiento
                End If
            End If

            ImporteVencimiento = xRound(((oDocFactura.HeaderRow("ImpTotal") - oDocFactura.HeaderRow("ImpRetencion") - oDocFactura.HeaderRow("ImpRetencionGar") - fImporteEntregasACta.Importe) * condicion("porcentaje") / 100) + fImporteFacturaFinal.Importe, oDocFactura.Moneda.NDecimalesImporte)
            ImporteVencimientoA = xRound(((oDocFactura.HeaderRow("ImpTotalA") - oDocFactura.HeaderRow("ImpRetencionA") - oDocFactura.HeaderRow("ImpRetencionGarA") - fImporteEntregasACta.ImporteA) * condicion("porcentaje") / 100) + fImporteFacturaFinal.ImporteA, oDocFactura.MonedaA.NDecimalesImporte)
            ImporteVencimientoB = xRound(((oDocFactura.HeaderRow("ImpTotalB") - oDocFactura.HeaderRow("ImpRetencionB") - oDocFactura.HeaderRow("ImpRetencionGarB") - fImporteEntregasACta.ImporteB) * condicion("porcentaje") / 100) + fImporteFacturaFinal.ImporteB, oDocFactura.MonedaB.NDecimalesImporte)

            fImporteVencimiento.Importe = ImporteVencimiento
            fImporteVencimiento.ImporteA = ImporteVencimientoA
            fImporteVencimiento.ImporteB = ImporteVencimientoB

            SumaImporteVencimiento = SumaImporteVencimiento + ImporteVencimiento
            SumaImporteVencimientoA = SumaImporteVencimientoA + ImporteVencimientoA
            SumaImporteVencimientoB = SumaImporteVencimientoB + ImporteVencimientoB

            fSumaImporteVencimiento.Importe = SumaImporteVencimiento
            fSumaImporteVencimiento.ImporteA = SumaImporteVencimientoA
            fSumaImporteVencimiento.ImporteB = SumaImporteVencimientoB

            If oDocFactura.HeaderRow("RecFinan") > 0 Then
                ImporteRecFinanciero = xRound(oDocFactura.HeaderRow("ImpRecFinan") * condicion("porcentaje") / 100, oDocFactura.Moneda.NDecimalesImporte)
                ImporteRecFinancieroA = xRound(oDocFactura.HeaderRow("ImpRecFinanA") * condicion("porcentaje") / 100, oDocFactura.MonedaA.NDecimalesImporte)
                ImporteRecFinancieroB = xRound(oDocFactura.HeaderRow("ImpRecFinanB") * condicion("porcentaje") / 100, oDocFactura.MonedaB.NDecimalesImporte)

                fImporteRecFinanciero.Importe = ImporteRecFinanciero
                fImporteRecFinanciero.ImporteA = ImporteRecFinancieroA
                fImporteRecFinanciero.ImporteB = ImporteRecFinancieroB

                SumaImporteRecFinanciero = SumaImporteRecFinanciero + ImporteRecFinanciero
                SumaImporteRecFinancieroA = SumaImporteRecFinancieroA + ImporteRecFinancieroA
                SumaImporteRecFinancieroB = SumaImporteRecFinancieroB + ImporteRecFinancieroB

                fSumaImporteRecFinanciero.Importe = SumaImporteRecFinanciero
                fSumaImporteRecFinanciero.ImporteA = SumaImporteRecFinancieroA
                fSumaImporteRecFinanciero.ImporteB = SumaImporteRecFinancieroB

            End If

            If fImporteVencimiento.Importe <> 0 Then
                datNuevosCobros.FechaVencimiento = FechaVencimiento
                datNuevosCobros.ImporteVencimiento = fImporteVencimiento
                datNuevosCobros.ImporteRecFinanciero = fImporteRecFinanciero
                ProcessServer.ExecuteTask(Of DataNuevosCobros)(AddressOf NuevoCobro, datNuevosCobros, services)
            End If
        Next

        '///Generacion de pago por retencion de garantía
        ProcessServer.ExecuteTask(Of DataNuevosCobros)(AddressOf NuevoCobroRetencionGarantia, datNuevosCobros, services)

        '///Ajuste de los vencimientos
        Dim datAjuste As New DataAjusteVencimientosFV(oDocFactura, fSumaImporteVencimiento, fSumaImporteRecFinanciero, fImporteFacturaFinal, fImporteEntregasACta)
        ProcessServer.ExecuteTask(Of DataAjusteVencimientosFV)(AddressOf AjusteVencimientos, datAjuste, services)

        ''///Añadimos los cobros de compensación
        'Dim dtCobroCompensacionOS As DataTable = ProcessServer.ExecuteTask(Of DocumentoFacturaVenta, DataTable)(AddressOf NuevosCobrosCompensacionOSFacturasCerradasConFianza, oDocFactura, services)

    End Sub

    <Task()> Public Shared Function TipoCobro(ByVal oDocFactura As DocumentoFacturaVenta, ByVal services As ServiceProvider) As Integer
        Dim AppParams As ParametroFacturaVenta = services.GetService(Of ParametroFacturaVenta)()
        Dim intTipoCobro As Integer = AppParams.TipoCobroFacturaVenta
        If Length(oDocFactura.HeaderRow("IDContador")) > 0 Then
            '//Si es una factura B
            If Not oDocFactura.AIva Then
                intTipoCobro = AppParams.TipoCobroFacturaVentaB
            End If
        End If
        Return intTipoCobro
    End Function

    <Task()> Public Shared Function DireccionCobro(ByVal oDocFactura As DocumentoFacturaVenta, ByVal services As ServiceProvider) As Integer
        Dim Direccion As Integer
        Dim cd As New ClienteDireccion

        If Length(oDocFactura.HeaderRow("IDDireccion")) > 0 Then
            Dim StDatosDirec As New ClienteDireccion.DataDirecDe(oDocFactura.HeaderRow("IDDireccion"), enumcdTipoDireccion.cdDireccionGiro)
            If ProcessServer.ExecuteTask(Of ClienteDireccion.DataDirecDe, Boolean)(AddressOf ClienteDireccion.EsDireccionDe, StDatosDirec, services) Then
                Direccion = oDocFactura.HeaderRow("IDDireccion")
            Else
                Dim StDatosDirecEnv As New ClienteDireccion.DataDirecEnvio(oDocFactura.HeaderRow("IDCliente"), enumcdTipoDireccion.cdDireccionGiro)
                Dim dir As DataTable = ProcessServer.ExecuteTask(Of ClienteDireccion.DataDirecEnvio, DataTable)(AddressOf ClienteDireccion.ObtenerDireccionEnvio, StDatosDirecEnv, services)
                If Not IsNothing(dir) AndAlso dir.Rows.Count Then
                    Direccion = Nz(dir.Rows(0)("IDDireccion"), 0)
                End If
            End If
        Else
            Dim StDatosDirecEnv As New ClienteDireccion.DataDirecEnvio(oDocFactura.HeaderRow("IDCliente"), enumcdTipoDireccion.cdDireccionGiro)
            Dim dir As DataTable = ProcessServer.ExecuteTask(Of ClienteDireccion.DataDirecEnvio, DataTable)(AddressOf ClienteDireccion.ObtenerDireccionEnvio, StDatosDirecEnv, services)
            If Not IsNothing(dir) AndAlso dir.Rows.Count Then
                Direccion = Nz(dir.Rows(0)("IDDireccion"), 0)
            End If
        End If
        Return Direccion
    End Function

    <Task()> Public Shared Sub NuevoCobro(ByVal data As DataNuevosCobros, ByVal services As ServiceProvider)
        If data.ImporteRecFinanciero Is Nothing Then data.ImporteRecFinanciero = New fImporte
        Dim newrow As DataRow = data.Doc.dtCobros.NewRow
        newrow("IDCobro") = AdminData.GetAutoNumeric
        If data.DireccionCobro <> 0 Then newrow("IDDireccion") = data.DireccionCobro
        Dim AppParamsConta As ParametroContabilidadVenta = services.GetService(Of ParametroContabilidadVenta)()
        If AppParamsConta.Contabilidad Then
            If Len(data.Doc.Cliente.CCCliente) = 0 Then ApplicationService.GenerateError("La Cuenta Contable del Cliente es un dato obligatorio.")
            newrow("CContable") = data.Doc.Cliente.CCCliente
        End If
        newrow("IDFactura") = data.Doc.HeaderRow("IDFactura")
        newrow("IdTipoCobro") = data.TipoCobro
        newrow("IdCliente") = data.Doc.HeaderRow("IdCliente")
        newrow("IdClienteBanco") = data.Doc.HeaderRow("IdClienteBanco")
        newrow("IDBancoPropio") = data.Doc.HeaderRow("IDBancoPropio")
        newrow("FechaVencimientoFactura") = data.FechaVencimiento
        newrow("FechaVencimiento") = data.FechaVencimiento
        newrow("Titulo") = Nz(data.Doc.HeaderRow("RazonSocial"), data.Doc.Cliente.RazonSocial)
        newrow("NFactura") = data.Doc.HeaderRow("NFactura")
        newrow("IDFormaPago") = data.Doc.HeaderRow("IDFormaPago")
        newrow("IDMoneda") = data.Doc.HeaderRow("IDMoneda")
        newrow("CambioA") = data.Doc.HeaderRow("CambioA")
        newrow("CambioB") = data.Doc.HeaderRow("CambioB")
        newrow("IDProveedor") = data.Doc.HeaderRow("IDProveedor")
        newrow("IDOperario") = data.Doc.HeaderRow("IDOperario")
        newrow("ComunicadoGestorCobro") = data.Doc.HeaderRow("ComunicadoGestorCobro")
        newrow("FechaComunicacionGestorCobro") = data.Doc.HeaderRow("FechaComunicacionGestorCobro")

        newrow("ImpVencimiento") = data.ImporteVencimiento.Importe
        newrow("ImporteRemesaAnticipo") = data.ImporteVencimiento.Importe
        newrow("RecargoFinanciero") = data.ImporteRecFinanciero.Importe

        newrow("ImpVencimientoA") = data.ImporteVencimiento.ImporteA
        newrow("ImporteRemesaAnticipoA") = data.ImporteVencimiento.ImporteA
        newrow("RecargoFinancieroA") = data.ImporteRecFinanciero.ImporteA

        newrow("ImpVencimientoB") = data.ImporteVencimiento.ImporteB
        newrow("RecargoFinancieroB") = data.ImporteRecFinanciero.ImporteB
        newrow("ImporteRemesaAnticipoB") = data.ImporteVencimiento.ImporteB

        If Length(data.Doc.HeaderRow("IDObra")) <> 0 Then
            newrow("IDObra") = data.Doc.HeaderRow("IDObra")
        End If

        If newrow.Table.Columns.Contains("IDMandato") Then
            newrow("IDMandato") = data.Doc.HeaderRow("IDMandato")
        End If

        Dim datEstSit As New DataAsignarEstadoSituacionFV(data.Doc.HeaderRow("IDTipoAsiento"), newrow)
        ProcessServer.ExecuteTask(Of DataAsignarEstadoSituacionFV)(AddressOf AsignarEstadoSituacion, datEstSit, services)

        data.Doc.dtCobros.Rows.Add(newrow.ItemArray)
    End Sub

    <Task()> Public Shared Function NuevoCobroRetencionGarantia(ByVal data As DataNuevosCobros, ByVal services As ServiceProvider) As fImporte
        If Nz(data.Doc.HeaderRow("Retencion"), 0) <> 0 And Nz(data.Doc.HeaderRow("ImpRetencionGar"), 0) <> 0 Then
            Dim newrow As DataRow = data.Doc.dtCobros.NewRow
            newrow("IDCobro") = AdminData.GetAutoNumeric
            If data.DireccionCobro <> 0 Then newrow("IDDireccion") = data.DireccionCobro

            Dim AppParamsConta As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
            If AppParamsConta.Contabilidad Then
                If Length(data.Doc.Cliente.CCRetencion) > 0 Then
                    newrow("CContable") = data.Doc.Cliente.CCRetencion
                Else
                    If Len(data.Doc.Cliente.CCCliente) = 0 Then ApplicationService.GenerateError("La Cuenta Contable es un dato obligatorio.")
                    newrow("CContable") = data.Doc.Cliente.CCCliente
                End If
            End If

            newrow("IDFactura") = data.Doc.HeaderRow("IDFactura")
            Dim AppParams As ParametroFacturaVenta = services.GetService(Of ParametroFacturaVenta)()
            newrow("IdTipoCobro") = AppParams.TipoCobroRetencion
            newrow("IdCliente") = data.Doc.HeaderRow("IdCliente")
            newrow("IdClienteBanco") = data.Doc.HeaderRow("IdClienteBanco")
            newrow("IDBancoPropio") = data.Doc.HeaderRow("IDBancoPropio")
            If Length(data.Doc.HeaderRow("TipoRetencion")) = 0 Then data.Doc.HeaderRow("TipoRetencion") = enumTipoRetencion.troSobreBI
            If Length(data.Doc.HeaderRow("FechaRetencion")) = 0 Then data.Doc.HeaderRow("FechaRetencion") = Today
            newrow("FechaVencimientoFactura") = data.Doc.HeaderRow("FechaRetencion")
            newrow("FechaVencimiento") = data.Doc.HeaderRow("FechaRetencion")
            newrow("Titulo") = "RETENCION " & Nz(data.Doc.HeaderRow("RazonSocial"), data.Doc.Cliente.DescCliente)
            newrow("NFactura") = data.Doc.HeaderRow("NFactura")
            newrow("IDFormaPago") = data.Doc.HeaderRow("IDFormaPago")
            newrow("IDMoneda") = data.Doc.HeaderRow("IDMoneda")
            newrow("CambioA") = data.Doc.HeaderRow("CambioA")
            newrow("CambioB") = data.Doc.HeaderRow("CambioB")
            newrow("IDProveedor") = data.Doc.HeaderRow("IDProveedor")
            newrow("IDOperario") = data.Doc.HeaderRow("IDOperario")
            newrow("ComunicadoGestorCobro") = data.Doc.HeaderRow("ComunicadoGestorCobro")
            newrow("FechaComunicacionGestorCobro") = data.Doc.HeaderRow("FechaComunicacionGestorCobro")

            Dim ValAyB As New ValoresAyB(CDbl(data.Doc.HeaderRow("ImpRetencionGar")), data.Doc.IDMoneda, data.Doc.CambioA, data.Doc.CambioB)
            Dim fImporteVencimiento As fImporte = ProcessServer.ExecuteTask(Of ValoresAyB, fImporte)(AddressOf NegocioGeneral.MantenimientoValoresImporteAyB, ValAyB, services)

            newrow("ImpVencimiento") = fImporteVencimiento.Importe
            newrow("ImpVencimientoA") = fImporteVencimiento.ImporteA
            newrow("ImpVencimientoB") = fImporteVencimiento.ImporteB

            newrow("IDObra") = data.Doc.HeaderRow("IDObra")


            If newrow.Table.Columns.Contains("IDMandato") Then
                newrow("IDMandato") = data.Doc.HeaderRow("IDMandato")
            End If

            Dim datEstSit As New DataAsignarEstadoSituacionFV(data.Doc.HeaderRow("IDTipoAsiento"), newrow)
            ProcessServer.ExecuteTask(Of DataAsignarEstadoSituacionFV)(AddressOf AsignarEstadoSituacion, datEstSit, services)

            data.Doc.dtCobros.Rows.Add(newrow)
        End If

    End Function

    Public Class DataAsignarEstadoSituacionFV
        Public TipoAsiento As enumTipoAsientoFV
        Public NewRow As DataRow

        Public Sub New(ByVal TipoAsiento As enumTipoAsientoFV, ByVal NewRow As DataRow)
            Me.TipoAsiento = TipoAsiento
            Me.NewRow = NewRow
        End Sub
    End Class
    <Task()> Public Shared Sub AsignarEstadoSituacion(ByVal data As DataAsignarEstadoSituacionFV, ByVal services As ServiceProvider)
        Select Case data.TipoAsiento
            Case enumTipoAsientoFV.taBancoSinCobro
                data.NewRow("Contabilizado") = enumCobroContabilizado.CobroContabilizado
                data.NewRow("Situacion") = enumCobroSituacion.Cobrado
            Case enumTipoAsientoFV.taClienteConCobroNNyNC
                data.NewRow("Contabilizado") = enumCobroContabilizado.CobroNoContabilizado
                data.NewRow("Situacion") = enumCobroSituacion.NoNegociado
            Case enumTipoAsientoFV.taClienteConCobroCyNC
                data.NewRow("Contabilizado") = enumCobroContabilizado.CobroNoContabilizado
                data.NewRow("Situacion") = enumCobroSituacion.Cobrado
            Case enumTipoAsientoFV.taClienteSinCobro
                data.NewRow("Contabilizado") = enumCobroContabilizado.CobroContabilizado
                data.NewRow("Situacion") = enumCobroSituacion.Cobrado
        End Select
    End Sub


    Public Class DataAjusteVencimientosFV
        Public Doc As DocumentoFacturaVenta
        Public SumaImporteVencimiento As fImporte
        Public SumaImporteRecFinanciero As fImporte
        Public SumaEntregasACuenta As fImporte
        Public ImporteFacturaFinal As fImporte

        Public Sub New(ByVal Doc As DocumentoFacturaVenta, ByVal SumaImporteVencimiento As fImporte, ByVal SumaImporteRecFinanciero As fImporte, ByVal ImporteFacturaFinal As fImporte, Optional ByVal SumaEntregasACuenta As fImporte = Nothing)
            Me.Doc = Doc
            Me.SumaImporteVencimiento = SumaImporteVencimiento
            Me.SumaImporteRecFinanciero = SumaImporteRecFinanciero
            Me.SumaEntregasACuenta = SumaEntregasACuenta
            Me.ImporteFacturaFinal = ImporteFacturaFinal
        End Sub
    End Class
    <Task()> Public Shared Sub AjusteVencimientos(ByVal data As DataAjusteVencimientosFV, ByVal services As ServiceProvider)
        If data.SumaEntregasACuenta Is Nothing Then data.SumaEntregasACuenta = New fImporte
        Dim AddedRows As DataTable = data.Doc.dtCobros.GetChanges(DataRowState.Added)
        If Not IsNothing(AddedRows) Then
            If AddedRows.Rows.Count > 0 Then
                Dim VtoAAjustar As DataRow
                If data.Doc.HeaderRow("ImpRetencionGar") = 0 Then
                    VtoAAjustar = data.Doc.dtCobros.Rows(data.Doc.dtCobros.Rows.Count - 1)
                Else
                    VtoAAjustar = data.Doc.dtCobros.Rows(data.Doc.dtCobros.Rows.Count - 2)
                End If

                'Importe Vencimientos
                If (data.SumaImporteVencimiento.Importe - (data.Doc.HeaderRow("ImpTotal") - Nz(data.Doc.HeaderRow("ImpRetencion"), 0)) - Nz(data.Doc.HeaderRow("ImpRetencionGar"), 0)) <> 0 Then
                    VtoAAjustar("ImpVencimiento") = VtoAAjustar("ImpVencimiento") + data.Doc.HeaderRow("ImpTotal") - Nz(data.Doc.HeaderRow("ImpRetencion"), 0) - Nz(data.Doc.HeaderRow("ImpRetencionGar"), 0) - data.SumaImporteVencimiento.Importe + data.ImporteFacturaFinal.Importe - data.SumaEntregasACuenta.Importe
                End If
                If (data.SumaImporteVencimiento.ImporteA - (data.Doc.HeaderRow("ImpTotalA") - Nz(data.Doc.HeaderRow("ImpRetencionA"), 0)) - Nz(data.Doc.HeaderRow("ImpRetencionGarA"), 0)) <> 0 Then
                    VtoAAjustar("ImpVencimientoA") = VtoAAjustar("ImpVencimientoA") + data.Doc.HeaderRow("ImpTotalA") - Nz(data.Doc.HeaderRow("ImpRetencionA"), 0) - Nz(data.Doc.HeaderRow("ImpRetencionGarA"), 0) - data.SumaImporteVencimiento.ImporteA + data.ImporteFacturaFinal.ImporteA - data.SumaEntregasACuenta.ImporteA
                End If
                If (data.SumaImporteVencimiento.ImporteB - (data.Doc.HeaderRow("ImpTotalB") - Nz(data.Doc.HeaderRow("ImpRetencionB"), 0)) - Nz(data.Doc.HeaderRow("ImpRetencionGarB"), 0)) <> 0 Then
                    VtoAAjustar("ImpVencimientoB") = VtoAAjustar("ImpVencimientoB") + data.Doc.HeaderRow("ImpTotalB") - Nz(data.Doc.HeaderRow("ImpRetencionB"), 0) - Nz(data.Doc.HeaderRow("ImpRetencionGarB"), 0) - data.SumaImporteVencimiento.ImporteB + data.ImporteFacturaFinal.ImporteB - data.SumaEntregasACuenta.ImporteB
                End If

                VtoAAjustar("ImporteRemesaAnticipo") = VtoAAjustar("ImpVencimiento")
                VtoAAjustar("ImporteRemesaAnticipoA") = VtoAAjustar("ImpVencimientoA")
                VtoAAjustar("ImporteRemesaAnticipoB") = VtoAAjustar("ImpVencimientoB")


                'Recargo Financiero
                If data.Doc.HeaderRow("RecFinan") > 0 Then
                    If (data.SumaImporteRecFinanciero.Importe - data.Doc.HeaderRow("ImpRecFinan")) <> 0 Then
                        VtoAAjustar("RecargoFinanciero") = Nz(VtoAAjustar("RecargoFinanciero"), 0) + data.Doc.HeaderRow("ImpRecFinan") - data.SumaImporteRecFinanciero.Importe
                    End If
                    If (data.SumaImporteRecFinanciero.ImporteA - data.Doc.HeaderRow("ImpRecFinanA")) <> 0 Then
                        VtoAAjustar("RecargoFinancieroA") = Nz(VtoAAjustar("RecargoFinancieroA"), 0) + data.Doc.HeaderRow("ImpRecFinanA") - data.SumaImporteRecFinanciero.ImporteA
                    End If
                    If (data.SumaImporteRecFinanciero.ImporteB - data.Doc.HeaderRow("ImpRecFinanB")) <> 0 Then
                        VtoAAjustar("RecargoFinancieroB") = Nz(VtoAAjustar("RecargoFinancieroB"), 0) + data.Doc.HeaderRow("ImpRecFinanB") - data.SumaImporteRecFinanciero.ImporteB
                    End If
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Function ImporteFacturaFinal(ByVal oDocFactura As DocumentoFacturaVenta, ByVal services As ServiceProvider) As fImporte
        Dim fImporteFacturaFinal As New fImporte
        If oDocFactura.HeaderRow("TipoFactura") = enumfvcTipoFactura.fvcFinal Then
            Dim dv As DataView = oDocFactura.dtLineas.DefaultView
            dv.RowFilter = "Cantidad=0"
            If dv.Count > 0 Then
                dv.Item(0).Row("Importe") = Nz(dv.Item(0).Row("Precio"), 0)
                Dim ValAyB As New ValoresAyB(CDbl(Nz(dv.Item(0).Row("Precio"), 0)), oDocFactura.IDMoneda, oDocFactura.CambioA, oDocFactura.CambioB)
                fImporteFacturaFinal = ProcessServer.ExecuteTask(Of ValoresAyB, fImporte)(AddressOf NegocioGeneral.MantenimientoValoresImporteAyB, ValAyB, services)
            End If
        End If

        Return fImporteFacturaFinal
    End Function

    <Task()> Public Shared Function NuevosCobrosCompensacionOSFacturasCerradasConFianza(ByVal oDocFactura As DocumentoFacturaVenta, ByVal services As ServiceProvider) As DataTable
        '//Búsqueda de los cobros de compensacion de las O.S. de las facturas cerradas y con Fianza
        Dim ValuesCobro(-1) As Object
        Dim ValuesTrabajo(-1) As Object
        Dim dtCobroCompensacionOS As DataTable
        If Not IsNothing(oDocFactura.dtLineas) AndAlso oDocFactura.dtLineas.Rows.Count > 0 Then
            Dim strIDTrabajoAnt As String
            Dim fTrabajo As New Filter
            fTrabajo.Add(New IsNullFilterItem("IDTrabajo", False))
            Dim WhereNotNullTrabajo As String = fTrabajo.Compose(New AdoFilterComposer)
            For Each drLineas As DataRow In oDocFactura.dtLineas.Select(WhereNotNullTrabajo, "IDTrabajo")
                If (drLineas("IDTrabajo") <> strIDTrabajoAnt) Then
                    strIDTrabajoAnt = drLineas("IDTrabajo")

                    Dim f As New Filter
                    f.Add(New NumberFilterItem("IDTrabajo", FilterOperator.Equal, drLineas("IDTrabajo")))
                    f.Add(New NumberFilterItem("Fianza", FilterOperator.GreaterThan, 0))
                    f.Add(New NumberFilterItem("Estado", FilterOperator.Equal, enumotEstado.otTerminado))
                    f.Add(New BooleanFilterItem("FianzaContabilizada", FilterOperator.Equal, True))
                    f.Add(New BooleanFilterItem("FianzaCompensada", FilterOperator.Equal, False))

                    Dim dtOT As DataTable = New BE.DataEngine().Filter("tbObraTrabajo", f, "IDCobroCompensacion")
                    If Not IsNothing(dtOT) AndAlso dtOT.Rows.Count > 0 Then
                        If Length(dtOT.Rows(0)("IDCobroCompensacion")) > 0 Then
                            ReDim Preserve ValuesCobro(UBound(ValuesCobro) + 1)
                            ReDim Preserve ValuesTrabajo(UBound(ValuesTrabajo) + 1)
                            ValuesCobro(UBound(ValuesCobro)) = dtOT.Rows(0)("IDCobroCompensacion")
                            ValuesTrabajo(UBound(ValuesTrabajo)) = drLineas("IDTrabajo")
                        End If
                    End If
                End If
            Next
            If ValuesCobro.Length > 0 Then
                Dim f As New Filter
                f.Add(New InListFilterItem("IDCobro", ValuesCobro, FilterType.Numeric))
                dtCobroCompensacionOS = New Cobro().Filter(f)
            End If
        End If

        Dim strAgrup As String
        If Not IsNothing(dtCobroCompensacionOS) AndAlso dtCobroCompensacionOS.Rows.Count > 0 Then
            Dim dteFechaComp As Date = oDocFactura.dtCobros.Rows(0)("FechaVencimiento")
            For Each drCobroComp As DataRow In dtCobroCompensacionOS.Select
                drCobroComp("IDFactura") = oDocFactura.HeaderRow("IDFactura")
                drCobroComp("IdCliente") = oDocFactura.HeaderRow("IdCliente")
                drCobroComp("IdClienteBanco") = oDocFactura.HeaderRow("IdClienteBanco")
                drCobroComp("FechaVencimientoFactura") = dteFechaComp
                drCobroComp("FechaVencimiento") = dteFechaComp
                drCobroComp("NFactura") = oDocFactura.HeaderRow("NFactura")

                Dim datEstSit As New DataAsignarEstadoSituacionFV(oDocFactura.HeaderRow("IDTipoAsiento"), drCobroComp)
                ProcessServer.ExecuteTask(Of DataAsignarEstadoSituacionFV)(AddressOf AsignarEstadoSituacion, datEstSit, services)
            Next
            If Not oDocFactura.dtCobros Is Nothing AndAlso oDocFactura.dtCobros.Rows.Count > 0 Then
                Dim drUltimoCobro As DataRow = oDocFactura.dtCobros.Rows(oDocFactura.dtCobros.Rows.Count - 1)
                Dim Cobro As New Cobro
                Dim dt As DataTable = Cobro.AddNewForm
                For Each dc As DataColumn In drUltimoCobro.Table.Columns
                    Select Case dc.ColumnName
                        Case "IDCobro"
                            dt.Rows(0)(dc.ColumnName) = AdminData.GetAutoNumeric
                        Case "IDFormaPago"
                            Dim AppParams As Parametro = services.GetService(GetType(Parametro))
                            dt.Rows(0)(dc.ColumnName) = AppParams.FormaPagoEnEfectivo
                        Case "CContable"
                            Dim AppParamsConta As ParametroContabilidadVenta = services.GetService(GetType(ParametroContabilidadVenta))
                            If AppParamsConta.Contabilidad AndAlso Length(drUltimoCobro("IdCliente")) > 0 Then
                                Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(GetType(EntityInfoCache(Of ClienteInfo)))
                                Dim ClteInfo As ClienteInfo = Clientes.GetEntity(drUltimoCobro("IdCliente"))
                                If Not IsNothing(ClteInfo) Then
                                    If Length(ClteInfo.CCCliente) > 0 Then
                                        dt.Rows(0)(dc.ColumnName) = ClteInfo.CCCliente
                                    End If
                                End If
                            End If
                        Case Else
                            dt.Rows(0)(dc.ColumnName) = drUltimoCobro(dc.ColumnName)
                    End Select
                Next

                Dim DblImporte As Double = drUltimoCobro("ImpVencimiento")
                DblImporte = DblImporte + dtCobroCompensacionOS.Compute("SUM(ImpVencimiento)", Nothing)

                Dim ValAyB As New ValoresAyB(DblImporte, oDocFactura.IDMoneda, oDocFactura.CambioA, oDocFactura.CambioB)
                Dim fImporteVtos As fImporte = ProcessServer.ExecuteTask(Of ValoresAyB, fImporte)(AddressOf NegocioGeneral.MantenimientoValoresImporteAyB, ValAyB, services)
                dt.Rows(0)("ImpVencimiento") = fImporteVtos.Importe
                dt.Rows(0)("ImpVencimientoA") = fImporteVtos.ImporteA
                dt.Rows(0)("ImpVencimientoB") = fImporteVtos.ImporteB

                If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
                    If Length(strAgrup) > 0 Then strAgrup = strAgrup & "X"
                    strAgrup = strAgrup & dt.Rows(0)("IDCobro")
                    strAgrup = strAgrup & "," & drUltimoCobro("IDCobro")
                    For Each drCobroComp As DataRow In dtCobroCompensacionOS.Rows
                        strAgrup = strAgrup & "," & drCobroComp("IDCobro")
                    Next

                    If ValuesTrabajo.Length > 0 Then
                        Dim ObraTrabajo As BusinessHelper = BusinessHelper.CreateBusinessObject(AdminData.GetEntityInfo("ObraTrabajo"))
                        Dim f As New Filter
                        f.Add(New InListFilterItem("IdTrabajo", ValuesTrabajo, FilterType.Numeric))
                        Dim dtOTs As DataTable = ObraTrabajo.Filter(f)
                        If Not IsNothing(dtOTs) AndAlso dtOTs.Rows.Count > 0 Then
                            For Each drOT As DataRow In dtOTs.Select
                                drOT("FianzaCompensada") = True
                            Next
                            ObraTrabajo.Update(dtOTs)
                        End If
                    End If
                End If
            End If
            Cobro.UpdateTable(dtCobroCompensacionOS)
        End If
    End Function

#End Region
#Region "Validar datos"

    '<Task()> Public Shared Sub ValidarDocumento(ByVal Doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)
    '    Dim FVC As New FacturaVentaCabecera
    '    FVC.Validate(Doc.HeaderRow.Table)

    '    Dim FVL As New FacturaVentaLinea
    '    FVL.Validate(Doc.dtLineas)

    '    Dim FVR As New FacturaVentaRepresentante
    '    FVR.Validate(Doc.dtVentaRepresentante)
    'End Sub


    '<Task()> Public Shared Sub ValidarNumeroFactura(ByVal data As DataRow, ByVal services As ServiceProvider)
    '    If data.RowState = DataRowState.Added Then
    '        Dim f As New Filter
    '        f.Add(New StringFilterItem("NFactura", data("NFactura")))
    '        If Length(data("IDContador")) > 0 Then
    '            f.Add(New StringFilterItem("IDContador", data("IDContador")))
    '        Else
    '            f.Add(New IsNullFilterItem("IDContador", True))
    '        End If
    '        'Dim AppParamsConta As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
    '        'If AppParamsConta.Contabilidad Then f.Add(New StringFilterItem("IDEjercicio", data("IDEjercicio")))
    '        f.Add(New StringFilterItem("YEAR(FechaFactura)", Year(data("FechaFactura"))))

    '        Dim dtFVC As DataTable = New FacturaVentaCabecera().Filter(f)
    '        If Not dtFVC Is Nothing AndAlso dtFVC.Rows.Count > 0 Then
    '            'If AppParamsConta.Contabilidad Then
    '            '    ApplicationService.GenerateError("La Factura {0} ya existe para el Ejercicio {1}.", Quoted(data("NFactura")), Quoted(data("IDEjercicio")))
    '            'Else
    '            '    ApplicationService.GenerateError("La Factura {0} ya existe.", Quoted(data("NFactura")))
    '            'End If
    '            ApplicationService.GenerateError("La Factura {0} ya existe para el año {1}.", Quoted(data("NFactura")), Quoted(Year(data("FechaFactura"))))
    '        End If
    '    End If
    'End Sub

    <Task()> Public Shared Sub ValidarFacturaContabilizada(ByVal DocHeaderRow As DataRow, ByVal services As ServiceProvider)
        If DocHeaderRow("Estado") = enumfvcEstado.fvcContabilizado Then
            If DocHeaderRow.RowState = DataRowState.Modified Then
                If Nz(DocHeaderRow("CifCliente")) <> Nz(DocHeaderRow("CifCliente", DataRowVersion.Original)) OrElse _
                    Nz(DocHeaderRow("Direccion")) <> Nz(DocHeaderRow("Direccion", DataRowVersion.Original)) OrElse _
                    Nz(DocHeaderRow("CodPostal")) <> Nz(DocHeaderRow("CodPostal", DataRowVersion.Original)) OrElse _
                    Nz(DocHeaderRow("Poblacion")) <> Nz(DocHeaderRow("Poblacion", DataRowVersion.Original)) OrElse _
                    Nz(DocHeaderRow("Provincia")) <> Nz(DocHeaderRow("Provincia", DataRowVersion.Original)) OrElse _
                    Nz(DocHeaderRow("IDPais")) <> Nz(DocHeaderRow("IDPais", DataRowVersion.Original)) OrElse _
                    Nz(DocHeaderRow("FechaParaDeclaracion")) <> Nz(DocHeaderRow("FechaParaDeclaracion", DataRowVersion.Original)) OrElse _
                    Nz(DocHeaderRow("Enviar347")) <> Nz(DocHeaderRow("Enviar347", DataRowVersion.Original)) OrElse _
                    Nz(DocHeaderRow("Enviar349")) <> Nz(DocHeaderRow("Enviar349", DataRowVersion.Original)) OrElse _
                    Nz(DocHeaderRow("Servicios349")) <> Nz(DocHeaderRow("Servicios349", DataRowVersion.Original)) OrElse _
                    Nz(DocHeaderRow("OpeTriangular")) <> Nz(DocHeaderRow("OpeTriangular", DataRowVersion.Original)) OrElse _
                    Nz(DocHeaderRow("ClaveOperacion")) <> Nz(DocHeaderRow("ClaveOperacion", DataRowVersion.Original)) OrElse _
                    Nz(DocHeaderRow("IDFacturaRectificada")) <> Nz(DocHeaderRow("IDFacturaRectificada", DataRowVersion.Original)) OrElse _
                    Nz(DocHeaderRow("TipoOperIntra")) <> Nz(DocHeaderRow("TipoOperIntra", DataRowVersion.Original)) OrElse _
                    Nz(DocHeaderRow("EnviadaEntidadAseguradora")) <> Nz(DocHeaderRow("EnviadaEntidadAseguradora", DataRowVersion.Original)) Then
                Else
                    If New Parametro().Contabilidad Then
                        ApplicationService.GenerateError("La Factura está Contabilizada.")
                    Else : ApplicationService.GenerateError("La Factura está Bloqueda y generados los vencimientos (o efectos)")
                    End If
                End If
            Else
                If New Parametro().Contabilidad Then
                    ApplicationService.GenerateError("La Factura está Contabilizada.")
                Else : ApplicationService.GenerateError("La Factura está Bloqueda y generados los vencimientos (o efectos)")
                End If
            End If
        End If
    End Sub


    <Task()> Public Shared Sub ValidarFacturaDeclarada(ByVal DocHeaderRow As DataRow, ByVal services As ServiceProvider)
        If Length(DocHeaderRow("AñoDeclaracionIva")) > 0 AndAlso Length(DocHeaderRow("NDeclaracionIva")) > 0 Then
            ApplicationService.GenerateError("La Factura está Declarada.")
        End If
    End Sub
    <Task()> Public Shared Sub ValidarFacturaArqueoCaja(ByVal DocHeaderRow As DataRow, ByVal services As ServiceProvider)
        If Length(DocHeaderRow("Arqueo")) > 0 AndAlso DocHeaderRow("Arqueo") Then
            ApplicationService.GenerateError("No se permite eliminar una factura Arqueada.")
        End If
    End Sub

#End Region

#Region " Grabar Documento "

    <Task()> Public Shared Sub GrabarDocumento(ByVal Doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        'AdminData.BeginTx()
        'ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ValidarDocumento, Doc, services)
        'AdminData.SetData(Doc.HeaderRow.Table)
        'AdminData.SetData(Doc.dtLineas)
        'AdminData.SetData(Doc.dtFVBI)
        'AdminData.SetData(Doc.dtCobros)
        'AdminData.SetData(Doc.dtAnalitica)
        'AdminData.SetData(Doc.dtVentaRepresentante)
        ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf Business.General.Comunes.UpdateDocument, Doc, services)

        'AdminData.CommitTx(True)
    End Sub

#End Region

#Region "Métodos Borrado"

    <Task()> Public Shared Sub ActualizarEntregasACuenta(ByVal DocHeaderRow As DataRow, ByVal services As ServiceProvider)
        Dim StDatos As New EntregasACuenta.DatosElimRestricEnt
        StDatos.IDFactura = DocHeaderRow("IDFactura")
        StDatos.Circuito = Circuito.Ventas
        ProcessServer.ExecuteTask(Of EntregasACuenta.DatosElimRestricEnt)(AddressOf EntregasACuenta.EliminarRestriccionesDeleteEntregaCuenta, StDatos, services)
    End Sub

    <Task()> Public Shared Sub ActualizarPromociones(ByVal DocHeaderRow As DataRow, ByVal services As ServiceProvider)
        If DocHeaderRow("TipoFactura") = enumfvcTipoFactura.fvcFinal Then
            If DocHeaderRow("IDFactura") > 0 Then
                Dim Obra As BusinessHelper
                Obra = BusinessHelper.CreateBusinessObject(AdminData.GetEntityInfo("ObraPromoLocalVencimiento"))

                Dim dtVtos As DataTable = Obra.Filter(New FilterItem("IDFactura", DocHeaderRow("IDFactura")))

                If Not dtVtos Is Nothing AndAlso dtVtos.Rows.Count > 0 Then
                    For Each drVtos As DataRow In dtVtos.Rows
                        drVtos("Facturado") = False
                        If Length(drVtos("IdPromoVencimiento")) > 0 Then
                            Dim dtCadencia As DataTable = New BE.DataEngine().Filter("tbObraPromoVencimiento", New FilterItem("IdPromoVencimiento", drVtos("IdPromoVencimiento")))
                            If Not dtCadencia Is Nothing AndAlso dtCadencia.Rows.Count > 0 Then
                                drVtos("TipoFactura") = dtCadencia.Rows(0)("TipoFActura")
                            End If
                        End If
                        drVtos("IDFactura") = System.DBNull.Value
                        drVtos("NFactura") = System.DBNull.Value
                    Next
                    BusinessHelper.UpdateTable(dtVtos)
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub EliminarCobrosCompensacionesOServicio(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim dtVtosFactura As DataTable = New Cobro().Filter(New NumberFilterItem("IDFactura", data("IDFactura")))
        If Not IsNothing(dtVtosFactura) AndAlso dtVtosFactura.Rows.Count > 0 Then
            Dim f As New Filter(FilterUnionOperator.Or)
            For Each drCobro As DataRow In dtVtosFactura.Rows
                f.Add(New NumberFilterItem("IDCobroCompensacion", drCobro("IDCobro")))
            Next

            Dim ot As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraTrabajo")
            Dim dtTrabajos As DataTable = ot.Filter(f)
            If Not IsNothing(dtTrabajos) AndAlso dtTrabajos.Rows.Count > 0 Then
                f.Clear()
                For Each drTrabajo As DataRow In dtTrabajos.Rows
                    f.Add(New NumberFilterItem("IDCobro", drTrabajo("IDCobroCompensacion")))
                    drTrabajo("FianzaCompensada") = False
                Next

                For Each drCobro As DataRow In dtVtosFactura.Select(f.Compose(New AdoFilterComposer))
                    drCobro("IDFactura") = System.DBNull.Value
                    drCobro("NFactura") = System.DBNull.Value
                Next
                AdminData.SetData(dtVtosFactura)
                AdminData.SetData(dtTrabajos)
            End If
        End If
    End Sub

#End Region

#Region " Actualización de Albaranes "

    <Task()> Public Shared Sub ActualizarAlbaran(ByVal DocFra As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If DocFra.HeaderRow("Estado") = enumfvcEstado.fvcNoContabilizado Then
            ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ActualizarQFacturadaAlbaran, DocFra, services)
            ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ActualizarImportesAlbaran, DocFra, services)
            ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf GrabarAlbaranes, DocFra, services)
        End If
    End Sub


    <Task()> Public Shared Sub ActualizarQFacturadaAlbaran(ByVal DocFra As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If DocFra Is Nothing Then Exit Sub
        Dim f As New Filter
        f.Add(New IsNullFilterItem("IDAlbaran", False))
        f.Add(New IsNullFilterItem("IDLineaAlbaran", False))
        Dim WhereNotNullAlbaran As String = f.Compose(New AdoFilterComposer)
        For Each lineaFactura As DataRow In DocFra.dtLineas.Select(WhereNotNullAlbaran, "IDAlbaran,IDLineaAlbaran")
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ActualizarQFacturadaLineaAlbaran, lineaFactura, services)
        Next
    End Sub

    '<Task()> Public Shared Sub ActualizarAlbaranEnProceso(ByVal DocFra As DocumentoFacturaVenta, ByVal services As ServiceProvider)
    '    If DocFra Is Nothing Then Exit Sub
    '    Dim f As New Filter
    '    f.Add(New IsNullFilterItem("IDAlbaran", False))
    '    f.Add(New IsNullFilterItem("IDLineaAlbaran", False))
    '    Dim WhereNotNullAlbaran As String = f.Compose(New AdoFilterComposer)
    '    For Each lineaFactura As DataRow In DocFra.dtLineas.Select(WhereNotNullAlbaran, "IDAlbaran,IDLineaAlbaran")
    '        ProcessServer.ExecuteTask(Of DataRow)(AddressOf ActualizarQFacturadaLineaAlbaran, lineaFactura, services)
    '        ProcessServer.ExecuteTask(Of DataRow)(AddressOf ActualizarTipoDocumentoAlbaranTickectFacturado, lineaFactura, services)
    '    Next
    '    ProcessServer.ExecuteTask(Of Object)(AddressOf GrabarAlbaranes, Nothing, services)
    'End Sub

    <Task()> Public Shared Sub ActualizarAlbaranEnProceso(ByVal DocFra As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If DocFra Is Nothing OrElse DocFra.dtLineas Is Nothing OrElse DocFra.dtLineas.Rows.Count = 0 Then Exit Sub

        Dim IDAlbaranesFactura As List(Of Object) = (From c In DocFra.dtLineas _
                                                      Where c.RowState <> DataRowState.Deleted AndAlso Not c.IsNull("IDAlbaran") AndAlso Not c.IsNull("IDLineaAlbaran") _
                                                      Select c("IDAlbaran") Distinct).ToList
        If Not IDAlbaranesFactura Is Nothing AndAlso IDAlbaranesFactura.Count > 0 Then
            For Each IDAlbaran As Integer In IDAlbaranesFactura
                Dim LineasFraAlbaran As List(Of DataRow) = (From c In DocFra.dtLineas _
                                                               Where (c.RowState = DataRowState.Added OrElse (c.RowState = DataRowState.Modified AndAlso c("Cantidad") <> c("cantidad", DataRowVersion.Original))) AndAlso _
                                                                      Not c.IsNull("IDAlbaran") AndAlso c("IDAlbaran") = IDAlbaran _
                                                               Select c).ToList
                If Not LineasFraAlbaran Is Nothing AndAlso LineasFraAlbaran.Count > 0 Then
                    For Each lineaFactura As DataRow In LineasFraAlbaran
                        ProcessServer.ExecuteTask(Of DataRow)(AddressOf ValidarAlbaranFacturado, lineaFactura, services)
                        ProcessServer.ExecuteTask(Of DataRow)(AddressOf ActualizarQFacturadaLineaAlbaran, lineaFactura, services)
                        ProcessServer.ExecuteTask(Of DataRow)(AddressOf ActualizarTipoDocumentoAlbaranTickectFacturado, lineaFactura, services)
                    Next
                End If
                ProcessServer.ExecuteTask(Of Object)(AddressOf GrabarAlbaranes, Nothing, services)
            Next
        End If
    End Sub

    <Task()> Public Shared Sub ValidarAlbaranFacturado(ByVal lineaFactura As DataRow, ByVal services As ServiceProvider)
        If Length(lineaFactura("IDAlbaran")) > 0 AndAlso Length(lineaFactura("IDLineaAlbaran")) > 0 Then
            If lineaFactura.RowState <> DataRowState.Modified OrElse lineaFactura("cantidad") <> lineaFactura("cantidad", DataRowVersion.Original) Then
                Dim Albaranes As DocumentInfoCache(Of DocumentoAlbaranVenta) = services.GetService(Of DocumentInfoCache(Of DocumentoAlbaranVenta))()
                Dim DocAlb As DocumentoAlbaranVenta = Albaranes.GetDocument(lineaFactura("IDAlbaran"))

                Dim LineaActualizar As List(Of DataRow) = (From c In DocAlb.dtLineas Where c("IDLineaAlbaran") = lineaFactura("IDLineaAlbaran") Select c).ToList
                If Not LineaActualizar Is Nothing AndAlso LineaActualizar.Count > 0 Then
                    If LineaActualizar(0)("EstadoFactura") = enumavlEstadoFactura.avlFacturado Then
                        ApplicationService.GenerateError("La linea del Albarán {0} ya está facturado. Linea correspondiente al articulo {1}.", Quoted(DocAlb.HeaderRow("NAlbaran")), Quoted(LineaActualizar(0)("IdArticulo")))
                    End If
                End If
            End If
        End If
    End Sub
    <Task()> Public Shared Sub ActualizarQFacturadaLineaAlbaran(ByVal lineaFactura As DataRow, ByVal services As ServiceProvider)
        If Length(lineaFactura("IDAlbaran")) > 0 AndAlso Length(lineaFactura("IDLineaAlbaran")) > 0 Then
            If lineaFactura.RowState <> DataRowState.Modified OrElse lineaFactura("cantidad") <> lineaFactura("cantidad", DataRowVersion.Original) Then
                Dim Albaranes As DocumentInfoCache(Of DocumentoAlbaranVenta) = services.GetService(Of DocumentInfoCache(Of DocumentoAlbaranVenta))()
                Dim DocAlb As DocumentoAlbaranVenta = Albaranes.GetDocument(lineaFactura("IDAlbaran"))

                Dim OriginalQFacturada As Double
                Dim ProposedQFacturada As Double = Nz(lineaFactura("cantidad"), 0)
                If lineaFactura.RowState = DataRowState.Modified Then
                    OriginalQFacturada = lineaFactura("cantidad", DataRowVersion.Original)
                End If
                DocAlb.SetQFacturada(lineaFactura("IDLineaAlbaran"), ProposedQFacturada - OriginalQFacturada, services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarTipoDocumentoAlbaranTickectFacturado(ByVal lineaFactura As DataRow, ByVal services As ServiceProvider)
        If Length(lineaFactura("IDAlbaran")) > 0 AndAlso Length(lineaFactura("IDLineaAlbaran")) > 0 Then
            If lineaFactura.RowState <> DataRowState.Modified Then
                Dim Albaranes As DocumentInfoCache(Of DocumentoAlbaranVenta) = services.GetService(Of DocumentInfoCache(Of DocumentoAlbaranVenta))()
                Dim DocAlb As DocumentoAlbaranVenta = Albaranes.GetDocument(lineaFactura("IDAlbaran"))
                DocAlb.SetTipoDocumentoFactura()
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarImportesAlbaran(ByVal DocFra As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If DocFra Is Nothing Then Exit Sub
        Dim f As New Filter
        f.Add(New IsNullFilterItem("IDAlbaran", False))
        f.Add(New IsNullFilterItem("IDLineaAlbaran", False))
        Dim WhereNotNullAlbaran As String = f.Compose(New AdoFilterComposer)
        For Each lineaFactura As DataRow In DocFra.dtLineas.Select(WhereNotNullAlbaran, "IDAlbaran,IDLineaAlbaran")
            ' ProcessServer.ExecuteTask(Of DataRow)(AddressOf ActualizarImportesLineaAlbaran, lineaFactura, services)

            'End Sub

            '<Task()> Public Shared Sub ActualizarImportesLineaAlbaran(ByVal lineaFactura As DataRow, ByVal services As ServiceProvider)
            Dim AlbPeriodoCerrado As New Dictionary(Of Integer, Boolean)

            If lineaFactura("Precio") <> lineaFactura("Precio", DataRowVersion.Original) Or lineaFactura("PrecioA") <> lineaFactura("PrecioA", DataRowVersion.Original) Or lineaFactura("PrecioB") <> lineaFactura("PrecioB", DataRowVersion.Original) Or _
            lineaFactura("Dto1") <> lineaFactura("Dto1", DataRowVersion.Original) Or lineaFactura("Dto2") <> lineaFactura("Dto2", DataRowVersion.Original) Or lineaFactura("Dto3") <> lineaFactura("Dto3", DataRowVersion.Original) Or _
            lineaFactura("Dto") <> lineaFactura("Dto", DataRowVersion.Original) Or lineaFactura("DtoProntoPago") <> lineaFactura("DtoProntoPago", DataRowVersion.Original) Then

                Dim AVL As New AlbaranVentaLinea
                Dim context As New BusinessData
                Dim strTipoAlbaran As String

                Dim dtAVC As DataTable = New AlbaranVentaCabecera().SelOnPrimaryKey(lineaFactura("IdAlbaran"))
                If Not dtAVC Is Nothing AndAlso dtAVC.Rows.Count > 0 Then
                    strTipoAlbaran = dtAVC.Rows(0)("IDTipoAlbaran") & String.Empty
                End If
                Dim AppParamsAlb As ParametroAlbaranVenta = services.GetService(Of ParametroAlbaranVenta)()
                If Length(strTipoAlbaran) > 0 AndAlso (strTipoAlbaran <> AppParamsAlb.TipoAlbaranRetornoAlquiler AndAlso strTipoAlbaran <> AppParamsAlb.TipoAlbaranDeDeposito) Then
                    Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
                    'Dim Factura As DocumentInfoCache(Of DocumentoFacturaVenta) = services.GetService(Of DocumentInfoCache(Of DocumentoFacturaVenta))()
                    'Dim DocFra As DocumentoFacturaVenta = Factura.GetDocument(lineaFactura("IDFactura"))
                    Dim Albaranes As DocumentInfoCache(Of DocumentoAlbaranVenta) = services.GetService(Of DocumentInfoCache(Of DocumentoAlbaranVenta))()
                    Dim DocAlb As DocumentoAlbaranVenta = Albaranes.GetDocument(lineaFactura("IDAlbaran"))

                    DocAlb.IDMoneda = DocFra.IDMoneda
                    DocAlb.CambioA = DocFra.CambioA
                    DocAlb.CambioB = DocFra.CambioB
                    context("IDMoneda") = DocAlb.IDMoneda
                    context("CambioA") = DocAlb.CambioA
                    context("CambioB") = DocAlb.CambioB

                    Dim blnAlbPeriodoCerrado As Boolean = False
                    If AlbPeriodoCerrado.ContainsKey(lineaFactura("IDAlbaran")) Then
                        blnAlbPeriodoCerrado = AlbPeriodoCerrado(lineaFactura("IDAlbaran"))
                    Else
                        blnAlbPeriodoCerrado = ProcessServer.ExecuteTask(Of Date, Boolean)(AddressOf ProcesoComunes.AlbaranEnPeriodoCerrado, DocAlb.HeaderRow("FechaAlbaran"), services)
                        AlbPeriodoCerrado(lineaFactura("IDAlbaran")) = blnAlbPeriodoCerrado
                    End If

                    If AppParams.ActualizarPrecioAlbaranPeriodoCerrado OrElse Not blnAlbPeriodoCerrado Then
                        Dim fLineaAlbaran As New Filter
                        fLineaAlbaran.Add(New NumberFilterItem("IDLineaAlbaran", lineaFactura("IDLineaAlbaran")))
                        Dim WhereLineaAlbaran As String = fLineaAlbaran.Compose(New AdoFilterComposer)
                        For Each lineaAlbaran As DataRow In DocAlb.dtLineas.Select(WhereLineaAlbaran)

                            Dim LineaAlb As IPropertyAccessor = New DataRowPropertyAccessor(lineaAlbaran)
                            If LineaAlb("Precio") <> lineaFactura("Precio") Then
                                LineaAlb("Precio") = lineaFactura("Precio")
                                LineaAlb = AVL.ApplyBusinessRule("Precio", lineaFactura("Precio"), LineaAlb, context)
                            End If
                            If LineaAlb("PrecioA") <> lineaFactura("PrecioA") Then
                                LineaAlb("PrecioA") = lineaFactura("PrecioA")
                                LineaAlb = AVL.ApplyBusinessRule("PrecioA", lineaFactura("PrecioA"), LineaAlb, context)
                            End If
                            If LineaAlb("PrecioB") <> lineaFactura("PrecioB") Then
                                LineaAlb("PrecioB") = lineaFactura("PrecioB")
                                LineaAlb = AVL.ApplyBusinessRule("PrecioB", lineaFactura("PrecioB"), LineaAlb, context)
                            End If
                            If LineaAlb("Dto1") <> lineaFactura("Dto1") Then
                                LineaAlb("Dto1") = lineaFactura("Dto1")
                                LineaAlb = AVL.ApplyBusinessRule("Dto1", lineaFactura("Dto1"), LineaAlb, context)
                            End If
                            If LineaAlb("Dto2") <> lineaFactura("Dto2") Then
                                LineaAlb("Dto2") = lineaFactura("Dto2")
                                LineaAlb = AVL.ApplyBusinessRule("Dto2", lineaFactura("Dto2"), LineaAlb, context)
                            End If
                            If LineaAlb("Dto3") <> lineaFactura("Dto3") Then
                                LineaAlb("Dto3") = lineaFactura("Dto3")
                                LineaAlb = AVL.ApplyBusinessRule("Dto3", lineaFactura("Dto3"), LineaAlb, context)
                            End If
                            If LineaAlb("Dto") <> lineaFactura("Dto") Then
                                LineaAlb("Dto") = lineaFactura("Dto")
                                LineaAlb = AVL.ApplyBusinessRule("Dto", lineaFactura("Dto"), LineaAlb, context)
                            End If
                            If LineaAlb("DtoProntoPago") <> lineaFactura("DtoProntoPago") Then
                                LineaAlb("DtoProntoPago") = lineaFactura("DtoProntoPago")
                                LineaAlb = AVL.ApplyBusinessRule("DtoProntoPago", lineaFactura("DtoProntoPago"), LineaAlb, context)
                            End If

                            Dim ctx As New DataDocRow(DocAlb, lineaAlbaran)
                            Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of DataDocRow, StockUpdateData)(AddressOf ProcesoAlbaranVenta.CorregirMovimiento, ctx, services)
                        Next
                    End If
                    ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoComercial.CalcularRepresentantes, DocAlb, services)
                    ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVentaPedidos.AnaliticaLineasNSerie, DocAlb, services)
                    ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf NegocioGeneral.CalcularAnalitica, DocAlb, services)
                    ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.CalcularBasesImponibles, DocAlb, services)
                    ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoComunes.TotalDocumento, DocAlb, services)
                End If
            End If
        Next
    End Sub


    <Task()> Public Shared Sub GrabarAlbaranes(ByVal data As Object, ByVal services As ServiceProvider)
        ' AdminData.BeginTx()
        Dim Albaranes As DocumentInfoCache(Of DocumentoAlbaranVenta) = services.GetService(Of DocumentInfoCache(Of DocumentoAlbaranVenta))()
        For Each key As Integer In Albaranes.Keys
            Dim DocAlb As DocumentoAlbaranVenta = Albaranes.GetDocument(key)
            DocAlb.SetData()
            DocAlb.ClearDoc()
        Next
        Albaranes.Clear() '//Hemos actualizado los albaranes, para que los podamos tener actualizados, debemos limpiar la lista
    End Sub


#End Region

#Region "Promociones y Regalos"

    '<Task()> Public Shared Sub TratarPromocionesLineas(ByVal Doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)
    '    Dim pl As New PromocionLinea

    '    For Each linea As DataRow In Doc.dtLineas.Select
    '        Select Case linea.RowState
    '            Case DataRowState.Added
    '                Dim dtPromLineaOLD As DataTable = pl.SelOnPrimaryKey(linea("IDPromocionLinea"))
    '                If Not IsNothing(dtPromLineaOLD) AndAlso dtPromLineaOLD.Rows.Count > 0 Then
    '                    If linea("Cantidad") >= dtPromLineaOLD.Rows(0)("QMinPedido") Then
    '                        If linea("Cantidad") > dtPromLineaOLD.Rows(0)("QMaxPedido") Then
    '                            linea("Cantidad") = dtPromLineaOLD.Rows(0)("QMaxPedido")
    '                        End If

    '                        Dim datosPromo As PromocionLinea.DatosActuaLinPromoDr
    '                        datosPromo.Dr = linea
    '                        ProcessServer.ExecuteTask(Of PromocionLinea.DatosActuaLinPromoDr)(AddressOf PromocionLinea.ActualizarLineaPromocion, datosPromo, services)
    '                        Dim datosRegalo As New DataNuevaLineaRegalo(Doc, linea)
    '                        ProcessServer.ExecuteTask(Of DataNuevaLineaRegalo)(AddressOf NuevaLineaRegalo, datosRegalo, services)
    '                    End If
    '                End If
    '            Case DataRowState.Modified
    '                If Length(linea("IDLineaAlbaran")) = 0 Then
    '                    Dim intIDPromocionLineaOLD As Integer
    '                    Dim strIDPromocionOLD As String
    '                    Dim dblQServidaOLD As Double
    '                    If Nz(linea("Cantidad", DataRowVersion.Original), 0) <> Nz(linea("Cantidad"), 0) Then
    '                        intIDPromocionLineaOLD = Nz(linea("IDPromocionLinea", DataRowVersion.Original), 0)
    '                        strIDPromocionOLD = linea("IDPromocion", DataRowVersion.Original) & String.Empty
    '                        dblQServidaOLD = linea("Cantidad", DataRowVersion.Original)
    '                    End If
    '                    If intIDPromocionLineaOLD = 0 Then
    '                        intIDPromocionLineaOLD = Nz(linea("IDPromocionLinea"), 0)
    '                    End If
    '                    If intIDPromocionLineaOLD > 0 Then
    '                        'Se actualiza la cantidad promocionada
    '                        Dim dtPromLineaOLD As DataTable = pl.SelOnPrimaryKey(intIDPromocionLineaOLD)
    '                        If Not IsNothing(dtPromLineaOLD) AndAlso dtPromLineaOLD.Rows.Count > 0 Then
    '                            If linea("Cantidad") < dtPromLineaOLD.Rows(0)("QMinPedido") Then
    '                                linea("IDPromocionLinea") = intIDPromocionLineaOLD
    '                                linea("IDPromocion") = strIDPromocionOLD
    '                                linea("Cantidad") = linea("Cantidad") - dblQServidaOLD
    '                                ' pl.ActualizarLineaPromocion(dr.Table)
    '                                Dim datosPromo As PromocionLinea.DatosActuaLinPromoDr
    '                                datosPromo.Dr = linea
    '                                ProcessServer.ExecuteTask(Of PromocionLinea.DatosActuaLinPromoDr)(AddressOf PromocionLinea.ActualizarLineaPromocion, datosPromo, services)
    '                            Else
    '                                Dim dblQ As Double = linea("Cantidad")
    '                                If linea("Cantidad") > dtPromLineaOLD.Rows(0)("QMaxPedido") Then
    '                                    linea("IDPromocionLinea") = intIDPromocionLineaOLD
    '                                    linea("IDPromocion") = strIDPromocionOLD
    '                                    dblQ = dtPromLineaOLD.Rows(0)("QMaxPedido")
    '                                End If
    '                                linea("Cantidad") = linea("Cantidad") - dblQServidaOLD
    '                                'pl.ActualizarLineaPromocion(dr.Table)
    '                                Dim datosPromo As PromocionLinea.DatosActuaLinPromoDr
    '                                datosPromo.Dr = linea
    '                                ProcessServer.ExecuteTask(Of PromocionLinea.DatosActuaLinPromoDr)(AddressOf PromocionLinea.ActualizarLineaPromocion, datosPromo, services)
    '                                linea("Cantidad") = dblQ

    '                                Dim datosRegalo As New DataNuevaLineaRegalo(Doc, linea, False)
    '                                ProcessServer.ExecuteTask(Of DataNuevaLineaRegalo)(AddressOf NuevaLineaRegalo, datosRegalo, services)
    '                            End If
    '                        End If
    '                    Else
    '                        Dim datosRegalo As New DataNuevaLineaRegalo(Doc, linea)
    '                        ProcessServer.ExecuteTask(Of DataNuevaLineaRegalo)(AddressOf NuevaLineaRegalo, datosRegalo, services)
    '                    End If
    '                End If
    '        End Select
    '    Next
    'End Sub

    <Task()> Public Shared Sub TratarPromocionesLineas(ByVal Doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        For Each linea As DataRow In Doc.dtLineas.Select
            If linea("Regalo") = 0 Then
                If Nz(linea("IDLineaAlbaran"), 0) <> 0 AndAlso Nz(linea("IDPromocionLinea"), 0) <> 0 Then
                    If linea.RowState = DataRowState.Modified AndAlso _
                      (linea("Cantidad") <> linea("Cantidad", DataRowVersion.Original) OrElse linea("IdArticulo") <> linea("IdArticulo", DataRowVersion.Original)) Then
                        ProcessServer.ExecuteTask(Of Integer)(AddressOf PromocionLinea.ActualizarLineaPromocionQ, linea("IDPromocionLinea"), services)
                    End If
                End If

                '10. Quitamos la información anterior
                If linea.RowState = DataRowState.Modified Then
                    If Nz(linea("IDPromocionLinea", DataRowVersion.Original), 0) <> 0 Then
                        If (linea("Cantidad") <> linea("Cantidad", DataRowVersion.Original) OrElse linea("IdArticulo") <> linea("IdArticulo", DataRowVersion.Original)) Then
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
                      (linea("Cantidad") <> linea("Cantidad", DataRowVersion.Original) OrElse linea("IdArticulo") <> linea("IdArticulo", DataRowVersion.Original)) Then
                        Dim pl As New PromocionLinea
                        Dim dtPromLinea As DataTable = pl.SelOnPrimaryKey(linea("IDPromocionLinea"))
                        If Not IsNothing(dtPromLinea) AndAlso dtPromLinea.Rows.Count > 0 Then
                            If linea("Cantidad") >= dtPromLinea.Rows(0)("QMinPedido") Then
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
    End Sub

    Public Class DataNuevaLineaRegalo
        Public Doc As DocumentoFacturaVenta
        Public Row As DataRow
        Public RowPromocion As DataRow
        Public ActualizarPromo As Boolean

        Public Sub New(ByVal Doc As DocumentoFacturaVenta, ByVal Row As DataRow, ByVal RowPromocion As DataRow, Optional ByVal ActualizarPromo As Boolean = True)
            Me.Doc = Doc
            Me.Row = Row
            Me.RowPromocion = RowPromocion
            Me.ActualizarPromo = ActualizarPromo
        End Sub
    End Class

    <Task()> Public Shared Sub NuevaLineaRegalo(ByVal data As DataNuevaLineaRegalo, ByVal services As ServiceProvider)
        If Not IsNothing(data.Row) AndAlso Length(data.Row("IDPromocionLinea")) > 0 Then
            Dim dblQServida As Double
            If data.Row("Cantidad") > data.RowPromocion("QMaxPedido") Then
                dblQServida = data.RowPromocion("QMaxPedido")
            Else
                dblQServida = data.Row("Cantidad")
            End If

            Dim f As New Filter
            f.Add(New NumberFilterItem("IDPromocionLinea", data.Row("IDPromocionLinea")))
            f.Add(New StringFilterItem("IDArticulo", data.Row("IDArticulo")))

            Dim dtArticuloRegalo As DataTable = AdminData.GetData("vNegPromocionArticulosRegaloFactura", f)
            If Not IsNothing(dtArticuloRegalo) AndAlso dtArticuloRegalo.Rows.Count > 0 Then
                Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
                Dim strAlmacenPred As String = AppParams.Almacen
                Dim FVL As New FacturaVentaLinea
                Dim strTipoLinea As String = ProcessServer.ExecuteTask(Of Object, String)(AddressOf TipoLinea.TipoLineaRegalo, Nothing, services)
                Dim intOrden As Integer = data.Doc.dtLineas.Compute("MAX(IDOrdenLinea)", Nothing)

                Dim context As New BusinessData(data.Doc.HeaderRow)
                f.Clear()
                f.Add(New NumberFilterItem("IDFactura", data.Row("IDFactura")))
                For Each drArticuloRegalo As DataRow In dtArticuloRegalo.Rows
                    'Nuevo registro
                    Dim drFVL As DataRow = data.Doc.dtLineas.NewRow
                    drFVL("IDLineaFactura") = AdminData.GetAutoNumeric
                    drFVL("IDTipoLinea") = strTipoLinea
                    drFVL("IDFactura") = data.Doc.HeaderRow("IDFactura")
                    drFVL("IDCentroGestion") = data.Doc.HeaderRow("IDCentroGestion")
                    drFVL("Cantidad") = 0
                    drFVL = FVL.ApplyBusinessRule("IDArticulo", drArticuloRegalo("IDArticuloRegalo"), drFVL, context)

                    context("Fecha") = data.Doc.HeaderRow("FechaFactura")
                    drFVL("Regalo") = True

                    'En el campo Cantidad guardamos la Cantidad indicada con el ArticuloRegalo
                    drFVL("Cantidad") = Fix((dblQServida / drArticuloRegalo("QPedida"))) * drArticuloRegalo("QRegalo")
                    If drFVL("Cantidad") = 0 Then
                        drFVL("Cantidad") = drArticuloRegalo("QRegalo")
                    End If

                    'Se incrementa el IDOrden para cada linea de regalo generada
                    intOrden = intOrden + 1
                    drFVL("IDOrdenLinea") = intOrden

                    drFVL = FVL.ApplyBusinessRule("Cantidad", drFVL("Cantidad"), drFVL, context)
                    drFVL("IDPromocion") = data.Row("IDPromocion")
                    drFVL("IDPromocionLinea") = data.Row("IDPromocionLinea")

                    data.Doc.dtLineas.Rows.Add(drFVL)
                Next
                If data.ActualizarPromo AndAlso Length(data.Row("IDLineaPedido")) = 0 AndAlso Length(data.Row("IDLineaAlbaran")) = 0 Then
                    'Actualización QPromocionada
                    Dim PL As New PromocionLinea
                    Dim drPromocionLinea As DataRow = PL.GetItemRow(data.Row("IDPromocionLinea"))
                    drPromocionLinea("QPromocionada") = drPromocionLinea("QPromocionada") + dblQServida
                    BusinessHelper.UpdateTable(drPromocionLinea.Table)
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarQLineasPromociones(ByVal Doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If Doc.HeaderRow("Estado") = enumfvcEstado.fvcNoContabilizado Then
            For Each linea As DataRow In Doc.dtLineas.Select
                If linea("Regalo") = 0 Then
                    '10. Quitamos la información anterior
                    If linea.RowState = DataRowState.Modified Then
                        If Nz(linea("IDPromocionLinea", DataRowVersion.Original), 0) <> 0 Then
                            If (linea("Cantidad") <> linea("Cantidad", DataRowVersion.Original) OrElse linea("IdArticulo") <> linea("IdArticulo", DataRowVersion.Original)) Then
                                ProcessServer.ExecuteTask(Of Integer)(AddressOf PromocionLinea.ActualizarLineaPromocionQ, linea("IDPromocionLinea", DataRowVersion.Original), services)
                            End If
                        End If
                    End If

                    '30. Actualizamos en función de la Cantidad.
                    If linea.RowState = DataRowState.Modified Then
                        If Nz(linea("IDPromocionLinea"), 0) <> 0 AndAlso Nz(linea("Cantidad"), 0) <> Nz(linea("Cantidad", DataRowVersion.Original), 0) Then
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

#Region " IVA Caja "

    <Task()> Public Shared Sub FechaParaDeclaracionComoProveedor(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider)
        If Not Nz(data("FechaDeclaracionManual"), False) AndAlso data.ContainsKey("IDCliente") Then
            Dim AppParams As ParametroVenta = services.GetService(Of ParametroVenta)()
            Dim IVACaja As Boolean = AppParams.IvaCajaCircuitoVentas
            If IVACaja Then
                data("FechaParaDeclaracion") = New Date(Year(Nz(data("FechaFactura"), Today)) + 1, 12, 31) 'NegocioGeneral.cnMAX_DATE
            End If
        End If
    End Sub

#End Region

End Class

Public Class LineasFacturaEliminadas
    Public IDLineas As Hashtable

    Public Sub New()
        IDLineas = New Hashtable
    End Sub
End Class

