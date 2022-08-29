Public Class ProcesoFacturacionCompra

    <Task()> Public Shared Function CrearDocumento(ByVal data As UpdatePackage, ByVal services As ServiceProvider) As DocumentoFacturaCompra
        Return New DocumentoFacturaCompra(data)
    End Function

#Region "Agrupaciones"
    <Task()> Public Shared Function AgruparAlbaranesCompra(ByVal data As DataPrcFacturacionCompra, ByVal services As ServiceProvider) As FraCabCompra()
        Dim dtLineas As DataTable

        'se seleccionan todas las lineas de albaran no facturadas

        Dim strViewName As String = "vNegCompraCrearFactura"

        If data.IDAlbaranes.Length > 0 Then
            Dim values(data.IDAlbaranes.Length - 1) As Object
            data.IDAlbaranes.CopyTo(values, 0)
            Dim oFltr As New Filter
            oFltr.Add(New InListFilterItem("IDAlbaran", values, FilterType.Numeric))
            oFltr.Add(New NumberFilterItem("EstadoFactura", FilterOperator.NotEqual, enumaclEstadoFactura.aclFacturado))
            oFltr.Add(New NumberFilterItem("TipoLineaAlbaran", FilterOperator.NotEqual, enumaclTipoLineaAlbaran.aclComponente))

            dtLineas = New BE.DataEngine().Filter(strViewName, oFltr)
        End If
        If Not dtLineas Is Nothing AndAlso dtLineas.Rows.Count > 0 Then
            Dim p As New Parametro
            Dim strCondicionPago As String = p.CondicionPago

            Dim oGrprUser As New GroupUserAlbaranCompra '(data.DteFechaFactura)

            Dim grpAlb As DataColumn() = ProcessServer.ExecuteTask(Of DataGetGroupColumns, DataColumn())(AddressOf GetGroupColumns, New DataGetGroupColumns(dtLineas, enummpAgrupFactura.mpAlbaran), services)
            Dim grpProv As DataColumn() = ProcessServer.ExecuteTask(Of DataGetGroupColumns, DataColumn())(AddressOf GetGroupColumns, New DataGetGroupColumns(dtLineas, enummpAgrupFactura.mpProveedor), services)

            Dim groupers(1) As GroupHelper
            groupers(enummpAgrupFactura.mpAlbaran) = New GroupHelper(grpAlb, oGrprUser)
            groupers(enummpAgrupFactura.mpProveedor) = New GroupHelper(grpProv, oGrprUser)

            For Each rwLin As DataRow In dtLineas.Rows
                groupers(rwLin("AgrupFactura")).Group(rwLin)
            Next

            Return oGrprUser.Fras
        Else
            ApplicationService.GenerateError("No hay elementos a Facturar. Revise sus datos.")
        End If
    End Function
    'David Velasco 25/7/22
    <Task()> Public Shared Function AgruparAlbaranesCompraPiso(ByVal data As DataPrcFacturacionCompraPiso, ByVal services As ServiceProvider) As FraCabCompra()
        Dim dtLineas As DataTable

        'se seleccionan todas las lineas de albaran no facturadas

        Dim strViewName As String = "vFacturasPisos"

        If data.IDPisoPago.Length > 0 Then
            Dim values(data.IDPisoPago.Length - 1) As Object
            data.IDPisoPago.CopyTo(values, 0)
            Dim oFltr As New Filter
            oFltr.Add(New InListFilterItem("IDPisoPago", values, FilterType.String))

            dtLineas = New BE.DataEngine().Filter(strViewName, oFltr)
        End If
        If Not dtLineas Is Nothing AndAlso dtLineas.Rows.Count > 0 Then
            Dim p As New Parametro
            Dim strCondicionPago As String = p.CondicionPago

            Dim oGrprUser As New GroupUserPisoCompra '(data.DteFechaFactura)

            Dim grpAlb As DataColumn() = ProcessServer.ExecuteTask(Of DataGetGroupColumns, DataColumn())(AddressOf GetGroupColumns, New DataGetGroupColumns(dtLineas, enummpAgrupFactura.mpAlbaran), services)
            Dim grpProv As DataColumn() = ProcessServer.ExecuteTask(Of DataGetGroupColumns, DataColumn())(AddressOf GetGroupColumns, New DataGetGroupColumns(dtLineas, enummpAgrupFactura.mpProveedor), services)

            Dim groupers(1) As GroupHelper
            groupers(enummpAgrupFactura.mpAlbaran) = New GroupHelper(grpAlb, oGrprUser)
            groupers(enummpAgrupFactura.mpProveedor) = New GroupHelper(grpProv, oGrprUser)

            For Each rwLin As DataRow In dtLineas.Rows
                groupers(rwLin("AgrupFactura")).Group(rwLin)
            Next

            Return oGrprUser.Fras
        Else
            ApplicationService.GenerateError("No hay elementos a Facturar. Revise sus datos.")
        End If
    End Function

    <Task()> Public Shared Function AgruparAlbaranesAutoFraCompra(ByVal data As DataPrcAutofacturacionCompra, ByVal services As ServiceProvider) As FraCabCompra()

        Dim dtLineas As DataTable

        'se seleccionan todas las lineas de albaran no facturadas

        Dim htLins As New Hashtable
        Dim ids(data.IDAlbaranes.Length - 1) As Object
        For i As Integer = 0 To data.IDAlbaranes.Length - 1
            ids(i) = data.IDAlbaranes(i).IDLineaAlbaran
            htLins.Add(data.IDAlbaranes(i).IDLineaAlbaran, data.IDAlbaranes(i))
        Next

        Dim strViewName As String = "vNegCompraAutoFacturacion"

        If ids.Length > 0 Then
            Dim oFltr As New Filter
            oFltr.Add(New InListFilterItem("IDLineaAlbaran", ids, FilterType.Numeric))
            oFltr.Add(New NumberFilterItem("EstadoFactura", FilterOperator.NotEqual, enumaclEstadoFactura.aclFacturado))
            oFltr.Add(New NumberFilterItem("TipoLineaAlbaran", FilterOperator.NotEqual, enumaclTipoLineaAlbaran.aclComponente))
            dtLineas = New BE.DataEngine().Filter(strViewName, oFltr)
        End If

        If Not dtLineas Is Nothing AndAlso dtLineas.Rows.Count > 0 Then
            Dim p As New Parametro
            Dim strCondicionPago As String = p.CondicionPago

            Dim oGrprUser As New GroupUserAlbaranCompra '(data.DteFechaFactura)

            Dim grpAlb As DataColumn() = ProcessServer.ExecuteTask(Of DataGetGroupColumns, DataColumn())(AddressOf GetGroupColumns, New DataGetGroupColumns(dtLineas, enummpAgrupFactura.mpAlbaran), services)
            Dim grpProv As DataColumn() = ProcessServer.ExecuteTask(Of DataGetGroupColumns, DataColumn())(AddressOf GetGroupColumns, New DataGetGroupColumns(dtLineas, enummpAgrupFactura.mpProveedor), services)

            Dim groupers(1) As GroupHelper
            groupers(enummcAgrupFactura.mcAlbaran) = New GroupHelper(grpAlb, oGrprUser)
            groupers(enummpAgrupFactura.mpProveedor) = New GroupHelper(grpProv, oGrprUser)

            For Each rwLin As DataRow In dtLineas.Rows
                groupers(rwLin("AgrupFactura")).Group(rwLin)
            Next

            For Each fra As FraCabCompra In oGrprUser.Fras
                For Each fralin As FraLinCompraAlbaran In CType(fra, FraCabCompraAlbaran).Lineas
                    fralin.QaFacturar = DirectCast(htLins(fralin.IDLineaAlbaran), DataAutoFact).QFacturar
                    fralin.QIntAFacturar = DirectCast(htLins(fralin.IDLineaAlbaran), DataAutoFact).QIntAFacturar
                Next
            Next

            Return oGrprUser.Fras
        Else
            ApplicationService.GenerateError("No hay elementos a Facturar. Revise sus datos.")
        End If
    End Function
    Public Class DataGetGroupColumns
        Public Table As DataTable
        Public Agrupacion As enummpAgrupFactura

        Public Sub New(ByVal Table As DataTable, ByVal Agrupacion As enummpAgrupFactura)
            Me.Table = Table
            Me.Agrupacion = Agrupacion
        End Sub
    End Class
    <Task()> Public Shared Function GetGroupColumns(ByVal data As DataGetGroupColumns, ByVal services As ServiceProvider) As DataColumn()
        Dim columns(5) As DataColumn
        columns(0) = data.Table.Columns("IDProveedor")
        columns(1) = data.Table.Columns("IDFormaPago")
        columns(2) = data.Table.Columns("IDCondicionPago")
        columns(3) = data.Table.Columns("IDBancoPropio") '// banco propio??
        columns(4) = data.Table.Columns("IdMoneda")
        'columns(5) = table.Columns("Dto")
        columns(5) = data.Table.Columns("IDDireccion")
        If data.Agrupacion = enummpAgrupFactura.mpAlbaran Then
            ReDim Preserve columns(6)
            columns(6) = data.Table.Columns("IDAlbaran")
        End If
        Return columns
    End Function
#End Region

#Region " Agrupaciones Leasing "

    <Task()> Public Shared Function AgruparPagosLeasing(ByVal data As DataPrcFacturacionCompraLeasing, ByVal services As ServiceProvider) As FraCabCompra()
        Dim dtLineas As DataTable
        Dim strViewName As String = "vNegCompraCrearFacturaLeasing"

        If data.IDPagos.Length > 0 Then
            Dim values(data.IDPagos.Length - 1) As Object
            data.IDPagos.CopyTo(values, 0)
            Dim oFltr As New Filter
            oFltr.Add(New InListFilterItem("IDPago", values, FilterType.Numeric))
            dtLineas = New BE.DataEngine().Filter(strViewName, oFltr)
        End If
        If Not dtLineas Is Nothing AndAlso dtLineas.Rows.Count > 0 Then
            Dim oGrprUser As New GroupUserLeasing  '(data.DteFechaFactura)
            Dim grpLeasing As DataColumn() = ProcessServer.ExecuteTask(Of DataTable, DataColumn())(AddressOf GetGroupColumnLeasing, dtLineas, services)

            Dim groupers(0) As GroupHelper
            groupers(0) = New GroupHelper(grpLeasing, oGrprUser)

            For Each rwLin As DataRow In dtLineas.Rows
                groupers(0).Group(rwLin)
            Next
            Dim PagosPeriodicos As EntityInfoCache(Of PagoPeriodicoInfo) = services.GetService(Of EntityInfoCache(Of PagoPeriodicoInfo))()
            Dim IDPagoPeriodicoAnt As Integer
            Dim ContPagoPer As Integer
            For Each fra As FraCabCompraLeasing In oGrprUser.Fras
                For Each fralin As FraLinCompraLeasing In fra.Lineas
                    If IDPagoPeriodicoAnt <> fralin.IDPagoPeriodico Then
                        ContPagoPer = 1

                        '// Sumamos al contador los Pagos ya contabilizados, asociados al Pago periódico.
                        Dim fPerContabilizado As New Filter
                        fPerContabilizado.Add(New NumberFilterItem("IdPagoPeriodo", fralin.IDPagoPeriodico))
                        fPerContabilizado.Add(New BooleanFilterItem("Contabilizado", True))
                        Dim dtPagoPerContabilizado As DataTable = New Pago().Filter(fPerContabilizado)
                        If dtPagoPerContabilizado.Rows.Count > 0 Then
                            ContPagoPer += dtPagoPerContabilizado.Rows.Count
                        End If
                    Else
                        ContPagoPer += 1
                    End If

                    '// Montamos SuFactura en función del inmovilizado asociado y de los pagos periodicos contabilizados
                    Dim PagoPer As PagoPeriodicoInfo = PagosPeriodicos.GetEntity(fralin.IDPagoPeriodico)
                    If Length(PagoPer.IDInmovilizado) > 0 Then
                        fra.SuFactura = PagoPer.IDInmovilizado & "/" & CStr(ContPagoPer)
                    End If

                    '// Recuperamos el Banco Predereteminado del Proveedor
                    Dim fBancoProv As New Filter
                    fBancoProv.Add(New BooleanFilterItem("Predeterminado", True))
                    fBancoProv.Add(New StringFilterItem("IDProveedor", fra.IDProveedor))
                    Dim dtBancoProv As DataTable = New ProveedorBanco().Filter(fBancoProv)
                    If dtBancoProv.Rows.Count > 0 Then
                        fra.IDProveedorBanco = dtBancoProv.Rows(0)("IDProveedorBanco") & String.Empty
                    End If

                    IDPagoPeriodicoAnt = fralin.IDPagoPeriodico

                Next
            Next

            Return oGrprUser.Fras
        Else
            ApplicationService.GenerateError("No hay elementos a Facturar. Revise sus datos.")
        End If
    End Function

    <Task()> Public Shared Function GetGroupColumnLeasing(ByVal data As DataTable, ByVal services As ServiceProvider) As DataColumn()
        Dim columns(4) As DataColumn
        columns(0) = data.Columns("IDProveedor")
        columns(1) = data.Columns("IDFormaPago")
        columns(2) = data.Columns("IDCondicionPago")
        columns(3) = data.Columns("IDPago")
        columns(4) = data.Columns("IDMoneda")
        Return columns
    End Function

#End Region

#Region "Ordenar Facturas"
    'Ordena las facturas teniendo en cuenta las fechas de los albaranes
    <Task()> Public Shared Sub Ordenar(ByVal data As FraCabCompra(), ByVal services As ServiceProvider)
        If data IsNot Nothing Then Array.Sort(data, New OrdenFacturasCompra)
    End Sub
#End Region

#Region "Proceso Facturación "

    <Task()> Public Shared Function CrearDocumentoFactura(ByVal fra As FraCabCompra, ByVal services As ServiceProvider) As DocumentoFacturaCompra
        Return New DocumentoFacturaCompra(fra, services)
    End Function

    <Task()> Public Shared Sub AsignarValoresPredeterminadosGenerales(ByVal fra As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        Dim fraRow As DataRow = fra.HeaderRow
        'David Velasco 29/8/22 Campo idpisopagos
        Try
            If fraRow.IsNull("IDPisoPagos") Then fraRow("IDPisoPagos") = System.DBNull.Value
        Catch ex As Exception

        End Try
        'David
        If fraRow.IsNull("IDFactura") Then fraRow("IDFactura") = AdminData.GetAutoNumeric
        If fraRow.IsNull("FechaFactura") Then fraRow("FechaFactura") = Date.Today
        If fraRow.IsNull("SuFechaFactura") Then fraRow("SuFechaFactura") = fraRow("FechaFactura")
        'If fraRow.IsNull("FechaParaDeclaracion") Then fraRow("FechaParaDeclaracion") = fraRow("FechaFactura")
        If fraRow.IsNull("FechaDeclaracionManual") Then fraRow("FechaDeclaracionManual") = False
        If fraRow.IsNull("FechaParaDeclaracion") Then
            fraRow("FechaParaDeclaracion") = fraRow("FechaFactura")
            ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ProcesoFacturacionCompra.FechaParaDeclaracionPorProveedor, New DataRowPropertyAccessor(fraRow), services)
        End If

        If fraRow.IsNull("Estado") Then fraRow("Estado") = enumfccEstado.fccNoContabilizado
        If fraRow.IsNull("IVAManual") Then fraRow("IVAManual") = False
        If fraRow.IsNull("VencimientosManuales") Then fraRow("VencimientosManuales") = False
        If fraRow.IsNull("Enviar347") Then fraRow("Enviar347") = False
        If fraRow.IsNull("Tipofactura") Then fraRow("Tipofactura") = enumfccTipoFactura.fccNormal

    End Sub

    <Task()> Public Shared Sub AsignarProveedorGrupo(ByVal fra As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
        If fra.Proveedor Is Nothing Then fra.Proveedor = Proveedores.GetEntity(fra.HeaderRow("IDProveedor"))

        If Len(fra.Proveedor.GrupoProveedor) > 0 And fra.Proveedor.GrupoFactura Then
            fra.HeaderRow("IDProveedorInicial") = fra.HeaderRow("IDProveedor")
            fra.HeaderRow("IDProveedor") = fra.Proveedor.GrupoProveedor
            fra.Proveedor = Proveedores.GetEntity(fra.HeaderRow("IDProveedor"))
        End If
    End Sub

    <Task()> Public Shared Sub AsignarDatosProveedorNuevoGasto(ByVal fra As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        If Not fra.Cabecera Is Nothing AndAlso TypeOf fra.Cabecera Is FraCabCompraNuevoGasto Then
            Dim FraCab As FraCabCompraNuevoGasto = CType(fra.Cabecera, FraCabCompraNuevoGasto)
            fra.HeaderRow("CifProveedor") = FraCab.CIF
            fra.HeaderRow("RazonSocial") = FraCab.RazonSocial
            fra.HeaderRow("IDDiaPago") = FraCab.IDDiaPago
            fra.HeaderRow("IDTipoAsiento") = FraCab.IDTipoAsiento
        End If
    End Sub
    <Task()> Public Shared Sub AsignarDatosProveedor(ByVal fra As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        'AsignarDatosProveedor(New DataRowPropertyAccessor(fra.HeaderRow), fra.Info)
        ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf AsignarDatosProveedorNuevoGasto, fra, services)
        Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
        If fra.Proveedor Is Nothing Then fra.Proveedor = Proveedores.GetEntity(fra.HeaderRow("IDProveedor"))
        If fra.HeaderRow.IsNull("IDProveedorInicial") Then fra.HeaderRow("IDProveedorInicial") = fra.Proveedor.IDProveedor
        If fra.HeaderRow.IsNull("CifProveedor") Then fra.HeaderRow("CifProveedor") = fra.Proveedor.CifProveedor
        If fra.HeaderRow.IsNull("RazonSocial") Then fra.HeaderRow("RazonSocial") = fra.Proveedor.RazonSocial
        If fra.HeaderRow.IsNull("Direccion") Then fra.HeaderRow("Direccion") = fra.Proveedor.Direccion
        If fra.HeaderRow.IsNull("CodPostal") Then fra.HeaderRow("CodPostal") = fra.Proveedor.CodPostal
        If fra.HeaderRow.IsNull("Poblacion") Then fra.HeaderRow("Poblacion") = fra.Proveedor.Poblacion
        If fra.HeaderRow.IsNull("Provincia") Then fra.HeaderRow("Provincia") = fra.Proveedor.Provincia
        If fra.HeaderRow.IsNull("IDPais") Then fra.HeaderRow("IDPais") = fra.Proveedor.IDPais

        If fra.HeaderRow.IsNull("IDTipoAsiento") Then fra.HeaderRow("IDTipoAsiento") = fra.Proveedor.IDTipoAsiento
        If fra.HeaderRow.IsNull("RetencionIRPF") Then fra.HeaderRow("RetencionIRPF") = fra.Proveedor.RetencionIRPF
        If fra.HeaderRow.IsNull("RegimenEspecial") Then fra.HeaderRow("RegimenEspecial") = fra.Proveedor.RegimenEspecial
        If fra.HeaderRow.IsNull("TipoRetencionIRPF") Then fra.HeaderRow("TipoRetencionIRPF") = fra.Proveedor.TipoRetencionIRPF
        ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf ProcesoCompra.AsignarDatosProveedor, fra, services)
        If fra.HeaderRow.IsNull("DtoFactura") Then fra.HeaderRow("DtoFactura") = fra.Proveedor.DtoComercial
        If fra.HeaderRow.IsNull("IDBancoPropio") Then fra.HeaderRow("IDBancoPropio") = fra.Proveedor.IDBancoPropio
        If fra.HeaderRow.IsNull("IDDiaPago") Then fra.HeaderRow("IDDiaPago") = fra.Proveedor.IDDiaPago
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
    End Sub

    <Task()> Public Shared Sub AsignarDireccion(ByVal fra As DocumentoFacturaCompra, ByVal services As ServiceProvider)

        If Not fra.HeaderRow.IsNull("IDDireccion") And fra.HeaderRow("IDProveedor") = fra.HeaderRow("IDProveedorInicial") Then
            Dim StDatosDirec As New ProveedorDireccion.DataDirecDe
            StDatosDirec.IDDireccion = fra.HeaderRow("IDDireccion")
            StDatosDirec.TipoDireccion = enumpdTipoDireccion.pdDireccionFactura
            If Not ProcessServer.ExecuteTask(Of ProveedorDireccion.DataDirecDe, Boolean)(AddressOf ProveedorDireccion.EsDireccionDe, StDatosDirec, services) Then
                Dim StDatosDirecEnvio As New ProveedorDireccion.DataDirecEnvio
                StDatosDirecEnvio.IDProveedor = fra.HeaderRow("IDProveedor")
                StDatosDirecEnvio.TipoDireccion = enumpdTipoDireccion.pdDireccionFactura
                fra.HeaderRow("IDDireccion") = CType(ProcessServer.ExecuteTask(Of ProveedorDireccion.DataDirecEnvio, DataTable)(AddressOf ProveedorDireccion.ObtenerDireccionEnvio, StDatosDirecEnvio, services), DataTable).Rows(0)("IDDireccion")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarBanco(ByVal fra As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        If fra.HeaderRow.IsNull("IDProveedorBanco") Then
            Dim IDBanco As Integer = ProcessServer.ExecuteTask(Of String, Integer)(AddressOf New ProveedorBanco().GetBancoPredeterminado, fra.HeaderRow("IDProveedor"), services)

            If IDBanco > 0 Then
                fra.HeaderRow("IDProveedorBanco") = IDBanco
            Else
                fra.HeaderRow("IDProveedorBanco") = System.DBNull.Value
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarContador(ByVal fra As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        If fra.HeaderRow.IsNull("IDContador") Then
            Dim Info As ProcessInfo = services.GetService(Of ProcessInfo)()
            If Len(Info.IDContador) > 0 Then
                fra.HeaderRow("IDContador") = Info.IDContador
            Else
                If Len(fra.Proveedor.IDContadorCargo) > 0 Then
                    fra.HeaderRow("IDContador") = fra.Proveedor.IDContadorCargo
                Else
                    ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf ProcesoComunes.AsignarContador, fra, services)
                End If
            End If
        End If
        'TODO acceso a contador duplicado
        fra.AIva = New Contador().GetItemRow(fra.HeaderRow("IDContador"))("AIva")
    End Sub

    'David Velasco 27/7/22 
    <Task()> Public Shared Sub AsignarContadorPiso(ByVal fra As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf ProcesoComunes.AsignarContadorPiso, fra, services)
    End Sub

    <Task()> Public Shared Sub AsignarNumeroFactura(ByVal fra As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        If fra.HeaderRow.RowState = DataRowState.Added Then
            If Not IsDBNull(fra.HeaderRow("IDContador")) Then
                Dim StDatos As New Contador.DatosCounterValue
                StDatos.IDCounter = fra.HeaderRow("IDContador")
                StDatos.TargetClass = New FacturaCompraCabecera
                StDatos.TargetField = "NFactura"
                StDatos.DateField = "FechaFactura"
                StDatos.DateValue = fra.HeaderRow("FechaFactura")
                StDatos.IDEjercicio = fra.HeaderRow("IDEjercicio") & String.Empty

                Dim NFraAnterior As String = fra.HeaderRow("NFactura") & String.Empty
                fra.HeaderRow("NFactura") = ProcessServer.ExecuteTask(Of Contador.DatosCounterValue, String)(AddressOf Contador.CounterValue, StDatos, services)
                If Length(NFraAnterior) > 0 AndAlso NFraAnterior = fra.HeaderRow("SuFactura") AndAlso NFraAnterior <> fra.HeaderRow("NFactura") Then
                    fra.HeaderRow("SuFactura") = System.DBNull.Value
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarNumeroFacturaPropuesta(ByVal fra As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        If fra.HeaderRow.IsNull("NFactura") Then
            Dim counters As ProvisionalCounter = services.GetService(Of ProvisionalCounter)()
            fra.HeaderRow("NFactura") = counters.GetCounterValue(fra.HeaderRow("IDContador"))
        End If
    End Sub

    <Task()> Public Shared Sub AsignarSuFactura(ByVal fra As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        Dim ProcInfo As ProcessInfoFra = services.GetService(Of ProcessInfoFra)()
        If fra.HeaderRow.IsNull("SuFactura") AndAlso Length(ProcInfo.SuFactura) > 0 Then
            fra.HeaderRow("SuFactura") = ProcInfo.SuFactura
        End If
        If fra.HeaderRow.IsNull("SuFactura") AndAlso Not fra.HeaderRow.IsNull("NFactura") Then
            fra.HeaderRow("SuFactura") = fra.HeaderRow("NFactura")
        End If

        If Not fra.Cabecera Is Nothing AndAlso TypeOf fra.Cabecera Is FraCabCompraLeasing AndAlso Length(CType(fra.Cabecera, FraCabCompraLeasing).SuFactura) > 0 Then
            fra.HeaderRow("SuFactura") = CType(fra.Cabecera, FraCabCompraLeasing).SuFactura
        End If
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoFacturacionCompra.ValidarSuFactura, fra.HeaderRow, services)
    End Sub

    <Task()> Public Shared Sub AsignarSuFechaFactura(ByVal fra As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        Dim ProcInfo As ProcessInfoFra = services.GetService(Of ProcessInfoFra)()
        If Length(ProcInfo.SuFechaFactura) > 0 Then
            fra.HeaderRow("SuFechaFactura") = ProcInfo.SuFechaFactura

            'David Velasco 27/7/22
            'Para vincular la linea de los pagos de los pisos con la factura
            Try
                fra.HeaderRow("IDPisoPagos") = ProcInfo.IDPisosPagos
            Catch ex As Exception

            End Try
            Dim AppParams As ParametroFacturaCompra = services.GetService(Of ParametroFacturaCompra)()
            If AppParams.ControlarFechaFCProveedor Then
                fra.HeaderRow("FechaFactura") = ProcInfo.SuFechaFactura
                fra.HeaderRow("FechaParaDeclaracion") = fra.HeaderRow("FechaFactura")
                ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ProcesoFacturacionCompra.FechaParaDeclaracionPorProveedor, New DataRowPropertyAccessor(fra.HeaderRow), services)
            End If

        End If
    End Sub

    <Task()> Public Shared Function Resultado(ByVal Doc As DocumentoFacturaCompra, ByVal services As ServiceProvider) As CreateElement
        Dim result As New CreateElement
        result.IDElement = Doc.HeaderRow("IDFactura")
        result.NElement = Doc.HeaderRow("NFactura")
        Return result
    End Function

    <Task()> Public Shared Sub AsignarClaveOperacion(ByVal fra As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        '  fra.HeaderRow("ClaveOperacion") = System.DBNull.Value     
        If fra.HeaderRow("Estado") = enumfccEstado.fccNoContabilizado Then
            If Not IsDBNull(fra.HeaderRow("IDContador")) Then
                Dim Contadores As EntityInfoCache(Of ContadorInfo) = services.GetService(Of EntityInfoCache(Of ContadorInfo))()
                Dim ContInfo As ContadorInfo = Contadores.GetEntity(fra.HeaderRow("IDContador"))
                If Length(ContInfo.IDTipoComprobante) > 0 AndAlso Length(ContInfo.ClaveOperacion) > 0 Then
                    '//Le asignamos la clave de operación del Tipo de Comprobante asociado al contador.
                    fra.HeaderRow("ClaveOperacion") = ContInfo.ClaveOperacion

                End If
            End If

            If Not fra.dtFCBI Is Nothing AndAlso fra.dtFCBI.Rows.Count > 0 Then
                Dim fLineaBI As New Filter
                fLineaBI.Add(New NumberFilterItem("BaseImponible", FilterOperator.NotEqual, 0))
                Dim WhereLineaBI As String = fLineaBI.Compose(New AdoFilterComposer)
                Dim adr() As DataRow = fra.dtFCBI.Select(WhereLineaBI, Nothing, DataViewRowState.CurrentRows)
                If adr.Length > 1 Then
                    fra.HeaderRow("ClaveOperacion") = ClaveOperacion.FacturaVariosTiposImpositivos
                Else
                    If Not fra.HeaderRow.IsNull("ClaveOperacion") AndAlso (fra.HeaderRow("ClaveOperacion") = ClaveOperacion.FacturaVariosTiposImpositivos Or fra.HeaderRow("ClaveOperacion") = ClaveOperacion.AdquisicionesIntracomunitariasBienes Or fra.HeaderRow("ClaveOperacion") = ClaveOperacion.FacturaRectificativa Or fra.HeaderRow("ClaveOperacion") = ClaveOperacion.InversionSujetoPasivo) Then
                        fra.HeaderRow("ClaveOperacion") = System.DBNull.Value
                    End If
                End If
                If fra.dtFCBI.Rows(0).RowState <> DataRowState.Deleted Then
                    If Nz(fra.dtFCBI.Rows(0)("ImpIntrastat"), 0) <> 0 Then
                        fra.HeaderRow("ClaveOperacion") = ClaveOperacion.AdquisicionesIntracomunitariasBienes
                    End If
                Else
                    If adr.Length > 0 AndAlso adr(0).RowState = DataRowState.Added Then
                        If Nz(adr(0)("ImpIntrastat"), 0) <> 0 Then
                            fra.HeaderRow("ClaveOperacion") = ClaveOperacion.AdquisicionesIntracomunitariasBienes
                        End If
                    End If
                End If
            End If

            If Length(fra.HeaderRow("IDFacturaRectificada")) > 0 Then
                fra.HeaderRow("ClaveOperacion") = ClaveOperacion.FacturaRectificativa
            End If

            If Nz(fra.HeaderRow("ImpSinRepercutir"), 0) <> 0 Then
                fra.HeaderRow("ClaveOperacion") = ClaveOperacion.InversionSujetoPasivo
            End If

        End If
    End Sub
    'David Velasco 26/07/22
    <Task()> Public Shared Sub CrearLineasDesdePiso(ByVal docFactura As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        '0.Recupero la linea de la que tengo que sacar el Pago del piso y de los suministros.
        Dim dt As New DataTable
        Dim filtro As New Filter
        filtro.Add("IDPisoPago", docFactura.HeaderRow("IDPisoPagos"))

        dt = New BE.DataEngine().Filter("vFrmPagoPisosMes", filtro)

        '1.Creo las instancias
        Dim lineas As DataTable = docFactura.dtLineas
        Dim fraCabAlb As FraCabCompra = docFactura.Cabecera
        'Dim oFCL As New FacturaCompraLinea
        '2.Relleno datos

        '2.1 Formo la linea para el pago del piso ->TotalAlquiler
        If dt(0)("TotalAlquiler") <> 0 Then

            Dim linea As DataRow = lineas.NewRow
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarValoresPredeterminadosLineaObligatorios, linea, services)
            linea("IDFactura") = docFactura.HeaderRow("IDFactura")
            linea("IDArticulo") = "PISO"
            Dim dtArticulo As New DataTable
            Dim f As New Filter
            f.Add("IDArticulo", FilterOperator.Equal, "PISO")
            dtArticulo = New BE.DataEngine().Filter("tbMaestroArticulo", f)

            linea("DescArticulo") = dtArticulo(0)("DescArticulo")
            linea("CContable") = dtArticulo(0)("CCCompra")
            linea("SeguimientoTarifa") = "PAGO GENERADO DE PISOS"
            linea("Precio") = dt(0)("TotalAlquiler")
            linea("Cantidad") = 1
            linea("QInterna") = 1
            lineas.Rows.Add(linea.ItemArray)
        End If
        '2.2 Formo la linea para el pago de los suminstros -->TotalPisos
        If dt(0)("TotalGastos") <> 0 Then

            Dim linea As DataRow = lineas.NewRow
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarValoresPredeterminadosLineaObligatorios, linea, services)
            linea("IDFactura") = docFactura.HeaderRow("IDFactura")
            linea("IDArticulo") = "SUMINISTRO"
            Dim dtArticulo As New DataTable
            Dim f As New Filter
            f.Add("IDArticulo", FilterOperator.Equal, "SUMINISTRO")
            dtArticulo = New BE.DataEngine().Filter("tbMaestroArticulo", f)

            linea("DescArticulo") = dtArticulo(0)("DescArticulo")
            linea("CContable") = dtArticulo(0)("CCCompra")
            linea("SeguimientoTarifa") = "PAGO GENERADO DE PISOS"
            linea("Precio") = dt(0)("TotalGastos")
            linea("Cantidad") = 1
            linea("QInterna") = 1
            lineas.Rows.Add(linea.ItemArray)
        End If

    End Sub
    <Task()> Public Shared Sub CrearLineasDesdeAlbaran(ByVal docFactura As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        Dim dtAlbaran As DataTable = ProcessServer.ExecuteTask(Of DocumentoFacturaCompra, DataTable)(AddressOf RecuperarDatosAlbaran, docFactura, services)
        Dim lineas As DataTable = docFactura.dtLineas
        If lineas Is Nothing Then
            Dim oFCL As New FacturaCompraLinea
            lineas = oFCL.AddNew
            docFactura.Add(GetType(FacturaCompraLinea).Name, lineas)
        End If

        Dim fraCabAlb As FraCabCompra = docFactura.Cabecera
        For Each albaran As DataRow In dtAlbaran.Rows
            Dim fralin As FraLinCompraAlbaran = Nothing
            For i As Integer = 0 To CType(fraCabAlb, FraCabCompraAlbaran).Lineas.Length - 1
                If albaran("IDLineaAlbaran") = CType(fraCabAlb, FraCabCompraAlbaran).Lineas(i).IDLineaAlbaran Then
                    fralin = CType(fraCabAlb, FraCabCompraAlbaran).Lineas(i)
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
                    linea("IdLineaAlbaran") = albaran("IdLineaAlbaran")
                    linea("IDAlbaran") = albaran("IDAlbaran")
                    linea("IDArticulo") = albaran("IDArticulo")
                    linea("DescArticulo") = albaran("DescArticulo")
                    linea("RefProveedor") = albaran("RefProveedor")
                    linea("DescRefProveedor") = albaran("DescRefProveedor")
                    linea("lote") = albaran("lote")
                    linea("IDTipoIva") = albaran("IDTipoIva")
                    If Length(albaran("IDCentroGestion")) > 0 Then
                        linea("IDCentroGestion") = albaran("IDCentroGestion")
                    Else
                        Dim drCabecera As DataRow = New AlbaranCompraCabecera().GetItemRow(albaran("IDAlbaran"))
                        linea("IDCentroGestion") = drCabecera("IDCentroGestion")
                    End If
                    linea("Cantidad") = dblCantidad
                    ''''
                    'linea("Factor") = albaran("Factor")
                    'linea("QInterna") = dblCantidad * albaran("Factor")
                    If TypeOf fralin Is FraLinCompraAlbaran AndAlso Not Double.IsNaN(fralin.QIntAFacturar) Then
                        '//Autofacturación
                        If linea("Cantidad") <> 0 Then
                            linea("QInterna") = fralin.QIntAFacturar
                            linea("Factor") = linea("QInterna") / linea("Cantidad")
                        Else
                            linea("QInterna") = 0
                            linea("Factor") = 0
                        End If
                    Else
                        '//Facturación General
                        linea("Factor") = albaran("Factor")
                        linea("QInterna") = dblCantidad * albaran("Factor")
                    End If
                    ''''
                    If ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, linea("IDArticulo"), services) AndAlso Length(albaran("IdLineaAlbaran")) > 0 Then
                        If linea.Table.Columns.Contains("QInterna2") Then
                            Dim ACL As New AlbaranCompraLinea
                            Dim dtACL As DataTable = ACL.SelOnPrimaryKey(albaran("IdLineaAlbaran"))
                            If dtACL.Rows.Count > 0 AndAlso dtACL.Columns.Contains("QInterna2") AndAlso Length(dtACL.Rows(0)("QInterna2")) > 0 Then
                                linea("QInterna2") = dtACL.Rows(0)("QInterna2")
                            End If
                        End If
                    End If
                    linea("IDUDInterna") = albaran("IDUDInterna")
                    linea("IDUDMedida") = albaran("IDUDMedida")
                    linea("UdValoracion") = albaran("UdValoracion")
                    linea("Precio") = albaran("Precio")
                    linea("Dto1") = albaran("Dto1")
                    linea("Dto2") = albaran("Dto2")
                    linea("Dto3") = albaran("Dto3")
                    linea("Dto") = albaran("Dto")
                    linea("DtoProntoPago") = albaran("DtoProntoPago")
                    'If AppParamsConta.Contabilidad Then 
                    linea("CContable") = albaran("CContable")
                    linea("IDMntoOTPrev") = albaran("IDMntoOTPrev")
                    linea("IDTrabajo") = albaran("IDTrabajo")
                    linea("IDObra") = albaran("IDObra")
                    linea("TipoGastoObra") = albaran("TipoGastoObra")
                    linea("IDConcepto") = albaran("IDConcepto")
                    linea("IDLineaPadre") = albaran("IDLineaPadre")
                    linea("IDOrdenLinea") = albaran("IDOrdenLinea")
                    linea("TipoLineaFactura") = albaran("TipoLineaAlbaran")
                    linea("SeguimientoTarifa") = albaran("SeguimientoTarifa")
                    linea("Texto") = albaran("Texto")
                    lineas.Rows.Add(linea.ItemArray)
                End If
            End If
        Next
    End Sub

    <Task()> Public Shared Function RecuperarDatosLeasing(ByVal Doc As DocumentoFacturaCompra, ByVal services As ServiceProvider) As DataTable
        Dim fraCab As FraCabCompraLeasing = Doc.Cabecera
        Dim f As New Filter(FilterUnionOperator.Or)
        For Each fraLin As FraLinCompraLeasing In fraCab.Lineas
            f.Add("IDPago", fraLin.IDPago)
        Next
        Return AdminData.GetData("vCtlCIPagoContGeneraFactura", f)
    End Function

    <Task()> Public Shared Sub AsignarEstadoFactura(ByVal Doc As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        'Doc.HeaderRow("VencimientosManuales") = True
        Doc.HeaderRow("Estado") = enumfccEstado.fccContabilizado
    End Sub

    <Task()> Public Shared Sub CrearLineasDesdePagosLeasing(ByVal Doc As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        Dim dtPagos As DataTable = ProcessServer.ExecuteTask(Of DocumentoFacturaCompra, DataTable)(AddressOf RecuperarDatosLeasing, Doc, services)
        Dim lineas As DataTable = Doc.dtLineas
        Dim oFCL As New FacturaCompraLinea
        If lineas Is Nothing Then
            lineas = oFCL.AddNew
            Doc.Add(GetType(FacturaCompraLinea).Name, lineas)
        End If
        Dim context As New BusinessData(Doc.HeaderRow)
        Dim ArticulosProveedores As EntityInfoCache(Of ArticuloProveedorInfo) = services.GetService(Of EntityInfoCache(Of ArticuloProveedorInfo))()
        Dim fraCab As FraCabCompraLeasing = Doc.Cabecera
        For Each pago As DataRow In dtPagos.Rows
            Dim linea As DataRow = lineas.NewRow

            ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarValoresPredeterminadosLinea, linea, services)

            linea("IDFactura") = Doc.HeaderRow("IDFactura")
            linea("IDCentroGestion") = Doc.HeaderRow("IDCentroGestion")
            linea("Cantidad") = 1
            linea("Factor") = 1
            linea("QInterna") = 1
            linea("UdValoracion") = 1
            linea("Dto1") = 0
            linea("Dto2") = 0
            linea("Dto3") = 0
            linea("Dto") = 0
            linea("DtoProntoPago") = 0

            linea = oFCL.ApplyBusinessRule("IDArticulo", pago("IDArticulo"), linea, context)
            If Length(pago("IDProveedor")) > 0 Then
                Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
                Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(pago("IDProveedor"))
                linea("CContable") = ProvInfo.CCInMovilizadoCortoPlazo
            End If

            If Length(pago("IDUDInterna")) > 0 AndAlso linea("IDUDInterna") & String.Empty <> pago("IDUDInterna") Then
                linea = oFCL.ApplyBusinessRule("IDUDInterna", pago("IDUDInterna"), linea, context)
            End If

            linea = oFCL.ApplyBusinessRule("Precio", pago("Importe"), linea, context)

            linea("TipoGastoObra") = enumfclTipoGastoObra.enumfclMaterial

            lineas.Rows.Add(linea.ItemArray)

            '// Asignamos los Pagos a la Factura
            Doc.HeaderRow("VencimientosManuales") = True
            pago("IDFactura") = Doc.HeaderRow("IDFactura")
            pago("NFactura") = Doc.HeaderRow("NFactura")
            Doc.dtPagos.ImportRow(pago)
        Next
    End Sub

    'Private Function NuevaLineaFacturaLeasing(ByVal lngIDFactura As Integer, ByVal strNFactura As String, _
    '                                         ByRef dvData As DataView, ByVal dtmFechaFactura As Date, ByVal strIDProveedor As String, _
    '                                         ByVal Lineas As DataTable, ByVal strIDCentroGestion As String, _
    '                                         ByVal strIDMoneda As String, ByRef DtPagoLeasing As DataTable) As DataTable

    '    Dim rcsComponentes As Recordset
    '    Dim strIN As String
    '    Dim strWhere As String
    '    Dim strIDAlmacen As String
    '    Dim dblCantidad As Double
    '    Dim lngContadorLineas As Integer
    '    Dim blnGestionStock As Boolean
    '    Dim ClsArtProv As New ArticuloProveedor
    '    Dim ClsProv As New Proveedor
    '    Dim ClsUds As New ArticuloUnidadAB
    '    Dim ClsCompra As New Compra
    '    '''''''''''''''''''''''''''''''''
    '    Dim fcl As New FacturaCompraLinea
    '    If Not dvData Is Nothing Then
    '        For Each dr As DataRowView In dvData
    '            If Len(strIN) Then strIN = strIN & ","
    '            strIN = strIN & dr.Row("IdPago")
    '        Next
    '        strWhere = "IDPago IN (" & strIN & ")"
    '        Dim DtPago As DataTable = AdminData.GetData("select * from vCtlCIPagoContGeneraFactura", , strWhere)

    '        If Not DtPago Is Nothing AndAlso DtPago.Rows.Count > 0 Then
    '            For Each pago As DataRow In DtPago.Rows
    '                dvData.RowFilter = "IDPago = " & pago("IDPago")
    '                If dvData.Count > 0 Then
    '                    lngContadorLineas = 0
    '                    Dim linea As DataRow = fcl.AddNewForm.Rows(0)
    '                    linea("IDFactura") = lngIDFactura

    '                    Dim DrNew As DataRow = DtPagoLeasing.NewRow()
    '                    DrNew("IDFactura") = lngIDFactura
    '                    DrNew("IDPago") = pago("IDPago")
    '                    DrNew("NFactura") = strNFactura
    '                    DtPagoLeasing.Rows.Add(DrNew)


    '                    linea("NFactura") = strNFactura
    '                    linea("IDCentroGestion") = strIDCentroGestion
    '                    linea("Cantidad") = 1
    '                    linea("Factor") = 1
    '                    linea("QInterna") = 1
    '                    linea("UdValoracion") = 1
    '                    linea("Precio") = pago("Importe")
    '                    linea("Importe") = pago("Importe")
    '                    If Lenght(pago("IDProveedor")) > 0 Then
    '                        Dim DtProv As DataTable = ClsProv.SelOnPrimaryKey(pago("IDProveedor"))
    '                        If Not DtProv Is Nothing AndAlso DtProv.Rows.Count > 0 Then
    '                            linea("CContable") = DtProv.Rows(0)("CCInMovilizadoCortoPlazo")
    '                        End If
    '                    End If
    '                    linea("Dto1") = 0
    '                    linea("Dto2") = 0
    '                    linea("Dto3") = 0
    '                    linea("Dto") = 0
    '                    linea("DtoProntoPago") = 0
    '                    linea("TipoGastoObra") = enumfclTipoGastoObra.enumfclMaterial
    '                    linea("IDUDInterna") = pago("IDUDInterna")
    '                    Dim StrWhereArt As String = "IDArticulo ='" & pago("IDArticulo") & "' AND Compra=1 AND Activo=1"
    '                    Dim DtArt As DataTable = AdminData.Filter("vNegCaractArticulo", , StrWhereArt)
    '                    If Not DtArt Is Nothing AndAlso DtArt.Rows.Count > 0 Then
    '                        linea("IDArticulo") = DtArt.Rows(0)("IDArticulo")
    '                        Dim DtRef As DataTable = ClsArtProv.SelOnPrimaryKey(pago("IDProveedor"), pago("IDArticulo"))
    '                        If Not DtRef Is Nothing AndAlso DtRef.Rows.Count > 0 Then
    '                            linea("RefProveedor") = DtRef.Rows(0)("RefProveedor") & String.Empty
    '                            linea("DescArticulo") = DtRef.Rows(0)("DescRefProveedor") & String.Empty
    '                            linea("IDUDMedida") = DtRef.Rows(0)("IdUdCompra") & String.Empty
    '                            linea("UdValoracion") = DtRef.Rows(0)("UdValoracion")
    '                        End If
    '                        If Lenght(linea("DescArticulo")) = 0 Then linea("DescArticulo") = DtArt.Rows(0)("DescArticulo")
    '                        linea("IDTipoIva") = ClsCompra.ObtenerIVA(pago("IDProveedor"), pago("IDArticulo"))
    '                        If linea("UdValoracion") = 0 Then linea("UdValoracion") = IIf(DtArt.Rows(0)("UdValoracion") > 0, DtArt.Rows(0)("UdValoracion"), 1)
    '                        If Lenght(linea("IDUDMedida")) = 0 Then linea("IDUDMedida") = DtArt.Rows(0)("IDUDInterna")
    '                    End If
    '                    MantenimientoValoresAyB(linea, strIDMoneda, Today.Date)
    '                    Lineas.Rows.Add(linea.ItemArray)
    '                End If
    '            Next
    '        End If
    '        dvData.RowFilter = String.Empty
    '        Return Lineas
    '    End If
    'End Function


    <Task()> Public Shared Sub CalcularImporteLineasFacturas(ByVal docFactura As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        If docFactura.HeaderRow("Estado") = enumfccEstado.fccNoContabilizado Then
            ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf ProcesoComunes.CalcularImporteLineas, docFactura, services)
        End If
    End Sub

    <Task()> Public Shared Sub CalcularAnaliticaFacturas(ByVal docFactura As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        If docFactura.HeaderRow("Estado") = enumfccEstado.fccNoContabilizado Then
            ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf NegocioGeneral.CalcularAnalitica, docFactura, services)
        End If

    End Sub

    <Task()> Public Shared Function RecuperarDatosAlbaran(ByVal docFactura As DocumentoFacturaCompra, ByVal services As ServiceProvider) As DataTable

        Dim fcl As New FacturaCompraLinea
        Dim acl As New AlbaranCompraLinea

        Dim fraCabAlb As FraCabCompra = docFactura.Cabecera

        Dim ids(CType(fraCabAlb, FraCabCompraAlbaran).Lineas.Length - 1) As Object
        For i As Integer = 0 To ids.Length - 1
            ids(i) = CType(fraCabAlb, FraCabCompraAlbaran).Lineas(i).IDLineaAlbaran
        Next

        Dim oFltr As New Filter
        oFltr.Add(New InListFilterItem("IDlineaAlbaran", ids, FilterType.Numeric))
        oFltr.Add(New NumberFilterItem("EstadoFactura", FilterOperator.NotEqual, enumaclEstadoFactura.aclFacturado))

        Return acl.Filter(oFltr)
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
        row("QInterna") = 0
        row("Cantidad") = 0
    End Sub

    <Task()> Public Shared Sub AsignarValoresPredeterminadosLineaObligatorios(ByVal row As DataRow, ByVal services As ServiceProvider)
        row("IdLineaFactura") = AdminData.GetAutoNumeric
        row("IDUdMedida") = "UND"
        row("Cantidad") = 0
        row("Precio") = 0
        row("PrecioA") = 0
        row("PrecioB") = 0
        row("UdValoracion") = 1
        row("Dto1") = 0
        row("Dto2") = 0
        row("Dto3") = 0
        row("Importe") = 0
        row("ImporteA") = 0
        row("ImporteB") = 0
        row("EstadoInmovilizado") = 0
        row("Marca") = 0
        row("IDTipoIVA") = "00"
        row("TipoLineaFactura") = 0
        row("IDCentroGestion") = "008"
        row("IDUdInterna") = "UND"
        row("Factor") = 1
        row("QInterna") = 0
        row("Dto") = 0
        row("DtoProntoPago") = 0
        row("NoDeclarar") = 0

    End Sub


#End Region

#Region "Resultado a mostrar en la pantalla intermedia"
    '  Guardamos la información para visualizar la pantalla intermedia de facturas a generar
    <Task()> Public Shared Function ResultadoPropuesta(ByVal data As Object, ByVal services As ServiceProvider) As ResultFacturacion
        'Elimina la información almacenada en memoria si previamente hemos cancelado la facturación
        AdminData.GetSessionData("__frax__")
        'Guardamos la información del documento en memoria, para recuperarla cuando volvamos del preview de presentación
        AdminData.SetSessionData("__frax__", services.GetService(Of ArrayList))

        Return services.GetService(Of ResultFacturacion)()

    End Function
#End Region

#Region "Analítica"

    <Task()> Public Shared Sub CopiarAnalitica(ByVal Doc As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        If Doc.HeaderRow("Estado") <> enumfccEstado.fccContabilizado Then
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
            If IDOrigen.Length > 0 Then
                Dim dtAnaliticaOrigen As DataTable = New AlbaranCompraAnalitica().Filter(New InListFilterItem("IDLineaAlbaran", IDOrigen, FilterType.Numeric))
                Dim datosCopia As New NegocioGeneral.DataCopiarAnalitica(dtAnaliticaOrigen, Doc)
                ProcessServer.ExecuteTask(Of NegocioGeneral.DataCopiarAnalitica)(AddressOf NegocioGeneral.CopiarAnalitica, datosCopia, services)
            End If
        End If
    End Sub
    'David Velasco 28/7/22 Pago desde pisos
    <Task()> Public Shared Sub CentroCosteEnAnalitica(ByVal Doc As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        Try
            'OBTENGO EL NOBRA DEL CONTRATO QUE SERÁ EL CENTRO DE COSTE
            Dim idpisopagos As String = Doc.HeaderRow("IDPisoPagos")
            Dim dt As New DataTable
            Dim f As New Filter
            f.Add("IDPisoPago", FilterOperator.Equal, idpisopagos)
            dt = New BE.DataEngine().Filter("tbPisosPagos", f)

            Dim idcontrato As String = dt(0)("IDContrato")

            Dim dt2 As New DataTable
            Dim f2 As New Filter
            f2.Add("IDContrato", FilterOperator.Equal, idcontrato)
            dt2 = New BE.DataEngine().Filter("tbPisosContrato", f2)
            Dim idobra As String
            idobra = dt2(0)("IDObra")

            Dim dt3 As New DataTable
            Dim f3 As New Filter
            f3.Add("IDObra", FilterOperator.Equal, idobra)
            dt3 = New BE.DataEngine().Filter("tbObraCabecera", f3)
            Dim nobra As String
            nobra = dt3(0)("NObra")
            'OBTENGO LA ESTRUCTURA DE LA TABLA CENTRO DE COSTE 
            Dim dtAnaliticaOrigen As DataTable = New BE.DataEngine().Filter("tbFacturaCompraAnalitica", New NoRowsFilterItem)
            Dim datosCopia As New NegocioGeneral.DataCopiarAnalitica(dtAnaliticaOrigen, Doc)

            'ProcessServer.ExecuteTask(Of NegocioGeneral.DataCopiarAnalitica)(AddressOf NegocioGeneral.RellenaAnaliticaFactura, datosCopia, services)
            'Creo las lineas del centro de coste
            Dim pkDestino As String
            pkDestino = "IDLineaFactura"
            For Each linea As DataRow In Doc.dtLineas.Select()
                Dim UltimaRow As DataRow
                Dim Acum As Double = 0 : Dim AcumA As Double = 0 : Dim AcumB As Double = 0
                Dim HayAnalitica As Boolean = False

                Dim lineaAnalitica As DataRow
                Dim drNewLine As DataRow = Doc.dtAnalitica.NewRow
                drNewLine(pkDestino) = linea(pkDestino)
                drNewLine("IDCentroCoste") = nobra
                drNewLine("Porcentaje") = "100"
                drNewLine("Importe") = linea("Importe")
                Dim ValAyB As New ValoresAyB(New DataRowPropertyAccessor(drNewLine), Doc.IDMoneda, Doc.CambioA, Doc.CambioB)
                ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)
                Acum += drNewLine("Importe")
                AcumA += drNewLine("ImporteA")
                AcumB += drNewLine("ImporteB")
                UltimaRow = drNewLine
                Doc.dtAnalitica.Rows.Add(drNewLine)
                HayAnalitica = True
                If HayAnalitica Then
                    If Acum <> linea("Importe") Then
                        UltimaRow("Importe") += linea("Importe") - Acum
                    End If
                    If AcumA <> linea("ImporteA") Then
                        UltimaRow("ImporteA") += linea("ImporteA") - AcumA
                    End If
                    If AcumB <> linea("ImporteB") Then
                        UltimaRow("ImporteB") += linea("ImporteB") - AcumB
                    End If
                End If
            Next

        Catch ex As Exception

        End Try

    End Sub


#End Region

#Region "Calcular factura"

    <Task()> Public Shared Sub CalcularBasesImponibles(ByVal oDocFra As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        If oDocFra.HeaderRow("Estado") = enumfccEstado.fccNoContabilizado Then
            If Not IsNothing(oDocFra.dtLineas) Then
                Dim desglose() As DataBaseImponible = ProcessServer.ExecuteTask(Of DocumentoFacturaCompra, DataBaseImponible())(AddressOf ProcesoComunes.DesglosarImporte, oDocFra, services)
                Dim datosCalculo As New ProcesoComunes.DataCalculoTotalesCab
                datosCalculo.Doc = oDocFra
                datosCalculo.BasesImponibles = desglose
                ProcessServer.ExecuteTask(Of ProcesoComunes.DataCalculoTotalesCab)(AddressOf CalcularTotalesCabecera, datosCalculo, services)

            End If
        End If
    End Sub

    <Task()> Public Shared Sub CalcularTotalesCabecera(ByVal data As ProcesoComunes.DataCalculoTotalesCab, ByVal services As ServiceProvider)

        If Not IsNothing(data.Doc.HeaderRow) Then
            Dim Bases As DataTable = CType(data.Doc, DocumentoFacturaCompra).dtFCBI
            Bases.DefaultView.Sort = "IDTipoIva"
            Dim IvaManual As Boolean = Nz(data.Doc.HeaderRow("IVAManual"), False)

            If Not IvaManual Then
                Dim notDeleted() As DataRow = Bases.Select(Nothing, Nothing, DataViewRowState.Added Or DataViewRowState.ModifiedCurrent Or DataViewRowState.Unchanged)
                For Each r As DataRow In notDeleted
                    r.Delete()
                Next
            End If

            Dim ImporteLineas As Double
            Dim BaseImponibleTotal As Double
            Dim ImporteIVATotal As Double
            Dim ImporteRETotal As Double
            Dim ImporteIntrastatTotal As Double
            Dim ImporteSinRepercutirTotal As Double


            Dim ImporteLineasA As Double
            Dim BaseImponibleTotalA As Double
            Dim ImporteIVATotalA As Double
            Dim ImporteRETotalA As Double
            Dim ImporteIntrastatTotalA As Double
            Dim ImporteSinRepercutirTotalA As Double

            Dim ImporteLineasB As Double
            Dim BaseImponibleTotalB As Double
            Dim ImporteIVATotalB As Double
            Dim ImporteRETotalB As Double
            Dim ImporteIntrastatTotalB As Double
            Dim ImporteSinRepercutirTotalB As Double


            If Not IsNothing(data.BasesImponibles) AndAlso data.BasesImponibles.Length > 0 Then
                Dim AppParams As ParametroCompra = services.GetService(Of ParametroCompra)()
                Dim IVA As New TipoIva
                Dim AplicarRE As Boolean = AppParams.EmpresaConRecargoEquivalencia
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

                        Dim TiposIVA As EntityInfoCache(Of TipoIvaInfo) = services.GetService(Of EntityInfoCache(Of TipoIvaInfo))()
                        ' HistoricoTipoIVA
                        Dim TIVAInfo As TipoIvaInfo = TiposIVA.GetEntity(lineaBase("IDTipoIva"), data.Doc.HeaderRow("SuFechaFactura"))
                        If data.Doc.AIva Then

                            If AddNew Then
                                'valor por defecto
                                factor = TIVAInfo.Factor
                                If TIVAInfo.SinRepercutir Then
                                    'Nuevo para los ivas especiales que no se repercuten
                                    factor = TIVAInfo.IVASinRepercutir
                                End If
                            End If
                            If Not IvaManual Or (IvaManual And AddNew) Then
                                If TIVAInfo.SinRepercutir Then
                                    lineaBase("ImpSinRepercutir") = xRound(Base * factor / 100, data.Doc.Moneda.NDecimalesImporte)
                                    lineaBase("ImpSinRepercutirA") = xRound(BaseA * factor / 100, data.Doc.MonedaA.NDecimalesImporte)
                                    lineaBase("ImpSinRepercutirB") = xRound(BaseB * factor / 100, data.Doc.MonedaB.NDecimalesImporte)
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
                                If TIVAInfo.SinRepercutir Then
                                    lineaBase("ImpSinRepercutir") = xRound(lineaBase("ImpSinRepercutir"), data.Doc.Moneda.NDecimalesImporte)
                                    lineaBase("ImpSinRepercutirA") = xRound(lineaBase("ImpSinRepercutir") * data.Doc.CambioA, data.Doc.MonedaA.NDecimalesImporte)
                                    lineaBase("ImpSinRepercutirB") = xRound(lineaBase("ImpSinRepercutir") * data.Doc.CambioB, data.Doc.MonedaB.NDecimalesImporte)
                                Else
                                    lineaBase("ImpIVA") = xRound(lineaBase("ImpIVA"), data.Doc.Moneda.NDecimalesImporte)
                                    lineaBase("ImpIVAA") = xRound(lineaBase("ImpIVA") * data.Doc.CambioA, data.Doc.MonedaA.NDecimalesImporte)
                                    lineaBase("ImpIVAB") = xRound(lineaBase("ImpIVA") * data.Doc.CambioB, data.Doc.MonedaB.NDecimalesImporte)
                                    'Gestión cambio base imponible para ajustar facturas 
                                    lineaBase("BaseImponible") = xRound(lineaBase("BaseImponible"), data.Doc.Moneda.NDecimalesImporte)
                                    lineaBase("BaseImponibleA") = xRound(lineaBase("BaseImponible") * data.Doc.CambioA, data.Doc.MonedaA.NDecimalesImporte)
                                    lineaBase("BaseImponibleB") = xRound(lineaBase("BaseImponible") * data.Doc.CambioB, data.Doc.MonedaB.NDecimalesImporte)
                                    'End
                                End If

                                If AplicarRE Then
                                    lineaBase("ImpRE") = xRound(lineaBase("ImpRE"), data.Doc.Moneda.NDecimalesImporte)
                                    lineaBase("ImpREA") = xRound(lineaBase("ImpRE") * data.Doc.CambioA, data.Doc.MonedaA.NDecimalesImporte)
                                    lineaBase("ImpREB") = xRound(lineaBase("ImpRE") * data.Doc.CambioB, data.Doc.MonedaB.NDecimalesImporte)
                                Else
                                    lineaBase("ImpRE") = 0
                                    lineaBase("ImpREA") = 0
                                    lineaBase("ImpREB") = 0
                                End If
                                lineaBase("ImpIntrastat") = xRound(lineaBase("ImpIntrastat"), data.Doc.Moneda.NDecimalesImporte)
                                lineaBase("ImpIntrastatA") = xRound(lineaBase("ImpIntrastat") * data.Doc.CambioA, data.Doc.MonedaA.NDecimalesImporte)
                                lineaBase("ImpIntrastatB") = xRound(lineaBase("ImpIntrastat") * data.Doc.CambioB, data.Doc.MonedaA.NDecimalesImporte)
                            End If
                        End If
                        ' If TIVAInfo.SinRepercutir AndAlso IvaManual Then 
                        ' Cambiado para el caso de un servicio de un extranjero sujeto a inversión de Sujeto pasivo con parte no deducible
                        If TIVAInfo.SinRepercutir And Not IvaManual Then
                            lineaBase("ImpIVANoDeducible") = 0
                            lineaBase("ImpIVANoDeducibleA") = 0
                            lineaBase("ImpIVANoDeducibleB") = 0
                        Else
                            If AddNew Then
                                lineaBase("ImpIVANoDeducible") = xRound(bi.ImporteIVANoDeducible * factor / 100, data.Doc.Moneda.NDecimalesImporte)
                                lineaBase("ImpIVANoDeducibleA") = xRound(bi.ImporteIVANoDeducibleA * factor / 100, data.Doc.MonedaA.NDecimalesImporte)
                                lineaBase("ImpIVANoDeducibleB") = xRound(bi.ImporteIVANoDeducibleB * factor / 100, data.Doc.MonedaB.NDecimalesImporte)
                            ElseIf IvaManual AndAlso Not AddNew Then
                                lineaBase("ImpIVANoDeducibleA") = xRound(lineaBase("ImpIVANoDeducible") * data.Doc.CambioA, data.Doc.MonedaA.NDecimalesImporte)
                                lineaBase("ImpIVANoDeducibleB") = xRound(lineaBase("ImpIVANoDeducible") * data.Doc.CambioB, data.Doc.MonedaB.NDecimalesImporte)
                            End If
                        End If

                        'Dim ValAyB As New ValoresAyB(New DataRowPropertyAccessor(lineaBase), data.Doc.IDMoneda, data.Doc.CambioA, data.Doc.CambioB)
                        'ProcessServer.ExecuteTask(Of ValoresAyB)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)

                        BaseImponibleTotal = BaseImponibleTotal + Nz(lineaBase("BaseImponible"), 0)
                        BaseImponibleTotalA = BaseImponibleTotalA + Nz(lineaBase("BaseImponibleA"), 0)
                        BaseImponibleTotalB = BaseImponibleTotalB + Nz(lineaBase("BaseImponibleB"), 0)

                        ImporteIVATotal = ImporteIVATotal + Nz(lineaBase("ImpIVA"), 0)
                        ImporteIVATotalA = ImporteIVATotalA + Nz(lineaBase("ImpIVAA"), 0)
                        ImporteIVATotalB = ImporteIVATotalB + Nz(lineaBase("ImpIVAB"), 0)

                        ImporteSinRepercutirTotal = ImporteSinRepercutirTotal + Nz(lineaBase("ImpSinRepercutir"), 0)
                        ImporteSinRepercutirTotalA = ImporteSinRepercutirTotalA + Nz(lineaBase("ImpSinRepercutirA"), 0)
                        ImporteSinRepercutirTotalB = ImporteSinRepercutirTotalB + Nz(lineaBase("ImpSinRepercutirB"), 0)

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

            data.Doc.HeaderRow("ImpSinRepercutir") = ImporteSinRepercutirTotal
            data.Doc.HeaderRow("ImpSinRepercutirA") = ImporteSinRepercutirTotalA
            data.Doc.HeaderRow("ImpSinRepercutirB") = ImporteSinRepercutirTotalB

            data.Doc.HeaderRow("ImpRE") = ImporteRETotal
            data.Doc.HeaderRow("ImpREA") = ImporteRETotalA
            data.Doc.HeaderRow("ImpREB") = ImporteRETotalB

            data.Doc.HeaderRow("ImpIntrastat") = ImporteIntrastatTotal
            data.Doc.HeaderRow("ImpIntrastatA") = ImporteIntrastatTotalA
            data.Doc.HeaderRow("ImpIntrastatB") = ImporteIntrastatTotalB

            data.Doc.HeaderRow("ImpLineas") = ImporteLineas
            data.Doc.HeaderRow("ImpLineasA") = ImporteLineasA
            data.Doc.HeaderRow("ImpLineasB") = ImporteLineasB

            Dim ImporteImpuestosTotal As Double = 0
            Dim ImporteImpuestosTotalA As Double = 0
            Dim ImporteImpuestosTotalB As Double = 0
            Dim dtImpuestos As DataTable = CType(data.Doc, DocumentoFacturaCompra).dtImpuestos
            If dtImpuestos.Rows.Count > 0 Then
                'For Each drImpuesto As DataRow In dtImpuestos.Rows
                '    If drImpuesto.RowState <> DataRowState.Deleted Then
                '        If drImpuesto.IsNull("IDLineaImpuesto") Then drImpuesto("IDLineaImpuesto") = AdminData.GetAutoNumeric
                '    End If
                'Next
                ImporteImpuestosTotal += Nz(dtImpuestos.Compute("SUM(Importe)", Nothing), 0)
                ImporteImpuestosTotalA += Nz(dtImpuestos.Compute("SUM(ImporteA)", Nothing), 0)
                ImporteImpuestosTotalB += Nz(dtImpuestos.Compute("SUM(ImporteB)", Nothing), 0)
            End If
      
            data.Doc.HeaderRow("ImpImpuestos") = ImporteImpuestosTotal
            data.Doc.HeaderRow("ImpImpuestosA") = ImporteImpuestosTotalA
            data.Doc.HeaderRow("ImpImpuestosB") = ImporteImpuestosTotalB

        End If
    End Sub

    <Task()> Public Shared Sub CalcularTotales(ByVal oCabFra As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        If oCabFra.HeaderRow("Estado") = enumfccEstado.fccNoContabilizado Then
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
                        Dim ImpLineasNormales As Double = 0
                        If Not oCabFra.dtLineas Is Nothing AndAlso oCabFra.dtLineas.Rows.Count > 0 Then
                            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                            For Each linea As DataRow In oCabFra.dtLineas.Rows
                                Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(linea("IDArticulo"))
                                If Not ArtInfo.Especial Then
                                    ImpLineasNormales += Nz(linea("Importe"), 0)
                                End If
                            Next
                        End If

                        Dim ValAyB As New ValoresAyB(ImpLineasNormales, oCabFra.IDMoneda, oCabFra.CambioA, oCabFra.CambioB)
                        Dim fImpLineasNormales As fImporte = ProcessServer.ExecuteTask(Of ValoresAyB, fImporte)(AddressOf NegocioGeneral.MantenimientoValoresImporteAyB, ValAyB, services)

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
                    If Not Nz(factura("RetencionManual"), False) Then
                        If Not Nz(factura("RegimenEspecial"), False) Then
                            factura("BaseRetencion") = Nz(factura("BaseImponible"), 0)
                            factura("BaseRetencionA") = Nz(factura("BaseImponibleA"), 0)
                            factura("BaseRetencionB") = Nz(factura("BaseImponibleB"), 0)
                        Else
                            factura("BaseRetencion") = Nz(factura("BaseImponible"), 0) + Nz(factura("ImpIVA"), 0) + Nz(factura("ImpRE"), 0) + Nz(factura("ImpImpuestos"), 0)
                            factura("BaseRetencionA") = Nz(factura("BaseImponibleA"), 0) + Nz(factura("ImpIVAA"), 0) + Nz(factura("ImpREA"), 0) + Nz(factura("ImpImpuestosA"), 0)
                            factura("BaseRetencionB") = Nz(factura("BaseImponibleB"), 0) + Nz(factura("ImpIVAB"), 0) + Nz(factura("ImpREB"), 0) + Nz(factura("ImpImpuestosB"), 0)
                        End If
                    Else
                        factura("BaseRetencion") = xRound(factura("BaseRetencion"), oCabFra.Moneda.NDecimalesImporte)
                        factura("BaseRetencionA") = xRound(factura("BaseRetencionA"), oCabFra.MonedaA.NDecimalesImporte)
                        factura("BaseRetencionB") = xRound(factura("BaseRetencionB"), oCabFra.MonedaB.NDecimalesImporte)
                    End If

                    If Length(factura("IDContador")) > 0 Then
                        If ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf NegocioGeneral.ContadorB, factura("IDContador"), services) Then factura("RetencionIRPF") = 0
                    Else
                        'Para importaciones de datos que vienen sin contador y necesitamos mantener el nfactura
                        ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf ProcesoComunes.AsignarContador, oCabFra, services)
                    End If

                    If Nz(factura("RetencionIRPF"), 0) <> 0 Then
                        factura("ImpRetencion") = xRound(factura("BaseRetencion") * factura("RetencionIRPF") / 100, oCabFra.Moneda.NDecimalesImporte)
                        factura("ImpRetencionA") = xRound(factura("BaseRetencionA") * factura("RetencionIRPF") / 100, oCabFra.MonedaA.NDecimalesImporte)
                        factura("ImpRetencionB") = xRound(factura("BaseRetencionB") * factura("RetencionIRPF") / 100, oCabFra.MonedaB.NDecimalesImporte)
                    Else
                        factura("ImpRetencion") = 0
                        factura("ImpRetencionA") = 0
                        factura("ImpRetencionB") = 0
                    End If

                    'If Nz(factura("RetencionIRPF"), 0) <> 0 Then
                    '    factura("ImpRetencion") = xRound(factura("BaseImponible") * factura("RetencionIRPF") / 100, oCabFra.Moneda.NDecimalesImporte)
                    '    factura("ImpRetencionA") = xRound(factura("BaseImponibleA") * factura("RetencionIRPF") / 100, oCabFra.MonedaA.NDecimalesImporte)
                    '    factura("ImpRetencionB") = xRound(factura("BaseImponibleB") * factura("RetencionIRPF") / 100, oCabFra.MonedaB.NDecimalesImporte)
                    'Else
                    '    factura("ImpRetencion") = 0
                    '    factura("ImpRetencionA") = 0
                    '    factura("ImpRetencionB") = 0
                    'End If

                    'If Not Nz(factura("RetencionManual"), False) Then
                    '    If Not Nz(factura("RegimenEspecial"), False) Then
                    '        factura("BaseRetencion") = Nz(factura("BaseImponible"), 0)
                    '        factura("BaseRetencionA") = Nz(factura("BaseImponibleA"), 0)
                    '        factura("BaseRetencionB") = Nz(factura("BaseImponibleB"), 0)
                    '    Else
                    '        factura("BaseRetencion") = Nz(factura("BaseImponible"), 0) + Nz(factura("ImpIVA"), 0) + Nz(factura("ImpRE"), 0) + factura("ImpRetencion")
                    '        factura("BaseRetencionA") = Nz(factura("BaseImponibleA"), 0) + Nz(factura("ImpIVAA"), 0) + Nz(factura("ImpREA"), 0) + factura("ImpRetencionA")
                    '        factura("BaseRetencionB") = Nz(factura("BaseImponibleB"), 0) + Nz(factura("ImpIVAB"), 0) + Nz(factura("ImpREB"), 0) + factura("ImpRetencionB")
                    '    End If
                    'Else
                    '    factura("BaseRetencion") = xRound(factura("BaseRetencion"), oCabFra.Moneda.NDecimalesImporte)
                    '    factura("BaseRetencionA") = xRound(factura("BaseRetencionA"), oCabFra.Moneda.NDecimalesImporte)
                    '    factura("BaseRetencionB") = xRound(factura("BaseRetencionB"), oCabFra.Moneda.NDecimalesImporte)
                    'End If
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
#End Region

#Region "Calcular Vencimientos "


    <Task()> Public Shared Sub CalcularVencimientos(ByVal FraCab As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        If FraCab.HeaderRow("Estado") = enumfccEstado.fccNoContabilizado Then
            If Not FraCab Is Nothing AndAlso Not FraCab.HeaderRow Is Nothing Then
                If Not FraCab.HeaderRow("VencimientosManuales") Then
                    ProcessServer.ExecuteTask(AddressOf PagosAutomaticos, FraCab, services)
                Else
                    Dim dtVtosModif As DataTable = FraCab.dtPagos
                    If Not dtVtosModif Is Nothing AndAlso dtVtosModif.Rows.Count > 0 AndAlso FraCab.dtPagos.Rows.Count > 0 Then
                        Dim dblImpVencimientoA As Double = FraCab.dtPagos.Compute("sum(ImpVencimientoA)", Nothing)
                        Dim dblDifA As Double = FraCab.HeaderRow("ImpTotalA") - dblImpVencimientoA
                        If dblDifA <> 0 Then
                            FraCab.dtPagos.Rows(0)("ImpVencimientoA") = FraCab.dtPagos.Rows(0)("ImpVencimientoA") + dblDifA
                        End If

                        Dim dblImpVencimientoB As Double = FraCab.dtPagos.Compute("sum(ImpVencimientoB)", Nothing)
                        Dim dblDifB As Double = FraCab.HeaderRow("ImpTotalB") - dblImpVencimientoB
                        If dblDifB <> 0 Then
                            FraCab.dtPagos.Rows(0)("ImpVencimientoB") = FraCab.dtPagos.Rows(0)("ImpVencimientoB") + dblDifB
                        End If
                    End If
                End If
            End If
        End If
    End Sub


    Public Class DataNuevosPagos
        Public Doc As DocumentoFacturaCompra
        Public DireccionPago As Integer
        Public TipoPago As Integer
        Public FechaVencimiento As Date
        Public ImporteVencimiento As fImporte
        Public ImporteRecFinanciero As fImporte

        Public Sub New(ByVal Doc As DocumentoFacturaCompra, ByVal DireccionPago As Integer, ByVal TipoPago As Integer)
            Me.Doc = Doc
            Me.DireccionPago = DireccionPago
            Me.TipoPago = TipoPago
        End Sub
    End Class
    <Task()> Public Shared Sub PagosAutomaticos(ByVal oDocFactura As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        Dim ImporteFacturaFinal As Double
        Dim SumaImporteVencimiento As Double : Dim fSumaImporteVencimiento As New fImporte
        Dim SumaImporteVencimientoA As Double
        Dim SumaImporteVencimientoB As Double

        Dim SumaImporteRecFinanciero As Double : Dim fSumaImporteRecFinanciero As New fImporte
        Dim SumaImporteRecFinancieroA As Double
        Dim SumaImporteRecFinancieroB As Double

        Dim fImporteEntregasACta As New fImporte
        '//Se borran todos los vencimientos existentes para la factura.
        For Each pago As DataRow In oDocFactura.dtPagos.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
            '//Borramos los vencimientos que no provienen de Entregas a Cuenta.
            If pago.RowState <> DataRowState.Deleted AndAlso Length(pago("IDEntrega")) = 0 Then
                pago.Delete()
            Else
                fImporteEntregasACta.Importe = fImporteEntregasACta.Importe + Nz(pago("ImpVencimiento"), 0)
            End If
        Next
        ''Para el caso de un abono
        'If oDocFactura.HeaderRow("ImpTotal") - oDocFactura.HeaderRow("ImpRetencion") - oDocFactura.HeaderRow("ImpRetencionGar") < 0 Then
        '    fImporteEntregasACta.Importe = -fImporteEntregasACta.Importe
        'End If

        Dim ValAyB As New ValoresAyB(fImporteEntregasACta.Importe, oDocFactura.IDMoneda, oDocFactura.CambioA, oDocFactura.CambioB)
        fImporteEntregasACta = ProcessServer.ExecuteTask(Of ValoresAyB, fImporte)(AddressOf NegocioGeneral.MantenimientoValoresImporteAyB, ValAyB, services)
        SumaImporteVencimiento = fImporteEntregasACta.Importe
        SumaImporteVencimientoA = fImporteEntregasACta.ImporteA
        SumaImporteVencimientoB = fImporteEntregasACta.ImporteB

        Dim FechaVencimiento As Date
        Dim primero As Boolean = True
        Dim fImporteVencimiento As New fImporte : Dim fImporteRecFinanciero As New fImporte

        Dim Direccion As Integer = ProcessServer.ExecuteTask(Of DocumentoFacturaCompra, Integer)(AddressOf DireccionPago, oDocFactura, services)
        Dim TPago As Integer = ProcessServer.ExecuteTask(Of DocumentoFacturaCompra, Integer)(AddressOf TipoPago, oDocFactura, services)
        Dim DiaPago As String = oDocFactura.HeaderRow("IDDiaPago") & String.Empty

        Dim datNuevosPagos As New DataNuevosPagos(oDocFactura, Direccion, TPago)
        '//Se recorren las líneas de condicion de pago para crear un cobro por cada una de ellas.
        Dim condiciones As DataTable = New CondicionPagoLinea().Filter(New StringFilterItem("IdCondicionPago", oDocFactura.HeaderRow("IDCondicionPago")))
        For Each condicion As DataRow In condiciones.Rows
            Dim ImporteVencimiento As Double = 0
            Dim ImporteVencimientoA As Double = 0
            Dim ImporteVencimientoB As Double = 0

            Dim ImporteRecFinanciero As Double = 0
            Dim ImporteRecFinancieroA As Double = 0
            Dim ImporteRecFinancieroB As Double = 0

            '//Guardamos la fecha de vencimiento más lejana para la cabecera.
            Dim dataVto As New NegocioGeneral.dataCalculoFechaVencimiento(oDocFactura.HeaderRow("suFechaFactura"), condicion("Periodo"), condicion("TipoPeriodo"), DiaPago, False, oDocFactura.HeaderRow("IdProveedor"))
            ProcessServer.ExecuteTask(Of NegocioGeneral.dataCalculoFechaVencimiento)(AddressOf NegocioGeneral.CalcularFechaVencimiento, dataVto, services)
            FechaVencimiento = dataVto.FechaVencimiento
            If primero Then
                primero = False
                oDocFactura.HeaderRow("FechaVencimiento") = FechaVencimiento
            End If

            ImporteVencimiento = xRound(((oDocFactura.HeaderRow("ImpTotal") - oDocFactura.HeaderRow("ImpRetencion") - oDocFactura.HeaderRow("ImpRetencionGar") - fImporteEntregasACta.Importe) * condicion("porcentaje") / 100), oDocFactura.Moneda.NDecimalesImporte)
            ImporteVencimientoA = xRound(((oDocFactura.HeaderRow("ImpTotalA") - oDocFactura.HeaderRow("ImpRetencionA") - oDocFactura.HeaderRow("ImpRetencionGarA") - fImporteEntregasACta.ImporteA) * condicion("porcentaje") / 100), oDocFactura.MonedaA.NDecimalesImporte)
            ImporteVencimientoB = xRound(((oDocFactura.HeaderRow("ImpTotalB") - oDocFactura.HeaderRow("ImpRetencionB") - oDocFactura.HeaderRow("ImpRetencionGarB") - fImporteEntregasACta.ImporteB) * condicion("porcentaje") / 100), oDocFactura.MonedaB.NDecimalesImporte)

            fImporteVencimiento.Importe = ImporteVencimiento
            fImporteVencimiento.ImporteA = ImporteVencimientoA
            fImporteVencimiento.ImporteB = ImporteVencimientoB

            SumaImporteVencimiento = CDec(SumaImporteVencimiento) + CDec(ImporteVencimiento)
            SumaImporteVencimientoA = CDec(SumaImporteVencimientoA) + CDec(ImporteVencimientoA)
            SumaImporteVencimientoB = CDec(SumaImporteVencimientoB) + CDec(ImporteVencimientoB)

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

                SumaImporteRecFinanciero = CDec(SumaImporteRecFinanciero) + CDec(ImporteRecFinanciero)
                SumaImporteRecFinancieroA = CDec(SumaImporteRecFinancieroA) + CDec(ImporteRecFinancieroA)
                SumaImporteRecFinancieroB = CDec(SumaImporteRecFinancieroB) + CDec(ImporteRecFinancieroB)

                fSumaImporteRecFinanciero.Importe = SumaImporteRecFinanciero
                fSumaImporteRecFinanciero.ImporteA = SumaImporteRecFinancieroA
                fSumaImporteRecFinanciero.ImporteB = SumaImporteRecFinancieroB

            End If

            If fImporteVencimiento.Importe <> 0 Then
                datNuevosPagos.FechaVencimiento = FechaVencimiento
                datNuevosPagos.ImporteVencimiento = fImporteVencimiento
                datNuevosPagos.ImporteRecFinanciero = fImporteRecFinanciero

                ProcessServer.ExecuteTask(Of DataNuevosPagos)(AddressOf NuevoPago, datNuevosPagos, services)
            End If
        Next

        '///Generacion de pago por retencion de garantía
        ProcessServer.ExecuteTask(Of DataNuevosPagos)(AddressOf NuevoPagoRetencionGarantia, datNuevosPagos, services)

        '///Generacion de pagos por retencion IRPF
        ProcessServer.ExecuteTask(Of DataNuevosPagos)(AddressOf NuevoPagoRetencionIRPF, datNuevosPagos, services)

        '//Ajuste de los vencimientos
        Dim datAjuste As New DataAjusteVencimientosFC(oDocFactura, fSumaImporteVencimiento, fSumaImporteRecFinanciero, fImporteEntregasACta)
        ProcessServer.ExecuteTask(Of DataAjusteVencimientosFC)(AddressOf AjusteVencimientos, datAjuste, services)
    End Sub

    <Task()> Public Shared Function TipoPago(ByVal oDocFactura As DocumentoFacturaCompra, ByVal services As ServiceProvider) As Integer
        Dim AppParams As ParametroFacturaCompra = services.GetService(Of ParametroFacturaCompra)()
        Dim intTipoPago As Integer = AppParams.TipoPagoFacturaCompra
        If Length(oDocFactura.HeaderRow("IDContador")) > 0 Then
            '//Si es una factura B
            If Not oDocFactura.AIva Then
                intTipoPago = AppParams.TipoPagoFacturaCompraB
            End If
        End If
        Return intTipoPago
    End Function

    <Task()> Public Shared Function DireccionPago(ByVal oDocFactura As DocumentoFacturaCompra, ByVal services As ServiceProvider) As Integer
        Dim Direccion As Integer
        Dim pd As New ProveedorDireccion
        Dim StDatosDirec As New ProveedorDireccion.DataDirecDe
        StDatosDirec.IDDireccion = Nz(oDocFactura.HeaderRow("IDDireccion"), 0)
        StDatosDirec.TipoDireccion = enumpdTipoDireccion.pdDireccionPago
        If Length(oDocFactura.HeaderRow("IDDireccion")) > 0 AndAlso ProcessServer.ExecuteTask(Of ProveedorDireccion.DataDirecDe, Boolean)(AddressOf ProveedorDireccion.EsDireccionDe, StDatosDirec, services) Then
            Direccion = Nz(oDocFactura.HeaderRow("IDDireccion"), 0)
        Else
            Dim StDatosDirecEnv As New ProveedorDireccion.DataDirecEnvio
            StDatosDirecEnv.IDProveedor = oDocFactura.HeaderRow("IDProveedor")
            StDatosDirecEnv.TipoDireccion = enumpdTipoDireccion.pdDireccionPago
            Dim dir As DataTable = ProcessServer.ExecuteTask(Of ProveedorDireccion.DataDirecEnvio, DataTable)(AddressOf ProveedorDireccion.ObtenerDireccionEnvio, StDatosDirecEnv, services)
            If Not IsNothing(dir) AndAlso dir.Rows.Count Then
                Direccion = Nz(dir.Rows(0)("IDDireccion"), 0)
            End If
        End If
        Return Direccion
    End Function

    <Task()> Public Shared Sub NuevoPago(ByVal data As DataNuevosPagos, ByVal services As ServiceProvider)
        If data Is Nothing Then Exit Sub
        If data.ImporteRecFinanciero Is Nothing Then data.ImporteRecFinanciero = New fImporte
        Dim newrow As DataRow = data.Doc.dtPagos.NewRow
        newrow("IDPago") = AdminData.GetAutoNumeric
        If data.DireccionPago <> 0 Then newrow("IDDireccion") = data.DireccionPago
        Dim AppParams As ParametroContabilidadCompra = services.GetService(GetType(ParametroContabilidadCompra))
        If AppParams.Contabilidad Then
            If Length(data.Doc.Proveedor.CCProveedor) = 0 Then ApplicationService.GenerateError("La Cuenta Contable del proveedor es un dato obligatorio.")
            newrow("CContable") = data.Doc.Proveedor.CCProveedor
        End If
        newrow("IDFactura") = data.Doc.HeaderRow("IDFactura")
        newrow("IdTipoPago") = data.TipoPago
        newrow("IdProveedor") = data.Doc.HeaderRow("IdProveedor")
        newrow("IdProveedorBanco") = data.Doc.HeaderRow("IdProveedorBanco")
        newrow("IDBancoPropio") = data.Doc.HeaderRow("IDBancoPropio")
        newrow("FechaVencimientoFactura") = data.FechaVencimiento
        newrow("FechaVencimiento") = data.FechaVencimiento
        newrow("Titulo") = Nz(data.Doc.HeaderRow("RazonSocial"), data.Doc.Proveedor.RazonSocial)
        newrow("NFactura") = data.Doc.HeaderRow("NFactura")
        newrow("IDFormaPago") = data.Doc.HeaderRow("IDFormaPago")
        newrow("IDMoneda") = data.Doc.HeaderRow("IDMoneda")
        newrow("CambioA") = data.Doc.HeaderRow("CambioA")
        newrow("CambioB") = data.Doc.HeaderRow("CambioB")

        newrow("ImpVencimiento") = data.ImporteVencimiento.Importe
        newrow("RecargoFinanciero") = data.ImporteRecFinanciero.Importe
        newrow("ImpVencimientoA") = data.ImporteVencimiento.ImporteA
        newrow("RecargoFinancieroA") = data.ImporteRecFinanciero.ImporteA
        newrow("ImpVencimientoB") = data.ImporteVencimiento.ImporteB
        newrow("RecargoFinancieroB") = data.ImporteRecFinanciero.ImporteB
        newrow("NOperacion") = 0
        newrow("Impreso") = False
        If Length(data.Doc.HeaderRow("IDObra")) <> 0 Then
            newrow("IDObra") = data.Doc.HeaderRow("IDObra")
        End If

        Dim datEstSit As New DataAsignarEstadoSituacionFC(data.Doc.HeaderRow("IDTipoAsiento"), newrow)
        ProcessServer.ExecuteTask(Of DataAsignarEstadoSituacionFC)(AddressOf AsignarEstadoSituacion, datEstSit, services)

        data.Doc.dtPagos.Rows.Add(newrow.ItemArray)
    End Sub

    <Task()> Public Shared Sub NuevoPagoRetencionGarantia(ByVal data As DataNuevosPagos, ByVal services As ServiceProvider)
        If Nz(data.Doc.HeaderRow("Retencion"), 0) <> 0 And Nz(data.Doc.HeaderRow("ImpRetencionGar"), 0) <> 0 Then
            Dim newrow As DataRow = data.Doc.dtPagos.NewRow
            newrow("IDPago") = AdminData.GetAutoNumeric
            If data.DireccionPago <> 0 Then newrow("IDDireccion") = data.DireccionPago
            Dim AppParamsCont As ParametroContabilidadCompra = services.GetService(Of ParametroContabilidadCompra)()
            If AppParamsCont.Contabilidad Then
                If Length(data.Doc.Proveedor.CCRetencion) > 0 Then
                    newrow("CContable") = data.Doc.Proveedor.CCRetencion
                Else
                    If Len(data.Doc.Proveedor.CCProveedor) = 0 Then ApplicationService.GenerateError("La Cuenta Contable del proveedor es un dato obligatorio.")
                    newrow("CContable") = data.Doc.Proveedor.CCProveedor
                End If
            End If
            newrow("IDFactura") = data.Doc.HeaderRow("IDFactura")
            Dim AppParams As ParametroFacturaCompra = services.GetService(Of ParametroFacturaCompra)()
            newrow("IdTipoPago") = AppParams.TipoPagoRetencion
            newrow("IdProveedor") = data.Doc.HeaderRow("IdProveedor")
            newrow("IdProveedorBanco") = data.Doc.HeaderRow("IdProveedorBanco")
            newrow("IDBancoPropio") = data.Doc.HeaderRow("IDBancoPropio")
            If Length(data.Doc.HeaderRow("TipoRetencion")) = 0 Then data.Doc.HeaderRow("TipoRetencion") = enumTipoRetencion.troSobreBI
            newrow("FechaVencimientoFactura") = data.Doc.HeaderRow("FechaRetencion")
            newrow("Permiso") = False
            newrow("FechaVencimiento") = data.Doc.HeaderRow("FechaRetencion")
            newrow("Titulo") = Nz(data.Doc.HeaderRow("RazonSocial"), data.Doc.Proveedor.RazonSocial)
            newrow("NFactura") = data.Doc.HeaderRow("NFactura")
            newrow("IDFormaPago") = data.Doc.HeaderRow("IDFormaPago")
            newrow("IDMoneda") = data.Doc.HeaderRow("IDMoneda")
            newrow("CambioA") = data.Doc.HeaderRow("CambioA")
            newrow("CambioB") = data.Doc.HeaderRow("CambioB")
            Dim ValAyB As New ValoresAyB(CDbl(data.Doc.HeaderRow("ImpRetencionGar")), data.Doc.IDMoneda, data.Doc.CambioA, data.Doc.CambioB)
            Dim fImporteVencimiento As fImporte = ProcessServer.ExecuteTask(Of ValoresAyB, fImporte)(AddressOf NegocioGeneral.MantenimientoValoresImporteAyB, ValAyB, services)
            newrow("ImpVencimiento") = fImporteVencimiento.Importe
            newrow("ImpVencimientoA") = fImporteVencimiento.ImporteA
            newrow("ImpVencimientoB") = fImporteVencimiento.ImporteB
            newrow("NOperacion") = 0
            newrow("Impreso") = False
            newrow("IDObra") = data.Doc.HeaderRow("IDObra")

            Dim datEstSit As New DataAsignarEstadoSituacionFC(data.Doc.HeaderRow("IDTipoAsiento"), newrow)
            ProcessServer.ExecuteTask(Of DataAsignarEstadoSituacionFC)(AddressOf AsignarEstadoSituacion, datEstSit, services)

            data.Doc.dtPagos.Rows.Add(newrow.ItemArray)
        End If
    End Sub

    <Task()> Public Shared Sub NuevoPagoRetencionIRPF(ByVal data As DataNuevosPagos, ByVal services As ServiceProvider)
        If Nz(data.Doc.HeaderRow("RetencionIRPF"), 0) <> 0 And Nz(data.Doc.HeaderRow("BaseImponible"), 0) <> 0 Then
            '//Obtenemos el Proveedor y la CContable de las retenciones
            Dim AppParams As ParametroFacturaCompra = services.GetService(Of ParametroFacturaCompra)()
            Dim strIDProvRetencion As String = AppParams.ProveedorRetencion()
            Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
            Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(strIDProvRetencion)
            If Not ProvInfo Is Nothing Then
                Dim newrow As DataRow = data.Doc.dtPagos.NewRow
                newrow("IDPago") = AdminData.GetAutoNumeric
                If data.DireccionPago <> 0 Then newrow("IDDireccion") = data.DireccionPago
                Dim Info As ParametroContabilidadCompra = services.GetService(Of ParametroContabilidadCompra)()
                If Info.Contabilidad Then newrow("CContable") = ProvInfo.CCProveedor
                newrow("IDFactura") = data.Doc.HeaderRow("IDFactura")
                newrow("IdTipoPago") = data.TipoPago
                newrow("IdProveedor") = strIDProvRetencion
                newrow("IdProveedorBanco") = data.Doc.HeaderRow("IdProveedorBanco")
                newrow("IDBancoPropio") = data.Doc.HeaderRow("IDBancoPropio")
                newrow("FechaVencimientoFactura") = data.Doc.HeaderRow("FechaFactura")
                newrow("FechaVencimiento") = data.Doc.HeaderRow("FechaFactura")
                newrow("Titulo") = ProvInfo.DescProveedor
                newrow("NFactura") = data.Doc.HeaderRow("NFactura")
                newrow("IDFormaPago") = ProvInfo.IDFormaPago
                newrow("IDMoneda") = data.Doc.HeaderRow("IDMoneda")
                newrow("CambioA") = data.Doc.HeaderRow("CambioA")
                newrow("CambioB") = data.Doc.HeaderRow("CambioB")
                Dim ValAyB As New ValoresAyB(CDbl(data.Doc.HeaderRow("ImpRetencion")), data.Doc.IDMoneda, data.Doc.CambioA, data.Doc.CambioB)
                Dim fImpRetencion As fImporte = ProcessServer.ExecuteTask(Of ValoresAyB, fImporte)(AddressOf NegocioGeneral.MantenimientoValoresImporteAyB, ValAyB, services)
                newrow("ImpVencimiento") = fImpRetencion.Importe
                newrow("ImpVencimientoA") = fImpRetencion.ImporteA
                newrow("ImpVencimientoB") = fImpRetencion.ImporteB
                newrow("NOperacion") = 0
                newrow("Impreso") = False
                newrow("IDObra") = data.Doc.HeaderRow("IDObra")
                newrow("Contabilizado") = enumPagoContabilizado.PagoNoContabilizado
                newrow("Situacion") = enumPagoSituacion.NoPagado
                'Dim datEstSit As New DataAsignarEstadoSituacionFC(data.Doc.HeaderRow("IDTipoAsiento"), newrow)
                'ProcessServer.ExecuteTask(Of DataAsignarEstadoSituacionFC)(AddressOf AsignarEstadoSituacion, datEstSit, services)

                data.Doc.dtPagos.Rows.Add(newrow.ItemArray)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub NuevaLineaFacturaObraEntregaCuenta(ByVal Doc As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        If (Doc.HeaderRow.RowState = DataRowState.Added AndAlso Length(Doc.HeaderRow("IDObra")) > 0) OrElse _
           (Doc.HeaderRow.RowState = DataRowState.Modified AndAlso Length(Doc.HeaderRow("IDObra")) > 0 AndAlso Nz(Doc.HeaderRow("IDObra"), 0) <> Nz(Doc.HeaderRow("IDObra", DataRowVersion.Original), 0)) Then
            '//Si tenemos retención por Garantía de una Obra
            If Length(Doc.HeaderRow("TipoRetencion")) > 0 AndAlso Length(Doc.HeaderRow("Retencion")) > 0 AndAlso Length(Doc.HeaderRow("FechaRetencion")) > 0 Then
                If Not IsNothing(Doc.dtLineas) AndAlso Doc.dtLineas.Rows.Count > 0 Then
                    Dim EC As New EntregasACuenta
                    Dim dtEntregas As DataTable = EC.AddNew
                    '//Creamos una Entrega nueva de Tipo Retención.
                    Dim datNuevaEntrega As New EntregasACuenta.DatosNuevaEntrega(Doc.HeaderRow.Table, Doc.dtLineas, dtEntregas, Circuito.Compras)
                    Dim drNuevaEntrega As DataRow = ProcessServer.ExecuteTask(Of EntregasACuenta.DatosNuevaEntrega, DataRow)(AddressOf EntregasACuenta.NuevaEntregaTipoRetencionFacturaObra, datNuevaEntrega, services)
                    If Not drNuevaEntrega Is Nothing Then
                        EC.Update(datNuevaEntrega.DtEntregas)
                        '//Creamos una nueva línea de factura de tipo Retención
                        Dim datFactEntrCta As New EntregasACuenta.DataFacturaCompraEntregas(Doc, datNuevaEntrega.DtEntregas)
                        ProcessServer.ExecuteTask(Of EntregasACuenta.DataFacturaCompraEntregas)(AddressOf EntregasACuenta.AddEntregasTipoFacturaCompras, datFactEntrCta, services)
                    End If
                End If
            End If

        End If
    End Sub

    Public Class DataAsignarEstadoSituacionFC
        Public TipoAsiento As enumTipoAsiento
        Public NewRow As DataRow

        Public Sub New(ByVal TipoAsiento As enumTipoAsiento, ByVal NewRow As DataRow)
            Me.TipoAsiento = TipoAsiento
            Me.NewRow = NewRow
        End Sub
    End Class
    <Task()> Public Shared Sub AsignarEstadoSituacion(ByVal data As DataAsignarEstadoSituacionFC, ByVal services As ServiceProvider)
        Select Case data.TipoAsiento
            Case enumTipoAsiento.taBancoSinPago
                data.NewRow("Contabilizado") = enumPagoContabilizado.PagoContabilizado
                data.NewRow("Situacion") = enumPagoSituacion.Pagado
            Case enumTipoAsiento.taProveedorConPagoNPyNC
                data.NewRow("Contabilizado") = enumPagoContabilizado.PagoNoContabilizado
                data.NewRow("Situacion") = enumPagoSituacion.NoPagado
            Case enumTipoAsiento.taProveedorConPagoPyNC
                data.NewRow("Contabilizado") = enumPagoContabilizado.PagoNoContabilizado
                data.NewRow("Situacion") = enumPagoSituacion.Pagado
            Case enumTipoAsiento.taProveedorSinPago
                data.NewRow("Contabilizado") = enumPagoContabilizado.PagoContabilizado
                data.NewRow("Situacion") = enumPagoSituacion.Pagado
        End Select
    End Sub

    Public Class DataAjusteVencimientosFC
        Public Doc As DocumentoFacturaCompra
        Public SumaImporteVencimiento As fImporte
        Public SumaImporteRecFinanciero As fImporte
        Public SumaEntregasACuenta As fImporte

        Public Sub New(ByVal Doc As DocumentoFacturaCompra, ByVal SumaImporteVencimiento As fImporte, ByVal SumaImporteRecFinanciero As fImporte, ByVal SumaEntregasACuenta As fImporte)
            Me.Doc = Doc
            Me.SumaImporteVencimiento = SumaImporteVencimiento
            Me.SumaImporteRecFinanciero = SumaImporteRecFinanciero
            Me.SumaEntregasACuenta = SumaEntregasACuenta
        End Sub
    End Class
    <Task()> Public Shared Sub AjusteVencimientos(ByVal data As DataAjusteVencimientosFC, ByVal services As ServiceProvider)
        Dim AddedRows As DataTable = data.Doc.dtPagos.GetChanges(DataRowState.Added)
        If Not IsNothing(AddedRows) Then
            If AddedRows.Rows.Count Then
                Dim VtoAAjustar As DataRow
                If data.Doc.HeaderRow("ImpRetencionGar") = 0 Then
                    VtoAAjustar = data.Doc.dtPagos.Rows(data.Doc.dtPagos.Rows.Count - 1)
                Else
                    VtoAAjustar = data.Doc.dtPagos.Rows(data.Doc.dtPagos.Rows.Count - 2)
                End If

                If (data.SumaImporteVencimiento.Importe - (Nz(data.Doc.HeaderRow("ImpTotal"), 0) - Nz(data.Doc.HeaderRow("ImpRetencion"), 0)) - Nz(data.Doc.HeaderRow("ImpRetencionGar"), 0)) <> 0 Then
                    Dim ImporteVencimiento As Double = Nz(VtoAAjustar("ImpVencimiento"), 0)
                    VtoAAjustar("ImpVencimiento") = ImporteVencimiento + Nz(data.Doc.HeaderRow("ImpTotal"), 0) - Nz(data.Doc.HeaderRow("ImpRetencion"), 0) - Nz(data.Doc.HeaderRow("ImpRetencionGar"), 0) - data.SumaImporteVencimiento.Importe + data.SumaEntregasACuenta.Importe
                End If
                If (data.SumaImporteVencimiento.ImporteA - (Nz(data.Doc.HeaderRow("ImpTotalA"), 0) - Nz(data.Doc.HeaderRow("ImpRetencionA"), 0) - Nz(data.Doc.HeaderRow("ImpRetencionGarA"), 0))) <> 0 Then
                    Dim ImporteVencimientoA As Double = Nz(VtoAAjustar("ImpVencimientoA"), 0)
                    VtoAAjustar("ImpVencimientoA") = ImporteVencimientoA + Nz(data.Doc.HeaderRow("ImpTotalA"), 0) - Nz(data.Doc.HeaderRow("ImpRetencionA"), 0) - Nz(data.Doc.HeaderRow("ImpRetencionGarA"), 0) - data.SumaImporteVencimiento.ImporteA + data.SumaEntregasACuenta.ImporteA
                End If
                If (data.SumaImporteVencimiento.ImporteB - (Nz(data.Doc.HeaderRow("ImpTotalB"), 0) - Nz(data.Doc.HeaderRow("ImpRetencionB"), 0) - Nz(data.Doc.HeaderRow("ImpRetencionGarB"), 0))) <> 0 Then
                    Dim ImporteVencimientoB As Double = Nz(VtoAAjustar("ImpVencimientoB"), 0)
                    VtoAAjustar("ImpVencimientoB") = ImporteVencimientoB + Nz(data.Doc.HeaderRow("ImpTotalB"), 0) - Nz(data.Doc.HeaderRow("ImpRetencionB"), 0) - Nz(data.Doc.HeaderRow("ImpRetencionGarB"), 0) - data.SumaImporteVencimiento.ImporteB + data.SumaEntregasACuenta.ImporteB
                End If

                If Nz(data.Doc.HeaderRow("RecFinan"), 0) > 0 Then
                    If (data.SumaImporteRecFinanciero.Importe - Nz(data.Doc.HeaderRow("ImpRecFinan"), 0)) <> 0 Then
                        Dim ImporteRecFinanciero As Double = Nz(VtoAAjustar("RecargoFinanciero"), 0)
                        VtoAAjustar("RecargoFinanciero") = ImporteRecFinanciero + Nz(data.Doc.HeaderRow("ImpRecFinan"), 0) - data.SumaImporteRecFinanciero.Importe
                    End If
                    If (data.SumaImporteRecFinanciero.ImporteA - Nz(data.Doc.HeaderRow("ImpRecFinanA"), 0)) <> 0 Then
                        Dim ImporteRecFinancieroA As Double = Nz(VtoAAjustar("RecargoFinancieroA"), 0)
                        VtoAAjustar("RecargoFinancieroA") = ImporteRecFinancieroA + Nz(data.Doc.HeaderRow("ImpRecFinanA"), 0) - data.SumaImporteRecFinanciero.ImporteA
                    End If
                    If (data.SumaImporteRecFinanciero.ImporteB - Nz(data.Doc.HeaderRow("ImpRecFinanB"), 0)) <> 0 Then
                        Dim ImporteRecFinancieroB As Double = Nz(VtoAAjustar("RecargoFinancieroB"), 0)
                        VtoAAjustar("RecargoFinancieroB") = ImporteRecFinancieroB + Nz(data.Doc.HeaderRow("ImpRecFinanB"), 0) - data.SumaImporteRecFinanciero.ImporteB
                    End If
                End If

            End If
        End If
    End Sub

#End Region

#Region "Validaciones"
    <Task()> Public Shared Sub ValidarDocumento(ByVal Doc As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        Dim FCC As New FacturaCompraCabecera
        FCC.Validate(Doc.HeaderRow.Table)

        Dim FCL As New FacturaCompraLinea
        FCL.Validate(Doc.dtLineas)
    End Sub

    <Task()> Public Shared Sub ValidarFacturaContabilizada(ByVal data As DataRow, ByVal services As ServiceProvider)
        'If data("Estado") = enumfccEstado.fccContabilizado Then
        '    If New Parametro().Contabilidad Then
        '        ApplicationService.GenerateError("La Factura está Contabilizada.")
        '    Else : ApplicationService.GenerateError("La Factura está Bloqueada y generado los vencimientos (o efectos)")
        '    End If
        'End If
    End Sub
    <Task()> Public Shared Sub ValidarFacturaDeclarada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("AñoDeclaracionIva")) > 0 AndAlso Length(data("NDeclaracionIva")) > 0 Then
            ApplicationService.GenerateError("La Factura está Declarada.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarFechaRetencion(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Nz(data("Retencion"), 0) <> 0 AndAlso Length(data("FechaRetencion")) = 0 Then ApplicationService.GenerateError("La Fecha de Rentención es obligatoria si introduce un Porcentaje de Retención.")
    End Sub

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

    '        Dim dtFCC As DataTable = New FacturaCompraCabecera().Filter(f)
    '        If Not dtFCC Is Nothing AndAlso dtFCC.Rows.Count > 0 Then
    '            'If AppParamsConta.Contabilidad Then
    '            '    ApplicationService.GenerateError("La Factura {0} ya existe para el Ejercicio {1}.", Quoted(data("NFactura")), Quoted(data("IDEjercicio")))
    '            'Else
    '            '    ApplicationService.GenerateError("La Factura {0} ya existe.", Quoted(data("NFactura")))
    '            'End If
    '            ApplicationService.GenerateError("La Factura {0} ya existe para el año {1}.", Quoted(data("NFactura")), Quoted(Year(data("FechaFactura"))))
    '        End If
    '    End If
    'End Sub

    <Task()> Public Shared Sub ValidarSuFactura(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("SuFactura")) > 0 Then
            If data.RowState = DataRowState.Added OrElse (data.RowState = DataRowState.Modified AndAlso data("SuFactura", DataRowVersion.Original) & String.Empty <> data("SuFactura") OrElse data("IDProveedor", DataRowVersion.Original) & String.Empty <> data("IDProveedor")) Then
                Dim f As New Filter
                f.Add(New StringFilterItem("SuFactura", data("SuFactura")))
                f.Add(New StringFilterItem("IDProveedor", data("IDProveedor")))
                Dim AppParamsConta As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
                If AppParamsConta.Contabilidad AndAlso Length(data("IDEjercicio")) > 0 Then
                    f.Add(New StringFilterItem("IDEjercicio", data("IDEjercicio")))
                Else
                    f.Add(New StringFilterItem("YEAR(FechaFactura)", CType(data("FechaFactura"), Date).Year))
                End If
                Dim dtFCC As DataTable = New FacturaCompraCabecera().Filter(f)
                If Not dtFCC Is Nothing AndAlso dtFCC.Rows.Count > 0 Then
                    If AppParamsConta.Contabilidad Then
                        ApplicationService.GenerateError("Ya existe una Factura con este SuFactura {0} para el Ejercicio {1}.", Quoted(data("SuFactura")), Quoted(data("IDEjercicio")))
                    Else
                        ApplicationService.GenerateError("Ya existe una Factura con este SuFactura {0}.", Quoted(data("SuFactura")))
                    End If
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarConceptosGastosObra(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDConcepto")) > 0 Then
            Dim strFrom As String : Dim strEntidad As String
            Dim FilWhere As New Filter
            If data("TipoGastoObra") = enumfclTipoGastoObra.enumfclGastos Then
                FilWhere.Add("IDGasto", FilterOperator.Equal, data("IDConcepto"))
                strFrom = "tbMaestroGasto"
                strEntidad = "Gastos"
            ElseIf data("TipoGastoObra") = enumfclTipoGastoObra.enumfclVarios Then
                FilWhere.Add("IDVarios", FilterOperator.Equal, data("IDConcepto"))
                strFrom = "tbMaestroVarios"
                strEntidad = "Varios"
            Else
                strFrom = String.Empty
            End If
            If Len(strFrom) > 0 Then
                Dim dt As DataTable = New BE.DataEngine().Filter(strFrom, FilWhere)
                If IsNothing(dt) OrElse dt.Rows.Count = 0 Then
                    ApplicationService.GenerateError("El Concepto introducido no existe en {0}.", Quoted(strEntidad))
                End If
            End If
        End If
    End Sub

#End Region

#Region " Grabar Documento "

    <Task()> Public Shared Sub GrabarDocumento(ByVal Doc As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        'AdminData.SetData(Doc.HeaderRow.Table, False)
        'AdminData.SetData(Doc.dtLineas, False)
        'AdminData.SetData(Doc.dtFCBI, False)
        'AdminData.SetData(Doc.dtPagos, False)
        'AdminData.SetData(Doc.dtAnalitica, False)
        ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf Business.General.Comunes.UpdateDocument, Doc, services)
    End Sub

#End Region

#Region "Métodos Borrado"

    <Task()> Public Shared Sub ActualizarEntregasACuenta(ByVal DocHeaderRow As DataRow, ByVal services As ServiceProvider)
        Dim StDatos As New EntregasACuenta.DatosElimRestricEnt
        StDatos.IDFactura = DocHeaderRow("IDFactura")
        StDatos.Circuito = Circuito.Compras
        ProcessServer.ExecuteTask(Of EntregasACuenta.DatosElimRestricEnt)(AddressOf EntregasACuenta.EliminarRestriccionesDeleteEntregaCuenta, StDatos, services)
    End Sub

#End Region

#Region " Actualización de Albaranes "

    <Task()> Public Shared Sub ActualizarAlbaran(ByVal DocFra As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        If DocFra.HeaderRow("Estado") = enumfccEstado.fccNoContabilizado Then
            ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf ActualizarQFacturadaAlbaran, DocFra, services)
            ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf ActualizarImportesAlbaran, DocFra, services)
            ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf GrabarAlbaranes, DocFra, services)
            'TODO Realquiler (Acción para actualizar lineas de albarán normal o de realquiler)
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarQFacturadaAlbaran(ByVal DocFra As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        If DocFra Is Nothing Then Exit Sub
        For Each lineaFactura As DataRow In DocFra.dtLineas.Select(Nothing, "IDAlbaran,IDLineaAlbaran")
            ActualizarQFacturadaLineaAlbaran(lineaFactura, services)
        Next
    End Sub
    <Task()> Public Shared Sub ActualizarQFacturadaAlbaranEnProceso(ByVal DocFra As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        If DocFra Is Nothing Then Exit Sub
        For Each lineaFactura As DataRow In DocFra.dtLineas.Select(Nothing, "IDAlbaran,IDLineaAlbaran")
            ActualizarQFacturadaLineaAlbaran(lineaFactura, services)
        Next

        ProcessServer.ExecuteTask(Of Object)(AddressOf GrabarAlbaranes, Nothing, services)
    End Sub
    <Task()> Public Shared Sub ActualizarQFacturadaLineaAlbaran(ByVal lineaFactura As DataRow, ByVal services As ServiceProvider)
        If Length(lineaFactura("IDAlbaran")) > 0 AndAlso Length(lineaFactura("IDLineaAlbaran")) > 0 Then
            If lineaFactura.RowState <> DataRowState.Modified OrElse lineaFactura("cantidad") <> lineaFactura("cantidad", DataRowVersion.Original) Then
                Dim Albaranes As DocumentInfoCache(Of DocumentoAlbaranCompra) = services.GetService(Of DocumentInfoCache(Of DocumentoAlbaranCompra))()
                Dim DocAlb As DocumentoAlbaranCompra = Albaranes.GetDocument(lineaFactura("IDAlbaran"))

                Dim OriginalQFacturada As Double
                Dim ProposedQFacturada As Double = Nz(lineaFactura("cantidad"), 0)
                If lineaFactura.RowState = DataRowState.Modified Then
                    OriginalQFacturada = lineaFactura("cantidad", DataRowVersion.Original)
                End If
                DocAlb.SetQFacturada(lineaFactura("IDLineaAlbaran"), ProposedQFacturada - OriginalQFacturada, services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarImportesAlbaran(ByVal DocFra As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
        If DocFra Is Nothing Then Exit Sub
        Dim IDAlbaranAnt As Integer
        Dim ACL As New AlbaranCompraLinea : Dim context As New BusinessData
        Dim AlbPeriodoCerrado As New Dictionary(Of Integer, Boolean)
        For Each lineaFactura As DataRow In DocFra.dtLineas.Select(Nothing, "IDAlbaran,IDLineaAlbaran")
            If Length(lineaFactura("IDAlbaran")) > 0 AndAlso Length(lineaFactura("IDLineaAlbaran")) > 0 Then
                Dim Albaranes As DocumentInfoCache(Of DocumentoAlbaranCompra) = services.GetService(Of DocumentInfoCache(Of DocumentoAlbaranCompra))()
                Dim DocAlb As DocumentoAlbaranCompra = Albaranes.GetDocument(lineaFactura("IDAlbaran"))

                If IDAlbaranAnt <> lineaFactura("IDAlbaran") Then
                    IDAlbaranAnt = lineaFactura("IDAlbaran")

                    DocAlb.IDMoneda = DocFra.IDMoneda
                    DocAlb.CambioA = DocFra.CambioA
                    DocAlb.CambioB = DocFra.CambioB

                    context("IDMoneda") = DocAlb.IDMoneda
                    context("CambioA") = DocAlb.CambioA
                    context("CambioB") = DocAlb.CambioB
                End If

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
                        If lineaAlbaran("Precio") <> lineaFactura("Precio") Or lineaAlbaran("PrecioA") <> lineaFactura("PrecioA") Or lineaAlbaran("PrecioB") <> lineaFactura("PrecioB") Or _
                           lineaAlbaran("Dto1") <> lineaFactura("Dto1") Or lineaAlbaran("Dto2") <> lineaFactura("Dto2") Or lineaAlbaran("Dto3") <> lineaFactura("Dto3") Or _
                           lineaAlbaran("Dto") <> lineaFactura("Dto") Or lineaAlbaran("DtoProntoPago") <> lineaFactura("DtoProntoPago") Then

                            Dim LineaAlb As IPropertyAccessor = New DataRowPropertyAccessor(lineaAlbaran)
                            If LineaAlb("Precio") <> lineaFactura("Precio") Then
                                LineaAlb("Precio") = lineaFactura("Precio")
                                LineaAlb = ACL.ApplyBusinessRule("Precio", lineaFactura("Precio"), LineaAlb, context)
                            End If
                            If LineaAlb("PrecioA") <> lineaFactura("PrecioA") Then
                                LineaAlb("PrecioA") = lineaFactura("PrecioA")
                                LineaAlb = ACL.ApplyBusinessRule("PrecioA", lineaFactura("PrecioA"), LineaAlb, context)
                            End If
                            If LineaAlb("PrecioB") <> lineaFactura("PrecioB") Then
                                LineaAlb("PrecioB") = lineaFactura("PrecioB")
                                LineaAlb = ACL.ApplyBusinessRule("PrecioB", lineaFactura("PrecioB"), LineaAlb, context)
                            End If
                            If LineaAlb("Dto1") <> lineaFactura("Dto1") Then
                                LineaAlb("Dto1") = lineaFactura("Dto1")
                                LineaAlb = ACL.ApplyBusinessRule("Dto1", lineaFactura("Dto1"), LineaAlb, context)
                            End If
                            If LineaAlb("Dto2") <> lineaFactura("Dto2") Then
                                LineaAlb("Dto2") = lineaFactura("Dto2")
                                LineaAlb = ACL.ApplyBusinessRule("Dto2", lineaFactura("Dto2"), LineaAlb, context)
                            End If
                            If LineaAlb("Dto3") <> lineaFactura("Dto3") Then
                                LineaAlb("Dto3") = lineaFactura("Dto3")
                                LineaAlb = ACL.ApplyBusinessRule("Dto3", lineaFactura("Dto3"), LineaAlb, context)
                            End If
                            If LineaAlb("Dto") <> lineaFactura("Dto") Then
                                LineaAlb("Dto") = lineaFactura("Dto")
                                LineaAlb = ACL.ApplyBusinessRule("Dto", lineaFactura("Dto"), LineaAlb, context)
                            End If
                            If LineaAlb("DtoProntoPago") <> lineaFactura("DtoProntoPago") Then
                                LineaAlb("DtoProntoPago") = lineaFactura("DtoProntoPago")
                                LineaAlb = ACL.ApplyBusinessRule("DtoProntoPago", lineaFactura("DtoProntoPago"), LineaAlb, context)
                            End If

                            '//Quitamos esta corrección del movimiento, por que entraremos a corregir el movimiento desde la tarea GrabarAlbaranes, 
                            '//y así se generará el movimiento sólo una vez.
                            Dim ctx As New DataDocRow(DocAlb, lineaAlbaran)
                            Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of DataDocRow, StockUpdateData)(AddressOf ProcesoAlbaranCompra.CorregirMovimiento, ctx, services)
                            'End If
                        End If
                    Next
                End If
            End If

        Next
    End Sub


    <Task()> Public Shared Sub GrabarAlbaranes(ByVal data As Object, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of Object)(AddressOf Comunes.BeginTransaction, Nothing, services)
        Dim Albaranes As DocumentInfoCache(Of DocumentoAlbaranCompra) = services.GetService(Of DocumentInfoCache(Of DocumentoAlbaranCompra))()
        Dim AC As New AlbaranCompraCabecera
        For Each key As Integer In Albaranes.Keys
            Dim DocAlb As DocumentoAlbaranCompra = Albaranes.GetDocument(key)
            DocAlb.SetData()
        Next
        Albaranes.Clear() '//Hemos actualizado los albaranes, para que los podamos tener actualizados, debemos limpiar la lista
    End Sub

#End Region

#Region " Actualización Control Obras "

    <Task()> Public Shared Sub ActualizarObras(ByVal FraCab As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        If FraCab.HeaderRow("Estado") = enumfccEstado.fccNoContabilizado Then
            If Not FraCab Is Nothing AndAlso Not FraCab.HeaderRow Is Nothing AndAlso Not FraCab.dtLineas Is Nothing Then
                For Each drLinea As DataRow In FraCab.dtLineas.Select
                    Dim GeneradoControl As Boolean = ProcessServer.ExecuteTask(Of DataRow, Boolean)(AddressOf ActualizacionControlObras.AlbaranGeneradoControl, drLinea, services)
                    If Not GeneradoControl Then
                        If (drLinea.RowState = DataRowState.Added AndAlso Length(drLinea("IDObra")) > 0) OrElse _
                          (drLinea.RowState = DataRowState.Modified AndAlso (drLinea("ImporteA") <> drLinea("ImporteA", DataRowVersion.Original) _
                          OrElse Nz(drLinea("TipoGastoObra")) <> Nz(drLinea("TipoGastoObra", DataRowVersion.Original)) _
                          OrElse Nz(drLinea("IDObra")) <> Nz(drLinea("IDObra", DataRowVersion.Original)) _
                          OrElse Nz(drLinea("IDTrabajo")) <> Nz(drLinea("IDTrabajo", DataRowVersion.Original)) _
                          OrElse Nz(drLinea("IDLineaPadre")) <> Nz(drLinea("IDLineaPadre", DataRowVersion.Original)))) _
                          OrElse (drLinea.RowState <> DataRowState.Added AndAlso FraCab.HeaderRow.RowState = DataRowState.Modified AndAlso FraCab.HeaderRow("FechaFactura") <> FraCab.HeaderRow("FechaFactura", DataRowVersion.Original)) Then
                            Dim info As New ActualizacionControlObras.dataControlObras(drLinea, FraCab.HeaderRow("FechaFactura"), ActualizacionControlObras.enumOrigen.Factura)
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
        End If
    End Sub

#End Region

#Region " Actualización Control OTs "

    <Task()> Public Shared Sub ActualizarOTs(ByVal FraCab As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        If FraCab.HeaderRow("Estado") = enumfccEstado.fccNoContabilizado Then
            If Not FraCab Is Nothing AndAlso Not FraCab.HeaderRow Is Nothing AndAlso Not FraCab.dtLineas Is Nothing Then
                For Each drLinea As DataRow In FraCab.dtLineas.Select
                    If (drLinea.RowState = DataRowState.Added AndAlso Length(drLinea("IDMntoOTPrev")) > 0) OrElse _
                      (drLinea.RowState = DataRowState.Modified AndAlso (drLinea("ImporteA") <> drLinea("ImporteA", DataRowVersion.Original) _
                      OrElse Nz(drLinea("IDArticulo")) <> Nz(drLinea("IDArticulo", DataRowVersion.Original)) _
                      OrElse Nz(drLinea("IDMntoOTPrev")) <> Nz(drLinea("IDMntoOTPrev", DataRowVersion.Original)))) _
                      OrElse (FraCab.HeaderRow.RowState = DataRowState.Modified AndAlso FraCab.HeaderRow("FechaFactura") <> FraCab.HeaderRow("FechaFactura", DataRowVersion.Original)) Then
                        Dim info As New dataControlOTs(drLinea, FraCab.HeaderRow("FechaFactura"))
                        ProcessServer.ExecuteTask(Of dataControlOTs)(AddressOf ActualizarOTMaterialControl, info, services)
                    End If
                Next
            End If
        End If
    End Sub

#Region " ActualizarOTMaterialControl "

    Public Class dataControlOTs
        Public drLinea As DataRow
        Public FechaFactura As Date

        Public Sub New(ByVal drLinea As DataRow, ByVal FechaFactura As Date)
            Me.drLinea = drLinea
            Me.FechaFactura = FechaFactura
        End Sub
    End Class

    <Task()> Public Shared Sub ActualizarOTMaterialControl(ByVal data As dataControlOTs, ByVal services As ServiceProvider)
        If Length(data.drLinea("IDMntoOTPrev")) > 0 Then
            ProcessServer.ExecuteTask(Of dataControlOTs)(AddressOf NuevaLineaOTMaterialControl, data, services)
        Else
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf DeleteOTMaterialControl, data.drLinea, services)
        End If
    End Sub

    <Task()> Public Shared Sub NuevaLineaOTMaterialControl(ByVal data As dataControlOTs, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf DeleteOTMaterialControl, data.drLinea, services)
        Dim dtPrev As DataTable = BusinessHelper.CreateBusinessObject("MntoOTPrevLinea").Filter(New NumberFilterItem("IDMntoOTPrev", data.drLinea("IDMntoOTPrev")))
        If Not IsNothing(dtPrev) AndAlso dtPrev.Rows.Count > 0 Then
            Dim IDOT As Integer = dtPrev.Rows(0)("IDOT")

            Dim dtMatControl As DataTable = BusinessHelper.CreateBusinessObject("MntoOTControlLinea").AddNewForm
            dtMatControl.Rows(0)("IDMntoOTControl") = AdminData.GetAutoNumeric
            dtMatControl.Rows(0)("IDOT") = IDOT
            dtMatControl.Rows(0)("Tipo") = enumOTTipoLineasPrev.OTMaterial
            dtMatControl.Rows(0)("IDArticulo") = data.drLinea("IDArticulo")
            dtMatControl.Rows(0)("DescArticulo") = data.drLinea("DescArticulo")
            dtMatControl.Rows(0)("QConsumida") = data.drLinea("Cantidad")
            dtMatControl.Rows(0)("PrecioMatA") = data.drLinea("PrecioA")
            dtMatControl.Rows(0)("Dto1") = data.drLinea("Dto1")
            dtMatControl.Rows(0)("Dto2") = data.drLinea("Dto2")
            dtMatControl.Rows(0)("Dto3") = data.drLinea("Dto3")
            dtMatControl.Rows(0)("Fecha") = data.FechaFactura
            dtMatControl.Rows(0)("IDLineaFactura") = data.drLinea("IDLineaFactura")
            dtMatControl.Rows(0)("IDUDMedida") = data.drLinea("IDUDMedida")
            dtMatControl.Rows(0)("IDAlmacen") = ProcessServer.ExecuteTask(Of DataRow, String)(AddressOf GetAlmacenOTMaterialControl, data.drLinea, services)
            dtMatControl.Rows(0)("UDValoracion") = data.drLinea("UDValoracion")
            Dim infoImporte As New dataCalcularImporte(data.drLinea("Cantidad"), data.drLinea("PrecioA"), data.drLinea("Dto1"), data.drLinea("Dto2"), data.drLinea("Dto3"), data.drLinea("UDValoracion"))
            dtMatControl.Rows(0)("ImporteA") = ProcessServer.ExecuteTask(Of dataCalcularImporte, Double)(AddressOf CalcularImporteMateriales, infoImporte, services)

            BusinessHelper.UpdateTable(dtMatControl)
        End If
    End Sub

    <Task()> Public Shared Function GetAlmacenOTMaterialControl(ByVal data As DataRow, ByVal services As ServiceProvider) As String
        Dim IDAlmacen As String = String.Empty
        If Length(IDAlmacen) = 0 AndAlso Length(data("IDLineaAlbaran")) > 0 Then
            Dim Albaranes As EntityInfoCache(Of AlbaranCompraLineaInfo) = services.GetService(Of EntityInfoCache(Of AlbaranCompraLineaInfo))()
            Dim Albaran As AlbaranCompraLineaInfo = Albaranes.GetEntity(data("IDLineaAlbaran"))

            IDAlmacen = Albaran.IDAlmacen
        End If
        If Len(IDAlmacen) = 0 Then
            Dim StDatos As New DataArtAlm(data("IDArticulo"))
            IDAlmacen = ProcessServer.ExecuteTask(Of DataArtAlm, String)(AddressOf ArticuloAlmacen.AlmacenPredeterminadoArticulo, StDatos, services)
        End If
        Return IDAlmacen
    End Function

#Region " CalcularImporteMateriales "

    <Serializable()> _
    Public Class dataCalcularImporte
        Public Cantidad, Tasa, Dto1, Dto2, Dto3 As Double
        Public UDValoracion As Integer

        Public Sub New(ByVal Cantidad As Double, ByVal Tasa As Double, ByVal Dto1 As Double, ByVal Dto2 As Double, ByVal Dto3 As Double, ByVal UDValoracion As Double)
            Me.Cantidad = Cantidad
            Me.Tasa = Tasa
            Me.Dto1 = Dto1
            Me.Dto2 = Dto2
            Me.Dto3 = Dto3
            Me.UDValoracion = UDValoracion
            If Me.UDValoracion = 0 Then Me.UDValoracion = 1
        End Sub
    End Class

    <Task()> Public Shared Function CalcularImporteMateriales(ByVal data As dataCalcularImporte, ByVal services As ServiceProvider) As Double
        Dim Importe As Double = data.Cantidad * (data.Tasa / data.UDValoracion)
        Importe = Importe * (1 - data.Dto1 / 100)
        Importe = Importe * (1 - data.Dto2 / 100)
        Importe = Importe * (1 - data.Dto3 / 100)

        Return Importe
    End Function

#End Region

#End Region

    <Task()> Public Shared Sub DeleteOTMaterialControl(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDMntoOTPrev")) > 0 AndAlso Length(data("IDLineaFactura")) > 0 Then
            Dim Control As BusinessHelper = BusinessHelper.CreateBusinessObject("MntoOTControlLinea")
            Dim dtCtrl As DataTable = Control.Filter(New NumberFilterItem("IDLineaFactura", data("IDLineaFactura")))
            If Not IsNothing(dtCtrl) AndAlso dtCtrl.Rows.Count > 0 Then
                Control.Delete(dtCtrl)
            End If
        End If
    End Sub

#End Region

#Region " IVA Caja "

    <Task()> Public Shared Sub FechaParaDeclaracionPorProveedor(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider)
        If Not Nz(data("FechaDeclaracionManual"), False) AndAlso data.ContainsKey("IDProveedor") AndAlso Length(data("IDProveedor")) > 0 Then
            Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
            Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(data("IDProveedor"))
            If ProvInfo.IVACaja Then
                data("FechaParaDeclaracion") = New Date(Year(Nz(data("FechaFactura"), Today)) + 1, 12, 31) 'NegocioGeneral.cnMAX_DATE
            Else
                data("FechaParaDeclaracion") = Nz(data("FechaFactura"), Today)
            End If
        End If
    End Sub

#End Region

End Class


