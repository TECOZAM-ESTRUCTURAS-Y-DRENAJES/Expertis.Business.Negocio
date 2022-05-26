Public Class PrcCopiaFacturaCompra
    Inherits Process(Of Integer, CreateElement)

    'Lista de tareas a ejecutar
    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of Integer, FraCabCompraCopia)(AddressOf DatosIniciales)
        Me.AddTask(Of FraCabCompraCopia, DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CrearDocumentoFactura)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarValoresPredeterminadosGenerales)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf CrearDocumentoCopia)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf NegocioGeneral.AsignarEnvio347Doc)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf NegocioGeneral.AsignarEnvio349Doc)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularImporteLineasFacturas)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf NegocioGeneral.CalcularImpuestos)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularAnaliticaFacturas)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularBasesImponibles)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularTotales)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularVencimientos)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarClaveOperacion)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf Business.General.Comunes.UpdateDocument)
        Me.AddTask(Of DocumentoFacturaCompra, CreateElement)(AddressOf ProcesoFacturacionCompra.Resultado)
    End Sub

    <Task()> Public Shared Function DatosIniciales(ByVal data As Integer, ByVal services As ServiceProvider) As FraCabCompraCopia
        Dim Dr As DataRow = New FacturaCompraCabecera().GetItemRow(data)
        Dim FCC As New FraCabCompraCopia(Dr)
        Return FCC
    End Function

    <Task()> Public Shared Sub CrearDocumentoCopia(ByVal data As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        Dim DtCab As DataTable = New FacturaCompraCabecera().SelOnPrimaryKey(CType(data.Cabecera, FraCabCompraCopia).IDFactura)
        Dim DtLineas As DataTable = New FacturaCompraLinea().Filter(New FilterItem("IDFactura", FilterOperator.Equal, CType(data.Cabecera, FraCabCompraCopia).IDFactura))
        Dim AppParamsConta As ParametroContabilidadCompra = services.GetService(GetType(ParametroContabilidadCompra))
        Dim DrOrigenCabecera As DataRow = DtCab.Rows(0)

        For Each dc As DataColumn In DtCab.Columns
            If dc.ColumnName <> "IDFactura" And dc.ColumnName <> "NFactura" And _
               dc.ColumnName <> "FechaContabilizacion" And dc.ColumnName <> "IDObra" And _
               dc.ColumnName <> "FechaIntrastat" And dc.ColumnName <> "NDeclaracionIVA" And _
               dc.ColumnName <> "AñoDeclaracionIva" And dc.ColumnName <> "NFacturaIva" And _
               dc.ColumnName <> "NDeclaracionIntrastat" And dc.ColumnName <> "AñoDeclaracionIntrastat" And _
               dc.ColumnName <> "IDFacturaVenta" And dc.ColumnName <> "NFacturaAutofactura" And _
               dc.ColumnName <> "FechaRegContable" And dc.ColumnName <> "EnviarSII" And dc.ColumnName <> "IDClaveTipoFactura" And _
               dc.ColumnName <> "IDClaveRegimenEspecial" And dc.ColumnName <> "IDClaveRegimenEspecial1" And dc.ColumnName <> "IDClaveRegimenEspecial2" And _
               dc.ColumnName <> "FechaOperacion" Then
                data.HeaderRow(dc.ColumnName) = DrOrigenCabecera(dc)
            End If
        Next
        If Length(data.HeaderRow("IDContador")) > 0 Then
            Dim StDatos As New Contador.DatosCounterValue
            StDatos.IDCounter = data.HeaderRow("IDContador")
            StDatos.TargetClass = New FacturaCompraCabecera
            StDatos.TargetField = "NFactura"
            StDatos.DateField = "FechaFactura"
            StDatos.DateValue = data.HeaderRow("FechaFactura")
            StDatos.IDEjercicio = data.HeaderRow("IDEjercicio") & String.Empty
            data.HeaderRow("NFactura") = ProcessServer.ExecuteTask(Of Contador.DatosCounterValue, String)(AddressOf Contador.CounterValue, StDatos, services)
        End If
        data.HeaderRow("SuFactura") = data.HeaderRow("NFactura")
        Dim FCC As New FacturaCompraCabecera
        data.HeaderRow.ItemArray = FCC.ApplyBusinessRule("FechaFactura", Date.Today, data.HeaderRow, Nothing).ItemArray
        data.HeaderRow("SuFechaFactura") = Date.Today
        data.HeaderRow("FechaDeclaracionManual") = False
        data.HeaderRow("FechaParaDeclaracion") = Date.Today
        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ProcesoFacturacionCompra.FechaParaDeclaracionPorProveedor, New DataRowPropertyAccessor(data.HeaderRow), services)
        If AppParamsConta.Contabilidad Then
            data.HeaderRow("IDEjercicio") = ProcessServer.ExecuteTask(Of Date, String)(AddressOf NegocioGeneral.EjercicioPredeterminado, Today, services)
        End If
        data.HeaderRow("Estado") = enumfccEstado.fccNoContabilizado
        data.HeaderRow("IVAManual") = 0
        data.HeaderRow("VencimientosManuales") = 0
        data.HeaderRow("IntrastatProcesado") = 0
        'data.HeaderRow("Enviar347") = 1
        data.HeaderRow("Exportado") = 0
        data.HeaderRow("Exportar") = 1
        data.HeaderRow("FacturaPagoPeriodicoSN") = 0
        data.HeaderRow("NoDescontabilizar") = 0
        data.HeaderRow("RetencionManual") = 0

      
        'Copia Líneas
        For Each drOrigenLinea As DataRow In DtLineas.Select
            Dim drDestinoLinea As DataRow = data.dtLineas.NewRow
            For Each dc As DataColumn In DtLineas.Columns
                If dc.ColumnName <> "IDLineaFactura" And dc.ColumnName <> "IDFactura" And _
                   dc.ColumnName <> "IDPedido" And dc.ColumnName <> "IDLineaPedido" And _
                   dc.ColumnName <> "IDAlbaran" And dc.ColumnName <> "IDLineaAlbaran" And _
                   dc.ColumnName <> "IDObra" And dc.ColumnName <> "IDTrabajo" And _
                   dc.ColumnName <> "IDLineaPadre" And dc.ColumnName <> "IDMntoOTPrev" And _
                   dc.ColumnName <> "EstadoInmovilizado" And _
                   dc.ColumnName <> "IDLineaOfertaDetalle" And dc.ColumnName <> "IDActivoAImputar" Then
                    drDestinoLinea(dc.ColumnName) = drOrigenLinea(dc)
                End If
            Next
            drDestinoLinea("IDLineaFactura") = AdminData.GetAutoNumeric
            drDestinoLinea("IDFactura") = data.HeaderRow("IDFactura")
            data.dtLineas.Rows.Add(drDestinoLinea)
        Next
    End Sub

End Class