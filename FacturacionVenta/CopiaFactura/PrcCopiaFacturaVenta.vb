Imports Solmicro.Expertis.Business.ClasesTecozam

Public Class PrcCopiaFacturaVenta
    'Ibis Adolfo 20220208
    'Inherits Process(Of Integer, CreateElement)
    Inherits Process(Of TECDataPrcCopiarFactura, CreateElement)

    'Lista de tareas a ejecutar
    Public Overrides Sub RegisterTasks()
        'Ibis Adolfo 20220208
        'Me.AddTask(Of Integer, FraCabVentaCopia)(AddressOf DatosIniciales)
        Me.AddTask(Of TECDataPrcCopiarFactura, FraCabVentaCopia)(AddressOf DatosIniciales)

        Me.AddTask(Of FraCabVentaCopia, DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CrearDocumentoFactura)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf CrearDocumentoCopia)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularImporteLineasFacturas)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf NegocioGeneral.CalcularImpuestos)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularRepresentantes)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularAnalitica)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularBasesImponibles)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularTotales)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularPuntoVerde)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularVencimientos)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarClaveOperacion)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf Business.General.Comunes.UpdateDocument)
        Me.AddTask(Of DocumentoFacturaVenta, CreateElement)(AddressOf Resultado)
    End Sub

    'Ibis Adolfo 20220208
    <Task()> Public Shared Function DatosIniciales(ByVal data As Integer, ByVal services As ServiceProvider) As FraCabVentaCopia
        Dim Dr As DataRow = New FacturaVentaCabecera().GetItemRow(data)
        Dim FVC As New FraCabVentaCopia(Dr)
        Return FVC
    End Function

    <Task()> Public Shared Function DatosIniciales(ByVal data As TECDataPrcCopiarFactura, ByVal services As ServiceProvider) As FraCabVentaCopia
        If Nz(data.idContador, "") <> "" Then services.RegisterService(New ProcessInfo(data.idContador))
        Dim Dr As DataRow = New FacturaVentaCabecera().GetItemRow(data.idFactura)
        Dim FVC As New FraCabVentaCopia(Dr)
        Return FVC
    End Function

    <Task()> Public Shared Sub CrearDocumentoCopia(ByVal data As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        Dim DtCab As DataTable = New FacturaVentaCabecera().SelOnPrimaryKey(CType(data.Cabecera, FraCabVentaCopia).IDFactura)
        Dim DtLineas As DataTable = New FacturaVentaLinea().Filter(New FilterItem("IDFactura", FilterOperator.Equal, CType(data.Cabecera, FraCabVentaCopia).IDFactura))
        Dim AppParamsConta As ParametroContabilidadVenta = services.GetService(GetType(ParametroContabilidadVenta))

        Dim drOrigenCabecera As DataRow = DtCab.Rows(0)
        For Each dc As DataColumn In DtCab.Columns
            If dc.ColumnName <> "IDFactura" And dc.ColumnName <> "NFactura" And _
               dc.ColumnName <> "FechaContabilizacion" And dc.ColumnName <> "IDObra" And _
               dc.ColumnName <> "IDFacturaCompra" And dc.ColumnName <> "NDeclaracionIVA" And _
               dc.ColumnName <> "AñoDeclaracionIva" And dc.ColumnName <> "NDeclaracionIntrastat" And _
               dc.ColumnName <> "AñoDeclaracionIntrastat" And dc.ColumnName <> "DirecFacturaPDF" And _
               dc.ColumnName <> "DirecFacturaXML" And dc.ColumnName <> "DirecCorreoPDF" And _
               dc.ColumnName <> "EnviarSII" And dc.ColumnName <> "IDClaveTipoFactura" And _
               dc.ColumnName <> "IDClaveRegimenEspecial" And dc.ColumnName <> "IDClaveRegimenEspecial1" And dc.ColumnName <> "IDClaveRegimenEspecial2" And _
               dc.ColumnName <> "FechaOperacion" And dc.ColumnName <> "FechaDeclaracionManual" And dc.ColumnName <> "EmitidaPorTerceros" Then
                data.HeaderRow(dc.ColumnName) = drOrigenCabecera(dc)
            End If
        Next

        data.HeaderRow("IDFactura") = AdminData.GetAutoNumeric
        'Ibis Adolfo 20220208
        Dim info As ProcessInfo = services.GetService(Of ProcessInfo)()
        If Length(info.IDContador) > 0 Then
            data.HeaderRow("IDContador") = info.IDContador
        Else
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf FacturaVentaCabecera.AsignarContador, data.HeaderRow, services)
        End If

        If Length(data.HeaderRow("IDContador")) > 0 Then
            ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf SetContadorPredeterminadoEntidad, data, services)
            Dim StDatos As New Contador.DatosCounterValue(data.HeaderRow("IDContador"), New FacturaVentaCabecera, "NFactura", "FechaFactura", data.HeaderRow("FechaFactura"))
            StDatos.IDEjercicio = data.HeaderRow("IDEjercicio") & String.Empty
            data.HeaderRow("NFactura") = ProcessServer.ExecuteTask(Of Contador.DatosCounterValue, String)(AddressOf Contador.CounterValue, StDatos, services)
        End If
        data.HeaderRow("FechaFactura") = Date.Today
        data.HeaderRow("FechaParaDeclaracion") = Date.Today

        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ProcesoFacturacionVenta.FechaParaDeclaracionComoProveedor, New DataRowPropertyAccessor(data.HeaderRow), services)

        If AppParamsConta.Contabilidad Then
            Dim DataEjer As New DataEjercicio(New DataRowPropertyAccessor(data.HeaderRow), Today.Date)
            ProcessServer.ExecuteTask(Of DataEjercicio)(AddressOf NegocioGeneral.AsignarEjercicioContable, DataEjer, services)
        End If

        data.HeaderRow("Estado") = enumfvcEstado.fvcNoContabilizado
        data.HeaderRow("IVAManual") = 0
        data.HeaderRow("VencimientosManuales") = 0
        data.HeaderRow("GeneradoFichero") = 0
        data.HeaderRow("EnviadaEntidadAseguradora") = 0
        data.HeaderRow("Exportado") = 0
        data.HeaderRow("Exportar") = 1

        'Copia Líneas
        For Each drOrigenLinea As DataRow In DtLineas.Select
            Dim drDestinoLinea As DataRow = data.dtLineas.NewRow
            For Each dc As DataColumn In DtLineas.Columns
                If dc.ColumnName <> "IDLineaFactura" And dc.ColumnName <> "IDFactura" And _
                   dc.ColumnName <> "IDPedido" And dc.ColumnName <> "IDLineaPedido" And _
                   dc.ColumnName <> "IDAlbaran" And dc.ColumnName <> "IDLineaAlbaran" And _
                   dc.ColumnName <> "IDVencimiento" And dc.ColumnName <> "IDLineaVencimiento" And _
                   dc.ColumnName <> "IDLineaMaterial" And dc.ColumnName <> "IDLineaMOD" And _
                   dc.ColumnName <> "IDLineaCentro" And dc.ColumnName <> "IDLineaGasto" And _
                   dc.ColumnName <> "IDLineaVarios" And dc.ColumnName <> "IDPromocionLinea" And _
                   dc.ColumnName <> "IDAlbaranRetorno" And dc.ColumnName <> "IDLineaAlbaranRetorno" And _
                   dc.ColumnName <> "IDLineaOfertaDetalle" And dc.ColumnName <> "IDCertificacion" Then
                    'dc.ColumnName <> "IDObra" And dc.ColumnName <> "IDTrabajo" And _
                    drDestinoLinea(dc.ColumnName) = drOrigenLinea(dc)
                End If
            Next
            drDestinoLinea("IDFactura") = data.HeaderRow("IDFactura")
            drDestinoLinea("IDLineaFactura") = AdminData.GetAutoNumeric
            data.dtLineas.Rows.Add(drDestinoLinea)
        Next
    End Sub

    <Task()> Public Shared Sub SetContadorPredeterminadoEntidad(ByVal Doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        Dim dtCount As DataTable = New Contador().SelOnPrimaryKey(Doc.HeaderRow("IDContador"))
        If dtCount.Rows.Count = 0 Then
            Dim fContPred As New Filter
            fContPred.Add(New StringFilterItem("Entidad", Doc.EntidadCabecera))
            fContPred.Add(New BooleanFilterItem("Predeterminado", True))
            Dim dtContPred As DataTable = New EntidadContador().Filter(fContPred)
            If dtContPred.Rows.Count > 0 Then
                Doc.HeaderRow("IDContador") = dtContPred.Rows(0)("IDContador")
            End If
        End If
    End Sub

    <Task()> Public Shared Function Resultado(ByVal Doc As DocumentoFacturaVenta, ByVal services As ServiceProvider) As CreateElement
        Dim result As New CreateElement
        result.IDElement = Doc.HeaderRow("IDFactura")
        result.NElement = Doc.HeaderRow("NFactura")
        Return result
    End Function

End Class