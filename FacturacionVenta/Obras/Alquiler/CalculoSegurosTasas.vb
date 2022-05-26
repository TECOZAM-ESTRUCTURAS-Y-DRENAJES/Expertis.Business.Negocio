Public Class CalculoSegurosTasas

#Region " AñadirSegurosFacturacion "

    <Task()> Public Shared Sub AñadirSegurosFacturacion(ByVal doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If Not doc.dtLineas Is Nothing AndAlso doc.dtLineas.Rows.Count > 0 Then
            Dim dtSeguros As DataTable = ProcessServer.ExecuteTask(Of DataTable, DataTable)(AddressOf AnalisisSeguros, doc.dtLineas, services)
            If Not dtSeguros Is Nothing AndAlso dtSeguros.Rows.Count > 0 Then
                Dim IDTipoLinea As String = ProcessServer.ExecuteTask(Of Object, String)(AddressOf TipoLinea.TipoLineaPorDefecto, Nothing, services)
                Dim DescConcepto As String = doc.dtLineas.Rows(0)("DescConcepto") & String.Empty

                Dim IDOrdenLinea As Integer = 0
                Dim context As New BusinessData(doc.HeaderRow)

                Dim fvl As New FacturaVentaLinea
                For Each drSeguros As DataRow In dtSeguros.Rows
                    Dim drlinea As DataRow = doc.dtLineas.NewRow
                    drlinea("IDLineaFactura") = AdminData.GetAutoNumeric
                    drlinea("IDFactura") = doc.HeaderRow("IDFactura")
                    IDOrdenLinea = IDOrdenLinea + 1
                    drlinea("IDOrdenLinea") = IDOrdenLinea
                    drlinea("IDArticulo") = drSeguros("IDArticulo")
                    drlinea = fvl.ApplyBusinessRule("IDArticulo", drSeguros("IDArticulo"), drlinea, context)
                    drlinea("IDArticulo") = drSeguros("IDArticulo")

                    Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                    Dim Articulo As ArticuloInfo = Articulos.GetEntity(drSeguros("IDArticulo"))
                    drlinea("DescArticulo") = Articulo.DescArticulo
                    drlinea("IDUDMedida") = Articulo.IDUDVenta
                    drlinea("IDUDInterna") = Articulo.IDUDInterna
                    If Length(Articulo.IDConcepto) = 0 Then drlinea("IDConcepto") = Articulo.IDConcepto
                    drlinea("DescConcepto") = DescConcepto

                    Dim datos As New ProcesoComercial.DataIvaArticuloCliente(doc.HeaderRow("IDCliente"), drlinea("IDArticulo"))
                    drlinea("IDTipoIva") = ProcessServer.ExecuteTask(Of ProcesoComercial.DataIvaArticuloCliente, String)(AddressOf ProcesoComercial.GetIva, datos, services)

                    drlinea("IDCentroGestion") = doc.HeaderRow("IDCentroGestion")
                    drlinea("Cantidad") = 1
                    drlinea("Precio") = drSeguros("Precio")
                    drlinea("IDObra") = drSeguros("IDObra")
                    drlinea("Importe") = drlinea("Precio")
                    drlinea("IDTipoLinea") = IDTipoLinea
                    drlinea("Factor") = 1
                    drlinea("QInterna") = 1
                    drlinea("UDValoracion") = 1

                    Dim lineaIProperty As New ValoresAyB(New DataRowPropertyAccessor(drlinea), doc.IDMoneda, doc.CambioA, doc.CambioB)
                    ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, lineaIProperty, services)

                    doc.dtLineas.Rows.Add(drlinea.ItemArray)
                Next
            End If
        End If
    End Sub

#End Region

    <Task()> Public Shared Sub AñadirTasasDeResiduos(ByVal doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If Not doc.dtLineas Is Nothing AndAlso doc.dtLineas.Rows.Count > 0 Then
            Dim IDObra As Integer = doc.dtLineas.Rows(0)("IDObra")
            Dim DescConcepto As String = doc.dtLineas.Rows(0)("DescConcepto") & String.Empty

            Dim dataAnalisis As New dataAnalisisTasas(doc.dtLineas, doc.HeaderRow("FechaFactura"))
            ProcessServer.ExecuteTask(Of dataAnalisisTasas)(AddressOf AnalisisTasas, dataAnalisis, services)
            If dataAnalisis.ImporteTasa > 0 Then
                Dim drlinea As DataRow = doc.dtLineas.NewRow
                drlinea("IDLineaFactura") = AdminData.GetAutoNumeric
                drlinea("IDFactura") = doc.HeaderRow("IDFactura")
                '  drlinea("NFactura") = doc.HeaderRow("NFactura")
                drlinea("IDOrdenLinea") = doc.dtLineas.Rows.Count + 1
                drlinea("IDArticulo") = dataAnalisis.ArticuloTasa

                Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                Dim Articulo As ArticuloInfo = Articulos.GetEntity(drlinea("IDArticulo"))
                drlinea("DescArticulo") = Articulo.DescArticulo
                If ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf Cliente.Nacional, doc.HeaderRow("IDCliente"), services) Then
                    drlinea("CContable") = Articulo.CCVenta
                Else
                    drlinea("CContable") = Articulo.CCExport
                End If
                drlinea("IDUDMedida") = Articulo.IDUDVenta
                drlinea("IDUDInterna") = Articulo.IDUDInterna
                If Length(Articulo.IDConcepto) = 0 Then drlinea("IDConcepto") = Articulo.IDConcepto

                drlinea("DescConcepto") = doc.dtLineas.Rows(0)("DescConcepto")

                Dim datos As New ProcesoComercial.DataIvaArticuloCliente(doc.HeaderRow("IDCliente"), String.Empty)
                drlinea("IDTipoIva") = ProcessServer.ExecuteTask(Of ProcesoComercial.DataIvaArticuloCliente, String)(AddressOf ProcesoComercial.GetIva, datos, services)

                drlinea("IDCentroGestion") = doc.HeaderRow("IDCentroGestion")
                drlinea("Cantidad") = 1
                drlinea("Precio") = dataAnalisis.ImporteTasa
                drlinea("IDObra") = doc.dtLineas.Rows(0)("IDObra")
                drlinea("Importe") = dataAnalisis.ImporteTasa
                drlinea("IDTipoLinea") = ProcessServer.ExecuteTask(Of Object, String)(AddressOf TipoLinea.TipoLineaPorDefecto, Nothing, services)
                drlinea("Factor") = 1
                drlinea("QInterna") = 1
                drlinea("UDValoracion") = 1

                Dim lineaIProperty As New ValoresAyB(New DataRowPropertyAccessor(drlinea), doc.IDMoneda, doc.CambioA, doc.CambioB)
                ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, lineaIProperty, services)

                doc.dtLineas.Rows.Add(drlinea.ItemArray)
            End If
        End If
    End Sub

#Region " AnalisisSeguros "

    <Task()> Public Shared Function AnalisisSeguros(ByVal dtLineas As DataTable, ByVal services As ServiceProvider) As DataTable
        If dtLineas.Rows.Count > 0 Then
            Dim dtDatos As DataTable = ProcessServer.ExecuteTask(Of DataTable, DataTable)(AddressOf GetDatosCalculoSeguro, dtLineas, services)
            If Not dtDatos Is Nothing AndAlso dtDatos.Rows.Count > 0 Then
                Dim Parametros As ParametroAlquiler = services.GetService(Of ParametroAlquiler)()
                Dim blnExcesoContador As Boolean = Parametros.IntervienenExcesosContadoresEnCalculoSeguros

                Dim dtLineasSeguro As New DataTable
                dtLineasSeguro.Columns.Add("IDObra", GetType(Integer))
                dtLineasSeguro.Columns.Add("IDArticulo", GetType(String))
                dtLineasSeguro.Columns.Add("Precio", GetType(Double))

                Dim dtLineasSeguroAux As DataTable = dtLineasSeguro.Clone
                Dim IDObra As Integer

                Dim dblImporteTotal As Double
                Dim dblImporteTotalAUX As Double
                Dim DE As New BE.DataEngine
                For Each drLineas As DataRow In dtLineas.Select
                    If drLineas("IDObra") <> IDObra Then
                        IDObra = drLineas("IDObra")

                        Dim Obras As EntityInfoCache(Of ObraCabeceraInfo) = services.GetService(Of EntityInfoCache(Of ObraCabeceraInfo))()
                        Dim Obra As ObraCabeceraInfo = Obras.GetEntity(IDObra)

                        Dim f As New Filter
                        f.Add(New NumberFilterItem("IDObra", IDObra))
                        f.Add(New NumberFilterItem("Tipo", 2))
                        Dim dt As DataTable = DE.Filter("tbObraTarifaAlquiler", f)
                        If Not dt Is Nothing Then
                            Dim dv As New DataView(dt)
                            For i As Integer = 0 To 3
                                Select Case i
                                    Case 0
                                        dv.RowFilter = "TipoDato= 1"
                                    Case 1
                                        dv.RowFilter = "TipoDato= 2"
                                    Case 2
                                        dv.RowFilter = "TipoDato= 3"
                                    Case 3
                                        dv.RowFilter = "TipoDato= 0"
                                End Select
                                Dim strWhere As String = String.Empty
                                If dv.Count > 0 Then
                                    Dim blnEncontrado As Boolean
                                    For Each drv As DataRowView In dv
                                        Select Case drv("Tipodato")
                                            Case 1
                                                strWhere = "IDArticulo='" & drv("IDArticulo") & "'"
                                            Case 3
                                                strWhere = "IDTipo='" & drv("IDTipo") & "' AND IDFamilia='" & drv("IDFamilia") & "'"
                                            Case 2
                                                strWhere = "IDTipo='" & drv("IDTipo") & "'"
                                            Case 0
                                                strWhere = String.Empty
                                        End Select

                                        Dim dvLinSeg As New DataView(dtDatos)

                                        If Length(strWhere) > 0 Then strWhere = strWhere & " AND "
                                        strWhere = strWhere & "IDObra=" & IDObra
                                        If Not blnExcesoContador Then strWhere = strWhere & " AND ExcesoContador = False"
                                        dvLinSeg.RowFilter = strWhere

                                        Dim IDArticulo As String
                                        Dim dblTotPrecio As Double = 0
                                        If dvLinSeg.Count > 0 Then
                                            Dim strMarcadas As String
                                            For Each drvLinSeg As DataRowView In dvLinSeg
                                                Dim blnMarcadas As Boolean = (InStr(1, strMarcadas, drvLinSeg("IdLineaFactura")) = 0)
                                                If blnMarcadas Then
                                                    IDArticulo = drv("IdArticuloFactura")
                                                    Dim dblTiempo As Double
                                                    If drvLinSeg("QTiempo") = 0 Then
                                                        dblTiempo = 1
                                                    Else
                                                        dblTiempo = drvLinSeg("QTiempo")
                                                    End If
                                                    Dim dblImporteLinea As Double
                                                    If drv("TipoCalculoSeguro") = CInt(otaTipoCalcSeg.PrecioUdFactAlquiler) Then
                                                        dblImporteLinea = dblTiempo * Nz(drv("PrecioSeguroFactAlquiler"), 0)
                                                    ElseIf Obra.TipoGeneracionSeguros = CInt(enumocSegurosTipoImporte.ocImporteNeto) Then 'Calcular sobre importe Netos
                                                        dblImporteLinea = drvLinSeg("Importe")
                                                    ElseIf Obra.TipoGeneracionSeguros = CInt(enumocSegurosTipoImporte.ocImporteBruto) Then 'Calcular sobre Importe Bruto
                                                        dblImporteLinea = drvLinSeg("Precio") * drvLinSeg("Cantidad") * dblTiempo
                                                    End If
                                                    strMarcadas = strMarcadas & drvLinSeg("IdLineaFactura") & ","
                                                    dblTotPrecio = dblTotPrecio + dblImporteLinea
                                                    blnEncontrado = True
                                                End If
                                            Next
                                        End If

                                        If blnEncontrado Then
                                            Dim DblImporte As Double = 0
                                            Select Case drv("TipoCalculoSeguro")
                                                Case 0
                                                    If drv("PorcentajeSeguro") > 0 Then
                                                        DblImporte = dblTotPrecio * drv("PorcentajeSeguro") / 100
                                                    End If
                                                Case 1
                                                    DblImporte = drv("ImporteFijo")
                                                Case 2
                                                    DblImporte = dblTotPrecio
                                            End Select

                                            If DblImporte > 0 Then
                                                If Not dtLineasSeguroAux Is Nothing Then
                                                    Dim drLineasSeguroAux As DataRow = dtLineasSeguroAux.NewRow
                                                    drLineasSeguroAux("IDObra") = IDObra
                                                    drLineasSeguroAux("IDArticulo") = IDArticulo
                                                    drLineasSeguroAux("Precio") = DblImporte
                                                    dtLineasSeguroAux.Rows.Add(drLineasSeguroAux)
                                                End If
                                            End If

                                            dblImporteTotal = dblImporteTotal + DblImporte
                                        End If
                                        blnEncontrado = False
                                    Next
                                End If
                            Next i
                        End If
                    End If
                Next
                If Not dtLineasSeguroAux Is Nothing And dtLineasSeguroAux.Rows.Count > 0 Then
                    Dim dvLineasAux As New DataView(dtLineasSeguroAux)
                    dvLineasAux.Sort = "IDArticulo, IDObra"
                    Dim IDArticuloAUX As String
                    Dim IDObraAUX As Integer
                    For Each drvLineaAux As DataRowView In dvLineasAux
                        If drvLineaAux("IDArticulo") & String.Empty <> IDArticuloAUX Or Nz(drvLineaAux("IDObra"), 0) <> IDObraAUX Then
                            If dblImporteTotalAUX > 0 Then
                                Dim drLineasSeguro As DataRow = dtLineasSeguro.NewRow
                                drLineasSeguro("IDObra") = IDObraAUX
                                drLineasSeguro("IDArticulo") = IDArticuloAUX
                                drLineasSeguro("Precio") = dblImporteTotalAUX
                                dtLineasSeguro.Rows.Add(drLineasSeguro)
                            End If
                            dblImporteTotalAUX = 0
                        End If

                        IDArticuloAUX = drvLineaAux("IDArticulo") & String.Empty
                        IDObraAUX = Nz(drvLineaAux("IDObra"), 0)
                        dblImporteTotalAUX = dblImporteTotalAUX + Nz(drvLineaAux("Precio"), 0)
                    Next

                    If dblImporteTotalAUX <> 0 Then
                        Dim drLineasSeguro As DataRow = dtLineasSeguro.NewRow
                        drLineasSeguro("IDObra") = IDObraAUX
                        drLineasSeguro("IDArticulo") = IDArticuloAUX
                        drLineasSeguro("Precio") = dblImporteTotalAUX
                        dtLineasSeguro.Rows.Add(drLineasSeguro)
                    End If
                End If

                Return dtLineasSeguro
            End If
        End If
        Return Nothing
    End Function

    <Task()> Public Shared Function GetDatosCalculoSeguro(ByVal dtLineas As DataTable, ByVal services As ServiceProvider) As DataTable
        If Not dtLineas Is Nothing AndAlso dtLineas.Rows.Count > 0 Then
            Dim dtLineasSeguro As New DataTable
            dtLineasSeguro.Columns.Add("IDLineafactura", GetType(Integer))
            dtLineasSeguro.Columns.Add("IDArticulo", GetType(String))
            dtLineasSeguro.Columns.Add("IDTipo", GetType(String))
            dtLineasSeguro.Columns.Add("IDFamilia", GetType(String))
            dtLineasSeguro.Columns.Add("IDObra", GetType(Integer))
            dtLineasSeguro.Columns.Add("Cantidad", GetType(Double))
            dtLineasSeguro.Columns.Add("QTiempo", GetType(Double))
            dtLineasSeguro.Columns.Add("Precio", GetType(Double))
            dtLineasSeguro.Columns.Add("Importe", GetType(Double))
            dtLineasSeguro.Columns.Add("ExcesoContador", GetType(Boolean))

            For Each drLineas As DataRow In dtLineas.Rows
                Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                Dim Articulo As ArticuloInfo = Articulos.GetEntity(drLineas("IDArticulo"))

                If Not Articulo.SinSeguroEnAlquiler And drLineas("TipoFactAlquiler") <> enumTipoFacturacionAlquiler.enumTFASinAlquiler Then
                    Dim drLineaSeguro As DataRow = dtLineasSeguro.NewRow
                    drLineaSeguro("IDLineaFactura") = drLineas("IDLineaFactura")
                    drLineaSeguro("IDArticulo") = drLineas("IDArticulo")
                    drLineaSeguro("IDTipo") = Articulo.IDTipo
                    drLineaSeguro("IDFamilia") = Articulo.IDFamilia
                    drLineaSeguro("IDObra") = drLineas("IDObra")
                    drLineaSeguro("Cantidad") = drLineas("Cantidad")
                    drLineaSeguro("QTiempo") = Nz(drLineas("QTiempo"), 1)
                    drLineaSeguro("Precio") = drLineas("Precio")
                    drLineaSeguro("Importe") = drLineas("Importe")
                    drLineaSeguro("ExcesoContador") = Nz(drLineas("ExcesoContador"), False)

                    dtLineasSeguro.Rows.Add(drLineaSeguro)
                End If
            Next

            Return dtLineasSeguro
        End If

        Return Nothing
    End Function

#End Region

#Region " AnalisisTasas "

    <Serializable()> _
    Public Class dataAnalisisTasas
        Public dtLineas As DataTable
        Public FechaFactura As Date?
        Public Agrupacion As String = String.Empty
        Public ArticuloTasa As String
        Public ImporteTasa As Double = 0

        Public Sub New(ByVal dtLineas As DataTable, ByVal FechaFactura As Date, Optional ByVal Agrupacion As String = Nothing)
            Me.dtLineas = dtLineas
            Me.FechaFactura = FechaFactura
            Me.Agrupacion = Agrupacion
        End Sub

        Public Sub New(ByVal dtLineas As DataTable)
            Me.dtLineas = dtLineas
        End Sub
    End Class
    <Task()> Public Shared Sub AnalisisTasas(ByVal data As dataAnalisisTasas, ByVal services As ServiceProvider)
        If data.dtLineas.Rows.Count > 0 Then
            Dim Parametros As ParametroAlquiler = services.GetService(Of ParametroAlquiler)()
            Dim TasaResiduos As Double = Parametros.PorcentajeTasaResiduos
            Dim LimiteTasaResiduos As Double = Parametros.ImporteLimiteTasaResiduos()
            data.ArticuloTasa = Parametros.ArticuloTasaResiduos()

            Dim FacturacionAnterior As Double
            Dim IDArticulo As String = String.Empty
            Dim Lote As String = String.Empty
            Dim IDLineaMaterial As Integer = -1
            Dim IDLineaAlbaranRetorno As Integer = -1
            Dim Importe As Double

            Dim Where As String = String.Empty
            If Length(data.Agrupacion) > 0 Then Where = "Agrupacion = " & data.Agrupacion
            For Each drLinea As DataRow In data.dtLineas.Select(Where, "IDLineaMaterial, IDArticulo, Lote")
                FacturacionAnterior = 0
                'If Nz(drLinea("IDLineaAlbaranRetorno"), -1) <> IDLineaAlbaranRetorno Or drLinea("IDArticulo") & String.Empty <> IDArticulo Or drLinea("Lote") & String.Empty <> Lote Or Nz(drLinea("IDLineaMaterial"), -1) <> IDLineaMaterial Then
                '    If Importe <> 0 Then
                '        Dim dataFacturacionAnterior As New dataFacturacionAnteriorOS(IDLineaMaterial, IDArticulo, Lote, data.FechaFactura)
                '        FacturacionAnterior = ProcessServer.ExecuteTask(Of dataFacturacionAnteriorOS, Double)(AddressOf GetFacturacionAnteriorOS, dataFacturacionAnterior, services)
                '        Importe = (Importe + FacturacionAnterior) * (TasaResiduos / 100)
                '        If Importe > LimiteTasaResiduos Then Importe = LimiteTasaResiduos
                '        data.ImporteTasa = data.ImporteTasa + Importe
                '        Importe = 0
                '    End If
                'End If

                IDLineaAlbaranRetorno = Nz(drLinea("IDLineaAlbaranRetorno"), -1)
                IDArticulo = drLinea("IDArticulo") & String.Empty
                Lote = drLinea("Lote") & String.Empty
                IDLineaMaterial = Nz(drLinea("IDLineaMaterial"), -1)
                If IDLineaAlbaranRetorno > -1 AndAlso drLinea("TipoFactAlquiler") <> enumTipoFacturacionAlquiler.enumTFASinAlquiler Then 'AndAlso Len(Lote) > 0 Then
                    Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                    Dim Articulo As ArticuloInfo = Articulos.GetEntity(drLinea("IDArticulo"))

                    Dim FacturarTasaResiduos As Boolean = Articulo.FactTasaResiduos
                    If Not FacturarTasaResiduos AndAlso Length(drLinea("IDObra")) > 0 Then
                        Dim Obras As EntityInfoCache(Of ObraCabeceraInfo) = services.GetService(Of EntityInfoCache(Of ObraCabeceraInfo))()
                        Dim Obra As ObraCabeceraInfo = Obras.GetEntity(drLinea("IDObra"))
                        FacturarTasaResiduos = Obra.FacturarTasaResiduos
                    End If
                    If FacturarTasaResiduos Then
                        Importe = (Importe + drLinea("Importe"))
                        If Importe <> 0 Then
                            Dim dataFacturacionAnterior As New dataFacturacionAnteriorOS(IDLineaMaterial, IDArticulo, Lote, data.FechaFactura)
                            FacturacionAnterior = ProcessServer.ExecuteTask(Of dataFacturacionAnteriorOS, Double)(AddressOf GetFacturacionAnteriorOS, dataFacturacionAnterior, services)
                            Importe = Importe + FacturacionAnterior
                        End If
                    End If
                End If
            Next
            If Importe <> 0 Then
                'Dim dataFacturacionAnterior As New dataFacturacionAnteriorOS(IDLineaMaterial, IDArticulo, Lote, data.FechaFactura)
                'FacturacionAnterior = ProcessServer.ExecuteTask(Of dataFacturacionAnteriorOS, Double)(AddressOf GetFacturacionAnteriorOS, dataFacturacionAnterior, services)
                'Importe = (Importe + FacturacionAnterior) * (TasaResiduos / 100)
                Importe = Importe * (TasaResiduos / 100)
                If Importe > LimiteTasaResiduos Then Importe = LimiteTasaResiduos
                data.ImporteTasa = Importe
            End If
        End If
    End Sub

    <Serializable()> _
    Public Class dataFacturacionAnteriorOS
        Public IDLineaMaterial As Integer
        Public IDArticulo As String
        Public Lote As String
        Public FechaFactura As Date?

        Public Sub New(ByVal IDLineaMaterial As Integer, ByVal IDArticulo As String, ByVal Lote As String, ByVal FechaFactura As Date?)
            Me.IDLineaMaterial = IDLineaMaterial
            Me.IDArticulo = IDArticulo
            Me.Lote = Lote
            Me.FechaFactura = FechaFactura
        End Sub
    End Class
    <Task()> Public Shared Function GetFacturacionAnteriorOS(ByVal data As dataFacturacionAnteriorOS, ByVal services As ServiceProvider) As Double
        Dim f As New Filter
        f.Add(New NumberFilterItem("IDLineaMaterial", data.IDLineaMaterial))
        f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
        If Length(data.Lote) > 0 Then f.Add(New StringFilterItem("Lote", data.Lote))
        If Not data.FechaFactura Is Nothing Then
            f.Add(New DateFilterItem("FechaFactura", FilterOperator.LessThan, data.FechaFactura))
        End If

        Dim Importe As Double = 0
        Dim dt As DataTable = New BE.DataEngine().Filter("vAlquilerNegDatosFacturacion", f)
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                Importe = Importe + dr("Importe")
            Next
        End If

        Return Importe
    End Function

#End Region

End Class
