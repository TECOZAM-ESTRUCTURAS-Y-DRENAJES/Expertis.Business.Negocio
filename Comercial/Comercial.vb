Imports Solmicro.Expertis.Business.Negocio.NegocioGeneral

<Transactional()> _
Public Class Comercial
    Inherits ContextBoundObject

    '#Region " BusinessRules - Quitar cuando TODO Comercial vaya por MetodosComercial"

    '    Public Function DetailBusinessRules(ByVal ColumnName As String, _
    '                                        ByVal Value As Object, _
    '                                        ByVal current As IPropertyAccessor, _
    '                                        ByVal services As ServiceProvider, _
    '                                        Optional ByVal Context As IPropertyAccessor = Nothing) As IPropertyAccessor

    '        If IsNothing(Context) Then Context = New BusinessData


    '        If Context.ContainsKey("IDMoneda") Then
    '            Dim Fecha As Date
    '            If Context.ContainsKey("Fecha") AndAlso Length(Context("Fecha")) > 0 Then
    '                Fecha = Context("Fecha")
    '            Else
    '                Fecha = Today
    '            End If
    '            'TODO ¿esto vale para algo?
    '            'If IsNothing(Monedas) Then Monedas = New MonedaCache
    '            'Monedas.GetMoneda(Context("IDMoneda"), Fecha)
    '        End If

    '        Select Case ColumnName
    '            Case "IDArticulo", "CodigoBarras", "RefCliente"
    '                CambioArticulo(ColumnName, Value, current, Context, services)
    '            Case "Cantidad"
    '                CambioCantidad(current, Value, Context, services)
    '            Case "QInterna"
    '                CambioQInterna(current, Value, Context, services)
    '            Case "Factor"
    '                CambioFactor(current, Value, Context, services)
    '            Case "IDUDMedida"
    '                CambioUDMedida(current, Value, Context, services)
    '            Case "Precio", "UDValoracion", "Dto1", "Dto2", "Dto3", "Dto", "DtoProntoPago"
    '                CambioPrecio(ColumnName, Value, current, Context, services)
    '            Case "PrecioA", "PrecioB"
    '                CalcularImportes(ColumnName, current, Value, Context, services)
    '            Case "CContable"
    '                CambioCContable(current, Value, Context)
    '            Case "NObra"
    '                CambioNObra(current, Value, Context, services)
    '            Case "CodTrabajo"
    '                CambioCodTrabajo(current, Value)
    '            Case "Regalo"
    '                CambioRegalo(current, Value, Context, services)
    '            Case "IDLineaAlbaran"
    '                CambioLineaAlbaran(current, Value, services)
    '        End Select

    '        Return current
    '    End Function

    '    Private Sub CambioArticulo(ByVal ColumnName As String, _
    '                               ByVal value As Object, ByVal current As IPropertyAccessor, _
    '                               ByVal context As IPropertyAccessor, _
    '                               ByVal services As ServiceProvider)

    '        Dim brd As New BusinessRuleData(ColumnName, value, current, context)
    '        ProcesoComercial.CambioArticulo(brd, services)
    '    End Sub

    '    Private Sub CambioCantidad(ByVal current As IPropertyAccessor, ByVal Value As Object, ByVal context As IPropertyAccessor, ByVal services As ServiceProvider)
    '        Dim brd As New BusinessRuleData("Cantidad", Value, current, context)
    '        ProcesoComercial.CambioCantidad(brd, services)
    '    End Sub

    '    Public Sub CambioPrecio(ByVal ColumnName As String, ByVal value As Object, ByVal current As IPropertyAccessor, ByVal context As IPropertyAccessor, ByVal services As ServiceProvider)
    '        Dim brd As New BusinessRuleData(ColumnName, value, current, context)
    '        General.CambioPrecio(brd, services)
    '    End Sub

    '    Private Sub CambioNObra(ByVal current As IPropertyAccessor, ByVal Value As Object, ByVal context As IPropertyAccessor, ByVal services As ServiceProvider)
    '        Dim brd As New BusinessRuleData("NObra", Value, current, context)
    '        ProcesoComercial.CambioNObra(brd, services)
    '    End Sub

    '    Private Sub CambioCodTrabajo(ByVal current As IPropertyAccessor, ByVal value As Object)
    '        Dim brd As New BusinessRuleData("CodTrabajo", value, current, Nothing)
    '        ProcesoComercial.CambioNObra(brd, New ServiceProvider)
    '    End Sub

    '    Public Sub CambioRegalo(ByVal current As IPropertyAccessor, ByVal value As Object, ByVal context As IPropertyAccessor, ByVal services As ServiceProvider)
    '        Dim brd As New BusinessRuleData("Regalo", value, current, context)
    '        ProcesoComercial.CambioRegalo(brd, services)
    '    End Sub

    '    'Friend Sub CambioLineaAlbaran(ByVal current As IPropertyAccessor, ByVal Value As Object, ByVal services As ServiceProvider)
    '    '    Dim brd As New BusinessRuleData("IDLineaAlbaran", Value, current, Nothing)
    '    '    ProcesoComercial.CambioLineaAlbaran(brd, services)
    '    'End Sub

    '    Friend Sub CambioQInterna(ByVal current As IPropertyAccessor, ByVal value As Object, ByVal context As IPropertyAccessor, ByVal services As ServiceProvider)
    '        Dim brd As New BusinessRuleData("QInterna", value, current, context)
    '        General.CalculoQInterna(brd, services)
    '    End Sub

    '    Friend Sub CambioFactor(ByVal current As IPropertyAccessor, ByVal Value As Object, ByVal Context As IPropertyAccessor, ByVal services As ServiceProvider)
    '        Dim brd As New BusinessRuleData("QInterna", Value, current, Context)
    '        General.CambioFactor(brd, services)
    '    End Sub

    '    Friend Sub CalcularImportes(ByVal ColumnName As String, ByVal current As IPropertyAccessor, ByVal Value As Object, ByVal Context As IPropertyAccessor, ByVal services As ServiceProvider)
    '        Dim brd As New BusinessRuleData(ColumnName, Value, current, Context)
    '        General.CalcularImportes(brd, services)
    '    End Sub

    '#End Region



    '#Region " CalculoRepresentantes "

    '    Public Shared Function CalculoRepresentantes(ByVal ArtInfo As ArticuloInfo, ByVal ClteInfo As ClienteInfo, ByVal IDObra As Integer, ByVal Cantidad As Double) As DataTable
    '        Dim intOrden As Integer
    '        Dim dtRepresentanteTot As DataTable
    '        Dim dtRepresentante As DataTable
    '        'Dim dtArticulo As DataTable
    '        'Funcionamiento de las comisiones a representantes
    '        ' 1º Ver si hay comisión por obra
    '        ' 2º Ver si hay comisión por cliente - Artículo
    '        ' 3º Ver si hay comisión por cliente cuando el campo comisión sea <>0
    '        ' 4º Ver si hay comisión por la combinación de cliente con comisión = 0 y por Tipo / Familia / Cantidad
    '        ' 5º Ver si hay comisión por la combinación de cliente con comisión = 0 y por Tipo / Cantidad
    '        ' 6º Ver si hay comisión por Zona / Tipo / Familia
    '        ' 7º Ver si hay comisión por Zona 

    '        '1º.-Miramos en la tabla Obra Representante si hay algún representante para esa obra
    '        intOrden = 1
    '        ' 1º Obra-Artículo
    '        'If Length(IDArticulo) > 0 Then dtArticulo = New Articulo().SelOnPrimaryKey(IDArticulo)
    '        Dim objFilter As New Filter
    '        If IDObra > 0 AndAlso Not IsNothing(ArtInfo) AndAlso Length(ArtInfo.IDArticulo) > 0 Then
    '            objFilter.Add(New NumberFilterItem("IDObra", IDObra))
    '            objFilter.Add(New StringFilterItem("IDArticulo", ArtInfo.IDArticulo))
    '            Dim ORepresentante As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraRepresentante")
    '            dtRepresentanteTot = ORepresentante.Filter(objFilter)

    '            If Not dtRepresentanteTot Is Nothing AndAlso dtRepresentanteTot.Rows.Count > 0 Then
    '                AñadirComisiones(dtRepresentanteTot, dtRepresentante, intOrden)
    '            Else
    '                'If Not dtArticulo Is Nothing AndAlso dtArticulo.Rows.Count > 0 Then
    '                If Length(ArtInfo.IDSubFamilia) > 0 Then
    '                    objFilter.Clear()
    '                    objFilter.Add(New NumberFilterItem("IDObra", IDObra))
    '                    objFilter.Add(New StringFilterItem("IDTipo", ArtInfo.IDTipo))
    '                    objFilter.Add(New StringFilterItem("IDFamilia", ArtInfo.IDFamilia))
    '                    objFilter.Add(New StringFilterItem("IDSubfamilia", ArtInfo.IDSubFamilia))
    '                    dtRepresentanteTot = ORepresentante.Filter(objFilter)
    '                    If Not dtRepresentanteTot Is Nothing AndAlso dtRepresentanteTot.Rows.Count > 0 Then
    '                        AñadirComisiones(dtRepresentanteTot, dtRepresentante, intOrden)
    '                    Else
    '                        objFilter.Clear()
    '                        objFilter.Add(New NumberFilterItem("IDObra", IDObra))
    '                        objFilter.Add(New StringFilterItem("IDTipo", ArtInfo.IDTipo))
    '                        objFilter.Add(New StringFilterItem("IDFamilia", ArtInfo.IDFamilia))
    '                        dtRepresentanteTot = ORepresentante.Filter(objFilter)
    '                        If Not dtRepresentanteTot Is Nothing AndAlso dtRepresentanteTot.Rows.Count > 0 Then
    '                            AñadirComisiones(dtRepresentanteTot, dtRepresentante, intOrden)
    '                        Else
    '                            objFilter.Clear()
    '                            objFilter.Add(New NumberFilterItem("IDObra", IDObra))
    '                            objFilter.Add(New StringFilterItem("IDTipo", ArtInfo.IDTipo))
    '                            dtRepresentanteTot = ORepresentante.Filter(objFilter)
    '                            If Not dtRepresentanteTot Is Nothing AndAlso dtRepresentanteTot.Rows.Count > 0 Then
    '                                AñadirComisiones(dtRepresentanteTot, dtRepresentante, intOrden)
    '                            Else
    '                                objFilter.Clear()
    '                                objFilter.Add(New NumberFilterItem("IDObra", IDObra))
    '                                dtRepresentanteTot = ORepresentante.Filter(objFilter)
    '                                If Not dtRepresentanteTot Is Nothing AndAlso dtRepresentanteTot.Rows.Count > 0 Then
    '                                    AñadirComisiones(dtRepresentanteTot, dtRepresentante, intOrden)
    '                                End If
    '                            End If
    '                        End If
    '                    End If
    '                End If
    '                'End If
    '            End If
    '        End If

    '        '2º.- Miramos en la tabla ClienteRepresentante si hay algún representante para ese cliente y ese artículo
    '        intOrden = 2
    '        Dim ClteRepresentante As New ClienteRepresentante
    '        If Not IsNothing(ClteInfo) AndAlso Length(ClteInfo.IDCliente) > 0 AndAlso Not IsNothing(ArtInfo) AndAlso Length(ArtInfo.IDArticulo) > 0 Then
    '            objFilter.Clear()
    '            objFilter.Add(New StringFilterItem("IDCliente", ClteInfo.IDCliente))
    '            objFilter.Add(New StringFilterItem("IDArticulo", ArtInfo.IDArticulo))
    '            dtRepresentanteTot = ClteRepresentante.Filter(objFilter)
    '            If Not dtRepresentanteTot Is Nothing AndAlso dtRepresentanteTot.Rows.Count > 0 Then
    '                AñadirComisiones(dtRepresentanteTot, dtRepresentante, intOrden)
    '            End If
    '        End If

    '        '3º.- Si no hay, consultamos la misma tabla, pero sin el artículo explícito
    '        intOrden = 3
    '        If Not IsNothing(ClteInfo) AndAlso Length(ClteInfo.IDCliente) > 0 Then
    '            objFilter.Clear()
    '            objFilter.Add(New StringFilterItem("IDCliente", ClteInfo.IDCliente))
    '            objFilter.Add(New IsNullFilterItem("IDArticulo"))
    '            objFilter.Add(New NumberFilterItem("Comision", FilterOperator.NotEqual, 0))
    '            dtRepresentanteTot = ClteRepresentante.Filter(objFilter)
    '            If Not dtRepresentanteTot Is Nothing AndAlso dtRepresentanteTot.Rows.Count > 0 Then
    '                AñadirComisiones(dtRepresentanteTot, dtRepresentante, intOrden)
    '            End If
    '        End If

    '        '4º Se incorpora la búsqueda por comisión tipo- familia cantidad
    '        intOrden = 4
    '        'Buscamos el Tipo y la Familia del Artículo y cantidad seleccionado
    '        If Not IsNothing(ClteInfo) AndAlso Length(ClteInfo.IDCliente) > 0 Then
    '            If Not IsNothing(ArtInfo) Then
    '                objFilter.Clear()
    '                objFilter.Add(New StringFilterItem("IDCliente", ClteInfo.IDCliente))
    '                objFilter.Add(New StringFilterItem("IDTipo", ArtInfo.IDTipo))
    '                objFilter.Add(New StringFilterItem("IDFamilia", ArtInfo.IDFamilia))
    '                objFilter.Add(New NumberFilterItem("QDesde", FilterOperator.LessThanOrEqual, Cantidad))

    '                dtRepresentanteTot = AdminData.GetData("ComisionRepresentante", objFilter, , "IDRepresentante asc, IDTipo asc, IDFamilia asc, QDesde desc")
    '                If Not dtRepresentanteTot Is Nothing AndAlso dtRepresentanteTot.Rows.Count > 0 Then
    '                    AñadirComisiones(dtRepresentanteTot, dtRepresentante, intOrden)
    '                End If

    '                '5º.- Si no hay, buscamos en la misma tabla los registros sin  familia y con cantidad
    '                intOrden = 5
    '                objFilter.Clear()
    '                objFilter.Add(New StringFilterItem("IDCliente", ClteInfo.IDCliente))
    '                objFilter.Add(New StringFilterItem("IDTipo", ArtInfo.IDTipo))
    '                objFilter.Add(New NumberFilterItem("QDesde", FilterOperator.LessThanOrEqual, Cantidad))
    '                dtRepresentanteTot = AdminData.GetData("ComisionRepresentante", objFilter, , "IDRepresentante asc, IDTipo asc, IDFamilia asc, QDesde desc")
    '                If Not dtRepresentanteTot Is Nothing AndAlso dtRepresentanteTot.Rows.Count > 0 Then
    '                    AñadirComisiones(dtRepresentanteTot, dtRepresentante, intOrden)
    '                End If
    '            End If

    '            Dim ZonaRepres As New ZonaRepresentante
    '            If Length(ClteInfo.Zona) > 0 AndAlso Not IsNothing(ArtInfo) Then
    '                '6º.- Si no hay, consultamos la tabla ZonaRepresentante con el Tipo y Familia del artículo de la línea
    '                'Buscamos la Zona del Cliente
    '                intOrden = 6
    '                objFilter.Clear()
    '                objFilter.Add(New StringFilterItem("IDZona", ClteInfo.Zona))
    '                objFilter.Add(New StringFilterItem("IDTipo", ArtInfo.IDTipo))
    '                objFilter.Add(New StringFilterItem("IDFamilia", ArtInfo.IDFamilia))
    '                dtRepresentanteTot = ZonaRepres.Filter(objFilter)
    '                If Not dtRepresentanteTot Is Nothing AndAlso dtRepresentanteTot.Rows.Count > 0 Then
    '                    AñadirComisiones(dtRepresentanteTot, dtRepresentante, intOrden)
    '                End If
    '            End If

    '            If Length(ClteInfo.Zona) > 0 Then
    '                '7º.- Si no hay, buscamos en la misma tabla (ZonaRepresentante) los registros sin tipo ni familia
    '                intOrden = 7
    '                objFilter.Clear()
    '                objFilter.Add(New StringFilterItem("IDZona", ClteInfo.Zona))
    '                objFilter.Add(New IsNullFilterItem("IDTipo"))
    '                objFilter.Add(New IsNullFilterItem("IDFamilia"))
    '                dtRepresentanteTot = ZonaRepres.Filter(objFilter)
    '                If Not dtRepresentanteTot Is Nothing AndAlso dtRepresentanteTot.Rows.Count > 0 Then
    '                    AñadirComisiones(dtRepresentanteTot, dtRepresentante, intOrden)
    '                End If
    '            End If
    '        End If

    '        CalculoRepresentantes = dtRepresentante
    '    End Function


    '    Private Shared Function CrearDTRepresentante() As DataTable
    '        Dim dt As New DataTable
    '        dt.Columns.Add("IDRepresentante", GetType(String))
    '        dt.Columns.Add("Comision", GetType(Double))
    '        dt.Columns.Add("Importe", GetType(Double))
    '        dt.Columns.Add("Porcentaje", GetType(Boolean))
    '        dt.Columns.Add("Orden", GetType(Integer))
    '        Return dt
    '    End Function

    '    Private Shared Sub AñadirComisiones(ByVal dtRepresentanteTot, ByRef dtRepresentante, ByVal intOrden)
    '        Dim strRepresentante As String

    '        If dtRepresentante Is Nothing Then
    '            dtRepresentante = CrearDTRepresentante()
    '        End If
    '        For Each dr As DataRow In dtRepresentanteTot.Rows
    '            If dr("IDrepresentante") <> strRepresentante Then
    '                Dim drr As DataRow = dtRepresentante.NewRow
    '                drr("IDRepresentante") = dr("IDrepresentante")
    '                drr("Comision") = dr("Comision")
    '                drr("Porcentaje") = dr("Porcentaje")
    '                drr("orden") = intOrden
    '                dtRepresentante.Rows.Add(drr)
    '                strRepresentante = dr("IDrepresentante")
    '            End If
    '        Next
    '    End Sub



    '    Public Function RepresentantesCommonBusinessRules(ByVal ColumnName As String, ByVal Value As Object, ByVal current As IPropertyAccessor, Optional ByVal Context As IPropertyAccessor = Nothing) As IPropertyAccessor
    '        Dim IncrementoImporte As Decimal
    '        Dim RecalcularImportesAyB As Boolean

    '        Select Case ColumnName
    '            Case "IDRepresentante"
    '                If Len(Value & String.Empty) Then
    '                    Dim comision As Hashtable
    '                    comision = ComisionRepresentante(Value, Context("IDCliente"), Context("IDArticulo"))
    '                    If Not comision Is Nothing Then
    '                        current("Comision") = comision("Comision")
    '                        current("Porcentaje") = comision("Porcentaje")
    '                    End If
    '                    If current("Porcentaje") Then
    '                        current = Me.RepresentantesCommonBusinessRules("Comision", current("Comision"), current, Context)
    '                    Else
    '                        current = Me.RepresentantesCommonBusinessRules("Importe", current("Importe"), current, Context)
    '                    End If
    '                    RecalcularImportesAyB = True
    '                End If
    '            Case "Importe"
    '                If Not IsNumeric(Value) Then Value = 0
    '                If Not IsNumeric(current("Comision")) Then current("Comision") = 0
    '                If Not IsNumeric(current("Importe")) Then current("Importe") = 0

    '                IncrementoImporte = Value - current("Importe")
    '                If Context("ImporteLinea") < 0 Then
    '                    If Context("SumaImporte") + IncrementoImporte > Context("ImporteLinea") Then
    '                        If Nz(Context("ImporteLinea"), 0) <> 0 Then
    '                            current("Comision") = 100 * (Value / Context("ImporteLinea"))
    '                        End If
    '                    Else
    '                        ApplicationService.GenerateError("Los importes asignados a los representantes superan el importe total de la línea.")
    '                    End If
    '                Else
    '                    If Context("SumaImporte") + IncrementoImporte > Context("ImporteLinea") Then
    '                        ApplicationService.GenerateError("Los importes asignados a los representantes superan el importe total de la línea.")
    '                    Else
    '                        If Nz(Context("ImporteLinea"), 0) <> 0 Then
    '                            current("Comision") = 100 * (Value / Context("ImporteLinea"))
    '                        End If
    '                    End If
    '                End If
    '                current("Importe") = Value
    '                RecalcularImportesAyB = True
    '            Case "Comision"
    '                If Not IsNumeric(Value) Then Value = 0
    '                If Not IsNumeric(current("Comision")) Then current("Comision") = 0
    '                If Not IsNumeric(current("Importe")) Then current("Importe") = 0

    '                If Value < 0 Or Value > 100 Then
    '                    ApplicationService.GenerateError("El valor de la comisión debe estar entre [0-100]%.")
    '                Else
    '                    Dim importe As Decimal
    '                    importe = Context("ImporteLinea") * (Value / 100)

    '                    IncrementoImporte = importe - current("Importe")
    '                    If Context("ImporteLinea") < 0 Then
    '                        If Context("SumaImporte") + IncrementoImporte > Context("ImporteLinea") Then
    '                            current("Importe") = importe
    '                        Else
    '                            ApplicationService.GenerateError("El importe total asignado a los representantes superan el importe total de la línea.")
    '                        End If
    '                    Else
    '                        If Context("SumaImporte") + IncrementoImporte > Context("ImporteLinea") Then
    '                            ApplicationService.GenerateError("El importe total asignado a los representantes superan el importe total de la línea.")
    '                        Else
    '                            current("Importe") = importe
    '                        End If
    '                    End If
    '                End If
    '                RecalcularImportesAyB = True
    '            Case "Porcentaje"
    '                If Value Then
    '                    current = Me.RepresentantesCommonBusinessRules("Comision", current("Comision"), current, Context)
    '                Else
    '                    current = Me.RepresentantesCommonBusinessRules("Importe", current("Importe"), current, Context)
    '                End If
    '        End Select

    '        If RecalcularImportesAyB Then
    '            If Context.ContainsKey("IDMoneda") Then
    '                If Context.ContainsKey("Fecha") Then
    '                    current = General.MantenimientoValoresAyB(current, Context("IDMoneda"), Context("Fecha"))
    '                Else
    '                    current = General.MantenimientoValoresAyB(current, Context("IDMoneda"))
    '                End If
    '            End If
    '        End If

    '        Return current
    '    End Function

    '    Public Function ComisionRepresentante(ByVal strIDRepresentante As String, ByVal strIDCliente As String, Optional ByVal strIDArticulo As String = Nothing) As Hashtable
    '        If Len(strIDRepresentante) And Len(strIDCliente) Then
    '            Dim data As New Hashtable
    '            data("Comision") = 0
    '            data("Porcentaje") = False

    '            Dim f As New Filter
    '            f.Add(New StringFilterItem("IDCliente", FilterOperator.Equal, strIDCliente))
    '            f.Add(New StringFilterItem("IDRepresentante", FilterOperator.Equal, strIDRepresentante))

    '            Dim cr As New ClienteRepresentante
    '            Dim dt As DataTable
    '            dt = cr.Filter(f, "IDArticulo DESC")
    '            f.Clear()
    '            If Not IsNothing(dt) AndAlso dt.Rows.Count Then
    '                Dim ComisionPorDefecto As Boolean = True

    '                If Len(strIDArticulo) Then
    '                    f.Add("IDArticulo", FilterOperator.Equal, strIDArticulo, FilterType.String)
    '                    dt.DefaultView.RowFilter = f.Compose(New AdoFilterComposer)
    '                    If dt.DefaultView.Count > 0 Then
    '                        ComisionPorDefecto = False
    '                        data("Comision") = dt.DefaultView(0).Row("Comision")
    '                        data("Porcentaje") = dt.DefaultView(0).Row("Porcentaje")
    '                    End If
    '                End If

    '                If ComisionPorDefecto Then
    '                    f.Add(New IsNullFilterItem("IDArticulo"))
    '                    dt.DefaultView.RowFilter = f.Compose(New AdoFilterComposer)
    '                    If dt.DefaultView.Count > 0 Then
    '                        data("Comision") = dt.DefaultView(0).Row("Comision")
    '                        data("Porcentaje") = dt.DefaultView(0).Row("Porcentaje")
    '                    End If
    '                End If
    '            End If

    '            Return data
    '        End If
    '    End Function

    '#Region " ActualizarRepresentantes "

    '    Public Function ActualizarRepresentantes(ByVal dr As DataRow) As DataTable
    '        Dim blnCambioArticulo, blnCambioImporte, blnCambioCantidad As Boolean
    '        Dim comision As Double
    '        Dim Representantes As DataTable

    '        If Not IsNothing(dr) Then
    '            blnCambioArticulo = (dr("IDArticulo") & vbNullString <> dr("IDArticulo", DataRowVersion.Original) & vbNullString)
    '            blnCambioImporte = (dr("Importe") <> dr("Importe", DataRowVersion.Original))
    '            blnCambioCantidad = (dr("QInterna") <> dr("QInterna", DataRowVersion.Original))

    '            If blnCambioArticulo Or blnCambioImporte Then
    '                Select Case dr.Table.TableName
    '                    Case "FacturaVentaLinea"
    '                        Dim r As New FacturaVentaRepresentante
    '                        Representantes = r.Filter(New NumberFilterItem("IDLineaFactura", FilterOperator.Equal, dr("IDLineaFactura")))
    '                    Case "AlbaranVentaLinea"
    '                        Dim r As New AlbaranVentaRepresentante
    '                        Representantes = r.Filter(New NumberFilterItem("IDLineaAlbaran", FilterOperator.Equal, dr("IDLineaAlbaran")))
    '                    Case "PedidoVentaLinea"
    '                        Dim r As New PedidoVentaRepresentante
    '                        Representantes = r.Filter(New NumberFilterItem("IDLineaPedido", FilterOperator.Equal, dr("IDLineaPedido")))
    '                End Select

    '                If Not IsNothing(Representantes) Then
    '                    If Representantes.Rows.Count > 0 Then
    '                        If blnCambioImporte And (Not blnCambioArticulo And Not blnCambioCantidad) Then
    '                            If blnCambioImporte And Not blnCambioArticulo Then
    '                                Dim IDMoneda As String
    '                                Dim CambioA, CambioB As Double
    '                                Dim dt As DataTable
    '                                Select Case dr.Table.TableName
    '                                    Case "FacturaVentaLinea"
    '                                        Dim c As New FacturaVentaCabecera
    '                                        dt = c.Filter("IDMoneda,CambioA,CambioB", "IDFactura=" & dr("IDFactura"))
    '                                    Case "AlbaranVentaLinea"
    '                                        Dim c As New AlbaranVentaCabecera
    '                                        dt = c.Filter("IDMoneda,CambioA,CambioB", "IDAlbaran=" & dr("IDAlbaran"))
    '                                    Case "PedidoVentaLinea"
    '                                        Dim c As New PedidoVentaCabecera
    '                                        dt = c.Filter("IDMoneda,CambioA,CambioB", "IDPedido=" & dr("IDPedido"))
    '                                End Select
    '                                If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
    '                                    IDMoneda = dt.Rows(0)("IDMoneda")
    '                                    CambioA = dt.Rows(0)("CambioA")
    '                                    CambioB = dt.Rows(0)("CambioB")
    '                                    For Each lineaRepresentante As DataRow In Representantes.Rows
    '                                        comision = lineaRepresentante("Comision")
    '                                        If Not lineaRepresentante("Porcentaje") Then
    '                                            If dr("ImporteA") <> 0 Then comision = lineaRepresentante("ImporteA") * 100 / dr("ImporteA")
    '                                        End If
    '                                        lineaRepresentante("Comision") = comision
    '                                        lineaRepresentante("Importe") = (dr("Importe") * comision) / 100

    '                                        MantenimientoValoresAyB(lineaRepresentante, IDMoneda, CambioA, CambioB)
    '                                    Next
    '                                End If
    '                            End If
    '                        Else
    '                            'Si se modifica el artículo o la cantidad, se elimina el desglose anterior y se vuelve a calcular

    '                            For Each Representante As DataRow In Representantes.Select
    '                                Representante.Delete()
    '                            Next
    '                            Dim DtAux As DataTable = NuevoRepresentante(dr)
    '                            If Not IsNothing(DtAux) AndAlso DtAux.Rows.Count Then
    '                                For Each newrow As DataRow In DtAux.Rows
    '                                    Representantes.Rows.Add(newrow.ItemArray)
    '                                Next
    '                            End If
    '                        End If
    '                    Else
    '                        Representantes = NuevoRepresentante(dr)
    '                    End If
    '                End If

    '                Return Representantes
    '            End If
    '        End If
    '    End Function

    '    Public Function ActualizarRepresentantes(ByVal dt As DataTable) As DataTable()
    '        Dim carrier(-1) As DataTable
    '        If Not IsNothing(dt) Then
    '            For Each dr As DataRow In dt.Rows
    '                Dim r As DataTable
    '                r = ActualizarRepresentantes(dr)
    '                If Not IsNothing(r) Then
    '                    ReDim Preserve carrier(UBound(carrier) + 1)
    '                    carrier(UBound(carrier)) = r
    '                End If
    '            Next
    '        End If

    '        Return carrier
    '    End Function

    '#End Region

    '#Region " NuevoRepresentante "

    '#Region " NuevoRepresentante Document"

    '    Public Shared Sub NuevoRepresentante(ByVal oDoc As Document, ByVal services As ServiceProvider)
    '        If Not oDoc Is Nothing Then
    '            Dim MarcaRepresentante As Boolean
    '            Dim ClteInfo As ClienteInfo
    '            Dim PKLinea, strEntidadLinea, strEntidadRepresentante As String
    '            Dim Moneda, MonedaA, MonedaB As MonedaInfo
    '            If TypeOf oDoc Is DocumentoFacturaVenta Then
    '                MarcaRepresentante = True
    '                ClteInfo = CType(oDoc, DocumentoFacturaVenta).Cliente
    '                Moneda = CType(oDoc, DocumentoFacturaVenta).Moneda
    '                MonedaA = CType(oDoc, DocumentoFacturaVenta).MonedaA
    '                MonedaB = CType(oDoc, DocumentoFacturaVenta).MonedaB
    '                PKLinea = "IDLineaFactura"
    '                strEntidadLinea = GetType(FacturaVentaLinea).Name
    '                strEntidadRepresentante = GetType(FacturaVentaRepresentante).Name
    '            ElseIf TypeOf oDoc Is DocumentoAlbaranVenta Then
    '                ClteInfo = CType(oDoc, DocumentoAlbaranVenta).Cliente
    '                Moneda = CType(oDoc, DocumentoAlbaranVenta).Moneda
    '                MonedaA = CType(oDoc, DocumentoAlbaranVenta).MonedaA
    '                MonedaB = CType(oDoc, DocumentoAlbaranVenta).MonedaB
    '                PKLinea = "IDLineaAlbaran"
    '                strEntidadLinea = GetType(AlbaranVentaLinea).Name
    '                strEntidadRepresentante = GetType(AlbaranVentaRepresentante).Name
    '                'ElseIf TypeOf oDoc Is DocumentoPedidoVenta Then
    '                '    ClteInfo = CType(oDoc, DocumentoPedidoVenta).Cliente
    '                '    Moneda = CType(oDoc, DocumentoPedidoVenta).Moneda
    '                '    MonedaA = CType(oDoc, DocumentoPedidoVenta).MonedaA
    '                '    MonedaB = CType(oDoc, DocumentoPedidoVenta).MonedaB
    '                '    PKLinea = "IDLineaPedido"
    '                '    strEntidadLinea = GetType(PedidoVentaLinea).Name
    '                '    strEntidadRepresentante = GetType(PedidoVentaRepresentante).Name
    '            End If

    '            Dim f As New Filter
    '            If oDoc(strEntidadLinea) Is Nothing OrElse oDoc(strEntidadLinea).Rows.Count = 0 Then Exit Sub
    '            For Each drLinea As DataRow In oDoc(strEntidadLinea).Rows
    '                Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(GetType(EntityInfoCache(Of ArticuloInfo)))
    '                Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(drLinea("IDArticulo"))
    '                'filtro = PrepararFiltrosRepresentantes(ClteInfo.IDCliente, dr("IDArticulo"), dr("QInterna"), intObra)
    '                'If Not IsNothing(filtro) AndAlso filtro.Rows.Count Then
    '                Dim intObra As Integer = 0
    '                If Length(drLinea("IDObra")) > 0 Then
    '                    intObra = drLinea("IDObra")
    '                End If

    '                '//Cogemos los Representantes definidos ya para el IDLinea____, si los hay.
    '                Dim Representantes As DataTable = oDoc(strEntidadRepresentante).Clone
    '                f.Clear()
    '                f.Add(New NumberFilterItem(PKLinea, drLinea(PKLinea)))
    '                For Each drRepresentante As DataRow In oDoc(strEntidadRepresentante).Select(f.Compose(New AdoFilterComposer))
    '                    Representantes.Rows.Add(drRepresentante.ItemArray)
    '                Next

    '                '//Quitamos los Representantes que teníamos, para el IDLinea____
    '                For Each drRepresentante As DataRow In oDoc(strEntidadRepresentante).Select(f.Compose(New AdoFilterComposer))
    '                    drRepresentante.Delete()
    '                Next

    '                Dim strOrder As String
    '                If Representantes Is Nothing OrElse Representantes.Rows.Count = 0 Then
    '                    '//Recuperamos los Representantes de Obras, Clientes, .....
    '                    Representantes = CalculoRepresentantes(ArtInfo, ClteInfo, intObra, Nz(drLinea("QInterna"), 0))
    '                    strOrder = "IDRepresentante,Orden"
    '                Else
    '                    strOrder = "IDRepresentante"
    '                End If

    '                f.Clear()
    '                f.Add(New NumberFilterItem("Comision", FilterOperator.NotEqual, 0))
    '                Dim strRepresentante As String
    '                If Representantes Is Nothing OrElse Representantes.Rows.Count = 0 Then Exit Sub
    '                For Each r As DataRow In Representantes.Select(f.Compose(New AdoFilterComposer), strOrder)
    '                    If strRepresentante <> r("IDRepresentante") Then
    '                        Dim newrow As DataRow = oDoc(strEntidadRepresentante).NewRow()
    '                        strRepresentante = r("IDRepresentante")
    '                        If MarcaRepresentante Then
    '                            newrow("MarcaRepresentante") = AdminData.GetAutoNumeric
    '                        End If

    '                        newrow(PKLinea) = drLinea(PKLinea)
    '                        newrow("IDRepresentante") = r("IDRepresentante")

    '                        Dim comision As Double = 0
    '                        Dim dblImporteRepr As Double = 0
    '                        If r("porcentaje") Then         '//Comisión por Porcentaje
    '                            comision = r("Comision")
    '                            If drLinea("Importe") <> 0 Then dblImporteRepr = (drLinea("Importe") * comision) / 100
    '                        Else                            '//Comisión de Importe fijo
    '                            If drLinea("Importe") <> 0 Then
    '                                dblImporteRepr = Nz(r("Importe"), 0)
    '                                comision = Nz(r("Importe"), 0) * 100 / drLinea("Importe")
    '                            End If
    '                        End If
    '                        newrow("porcentaje") = Nz(r("porcentaje"), 0)
    '                        newrow("Importe") = dblImporteRepr
    '                        newrow("Comision") = comision

    '                        MantenimientoValoresAyB(newrow, Moneda, MonedaA, MonedaB)

    '                        oDoc(strEntidadRepresentante).Rows.Add(newrow.ItemArray)
    '                    End If
    '                Next

    '            Next
    '        End If
    '    End Sub

    '#End Region

    '    Public Function NuevoRepresentante(ByVal dt As DataTable, Optional ByVal Monedas As MonedaCache = Nothing) As DataTable()

    '        Dim carrier(-1) As DataTable
    '        If Not IsNothing(dt) Then
    '            If Monedas Is Nothing Then Monedas = New MonedaCache
    '            For Each dr As DataRow In dt.Rows
    '                Dim r As DataTable = NuevoRepresentante(dr, Monedas)
    '                If Not IsNothing(r) Then
    '                    ReDim Preserve carrier(UBound(carrier) + 1)
    '                    carrier(UBound(carrier)) = r
    '                End If
    '            Next
    '        End If

    '        Return carrier
    '    End Function

    '    Public Function NuevoRepresentante(ByVal dt As DataTable, _
    '            ByVal IDCliente As String, _
    '            ByVal IDMoneda As String, _
    '            ByVal CambioA As Double, _
    '            ByVal CambioB As Double, _
    '            ByVal Monedas As MonedaCache) As DataTable()

    '        Dim carrier(-1) As DataTable
    '        If Not IsNothing(dt) Then
    '            For Each dr As DataRow In dt.Rows
    '                Dim r As DataTable = NuevoRepresentante(dr, IDCliente, IDMoneda, CambioA, CambioB, Monedas)
    '                If Not IsNothing(r) Then
    '                    ReDim Preserve carrier(UBound(carrier) + 1)
    '                    carrier(UBound(carrier)) = r
    '                End If
    '            Next
    '        End If

    '        Return carrier
    '    End Function

    '    Public Function NuevoRepresentante(ByVal dr As DataRow, Optional ByVal Monedas As MonedaCache = Nothing) As DataTable
    '        If Not IsNothing(dr) Then
    '            If Not Monedas Is Nothing Then Monedas = New MonedaCache
    '            Dim dt As DataTable
    '            Select Case dr.Table.TableName
    '                Case "FacturaVentaLinea"
    '                    Dim c As New FacturaVentaCabecera
    '                    dt = c.SelOnPrimaryKey(dr("IDFactura"))
    '                Case "AlbaranVentaLinea"
    '                    Dim c As New AlbaranVentaCabecera
    '                    dt = c.SelOnPrimaryKey(dr("IDAlbaran"))
    '                Case "PedidoVentaLinea"
    '                    Dim c As New PedidoVentaCabecera
    '                    dt = c.SelOnPrimaryKey(dr("IDPedido"))
    '            End Select

    '            If Not IsNothing(dt) AndAlso dt.Rows.Count Then
    '                Return NuevoRepresentante(dr, dt.Rows(0)("IDCliente"), dt.Rows(0)("IDMoneda"), dt.Rows(0)("CambioA"), dt.Rows(0)("CambioB"), Monedas)
    '            End If
    '        End If
    '    End Function

    '    Public Function NuevoRepresentante(ByVal dr As DataRow, ByVal IDCliente As String, ByVal IDMoneda As String, ByVal CambioA As Double, ByVal CambioB As Double, ByVal Monedas As MonedaCache) As DataTable
    '        If CambioA <= 0 Then CambioA = 1
    '        If CambioB <= 0 Then CambioB = 1
    '        If IsNothing(Monedas) Then Monedas = New MonedaCache
    '        Dim MonInfo As MonedaInfo = Monedas.GetMoneda(IDMoneda)
    '        MonInfo.CambioA = CambioA
    '        MonInfo.CambioB = CambioB

    '        Return NuevoRepresentante(dr, IDCliente, MonInfo, Monedas.MonedaA, Monedas.MonedaB)
    '    End Function
    '    Public Function NuevoRepresentante(ByVal dr As DataRow, ByVal ArtInfo As ArticuloInfo, ByVal ClteInfo As ClienteInfo, ByVal IDMoneda As String, ByVal CambioA As Double, ByVal CambioB As Double, ByVal Monedas As MonedaCache) As DataTable
    '        If CambioA <= 0 Then CambioA = 1
    '        If CambioB <= 0 Then CambioB = 1
    '        If IsNothing(Monedas) Then Monedas = New MonedaCache
    '        Dim MonInfo As MonedaInfo = Monedas.GetMoneda(IDMoneda)
    '        MonInfo.CambioA = CambioA
    '        MonInfo.CambioB = CambioB

    '        Return NuevoRepresentante(dr, ArtInfo, ClteInfo, MonInfo, Monedas.MonedaA, Monedas.MonedaB)
    '    End Function

    '    Public Function NuevoRepresentante(ByVal dr As DataRow, ByVal IDCliente As String, ByVal Moneda As MonedaInfo, ByVal MonedaA As MonedaInfo, ByVal MonedaB As MonedaInfo) As DataTable
    '        Dim ClteInfo As ClienteInfo
    '        Dim ArtInfo As ArticuloInfo
    '        If Length(IDCliente) > 0 Then ClteInfo = New EntityInfoCache(Of ClienteInfo)().GetEntity(IDCliente) 'ClteInfo = New Cliente().InformacionCliente(IDCliente)
    '        If Length(dr("IDArticulo")) > 0 Then ArtInfo = New Articulo().InformacionArticulo(dr("IDArticulo"))

    '        Return NuevoRepresentante(dr, ArtInfo, ClteInfo, Moneda, MonedaA, MonedaB)
    '    End Function

    '    Public Function NuevoRepresentante(ByVal dr As DataRow, ByVal ArtInfo As ArticuloInfo, ByVal ClteInfo As ClienteInfo, ByVal Moneda As MonedaInfo, ByVal MonedaA As MonedaInfo, ByVal MonedaB As MonedaInfo) As DataTable
    '        Dim MarcaRepresentante As Boolean
    '        Dim comision As Double
    '        Dim PKLinea As String
    '        Dim newData As DataTable
    '        Dim admin As New AdminData
    '        Dim intObra As Integer

    '        If Not IsNothing(dr) Then
    '            Select Case dr.Table.TableName
    '                Case "FacturaVentaLinea"
    '                    PKLinea = "IDLineaFactura"
    '                    Dim representante As New FacturaVentaRepresentante
    '                    newData = representante.AddNew()
    '                    MarcaRepresentante = True
    '                    If Length(dr("IDObra")) > 0 Then
    '                        intObra = dr("IDObra")
    '                    End If
    '                Case "AlbaranVentaLinea"
    '                    PKLinea = "IDLineaAlbaran"
    '                    Dim representante As New AlbaranVentaRepresentante
    '                    newData = representante.AddNew()

    '                Case "PedidoVentaLinea"
    '                    PKLinea = "IDLineaPedido"
    '                    Dim representante As New PedidoVentaRepresentante
    '                    newData = representante.AddNew()
    '            End Select
    '            If IsNothing(Moneda) Then Moneda = New MonedaInfo
    '            If Not IsNothing(newData) Then
    '                If Len(ClteInfo.IDCliente) > 0 Then

    '                    'filtro = PrepararFiltrosRepresentantes(ClteInfo.IDCliente, dr("IDArticulo"), dr("QInterna"), intObra)
    '                    'If Not IsNothing(filtro) AndAlso filtro.Rows.Count Then
    '                    Dim Representantes As DataTable = CalculoRepresentantes(ArtInfo, ClteInfo, intObra, Nz(dr("QInterna"), 0))

    '                    If Not IsNothing(Representantes) AndAlso Representantes.Rows.Count > 0 Then

    '                        Dim strRepresentante As String
    '                        Dim dv As DataView = Representantes.DefaultView
    '                        dv.RowFilter = "Comision<>0"
    '                        dv.Sort = "IDRepresentante, Orden"

    '                        For Each r As DataRowView In Representantes.DefaultView

    '                            If strRepresentante <> r("IDRepresentante") Then
    '                                Dim newrow As DataRow = newData.NewRow()

    '                                strRepresentante = r("IDRepresentante")

    '                                If MarcaRepresentante Then
    '                                    newrow("MarcaRepresentante") = admin.GetAutoNumeric
    '                                End If

    '                                newrow(PKLinea) = dr(PKLinea)
    '                                newrow("IDRepresentante") = r("IDRepresentante")

    '                                comision = r("Comision")
    '                                If Not r("Porcentaje") Then
    '                                    'Calculo el % que supone el fijo
    '                                    If dr("Importe") <> 0 Then comision = comision * 100 / dr("Importe")
    '                                End If
    '                                newrow("Comision") = comision
    '                                newrow("Porcentaje") = r("Porcentaje")
    '                                newrow("Importe") = (dr("Importe") * comision) / 100

    '                                MantenimientoValoresAyB(newrow, Moneda, MonedaA, MonedaB)

    '                                newData.Rows.Add(newrow.ItemArray)
    '                            End If

    '                        Next
    '                    End If
    '                    'End If
    '                End If
    '            End If

    '            Return newData
    '        End If
    '    End Function

    '#End Region

    '    Public Sub CopiarRepresentantes(ByVal representantesOrigen As DataTable, ByVal representantesDestino As DataTable, ByVal intIDLineaDestino As Integer, ByVal dblImporte As Double)
    '        If Not IsNothing(representantesOrigen) And Not IsNothing(representantesDestino) Then
    '            If representantesOrigen.Rows.Count Then

    '                Dim pkOrigen, pkDestino As String

    '                If representantesOrigen.TableName = "PedidoVentaRepresentante" And representantesDestino.TableName = "AlbaranVentaRepresentante" Then
    '                    pkOrigen = "IDLineaPedido"
    '                    pkDestino = "IDLineaAlbaran"
    '                ElseIf representantesOrigen.TableName = "AlbaranVentaRepresentante" And representantesDestino.TableName = "FacturaVentaRepresentante" Then
    '                    pkOrigen = "IDLineaAlbaran"
    '                    pkDestino = "IDLineaFactura"
    '                ElseIf representantesOrigen.TableName = "PedidoVentaRepresentante" And representantesDestino.TableName = "PedidoVentaAnalitica" Then
    '                    pkOrigen = "IDLineaPedido"
    '                    pkDestino = "IDLineaPedido"
    '                ElseIf representantesOrigen.TableName = "AlbaranVentaRepresentante" And representantesDestino.TableName = "AlbaranVentaRepresentante" Then
    '                    pkOrigen = "IDLineaAlbaran"
    '                    pkDestino = "IDLineaAlbaran"
    '                ElseIf representantesOrigen.TableName = "FacturaVentaRepresentante" And representantesDestino.TableName = "FacturaVentaRepresentante" Then
    '                    pkOrigen = "IDLineaFactura"
    '                    pkDestino = "IDLineaFactura"
    '                End If

    '                For Each origen As DataRow In representantesOrigen.Rows
    '                    Dim destino As DataRow = representantesDestino.NewRow
    '                    For Each dc As DataColumn In representantesOrigen.Columns
    '                        If Not (dc.ColumnName = pkOrigen And pkOrigen <> pkDestino) Then
    '                            If dc.ColumnName <> pkDestino Then
    '                                destino(dc.ColumnName) = origen(dc)
    '                                If dc.ColumnName = "Importe" Then
    '                                    destino(dc.ColumnName) = (destino(dc.ColumnName) * origen("Porcentaje")) / 100
    '                                End If
    '                            End If
    '                        End If
    '                    Next
    '                    destino(pkDestino) = intIDLineaDestino
    '                    If representantesDestino.TableName = "FacturaVentaRepresentante" Then
    '                        destino("MarcaRepresentante") = AdminData.GetAutoNumeric
    '                    End If
    '                    representantesDestino.Rows.Add(destino)
    '                Next
    '            End If
    '        End If
    '    End Sub

    '#End Region

End Class