Imports Solmicro.Expertis.Business.Negocio.NegocioGeneral
<Transactional()> _
Public Class Compra
    Inherits ContextBoundObject

    'Private Const cmstrFacturacion_tbCabecera As String = "tbFacturaCompraCabecera"
    'Private Const cmstrFacturacion_tbLinea As String = "tbFacturaCompraLinea"


    'Public Function TarifaCompra(ByVal StrIDArticulo As String, _
    '                             ByVal StrIDProveedor As String, _
    '                             ByVal DblCantidad As Double, _
    '                             Optional ByVal DtmFecha As Date = cnMinDate, _
    '                             Optional ByVal IntIDLineaContrato As Integer = 0) As DataTable
    '    Dim BlnEncontrado As Boolean = False
    '    Dim DblPrecio, DblDto1, DblDto2, DblDto3 As Double
    '    Dim StrIDMoneda, StrIdUdCompra, StrRefProv, StrDescRefProv As String
    '    Dim IntUdValor As Integer = 1
    '    Dim DtArt, DtFinal As DataTable

    '    If IntIDLineaContrato > 0 Then
    '        Dim ClsContLin As New BusinessHelper("ContratoLinea")
    '        Dim DtContrato As DataTable = ClsContLin.Filter(New FilterItem("IDLineaContrato", FilterOperator.Equal, IntIDLineaContrato, FilterType.Numeric))
    '        If Not DtContrato Is Nothing AndAlso DtContrato.Rows.Count > 0 Then
    '            Dim DtProv As DataTable = New Proveedor().SelOnPrimaryKey(StrIDProveedor)
    '            DblPrecio = DtContrato.Rows(0)("Precio")
    '            DblDto1 = 0 : DblDto2 = 0 : DblDto3 = 0
    '            StrIdUdCompra = DtContrato.Rows(0)("IdUdCompra") & String.Empty
    '            IntUdValor = DtContrato.Rows(0)("UdValoracion")
    '            StrIDMoneda = DtProv.Rows(0)("IDMoneda") & String.Empty
    '            StrIdUdCompra = DtProv.Rows(0)("IdUdCompra") & String.Empty
    '            BlnEncontrado = True
    '        End If
    '    End If

    '    If Not BlnEncontrado Then
    '        Dim DtArtProv As DataTable = New ArticuloProveedor().SelOnPrimaryKey(StrIDProveedor, StrIDArticulo)
    '        If Not DtArtProv Is Nothing AndAlso DtArtProv.Rows.Count > 0 Then
    '            'Si el articulo esta asociado con el proveedor
    '            'Obtenemos el Precio correspondiente a la cantidad
    '            StrRefProv = DtArtProv.Rows(0)("RefProveedor") & String.Empty
    '            StrDescRefProv = DtArtProv.Rows(0)("DescRefProveedor") & String.Empty
    '            Dim FilArtProv As New Filter
    '            FilArtProv.Add("IdProveedor", FilterOperator.Equal, StrIDProveedor, FilterType.String)
    '            FilArtProv.Add("IDArticulo", FilterOperator.Equal, StrIDArticulo, FilterType.String)
    '            FilArtProv.Add("QDesde", FilterOperator.LessThanOrEqual, DblCantidad, FilterType.Numeric)
    '            Dim DtAPLinea As DataTable = New ArticuloProveedorLinea().Filter(FilArtProv, "QDesde DESC")
    '            If Not DtAPLinea Is Nothing AndAlso DtAPLinea.Rows.Count = 0 Then
    '                DtAPLinea = DtArtProv.Copy
    '            End If
    '            If Not DtAPLinea Is Nothing AndAlso DtAPLinea.Rows.Count > 0 Then
    '                Dim DtProv As DataTable = New Proveedor().SelOnPrimaryKey(StrIDProveedor)
    '                DblPrecio = DtAPLinea.Rows(0)("Precio")
    '                DblDto1 = DtAPLinea.Rows(0)("Dto1")
    '                DblDto2 = DtAPLinea.Rows(0)("Dto2")
    '                DblDto3 = DtAPLinea.Rows(0)("Dto3")
    '                IntUdValor = DtArtProv.Rows(0)("UdValoracion")
    '                StrIDMoneda = DtProv.Rows(0)("IDMoneda") & String.Empty
    '                StrIdUdCompra = DtArtProv.Rows(0)("IdUdCompra") & String.Empty
    '                BlnEncontrado = True
    '            End If
    '        End If

    '        ' Puede que encontremos la relación artículo - proveedor pero que unicamente tengamos descuentos sin precio
    '        If DblPrecio = 0 Then
    '            DtArt = New Articulo().SelOnPrimaryKey(StrIDArticulo)
    '            DblPrecio = Nz(DtArt.Rows(0)("PrecioEstandarA"), 0)
    '        End If

    '        If Not BlnEncontrado Then
    '            DescuentosFamilia(StrIDProveedor, StrIDArticulo, DblDto1, DblDto2, DblDto3)
    '            IntUdValor = DtArt.Rows(0)("UdValoracion")
    '            Dim MA As MonedaInfo = New Moneda().MonedaA
    '            StrIDMoneda = MA.ID
    '            StrIdUdCompra = DtArt.Rows(0)("IdUdCompra") & String.Empty
    '            BlnEncontrado = True
    '        End If

    '        DtFinal = New DataTable
    '        With DtFinal
    '            .Columns.Add("Precio", GetType(Double))
    '            .Columns.Add("Dto1", GetType(Double))
    '            .Columns.Add("Dto2", GetType(Double))
    '            .Columns.Add("Dto3", GetType(Double))
    '            .Columns.Add("UdValoracion", GetType(Integer))
    '            .Columns.Add("IDMoneda", GetType(String))
    '            .Columns.Add("IdUdCompra", GetType(String))
    '            .Columns.Add("RefProveedor", GetType(String))
    '            .Columns.Add("DescRefProveedor", GetType(String))
    '        End With
    '    End If
    '    If BlnEncontrado Then
    '        Dim DrNew As DataRow = DtFinal.NewRow
    '        DrNew("Precio") = DblPrecio
    '        DrNew("Dto1") = DblDto1
    '        DrNew("Dto2") = DblDto2
    '        DrNew("Dto3") = DblDto3
    '        DrNew("UdValoracion") = IntUdValor
    '        DrNew("IDMoneda") = StrIDMoneda
    '        DrNew("IdUdCompra") = StrIdUdCompra
    '        DrNew("RefProveedor") = StrRefProv
    '        DrNew("DescRefProveedor") = StrDescRefProv
    '        DtFinal.Rows.Add(DrNew)
    '    End If
    '    Return DtFinal
    'End Function

    'Private Sub DescuentosFamilia(ByVal strIDProveedor As String, ByVal strIDArticulo As String, ByRef dblDto1 As Double, ByRef dblDto2 As Double, ByRef dblDto3 As Double)
    '    Dim DtFamiliaDtosProv As DataTable
    '    Dim strWhere As String

    '    Dim DtArticulo As DataTable = New Articulo().SelOnPrimaryKey(strIDArticulo)

    '    If Not DtArticulo Is Nothing AndAlso DtArticulo.Rows.Count > 0 Then
    '        If Len(DtArticulo.Rows(0)("IDTipo")) > 0 Then
    '            'Primero se buscan los Descuentos por Proveedor-Tipo-Familia
    '            strWhere = "IdProveedor= '" & strIDProveedor & "'"
    '            strWhere = strWhere & " AND IDTipo= '" & DtArticulo.Rows(0)("IDTipo") & "'"
    '            If Len(DtArticulo.Rows(0)("IDFamilia")) > 0 Then
    '                strWhere = strWhere & " AND IDFamilia= '" & DtArticulo.Rows(0)("IDFamilia") & "'"
    '            End If
    '            DtFamiliaDtosProv = New ProveedorDescuentoFamilia().Filter(, strWhere)

    '            'Si no encuentra Descuentos por Proveedor-Tipo-Familia se buscará Cliente-Tipo
    '            If DtFamiliaDtosProv.Rows.Count = 0 And Len(DtArticulo.Rows(0)("IDFamilia")) > 0 Then
    '                strWhere = "IdProveedor= '" & strIDProveedor & "'"
    '                strWhere = strWhere & " AND IDTipo= '" & DtArticulo.Rows(0)("IDTipo") & "' and IDFamilia is null"
    '                DtFamiliaDtosProv = New ProveedorDescuentoFamilia().Filter(, strWhere)
    '            End If

    '            If DtFamiliaDtosProv.Rows.Count > 0 Then
    '                dblDto1 = DtFamiliaDtosProv.Rows(0)("Dto1")
    '                dblDto2 = DtFamiliaDtosProv.Rows(0)("Dto2")
    '                dblDto3 = DtFamiliaDtosProv.Rows(0)("Dto3")
    '            End If
    '        End If
    '    End If

    '    Exit Sub

    'End Sub


    '#Region " DetailCommonUpdateRules "

    '    Public Function DetailCommonUpdateRules(ByVal dttSource As DataTable) As Boolean
    '        If Not IsNothing(dttSource) AndAlso dttSource.Rows.Count Then
    '            Dim services As New ServiceProvider
    '            For Each dr As DataRow In dttSource.Rows
    '                If DetailCommonUpdateRules(dr, services) Then
    '                    DetailCommonUpdateRules = True
    '                Else
    '                    Return False
    '                End If
    '            Next
    '        End If
    '    End Function

    '    Public Function DetailCommonUpdateRules(ByVal dr As DataRow, ByVal services As ServiceProvider) As Boolean
    '        If Length(dr("IdArticulo")) = 0 Then
    '            ApplicationService.GenerateError("El artículo es obligatorio.")
    '        ElseIf Length(dr("DescArticulo")) = 0 Then
    '            ApplicationService.GenerateError("La descripción es obligatoria.")
    '        ElseIf Length(dr("IdTipoIva")) = 0 Then
    '            ApplicationService.GenerateError("El tipo de IVA es obligatorio.")
    '        ElseIf Length(dr("CContable")) = 0 Then
    '            Dim AppParamsConta As ParametroContabilidadCompra = services.GetService(GetType(ParametroContabilidadCompra))
    '            If AppParamsConta.Contabilidad Then ApplicationService.GenerateError("La cuenta contable es obligatoria.")
    '        ElseIf Not IsNumeric(dr("Precio")) Then
    '            ApplicationService.GenerateError("El precio no es válido.")
    '        ElseIf Not IsNumeric(dr("Dto1")) Then
    '            ApplicationService.GenerateError("El Dto. 1 no es válido.")
    '        ElseIf Not IsNumeric(dr("Dto2")) Then
    '            ApplicationService.GenerateError("El Dto. 2 no es válido.")
    '        ElseIf Not IsNumeric(dr("Dto3")) Then
    '            ApplicationService.GenerateError("El Dto. 3 no es válido.")
    '        ElseIf Not IsNumeric(dr("Dto")) Then
    '            ApplicationService.GenerateError("El Dto. Comercial no es válido.")
    '        ElseIf Not IsNumeric(dr("DtoProntoPago")) Then
    '            ApplicationService.GenerateError("El Dto. Pronto Pago no es válido.")
    '        ElseIf dr.Table.Columns.Contains("Cantidad") AndAlso (Not IsNumeric(dr("Cantidad")) OrElse dr("Cantidad") = 0) Then
    '            ApplicationService.GenerateError("La cantidad no es válida.")
    '        ElseIf Not IsNumeric(dr("Factor")) Then
    '            ApplicationService.GenerateError("El factor no es válido.")
    '        ElseIf Not IsNumeric(dr("QInterna")) Then
    '            ApplicationService.GenerateError("La cantidad interna no es válida")
    '        ElseIf Not IsNumeric(dr("UdValoracion")) Then
    '            ApplicationService.GenerateError("La Unidad de Valoración no es válida.")
    '        Else
    '            If dr("UdValoracion") <= 0 Then
    '                ApplicationService.GenerateError("La unidad de valoración debe ser positiva.")
    '            ElseIf dr("Factor") <= 0 Then
    '                ApplicationService.GenerateError("El factor debe ser positivo.")
    '            Else
    '                Return True
    '            End If
    '        End If
    '    End Function

    '#End Region

    'TODO: Provisionar hasta que se pase todo a las B.Rules
    'Public Function DetailBusinessRules(ByVal ColumnName As String, _
    '                                    ByVal Value As Object, _
    '                                    ByVal current As IPropertyAccessor, _
    '                                    ByVal services As ServiceProvider, _
    '                                    Optional ByVal Context As IPropertyAccessor = Nothing) As IPropertyAccessor
    '    'Dim obrl As New BusinessRules
    '    'ProcesoCompra.DetailBusinessRulesLin(obrl)
    'End Function

End Class

#Region " Borrar "


''#Region " Obtener IVA "

''    'Public Function ObtenerIVA(ByVal strIDProveedor As String, ByVal strIDArticulo As String, _
''    '                        Optional ByVal dtProveedor As DataTable = Nothing) As String

''    '    Dim strIVA As String
''    '    If Len(strIDProveedor) > 0 Then
''    '        If IsNothing(dtProveedor) OrElse dtProveedor.Rows.Count = 0 Then
''    '            dtProveedor = New Proveedor().SelOnPrimaryKey(strIDProveedor)
''    '        End If

''    '        If Not dtProveedor Is Nothing AndAlso dtProveedor.Rows.Count > 0 Then
''    '            'Si el parametro IVAProveedor está a 1, entonces la función devolverá siempre el iva del proveedor, sin importar la nacionalidad
''    '            Dim p As New Parametro
''    '            If p.IVAProveedor() Then
''    '                strIVA = dtProveedor.Rows(0)("IDTipoIva") & String.Empty
''    '            Else
''    '                'En caso contrario, se comprueba la nacionalidad
''    '                Dim dtPais As DataTable = New Pais().SelOnPrimaryKey(dtProveedor.Rows(0)("IDPais"))
''    '                If Not dtPais Is Nothing AndAlso dtPais.Rows.Count > 0 Then
''    '                    If dtPais.Rows(0)("Extranjero") Or dtPais.Rows(0)("CanariasCeutaMelilla") Then
''    '                        'Si el Prov es extranjero o "CanariasCeutaMelilla" se devuelve el iva del Proveedor.
''    '                        strIVA = dtProveedor.Rows(0)("IDTipoIva") & String.Empty
''    '                    Else
''    '                        'Si el Prov. es nacional se devuelve el iva del artículo
''    '                        If Len(strIDArticulo) > 0 Then
''    '                            Dim dtArticulo As DataTable = New Articulo().SelOnPrimaryKey(strIDArticulo)
''    '                            If Not dtArticulo Is Nothing AndAlso dtArticulo.Rows.Count > 0 Then
''    '                                strIVA = dtArticulo.Rows(0)("IDTipoIva") & String.Empty
''    '                            End If
''    '                        End If
''    '                        'Si el artículo no tiene iva, se devuelve el del Proveedor.
''    '                        If Len(strIVA) = 0 Then strIVA = dtProveedor.Rows(0)("IDTipoIva") & String.Empty
''    '                    End If
''    '                End If
''    '            End If
''    '        End If
''    '    End If

''    '    Return strIVA
''    'End Function

''    'Public Function ObtenerIVA(ByVal IDProveedor As String, ByVal IDArticulo As String) As String
''    '    Return ObtenerIVA(IDProveedor, IDArticulo, New Parametro().IVAProveedor)
''    'End Function

''    'Public Function ObtenerIVA(ByVal IDProveedor As String, ByVal IDArticulo As String, ByVal blnIVAProveedor As Boolean) As String
''    '    Dim strIVA As String
''    '    If Length(IDProveedor) > 0 Then
''    '        Dim ProvInfo As ProveedorInfo = New Proveedor().InformacionProveedor(IDProveedor)
''    '        Dim context As New BusinessData
''    '        context("IDProveedor") = IDProveedor
''    '        Dim ArtInfo As ArticuloInfo = New Articulo().InformacionArticulo(IDArticulo)
''    '        strIVA = ObtenerIVA(ProvInfo, ArtInfo.TipoIVA, blnIVAProveedor)
''    '    End If

''    '    Return strIVA
''    'End Function

''    'Public Function ObtenerIVA(ByVal ProvInfo As ProveedorInfo, ByVal TipoIVAArticulo As String, ByVal blnIVAProveedor As Boolean) As String
''    '    ObtenerIVA = String.Empty
''    '    If Not IsNothing(ProvInfo) Then
''    '        '//Si el parametro IVAProveedor está a 1, entonces la función devolverá siempre el iva del proveedor, sin importar la nacionalidad
''    '        If blnIVAProveedor Then
''    '            ObtenerIVA = ProvInfo.IDTipoIVA
''    '        Else
''    '            If (ProvInfo.Extranjero) OrElse ProvInfo.CanariasCeutaMelilla Then
''    '                '//Si el cliente es extranjero o "CanariasCeutaMelilla" se devuelve el iva del cliente
''    '                ObtenerIVA = ProvInfo.IDTipoIVA
''    '            Else
''    '                '//Si el cliente es nacional, se devuelve el iva del artículo
''    '                If Length(TipoIVAArticulo) > 0 Then
''    '                    ObtenerIVA = TipoIVAArticulo
''    '                End If

''    '                '//Si el artículo no tiene iva, se devuelve el del cliente.
''    '                If Length(ObtenerIVA) = 0 Then ObtenerIVA = ProvInfo.IDTipoIVA
''    '            End If
''    '        End If
''    '    End If
''    'End Function

''#End Region

'#Region " BusinessRules "

''Private Enum TipoCContable
''    CCImport
''    CCCompra
''    CCImportGRUPO
''    CCCompraGRUPO
''End Enum

'#Region " TratarCContable "

''Private Sub GetCContableArticulo(ByVal current As IPropertyAccessor, ByVal ArtInfo As ArticuloInfo, _
''                                 ByVal tipo As TipoCContable, ByVal services As ServiceProvider)

''    Dim AppParams As ParametroContabilidadCompra = services.GetService(GetType(ParametroContabilidadCompra))
''    If Not AppParams.Contabilidad Then Exit Sub

''    Dim strCContable, strField As String
''    Select Case tipo
''        Case TipoCContable.CCImport
''            strCContable = ArtInfo.CCImport : strField = "CCImport"
''        Case TipoCContable.CCImportGRUPO
''            strCContable = ArtInfo.CCImportGrupo : strField = "CCImportGRUPO"
''        Case TipoCContable.CCCompra
''            strCContable = ArtInfo.CCCompra : strField = "CCCompra"
''        Case TipoCContable.CCCompraGRUPO
''            strCContable = ArtInfo.CCCompraGrupo : strField = "CCCompraGRUPO"
''    End Select
''    If Length(strCContable) > 0 Then
''        current("CContable") = strCContable
''    Else
''        Dim f As New Filter
''        f.Add(New StringFilterItem("IDTipo", ArtInfo.IDTipo))
''        f.Add(New StringFilterItem("IDFamilia", ArtInfo.IDFamilia))
''        Dim dtFam As DataTable = New Familia().Filter(f)
''        If Not dtFam Is Nothing AndAlso dtFam.Rows.Count > 0 Then
''            current("CContable") = dtFam.Rows(0)(strField)
''        End If

''        If Length(current("CContable")) = 0 Then
''            Select Case tipo
''                Case TipoCContable.CCImport
''                    current("CContable") = AppParams.CuentaImportacion
''                Case TipoCContable.CCImportGRUPO
''                    current("CContable") = AppParams.CuentaImportacionGrupo
''                Case TipoCContable.CCCompra
''                    current("CContable") = AppParams.CuentaCompra
''                Case TipoCContable.CCCompraGRUPO
''                    current("CContable") = AppParams.CuentaCompraGrupo
''            End Select
''        End If
''    End If
''End Sub

''Public Function TratarCContable(ByVal current As IPropertyAccessor, ByVal Context As IPropertyAccessor, _
''                             ByVal ArtInfo As ArticuloInfo) As String

''    Dim services As New ServiceProvider
''    TratarCContable(current, Context, ArtInfo, services)

''    Return current("CContable") & String.Empty
''End Function

''Friend Sub TratarCContable(ByVal current As IPropertyAccessor, ByVal Context As IPropertyAccessor, _
''                         ByVal ArtInfo As ArticuloInfo, ByVal services As ServiceProvider)

''    Dim AppParams As ParametroContabilidadCompra = services.GetService(GetType(ParametroContabilidadCompra))
''    If Not AppParams.Contabilidad Then Exit Sub

''    Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(GetType(EntityInfoCache(Of ProveedorInfo)))
''    Dim ProvInfo As ProveedorInfo
''    If Context.ContainsKey("IDProveedor") AndAlso Length(Context("IDProveedor")) > 0 Then
''        ProvInfo = Proveedores.GetEntity(Context("IDProveedor"))
''    End If

''    If ProvInfo.EmpresaGrupo Then
''        If Not ProvInfo.Extranjero Then
''            GetCContableArticulo(current, ArtInfo, TipoCContable.CCCompraGRUPO, services)
''        Else
''            GetCContableArticulo(current, ArtInfo, TipoCContable.CCImportGRUPO, services)
''        End If
''    ElseIf Not ProvInfo.Extranjero Then
''        GetCContableArticulo(current, ArtInfo, TipoCContable.CCCompra, services)
''    Else
''        GetCContableArticulo(current, ArtInfo, TipoCContable.CCImport, services)
''    End If
''    FormatoCuentaContable(Context, current, services)
''End Sub

'#End Region

''Public Function DetailBusinessRules(ByVal ColumnName As String, ByVal Value As Object, ByVal current As IPropertyAccessor, Optional ByVal Context As IPropertyAccessor = Nothing) As Solmicro.Expertis.IPropertyAccessor
''    If Context Is Nothing Then Context = New BusinessData

''    Dim dtProveedor As DataTable
''    Dim blnEmpresaGrupo As Boolean
''    If Context.ContainsKey("IDProveedor") AndAlso Length(Context("IDProveedor")) > 0 Then
''        dtProveedor = New Proveedor().SelOnPrimaryKey(Context("IDProveedor"))
''        If Not dtProveedor Is Nothing AndAlso dtProveedor.Rows.Count > 0 Then
''            blnEmpresaGrupo = CBool(dtProveedor.Rows(0)("EmpresaGrupo"))
''        End If
''    End If

''    Dim blnSubcontratacion As Boolean = False
''    If Context.ContainsKey("IDOrdenRuta") AndAlso Length(Context("IDOrdenRuta")) > 0 Then
''        blnSubcontratacion = True
''    End If

''    Select Case ColumnName
''        Case "IDArticulo"
''            If Length(Value) Then
''                Dim dt As DataTable
''                Dim a As New Articulo
''                dt = a.SelOnPrimaryKey(Value)
''                If IsNothing(dt) OrElse dt.Rows.Count = 0 Then
''                    ApplicationService.GenerateError("El artículo | no existe.", Value)
''                Else
''                    current("IDArticulo") = Value
''                    If dt.Rows(0)("NSerieObligatorio") AndAlso Nz(current("Cantidad"), 0) = 0 Then
''                        current("Cantidad") = 1
''                        current("QInterna") = 1
''                    End If
''                    Dim dtCaract As DataTable = AdminData.GetData("vNegCaractArticulo", New StringFilterItem("IDArticulo", FilterOperator.Equal, Value), "Compra,Subcontratacion, Activo")
''                    If Not IsNothing(dtCaract) AndAlso dtCaract.Rows.Count > 0 Then
''                        If Not (dtCaract.Rows(0)("Compra") Or dtCaract.Rows(0)("subcontratacion")) And Not blnSubcontratacion Then
''                            ApplicationService.GenerateError("El artículo | no es de tipo compra.", Value)
''                        ElseIf Not dtCaract.Rows(0)("Activo") Then
''                            ApplicationService.GenerateError("El artículo | no está activo.", Value)
''                        Else
''                            '//ARTICULO
''                            If dtCaract.Rows(0)("Subcontratacion") Then
''                                current("TipoLineaCompra") = enumaclTipoLineaAlbaran.aclSubcontratacion
''                            Else
''                                current("TipoLineaCompra") = enumaclTipoLineaAlbaran.aclNormal
''                            End If
''                            current("DescArticulo") = dt.Rows(0)("DescArticulo")
''                            current("IDUDMedida") = dt.Rows(0)("IDUDCompra")
''                            current("IDUDInterna") = dt.Rows(0)("IDUDInterna")
''                            current("UdValoracion") = dt.Rows(0)("UdValoracion")
''                            current("IDTipoIva") = dt.Rows(0)("IDTipoIva")
''                            current("IDUDCompra") = dt.Rows(0)("IDUDCompra")

''                            '//PROVEEDOR
''                            If Not dtProveedor Is Nothing AndAlso dtProveedor.Rows.Count Then
''                                '//C.CONTABLES
''                                TratarCContable(current, Context, dt, blnEmpresaGrupo)
''                                '//TIPO IVA
''                                Dim strIDTipoIva As String = ObtenerIVA(Context("IDProveedor"), Value, dtProveedor)
''                                If Length(current("IDTipoIva")) = 0 Then
''                                    ApplicationService.GenerateError("El código de IVA es obligatorio. Revise la relación Articulo-Proveedor.")
''                                Else
''                                    current("IDTipoIva") = strIDTipoIva
''                                End If

''                                '//ARTICULO-PROVEEDOR
''                                Dim dtArticuloProveedor As DataTable
''                                dtArticuloProveedor = New ArticuloProveedor().SelOnPrimaryKey(Context("IDProveedor"), Value)
''                                If Not IsNothing(dtArticuloProveedor) AndAlso dtArticuloProveedor.Rows.Count Then
''                                    If Length(dtArticuloProveedor.Rows(0)("RefProveedor")) > 0 Then
''                                        current("RefProveedor") = dtArticuloProveedor.Rows(0)("RefProveedor")
''                                    End If
''                                    current("DescRefProveedor") = dtArticuloProveedor.Rows(0)("DescRefProveedor")
''                                    current("IDUDMedida") = dtArticuloProveedor.Rows(0)("IDUDCompra")
''                                    current("IDUDCompra") = dtArticuloProveedor.Rows(0)("IDUDCompra")
''                                    current("UdValoracion") = dtArticuloProveedor.Rows(0)("UdValoracion")
''                                End If
''                            End If

''                            '//ALMACEN
''                            Dim strAlmacen As String
''                            Dim strCentroGestion As String
''                            If Context.ContainsKey("IDCentroGestion") AndAlso Length(Context("IDCentroGestion")) > 0 Then
''                                strCentroGestion = Context("IDCentroGestion")
''                            Else
''                                If current.ContainsKey("IDCentroGestion") AndAlso Length(current("IDCentroGestion")) > 0 Then
''                                    strCentroGestion = current("IDCentroGestion")
''                                End If
''                            End If

''                            If Length(current("IDAlmacen")) > 0 AndAlso New Parametro().AlmacenCentroGestionActivo Then
''                                current("IDAlmacen") = current("IDAlmacen")
''                            Else
''                                strAlmacen = New ArticuloAlmacen().AlmacenPredeterminadoArticulo(Value, strCentroGestion)
''                                If Len(strAlmacen) > 0 Then
''                                    current("IDAlmacen") = strAlmacen
''                                End If
''                            End If
''                            current("Factor") = a.FactorConversion(Value, current("IDUDMedida"), current("IDUDInterna"))
''                            If current("Factor") = 0 Then current("Factor") = 1
''                            If current("UdValoracion") = 0 Then current("UdValoracion") = 1

''                            '//TARIFA
''                            If Not blnEmpresaGrupo Then current = TarifaCompra(Context, current)

''                            If Context.ContainsKey("IDMoneda") Then
''                                If Context("IDMoneda") & String.Empty <> current("IDMoneda") & String.Empty Then
''                                    '//CAMBIO DE MONEDA
''                                    current = CambioMoneda(current, current("IDMoneda"), Context("IDMoneda"), Nz(Context("Fecha"), cnMinDate))
''                                    CalcularImportes(current, Context("IDMoneda"))
''                                Else
''                                    CalcularImportes(current, current("IDMoneda"), Context("CambioA"), Context("CambioB"))
''                                End If
''                                current.Remove("IDMoneda")
''                            End If
''                        End If
''                    End If
''                End If
''            Else
''                'Inicializar todo lo que depende de Articulo
''                current("TipoLineaCompra") = enumaclTipoLineaAlbaran.aclNormal
''                current("DescArticulo") = Nothing
''                current("IDUDMedida") = Nothing
''                current("IDUDInterna") = Nothing
''                current("UdValoracion") = 1
''                current("IDTipoIva") = Nothing
''                current("CContable") = Nothing
''                current("IDTipoIva") = Nothing
''                current("Nacional") = False
''                current("CanariasCeutaMelilla") = False
''                current("IDAlmacen") = Nothing
''                current("Lote") = Nothing
''                current("Ubicacion") = Nothing
''                current("RefProveedor") = Nothing
''                current("DescRefProveedor") = Nothing
''                current("Factor") = 1
''                current("Cantidad") = 0
''                current("QInterna") = 0
''                current("Precio") = 0
''                current("Dto1") = 0
''                current("Dto2") = 0
''                current("Dto3") = 0
''            End If

''            If Context.ContainsKey("TipoCompraCabecera") AndAlso Length(Context("TipoCompraCabecera")) > 0 Then
''                If current("TipoLineaCompra") <> Context("TipoCompraCabecera") Then
''                    current("TipoLineaCompra") = Context("TipoCompraCabecera")
''                End If
''            End If
''        Case "RefProveedor"
''            If Length(Value) Then
''                If Context.ContainsKey("IDProveedor") Then
''                    Dim ref As String
''                    ref = New ArticuloProveedor().ObtenerArticuloRef(Context("IDProveedor"), Value)
''                    If Len(ref) = 0 Then
''                        ApplicationService.GenerateError("La referencia | no está asociada a ningún artículo", Value)
''                    Else
''                        current = Me.DetailBusinessRules("IDArticulo", ref, current, Context)
''                        current("IDArticulo") = ref
''                    End If
''                End If
''            ElseIf Length(current("IDArticulo")) > 0 Then
''                current = Me.DetailBusinessRules("IDArticulo", current("IDArticulo"), current, Context)
''            End If

''        Case "IDContrato"
''            If Length(Value) Then
''                Dim dtContrato As DataTable
''                dtContrato = ValidarContratos(Value, current("IDArticulo"), Context("IDProveedor"), current("cantidad"), Context("fecha"))
''                If Not IsNothing(dtContrato) AndAlso dtContrato.Rows.Count Then
''                    current("Precio") = dtContrato.Rows(0)("Precio")
''                    current("IdContrato") = dtContrato.Rows(0)("IdContrato")
''                    current("IdLineaContrato") = dtContrato.Rows(0)("IdLineaContrato")
''                    current("UdValoracion") = dtContrato.Rows(0)("UdValoracion")
''                    current("IDUdMedida") = dtContrato.Rows(0)("IDUdCompra")
''                    '//Se ha modificado el precio, hay que recalcular el importe.
''                    CalcularImportes(current)
''                Else
''                    ApplicationService.GenerateError("EL contrato | no es válido", Value)
''                End If
''            End If

''        Case "Cantidad", "QTiempo"
''            current(ColumnName) = Value
''            If IsNumeric(current(ColumnName)) Then
''                current("QInterna") = Nz(current("Factor"), 1) * current("Cantidad")

''                '//TARIFA
''                If Not blnEmpresaGrupo Then
''                    current = TarifaCompra(Context, current)
''                End If

''                If Context.ContainsKey("IDMoneda") And current.ContainsKey("IDMoneda") Then
''                    If AreDifferents(Context("IDMoneda"), current("IDMoneda")) Then
''                        '//CAMBIO DE MONEDA
''                        current = CambioMoneda(current, current("IDMoneda"), Context("IDMoneda"), Nz(Context("Fecha"), cnMinDate))
''                    End If
''                    current.Remove("IDMoneda")
''                End If
''                CalcularImportes(current, current("IDMoneda"))
''            Else
''                ApplicationService.GenerateError("El valor no es válido.")
''            End If
''        Case "QInterna"
''            If IsNumeric(Value) Then
''                If Nz(current("Cantidad"), 0) <> 0 Then
''                    current("Factor") = Value / current("Cantidad")
''                End If
''            Else
''                ApplicationService.GenerateError("El valor no es válido.")
''            End If
''        Case "Factor"
''            If IsNumeric(Value) Then
''                current("Factor") = Nz(Value, 1)
''                If current("Factor") <= 0 Then
''                    ApplicationService.GenerateError("El factor no es válido.")
''                Else
''                    current("QInterna") = current("Factor") * Nz(current("Cantidad"), 0)
''                End If
''            Else
''                ApplicationService.GenerateError("El valor no es válido.")
''            End If
''        Case "IDUDMedida"
''            If Len(Value & String.Empty) Then
''                current("Factor") = New Articulo().FactorConversion(current("IDArticulo") & String.Empty, current("IDUDMedida") & String.Empty, current("IDUdInterna") & String.Empty)
''                current("QInterna") = current("Factor") * Nz(current("Cantidad"), 0)
''            End If
''        Case "Precio", "UDValoracion", "Dto1", "Dto2", "Dto3"
''            If Length(Value) > 0 Then
''                If IsNumeric(Value) Then
''                    current(ColumnName) = Value
''                    If Context.ContainsKey("IDMoneda") Then
''                        CalcularImportes(current, Context("IDMoneda"))
''                    End If
''                Else
''                    ApplicationService.GenerateError("El valor no es válido.")
''                End If
''            End If
''        Case "PrecioA", "PrecioB"
''            CalcularImportes(current, Context("IDMoneda"), Context("CambioA"), Context("CambioB"))

''        Case "CContable"
''            If Length(Value) > 0 Then
''                current("CContable") = Value
''                ComprobarCContable(Context, current)
''                If current.ContainsKey("Inmovilizado") AndAlso Context.ContainsKey("IDEjercicio") Then
''                    CCInmovilizado(Context("IDEjercicio") & String.Empty, Nz(current("Inmovilizado"), False), current("CContable"))
''                End If
''            End If
''        Case "NObra", "IDObra"
''            If ColumnName = "NObra" AndAlso Length(Value) = 0 Then
''                current("IDObra") = System.DBNull.Value
''            End If
''            current("IDTrabajo") = DBNull.Value
''            current("CodTrabajo") = DBNull.Value
''            If Length(current("IDObra")) > 0 Then
''                Dim obra As BusinessHelper
''                obra = BusinessHelper.CreateBusinessObject("ObraCabecera")
''                Dim dtObra As DataTable = obra.SelOnPrimaryKey(current("IDObra"))
''                If IsNothing(dtObra) OrElse dtObra.Rows.Count = 0 Then
''                    ApplicationService.GenerateError("La Obra | no existe.", Value)
''                Else
''                    Dim dv As New DataView(dtObra)
''                    dv.RowFilter = "Estado<>" & enumocEstado.ocTerminado
''                    If dv.Count > 0 Then
''                        If current.ContainsKey("TipoGastoObra") Then
''                            If Length(current("TipoGastoObra")) = 0 Then current("TipoGastoObra") = CInt(enumfclTipoGastoObra.enumfclMaterial)
''                            If current("TipoGastoObra") = enumfclTipoGastoObra.enumfclMaterial Then
''                                current("IDConcepto") = current("IDArticulo")
''                            End If
''                        End If
''                    Else
''                        ApplicationService.GenerateError("La Obra | está terminada.", Value)
''                    End If
''                End If
''            Else
''                current("IdTrabajo") = System.DBNull.Value
''                current("CodTrabajo") = System.DBNull.Value
''            End If
''        Case "CodTrabajo"
''            current("TipoGastoObra") = enumfclTipoGastoObra.enumfclMaterial
''            If Length(Value) = 0 Then
''                current("IDTrabajo") = DBNull.Value
''            End If
''            If Length(current("IdTrabajo")) > 0 AndAlso Length(current("IDObra")) > 0 Then
''                Dim ClsObra As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraCabecera")
''                Dim dtObra As DataTable = ClsObra.SelOnPrimaryKey(current("IDObra"))
''                If Not dtObra Is Nothing AndAlso dtObra.Rows.Count > 0 Then
''                    current("NObra") = dtObra.Rows(0)("NObra")
''                End If
''            End If

''        Case "IDLineaAlbaran"
''            If IsNumeric(Value) Then
''                Dim dtACL As DataTable
''                dtACL = New AlbaranCompraLinea().SelOnPrimaryKey(Value)
''                If Not IsNothing(dtACL) AndAlso dtACL.Rows.Count Then
''                    current("IdLineaAlbaran") = DBNull.Value
''                    current("IdAlbaran") = dtACL.Rows(0)("IdAlbaran")
''                    current("IDArticulo") = dtACL.Rows(0)("IDArticulo")
''                    current("DescArticulo") = dtACL.Rows(0)("DescArticulo")
''                    current("RefPorveedor") = dtACL.Rows(0)("RefProveedor")
''                    current("Precio") = dtACL.Rows(0)("Precio")
''                    current("UdValoracion") = dtACL.Rows(0)("UdValoracion")
''                    current("Lote") = dtACL.Rows(0)("Lote")
''                    current("IDUdMedida") = dtACL.Rows(0)("IDUdMedida")
''                    current("IDUdInterna") = dtACL.Rows(0)("IDUdInterna")
''                    current("IDTipoIVA") = dtACL.Rows(0)("IDTipoIVA")
''                    If Len(dtACL.Rows(0)("IDCentroGestion") & String.Empty) > 0 Then
''                        current("IDCentroGestion") = dtACL.Rows(0)("IDCentroGestion")
''                    End If
''                    current("IDObra") = dtACL.Rows(0)("IdObra")
''                    current("IDTrabajo") = dtACL.Rows(0)("IdTrabajo")
''                    current("IdLineaMaterial") = dtACL.Rows(0)("IdLineaMaterial")
''                    current("Dto1") = dtACL.Rows(0)("Dto1")
''                    current("Dto2") = dtACL.Rows(0)("Dto2")
''                    current("Dto3") = dtACL.Rows(0)("Dto3")
''                    current("CContable") = dtACL.Rows(0)("CContable")
''                End If
''            End If
''        Case "Inmovilizado"
''            If Length(Value) > 0 Then
''                If current.ContainsKey("CContable") AndAlso Context.ContainsKey("IDEjercicio") Then
''                    CCInmovilizado(Context("IDEjercicio") & String.Empty, Value, current("CContable") & String.Empty)
''                End If
''            End If
''        Case "FechaEntrega"
''            If current.ContainsKey("FechaEntregaModificadoPedido") Then
''                current("FechaEntregaModificadoPedido") = Value
''            End If
''    End Select

''    Return current
''End Function
'Public Function DetailBusinessRules(ByVal ColumnName As String, _
'                                    ByVal Value As Object, _
'                                    ByVal current As IPropertyAccessor, _
'                                    ByVal services As ServiceProvider, _
'                                    Optional ByVal Context As IPropertyAccessor = Nothing) As IPropertyAccessor

'    'If Context Is Nothing Then Context = New BusinessData


'    'Select Case ColumnName
'    '    Case "IDArticulo", "RefProveedor"
'    '        CambioArticulo(ColumnName, Value, current, Context, services)
'    '    Case "IDContrato"
'    '        CambioContrato(current, Value, Context, services)
'    '    Case "Cantidad", "QTiempo"
'    '        CambioCantidad(current, Value, Context, services)
'    '    Case "QInterna"
'    '        CambioQInterna(current, Value, Context, services)
'    '    Case "Factor"
'    '        CambioFactor(current, Value, Context, services)
'    '    Case "IDUDMedida"
'    '        CambioUDMedida(current, Value, Context, services)
'    '    Case "Precio", "UDValoracion", "Dto1", "Dto2", "Dto3", "Dto", "DtoProntoPago"
'    '        CambioPrecio(ColumnName, Value, current, Context, services)
'    '    Case "PrecioA", "PrecioB"
'    '        CalcularImportes(current, Context("IDMoneda"), Context("CambioA"), Context("CambioB"))
'    '    Case "CContable"
'    '        General.CambioCContable(current, Value, Context)
'    '        If current.ContainsKey("Inmovilizado") AndAlso Context.ContainsKey("IDEjercicio") Then
'    '            CCInmovilizado(Context("IDEjercicio") & String.Empty, Nz(current("Inmovilizado"), False), current("CContable") & String.Empty)
'    '        End If
'    '    Case "NObra", "IDObra"
'    '        CambioObra(ColumnName, current, Value)
'    '    Case "CodTrabajo"
'    '        CambioTrabajo(current, Value)
'    '    Case "IDLineaAlbaran"
'    '        CambioLineaAlbaran(current, Value)
'    '    Case "Inmovilizado"
'    '        If current.ContainsKey("Inmovilizado") AndAlso Context.ContainsKey("IDEjercicio") Then
'    '            CCInmovilizado(Context("IDEjercicio") & String.Empty, Nz(current("Inmovilizado"), False), current("CContable") & String.Empty)
'    '        End If
'    '    Case "FechaEntrega"
'    '        If current.ContainsKey("FechaEntregaModificadoPedido") Then current("FechaEntregaModificadoPedido") = Value
'    'End Select

'    Return current
'End Function


''Private Sub CambioArticulo(ByVal ColumnName As String, _
''                           ByVal value As Object, _
''                           ByVal current As IPropertyAccessor, _
''                           ByVal context As IPropertyAccessor, _
''                           ByVal services As ServiceProvider)
''    If Length(value) Then
''        Dim ArtInfo As ArticuloInfo
''        Select Case ColumnName
''            Case "IDArticulo"
''                ArtInfo = New Articulo().InformacionArticulo(value)
''            Case "RefProveedor"
''                ArtInfo = New Articulo().InformacionArticulo(Nothing, Nothing, value, context)
''        End Select

''        CambioArticulo(ArtInfo, current, context, services)
''    End If

''    current(ColumnName) = value
''End Sub

''Friend Sub CambioArticulo(ByVal ArtInfo As ArticuloInfo, _
''                          ByVal current As IPropertyAccessor, _
''                          ByVal context As IPropertyAccessor, _
''                          ByVal services As ServiceProvider)

''    Dim AppParams As ParametroCompra = services.GetService(GetType(ParametroCompra))

''    If Not IsNothing(ArtInfo) Then
''        Dim blnSubcontratacion As Boolean
''        If context.ContainsKey("IDOrdenRuta") AndAlso Length(context("IDOrdenRuta")) > 0 Then
''            blnSubcontratacion = True
''        End If

''        If ArtInfo.NSerieObligatorio AndAlso Nz(current("Cantidad"), 0) = 0 Then
''            current("Cantidad") = 1
''            current("QInterna") = 1
''        End If

''        If Not (ArtInfo.Compra Or ArtInfo.Subcontratacion) And Not blnSubcontratacion Then
''            ApplicationService.GenerateError("El artículo | no es de tipo compra.", Quoted(ArtInfo.IDArticulo))
''        ElseIf Not ArtInfo.Activo Then
''            ApplicationService.GenerateError("El artículo | no está activo.", Quoted(ArtInfo.IDArticulo))
''        Else
''            '//ARTICULO
''            If ArtInfo.Subcontratacion Then
''                current("TipoLineaCompra") = enumaclTipoLineaAlbaran.aclSubcontratacion
''            Else
''                current("TipoLineaCompra") = enumaclTipoLineaAlbaran.aclNormal
''            End If
''            current("DescArticulo") = ArtInfo.DescArticulo
''            current("IDUDMedida") = ArtInfo.IDUDCompra
''            current("IDUDInterna") = ArtInfo.IDUDInterna
''            current("UdValoracion") = ArtInfo.UDValoracion
''            current("IDTipoIva") = ArtInfo.TipoIVA
''            current("IDUDCompra") = ArtInfo.IDUDCompra
''            If ArtInfo.Especial Then
''                If current.ContainsKey("Especial") Then
''                    current("Especial") = ArtInfo.Especial
''                End If
''                current("Dto") = 0
''                current("DtoProntoPago") = 0
''            End If


''            '//ALMACEN                       
''            General.AsignarArticuloAlmacen(current, context, services)

''            General.FactorConversion(current, services)
''            If current("UDValoracion") = 0 Then current("UDValoracion") = 1

''            '//PROVEEDOR
''            Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(GetType(EntityInfoCache(Of ProveedorInfo)))
''            Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(context("IDProveedor"))
''            If Not IsNothing(ProvInfo) Then
''                '//C.CONTABLES
''                TratarCContable(current, context, ArtInfo, services)

''                '//TIPO IVA
''                Dim strIDTipoIva As String = ObtenerIVA(ProvInfo, ArtInfo.TipoIVA, AppParams.IVAProveedor)
''                If Length(strIDTipoIva) = 0 Then
''                    ApplicationService.GenerateError("El código de IVA es obligatorio. Revise la relación Articulo-Proveedor.")
''                Else
''                    current("IDTipoIva") = strIDTipoIva
''                End If

''                '//ARTICULO-PROVEEDOR
''                ArticuloProveedor(current, ProvInfo)

''                '//TARIFA
''                Tarifa(current, context, services)
''            End If
''        End If

''        current("IDArticulo") = ArtInfo.IDArticulo
''        If context.ContainsKey("TipoCompraCabecera") AndAlso Length(context("TipoCompraCabecera")) > 0 Then
''            If current("TipoLineaCompra") <> context("TipoCompraCabecera") Then
''                current("TipoLineaCompra") = context("TipoCompraCabecera")
''            End If
''        End If
''    End If
''End Sub

''Friend Sub CambioArticulo(ByVal ArtInfo As ArticuloInfo, ByVal current As DataRow, ByVal context As IPropertyAccessor, ByVal services As ServiceProvider)
''    CambioArticulo(ArtInfo, New DataRowPropertyAccessor(current), context, services)
''End Sub

''Friend Sub ArticuloProveedor(ByVal current As IPropertyAccessor, ByVal ProvInfo As ProveedorInfo)
''    Dim ArtProvInfo As ArticuloProveedorInfo = New ArticuloProveedor().InformacionArticuloProveedor(current("IDArticulo"), ProvInfo.IDProveedor)
''    If Not IsNothing(ArtProvInfo) Then
''        current("RefProveedor") = ArtProvInfo.RefProveedor
''        If Length(ArtProvInfo.DescRefProveedor) > 0 Then
''            current("DescRefProveedor") = ArtProvInfo.DescRefProveedor
''        End If
''        If Length(ArtProvInfo.IDUDCompra) > 0 Then
''            current("IDUDMedida") = ArtProvInfo.IDUDCompra
''            current("IDUDCompra") = ArtProvInfo.IDUDCompra
''        End If
''        current("UdValoracion") = ArtProvInfo.UdValoracion
''    End If
''End Sub

''Private Sub Tarifa(ByVal current As IPropertyAccessor, ByVal context As IPropertyAccessor, ByVal services As ServiceProvider)

''    Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(GetType(EntityInfoCache(Of ProveedorInfo)))
''    Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(context("IDProveedor"))
''    If Not ProvInfo.EmpresaGrupo Then
''        TarifaCompra(context, current, services)

''        If context.ContainsKey("IDMoneda") And current.ContainsKey("IDMoneda") Then
''            '//Tendremos un current("IDMoneda") cuando cambiamos la tarifa
''            If Length(current("IDMoneda")) > 0 Then
''                If context("IDMoneda") & String.Empty <> current("IDMoneda") & String.Empty Then
''                    '//CAMBIO DE MONEDA (aplicamos el cambio de la tarifa)
''                    current = CambioMoneda(current, current("IDMoneda"), context("IDMoneda"), Nz(context("Fecha"), cnMinDate), services.GetService(GetType(MonedaCache)))
''                End If
''            End If

''            '//IMPORTE (aplicamos el cambio de la cabecera)
''            If context.ContainsKey("CambioA") AndAlso context.ContainsKey("CambioB") Then
''                CalcularImportes(current, context("IDMoneda"), context("CambioA"), context("CambioB"), services)
''            Else
''                CalcularImportes(current, context("IDMoneda"), , services)
''            End If
''        End If
''    End If
''End Sub

''Private Sub CambioContrato(ByVal current As IPropertyAccessor, ByVal value As Object, ByVal context As IPropertyAccessor, ByVal services As ServiceProvider)
''    If Length(value) Then
''        Dim dtContrato As DataTable = ValidarContratos(value, current("IDArticulo"), context("IDProveedor"), current("cantidad"), context("fecha"))
''        If Not IsNothing(dtContrato) AndAlso dtContrato.Rows.Count Then
''            current("Precio") = dtContrato.Rows(0)("Precio")
''            current("IdContrato") = dtContrato.Rows(0)("IdContrato")
''            current("IdLineaContrato") = dtContrato.Rows(0)("IdLineaContrato")
''            current("UdValoracion") = dtContrato.Rows(0)("UdValoracion")
''            current("IDUdMedida") = dtContrato.Rows(0)("IDUdCompra")
''            '//Se ha modificado el precio, hay que recalcular el importe.
''            CalcularImportes(current, services)
''        Else
''            ApplicationService.GenerateError("EL contrato | no es válido", Quoted(value))
''        End If
''    End If
''End Sub

''Private Sub CambioCantidad(ByVal current As IPropertyAccessor, ByVal Value As Object, ByVal context As IPropertyAccessor, ByVal services As ServiceProvider)
''    If IsNumeric(Value) Then
''        current("Cantidad") = Value
''        CalculoQInterna(current, services)
''        '///TARIFA
''        Tarifa(current, context, services)
''    Else
''        ApplicationService.GenerateError("Campo no numérico.")
''    End If
''End Sub

''Public Sub CambioCantidad(ByVal current As DataRow, ByVal Value As Object, ByVal context As IPropertyAccessor, ByVal blnEmpresaGrupo As Boolean, ByVal services As ServiceProvider)
''    CambioCantidad(New DataRowPropertyAccessor(current), Value, context, services)
''End Sub

''Private Sub CambioLineaAlbaran(ByVal current As IPropertyAccessor, ByVal Value As Object)
''    If IsNumeric(Value) Then
''        Dim dtACL As DataTable = New AlbaranCompraLinea().SelOnPrimaryKey(Value)
''        If Not IsNothing(dtACL) AndAlso dtACL.Rows.Count Then
''            current("IdLineaAlbaran") = DBNull.Value
''            current("IdAlbaran") = dtACL.Rows(0)("IdAlbaran")
''            current("IDArticulo") = dtACL.Rows(0)("IDArticulo")
''            current("DescArticulo") = dtACL.Rows(0)("DescArticulo")
''            current("RefPorveedor") = dtACL.Rows(0)("RefProveedor")
''            current("Precio") = dtACL.Rows(0)("Precio")
''            current("UdValoracion") = dtACL.Rows(0)("UdValoracion")
''            current("Lote") = dtACL.Rows(0)("Lote")
''            current("IDUdMedida") = dtACL.Rows(0)("IDUdMedida")
''            current("IDUdInterna") = dtACL.Rows(0)("IDUdInterna")
''            current("IDTipoIVA") = dtACL.Rows(0)("IDTipoIVA")
''            If Length(dtACL.Rows(0)("IDCentroGestion")) > 0 Then
''                current("IDCentroGestion") = dtACL.Rows(0)("IDCentroGestion")
''            End If
''            current("IDObra") = dtACL.Rows(0)("IdObra")
''            current("IDTrabajo") = dtACL.Rows(0)("IdTrabajo")
''            current("IdLineaMaterial") = dtACL.Rows(0)("IdLineaMaterial")
''            current("Dto1") = dtACL.Rows(0)("Dto1")
''            current("Dto2") = dtACL.Rows(0)("Dto2")
''            current("Dto3") = dtACL.Rows(0)("Dto3")
''            current("CContable") = dtACL.Rows(0)("CContable")
''        End If
''    End If
''End Sub

''Private Sub CambioObra(ByVal ColumnName As String, ByVal current As IPropertyAccessor, ByVal Value As Object)
''    If ColumnName = "NObra" AndAlso Length(Value) = 0 Then
''        current("IDObra") = System.DBNull.Value
''    End If
''    current("IDTrabajo") = DBNull.Value
''    current("CodTrabajo") = DBNull.Value
''    If Length(current("IDObra")) > 0 Then
''        Dim obra As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraCabecera")
''        Dim dtObra As DataTable = obra.SelOnPrimaryKey(current("IDObra"))
''        If IsNothing(dtObra) OrElse dtObra.Rows.Count = 0 Then
''            ApplicationService.GenerateError("La Obra | no existe.", Quoted(Value))
''        Else
''            Dim objFilter As New Filter
''            objFilter.Add(New NumberFilterItem("Estado", FilterOperator.NotEqual, enumocEstado.ocTerminado))
''            Dim dv As New DataView(dtObra)
''            dv.RowFilter = objFilter.Compose(New AdoFilterComposer)
''            If dv.Count > 0 Then
''                If current.ContainsKey("TipoGastoObra") Then
''                    If Length(current("TipoGastoObra")) = 0 Then current("TipoGastoObra") = CInt(enumfclTipoGastoObra.enumfclMaterial)
''                    If current("TipoGastoObra") = enumfclTipoGastoObra.enumfclMaterial Then
''                        current("IDConcepto") = current("IDArticulo")
''                    End If
''                End If
''            Else
''                ApplicationService.GenerateError("La Obra | está terminada.", Quoted(Value))
''            End If
''        End If
''    End If
''End Sub

''Private Sub CambioTrabajo(ByVal current As IPropertyAccessor, ByVal Value As Object)
''    current("TipoGastoObra") = enumfclTipoGastoObra.enumfclMaterial
''    If Length(Value) = 0 Then
''        current("IDTrabajo") = DBNull.Value
''    End If
''    If Length(current("IdTrabajo")) > 0 AndAlso Length(current("IDObra")) > 0 Then
''        Dim ClsObra As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraCabecera")
''        Dim dtObra As DataTable = ClsObra.SelOnPrimaryKey(current("IDObra"))
''        If Not dtObra Is Nothing AndAlso dtObra.Rows.Count > 0 Then
''            current("NObra") = dtObra.Rows(0)("NObra")
''        End If
''    End If
''End Sub

'''QUITAR cuando tengamos las nuevas BRules
''Friend Sub CambioQInterna(ByVal current As IPropertyAccessor, ByVal value As Object, ByVal context As IPropertyAccessor, ByVal services As ServiceProvider)
''    Dim brd As New BusinessRuleData("QInterna", value, current, context)
''    General.CambioQInterna(brd, services)
''End Sub

'''QUITAR cuando tengamos las nuevas BRules
''Friend Sub CambioFactor(ByVal current As IPropertyAccessor, ByVal Value As Object, ByVal Context As IPropertyAccessor, ByVal services As ServiceProvider)
''    Dim brd As New BusinessRuleData("QInterna", Value, current, Context)
''    General.CambioFactor(brd, services)
''End Sub

'''QUITAR cuando tengamos las nuevas BRules
''Friend Sub CambioUDMedida(ByVal current As IPropertyAccessor, ByVal Value As Object, ByVal Context As IPropertyAccessor, ByVal services As ServiceProvider)
''    Dim brd As New BusinessRuleData("IDUDMedida", Value, current, Context)
''    General.CambioUDMedida(brd, services)
''End Sub

'#End Region

''#Region " Contratos "

''    Public Function ValidarContratos(ByVal strIdContrato As String, ByVal strIDArticulo As String, ByVal strIDProveedor As String, ByVal dblCantidad As Double, ByVal dtFecha As Date) As DataTable
''        Const cnViewName As String = "vFrmPedidoCompraContratosArticulo"
''        Dim strSelect As String = "IdContrato,IDLineaContrato,QSobrante,Precio, IDUdCompra, UdValoracion"

''        Dim f As New Filter
''        f.Add(New StringFilterItem("IDContrato", strIdContrato))
''        f.Add(New StringFilterItem("IDArticulo", strIDArticulo))
''        f.Add(New StringFilterItem("IDProveedor", strIDProveedor))
''        f.Add(New DateFilterItem("FechaInicioContrato", FilterOperator.LessThanOrEqual, dtFecha))
''        f.Add(New DateFilterItem("FechaFinContrato", FilterOperator.GreaterThanOrEqual, dtFecha))
''        f.Add(New NumberFilterItem("QSobrante", FilterOperator.GreaterThanOrEqual, dblCantidad))

''        Return AdminData.GetData(cnViewName, f, strSelect)
''    End Function

''    Public Function ObtenerContratos(ByVal strIDArticulo As String, ByVal strIDProveedor As String, ByVal dblCantidad As Double, ByVal dtFecha As Date) As DataTable
''        Const cnViewName As String = "vFrmPedidoCompraContratosArticulo"
''        Dim strSelect As String = "IdContrato,IDLineaContrato,QSobrante,Precio, IDUdCompra, UdValoracion"

''        Dim f As New Filter
''        f.Add(New StringFilterItem("IDArticulo", strIDArticulo))
''        f.Add(New StringFilterItem("IDProveedor", strIDProveedor))
''        f.Add(New DateFilterItem("FechaInicioContrato", FilterOperator.LessThanOrEqual, dtFecha))
''        f.Add(New DateFilterItem("FechaFinContrato", FilterOperator.GreaterThanOrEqual, dtFecha))
''        f.Add(New NumberFilterItem("QSobrante", FilterOperator.GreaterThanOrEqual, dblCantidad))

''        Return AdminData.GetData(cnViewName, f, strSelect)
''    End Function

''#End Region

'#Region " TarifaCompra "

''<Task()> Public Shared Sub TarifaCompra(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
''    Dim C As New Compra
''    C.TarifaCompra(data.Context, data.Current, services)
''End Sub

''Private Function TarifaCompra(ByVal context As IPropertyAccessor, ByVal current As IPropertyAccessor, ByVal services As ServiceProvider) As IPropertyAccessor
''    If context.ContainsKey("IDOrdenRuta") AndAlso Length(context("IDOrdenRuta")) > 0 Then
''        'No hay que recuperar nada si se trata de una Subcontratación.
''    ElseIf current.ContainsKey("IDArticulo") And context.ContainsKey("IDProveedor") And current.ContainsKey("Cantidad") Then
''        If Not IsNumeric(current("Precio")) Then current("Precio") = 0
''        If Not IsNumeric(current("Dto1")) Then current("Dto1") = 0
''        If Not IsNumeric(current("Dto2")) Then current("Dto2") = 0
''        If Not IsNumeric(current("Dto3")) Then current("Dto3") = 0
''        If Not IsNumeric(current("UdValoracion")) Then current("UdValoracion") = 1
''        Dim BlnCambio As Boolean = False

''        If Length(current("IDArticulo")) > 0 And Length(context("IDProveedor")) > 0 And Nz(current("Cantidad"), 0) <> 0 Then
''            Dim FechaFactura As Date
''            If context.ContainsKey("Fecha") Then FechaFactura = context("Fecha")
''            Dim dtTarifa As DataTable = ObtenerContratos(current("IDArticulo"), context("IDProveedor"), current("Cantidad"), FechaFactura)
''            If Not IsNothing(dtTarifa) AndAlso dtTarifa.Rows.Count Then
''                If dtTarifa.Rows(0)("Precio") <> 0 Then
''                    current("Precio") = dtTarifa.Rows(0)("Precio")
''                End If
''                current("IdContrato") = dtTarifa.Rows(0)("IdContrato")
''                current("IdLineaContrato") = dtTarifa.Rows(0)("IdLineaContrato")
''                current("UdValoracion") = dtTarifa.Rows(0)("UdValoracion")
''                current("IDUdMedida") = dtTarifa.Rows(0)("IDUdCompra")
''                current("IDMoneda") = context("IDMoneda")
''                BlnCambio = False
''            Else
''                current("IdContrato") = DBNull.Value
''                current("IdLineaContrato") = DBNull.Value
''                Dim DblPrecio, DblDto1, DblDto2, DblDto3 As Double
''                Dim IntUdValor As Integer
''                Dim StrIDMoneda, StrIDUdCompra As String
''                Dim BlnEncontrado As Boolean = False
''                Dim DtArtProv As DataTable = New ArticuloProveedor().SelOnPrimaryKey(context("IDProveedor"), current("IDArticulo"))
''                If Not DtArtProv Is Nothing AndAlso DtArtProv.Rows.Count > 0 Then
''                    'Si el articulo esta asociado con el proveedor
''                    'Obtenemos el Precio correspondiente a la cantidad
''                    Dim FilProv As New Filter
''                    FilProv.Add("IdProveedor", FilterOperator.Equal, context("IDProveedor"))
''                    FilProv.Add("IDArticulo", FilterOperator.Equal, current("IDArticulo"))
''                    FilProv.Add("QDesde", FilterOperator.LessThanOrEqual, current("Cantidad"))
''                    Dim DtAPLinea As DataTable = New ArticuloProveedorLinea().Filter(FilProv, "QDesde DESC")
''                    If Not DtAPLinea Is Nothing AndAlso DtAPLinea.Rows.Count = 0 Then
''                        DtAPLinea = DtArtProv.Copy
''                    End If
''                    If Not DtAPLinea Is Nothing AndAlso DtAPLinea.Rows.Count > 0 Then
''                        Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
''                        Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(context("IDProveedor"))
''                        DblPrecio = DtAPLinea.Rows(0)("Precio")
''                        DblDto1 = DtAPLinea.Rows(0)("Dto1")
''                        DblDto2 = DtAPLinea.Rows(0)("Dto2")
''                        DblDto3 = DtAPLinea.Rows(0)("Dto3")
''                        IntUdValor = DtArtProv.Rows(0)("UdValoracion")
''                        StrIDMoneda = ProvInfo.IDMoneda
''                        StrIDUdCompra = DtArtProv.Rows(0)("IdUdCompra") & String.Empty
''                        BlnEncontrado = True
''                        BlnCambio = False
''                    End If
''                End If
''                ' Puede que encontremos la relación artículo - proveedor pero que unicamente tengamos descuentos sin precio
''                If DblPrecio = 0 Then
''                    Dim DrArt As DataRow = New Articulo().GetItemRow(current("IDArticulo"))
''                    DblPrecio = Nz(DrArt("PrecioEstandarA"), 0)
''                    If Len(StrIDUdCompra) = 0 Then
''                        If Length(current("IDUDMedida")) > 0 Then
''                            StrIDUdCompra = current("IDUDMedida")
''                        ElseIf Len(DrArt("IDUDInterna")) > 0 Then
''                            StrIDUdCompra = DrArt("IDUDInterna")
''                        End If
''                    End If
''                    BlnCambio = True
''                End If
''                If Not BlnEncontrado Then
''                    DescuentosFamilia(context("IDProveedor"), current("IDArticulo"), DblDto1, DblDto2, DblDto3)
''                    Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
''                    StrIDMoneda = Monedas.MonedaA.ID

''                    BlnEncontrado = True
''                End If
''                If BlnEncontrado Then
''                    If DblPrecio <> 0 Then
''                        current("Precio") = DblPrecio : current("IdMoneda") = StrIDMoneda
''                    End If
''                    current("Dto1") = DblDto1 : current("Dto2") = DblDto2
''                    current("Dto3") = DblDto3
''                    If IntUdValor <> 0 Then current("UdValoracion") = IntUdValor
''                    current("IDUDMedida") = StrIDUdCompra
''                End If

''            End If
''            If current("IDUDInterna") <> current("IDUDMedida") Then
''                Dim DblFactor As Double = New Articulo().FactorConversion(current("IDArticulo"), current("IDUDMedida"), current("IDUDInterna"))
''                If Length(DblFactor) > 0 Then
''                    current("Factor") = DblFactor
''                Else
''                    current("Factor") = 1
''                End If
''            End If
''            If BlnCambio = True Then
''                current("Precio") = current("Precio") * current("Factor")
''            End If
''        End If
''    End If

''    Return current
''End Function



'#End Region

#End Region


