Public Class ProcesoCompra

#Region " DetailCommonUpdateRules "

    'Public Function DetailCommonUpdateRules(ByVal dttSource As DataTable) As Boolean
    '    If Not IsNothing(dttSource) AndAlso dttSource.Rows.Count Then
    '        Dim services As New ServiceProvider
    '        For Each dr As DataRow In dttSource.Rows
    '            If DetailCommonUpdateRules(dr, services) Then
    '                DetailCommonUpdateRules = True
    '            Else
    '                Return False
    '            End If
    '        Next
    '    End If
    'End Function

    <Task()> Public Shared Sub DetailCommonUpdateRules(ByVal dr As DataRow, ByVal services As ServiceProvider)
        If Length(dr("IdArticulo")) = 0 Then
            ApplicationService.GenerateError("El artículo es obligatorio.")
        ElseIf Length(dr("DescArticulo")) = 0 Then
            ApplicationService.GenerateError("La descripción es obligatoria.")
        ElseIf Length(dr("IDCentroGestion")) = 0 Then
            If Not dr.Table.Columns.Contains("TipoLineaAlbaran") OrElse dr("TipoLineaAlbaran") = enumaclTipoLineaAlbaran.aclNormal Then
                ApplicationService.GenerateError("El Centro de Gestión es obligatorio.")
            End If
        ElseIf Length(dr("IdTipoIva")) = 0 Then
            ApplicationService.GenerateError("El tipo de IVA es obligatorio.")
        ElseIf Length(dr("CContable")) = 0 Then
            Dim AppParamsConta As ParametroContabilidadCompra = services.GetService(Of ParametroContabilidadCompra)()
            If AppParamsConta.Contabilidad Then ApplicationService.GenerateError("La cuenta contable es obligatoria.")
        ElseIf Not IsNumeric(dr("Precio")) Then
            ApplicationService.GenerateError("El precio no es válido.")
        ElseIf Not IsNumeric(dr("Dto1")) Then
            ApplicationService.GenerateError("El Dto. 1 no es válido.")
        ElseIf Not IsNumeric(dr("Dto2")) Then
            ApplicationService.GenerateError("El Dto. 2 no es válido.")
        ElseIf Not IsNumeric(dr("Dto3")) Then
            ApplicationService.GenerateError("El Dto. 3 no es válido.")
        ElseIf Not IsNumeric(dr("Dto")) Then
            ApplicationService.GenerateError("El Dto. Comercial no es válido.")
        ElseIf Not IsNumeric(dr("DtoProntoPago")) Then
            ApplicationService.GenerateError("El Dto. Pronto Pago no es válido.")
        ElseIf Not IsNumeric(dr("Factor")) Then
            ApplicationService.GenerateError("El factor no es válido.")
        ElseIf Not IsNumeric(dr("QInterna")) Then
            ApplicationService.GenerateError("La cantidad interna no es válida")
        ElseIf Not IsNumeric(dr("UdValoracion")) Then
            ApplicationService.GenerateError("La unidad de valoración no es válida.")
        ElseIf dr.Table.Columns.Contains("IDTipoLinea") AndAlso Length(dr("IDTipoLinea")) = 0 Then
            ApplicationService.GenerateError("El tipo de línea no es válido.")
        ElseIf dr.Table.Columns.Contains("TipoLineaCompra") AndAlso Length(dr("TipoLineaCompra")) = 0 Then
            ApplicationService.GenerateError("El tipo de línea no es válido.")
        Else
            If dr("UdValoracion") <= 0 Then
                ApplicationService.GenerateError("La unidad de valoración debe ser positiva.")
            ElseIf dr("Factor") <= 0 Then
                ApplicationService.GenerateError("El factor debe ser positivo.")
            End If
        End If
    End Sub

#End Region

#Region " DetailCommonBusinessRules - CABECERAS"

    <Task()> Public Shared Function DetailBusinessRulesCab(ByVal oBRL As BusinessRules, ByVal services As ServiceProvider) As BusinessRules
        If oBRL Is Nothing Then oBRL = New BusinessRules
        oBRL.Add("IDProveedor", AddressOf ProcesoCompra.CambioProveedor)
        oBRL.Add("IDCondicionPago", AddressOf NegocioGeneral.CambioCondicionPago)
        oBRL.Add("IDMoneda", AddressOf ProcesoComunes.CambioMoneda)
        Return oBRL
    End Function

    <Task()> Public Shared Sub CambioProveedor(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If data.ColumnName = "IDProveedor" Then data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDProveedor")) Then
            Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
            Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(data.Current("IDProveedor"))
            data.Current("CifProveedor") = ProvInfo.CifProveedor
            data.Current("RazonSocial") = ProvInfo.RazonSocial
            data.Current("Direccion") = ProvInfo.Direccion
            data.Current("CodPostal") = ProvInfo.CodPostal
            data.Current("Poblacion") = ProvInfo.Poblacion
            data.Current("Provincia") = ProvInfo.Provincia
            data.Current("IdMoneda") = ProvInfo.IDMoneda

            data.Current("IDFormaPago") = ProvInfo.IDFormaPago
            data.Current("IdCondicionPago") = ProvInfo.IDCondicionPago
            ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf NegocioGeneral.CambioCondicionPago, data, services)
            data.Current("IdFormaEnvio") = ProvInfo.IDFormaEnvio
            data.Current("IDCondicionEnvio") = ProvInfo.IDCondicionEnvio
            data.Current("IDDiaPago") = ProvInfo.IDDiaPago
        Else
            data.Current("CifProveedor") = System.DBNull.Value
            data.Current("RazonSocial") = System.DBNull.Value
            data.Current("Direccion") = System.DBNull.Value
            data.Current("CodPostal") = System.DBNull.Value
            data.Current("Poblacion") = System.DBNull.Value
            data.Current("Provincia") = System.DBNull.Value
            data.Current("IDFormaPago") = System.DBNull.Value
            data.Current("IdCondicionPago") = System.DBNull.Value
            ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf NegocioGeneral.CambioCondicionPago, data, services)
            data.Current("IdFormaEnvio") = System.DBNull.Value
            data.Current("IDCondicionEnvio") = System.DBNull.Value
            data.Current("IDDiaPago") = System.DBNull.Value
            data.Current("IdMoneda") = System.DBNull.Value
        End If
    End Sub

    <Task()> Public Shared Sub AsignarDireccionProveedor(ByVal data As DataDireccionProv, ByVal services As ServiceProvider)
        If Length(data.Datos("IDProveedor")) = 0 Then Exit Sub
        Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
        Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(data.Datos("IDProveedor"))

        Dim strProveedor As String = ProvInfo.IDProveedor
        ' If ProvInfo.GrupoDireccion Then strProveedor = ProvInfo.GrupoProveedor
        Dim stDatosDirec As New ProveedorDireccion.DataDirecEnvio
        stDatosDirec.IDProveedor = strProveedor
        stDatosDirec.TipoDireccion = data.TipoDireccion
        Dim dtDireccion As DataTable = ProcessServer.ExecuteTask(Of ProveedorDireccion.DataDirecEnvio, DataTable)(AddressOf ProveedorDireccion.ObtenerDireccionEnvio, stDatosDirec, services)
        If Not dtDireccion Is Nothing AndAlso dtDireccion.Rows.Count > 0 Then
            data.Datos(data.Field) = dtDireccion.Rows(0)("IDDireccion")
        End If
    End Sub

    <Task()> Public Shared Sub AsignarObservacionesProveedor(ByVal data As DataObservaciones, ByVal services As ServiceProvider)
        If Length(data.Datos("IDProveedor")) > 0 Then
            Dim StDatos As New Observacion.DatosObv
            StDatos.IDEntidad = data.Entity
            StDatos.IDPrimaryKey = data.Datos("IDProveedor")
            data.Datos(data.Field) = ProcessServer.ExecuteTask(Of Observacion.DatosObv, String)(AddressOf Observacion.ObtenerObservacionesProveedor, StDatos, services)
        Else : data.Datos(data.Field) = System.DBNull.Value
        End If
    End Sub
#End Region

#Region " DetailCommonBusinessRules - LINEAS"

    <Task()> Public Shared Function DetailBusinessRulesLin(ByVal oBRL As BusinessRules, ByVal services As ServiceProvider) As BusinessRules
        oBRL.Add("IDArticulo", AddressOf ProcesoCompra.CambioArticulo)
        oBRL.Add("RefProveedor", AddressOf ProcesoCompra.CambioArticulo)
        oBRL.Add("Cantidad", AddressOf ProcesoCompra.CambioCantidad)
        oBRL.Add("QTiempo", AddressOf ProcesoCompra.CambioCantidad)
        oBRL.Add("IDContrato", AddressOf ProcesoCompra.CambioContrato)
        oBRL.Add("QInterna", AddressOf ProcesoCompra.CambioQInterna)
        oBRL.Add("QInterna2", AddressOf CambioCantidadInterna2)
        oBRL.Add("Factor", AddressOf NegocioGeneral.CambioFactor)
        oBRL.Add("IDUDMedida", AddressOf NegocioGeneral.CambioUDMedida)
        oBRL.Add("Precio", AddressOf NegocioGeneral.CambioPrecio)
        oBRL.Add("UDValoracion", AddressOf NegocioGeneral.CambioPrecio)
        oBRL.Add("Dto1", AddressOf NegocioGeneral.CambioPrecio)
        oBRL.Add("Dto2", AddressOf NegocioGeneral.CambioPrecio)
        oBRL.Add("Dto3", AddressOf NegocioGeneral.CambioPrecio)
        oBRL.Add("Dto", AddressOf NegocioGeneral.CambioPrecio)
        oBRL.Add("DtoProntoPago", AddressOf NegocioGeneral.CambioPrecio)
        oBRL.Add("PrecioA", AddressOf NegocioGeneral.CalcularImportes)
        oBRL.Add("PrecioB", AddressOf NegocioGeneral.CalcularImportes)
        oBRL.Add("CContable", AddressOf CambioCContableCompra)
        oBRL.Add("IDObra", AddressOf ProcesoCompra.CambioObra)
        oBRL.Add("NObra", AddressOf ProcesoCompra.CambioObra)
        oBRL.Add("CodTrabajo", AddressOf ProcesoCompra.CambioCodTrabajo)
        oBRL.Add("IDLineaAlbaran", AddressOf ProcesoCompra.CambioLineaAlbaran)
        oBRL.Add("Inmovilizado", AddressOf ProcesoCompra.CambioInmovilizado)
        oBRL.Add("FechaEntrega", AddressOf ProcesoCompra.CambioFechaEntrega)

        Return oBRL
    End Function

#Region " CambioArticulo "

    <Task()> Public Shared Sub CambioArticulo(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ValidarProveedorCabecera, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf RecuperarInformacionArticulo, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf AsignarValoresPredeterminados, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf AsignarDatosArticulo, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf NegocioGeneral.AsignarArticuloAlmacen, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf TratarCContable, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ObtenerIvaArticuloProveedor, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf RecuperarInformacionArticuloProveedor, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf NegocioGeneral.FactorConversion, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf AplicarTarifaCompra, data, services)
    End Sub

    <Task()> Public Shared Sub ValidarProveedorCabecera(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Context("Origen")) > 0 AndAlso data.Context("Origen") = "SolicitudCompraLinea" Then Exit Sub
        If data.Context.ContainsKey("IDProveedor") AndAlso Length(data.Context("IDProveedor")) = 0 Then
            If Not data.Context.ContainsKey("MensajeFaltaProveedor") OrElse Nz(data.Context("MensajeFaltaProveedor"), False) Then
                ApplicationService.GenerateError("No ha especificado el Proveedor.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub RecuperarInformacionArticulo(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current(data.ColumnName)) Then
            Dim ArtInfo As ArticuloInfo
            'TODO: Pensar como taskear InformacionArticulo
            Select Case data.ColumnName
                Case "IDArticulo"
                    Dim stArticulo As New Articulo.DataInfoArticulo(data.Current("IDArticulo"))
                    ArtInfo = ProcessServer.ExecuteTask(Of Articulo.DataInfoArticulo, ArticuloInfo)(AddressOf Articulo.InformacionArticulo, stArticulo, services)
                    'Case "CodigoBarras"
                    '    ArtInfo = New Articulo().InformacionArticulo(Nothing, data.Current("CodigoBarras"))
                Case "RefProveedor"
                    Dim stArticulo As New Articulo.DataInfoArticulo(Nothing, Nothing, data.Current("RefProveedor"), data.Context)
                    ArtInfo = ProcessServer.ExecuteTask(Of Articulo.DataInfoArticulo, ArticuloInfo)(AddressOf Articulo.InformacionArticulo, stArticulo, services)
                    If Not ArtInfo Is Nothing AndAlso Length(ArtInfo.IDArticulo) > 0 Then
                        data.Current("IDArticulo") = ArtInfo.IDArticulo
                    End If
            End Select

            If Not ArtInfo Is Nothing AndAlso Length(ArtInfo.IDArticulo) > 0 Then
                '//Registramos el artículo de manera manual, en lugar de hacerlo a través del services, por que dependiendo de como
                '// se intente acceder recuperaremos el artículo de una manera u otra.
                Dim Articulos As New EntityInfoCache(Of ArticuloInfo)
                Articulos.Add(data.Current("IDArticulo")) = ArtInfo
                services.RegisterService(Articulos, GetType(EntityInfoCache(Of ArticuloInfo)))
            End If
        ElseIf Length(data.Current(data.ColumnName)) = 0 AndAlso data.ColumnName = "RefProveedor" Then
            If data.Context.ContainsKey("IDArticuloRef") Then
                data.Current("IDArticulo") = data.Context("IDArticuloRef")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarValoresPredeterminados(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Current("UDValoracion")) = 0 OrElse data.Current("UDValoracion") = 0 Then data.Current("UDValoracion") = 1
        If Length(data.Current("IDArticulo")) > 0 Then
            If ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.Current("IDArticulo"), services) Then
                If data.Current.ContainsKey("Cantidad") AndAlso Length(data.Current("Cantidad")) = 0 Then
                    data.Current("Cantidad") = 0
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarDatosArticulo(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Current("IDArticulo")) = 0 Then Exit Sub

        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.Current("IDArticulo"))

        data.Current("IDArticulo") = ArtInfo.IDArticulo
        data.Current("DescArticulo") = ArtInfo.DescArticulo
        If Not ArtInfo.Activo Then
            ApplicationService.GenerateError("El artículo | no está activo.", Quoted(data.Current("IDArticulo")))
        ElseIf Not ArtInfo.Compra Then
            If Not data.Context.ContainsKey("IDOrdenRuta") OrElse Length(data.Context("IDOrdenRuta")) = 0 Then
                If Not data.Context.ContainsKey("TipoLineaCompra") OrElse data.Context("TipoLineaCompra") <> enumaclTipoLineaAlbaran.aclComponente Then
                    ApplicationService.GenerateError("El artículo | no es de tipo compra.", Quoted(data.Current("IDArticulo")))
                End If
            End If
            Else
                If ArtInfo.NSerieObligatorio AndAlso Nz(data.Current("Cantidad"), 0) = 0 Then
                    data.Current("Cantidad") = 1
                    data.Current("QInterna") = 1
                End If
            Dim AppParamsCompra As ParametroCompra = services.GetService(Of ParametroCompra)()
            If ArtInfo.Subcontratacion AndAlso Length(data.Context("IDTipoCompraCabecera")) > 0 AndAlso data.Context("IDTipoCompraCabecera") = AppParamsCompra.TipoCompraSubcontratacion Then
                data.Current("TipoLineaCompra") = enumaclTipoLineaAlbaran.aclSubcontratacion
            Else
                data.Current("TipoLineaCompra") = enumaclTipoLineaAlbaran.aclNormal
            End If
                If Length(ArtInfo.IDUDCompra) = 0 Then
                    ApplicationService.GenerateError("Debe indicar una unidad de compra para el artículo {0}.", Quoted(data.Current("IDArticulo")))
                Else
                    data.Current("IDUDMedida") = ArtInfo.IDUDCompra
                End If
                data.Current("IDUDInterna") = ArtInfo.IDUDInterna

                If data.Current.ContainsKey("IDUDInterna2") AndAlso ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.Current("IDArticulo"), services) Then
                    data.Current("IDUDInterna2") = ArtInfo.IDUDInterna2
                End If
                data.Current("UdValoracion") = ArtInfo.UDValoracion
                data.Current("IDTipoIva") = ArtInfo.IDTipoIVA
                data.Current("IDUDCompra") = ArtInfo.IDUDCompra
                If ArtInfo.Especial Then
                    If data.Current.ContainsKey("Especial") Then
                        data.Current("Especial") = ArtInfo.Especial
                    End If
                    data.Current("Dto") = 0
                    data.Current("DtoProntoPago") = 0
                End If

                If data.Current.ContainsKey("NSerieObligatorio") Then
                    data.Current("NSerieObligatorio") = ArtInfo.GestionPorNumeroSerie
                End If

                If data.Current.ContainsKey("GestionStockPorLotes") Then
                    data.Current("GestionStockPorLotes") = ArtInfo.GestionStockPorLotes
                End If
            End If
    End Sub

    <Task()> Public Shared Sub RecuperarInformacionArticuloProveedor(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim ArtProvs As EntityInfoCache(Of ArticuloProveedorInfo) = services.GetService(Of EntityInfoCache(Of ArticuloProveedorInfo))()
        Dim ArtProv As ArticuloProveedorInfo = ArtProvs.GetEntity(data.Context("IDProveedor"), data.Current("IDArticulo"))
        If Not IsNothing(ArtProv) AndAlso Length(ArtProv.IDProveedor) > 0 AndAlso Length(ArtProv.IDArticulo) > 0 Then
            data.Current("RefProveedor") = ArtProv.RefProveedor
            data.Current("DescRefProveedor") = ArtProv.DescRefProveedor
            If Length(ArtProv.IDUDCompra) > 0 Then data.Current("IDUDMedida") = ArtProv.IDUDCompra : data.Current("IDUDCompra") = ArtProv.IDUDCompra
            If ArtProv.UdValoracion <> 0 Then data.Current("UdValoracion") = ArtProv.UdValoracion

            ''Incluir control de precio de relación artículo - proveedor
            'If Length(ArtProv.Precio) > 0 AndAlso Length(ArtProv.Precio) <> 0 Then
            '    data.Current("Precio") = ArtProv.Precio
            'End If
        End If

        If Length(data.Current("DescRefProveedor")) = 0 Then
            Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
            If Length(data.Context("IDProveedor")) > 0 Then
                Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(data.Context("IDProveedor"))
                If Length(ProvInfo.IDIdioma) > 0 Then
                    Dim dtIdioma As DataTable = New ArticuloIdioma().SelOnPrimaryKey(data.Current("IDArticulo"), ProvInfo.IDIdioma)
                    If Not IsNothing(dtIdioma) AndAlso dtIdioma.Rows.Count > 0 Then
                        data.Current("DescRefProveedor") = dtIdioma.Rows(0)("DescArticuloIdioma")
                    End If
                End If
            End If
        End If
    End Sub


#End Region

#Region " TratarCContable "

    <Serializable()> _
    Public Class DataGetCContableCompra
        Public Datos As IPropertyAccessor
        Public TipoCuenta As TipoCContableCompra
    End Class

    Public Enum TipoCContableCompra
        CCImport
        CCCompra
        CCImportGRUPO
        CCCompraGRUPO
    End Enum

    <Task()> Public Shared Sub TratarCContable(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim AppParams As ParametroContabilidadCompra = services.GetService(Of ParametroContabilidadCompra)()
        If Not AppParams.Contabilidad Then Exit Sub

        Dim ProvInfo As ProveedorInfo
        If data.Context.ContainsKey("IDProveedor") AndAlso Length(data.Context("IDProveedor")) > 0 Then
            Dim datGetCC As New DataGetCContableArticuloProveedor(data.Current("IDArticulo") & String.Empty, data.Context("IDProveedor"))
            data.Current("CContable") = ProcessServer.ExecuteTask(Of DataGetCContableArticuloProveedor, String)(AddressOf GetCContableArticuloProveedor, datGetCC, services)
            ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf NegocioGeneral.FormatoCuentaContable, data, services)
        End If
    End Sub


    <Serializable()> _
    Public Class DataGetCContableArticuloProveedor
        Public IDArticulo As String
        Public IDProveedor As String

        Public Sub New(ByVal IDArticulo As String, ByVal IDProveedor As String)
            Me.IDArticulo = IDArticulo
            Me.IDProveedor = IDProveedor
        End Sub
    End Class
    <Task()> Public Shared Function GetCContableArticuloProveedor(ByVal data As DataGetCContableArticuloProveedor, ByVal services As ServiceProvider) As String
        If Length(data.IDArticulo) > 0 AndAlso Length(data.IDProveedor) > 0 Then
            Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
            Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(data.IDProveedor)

            Dim current As New BusinessData
            current("CContable") = Nothing
            current("IDArticulo") = data.IDArticulo

            Dim Cta As New DataGetCContableCompra
            If ProvInfo.EmpresaGrupo Then
                If Not ProvInfo.Extranjero Then
                    Cta.Datos = current
                    Cta.TipoCuenta = TipoCContableCompra.CCCompraGRUPO
                Else
                    Cta.Datos = current
                    Cta.TipoCuenta = TipoCContableCompra.CCImportGRUPO
                End If
            ElseIf Not ProvInfo.Extranjero Then
                Cta.Datos = current
                Cta.TipoCuenta = TipoCContableCompra.CCCompra
            Else
                Cta.Datos = current
                Cta.TipoCuenta = TipoCContableCompra.CCImport
            End If
            ProcessServer.ExecuteTask(Of DataGetCContableCompra)(AddressOf GetCContableArticulo, Cta, services)

            Return current("CContable")
        End If

    End Function


    <Task()> Public Shared Sub GetCContableArticulo(ByVal data As DataGetCContableCompra, ByVal services As ServiceProvider)
        Dim AppParamsConta As ParametroContabilidadCompra = services.GetService(Of ParametroContabilidadCompra)()
        If Not AppParamsConta.Contabilidad Then Exit Sub

        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.Datos("IDArticulo"))
        Dim strCContable As String : Dim strField As String

        Select Case data.TipoCuenta
            Case TipoCContableCompra.CCImport
                strField = "CCImport"
                strCContable = ArtInfo.CCImport
            Case TipoCContableCompra.CCImportGRUPO
                strField = "CCImportGRUPO"
                strCContable = ArtInfo.CCImportGrupo
            Case TipoCContableCompra.CCCompra
                strField = "CCCompra"
                strCContable = ArtInfo.CCCompra
            Case TipoCContableCompra.CCCompraGRUPO
                strField = "CCCompraGRUPO"
                strCContable = ArtInfo.CCCompraGrupo
        End Select

        If Length(strCContable) > 0 Then
            data.Datos("CContable") = strCContable
        Else
            Dim f As New Filter
            f.Add(New StringFilterItem("IDTipo", ArtInfo.IDTipo))
            f.Add(New StringFilterItem("IDFamilia", ArtInfo.IDFamilia))
            Dim dtFam As DataTable = New Familia().Filter(f)
            If Not dtFam Is Nothing AndAlso dtFam.Rows.Count > 0 Then
                data.Datos("CContable") = dtFam.Rows(0)(strField)
            End If
            If Length(data.Datos("CContable")) = 0 Then
                Select Case data.TipoCuenta
                    Case TipoCContableCompra.CCImport
                        data.Datos("CContable") = AppParamsConta.CuentaImportacion
                    Case TipoCContableCompra.CCImportGRUPO
                        data.Datos("CContable") = AppParamsConta.CuentaImportacionGrupo
                    Case TipoCContableCompra.CCCompra
                        data.Datos("CContable") = AppParamsConta.CuentaCompra
                    Case TipoCContableCompra.CCCompraGRUPO
                        data.Datos("CContable") = AppParamsConta.CuentaCompraGrupo
                End Select
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioCContableCompra(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf NegocioGeneral.CambioCContable, data, services)
        Dim dataCCI As New ProcesoCompra.DataCContableInmovilizado
        dataCCI.IDEjercicio = data.Context("IDEjercicio") & String.Empty
        dataCCI.CContable = data.Current("CContable") & String.Empty
        dataCCI.Inmovilizado = Nz(data.Current("Inmovilizado"), False)
        '   ProcessServer.ExecuteTask(Of ProcesoCompra.DataCContableInmovilizado)(AddressOf ProcesoCompra.ValidarCuentaInmovilizado, dataCCI, services)
    End Sub
#End Region

#Region " Cambio Cantidad "

    <Task()> Public Shared Sub CambioCantidadInterna2(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If IsNumeric(data.Current("Qinterna2")) Then
            ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoComunes.CambioSegundaUnidad, data, services)
            If data.Current.ContainsKey("CambioQInterna") AndAlso data.Current("CambioQInterna") Then ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoCompra.AplicarTarifaCompra, data, services)
        Else
            ApplicationService.GenerateError("Campo no numérico.")
        End If
    End Sub

    <Task()> Public Shared Sub CambioCantidad(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If IsNumeric(data.Value) Then
            ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf NegocioGeneral.CalculoQInterna, data, services)
            ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoCompra.AplicarTarifaCompra, data, services)
            If data.Context.ContainsKey("IDMoneda") AndAlso Length(data.Context("IDMoneda")) > 0 Then
                Dim ValAyB As New ValoresAyB(data.Current, data.Context("IDMoneda"), data.Context("CambioA"), data.Context("CambioB"))
                ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)

                ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf NegocioGeneral.CalcularImportes, data, services)
            End If
        Else
            ApplicationService.GenerateError("Campo no numérico.")
        End If
    End Sub


#End Region

#Region " CambioQInterna "

    <Task()> Public Shared Sub CambioQInterna(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf NegocioGeneral.CambioQInterna, data, services)
        If data.Current.ContainsKey("SimularCambioCantidad") AndAlso data.Current("SimularCambioCantidad") Then
            '// Para evitar modificar la ubicación de las tareas, por si hay sobreescrituras y debido a que no va a ser necesario en todos los casos,
            '// incluimos en el Current una variable para ver si hay que simular el Cambio Manual del campo cantidad, para el Recalculo de Tarifas 
            '// y otros posibles cambios referentes al cambio de Cantidad.  (Ver NegocioGeneral.CambioQInterna)
            ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoCompra.CambioCantidad, data, services)
        End If
    End Sub

#End Region

#Region " Contratos "

    <Serializable()> _
    Public Class DataContrato
        Public IDContrato As String
        Public IDArticulo As String
        Public IDProveedor As String
        Public Fecha As Date
        Public QSobrante As Double
    End Class

    <Task()> Public Shared Sub CambioContrato(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDContrato")) > 0 Then
            Dim contrato As New DataContrato
            contrato.IDContrato = data.Current("IDContrato")
            contrato.IDArticulo = data.Current("IDArticulo")
            contrato.IDProveedor = data.Context("IDProveedor")
            contrato.Fecha = data.Context("Fecha")
            contrato.QSobrante = data.Current("Cantidad")

            Dim dtContrato As DataTable = ProcessServer.ExecuteTask(Of DataContrato, DataTable)(AddressOf ValidarContratos, contrato, services)
            If Not IsNothing(dtContrato) AndAlso dtContrato.Rows.Count Then
                '//Se le había dato precio en otra moneda diferente a la del Proveedor,
                '//Por lo que tenemos que hacer la conversión.
                Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
                Dim m As MonedaInfo = Monedas.GetMoneda(data.Context("IDMoneda"), data.Context("Fecha"))
                data.Current("Precio") = xRound(dtContrato.Rows(0)("Precio") * m.CambioA, m.NDecimalesPrecio)

                'data.Current("Precio") = dtContrato.Rows(0)("Precio")
                data.Current("IdContrato") = dtContrato.Rows(0)("IdContrato")
                data.Current("IdLineaContrato") = dtContrato.Rows(0)("IdLineaContrato")
                data.Current("UdValoracion") = dtContrato.Rows(0)("UdValoracion")
                data.Current("IDUdMedida") = dtContrato.Rows(0)("IDUdCompra")
                '//Se ha modificado el precio, hay que recalcular el importe.
                ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf NegocioGeneral.CalcularImportes, data, services)
            Else
                ApplicationService.GenerateError("EL contrato | no es válido", Quoted(data.Current("IDContrato")))
            End If
        End If
    End Sub

    <Task()> Public Shared Function ValidarContratos(ByVal data As DataContrato, ByVal services As ServiceProvider) As DataTable
        Const cnViewName As String = "vFrmPedidoCompraContratosArticulo"
        Dim strSelect As String = "IdContrato,IDLineaContrato,QSobrante,Precio, IDUdCompra, UdValoracion, IDMoneda"

        Dim f As New Filter
        f.Add(New StringFilterItem("IDContrato", data.IDContrato))
        f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
        f.Add(New StringFilterItem("IDProveedor", data.IDProveedor))
        f.Add(New DateFilterItem("FechaInicioContrato", FilterOperator.LessThanOrEqual, data.Fecha))
        f.Add(New DateFilterItem("FechaFinContrato", FilterOperator.GreaterThanOrEqual, data.Fecha))
        f.Add(New NumberFilterItem("QSobrante", FilterOperator.GreaterThanOrEqual, data.QSobrante))

        Return New BE.DataEngine().Filter(cnViewName, f, strSelect)
    End Function

    <Task()> Public Shared Function ObtenerContratos(ByVal data As DataContrato, ByVal services As ServiceProvider) As DataTable
        Const cnViewName As String = "vFrmPedidoCompraContratosArticulo"
        Dim strSelect As String = "IdContrato,IDLineaContrato,QSobrante,Precio, IDUdCompra, UdValoracion, IDMoneda"

        Dim f As New Filter
        f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
        f.Add(New StringFilterItem("IDProveedor", data.IDProveedor))
        f.Add(New DateFilterItem("FechaInicioContrato", FilterOperator.LessThanOrEqual, data.Fecha))
        f.Add(New DateFilterItem("FechaFinContrato", FilterOperator.GreaterThanOrEqual, data.Fecha))
        f.Add(New NumberFilterItem("QSobrante", FilterOperator.GreaterThanOrEqual, data.QSobrante))

        Return New BE.DataEngine().Filter(cnViewName, f, strSelect)
    End Function

#End Region

#Region " Inmovilizado "

    <Serializable()> _
    Public Class DataCContableInmovilizado
        Public IDEjercicio As String
        Public CContable As String
        Public Inmovilizado As Boolean
    End Class

    <Task()> Public Shared Sub ValidarCuentaInmovilizado(ByVal data As DataCContableInmovilizado, ByVal services As ServiceProvider)
        If Length(data.CContable) > 0 AndAlso Length(data.IDEjercicio) > 0 Then
            Dim objFilter As New Filter
            objFilter.Add(New StringFilterItem("IDEjercicio", data.IDEjercicio))
            objFilter.Add(New StringFilterItem("IDCContable", data.CContable))

            Dim PlanContable As BusinessHelper = BusinessHelper.CreateBusinessObject("PlanContable")
            Dim dtCContable As DataTable = PlanContable.Filter(objFilter)
            If Not IsNothing(dtCContable) AndAlso dtCContable.Rows.Count > 0 Then
                If data.Inmovilizado AndAlso Not Nz(dtCContable.Rows(0)("Inversion"), False) Then
                    ApplicationService.GenerateError("La C.Contable | no es de Inmovilizado en el Ejercicio |.", Quoted(data.CContable), Quoted(data.IDEjercicio))
                ElseIf Not data.Inmovilizado AndAlso Nz(dtCContable.Rows(0)("Inversion"), False) Then
                    ApplicationService.GenerateError("La C.Contable | es de Inmovilizado en el Ejercicio |.", Quoted(data.CContable), Quoted(data.IDEjercicio))
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioInmovilizado(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("Inmovilizado")) > 0 AndAlso data.Current.ContainsKey("CContable") AndAlso data.Context.ContainsKey("IDEjercicio") Then
            Dim dataCCI As New DataCContableInmovilizado
            dataCCI.IDEjercicio = data.Context("IDEjercicio") & String.Empty
            dataCCI.Inmovilizado = data.Current("Inmovilizado")
            dataCCI.CContable = data.Current("CContable") & String.Empty
            ProcessServer.ExecuteTask(Of DataCContableInmovilizado)(AddressOf ValidarCuentaInmovilizado, dataCCI, services)
        End If
    End Sub


#End Region

#Region " Cambio Obra / Trabajo "

    <Task()> Public Shared Sub CambioObra(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If data.ColumnName = "NObra" AndAlso Length(data.Value) = 0 Then
            data.Current("IDObra") = System.DBNull.Value
        End If
        data.Current("IDTrabajo") = DBNull.Value
        data.Current("CodTrabajo") = DBNull.Value
        If data.Current.ContainsKey("IDLineaMaterial") Then data.Current("IDLineaMaterial") = System.DBNull.Value
        If data.Current.ContainsKey("IDLineaPadre") Then data.Current("IDLineaPadre") = System.DBNull.Value

        If Length(data.Current("IDObra")) > 0 Then
            Dim obra As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraCabecera")
            Dim dtObra As DataTable = obra.SelOnPrimaryKey(data.Current("IDObra"))
            If IsNothing(dtObra) OrElse dtObra.Rows.Count = 0 Then
                ApplicationService.GenerateError("El Proyecto | no existe.", Quoted(data.Current("IDObra")))
            Else
                Dim objFilter As New Filter
                objFilter.Add(New NumberFilterItem("Estado", FilterOperator.NotEqual, enumocEstado.ocTerminado))
                Dim dv As New DataView(dtObra)
                dv.RowFilter = objFilter.Compose(New AdoFilterComposer)
                If dv.Count > 0 Then
                    If data.Current.ContainsKey("TipoGastoObra") Then
                        If Length(data.Current("TipoGastoObra")) = 0 Then data.Current("TipoGastoObra") = CInt(enumfclTipoGastoObra.enumfclMaterial)
                        If data.Current("TipoGastoObra") = enumfclTipoGastoObra.enumfclMaterial Then
                            data.Current("IDConcepto") = data.Current("IDArticulo")
                        End If
                        '    If data.Current("EstadoFactura") = enumaclEstadoFactura.aclNoFacturado Then data.Current("GeneradoControl") = True
                    End If
                Else
                    ApplicationService.GenerateError("El Proyecto | está terminado.", Quoted(data.Current("NObra")))
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioCodTrabajo(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        data.Current("TipoGastoObra") = enumfclTipoGastoObra.enumfclMaterial
        If data.Current.ContainsKey("IDLineaMaterial") Then data.Current("IDLineaMaterial") = System.DBNull.Value
        If data.Current.ContainsKey("IDLineaPadre") Then data.Current("IDLineaPadre") = System.DBNull.Value
        If Length(data.Current("CodTrabajo")) = 0 Then
            data.Current("IDTrabajo") = DBNull.Value
        End If
        If Length(data.Current("IdTrabajo")) > 0 AndAlso Length(data.Current("IDObra")) > 0 Then
            Dim ClsObra As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraCabecera")
            Dim dtObra As DataTable = ClsObra.SelOnPrimaryKey(data.Current("IDObra"))
            If Not dtObra Is Nothing AndAlso dtObra.Rows.Count > 0 Then
                data.Current("NObra") = dtObra.Rows(0)("NObra")
            End If
        End If
    End Sub

#End Region

#Region " Cambio IDLineaAlbaran "

    <Task()> Public Shared Sub CambioLineaAlbaran(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If IsNumeric(data.Current("IDLineaAlbaran")) Then
            Dim dtACL As DataTable = New AlbaranCompraLinea().SelOnPrimaryKey(data.Current("IDLineaAlbaran"))
            If Not IsNothing(dtACL) AndAlso dtACL.Rows.Count Then
                data.Current("IdLineaAlbaran") = DBNull.Value
                data.Current("IdAlbaran") = dtACL.Rows(0)("IdAlbaran")
                data.Current("IDArticulo") = dtACL.Rows(0)("IDArticulo")
                data.Current("DescArticulo") = dtACL.Rows(0)("DescArticulo")
                data.Current("RefPorveedor") = dtACL.Rows(0)("RefProveedor")
                data.Current("Precio") = dtACL.Rows(0)("Precio")
                data.Current("UdValoracion") = dtACL.Rows(0)("UdValoracion")
                data.Current("Lote") = dtACL.Rows(0)("Lote")
                data.Current("IDUdMedida") = dtACL.Rows(0)("IDUdMedida")
                data.Current("IDUdInterna") = dtACL.Rows(0)("IDUdInterna")
                data.Current("IDTipoIVA") = dtACL.Rows(0)("IDTipoIVA")
                If Length(dtACL.Rows(0)("IDCentroGestion")) > 0 Then
                    data.Current("IDCentroGestion") = dtACL.Rows(0)("IDCentroGestion")
                End If
                data.Current("IDObra") = dtACL.Rows(0)("IdObra")
                data.Current("IDTrabajo") = dtACL.Rows(0)("IdTrabajo")
                data.Current("IdLineaMaterial") = dtACL.Rows(0)("IdLineaMaterial")
                data.Current("Dto1") = dtACL.Rows(0)("Dto1")
                data.Current("Dto2") = dtACL.Rows(0)("Dto2")
                data.Current("Dto3") = dtACL.Rows(0)("Dto3")
                data.Current("CContable") = dtACL.Rows(0)("CContable")
            End If
        End If
    End Sub

#End Region

#Region " Cambio FechaEntrega "

    <Task()> Public Shared Sub CambioFechaEntrega(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If data.Current.ContainsKey("FechaEntregaModificadoPedido") Then
            data.Current("FechaEntregaModificadoPedido") = data.Current("FechaEntega")
        End If
    End Sub

#End Region

#End Region

#Region " Obtener IVA "

    <Task()> Public Shared Function ObtenerIvaArticuloProveedor(ByVal data As BusinessRuleData, ByVal services As ServiceProvider) As String
        Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
        Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(data.Context("IDProveedor"))
        Dim IVA As String = String.Empty
        If Not IsNothing(ProvInfo) Then
            '//Si el parametro IVAProveedor está a 1, entonces la función devolverá siempre el iva del proveedor, sin importar la nacionalidad
            '//Si el proveedor es extranjero o "CanariasCeutaMelilla" se devuelve el iva del proveedor
            Dim AppParamsCompra As ParametroCompra = services.GetService(Of ParametroCompra)()
            If AppParamsCompra.IVAProveedor OrElse ProvInfo.Extranjero OrElse ProvInfo.CanariasCeutaMelilla Then
                IVA = ProvInfo.IDTipoIVA
                '//Si el Proveedor no tiene IVA, devolvemos el IVA del artículo.
                If Length(IVA) = 0 Then
                    Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                    Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.Current("IDArticulo"))

                    IVA = ArtInfo.IDTipoIVA
                End If
            Else
                '//Si el cliente es nacional, se devuelve el iva del artículo
                Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.Current("IDArticulo"))

                If ProvInfo.TipoRetencionIRPF = TipoRetencionIRPF.RegimenAgricola Then
                    IVA = ArtInfo.IDTipoIVAReducido
                Else
                    IVA = ArtInfo.IDTipoIVA
                End If

                '//Si el artículo no tiene iva, se devuelve el del proveedor.
                If Length(IVA) = 0 Then IVA = ProvInfo.IDTipoIVA
            End If

            If Length(IVA) = 0 Then
                ApplicationService.GenerateError("El código de IVA es obligatorio. Revise la relación Articulo-Proveedor o que el Artículo tenga IVA.")
            Else
                data.Current("IDTipoIva") = IVA
            End If

            Return IVA
        End If
    End Function

    <Serializable()> _
    Public Class DataIVAArtProvERP
        Public IDArticulo As String
        Public IDProveedor As String

        Public Sub New()
        End Sub
        Public Sub New(ByVal IDArticulo As String, ByVal IDProveedor As String)
            Me.IDArticulo = IDArticulo
            Me.IDProveedor = IDProveedor
        End Sub
    End Class

    <Task()> Public Shared Function ObtenerIvaArtProvERP(ByVal data As DataIVAArtProvERP, ByVal services As ServiceProvider) As String
        Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
        Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(data.IDProveedor)
        Dim IVA As String = String.Empty
        If Not IsNothing(ProvInfo) Then
            '//Si el parametro IVAProveedor está a 1, entonces la función devolverá siempre el iva del proveedor, sin importar la nacionalidad
            '//Si el proveedor es extranjero o "CanariasCeutaMelilla" se devuelve el iva del proveedor
            Dim AppParamsCompra As ParametroCompra = services.GetService(Of ParametroCompra)()
            If AppParamsCompra.IVAProveedor OrElse ProvInfo.Extranjero OrElse ProvInfo.CanariasCeutaMelilla Then
                IVA = ProvInfo.IDTipoIVA
                '//Si el Proveedor no tiene IVA, devolvemos el IVA del artículo.
                If Length(IVA) = 0 Then
                    Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                    Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.IDArticulo)

                    IVA = ArtInfo.IDTipoIVA
                End If
            Else
                '//Si el cliente es nacional, se devuelve el iva del artículo
                Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.IDArticulo)

                If ProvInfo.TipoRetencionIRPF = TipoRetencionIRPF.RegimenAgricola Then
                    IVA = ArtInfo.IDTipoIVAReducido
                Else
                    IVA = ArtInfo.IDTipoIVA
                End If

                '//Si el artículo no tiene iva, se devuelve el del proveedor.
                If Length(IVA) = 0 Then IVA = ProvInfo.IDTipoIVA
            End If

            If Length(IVA) = 0 Then ApplicationService.GenerateError("El código de IVA es obligatorio. Revise la relación Articulo-Proveedor o que el Artículo tenga IVA.")
            Return IVA
        End If
    End Function

#End Region

#Region " Tarifa Compra "

    <Task()> Public Shared Function TarifaCompra(ByVal data As DataCalculoTarifaCompra, ByVal services As ServiceProvider) As DataTarifaCompra
        Dim strSeguimientoTarifa As String
        If Length(data.IDArticulo) > 0 And data.Cantidad <> 0 Then
            data.DatosTarifa = New DataTarifaCompra
            data.DatosTarifa.Precio = 0
            data.DatosTarifa.Dto1 = 0
            data.DatosTarifa.Dto2 = 0
            data.DatosTarifa.Dto3 = 0
            data.DatosTarifa.UDValoracion = Nz(data.UDValoracion, 1)

            ProcessServer.ExecuteTask(Of DataCalculoTarifaCompra)(AddressOf TarifaContratoActivo, data, services)
            ProcessServer.ExecuteTask(Of DataCalculoTarifaCompra)(AddressOf TarifaArticuloProveedor, data, services)
            ProcessServer.ExecuteTask(Of DataCalculoTarifaCompra)(AddressOf AplicarTarifaPadre, data, services)
            ProcessServer.ExecuteTask(Of DataCalculoTarifaCompra)(AddressOf TarifaPrecioEstandar, data, services)
            ProcessServer.ExecuteTask(Of DataCalculoTarifaCompra)(AddressOf TarifaUltimaFacturaArticuloProveedor, data, services)
            ProcessServer.ExecuteTask(Of DataCalculoTarifaCompra)(AddressOf DescuentosPorFamilia, data, services)
            ProcessServer.ExecuteTask(Of DataCalculoTarifaCompra)(AddressOf PrecioEnMonedaContexto, data, services)
        End If

        Return data.DatosTarifa
    End Function

    <Task()> Public Shared Sub TarifaContratoActivo(ByVal data As DataCalculoTarifaCompra, ByVal services As ServiceProvider)
        Dim contrato As New DataContrato
        If Length(data.IDProveedor) > 0 Then
            contrato.IDArticulo = data.IDArticulo
            contrato.IDProveedor = data.IDProveedor
            contrato.QSobrante = data.Cantidad
            If Not data.Fecha Is Nothing Then contrato.Fecha = data.Fecha
            Dim dtTarifaContrato As DataTable = ProcessServer.ExecuteTask(Of DataContrato, DataTable)(AddressOf ObtenerContratos, contrato, services)
            If Not IsNothing(dtTarifaContrato) AndAlso dtTarifaContrato.Rows.Count Then
                If dtTarifaContrato.Rows(0)("Precio") <> 0 Then
                    data.DatosTarifa.Precio = Nz(dtTarifaContrato.Rows(0)("Precio"), 0)
                    data.DatosTarifa.IDContrato = dtTarifaContrato.Rows(0)("IdContrato")
                    data.DatosTarifa.IDLineaContrato = dtTarifaContrato.Rows(0)("IdLineaContrato")
                    data.DatosTarifa.UDValoracion = Nz(dtTarifaContrato.Rows(0)("UdValoracion"), 1)
                    data.DatosTarifa.IDUDCompra = dtTarifaContrato.Rows(0)("IDUdCompra")
                    data.DatosTarifa.IDMoneda = data.IDMoneda
                    data.DatosTarifa.SeguimientoTarifa = "PRECIO DE CONTRATO ACTIVO"
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub TarifaArticuloProveedor(ByVal data As DataCalculoTarifaCompra, ByVal services As ServiceProvider)
        If data.DatosTarifa.Precio <> 0 Or Length(data.IDProveedor) = 0 Then Exit Sub
        Dim ArtProvs As EntityInfoCache(Of ArticuloProveedorInfo) = services.GetService(Of EntityInfoCache(Of ArticuloProveedorInfo))()
        Dim ArtProv As ArticuloProveedorInfo = ArtProvs.GetEntity(data.IDProveedor, data.IDArticulo)
        If Not ArtProv Is Nothing AndAlso Length(ArtProv.IDArticulo) > 0 AndAlso Length(ArtProv.IDProveedor) > 0 Then
            If Length(data.Fecha) = 0 Then data.Fecha = Today 'cuando la busqueda del precio viene de obras no tenemos fecha pedido, lo hacemos con el día de hoy
            'TODO HISTREV
            '//Si el articulo esta asociado con el proveedor, obtenemos el Precio correspondiente a la cantidad
            Dim FilHistQDesdeProv As New Filter
            FilHistQDesdeProv.Add("FechaDesde", FilterOperator.LessThanOrEqual, data.Fecha)
            FilHistQDesdeProv.Add("FechaHasta", FilterOperator.GreaterThanOrEqual, data.Fecha)
            FilHistQDesdeProv.Add("QDesde", FilterOperator.LessThanOrEqual, data.Cantidad)
            FilHistQDesdeProv.Add("IDProveedor", FilterOperator.Equal, data.IDProveedor)
            FilHistQDesdeProv.Add("IDArticulo", FilterOperator.Equal, data.IDArticulo)
            Dim DtHistQDesdeProv As DataTable = New HistoricoPreciosProveedorQDesde().Filter(FilHistQDesdeProv)
            If Not DtHistQDesdeProv Is Nothing AndAlso DtHistQDesdeProv.Rows.Count > 0 Then
                data.DatosTarifa.Precio = Nz(DtHistQDesdeProv.Rows(0)("Precio"), 0)
                data.DatosTarifa.Dto1 = Nz(DtHistQDesdeProv.Rows(0)("Dto1"), 0)
                data.DatosTarifa.Dto2 = Nz(DtHistQDesdeProv.Rows(0)("Dto2"), 0)
                data.DatosTarifa.Dto3 = Nz(DtHistQDesdeProv.Rows(0)("Dto3"), 0)
            Else
                Dim FilProv As New Filter
                FilProv.Add("IdProveedor", FilterOperator.Equal, data.IDProveedor)
                FilProv.Add("IDArticulo", FilterOperator.Equal, data.IDArticulo)
                FilProv.Add("QDesde", FilterOperator.LessThanOrEqual, data.Cantidad)
                Dim dtAPLinea As DataTable = New ArticuloProveedorLinea().Filter(FilProv, "QDesde DESC")
                If Not dtAPLinea Is Nothing AndAlso dtAPLinea.Rows.Count > 0 Then
                    data.DatosTarifa.Precio = Nz(dtAPLinea.Rows(0)("Precio"), 0)
                    data.DatosTarifa.Dto1 = Nz(dtAPLinea.Rows(0)("Dto1"), 0)
                    data.DatosTarifa.Dto2 = Nz(dtAPLinea.Rows(0)("Dto2"), 0)
                    data.DatosTarifa.Dto3 = Nz(dtAPLinea.Rows(0)("Dto3"), 0)
                Else
                    Dim FilHistProv As New Filter
                    FilHistProv.Add("FechaDesde", FilterOperator.LessThanOrEqual, data.Fecha)
                    FilHistProv.Add("FechaHasta", FilterOperator.GreaterThanOrEqual, data.Fecha)
                    FilHistProv.Add("IDArticulo", FilterOperator.Equal, data.IDArticulo)
                    FilHistProv.Add("IDProveedor", FilterOperator.Equal, data.IDProveedor)
                    Dim DtHistProv As DataTable = New HistoricoPreciosProveedor().Filter(FilHistProv)
                    If Not DtHistProv Is Nothing AndAlso DtHistProv.Rows.Count > 0 Then
                        data.DatosTarifa.Precio = Nz(DtHistProv.Rows(0)("Precio"), 0)
                        data.DatosTarifa.Dto1 = Nz(DtHistProv.Rows(0)("Dto1"), 0)
                        data.DatosTarifa.Dto2 = Nz(DtHistProv.Rows(0)("Dto2"), 0)
                        data.DatosTarifa.Dto3 = Nz(DtHistProv.Rows(0)("Dto3"), 0)
                    Else
                        data.DatosTarifa.Precio = ArtProv.Precio
                        data.DatosTarifa.Dto1 = ArtProv.Dto1
                        data.DatosTarifa.Dto2 = ArtProv.Dto2
                        data.DatosTarifa.Dto3 = ArtProv.Dto3
                    End If
                End If
            End If

            Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
            Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(data.IDProveedor)
            data.DatosTarifa.UDValoracion = ArtProv.UdValoracion
            If (data.DatosTarifa.Precio) <> 0 Then data.DatosTarifa.IDMoneda = ProvInfo.IDMoneda
            data.DatosTarifa.SeguimientoTarifa = "PRECIO OBTENIDO DE LA RELACION ARTÍCULO PROVEEDOR"
            data.DatosTarifa.IDUDCompra = ArtProv.IDUDCompra
        End If
    End Sub

    <Task()> Public Shared Sub AplicarTarifaPadre(ByVal data As DataCalculoTarifaCompra, ByVal services As ServiceProvider)
        If data.DatosTarifa.Precio = 0 Then
            Dim DrArt As DataRow = New Articulo().GetItemRow(data.IDArticulo)
            If Length(DrArt("IDArticuloPadre")) > 0 Then
                data.IDArticulo = DrArt("IDArticuloPadre")
                ProcessServer.ExecuteTask(Of DataCalculoTarifaCompra)(AddressOf TarifaCompra, data, services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub TarifaPrecioEstandar(ByVal data As DataCalculoTarifaCompra, ByVal services As ServiceProvider)
        If data.DatosTarifa.Precio = 0 Then
            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
            Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.IDArticulo)
            data.DatosTarifa.Precio = ArtInfo.PrecioEstandarA
            If data.DatosTarifa.Precio <> 0 Then
                If Len(data.DatosTarifa.IDUDCompra) = 0 Then
                    If Length(data.IDUDMedida) > 0 Then
                        '//Asignamos la medida interna a la medida de compra
                        data.DatosTarifa.IDUDCompra = data.IDUDMedida
                    ElseIf Len(ArtInfo.IDUDInterna) > 0 Then
                        data.DatosTarifa.IDUDCompra = ArtInfo.IDUDInterna
                    End If
                End If
                If data.DatosTarifa.IDUDCompra <> ArtInfo.IDUDInterna Then
                    Dim StDatos As New ArticuloUnidadAB.DatosFactorConversion
                    StDatos.IDArticulo = data.IDArticulo
                    StDatos.IDUdMedidaA = data.DatosTarifa.IDUDCompra
                    StDatos.IDUdMedidaB = ArtInfo.IDUDInterna
                    StDatos.UnoSiNoExiste = False
                    Dim dblFactor As Double = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, StDatos, services)
                    If Nz(dblFactor, 0) <> 0 Then
                        data.DatosTarifa.Precio = data.DatosTarifa.Precio * dblFactor
                    End If
                End If
                Dim IDMonedaA As String = New Parametro().MonedaInternaA
                data.DatosTarifa.IDMoneda = IDMonedaA
                data.DatosTarifa.SeguimientoTarifa = "PRECIO OBTENIDO DEL PRECIO ESTANDAR DEL ARTICULO"
            End If
        End If
    End Sub
    <Task()> Public Shared Sub TarifaUltimaFacturaArticuloProveedor(ByVal data As DataCalculoTarifaCompra, ByVal services As ServiceProvider)
        If data.DatosTarifa.Precio <> 0 Or Length(data.IDProveedor) = 0 Then Exit Sub
        Dim f As New Filter
        f.Add(New StringFilterItem("IDProveedor", data.IDProveedor))
        f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
        Dim dt As DataTable = New BE.DataEngine().Filter("vNegFacturaUltimaCompra", f, , "FechaFactura DESC")
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            If dt.Rows(0)("Precio") > 0 Then data.DatosTarifa.Precio = dt.Rows(0)("Precio")
            If dt.Rows(0)("Dto1") > 0 Then data.DatosTarifa.Dto1 = dt.Rows(0)("Dto1")
            If dt.Rows(0)("Dto2") > 0 Then data.DatosTarifa.Dto2 = dt.Rows(0)("Dto2")
            If dt.Rows(0)("Dto3") > 0 Then data.DatosTarifa.Dto3 = dt.Rows(0)("Dto3")
            If dt.Rows(0)("UdValoracion") > 0 Then data.DatosTarifa.UDValoracion = dt.Rows(0)("UdValoracion")
            If Length(dt.Rows(0)("IDMoneda")) > 0 Then data.DatosTarifa.IDMoneda = dt.Rows(0)("IDMoneda")
            If Length(dt.Rows(0)("IDUdMedida")) > 0 Then data.DatosTarifa.IDUDCompra = dt.Rows(0)("IDUdMedida")
            If Length(dt.Rows(0)("RefProveedor")) > 0 Then data.DatosTarifa.Referencia = dt.Rows(0)("RefProveedor")
            If Length(dt.Rows(0)("DescRefProveedor")) > 0 Then data.DatosTarifa.DescReferencia = dt.Rows(0)("DescRefProveedor")

            If data.DatosTarifa.Precio <> 0 Then
                data.DatosTarifa.SeguimientoTarifa = "PRECIO OBTENIDO DE LA ULTIMA FACTURA"
            End If
        End If
    End Sub

    <Task()> Public Shared Sub DescuentosPorFamilia(ByVal data As DataCalculoTarifaCompra, ByVal services As ServiceProvider)

        If data.DatosTarifa.Dto1 <> 0 Or Length(data.IDProveedor) = 0 Then Exit Sub
        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.IDArticulo)

        If Not ArtInfo Is Nothing Then
            If Len(ArtInfo.IDTipo) > 0 Then
                '//Primero se buscan los Descuentos por Proveedor-Tipo-Familia
                Dim f As New Filter
                f.Add(New StringFilterItem("IDProveedor", data.IDProveedor))
                f.Add(New StringFilterItem("IDTipo", ArtInfo.IDTipo))
                If Len(ArtInfo.IDFamilia) > 0 Then
                    f.Add(New StringFilterItem("IDFamilia", ArtInfo.IDFamilia))
                End If
                Dim dtFamiliaDtosProv As DataTable = New ProveedorDescuentoFamilia().Filter(f)

                '//Si no encuentra Descuentos por Proveedor-Tipo-Familia se buscará Cliente-Tipo
                If dtFamiliaDtosProv.Rows.Count = 0 And Len(ArtInfo.IDFamilia) > 0 Then
                    f.Clear()
                    f.Add(New StringFilterItem("IDProveedor", data.IDProveedor))
                    f.Add(New StringFilterItem("IDTipo", ArtInfo.IDTipo))
                    f.Add(New IsNullFilterItem("IDFamilia"))

                    dtFamiliaDtosProv = New ProveedorDescuentoFamilia().Filter(f)
                End If

                If dtFamiliaDtosProv.Rows.Count > 0 Then
                    data.DatosTarifa.Dto1 = dtFamiliaDtosProv.Rows(0)("Dto1")
                    data.DatosTarifa.Dto2 = dtFamiliaDtosProv.Rows(0)("Dto2")
                    data.DatosTarifa.Dto3 = dtFamiliaDtosProv.Rows(0)("Dto3")

                    data.DatosTarifa.SeguimientoTarifa = data.DatosTarifa.SeguimientoTarifa & " DESCUENTOS POR FAMILIA"
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub PrecioEnMonedaContexto(ByVal data As DataCalculoTarifaCompra, ByVal services As ServiceProvider)
        If Length(data.DatosTarifa.IDMoneda) > 0 AndAlso data.IDMoneda <> data.DatosTarifa.IDMoneda Then
            If data.Fecha Is Nothing Then data.Fecha = Today
            Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
            Dim MonTarifa As MonedaInfo = Monedas.GetMoneda(data.DatosTarifa.IDMoneda, data.Fecha)
            Dim MonContexto As MonedaInfo = Monedas.GetMoneda(data.IDMoneda, data.Fecha)

            If MonContexto.CambioA <> 0 Then
                data.DatosTarifa.Precio = data.DatosTarifa.Precio * (MonTarifa.CambioA / MonContexto.CambioA)
            End If
        End If
    End Sub
    <Task()> Public Shared Sub AplicarTarifaCompra(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Not data.Context Is Nothing AndAlso Length(data.Context("IDProveedor")) > 0 Then

            Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
            Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(data.Context("IDProveedor"))
            If (data.Current.ContainsKey("PedidoVentaOrigen") AndAlso Length(data.Current("PedidoVentaOrigen")) > 0) And ProvInfo.EmpresaGrupo Then Exit Sub
            If (Length(data.Current("IDLineaMaterial")) = 0 And Length(data.Current("IDOferta")) = 0 And (Length(data.Current("IdLineaSolicitud")) = 0 Or (data.Context.Contains("Origen") AndAlso data.Context("Origen") = "SolicitudCompraLinea")) And Length(data.Current("IDLineaPrograma")) = 0) Then
                If data.Context.ContainsKey("IDOrdenRuta") AndAlso Length(data.Context("IDOrdenRuta")) > 0 Then
                    '//No hay que recuperar nada si se trata de una Subcontratación.
                ElseIf data.Current.ContainsKey("IDOrdenRuta") AndAlso Length(data.Current("IDOrdenRuta")) > 0 Then
                    '//No hay que recuperar nada si se trata de una Subcontratación.
                ElseIf data.Current.ContainsKey("IDArticulo") AndAlso data.Context.ContainsKey("IDProveedor") AndAlso data.Current.ContainsKey("Cantidad") Then
                    Dim dataTarifa As New DataCalculoTarifaCompra
                    dataTarifa.IDArticulo = data.Current("IDArticulo")
                    dataTarifa.IDProveedor = data.Context("IDProveedor")
                    dataTarifa.Cantidad = Nz(data.Current("Cantidad"), 0)
                    dataTarifa.Fecha = data.Context("Fecha")
                    dataTarifa.IDUDMedida = Nz(data.Current("IDUDMedida"), String.Empty)
                    If Length(data.Current("UDValoracion")) > 0 Then dataTarifa.UDValoracion = CInt(data.Current("UDValoracion"))
                    ProcessServer.ExecuteTask(Of DataCalculoTarifaCompra, DataTarifaCompra)(AddressOf ProcesoCompra.TarifaCompra, dataTarifa, services)
                    If Not dataTarifa.DatosTarifa Is Nothing AndAlso dataTarifa.DatosTarifa.Precio <> 0 Then
                        'data.Current("Precio") = dataTarifa.DatosTarifa.Precio

                        If data.Context.Contains("IDMoneda") AndAlso (data.Context("IDMoneda") <> dataTarifa.DatosTarifa.IDMoneda) Then
                            '//Se le había dato precio en otra moneda diferente a la del Proveedor,
                            '//Por lo que tenemos que hacer la conversión.
                            Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
                            Dim m As MonedaInfo = Monedas.GetMoneda(data.Context("IDMoneda"), data.Context("Fecha"))
                            data.Current("Precio") = xRound(dataTarifa.DatosTarifa.Precio / m.CambioA, m.NDecimalesPrecio)
                        Else
                            data.Current("Precio") = dataTarifa.DatosTarifa.Precio
                        End If

                        If Length(dataTarifa.DatosTarifa.IDContrato) > 0 AndAlso Length(dataTarifa.DatosTarifa.IDLineaContrato) > 0 Then
                            data.Current("IDContrato") = dataTarifa.DatosTarifa.IDContrato
                            data.Current("IDLineaContrato") = dataTarifa.DatosTarifa.IDLineaContrato
                        End If
                        data.Current("IDMoneda") = dataTarifa.DatosTarifa.IDMoneda
                        data.Current("Dto1") = dataTarifa.DatosTarifa.Dto1
                        data.Current("Dto2") = dataTarifa.DatosTarifa.Dto2
                        data.Current("Dto3") = dataTarifa.DatosTarifa.Dto3
                        If dataTarifa.DatosTarifa.UDValoracion <> 0 Then data.Current("UDValoracion") = dataTarifa.DatosTarifa.UDValoracion
                        'data.Current("IDUDMedida") = dataTarifa.DatosTarifa.IDUDCompra
                        data.Current("SeguimientoTarifa") = dataTarifa.DatosTarifa.SeguimientoTarifa

                        If data.Current("IDUDInterna") <> dataTarifa.DatosTarifa.IDUDCompra AndAlso data.Current("IDUDMedida") <> dataTarifa.DatosTarifa.IDUDCompra Then
                            data.Current("IDUDMedida") = dataTarifa.DatosTarifa.IDUDCompra
                            Dim StDatos As New ArticuloUnidadAB.DatosFactorConversion
                            StDatos.IDArticulo = data.Current("IDArticulo")
                            StDatos.IDUdMedidaA = data.Current("IDUDMedida")
                            StDatos.IDUdMedidaB = data.Current("IDUDInterna")
                            StDatos.UnoSiNoExiste = False
                            Dim dblFactor As Double = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, StDatos, services)
                            If Nz(dblFactor, 0) <> 0 Then
                                data.Current("Factor") = dblFactor
                            Else
                                If Nz(data.Current("Cantidad"), 0) <> 0 Then data.Current("Factor") = data.Current("QInterna") / data.Current("Cantidad")
                            End If
                        End If
                        data.Current("SeguimientoTarifa") = dataTarifa.DatosTarifa.SeguimientoTarifa
                    End If


                    If data.Context.ContainsKey("IDMoneda") And data.Current.ContainsKey("IDMoneda") Then
                        'TODO: Está incluido en obtención de la Tarifa ?¿
                        ''//Tendremos un current("IDMoneda") cuando cambiamos la tarifa
                        'If Length(data.Current("IDMoneda")) > 0 Then
                        '    If data.Context("IDMoneda") & String.Empty <> data.Current("IDMoneda") & String.Empty Then
                        '        '//CAMBIO DE MONEDA (aplicamos el cambio de la tarifa)
                        '        data.Current = CambioMoneda(data.Current, data.Current("IDMoneda"), data.Context("IDMoneda"), Nz(data.Context("Fecha"), cnMinDate), services.GetService(Of MonedaCache))
                        '    End If
                        'End If

                        '//IMPORTE (aplicamos el cambio de la cabecera)
                        If data.Context.ContainsKey("CambioA") AndAlso data.Context.ContainsKey("CambioB") Then
                            ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf NegocioGeneral.CalcularPrecioImporte, data.Current, services)
                            Dim ValAyB As New ValoresAyB(data.Current, data.Context("IDMoneda"), data.Context("CambioA"), data.Context("CambioB"))
                            ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)
                        Else
                            ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf NegocioGeneral.CalcularPrecioImporte, data.Current, services)
                            Dim ValAyB As New ValoresAyB(data.Current, data.Context("IDMoneda"), 0, 0)
                            ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)
                        End If
                        data.Current("IDMoneda") = data.Context("IDMoneda")
                    End If
                End If
            End If
        End If

    End Sub

#End Region

#Region " Métodos comunes a los procesos de creación de documentos del circuito compra "
    <Task()> Public Shared Sub AsignarDatosProveedor(ByVal Doc As DocumentoCompra, ByVal services As ServiceProvider)
        If Doc.Proveedor Is Nothing Then
            Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
            Doc.Proveedor = Proveedores.GetEntity(Doc.HeaderRow("IDProveedor"))
        End If
        If Doc.HeaderRow.IsNull("IDFormaPago") Then Doc.HeaderRow("IDFormaPago") = Doc.Proveedor.IDFormaPago
        If Doc.HeaderRow.IsNull("IDCondicionPago") Then Doc.HeaderRow("IDCondicionPago") = Doc.Proveedor.IDCondicionPago
        'If Doc.HeaderRow.IsNull("IDDiaPago") Then Doc.HeaderRow("IDDiaPago") = Doc.Proveedor.IDDiaPago
        If Doc.HeaderRow.IsNull("IDMoneda") Then Doc.HeaderRow("IDMoneda") = Doc.Proveedor.IDMoneda

    End Sub

    <Task()> Public Shared Sub AsignarObservacionesCompra(ByVal Doc As DocumentoCompra, ByVal services As ServiceProvider)
        Dim obs As New DataObservaciones(Doc.EntidadCabecera, "Texto", New DataRowPropertyAccessor(Doc.HeaderRow))
        ProcessServer.ExecuteTask(Of DataObservaciones)(AddressOf ProcesoCompra.AsignarObservacionesProveedor, obs, services)
        ProcessServer.ExecuteTask(Of DocumentoCompra)(AddressOf AsignarObservacionesAlbaran, Doc, services)
    End Sub

    <Task()> Public Shared Sub AsignarObservacionesAlbaran(ByVal Doc As DocumentoCompra, ByVal services As ServiceProvider)
        If Not Doc.Cabecera Is Nothing AndAlso TypeOf Doc.Cabecera Is FraCabCompraAlbaran Then
            Doc.HeaderRow("Texto") &= CType(Doc.Cabecera, FraCabCompraAlbaran).Texto
        End If
    End Sub

    <Task()> Public Shared Sub AsignarEjercicio(ByVal Doc As DocumentoCompra, ByVal services As ServiceProvider)
        Dim AppParamsConta As ParametroContabilidadCompra = services.GetService(Of ParametroContabilidadCompra)()
        If Not AppParamsConta.Contabilidad Then Exit Sub
        Dim DE As New DataEjercicio(New DataRowPropertyAccessor(Doc.HeaderRow), Doc.Fecha)
        ProcessServer.ExecuteTask(Of DataEjercicio)(AddressOf NegocioGeneral.AsignarEjercicioContable, DE, services)
    End Sub
    <Task()> Public Shared Sub AsignarTipoCompra(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim AppParamsCompra As ParametroCompra = services.GetService(Of ParametroCompra)()
        If data.IsNull("IDTipoCompra") Then data("IDTipoCompra") = AppParamsCompra.TipoCompraNormal
    End Sub

#End Region

End Class


