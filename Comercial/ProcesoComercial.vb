Public Class ProcesoComercial

#Region " DetailCommonUpdateRules "

    <Task()> Public Shared Sub DetailCommonUpdateRules(ByVal dr As DataRow, ByVal services As ServiceProvider)
        If Length(dr("IdArticulo")) = 0 Then
            ApplicationService.GenerateError("El artículo es obligatorio.")
        ElseIf Length(dr("DescArticulo")) = 0 Then
            ApplicationService.GenerateError("La descripción es obligatoria.")
        ElseIf Length(dr("IdTipoIva")) = 0 Then
            ApplicationService.GenerateError("El tipo de IVA es obligatorio.")
        ElseIf Length(dr("CContable")) = 0 Then
            Dim AppParamsConta As ParametroContabilidadVenta = services.GetService(Of ParametroContabilidadVenta)()
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
        ElseIf Length(dr("IDTipoLinea")) = 0 Then
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

#Region " DetailCommonBusinessRules - CABECERAS "

    <Task()> Public Shared Sub CambioDireccion(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        '//CUIDADO esto sólo desde PV y AV.
        data.Current(data.ColumnName) = data.Value

        If data.Context.ContainsKey("Lineas") Then
            Dim BH As BusinessHelper
            If data.Current.ContainsKey("IDDireccionEnvio") Then
                BH = BusinessHelper.CreateBusinessObject("PedidoVentaLinea")
            ElseIf data.Current.ContainsKey("IDDireccion") Then
                BH = BusinessHelper.CreateBusinessObject("AlbaranVentaLinea")
            End If
            Dim dtLineas As DataTable = data.Context("Lineas")
            If Not dtLineas Is Nothing AndAlso dtLineas.Rows.Count > 0 Then
                Dim contextLin As BusinessData = data.Current
                For Each linea As DataRow In dtLineas.Rows
                    If linea.RowState <> DataRowState.Deleted Then
                        Dim currentLin As New DataRowPropertyAccessor(linea)
                        Dim bRules As New BusinessRuleData(String.Empty, Nothing, currentLin, contextLin)

                        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoComercial.ObtenerIvaArticuloCliente, bRules, services)
                    End If
                Next
                '//se devuelve en el current, en lugar de en el context para poder tener acceso a él desde presentación
                data.Current("Lineas") = dtLineas
            End If
        End If

    End Sub

    <Task()> Public Shared Function DetailBusinessRulesCab(ByVal oBRL As BusinessRules, ByVal services As ServiceProvider) As BusinessRules
        If oBRL Is Nothing Then oBRL = New BusinessRules
        oBRL.Add("IDCliente", AddressOf ProcesoComercial.CambioCliente)
        oBRL.Add("IDCondicionPago", AddressOf NegocioGeneral.CambioCondicionPago)
        oBRL.Add("IDMoneda", AddressOf ProcesoComunes.CambioMoneda)
        Return oBRL
    End Function

    <Task()> Public Shared Sub CambioCliente(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If data.ColumnName = "IDCliente" Then data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDCliente")) Then
            Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
            Dim ClteInfo As ClienteInfo = Clientes.GetEntity(data.Current("IDCliente"))
            data.Current("CifCliente") = ClteInfo.CifCliente
            data.Current("RazonSocial") = ClteInfo.RazonSocial
            data.Current("Direccion") = ClteInfo.Direccion
            data.Current("CodPostal") = ClteInfo.CodPostal
            data.Current("Poblacion") = ClteInfo.Poblacion
            data.Current("Provincia") = ClteInfo.Provincia
            data.Current("IdMoneda") = ClteInfo.Moneda

            data.Current("IDFormaPago") = ClteInfo.FormaPago
            data.Current("IdCondicionPago") = ClteInfo.CondicionPago
            ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf NegocioGeneral.CambioCondicionPago, data, services)
            data.Current("IdFormaEnvio") = ClteInfo.FormaEnvio
            data.Current("IDCondicionEnvio") = ClteInfo.CondicionEnvio
            data.Current("IDDiaPago") = ClteInfo.DiaPago
        Else
            data.Current("CifCliente") = System.DBNull.Value
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
            data.Current("CambioA") = 0
            data.Current("CambioB") = 0

            'data.Current("IDPais") = System.DBNull.Value
            'data.Current("Telefono") = System.DBNull.Value
            'data.Current("Fax") = System.DBNull.Value
            'data.Current("IdClienteBanco") = System.DBNull.Value
            'data.Current("IdBancoPropio") = System.DBNull.Value
            'data.Current("DtoFactura") = 0
            'data.Current("RetencionIRPF") = 0
            'data.Current("Texto") = System.DBNull.Value
        End If
        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf AsignarClienteBanco, data.Current, services)
    End Sub

    <Task()> Public Shared Sub AsignarClienteBanco(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider)
        If Length(data("IDCliente")) > 0 Then
            Dim IDBanco As Integer = ProcessServer.ExecuteTask(Of String, Integer)(AddressOf ClienteBanco.GetBancoPredeterminado, data("IDCliente"), services)
            If IDBanco > 0 Then
                data("IDClienteBanco") = IDBanco
            Else
                data("IDClienteBanco") = System.DBNull.Value
            End If
        Else
            data("IDClienteBanco") = System.DBNull.Value
        End If
    End Sub

    <Task()> Public Shared Sub AsignarDireccionCliente(ByVal data As DataDireccionClte, ByVal services As ServiceProvider)
        If Length(data.Datos("IDCliente")) = 0 Then Exit Sub
        Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
        Dim ClteInfo As ClienteInfo = Clientes.GetEntity(data.Datos("IDCliente"))

        Dim strCliente As String = ClteInfo.IDCliente
        If ClteInfo.GrupoDireccion Then strCliente = ClteInfo.GrupoCliente
        Dim StDatosDirec As New ClienteDireccion.DataDirecEnvio(strCliente, data.TipoDireccion)
        Dim dtDireccion As DataTable = ProcessServer.ExecuteTask(Of ClienteDireccion.DataDirecEnvio, DataTable)(AddressOf ClienteDireccion.ObtenerDireccionEnvio, StDatosDirec, services)
        If Not dtDireccion Is Nothing AndAlso dtDireccion.Rows.Count > 0 Then
            data.Datos(data.Field) = dtDireccion.Rows(0)("IDDireccion")
            data.Datos("IDOficinaContable") = Nz(dtDireccion.Rows(0)("IDOficinaContable"), String.Empty)
            data.Datos("IDOrganoGestor") = Nz(dtDireccion.Rows(0)("IDOrganoGestor"), String.Empty)
            data.Datos("IDUnidadTramitadora") = Nz(dtDireccion.Rows(0)("IDUnidadTramitadora"), String.Empty)
        End If
    End Sub

    <Task()> Public Shared Sub AsignarObservacionesCliente(ByVal data As DataObservaciones, ByVal services As ServiceProvider)
        If Length(data.Datos("IDCliente")) > 0 Then
            Dim StDatos As New Observacion.DatosObv
            StDatos.IDEntidad = data.Entity
            StDatos.IDPrimaryKey = data.Datos("IDCliente")
            Dim Observaciones As String = ProcessServer.ExecuteTask(Of Observacion.DatosObv, String)(AddressOf Observacion.ObtenerObservacionesCliente, StDatos, services)
            If Length(data.Datos(data.Field)) > 0 Then
                data.Datos(data.Field) = data.Datos(data.Field) & vbNewLine & Observaciones
            Else
                data.Datos(data.Field) = Observaciones
            End If
        End If
    End Sub


#End Region

#Region " DetailCommonBusinessRules - LINEAS "
    <Task()> Public Shared Function DetailBusinessRulesLin(ByVal oBRL As BusinessRules, ByVal services As ServiceProvider) As BusinessRules
        oBRL.Add("IDArticulo", AddressOf ProcesoComercial.CambioArticulo)
        oBRL.Add("CodigoBarras", AddressOf ProcesoComercial.CambioArticulo)
        oBRL.Add("RefCliente", AddressOf ProcesoComercial.CambioArticulo)
        oBRL.Add("Cantidad", AddressOf ProcesoComercial.CambioCantidad)
        oBRL.Add("QInterna", AddressOf ProcesoComercial.CambioQInterna)
        oBRL.Add("QInterna2", AddressOf CambioCantidadInterna2)
        oBRL.Add("Factor", AddressOf NegocioGeneral.CambioFactor)
        oBRL.Add("IDUDMedida", AddressOf NegocioGeneral.CambioUDMedida)
        oBRL.Add("PVP", AddressOf NegocioGeneral.CambioPrecio)
        oBRL.Add("Precio", AddressOf NegocioGeneral.CambioPrecio)
        oBRL.Add("UDValoracion", AddressOf NegocioGeneral.CambioPrecio)
        oBRL.Add("Dto1", AddressOf NegocioGeneral.CambioPrecio)
        oBRL.Add("Dto2", AddressOf NegocioGeneral.CambioPrecio)
        oBRL.Add("Dto3", AddressOf NegocioGeneral.CambioPrecio)
        oBRL.Add("Dto", AddressOf NegocioGeneral.CambioPrecio)
        oBRL.Add("DtoProntoPago", AddressOf NegocioGeneral.CambioPrecio)
        oBRL.Add("PrecioA", AddressOf NegocioGeneral.CalcularImportes)
        oBRL.Add("PrecioB", AddressOf NegocioGeneral.CalcularImportes)
        oBRL.Add("CContable", AddressOf NegocioGeneral.CambioCContable)
        oBRL.Add("NObra", AddressOf ProcesoComercial.CambioNObra)
        oBRL.Add("CodTrabajo", AddressOf ProcesoComercial.CambioCodTrabajo)
        oBRL.Add("Regalo", AddressOf ProcesoComercial.CambioRegalo)
        oBRL.Add("IDTarifa", AddressOf ProcesoComercial.CambioTarifa)
        oBRL.Add("IDTipoLinea", AddressOf ProcesoComercial.CambioTipoLinea)

        'oBRL.Add("IDLineaAlbaran", AddressOf ProcesoComercial.CambioLineaAlbaran)

        Return oBRL
    End Function

#Region " CambioArticulo "

    <Task()> Public Shared Sub CambioArticulo(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ValidarBloqueoArticulo, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ValidarClienteCabecera, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf RecuperarInformacionArticulo, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf AsignarValoresPredeterminados, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf AsignarDatosArticulo, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf TratarCContable, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ObtenerIvaArticuloCliente, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ArticuloCliente, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf NegocioGeneral.FactorConversion, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf Tarifa, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf GetStockFisico, data, services)
    End Sub

    <Task()> Public Shared Sub ValidarClienteCabecera(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If data.Context.ContainsKey("IDCliente") AndAlso Length(data.Context("IDCliente")) = 0 Then
            ApplicationService.GenerateError("No ha especificado el Cliente.")
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
                Case "CodigoBarras"
                    Dim stArticulo As New Articulo.DataInfoArticulo(Nothing, data.Current("CodigoBarras"))
                    ArtInfo = ProcessServer.ExecuteTask(Of Articulo.DataInfoArticulo, ArticuloInfo)(AddressOf Articulo.InformacionArticulo, stArticulo, services)
                    data.Current("IDArticulo") = ArtInfo.IDArticulo
                Case "RefCliente"
                    Dim stArticulo As New Articulo.DataInfoArticulo(Nothing, Nothing, data.Current("RefCliente"), data.Context)
                    ArtInfo = ProcessServer.ExecuteTask(Of Articulo.DataInfoArticulo, ArticuloInfo)(AddressOf Articulo.InformacionArticulo, stArticulo, services)
                    data.Current("IDArticulo") = ArtInfo.IDArticulo
            End Select

            If Not ArtInfo Is Nothing AndAlso Length(ArtInfo.IDArticulo) > 0 Then
                '//Registramos el artículo de manera manual, en lugar de hacerlo a través del services, por que dependiendo de como
                '// se intente acceder recuperaremos el artículo de una manera u otra.
                Dim Articulos As New EntityInfoCache(Of ArticuloInfo)
                Articulos.Add(data.Current("IDArticulo")) = ArtInfo
                services.RegisterService(Articulos, GetType(EntityInfoCache(Of ArticuloInfo)))
            End If
        ElseIf Length(data.Current(data.ColumnName)) = 0 AndAlso data.ColumnName = "RefCliente" Then
            If data.Context.ContainsKey("IDArticuloRef") Then
                data.Current("IDArticulo") = data.Context("IDArticuloRef")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarValoresPredeterminados(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current("IDLineaOfertaDetalle") = DBNull.Value
        data.Current("NumOferta") = DBNull.Value
        If Nz(data.Current("UDValoracion"), 0) = 0 Then data.Current("UDValoracion") = 1
        data.Current("IDTipoLinea") = ProcessServer.ExecuteTask(Of Object, String)(AddressOf TipoLinea.TipoLineaPorDefecto, Nothing, services)

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
        data.Current("IDConcepto") = ArtInfo.IDConcepto

        If Not ArtInfo.Activo Then
            ApplicationService.GenerateError("El artículo | no está activo.", Quoted(data.Current("IDArticulo")))
        ElseIf Not ArtInfo.Venta Then
            ApplicationService.GenerateError("El artículo | no es de tipo venta.", Quoted(data.Current("IDArticulo")))
        Else
            data.Current("Configurable") = ArtInfo.Configurable
            If Length(ArtInfo.IDUDVenta) = 0 Then
                ApplicationService.GenerateError("Debe indicar una unidad de venta para el artículo {0}.", Quoted(data.Current("IDArticulo")))
            Else
                data.Current("IDUDMedida") = ArtInfo.IDUDVenta
            End If
            data.Current("IDUDInterna") = ArtInfo.IDUDInterna
            If data.Current.ContainsKey("IDUDInterna2") AndAlso ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.Current("IDArticulo"), services) Then
                data.Current("IDUDInterna2") = ArtInfo.IDUDInterna2
            End If

            data.Current("UdValoracion") = ArtInfo.UDValoracion
            data.Current("IDTipoIva") = ArtInfo.IDTipoIVA
            data.Current("CodigoBarras") = ArtInfo.CodigoBarras
            If ArtInfo.Especial Then
                If data.Current.ContainsKey("Especial") Then
                    data.Current("Especial") = ArtInfo.Especial
                End If
                data.Current("Dto") = 0
                data.Current("DtoProntoPago") = 0
            End If
        End If

        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf NegocioGeneral.AsignarArticuloAlmacen, data, services)
    End Sub

#Region " TratarCContable "

    <Serializable()> _
    Public Class DataGetCContable
        Public Datos As IPropertyAccessor
        Public TipoCuenta As TipoCContable
    End Class

    Public Enum TipoCContable
        CCExport
        CCVenta
        CCExportGRUPO
        CCVentaGRUPO
    End Enum

    <Task()> Public Shared Sub TratarCContable(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim AppParams As ParametroContabilidadVenta = services.GetService(Of ParametroContabilidadVenta)()
        If Not AppParams.Contabilidad Then Exit Sub

        Dim ClteInfo As New ClienteInfo
        If data.Context.ContainsKey("IDCliente") AndAlso Length(data.Context("IDCliente")) > 0 Then
            Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
            ClteInfo = Clientes.GetEntity(data.Context("IDCliente"))
        End If

        Dim Cta As New DataGetCContable
        If ClteInfo.EmpresaGrupo Then
            If Not ClteInfo.Extranjero Then
                Cta.Datos = data.Current
                Cta.TipoCuenta = TipoCContable.CCVentaGRUPO
            Else
                Cta.Datos = data.Current
                Cta.TipoCuenta = TipoCContable.CCExportGRUPO
            End If
        ElseIf Not ClteInfo.Extranjero Then
            Cta.Datos = data.Current
            Cta.TipoCuenta = TipoCContable.CCVenta
        Else
            Cta.Datos = data.Current
            Cta.TipoCuenta = TipoCContable.CCExport
        End If
        ProcessServer.ExecuteTask(Of DataGetCContable)(AddressOf GetCContableArticulo, Cta, services)

        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf NegocioGeneral.FormatoCuentaContable, data, services)
    End Sub

    <Task()> Public Shared Sub GetCContableArticulo(ByVal data As DataGetCContable, ByVal services As ServiceProvider)
        Dim AppParamsConta As ParametroContabilidadVenta = services.GetService(Of ParametroContabilidadVenta)()
        If Not AppParamsConta.Contabilidad Then Exit Sub
        If Length(data.Datos("IDArticulo")) = 0 Then Exit Sub
        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.Datos("IDArticulo"))
        Dim strCContable As String : Dim strField As String

        Select Case data.TipoCuenta
            Case TipoCContable.CCExport
                strField = "CCExport"
                strCContable = ArtInfo.CCExport
            Case TipoCContable.CCExportGRUPO
                strField = "CCExportGRUPO"
                strCContable = ArtInfo.CCExportGrupo
            Case TipoCContable.CCVenta
                strField = "CCVenta"
                strCContable = ArtInfo.CCVenta
            Case TipoCContable.CCVentaGRUPO
                strField = "CCVentaGRUPO"
                strCContable = ArtInfo.CCVentaGrupo
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
                    Case TipoCContable.CCExport
                        data.Datos("CContable") = AppParamsConta.CuentaExportacion
                    Case TipoCContable.CCExportGRUPO
                        data.Datos("CContable") = AppParamsConta.CuentaExportacionGrupo
                    Case TipoCContable.CCVenta
                        data.Datos("CContable") = AppParamsConta.CuentaVenta
                    Case TipoCContable.CCVentaGRUPO
                        data.Datos("CContable") = AppParamsConta.CuentaVentaGrupo
                End Select
            End If
        End If
    End Sub

#End Region

    <Task()> Public Shared Sub ArticuloCliente(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim ArtCltes As EntityInfoCache(Of ArticuloClienteInfo) = services.GetService(Of EntityInfoCache(Of ArticuloClienteInfo))()
        Dim ArtClte As ArticuloClienteInfo = ArtCltes.GetEntity(data.Context("IDCliente"), data.Current("IDArticulo"))
        If Not IsNothing(ArtClte) AndAlso Length(ArtClte.IDCliente) > 0 AndAlso Length(ArtClte.IDArticulo) > 0 Then
            data.Current("RefCliente") = ArtClte.RefCliente
            data.Current("DescRefCliente") = ArtClte.DescRefCliente
            data.Current("IDUDMedida") = ArtClte.IDUDVenta
            data.Current("IDUDExpedicion") = ArtClte.IDUDExpedicion
            data.Current("UdValoracion") = ArtClte.UdValoracion
            data.Current("Revision") = ArtClte.Revision
            If ArtClte.IDLineaOfertaDetalle > 0 Then data.Current("IdLineaOfertaDetalle") = ArtClte.IDLineaOfertaDetalle
        Else
            data.Current("RefCliente") = String.Empty
            data.Current("DescRefCliente") = String.Empty
            If Length(data.Context("IDCliente")) > 0 Then
                If Length(data.Current("IDUDExpedicion")) = 0 Then
                    Dim DrClie As DataRow = New Cliente().GetItemRow(data.Context("IDCliente"))
                    data.Current("IDUDExpedicion") = Nz(DrClie("IDUDExpedicion"), String.Empty)
                End If
            End If
        End If
        If Length(data.Current("DescRefCliente")) = 0 Then
            Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
            Dim ClteInfo As ClienteInfo = Clientes.GetEntity(data.Context("IDCliente"))
            If Length(ClteInfo.GrupoCliente) > 0 AndAlso ClteInfo.GrupoArticulo Then
                ArtClte = ArtCltes.GetEntity(ClteInfo.GrupoCliente, data.Current("IDArticulo"))
                If Not IsNothing(ArtClte) AndAlso Length(ArtClte.IDCliente) > 0 AndAlso Length(ArtClte.IDArticulo) > 0 Then
                    data.Current("RefCliente") = ArtClte.RefCliente
                    If Length(ArtClte.DescRefCliente) > 0 Then
                        data.Current("DescRefCliente") = ArtClte.DescRefCliente
                    End If
                    If Length(ArtClte.IDUDVenta) > 0 Then
                        data.Current("IDUDMedida") = ArtClte.IDUDVenta
                    End If
                    data.Current("UdValoracion") = ArtClte.UdValoracion
                    data.Current("Revision") = ArtClte.Revision
                    If ArtClte.IDLineaOfertaDetalle > 0 Then data.Current("IdLineaOfertaDetalle") = ArtClte.IDLineaOfertaDetalle
                End If
            End If

            If Length(data.Current("DescRefCliente")) = 0 Then
                If Length(ClteInfo.Idioma) > 0 Then
                    Dim dtIdioma As DataTable = New ArticuloIdioma().SelOnPrimaryKey(data.Current("IDArticulo"), ClteInfo.Idioma)
                    If Not IsNothing(dtIdioma) AndAlso dtIdioma.Rows.Count > 0 Then
                        data.Current("DescRefCliente") = dtIdioma.Rows(0)("DescArticuloIdioma")
                    End If
                End If
            End If
        End If
    End Sub

#End Region

#Region " CambioCantidad "

    <Task()> Public Shared Sub CambioCantidadInterna2(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If IsNumeric(data.Current("Qinterna2")) Then
            ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoComunes.CambioSegundaUnidad, data, services)
            If data.Current.ContainsKey("CambioQInterna") AndAlso data.Current("CambioQInterna") Then ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf Tarifa, data, services)
        Else
            ApplicationService.GenerateError("Campo no numérico.")
        End If
    End Sub
    <Task()> Public Shared Sub CambioCantidad(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current("CantidadAnterior") = data.Current("Cantidad") '//Necesario para la actualización de las promociones
        data.Current(data.ColumnName) = data.Value
        If IsNumeric(data.Current("Cantidad")) Then
            ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf NegocioGeneral.CalculoQInterna, data, services)
            ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf Tarifa, data, services)
        Else
            ApplicationService.GenerateError("Campo no numérico.")
        End If
    End Sub

    <Task()> Public Shared Sub GetStockFisico(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If data.Current.ContainsKey("StockFisico") Then
            If Length(data.Current("IDArticulo")) > 0 AndAlso Length(data.Current("IDAlmacen")) > 0 AndAlso Length(data.Current("StockFisico")) = 0 Then
                Dim DtArtAlm As DataTable = New ArticuloAlmacen().SelOnPrimaryKey(data.Current("IDArticulo"), data.Current("IDAlmacen"))
                If Not DtArtAlm Is Nothing AndAlso DtArtAlm.Rows.Count > 0 Then
                    data.Current("StockFisico") = Nz(DtArtAlm.Rows(0)("StockFisico"), 0)
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub Tarifa(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim AppParamsVenta As ParametroVenta = services.GetService(Of ParametroVenta)()

        If AppParamsVenta.General.AplicacionGestionAlquiler AndAlso Length(data.Current("IDObra")) > 0 Then
            ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf TarifaAlquiler, data, services)
        Else
            ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf AplicarTarifaComercial, data, services)
        End If
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CambioMoneda, data, services)
    End Sub

    <Task()> Public Shared Sub CambioMoneda(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If data.Context.ContainsKey("IDMoneda") Then
            '//Tendremos un current("IDMoneda") cuando cambiamos la tarifa
            If Length(data.Current("IDMoneda")) > 0 Then
                If data.Context("IDMoneda") & String.Empty <> data.Current("IDMoneda") & String.Empty Then
                    '//CAMBIO DE MONEDA (aplicamos el cambio de la tarifa)
                    Dim datos As New DataCambioMoneda(data.Current, data.Current("IDMoneda"), data.Context("IDMoneda"), Nz(data.Context("Fecha"), cnMinDate))
                    ProcessServer.ExecuteTask(Of DataCambioMoneda)(AddressOf NegocioGeneral.CambioMoneda, datos, services)
                End If
            End If

            '//IMPORTE (aplicamos el cambio de la cabecera)
            ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf NegocioGeneral.CalcularImportes, data, services)
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
            ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoComercial.CambioCantidad, data, services)
        End If
    End Sub

#End Region

    <Task()> Public Shared Sub CambioNObra(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current("IDTrabajo") = DBNull.Value
        data.Current("CodTrabajo") = DBNull.Value
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDObra")) Then
            Dim obra As BusinessHelper = BusinessHelper.CreateBusinessObject(AdminData.GetEntityInfo("ObraCabecera"))
            Dim dtObra As DataTable = obra.SelOnPrimaryKey(data.Current("IDObra"))
            If IsNothing(dtObra) OrElse dtObra.Rows.Count = 0 Then
                ApplicationService.GenerateError("El Proyecto | no existe.", Quoted(data.Value))
            Else
                '///TARIFA
                ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf Tarifa, data, services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioCodTrabajo(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Value) = 0 Then
            data.Current("IDTrabajo") = DBNull.Value
        End If
    End Sub

    <Task()> Public Shared Sub CambioRegalo(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Nz(data.Current("Regalo"), False) Then
            data.Current("Importe") = 0
            data.Current("ImporteA") = 0
            data.Current("ImporteB") = 0
        Else
            '///TARIFA
            ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf Tarifa, data, services)
        End If
    End Sub

    <Task()> Public Shared Sub CambioTarifa(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf Tarifa, data, services)
    End Sub

    <Task()> Public Shared Sub CambioTipoLinea(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim dtTipoLinea As DataTable = New TipoLinea().SelOnPrimaryKey(data.Value)
        If dtTipoLinea Is Nothing OrElse dtTipoLinea.Rows.Count = 0 Then
            ApplicationService.GenerateError("El Tipo de Línea {0} no existe.", Quoted(data.Value))
        Else
            data.Current("Regalo") = dtTipoLinea.Rows(0)("Regalo")
            ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoComercial.CambioRegalo, data, services)
        End If
    End Sub

    <Task()> Public Shared Sub ValidarBloqueoArticulo(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            If Not data.Context Is Nothing AndAlso data.Context.Contains("ValidarBloqueoArticulo") AndAlso data.Context("ValidarBloqueoArticulo") Then
                Dim infoBloqueo As New Cliente.DataBloqArtClie
                infoBloqueo.IDCliente = data.Context("IDCliente")
                If data.ColumnName = "IDArticulo" Then
                    infoBloqueo.IDArticulo = data.Value
                ElseIf data.ColumnName = "RefCliente" Then
                    infoBloqueo.RefCliente = data.Value
                End If
                If ProcessServer.ExecuteTask(Of Cliente.DataBloqArtClie, Boolean)(AddressOf Cliente.ComprobarBloqueoArticuloCliente, infoBloqueo, services) Then
                    ApplicationService.GenerateError("El Artículo está bloqueado para este Cliente.")
                End If
            End If
        End If
    End Sub

#End Region

#Region " Obtener IVA "

    <Task()> Public Shared Sub ObtenerIvaArticuloCliente(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Context("IDCliente")) > 0 AndAlso Length(data.Current("IDArticulo")) > 0 Then
            Dim datos As New DataIvaArticuloCliente
            datos.IDCliente = data.Context("IDCliente")
            datos.IDArticulo = data.Current("IDArticulo")

            '//PV: IDDireccionEnvio, IDDireccionFra
            '//AV: IDDireccion, IDDireccionFra
            '//FV: IDDireccion
            If data.Context.ContainsKey("IDDireccionEnvio") Then
                '//PV
                datos.IDDireccionEnvio = Nz(data.Context("IDDireccionEnvio"), 0)
            ElseIf Not data.Context.ContainsKey("IDDireccionEnvio") AndAlso (data.Context.ContainsKey("IDDireccion") AndAlso data.Context.ContainsKey("IDAlbaran")) Then
                '//AV
                datos.IDDireccionEnvio = Nz(data.Context("IDDireccion"), 0)
            End If

            data.Current("IDTipoIva") = ProcessServer.ExecuteTask(Of DataIvaArticuloCliente, String)(AddressOf GetIva, datos, services)

        End If
    End Sub

    <Serializable()> _
    Public Class DataIvaArticuloCliente
        Public IDCliente As String
        Public IDArticulo As String
        Public IDDireccionEnvio As Integer

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDCliente As String, ByVal IDArticulo As String, Optional ByVal IDDireccionEnvio As Integer = 0)
            Me.IDCliente = IDCliente
            Me.IDArticulo = IDArticulo
            Me.IDDireccionEnvio = IDDireccionEnvio
        End Sub
    End Class

    <Task()> Public Shared Function GetIva(ByVal data As ProcesoComercial.DataIvaArticuloCliente, ByVal services As ServiceProvider) As String
        Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
        Dim ClteInfo As ClienteInfo = Clientes.GetEntity(data.IDCliente)
        Dim IVA As String = String.Empty
        If Not IsNothing(ClteInfo) Then
            '//Si el parametro IVACliente está a 1, entonces la función devolverá siempre el iva del cliente, sin importar la nacionalidad.
            '//Si el cliente es extranjero o "CanariasCeutaMelilla" se devuelve el iva del cliente

            'Mirar el tema del dato nuevo en direccion de cliente
            Dim DtDirecEnvio As DataTable = New ClienteDireccion().SelOnPrimaryKey(data.IDDireccionEnvio)
            If Not DtDirecEnvio Is Nothing AndAlso DtDirecEnvio.Rows.Count > 0 Then
                If DtDirecEnvio.Columns.Contains("IDTipoIVA") Then
                    If Length(DtDirecEnvio.Rows(0)("IDTipoIVA")) > 0 Then
                        IVA = DtDirecEnvio.Rows(0)("IDTipoIVA")
                    End If
                End If
            End If

            If Length(IVA) = 0 Then
                Dim AppParamsVenta As ParametroVenta = services.GetService(Of ParametroVenta)()
                If AppParamsVenta.IVACliente OrElse ClteInfo.Extranjero OrElse ClteInfo.CanariasCeutaMelilla Then
                    IVA = ClteInfo.TipoIVA

                    '//Si el cliente no tiene IVA, devolvemos el IVA del artículo.
                    If Length(IVA) = 0 AndAlso Length(data.IDArticulo) > 0 Then
                        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.IDArticulo)
                        IVA = ArtInfo.IDTipoIVA
                    End If
                Else
                    If Length(data.IDArticulo) > 0 Then
                        '//Si el cliente es nacional, se devuelve el IVA del artículo
                        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.IDArticulo)

                        Dim blnAplicarIVAReducido As Boolean = False
                        Dim dtDireccionEnvio As DataTable
                        If data.IDDireccionEnvio <> 0 Then
                            dtDireccionEnvio = New ClienteDireccion().SelOnPrimaryKey(data.IDDireccionEnvio)
                            If dtDireccionEnvio.Rows.Count > 0 AndAlso dtDireccionEnvio.Columns.Contains("CodigoTipoDestino") AndAlso Nz(dtDireccionEnvio.Rows(0)("CodigoTipoDestino"), -1) = 1 Then
                                '//CodigoTipoDestino = 1 (Depósito fiscal)
                                blnAplicarIVAReducido = True
                            Else
                                blnAplicarIVAReducido = False
                            End If
                        Else
                            '//En FV
                            blnAplicarIVAReducido = ClteInfo.IVAReducido
                        End If

                        If blnAplicarIVAReducido Then
                            IVA = ArtInfo.IDTipoIVAReducido
                        Else
                            IVA = ArtInfo.IDTipoIVA
                        End If
                    End If

                    '//Si el artículo no tiene iva, se devuelve el del cliente.
                    If Length(IVA) = 0 Then IVA = ClteInfo.TipoIVA
                End If
            End If

            If Length(IVA) = 0 Then
                ApplicationService.GenerateError("El código de IVA es obligatorio. Revise la relación Articulo-Cliente o que el Artículo tenga IVA.")
            End If

            Return IVA
        End If
    End Function


    <Task()> Public Shared Sub RecuperarTiposIVADireccionEnvio(ByVal doc As DocumentCabLin, ByVal services As ServiceProvider)
        If doc.EntidadCabecera <> "PedidoVentaCabecera" AndAlso doc.EntidadCabecera <> "AlbaranVentaCabecera" Then Exit Sub
       
        Dim context As New DataRowPropertyAccessor(doc.HeaderRow)
        For Each linea As DataRow In doc.dtLineas.Rows
            If linea.RowState <> DataRowState.Deleted Then
                Dim current As New DataRowPropertyAccessor(linea)
                Dim bRules As New BusinessRuleData(String.Empty, Nothing, current, context)

                ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoComercial.ObtenerIvaArticuloCliente, bRules, services)
            End If
        Next
    End Sub

#End Region

#Region " Cálculo Tarifa Comercial "

    'Procedimiento que calcula los campos: PRECIO DTO's , MONEDA y UNIDAD VALORACION teniendo
    'en cuenta el ARTICULO, el CLIENTE y la CANTIDAD en la linea de pedido.
    'Una vez asignados los rdos. obtenidos se comprueba si la moneda obtenida coincide con la
    'moneda que consta en la cabecera del pedido, si no es así, se lleva a cabo el cambio.
    <Task()> Public Shared Sub AplicarTarifaComercial(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If data.Current.ContainsKey("IDArticulo") And data.Context.ContainsKey("IDCliente") And data.Current.ContainsKey("Cantidad") Then
            If Length(data.Current("IDArticulo")) > 0 And Length(data.Context("IDCliente")) > 0 And Nz(data.Current("Cantidad"), 0) <> 0 Then
                Dim dtmFecha As Date
                If data.Context.ContainsKey("Fecha") Then
                    dtmFecha = data.Context("Fecha")
                ElseIf data.Context.ContainsKey("FechaFactura") Then
                    dtmFecha = data.Context("FechaFactura")
                ElseIf data.Context.ContainsKey("FechaAlbaran") Then
                    dtmFecha = data.Context("FechaAlbaran")
                End If
                Dim dataTarifa As New DataCalculoTarifaComercial(data.Current("IDArticulo"), data.Context("IDCliente"), data.Current("Cantidad"), dtmFecha)
                If data.Current.Contains("UDValoracion") Then
                    dataTarifa.UDValoracion = CDbl(Nz(data.Current("UDValoracion"), 1))
                End If

                If data.Context.Contains("Ticket") AndAlso Nz(data.Context("Ticket"), False) AndAlso data.Current.ContainsKey("IDTipoIVA") Then
                    dataTarifa.IDTipoIVA = data.Current("IDTipoIVA") & String.Empty
                End If
                If data.Current.ContainsKey("Regalo") Then
                    dataTarifa.EsRegalo = Nz(data.Current("Regalo"), 0)
                    If data.Current.ContainsKey("IDPromocion") AndAlso Length(data.Current("IDPromocion")) > 0 Then
                        dataTarifa.IDPromocion = data.Current("IDPromocion")
                        dataTarifa.IDPromocionLinea = data.Current("IDPromocionLinea")
                    End If
                End If

                If data.Current.ContainsKey("CantidadAnterior") Then dataTarifa.CantidadAnterior = Nz(data.Current("CantidadAnterior"), 0)
                If data.Current.ContainsKey("IDTarifa") AndAlso Length(data.Current("IDTarifa")) > 0 Then dataTarifa.IDTarifa = data.Current("IDTarifa")
                If data.Current.ContainsKey("IDAlmacen") AndAlso Length(data.Current("IDAlmacen")) > 0 Then dataTarifa.IDAlmacen = data.Current("IDAlmacen")
                ProcessServer.ExecuteTask(Of DataCalculoTarifaComercial, DataTarifaComercial)(AddressOf TarifaComercial, dataTarifa, services)
                If Not dataTarifa.DatosTarifa Is Nothing AndAlso (dataTarifa.DatosTarifa.Precio <> 0 OrElse dataTarifa.DatosTarifa.Dto1 <> 0 OrElse dataTarifa.DatosTarifa.Dto2 <> 0 OrElse dataTarifa.DatosTarifa.Dto3 <> 0 OrElse Length(dataTarifa.DatosTarifa.IDPromocion) > 0) Then
                    If Length(dataTarifa.DatosTarifa.Precio) > 0 Then
                        data.Current("Precio") = dataTarifa.DatosTarifa.Precio
                    End If
                    If Not Nz(data.Current("Regalo")) Then
                        data.Current("IDTarifa") = dataTarifa.DatosTarifa.IDTarifa
                        data.Current("IDPromocion") = dataTarifa.DatosTarifa.IDPromocion
                        data.Current("IDPromocionLinea") = Nz(dataTarifa.DatosTarifa.IDPromocionLinea, DBNull.Value)

                        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf GetDatosPromocion, data, services)
                    End If
                    If data.Context.Contains("Ticket") Then data.Context("Ticket") = dataTarifa.DatosTarifa.TarifaPVP
                    If data.Context.Contains("Ticket") AndAlso Nz(data.Context("Ticket"), False) Then
                        data.Current("PVP") = dataTarifa.DatosTarifa.PVP
                        If data.Current("PVP") = 0 Then
                            data.Current("Precio") = 0
                        End If
                    Else
                        data.Current("PVP") = 0
                    End If
                    data.Current("Dto1") = dataTarifa.DatosTarifa.Dto1
                    data.Current("Dto2") = dataTarifa.DatosTarifa.Dto2
                    data.Current("Dto3") = dataTarifa.DatosTarifa.Dto3
                    If dataTarifa.DatosTarifa.UDValoracion <> 0 Then
                        data.Current("UDValoracion") = dataTarifa.DatosTarifa.UDValoracion
                    Else
                        data.Current("UDValoracion") = 1
                    End If
                    If Length(dataTarifa.DatosTarifa.IDUDVenta) > 0 Then data.Current("IDUDMedida") = dataTarifa.DatosTarifa.IDUDVenta
                    data.Current("SeguimientoTarifa") = dataTarifa.DatosTarifa.SeguimientoTarifa
                    If Length(dataTarifa.DatosTarifa.IDMoneda) > 0 Then data.Current("IDMoneda") = dataTarifa.DatosTarifa.IDMoneda
                    If dataTarifa.DatosTarifa.IDLineaOfertaDetalle <> 0 Then data.Current("IDLineaOfertaDetalle") = dataTarifa.DatosTarifa.IDLineaOfertaDetalle
                Else
                    If data.Current.ContainsKey("IDTarifa") AndAlso Length(data.Current("IDTarifa")) > 0 Then
                        data.Current("IDTarifa") = System.DBNull.Value
                    End If
                End If
                If Not dataTarifa.DatosTarifa Is Nothing Then
                    data.Current("PrecioCosteA") = dataTarifa.DatosTarifa.PrecioCosteA
                    data.Current("PrecioCosteB") = dataTarifa.DatosTarifa.PrecioCosteB
                    data.Current("ImportePVP") = dataTarifa.DatosTarifa.PVP * dataTarifa.Cantidad
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub GetDatosPromocion(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        '//Si venimos del TPV, recogemos la información de la promoción
        If data.Context.ContainsKey("IDTPV") AndAlso Length(data.Context("IDTPV")) > 0 AndAlso Length(data.Current("IDPromocionLinea")) > 0 Then
            Dim dtPromociones As DataTable = AdminData.GetData("vNegDatosPromocion", New NumberFilterItem("IDPromocionLinea", data.Current("IDPromocionLinea")))
            dtPromociones.Columns("QRegaloPromocion").ReadOnly = False
            dtPromociones.Columns("QMaxPromocion").ReadOnly = False

            For Each drPromociones As DataRow In dtPromociones.Rows
                drPromociones("QRegaloPromocion") = 0
                If (data.Current("Cantidad") >= drPromociones("QMinPedido")) Then
                    Dim QServida As Double
                    If data.Current("Cantidad") > drPromociones("QMaxPedido") Then
                        QServida = drPromociones("QMaxPedido")
                    Else
                        QServida = data.Current("Cantidad")
                    End If

                    drPromociones("QRegaloPromocion") = Fix((QServida / Nz(drPromociones("QPedida"), 0))) * Nz(drPromociones("QRegalo"), 0)

                    If Nz(drPromociones("QPedida"), 0) <> 0 Then
                        drPromociones("QMaxPromocion") = (Fix((Nz(drPromociones("QMaxPedido"), 0) / drPromociones("QPedida"))) * Nz(drPromociones("QRegalo"), 0))
                    End If

                    If drPromociones("IDArticulo") & String.Empty = drPromociones("IDArticuloRegalo") & String.Empty Then
                        drPromociones("QMaxPromocion") += Nz(drPromociones("QMaxPedido"), 0)
                    End If
                End If
            Next

            data.Current("DatosPromocion") = dtPromociones
        End If
    End Sub

    <Task()> Public Shared Function TarifaComercial(ByVal data As DataCalculoTarifaComercial, ByVal services As ServiceProvider) As DataTarifaComercial
        Dim strSeguimientoTarifa As String
        If Length(data.IDArticulo) > 0 And data.Cantidad <> 0 Then
            data.DatosTarifa = New DataTarifaComercial
            data.DatosTarifa.Precio = 0
            data.DatosTarifa.Dto1 = 0
            data.DatosTarifa.Dto2 = 0
            data.DatosTarifa.Dto3 = 0
            data.DatosTarifa.UDValoracion = 1
            data.DatosTarifa.SeguimientoTarifa = String.Empty
            ProcessServer.ExecuteTask(Of DataCalculoTarifaComercial)(AddressOf TarifaGeneralMaximaPrioridad, data, services)
            ProcessServer.ExecuteTask(Of DataCalculoTarifaComercial)(AddressOf TarifaPromocionArticulo, data, services)
            ProcessServer.ExecuteTask(Of DataCalculoTarifaComercial)(AddressOf TarifaArticuloCliente, data, services)
            ProcessServer.ExecuteTask(Of DataCalculoTarifaComercial)(AddressOf TarifaSeleccionada, data, services)  'TPV
            ProcessServer.ExecuteTask(Of DataCalculoTarifaComercial)(AddressOf TarifaCliente, data, services)
            ProcessServer.ExecuteTask(Of DataCalculoTarifaComercial)(AddressOf TarifaCentroGestion, data, services)
            ProcessServer.ExecuteTask(Of DataCalculoTarifaComercial)(AddressOf TarifaGeneral, data, services)
            ProcessServer.ExecuteTask(Of DataCalculoTarifaComercial)(AddressOf RecuperarDatosTarifa, data, services)
            ProcessServer.ExecuteTask(Of DataCalculoTarifaComercial)(AddressOf AplicarTarifaPadre, data, services)
            ProcessServer.ExecuteTask(Of DataCalculoTarifaComercial)(AddressOf PrecioEnMonedaContexto, data, services)
            ProcessServer.ExecuteTask(Of DataCalculoTarifaComercial)(AddressOf TratarSeguimiento, data, services)
            ProcessServer.ExecuteTask(Of DataCalculoTarifaComercial)(AddressOf TarifaCosteArticulo, data, services)
        End If
        Return data.DatosTarifa
    End Function

    <Task()> Public Shared Sub TarifaGeneralMaximaPrioridad(ByVal data As DataCalculoTarifaComercial, ByVal services As ServiceProvider)
        data.Cantidad = Math.Abs(data.Cantidad)
        '//1º- TARIFA DE UTILIZACIÓN GENERAL DE MÁXIMA PRIORIDAD.-
        '//    Buscamos en las tarifas generales que tengan la marca de 'Máxima Prioridad'

        Dim dtmFecha As Date = Today
        If Not data.Fecha Is Nothing AndAlso data.Fecha <> cnMinDate Then dtmFecha = data.Fecha

        Dim dtTarifa As DataTable = New Tarifa().Filter(New BooleanFilterItem("MaxPrioridad", True))
        If Not dtTarifa Is Nothing AndAlso dtTarifa.Rows.Count > 0 Then
            '//Comprobamos la vigencia de las tarifas con Máxima Prioridad
            Dim dtTarifaArt As DataTable
            Dim TA As New TarifaArticulo
            For Each drTarifa As DataRow In dtTarifa.Rows
                Dim blnTarifaVigente As Boolean = ProcessServer.ExecuteTask(Of DataTarifaVigente, Boolean)(AddressOf TarifaVigente, New DataTarifaVigente(drTarifa("IdTarifa"), dtmFecha), services)
                If blnTarifaVigente Then
                    '//Comprobamos si el artículo está en la tarifa
                    dtTarifaArt = TA.SelOnPrimaryKey(drTarifa("IdTarifa"), data.IDArticulo)
                    If Not dtTarifaArt Is Nothing AndAlso dtTarifaArt.Rows.Count > 0 Then
                        Exit For
                    End If
                End If
            Next

            If Not dtTarifaArt Is Nothing AndAlso dtTarifaArt.Rows.Count > 0 Then
                data.DatosTarifa.IDTarifa = dtTarifaArt.Rows(0)("IdTarifa") & String.Empty
                data.DatosTarifa.TarifaPVP = Nz(dtTarifa.Rows(0)("TarifaPVP"), 0)
                data.DatosTarifa.SeguimientoTarifa = "TARIFA MAXIMA PRIORIDAD"
            End If
        End If
    End Sub

    <Serializable()> _
    Public Class DataTarifaVigente
        Public IDTarifa As String
        Public Fecha As Date

        Public Sub New(ByVal IDTarifa As String, ByVal Fecha As Date)
            Me.IDTarifa = IDTarifa
            Me.Fecha = Fecha
        End Sub
    End Class
    <Task()> Public Shared Function TarifaVigente(ByVal data As DataTarifaVigente, ByVal services As ServiceProvider) As Boolean
        Dim dtTarifa As DataTable = New BE.DataEngine().Filter("vTarifasVigentes", New StringFilterItem("IDTarifa", data.IDTarifa))
        If Not IsNothing(dtTarifa) AndAlso dtTarifa.Rows.Count > 0 Then
            If Length(dtTarifa.Rows(0)("FechaDesde")) > 0 And Length(dtTarifa.Rows(0)("FechaHasta")) > 0 Then
                If (dtTarifa.Rows(0)("FechaDesde") <= data.Fecha And dtTarifa.Rows(0)("FechaHasta") >= data.Fecha) Then
                    TarifaVigente = True
                Else
                    TarifaVigente = False
                End If
            Else
                TarifaVigente = True
            End If
        End If
    End Function

    <Serializable()> _
    Public Class DataOfertaComercialVigente
        Public IDLineaOfertaDetalle As Integer
        Public Fecha As Date

        Public Sub New(ByVal IDLineaOfertaDetalle As Integer, ByVal Fecha As Date)
            Me.IDLineaOfertaDetalle = IDLineaOfertaDetalle
            Me.Fecha = Fecha
        End Sub
    End Class
    <Task()> Public Shared Function OfertaComercialVigente(ByVal data As DataOfertaComercialVigente, ByVal services As ServiceProvider) As Boolean
        Dim blnVigente As Boolean = True
        If data.IDLineaOfertaDetalle > 0 Then
            Dim objFilter As New Filter
            objFilter.Add(New NumberFilterItem("IDLineaOfertaDetalle", data.IDLineaOfertaDetalle))
            Dim dt As DataTable = AdminData.GetData("tbOfertaComercialDetalle", objFilter, "IDOfertaComercial")
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                objFilter.Clear()
                objFilter.Add(New NumberFilterItem("IDOfertaComercial", dt.Rows(0)("IDOfertaComercial")))
                dt = AdminData.GetData("tbOfertaComercialCabecera", objFilter, "FechaInicioOferta,FechaFinOferta")
                If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                    If IsDate(dt.Rows(0)("FechaInicioOferta")) Then
                        blnVigente = dt.Rows(0)("FechaInicioOferta") <= data.Fecha
                    Else
                        blnVigente = True
                    End If
                    If IsDate(dt.Rows(0)("FechaFinOferta")) Then
                        blnVigente = blnVigente And (dt.Rows(0)("FechaFinOferta") >= data.Fecha)
                    End If
                End If
            End If
        End If

        Return blnVigente
    End Function

    <Task()> Public Shared Sub TarifaPromocionArticulo(ByVal data As DataCalculoTarifaComercial, ByVal services As ServiceProvider)
        If data.DatosTarifa.Precio <> 0 OrElse Length(data.DatosTarifa.IDTarifa) > 0 Then Exit Sub
        If Length(data.DatosTarifa.IDTarifa) > 0 Then Exit Sub
        '//2º .- TARIFA DE PROMOCIÓN APLICABLE AL PRODUCTO
        '//      Buscamos si hay una promoción aplicable al artículo.
        Dim dtTarifa As DataTable
        Dim TarifaPVP As Boolean
        Dim StDatos As New PromocionCabecera.DatosCompPromo(data.IDArticulo, data.Cantidad, data.CantidadAnterior, data.IDCliente, data.Fecha)
        Dim dtPromocioncab As DataTable = ProcessServer.ExecuteTask(Of PromocionCabecera.DatosCompPromo, DataTable)(AddressOf PromocionCabecera.ComprobarPromociones, StDatos, services)
        Dim BlnHayPromocion As Boolean = False
        If Not dtPromocioncab Is Nothing AndAlso dtPromocioncab.Rows.Count > 0 Then
            '//Existe una promoción para aplicar
            '//Comprobamos que no se exceda la cantidad máxima promocionable que ya tiene la cantidad disponible
            If Nz(dtPromocioncab.Rows(0)("QMaxPromocionable"), 0) > 0 Then
                '//Comprobamos la vigencia de la Tarifa
                'If TarifaVigente(dtPromocioncab.Rows(0)("IDTarifa"), data.Fecha) Then
                If Length(dtPromocioncab.Rows(0)("IDTarifa")) > 0 Then
                    Dim blnTarifaVigente As Boolean = ProcessServer.ExecuteTask(Of DataTarifaVigente, Boolean)(AddressOf TarifaVigente, New DataTarifaVigente(dtPromocioncab.Rows(0)("IDTarifa"), data.Fecha), services)
                    If blnTarifaVigente Then
                        BlnHayPromocion = True
                        dtTarifa = dtPromocioncab.Copy
                        Dim dt As DataTable = New Tarifa().SelOnPrimaryKey(dtPromocioncab.Rows(0)("IDTarifa"))
                        If dt.Rows.Count > 0 Then
                            TarifaPVP = Nz(dt.Rows(0)("TarifaPVP"), False)
                        End If
                    End If
                Else
                    BlnHayPromocion = True
                    dtTarifa = dtPromocioncab.Copy
                End If
            Else
                'La cantidad total promocionada supera la cantidad máxima promocionable
            End If
        ElseIf data.EsRegalo AndAlso Length(data.IDPromocion) > 0 Then
            Dim DtPromo As DataTable = New PromocionCabecera().SelOnPrimaryKey(data.IDPromocion)
            If Not DtPromo Is Nothing AndAlso DtPromo.Rows.Count > 0 Then
                If Length(DtPromo.Rows(0)("IDTarifa")) > 0 Then
                    data.DatosTarifa.IDTarifa = DtPromo.Rows(0)("IDTarifa")
                    Dim dt As DataTable = New Tarifa().SelOnPrimaryKey(DtPromo.Rows(0)("IDTarifa"))
                    If dt.Rows.Count > 0 Then
                        data.DatosTarifa.TarifaPVP = Nz(dt.Rows(0)("TarifaPVP"), False)
                    End If
                    data.DatosTarifa.IDPromocion = data.IDPromocion
                    data.DatosTarifa.IDPromocionLinea = data.IDPromocionLinea
                    data.DatosTarifa.SeguimientoTarifa &= " TARIFA EN PROMOCION"
                End If
            End If
        End If

        If BlnHayPromocion Then
            If Not dtTarifa Is Nothing AndAlso dtTarifa.Rows.Count > 0 Then
                data.DatosTarifa.IDTarifa = dtTarifa.Rows(0)("IdTarifa") & String.Empty
                data.DatosTarifa.TarifaPVP = TarifaPVP 'Nz(dtTarifa.Rows(0)("TarifaPVP"), 0)
            End If
            data.DatosTarifa.IDPromocion = Nz(dtTarifa.Rows(0)("IDPromocion"), String.Empty)
            data.DatosTarifa.IDPromocionLinea = Nz(dtTarifa.Rows(0)("IDPromocionLinea"), -1)
            data.DatosTarifa.SeguimientoTarifa &= " TARIFA EN PROMOCION"
        End If
    End Sub

    <Task()> Public Shared Sub TarifaArticuloCliente(ByVal data As DataCalculoTarifaComercial, ByVal services As ServiceProvider)
        If data.DatosTarifa.Precio <> 0 OrElse Length(data.DatosTarifa.IDTarifa) > 0 Then Exit Sub
        If Length(data.IDCliente) > 0 Then
            '//3º .- PRECIO ARTICULO DEL CLIENTE
            '//      Buscamos si el cliente tiene asociado el artículo.
            Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
            Dim ClteInfo As ClienteInfo = Clientes.GetEntity(data.IDCliente)

            Dim ArtCltes As EntityInfoCache(Of ArticuloClienteInfo) = services.GetService(Of EntityInfoCache(Of ArticuloClienteInfo))()
            Dim ArtClte As ArticuloClienteInfo = ArtCltes.GetEntity(data.IDCliente, data.IDArticulo)
            If Not ArtClte Is Nothing AndAlso Length(ArtClte.IDArticulo) > 0 AndAlso Length(ArtClte.IDCliente) > 0 Then
                'TODO HISTREV
                '//Si el artículo esta asociado con el cliente y no tenemos un precio o descuentos establecidos para el Articulo - Cliente a nivel de linea,
                '// obtenemos el Precio correspondiente al Articulo Cliente.
                'Metemos 
                Dim FilHistQDesdeClie As New Filter
                FilHistQDesdeClie.Add("FechaDesde", FilterOperator.LessThanOrEqual, data.Fecha)
                FilHistQDesdeClie.Add("FechaHasta", FilterOperator.GreaterThanOrEqual, data.Fecha)
                FilHistQDesdeClie.Add("IDCliente", FilterOperator.Equal, data.IDCliente)
                FilHistQDesdeClie.Add("IDArticulo", FilterOperator.Equal, data.IDArticulo)
                FilHistQDesdeClie.Add("QDesde", FilterOperator.LessThanOrEqual, data.Cantidad)
                Dim DtHistQDesdeClie As DataTable = New HistoricoPreciosClienteQDesde().Filter(FilHistQDesdeClie)
                If Not DtHistQDesdeClie Is Nothing AndAlso DtHistQDesdeClie.Rows.Count > 0 Then
                    data.DatosTarifa.Precio = Nz(DtHistQDesdeClie.Rows(0)("Precio"), 0)
                    data.DatosTarifa.Dto1 = Nz(DtHistQDesdeClie.Rows(0)("Dto1"), 0)
                    data.DatosTarifa.Dto2 = Nz(DtHistQDesdeClie.Rows(0)("Dto2"), 0)
                    data.DatosTarifa.Dto3 = Nz(DtHistQDesdeClie.Rows(0)("Dto3"), 0)
                Else
                    Dim FilClte As New Filter
                    FilClte.Add("IDCliente", FilterOperator.Equal, data.IDCliente)
                    FilClte.Add("IDArticulo", FilterOperator.Equal, data.IDArticulo)
                    FilClte.Add("QDesde", FilterOperator.LessThanOrEqual, data.Cantidad)

                    Dim dtACLinea As DataTable = New ArticuloClienteLinea().Filter(FilClte, "QDesde DESC", "TOP 1 *")
                    If Not dtACLinea Is Nothing AndAlso dtACLinea.Rows.Count > 0 Then
                        data.DatosTarifa.Precio = Nz(dtACLinea.Rows(0)("Precio"), 0)
                        data.DatosTarifa.PVP = Nz(dtACLinea.Rows(0)("PVP"), 0)
                        data.DatosTarifa.Dto1 = Nz(dtACLinea.Rows(0)("Dto1"), 0)
                        data.DatosTarifa.Dto2 = Nz(dtACLinea.Rows(0)("Dto2"), 0)
                        data.DatosTarifa.Dto3 = Nz(dtACLinea.Rows(0)("Dto3"), 0)
                    Else
                        Dim FilHistClie As New Filter
                        FilHistClie.Add("FechaDesde", FilterOperator.LessThanOrEqual, data.Fecha)
                        FilHistClie.Add("FechaHasta", FilterOperator.GreaterThanOrEqual, data.Fecha)
                        FilHistClie.Add("IDCliente", FilterOperator.Equal, data.IDCliente)
                        FilHistClie.Add("IDArticulo", FilterOperator.Equal, data.IDArticulo)
                        Dim DtHistClie As DataTable = New HistoricoPreciosCliente().Filter(FilHistClie)
                        If Not DtHistClie Is Nothing AndAlso DtHistClie.Rows.Count > 0 Then
                            data.DatosTarifa.Precio = Nz(DtHistClie.Rows(0)("Precio"), 0)
                            data.DatosTarifa.Dto1 = Nz(DtHistClie.Rows(0)("Dto1"), 0)
                            data.DatosTarifa.Dto2 = Nz(DtHistClie.Rows(0)("Dto2"), 0)
                            data.DatosTarifa.Dto3 = Nz(DtHistClie.Rows(0)("Dto3"), 0)
                        Else
                            data.DatosTarifa.Precio = ArtClte.Precio
                            data.DatosTarifa.Dto1 = ArtClte.Dto1
                            data.DatosTarifa.Dto2 = ArtClte.Dto2
                            data.DatosTarifa.Dto3 = ArtClte.Dto3
                            data.DatosTarifa.PVP = ArtClte.PVP
                        End If
                    End If
                End If

                data.DatosTarifa.UDValoracion = ArtClte.UdValoracion
                data.DatosTarifa.IDMoneda = ClteInfo.Moneda
                data.DatosTarifa.IDUDVenta = ArtClte.IDUDVenta
                If data.DatosTarifa.Precio = 0 AndAlso (data.DatosTarifa.Dto1 <> 0 OrElse data.DatosTarifa.Dto2 <> 0 OrElse data.DatosTarifa.Dto3 <> 0) Then
                    data.DatosTarifa.SeguimientoDtos &= " DESCUENTOS OBTENIDO DE LA RELACION CLIENTE - ARTÍCULO"

                Else
                    data.DatosTarifa.SeguimientoTarifa &= " PRECIO OBTENIDO DE LA RELACION CLIENTE - ARTÍCULO"
                End If
                If ProcessServer.ExecuteTask(Of DataOfertaComercialVigente, Boolean)(AddressOf OfertaComercialVigente, New DataOfertaComercialVigente(ArtClte.IDLineaOfertaDetalle, data.Fecha), services) Then
                    data.DatosTarifa.IDLineaOfertaDetalle = ArtClte.IDLineaOfertaDetalle
                Else
                    '///Si la oferta no esta vigente, se pone inicializan los datos recuperados de la Tarifa, para que sigua con el proceso de búsqueda.
                    data.DatosTarifa.Precio = 0
                    data.DatosTarifa.PVP = 0
                    data.DatosTarifa.Dto1 = 0
                    data.DatosTarifa.Dto2 = 0
                    data.DatosTarifa.Dto3 = 0
                    data.DatosTarifa.UDValoracion = 0
                    data.DatosTarifa.IDMoneda = String.Empty
                    data.DatosTarifa.IDUDVenta = String.Empty
                    data.DatosTarifa.SeguimientoTarifa = String.Empty
                End If
            End If

            '//4º y 5º .- DTO. FAMILIA CLIENTE. Sólo obtendremos los descuentos en caso que los haya.
            If Not IsNothing(ClteInfo) Then
                ProcessServer.ExecuteTask(Of DataCalculoTarifaComercial)(AddressOf DescuentosPorFamilia, data, services)

                '//6º .- GRUPO CLIENTE. Miramos si el cliente tiene el 'GrupoTarifa' activo
                '//      Si lo tiene, el IdCliente con el que trabajaremos será el del grupo.
                If Length(ClteInfo.GrupoCliente) > 0 AndAlso ClteInfo.GrupoTarifa Then
                    data.IDCliente = ClteInfo.GrupoCliente
                    ProcessServer.ExecuteTask(Of DataCalculoTarifaComercial)(AddressOf TarifaArticuloCliente, data, services)
                    If Len(data.DatosTarifa.SeguimientoTarifa) > 0 Then data.DatosTarifa.SeguimientoTarifa = data.DatosTarifa.SeguimientoTarifa & "\"
                    data.DatosTarifa.SeguimientoTarifa = data.DatosTarifa.SeguimientoTarifa & " CLIENTE TIENE GRUPO"
                    data.IDCliente = ClteInfo.IDCliente
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub DescuentosPorFamilia(ByVal data As DataCalculoTarifaComercial, ByVal services As ServiceProvider)
        If data.DatosTarifa.Dto1 <> 0 OrElse data.DatosTarifa.Dto2 <> 0 OrElse data.DatosTarifa.Dto3 <> 0 OrElse Length(data.DatosTarifa.IDTarifa) > 0 Then Exit Sub
        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.IDArticulo)

        If Not ArtInfo Is Nothing Then
            If Len(ArtInfo.IDTipo) > 0 Then
                '//Primero se buscan los Descuentos por Cliente-Tipo-Familia
                Dim f As New Filter
                f.Add(New StringFilterItem("IDCliente", data.IDCliente))
                f.Add(New StringFilterItem("IDTipo", ArtInfo.IDTipo))
                If Len(ArtInfo.IDFamilia) > 0 Then
                    f.Add(New StringFilterItem("IDFamilia", ArtInfo.IDFamilia))
                End If
                Dim CDF As New ClienteDescuentoFamilia
                Dim dtFamiliaDtosClte As DataTable = CDF.Filter(f)

                '//Si no encuentra Descuentos por Cliente-Tipo-Familia se buscará Cliente-Tipo
                If dtFamiliaDtosClte.Rows.Count = 0 And Len(ArtInfo.IDFamilia) > 0 Then
                    f.Clear()
                    f.Add(New StringFilterItem("IDCliente", data.IDCliente))
                    f.Add(New StringFilterItem("IDTipo", ArtInfo.IDTipo))
                    f.Add(New IsNullFilterItem("IDFamilia"))

                    dtFamiliaDtosClte = CDF.Filter(f)
                    data.DatosTarifa.SeguimientoDtos = "CLIENTE-TIPO"
                Else
                    data.DatosTarifa.SeguimientoDtos = "CLIENTE-TIPO-FAMILIA"
                End If

                If dtFamiliaDtosClte.Rows.Count > 0 Then
                    data.DatosTarifa.Dto1 = dtFamiliaDtosClte.Rows(0)("Dto1")
                    data.DatosTarifa.Dto2 = dtFamiliaDtosClte.Rows(0)("Dto2")
                    data.DatosTarifa.Dto3 = dtFamiliaDtosClte.Rows(0)("Dto3")

                    ' If Length(data.DatosTarifa.SeguimientoTarifa) > 0 Then data.DatosTarifa.SeguimientoTarifa = data.DatosTarifa.SeguimientoTarifa & "\"
                    ' data.DatosTarifa.SeguimientoTarifa = data.DatosTarifa.SeguimientoTarifa & "DESCUENTOS POR FAMILIA"
                    If data.DatosTarifa.Precio = 0 Then
                        If ProcessServer.ExecuteTask(Of DataOfertaComercialVigente, Boolean)(AddressOf OfertaComercialVigente, New DataOfertaComercialVigente(Nz(dtFamiliaDtosClte.Rows(0)("IDLineaOfertaDetalle"), 0), data.Fecha), services) Then
                            data.DatosTarifa.IDLineaOfertaDetalle = Nz(dtFamiliaDtosClte.Rows(0)("IDLineaOfertaDetalle"), 0)
                        Else
                            '// Si la oferta no está vigente, se inicializan los valores para que sigua realizándose la búsqueda.
                            data.DatosTarifa.Precio = 0
                            data.DatosTarifa.PVP = 0
                            data.DatosTarifa.Dto1 = 0
                            data.DatosTarifa.Dto2 = 0
                            data.DatosTarifa.Dto3 = 0
                            data.DatosTarifa.IDMoneda = String.Empty
                            data.DatosTarifa.IDPromocion = String.Empty
                            data.DatosTarifa.IDPromocionLinea = Nothing
                            data.DatosTarifa.IDTarifa = String.Empty
                            data.DatosTarifa.IDUDVenta = String.Empty
                            data.DatosTarifa.SeguimientoTarifa = String.Empty
                            data.DatosTarifa.SeguimientoDtos = String.Empty
                            data.DatosTarifa.UDValoracion = 0
                        End If
                    End If
                Else
                    data.DatosTarifa.SeguimientoDtos = String.Empty
                    Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
                    Dim ClieInfo As ClienteInfo = Clientes.GetEntity(data.IDCliente)
                    If ClieInfo.DtoComercialLinea > 0 Then
                        data.DatosTarifa.SeguimientoDtos = "CLIENTE. DESCUENTO COMERCIAL LINEA"
                        data.DatosTarifa.Dto1 = ClieInfo.DtoComercialLinea
                        data.DatosTarifa.Dto2 = 0
                        data.DatosTarifa.Dto3 = 0
                    Else
                        '//Si no hay descuentos por Familia-Cliente se buscan por Familia
                        If Length(ArtInfo.IDFamilia) > 0 Then
                            Dim dtFamDtos As DataTable = New FamiliaDescuentos().SelOnPrimaryKey(ArtInfo.IDTipo, ArtInfo.IDFamilia)
                            If dtFamDtos.Rows.Count > 0 Then
                                data.DatosTarifa.SeguimientoDtos = "TIPO-FAMILIA"
                                data.DatosTarifa.Dto1 = Nz(dtFamDtos.Rows(0)("Dto1"), 0)
                                data.DatosTarifa.Dto2 = Nz(dtFamDtos.Rows(0)("Dto2"), 0)
                                data.DatosTarifa.Dto3 = Nz(dtFamDtos.Rows(0)("Dto3"), 0)
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub TarifaSeleccionada(ByVal data As DataCalculoTarifaComercial, ByVal services As ServiceProvider)
        If data.DatosTarifa.Precio <> 0 OrElse Length(data.DatosTarifa.IDTarifa) > 0 Then Exit Sub
        '//Tarifa indicada por el usuario (TPV)
        If Length(data.IDTarifa) > 0 Then
            Dim dtTarifa As DataTable = New Tarifa().SelOnPrimaryKey(data.IDTarifa)
            If Not dtTarifa Is Nothing AndAlso dtTarifa.Rows.Count > 0 Then
                'Compruebo si la tarifa es vigente
                Dim blnTarifaVigente As Boolean = ProcessServer.ExecuteTask(Of DataTarifaVigente, Boolean)(AddressOf TarifaVigente, New DataTarifaVigente(data.IDTarifa, data.Fecha), services)
                If blnTarifaVigente Then
                    'Compruebo si el artículo está en la tarifa
                    dtTarifa = New TarifaArticulo().SelOnPrimaryKey(data.IDTarifa, data.IDArticulo)
                    If Not dtTarifa Is Nothing AndAlso dtTarifa.Rows.Count > 0 Then
                        data.DatosTarifa.IDTarifa = data.IDTarifa
                        If Len(data.DatosTarifa.SeguimientoTarifa) > 0 Then data.DatosTarifa.SeguimientoTarifa = data.DatosTarifa.SeguimientoTarifa & "\"
                        data.DatosTarifa.SeguimientoTarifa = data.DatosTarifa.SeguimientoTarifa & " TARIFA SELECCIONADA"
                    End If
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub TarifaCliente(ByVal data As DataCalculoTarifaComercial, ByVal services As ServiceProvider)
        If data.DatosTarifa.Precio <> 0 OrElse Length(data.DatosTarifa.IDTarifa) > 0 Then Exit Sub
        If Length(data.IDCliente) > 0 Then
            '//7º .- TARIFAS PROPIAS DEL CLIENTE
            '//      Buscamos si hay en las tarifa del cliente.
            '//      Buscamos todas las tarifas del cliente que no sean predeterminadas ordenadas por el campo "Orden"
            Dim dtClienteTarifa As DataTable = New ClienteTarifa().Filter(New StringFilterItem("IDCliente", data.IDCliente), "Orden")
            If Not dtClienteTarifa Is Nothing AndAlso dtClienteTarifa.Rows.Count > 0 Then
                Dim TA As New TarifaArticulo
                For Each drCT As DataRow In dtClienteTarifa.Rows
                    Dim blnTarifaVigente As Boolean = ProcessServer.ExecuteTask(Of DataTarifaVigente, Boolean)(AddressOf TarifaVigente, New DataTarifaVigente(drCT("IdTarifa"), data.Fecha), services)
                    If blnTarifaVigente Then
                        'Compruebo si el artículo está en la tarifa
                        Dim dtTarifaArt As DataTable = TA.SelOnPrimaryKey(drCT("IdTarifa"), data.IDArticulo)
                        If Not dtTarifaArt Is Nothing AndAlso dtTarifaArt.Rows.Count > 0 Then
                            If ProcessServer.ExecuteTask(Of DataOfertaComercialVigente, Boolean)(AddressOf OfertaComercialVigente, New DataOfertaComercialVigente(Nz(drCT("IDLineaOfertaDetalle"), 0), data.Fecha), services) Then
                                data.DatosTarifa.IDTarifa = drCT("IdTarifa")
                                data.DatosTarifa.IDLineaOfertaDetalle = Nz(drCT("IDLineaOfertaDetalle"), 0)
                                If Len(data.DatosTarifa.SeguimientoTarifa) > 0 Then data.DatosTarifa.SeguimientoTarifa = data.DatosTarifa.SeguimientoTarifa & "\"
                                data.DatosTarifa.SeguimientoTarifa = data.DatosTarifa.SeguimientoTarifa & " TARIFA CLIENTE"
                                Exit For
                            End If
                        End If
                    End If
                Next
            End If

            '//8º .- TARIFAS  DEL CLIENTE GRUPO 
            '//      Buscamos si hay en las tarifa del cliente.
            '//      Buscamos todas las tarifas del cliente que no sean predeterminadas ordenadas por el campo "Orden"
            Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
            Dim ClteInfo As ClienteInfo = Clientes.GetEntity(data.IDCliente)

            If Length(data.DatosTarifa.IDTarifa) = 0 AndAlso Length(ClteInfo.GrupoCliente) > 0 AndAlso ClteInfo.GrupoTarifa Then
                data.IDCliente = ClteInfo.GrupoCliente
                ProcessServer.ExecuteTask(Of DataCalculoTarifaComercial)(AddressOf TarifaCliente, data, services)
                If Len(data.DatosTarifa.SeguimientoTarifa) > 0 Then data.DatosTarifa.SeguimientoTarifa = data.DatosTarifa.SeguimientoTarifa & "\"
                data.DatosTarifa.SeguimientoTarifa = data.DatosTarifa.SeguimientoTarifa & " CLIENTE TIENE GRUPO"
                data.IDCliente = ClteInfo.IDCliente
            End If
        End If
    End Sub

    <Task()> Public Shared Sub TarifaCentroGestion(ByVal data As DataCalculoTarifaComercial, ByVal services As ServiceProvider)
        If data.DatosTarifa.Precio <> 0 OrElse Length(data.DatosTarifa.IDTarifa) > 0 Then Exit Sub
        '//8º .- TARIFAS RELACIONADAS CENTRO GESTION
        '//      Buscamos en la tarifas relacionadas con el centro de gestión.
        If Length(data.IDCliente) > 0 Then
            Dim dtCentroGestion As DataTable
            Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
            Dim ClteInfo As ClienteInfo = Clientes.GetEntity(data.IDCliente)
            If Length(ClteInfo.CentroGestion) > 0 Then dtCentroGestion = New CentroGestion().SelOnPrimaryKey(ClteInfo.CentroGestion)

            If Not dtCentroGestion Is Nothing AndAlso dtCentroGestion.Rows.Count > 0 AndAlso Length(dtCentroGestion.Rows(0)("IdTarifa")) > 0 Then
                Dim blnTarifaVigente As Boolean = ProcessServer.ExecuteTask(Of DataTarifaVigente, Boolean)(AddressOf TarifaVigente, New DataTarifaVigente(dtCentroGestion.Rows(0)("IdTarifa"), data.Fecha), services)
                If blnTarifaVigente Then
                    'Compruebo si el artículo está en la tarifa
                    Dim TA As New TarifaArticulo
                    Dim dtTarifaOk As DataTable = TA.SelOnPrimaryKey(dtCentroGestion.Rows(0)("IdTarifa"), data.IDArticulo)
                    If Not dtTarifaOk Is Nothing AndAlso dtTarifaOk.Rows.Count > 0 Then
                        data.DatosTarifa.IDTarifa = dtCentroGestion.Rows(0)("IdTarifa") & String.Empty
                        data.DatosTarifa.SeguimientoTarifa = " TARIFA DEL CENTRO DE GESTION DEL CLIENTE"
                    End If
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub TarifaGeneral(ByVal data As DataCalculoTarifaComercial, ByVal services As ServiceProvider)
        If data.DatosTarifa.Precio <> 0 OrElse Length(data.DatosTarifa.IDTarifa) > 0 Then Exit Sub
        '//9º .- TARIFAS GENERALES
        '//      Buscamos en la tarifas generales. Son las tarifas que están en curso
        'Ordenamos las Tarifas generales por el campo "Orden".
        Dim dtTarifa As DataTable = New TarifaOrden().Filter(, , "Orden")
        If Not dtTarifa Is Nothing AndAlso dtTarifa.Rows.Count > 0 Then
            Dim TA As New TarifaArticulo
            For Each drTarifa As DataRow In dtTarifa.Rows
                'Compruebo si la tarifa es vigente
                Dim blnTarifaVigente As Boolean = ProcessServer.ExecuteTask(Of DataTarifaVigente, Boolean)(AddressOf TarifaVigente, New DataTarifaVigente(drTarifa("IdTarifa"), data.Fecha), services)
                If blnTarifaVigente Then
                    'Compruebo si el artículo está en la tarifa
                    Dim dtTarifaOk As DataTable = TA.SelOnPrimaryKey(drTarifa("IdTarifa"), data.IDArticulo)
                    If Not dtTarifaOk Is Nothing AndAlso dtTarifaOk.Rows.Count > 0 Then
                        data.DatosTarifa.IDTarifa = dtTarifaOk.Rows(0)("IdTarifa") & String.Empty
                        data.DatosTarifa.SeguimientoTarifa = " TARIFAS GENERALES"
                        Exit For
                    End If
                End If
            Next
        End If
    End Sub

    <Task()> Public Shared Sub RecuperarDatosTarifa(ByVal data As DataCalculoTarifaComercial, ByVal services As ServiceProvider)
        If data.DatosTarifa.Precio <> 0 Then Exit Sub

        If Length(data.DatosTarifa.IDTarifa) > 0 Then
            '//Obtenemos los datos de TarifaArticulo
            Dim dtTarifaArticulo As DataTable = New TarifaArticulo().SelOnPrimaryKey(data.DatosTarifa.IDTarifa, data.IDArticulo)
            If Not dtTarifaArticulo Is Nothing AndAlso dtTarifaArticulo.Rows.Count > 0 Then
                '//Obtenemos los datos de la tarifa a nivel de TarifaArticuloLinea
                Dim f As New Filter
                f.Add("IDTarifa", FilterOperator.Equal, dtTarifaArticulo.Rows(0)("IdTarifa"))
                f.Add("IDArticulo", FilterOperator.Equal, data.IDArticulo)
                f.Add("QDesde", FilterOperator.LessThanOrEqual, data.Cantidad)

                Dim TAL As New TarifaArticuloLinea
                Dim dtDatosLinea As DataTable = TAL.Filter(f, "QDesde DESC", "TOP 1 *")
                If Not dtDatosLinea Is Nothing AndAlso dtDatosLinea.Rows.Count > 0 Then

                Else
                    'Si no tenemos líneas copiaremos loa datos de la cabecera
                    dtDatosLinea = dtTarifaArticulo.Copy
                End If

                Dim Tarifas As EntityInfoCache(Of TarifaInfo) = services.GetService(Of EntityInfoCache(Of TarifaInfo))()
                Dim TarInfo As TarifaInfo = Tarifas.GetEntity(data.DatosTarifa.IDTarifa)

                data.DatosTarifa.Precio = dtDatosLinea.Rows(0)("Precio")
                data.DatosTarifa.TarifaPVP = TarInfo.TarifaPVP
                If data.DatosTarifa.TarifaPVP Then
                    data.DatosTarifa.PVP = dtDatosLinea.Rows(0)("PVP")
                    If data.DatosTarifa.PVP = 0 AndAlso data.DatosTarifa.Precio <> 0 Then
                        If Length(data.IDTipoIVA) = 0 Then
                            Dim TiposIVA As EntityInfoCache(Of TipoIvaInfo) = services.GetService(Of EntityInfoCache(Of TipoIvaInfo))()
                            Dim TIVAInfo As TipoIvaInfo = TiposIVA.GetEntity(data.IDTipoIVA)
                            ' HistoricoTipoIVA
                            If data.Fecha = cnMinDate Then data.Fecha = Today
                            Dim IVAInfo As TipoIvaInfo = TiposIVA.GetEntity(data.IDTipoIVA, data.Fecha)
                            Dim FactorIVA As Double
                            If Not IsNothing(IVAInfo) Then
                                FactorIVA = IVAInfo.Factor
                                If IVAInfo.SinRepercutir Then FactorIVA = IVAInfo.IVASinRepercutir
                            End If
                            data.DatosTarifa.PVP = data.DatosTarifa.Precio * (1 + FactorIVA / 100)

                            'data.DatosTarifa.PVP = data.DatosTarifa.Precio * (1 - data.DatosTarifa.Dto1 / 100) * (1 - data.DatosTarifa.Dto2 / 100) * (1 - data.DatosTarifa.Dto3 / 100) * (1 - dblDto / 100) * (1 - dblDtoProntoPago / 100) * (1 + FactorIVA / 100)
                        End If
                    End If
                Else
                    data.DatosTarifa.PVP = 0
                End If
                If data.DatosTarifa.Dto1 = 0 AndAlso data.DatosTarifa.Dto2 = 0 AndAlso data.DatosTarifa.Dto3 = 0 Then
                    data.DatosTarifa.Dto1 = dtDatosLinea.Rows(0)("Dto1")
                    data.DatosTarifa.Dto2 = dtDatosLinea.Rows(0)("Dto2")
                    data.DatosTarifa.Dto3 = dtDatosLinea.Rows(0)("Dto3")
                End If
                data.DatosTarifa.UDValoracion = dtTarifaArticulo.Rows(0)("UdValoracion")
                data.DatosTarifa.IDMoneda = TarInfo.IDMoneda
            End If
        End If

    End Sub

    <Task()> Public Shared Sub AplicarTarifaPadre(ByVal data As DataCalculoTarifaComercial, ByVal services As ServiceProvider)
        If data.DatosTarifa.Precio = 0 Then
            Dim DrArt As DataRow = New Articulo().GetItemRow(data.IDArticulo)
            If DrArt.Table.Columns.Contains("IDArticuloPadre") AndAlso Length(DrArt("IDArticuloPadre")) > 0 Then
                Dim IDArticulo As String = data.IDArticulo
                data.IDArticulo = DrArt("IDArticuloPadre")
                ProcessServer.ExecuteTask(Of DataCalculoTarifaComercial)(AddressOf TarifaComercial, data, services)
                data.IDArticulo = IDArticulo
            End If
        End If
    End Sub

    <Task()> Public Shared Sub PrecioEnMonedaContexto(ByVal data As DataCalculoTarifaComercial, ByVal services As ServiceProvider)
        If Length(data.DatosTarifa.IDMoneda) > 0 AndAlso data.IDMoneda <> data.DatosTarifa.IDMoneda Then
            If data.Fecha Is Nothing Then data.Fecha = Today
            Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
            Dim MonTarifa As MonedaInfo = Monedas.GetMoneda(data.DatosTarifa.IDMoneda, data.Fecha)
            Dim MonContexto As MonedaInfo = Monedas.GetMoneda(data.IDMoneda, data.Fecha)

            If MonContexto.CambioA <> 0 Then
                data.DatosTarifa.Precio = xRound(data.DatosTarifa.Precio * (MonTarifa.CambioA / MonContexto.CambioA), MonContexto.NDecimalesPrecio)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub TratarSeguimiento(ByVal data As DataCalculoTarifaComercial, ByVal services As ServiceProvider)
        Dim strSeguimiento As String
        If Length(data.DatosTarifa.SeguimientoDtos) = 0 Then
            strSeguimiento = "RUTA PARA OBTENER EL PRECIO Y LOS DESCUENTOS: "
            strSeguimiento = strSeguimiento & data.DatosTarifa.SeguimientoTarifa
        Else
            strSeguimiento = "RUTA PARA OBTENER EL PRECIO: "
            strSeguimiento = strSeguimiento & data.DatosTarifa.SeguimientoTarifa & vbNewLine
            If data.DatosTarifa.Dto1 <> 0 OrElse data.DatosTarifa.Dto2 <> 0 OrElse data.DatosTarifa.Dto3 <> 0 Then
                strSeguimiento = strSeguimiento & "RUTA PARA OBTENER LOS DESCUENTOS: "
                strSeguimiento = strSeguimiento & data.DatosTarifa.SeguimientoDtos
            End If
        End If
        data.DatosTarifa.SeguimientoDtos = String.Empty
        data.DatosTarifa.SeguimientoTarifa = strSeguimiento
    End Sub

    <Task()> Public Shared Sub TarifaCosteArticulo(ByVal data As DataCalculoTarifaComercial, ByVal services As ServiceProvider)
        Dim dblPrecioA, dblPrecioB, dblUltimoB As Double
        Dim strWhere, strSelect As String
        Dim dtData As DataTable

        Dim dblStock As Double

        If Not IsNothing(data) AndAlso Length(data.IDArticulo) > 0 Then
            Dim dtArticulo As DataTable
            dtArticulo = New Articulo().SelOnPrimaryKey(data.IDArticulo)
            If Not dtArticulo Is Nothing AndAlso dtArticulo.Rows.Count > 0 Then
                'Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                'Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.IDArticulo)

                If dtArticulo.Rows(0)("CriterioValoracion") = enumtaValoracion.taPrecioFIFOFecha Or _
                   dtArticulo.Rows(0)("CriterioValoracion") = enumtaValoracion.taPrecioFIFOMvto Or _
                   dtArticulo.Rows(0)("CriterioValoracion") = enumtaValoracion.taPrecioMedio Then
                    Dim dtArticuloAlmacen As DataTable
                    If Length(data.IDAlmacen) = 0 Then
                        Dim objFilter As New Filter
                        objFilter.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
                        objFilter.Add(New BooleanFilterItem("Predeterminado", True))
                        dtArticuloAlmacen = New ArticuloAlmacen().Filter(objFilter)
                        If Not dtArticuloAlmacen Is Nothing AndAlso dtArticuloAlmacen.Rows.Count > 0 Then
                            data.IDAlmacen = dtArticuloAlmacen.Rows(0)("IDAlmacen")
                            dblStock = dtArticuloAlmacen.Rows(0)("StockFisico")
                        End If
                    Else
                        dtArticuloAlmacen = New ArticuloAlmacen().SelOnPrimaryKey(data.IDArticulo, data.IDAlmacen)
                        If Not dtArticuloAlmacen Is Nothing AndAlso dtArticuloAlmacen.Rows.Count > 0 Then
                            dblStock = dtArticuloAlmacen.Rows(0)("StockFisico")
                        End If
                    End If
                End If

                Dim f As New Filter
                f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
                f.Add(New StringFilterItem("IDAlmacen", data.IDAlmacen))

                Dim datCalc As New ProcesoStocks.DataCalculoValoracionEnFechaAlmacen(data.Fecha, Nothing, f)
                Dim ViewName As String = ProcessServer.ExecuteTask(Of ProcesoStocks.DataCalculoValoracionEnFechaAlmacen, String)(AddressOf ProcesoStocks.CalculoValoracionEnFechaAlmacen, datCalc, services)
                Dim dtValoracion As DataTable = New BE.DataEngine().Filter(ViewName, f)

                If Not dtValoracion Is Nothing AndAlso dtValoracion.Rows.Count > 0 Then
                    Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
                    Dim MonInfoA As MonedaInfo = Monedas.MonedaA
                    Dim MonInfoB As MonedaInfo = Monedas.MonedaB
                    Select Case dtArticulo.Rows(0)("CriterioValoracion")
                        Case enumtaValoracion.taPrecioEstandar
                            data.DatosTarifa.PrecioCosteA = xRound(Nz(dtValoracion.Rows(0)("PrecioEstandar"), 0), MonInfoA.NDecimalesPrecio) * Nz(data.DatosTarifa.UDValoracion, 1) 'valoracion.PrecioA
                            data.DatosTarifa.PrecioCosteB = xRound(Nz(dtValoracion.Rows(0)("PrecioEstandar") * MonInfoA.CambioB, 0), MonInfoB.NDecimalesPrecio) * Nz(data.DatosTarifa.UDValoracion, 1) 'valoracion.PrecioB
                            If data.DatosTarifa.PrecioCosteA = 0 Then
                                data.DatosTarifa.PrecioCosteA = dtArticulo.Rows(0)("PrecioEstandarA")
                                data.DatosTarifa.PrecioCosteB = dtArticulo.Rows(0)("PrecioEstandarB")
                            End If
                        Case enumtaValoracion.taPrecioFIFOFecha
                            data.DatosTarifa.PrecioCosteA = xRound(Nz(dtValoracion.Rows(0)("FifoFD"), 0), MonInfoA.NDecimalesPrecio) * Nz(data.DatosTarifa.UDValoracion, 1) 'valoracion.PrecioA
                            data.DatosTarifa.PrecioCosteB = xRound(Nz(dtValoracion.Rows(0)("FifoFD") * MonInfoA.CambioB, 0), MonInfoB.NDecimalesPrecio) * Nz(data.DatosTarifa.UDValoracion, 1) 'valoracion.PrecioB
                        Case enumtaValoracion.taPrecioFIFOMvto
                            data.DatosTarifa.PrecioCosteA = xRound(Nz(dtValoracion.Rows(0)("FifoF"), 0), MonInfoA.NDecimalesPrecio) * Nz(data.DatosTarifa.UDValoracion, 1) 'valoracion.PrecioA
                            data.DatosTarifa.PrecioCosteB = xRound(Nz(dtValoracion.Rows(0)("FifoF") * MonInfoA.CambioB, 0), MonInfoB.NDecimalesPrecio) * Nz(data.DatosTarifa.UDValoracion, 1) 'valoracion.PrecioB
                        Case enumtaValoracion.taPrecioMedio
                            data.DatosTarifa.PrecioCosteA = xRound(Nz(dtValoracion.Rows(0)("PrecioMedio"), 0), MonInfoA.NDecimalesPrecio) * Nz(data.DatosTarifa.UDValoracion, 1) 'valoracion.PrecioA
                            data.DatosTarifa.PrecioCosteB = xRound(Nz(dtValoracion.Rows(0)("PrecioMedio") * MonInfoA.CambioB, 0), MonInfoB.NDecimalesPrecio) * Nz(data.DatosTarifa.UDValoracion, 1) 'valoracion.PrecioB
                        Case enumtaValoracion.taPrecioUltCompra
                            data.DatosTarifa.PrecioCosteA = xRound(Nz(dtValoracion.Rows(0)("PrecioUltimaCompra"), 0), MonInfoA.NDecimalesPrecio) * Nz(data.DatosTarifa.UDValoracion, 1) 'valoracion.PrecioA
                            data.DatosTarifa.PrecioCosteB = xRound(Nz(dtValoracion.Rows(0)("PrecioUltimaCompra") * MonInfoA.CambioB, 0), MonInfoB.NDecimalesPrecio) * Nz(data.DatosTarifa.UDValoracion, 1) 'valoracion.PrecioB
                            If data.DatosTarifa.PrecioCosteA = 0 Then
                                data.DatosTarifa.PrecioCosteA = dtArticulo.Rows(0)("PrecioUltimaCompraA")
                                data.DatosTarifa.PrecioCosteB = dtArticulo.Rows(0)("PrecioUltimaCompraB")
                            End If
                    End Select
                Else
                    If data.DatosTarifa.PrecioCosteA = 0 Then
                        data.DatosTarifa.PrecioCosteA = dtArticulo.Rows(0)("PrecioEstandarA")
                        data.DatosTarifa.PrecioCosteB = dtArticulo.Rows(0)("PrecioEstandarB")
                    End If
                End If
            End If
        End If
    End Sub

#End Region

#Region " Cálculo Tarifa Alquiler "

    <Task()> Public Shared Sub TarifaAlquiler(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If data.Current.ContainsKey("IDObra") AndAlso data.Current.ContainsKey("IDArticulo") AndAlso data.Context.ContainsKey("IDCliente") AndAlso data.Current.ContainsKey("Cantidad") Then
            If Length(data.Current("IDObra")) > 0 AndAlso Length(data.Current("IDArticulo")) > 0 AndAlso Length(data.Context("IDCliente")) > 0 AndAlso Nz(data.Current("Cantidad"), 0) <> 0 Then
                Dim FechaFactura As Date
                If data.Context.ContainsKey("FechaFactura") Then
                    FechaFactura = data.Context("FechaFactura")
                End If

                Dim dataTarifa As New CalculoTarifaAlquiler.DataCalculoTarifaAlquiler(data.Current("IDObra"), data.Current("IDArticulo"), _
                                                                                      data.Context("IDCliente"), data.Current("Cantidad"), FechaFactura)
                ProcessServer.ExecuteTask(Of CalculoTarifaAlquiler.DataCalculoTarifaAlquiler)(AddressOf CalculoTarifaAlquiler.TarifaAlquiler, dataTarifa, services)

                If dataTarifa.DatosTarifa.Precio <> 0 Then
                    data.Current("Precio") = dataTarifa.DatosTarifa.Precio
                End If
                data.Current("Dto1") = dataTarifa.DatosTarifa.Dto1
                data.Current("Dto2") = dataTarifa.DatosTarifa.Dto2
                data.Current("Dto3") = dataTarifa.DatosTarifa.Dto3
                data.Current("UdValoracion") = dataTarifa.DatosTarifa.UDValoracion
                data.Current("IDMoneda") = dataTarifa.DatosTarifa.IDMoneda
            End If
        End If
    End Sub

#End Region

#Region "Representantes"

#Region " BusinessRules Representantes "

    <Serializable()> _
    Public Class DataComisionRepresentante
        Public IDRepresentante As String
        Public IDCliente As String
        Public IDArticulo As String
    End Class

    <Serializable()> _
    Public Class DataComision
        Public Comision As Double
        Public Porcentaje As Boolean
    End Class

    <Task()> Public Shared Function RepresentantesCommonBusinessRules(ByVal oBRL As BusinessRules, ByVal services As ServiceProvider) As BusinessRules
        If oBRL Is Nothing Then oBRL = New BusinessRules
        oBRL.Add("IDRepresentante", AddressOf ProcesoComercial.CambioRepresentante)
        oBRL.Add("Importe", AddressOf ProcesoComercial.CambioImporte)
        oBRL.Add("Comision", AddressOf ProcesoComercial.CambioComision)
        oBRL.Add("Porcentaje", AddressOf ProcesoComercial.CambioPorcentaje)
        Return oBRL
    End Function

    <Task()> Public Shared Sub CambioRepresentante(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDRepresentante")) > 0 Then
            ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ValidarRepresentanteExistente, data.Current, services)
            Dim datos As New DataComisionRepresentante
            datos.IDRepresentante = data.Current("IDRepresentante")
            datos.IDCliente = data.Context("IDCliente") & String.Empty
            datos.IDArticulo = data.Context("IDArticulo") & String.Empty
            Dim datosComision As DataComision = ProcessServer.ExecuteTask(Of DataComisionRepresentante, DataComision)(AddressOf GetComisionRepresentante, datos, services)
            data.Current("Comision") = datosComision.Comision
            data.Current("Porcentaje") = datosComision.Porcentaje
            ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CambioPorcentaje, data, services)
        End If
    End Sub

    <Task()> Public Shared Function GetComisionRepresentante(ByVal data As DataComisionRepresentante, ByVal service As ServiceProvider) As DataComision
        Dim comision As New DataComision
        If Len(data.IDRepresentante) > 0 AndAlso Len(data.IDCliente) > 0 Then
            Dim f As New Filter
            f.Add(New StringFilterItem("IDCliente", data.IDCliente))
            f.Add(New StringFilterItem("IDRepresentante", data.IDRepresentante))

            Dim dt As DataTable = New ClienteRepresentante().Filter(f, "IDArticulo DESC")

            If Not IsNothing(dt) AndAlso dt.Rows.Count Then
                Dim ComisionPorDefecto As Boolean = True

                If Len(data.IDArticulo) Then
                    Dim ComisionArticulo As List(Of DataRow) = (From c In dt Where Not c.IsNull("IDArticulo") AndAlso c("IDArticulo") = data.IDArticulo Select c).ToList
                    If Not ComisionArticulo Is Nothing AndAlso ComisionArticulo.Count > 0 Then
                        ComisionPorDefecto = False
                        comision.Comision = Nz(ComisionArticulo(0)("Comision"), 0)
                        comision.Porcentaje = Nz(ComisionArticulo(0)("Porcentaje"), False)
                    End If
                End If

                If ComisionPorDefecto Then
                    Dim ComisionDefecto As List(Of DataRow) = (From c In dt Where c.IsNull("IDArticulo") Select c).ToList
                    If Not ComisionDefecto Is Nothing AndAlso ComisionDefecto.Count > 0 Then
                        comision.Comision = Nz(ComisionDefecto(0)("Comision"), 0)
                        comision.Porcentaje = Nz(ComisionDefecto(0)("Porcentaje"), False)
                    End If
                End If
            End If
        End If

        Return comision
    End Function

    <Task()> Public Shared Sub CambioImporte(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim IncrementoImporte As Double
        If data.ColumnName = "Importe" Then
            If Not IsNumeric(data.Value) Then data.Value = 0
            IncrementoImporte = data.Value - data.Current("Importe")
        End If
        data.Current(data.ColumnName) = data.Value
        If Not IsNumeric(data.Current("Comision")) Then data.Current("Comision") = 0
        If Not IsNumeric(data.Current("Importe")) Then data.Current("Importe") = 0

        '   ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ValidarImporteValido, data.Current, services)
        If data.Context("ImporteLinea") < 0 Then
            If data.Context("SumaImporte") + IncrementoImporte > data.Context("ImporteLinea") Then
                If Nz(data.Context("ImporteLinea"), 0) <> 0 Then
                    data.Current("Comision") = 100 * (data.Current("Importe") / data.Context("ImporteLinea"))
                End If
            Else
                ApplicationService.GenerateError("Los importes asignados a los representantes superan el importe total de la línea.")
            End If
        Else
            If data.Context("SumaImporte") + IncrementoImporte > data.Context("ImporteLinea") Then
                ApplicationService.GenerateError("Los importes asignados a los representantes superan el importe total de la línea.")
            Else
                If Nz(data.Context("ImporteLinea"), 0) <> 0 Then
                    data.Current("Comision") = 100 * (data.Current("Importe") / data.Context("ImporteLinea"))
                End If
            End If
        End If

        If data.Context.ContainsKey("IDMoneda") AndAlso data.Context.ContainsKey("CambioA") AndAlso data.Context.ContainsKey("CambioB") Then
            Dim ValAyB As New ValoresAyB(data.Current, data.Context("IDMoneda"), data.Context("CambioA"), data.Context("CambioB"))
            ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)
        End If
    End Sub

    <Task()> Public Shared Sub CambioComision(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If data.ColumnName = "Comision" Then If Not IsNumeric(data.Value) Then data.Value = 0

        data.Current(data.ColumnName) = data.Value
        If Not IsNumeric(data.Current("Comision")) Then data.Current("Comision") = 0
        If Not IsNumeric(data.Current("Importe")) Then data.Current("Importe") = 0

        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ValidarComisionValida, data.Current, services)
        Dim importe As Double = data.Context("ImporteLinea") * (data.Current("Comision") / 100)
        Dim IncrementoImporte As Double = importe - data.Current("Importe")
        If data.Context("ImporteLinea") < 0 Then
            If data.Context("SumaImporte") + IncrementoImporte > data.Context("ImporteLinea") Then
                data.Current("Importe") = importe
            Else
                ApplicationService.GenerateError("El importe total asignado a los representantes superan el importe total de la línea.")
            End If
        Else
            If data.Context("SumaImporte") + IncrementoImporte > data.Context("ImporteLinea") Then
                ApplicationService.GenerateError("El importe total asignado a los representantes superan el importe total de la línea.")
            Else
                data.Current("Importe") = importe
            End If
        End If

        If data.Context.ContainsKey("IDMoneda") AndAlso data.Context.ContainsKey("CambioA") AndAlso data.Context.ContainsKey("CambioB") Then
            Dim ValAyB As New ValoresAyB(data.Current, data.Context("IDMoneda"), data.Context("CambioA"), data.Context("CambioB"))
            ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)
        End If
    End Sub

    <Task()> Public Shared Sub CambioPorcentaje(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Nz(data.Current("Porcentaje"), False) Then
            ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CambioComision, data, services)
        Else
            ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CambioImporte, data, services)
        End If
    End Sub

#End Region

#Region " Validaciones "

    <Task()> Public Shared Sub ValidarRepresentanteExistente(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider)
        If Length(data("IDRepresentante")) > 0 Then
            Dim DtRepresen As DataTable = New Representante().SelOnPrimaryKey(data("IDRepresentante"))
            If DtRepresen Is Nothing OrElse DtRepresen.Rows.Count = 0 Then
                ApplicationService.GenerateError("El Representante no existe.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarImporteValido(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider)
        '  If Nz(data("Importe"), 0) = 0 Then ApplicationService.GenerateError("El importe no es válido.")
    End Sub

    <Task()> Public Shared Sub ValidarComisionValida(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider)
        Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
        If AppParams.SAAS Then Exit Sub
        If data("Comision") < 0 OrElse data("Comision") > 100 Then ApplicationService.GenerateError("La comisión debe tener un valor entre 0 y 100.")
    End Sub

#End Region

#Region "Creación/Actualización de Representantes"

    <Task()> Public Shared Sub CalcularRepresentantes(ByVal doc As DocumentoComercial, ByVal services As ServiceProvider)
        For Each linea As DataRow In doc.dtLineas.Rows
            If linea.RowState = DataRowState.Added Then
                Dim ctx As New DataDocRow(doc, linea)
                ProcessServer.ExecuteTask(Of DataDocRow)(AddressOf ProcesoComercial.NuevoRepresentante, ctx, services)
            End If
            If linea.RowState = DataRowState.Modified Then
                Dim ctx As New DataDocRow(doc, linea)
                ProcessServer.ExecuteTask(Of DataDocRow)(AddressOf ProcesoComercial.ActualizarRepresentantes, ctx, services)
            End If
        Next
    End Sub

    <Task()> Public Shared Sub NuevoRepresentante(ByVal data As DataDocRow, ByVal services As ServiceProvider)
        If Not IsNothing(data.Row) Then
            If Len(CType(data.Doc, DocumentoComercial).IDCliente) > 0 Then
                If Length(data.Row("IDArticulo")) > 0 Then
                    Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                    Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.Row("IDArticulo"))
                    If ArtInfo.Especial Then Exit Sub
                End If
                Dim PKLinea As String
                Dim newData As DataTable
                Dim MarcaRepresentante As Boolean
                Select Case CType(data.Doc, DocumentCabLin).EntidadLineas
                    Case GetType(FacturaVentaLinea).Name
                        PKLinea = "IDLineaFactura"
                        MarcaRepresentante = True
                    Case GetType(AlbaranVentaLinea).Name
                        PKLinea = "IDLineaAlbaran"
                        If Length(data.Row("TipoLineaAlbaran")) > 0 AndAlso data.Row("TipoLineaAlbaran") = enumavlTipoLineaAlbaran.avlComponente Then Exit Sub
                    Case GetType(PedidoVentaLinea).Name
                        PKLinea = "IDLineaPedido"
                End Select

                newData = CType(data.Doc, DocumentoComercial).dtVentaRepresentante

                Dim Representantes As DataTable
                'Analizar si es el alta de nuevas lineas de albaran con NSerie de nuevo reparto, para coger previamente el reparto de representantes
                'de la linea de origen que se borró.
                If CType(data.Doc, DocumentCabLin).EntidadLineas = GetType(AlbaranVentaLinea).Name AndAlso Length(data.Row("Lote")) > 0 _
                   AndAlso Length(data.Row("IDLineaPedido")) > 0 Then
                    Dim ClsAlbRep As New PedidoVentaRepresentante
                    Dim DtPedRepOrigen As DataTable = ClsAlbRep.Filter(New FilterItem("IDLineaPedido", FilterOperator.Equal, data.Row("IDLineaPedido")))
                    If Not DtPedRepOrigen Is Nothing AndAlso DtPedRepOrigen.Rows.Count > 0 Then
                        Representantes = DtPedRepOrigen.Copy
                    End If
                End If
                If Representantes Is Nothing OrElse Representantes.Rows.Count = 0 Then
                    Representantes = ProcessServer.ExecuteTask(Of DataDocRow, DataTable)(AddressOf CalculoRepresentantes, data, services)
                End If

                If Not IsNothing(Representantes) AndAlso Representantes.Rows.Count > 0 Then
                    Dim strRepresentante As String

                    Dim ComisionesRepresentantes As List(Of DataRow) = (From c In Representantes Order By c("IDRepresentante")).ToList
                    For Each r As DataRow In ComisionesRepresentantes
                        If strRepresentante <> r("IDRepresentante") Then
                            Dim newrow As DataRow = newData.NewRow()

                            strRepresentante = r("IDRepresentante")

                            If MarcaRepresentante Then
                                newrow("MarcaRepresentante") = AdminData.GetAutoNumeric
                            End If

                            newrow(PKLinea) = data.Row(PKLinea)
                            newrow("IDRepresentante") = r("IDRepresentante")

                            Dim comision As Double = Nz(r("Comision"), 0)
                            If Not r("Porcentaje") Then
                                'Calculo el % que supone el fijo
                                If Nz(data.Row("Importe"), 0) <> 0 Then comision = comision * 100 / data.Row("Importe")
                            End If
                            newrow("Comision") = comision
                            newrow("Porcentaje") = r("Porcentaje")
                            newrow("Importe") = (Nz(data.Row("Importe"), 0) * comision) / 100
                            Dim ValAyB As New ValoresAyB(New DataRowPropertyAccessor(newrow), CType(data.Doc, DocumentCabLin).IDMoneda, CType(data.Doc, DocumentCabLin).CambioA, CType(data.Doc, DocumentCabLin).CambioB)
                            ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)

                            newData.Rows.Add(newrow.ItemArray)
                        End If
                    Next
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarRepresentantes(ByVal data As DataDocRow, ByVal services As ServiceProvider)
        Dim blnCambioArticulo, blnCambioImporte, blnCambioCantidad As Boolean
        Dim comision As Double
        Dim Representantes As DataTable = CType(data.Doc, DocumentoComercial).dtVentaRepresentante
        Dim dr As DataRow = data.Row
        If Not IsNothing(dr) Then
            If dr.RowState = DataRowState.Modified Then
                blnCambioArticulo = (dr("IDArticulo") & String.Empty <> dr("IDArticulo", DataRowVersion.Original) & String.Empty)
                blnCambioImporte = (dr("Importe") <> dr("Importe", DataRowVersion.Original))
                blnCambioCantidad = (dr("QInterna") <> dr("QInterna", DataRowVersion.Original))
            Else
                blnCambioArticulo = True
                blnCambioImporte = True
                blnCambioCantidad = True
            End If
            If blnCambioArticulo Or blnCambioImporte Then
                Dim f As New Filter
                Select Case CType(data.Doc, DocumentCabLin).EntidadLineas
                    Case "FacturaVentaLinea"
                        f.Add(New NumberFilterItem("IDLineaFactura", data.Row("IDLineaFactura")))
                    Case "AlbaranVentaLinea"
                        f.Add(New NumberFilterItem("IDLineaAlbaran", data.Row("IDLineaAlbaran")))
                    Case "PedidoVentaLinea"
                        f.Add(New NumberFilterItem("IDLineaPedido", data.Row("IDLineaPedido")))
                End Select

                Dim WhereRepresLinea As String = f.Compose(New AdoFilterComposer)
                Dim RepresLinea() As DataRow = Representantes.Select(WhereRepresLinea)
                If Not IsNothing(RepresLinea) AndAlso RepresLinea.Length > 0 Then
                    If blnCambioImporte And (Not blnCambioArticulo And Not blnCambioCantidad) Then
                        If blnCambioImporte And Not blnCambioArticulo Then

                            If Not IsNothing(data.Doc.HeaderRow) Then

                                For Each lineaRepresentante As DataRow In RepresLinea
                                    comision = lineaRepresentante("Comision")
                                    If Not lineaRepresentante("Porcentaje") Then
                                        If dr("ImporteA") <> 0 Then comision = lineaRepresentante("ImporteA") * 100 / dr("ImporteA")
                                    End If
                                    lineaRepresentante("Comision") = comision
                                    lineaRepresentante("Importe") = (dr("Importe") * comision) / 100

                                    Dim ValAyB As New ValoresAyB(New DataRowPropertyAccessor(lineaRepresentante), CType(data.Doc, DocumentCabLin).IDMoneda, CType(data.Doc, DocumentCabLin).CambioA, CType(data.Doc, DocumentCabLin).CambioB)
                                    ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)
                                Next
                            End If
                        End If
                    Else
                        'Si se modifica el artículo o la cantidad, se elimina el desglose anterior y se vuelve a calcular

                        For Each Representante As DataRow In RepresLinea
                            Representante.Delete()
                        Next
                        ProcessServer.ExecuteTask(Of DataDocRow)(AddressOf NuevoRepresentante, data, services)
                    End If
                Else
                    ProcessServer.ExecuteTask(Of DataDocRow)(AddressOf NuevoRepresentante, data, services)
                End If

            End If
        End If
    End Sub

#End Region

#Region " Calculo Representantes "

    <Task()> Public Shared Function CalculoRepresentantes(ByVal data As DataDocRow, ByVal services As ServiceProvider) As DataTable
        Dim datosCalculo As DataCalculoRepresentantes = ProcessServer.ExecuteTask(Of DataDocRow, DataCalculoRepresentantes)(AddressOf PrepararInformacionCalculoRepresentantes, data, services)
        ProcessServer.ExecuteTask(Of DataCalculoRepresentantes)(AddressOf ComisionesPorObra, datosCalculo, services)
        ProcessServer.ExecuteTask(Of DataCalculoRepresentantes)(AddressOf ComisionesPorCliente, datosCalculo, services)
        ProcessServer.ExecuteTask(Of DataCalculoRepresentantes)(AddressOf ComisionesPorZona, datosCalculo, services)
        ProcessServer.ExecuteTask(Of DataCalculoRepresentantes)(AddressOf ComisionesPorTipoFamilia, datosCalculo, services)
        Return datosCalculo.Representantes
    End Function

    <Task()> Public Shared Function PrepararInformacionCalculoRepresentantes(ByVal data As DataDocRow, ByVal services As ServiceProvider) As DataCalculoRepresentantes
        Dim datos As DataCalculoRepresentantes
        If data.Row.Table.Columns.Contains("IDObra") AndAlso Nz(data.Row("IDObra"), 0) > 0 Then
            datos = New DataCalculoRepresentantes(CType(data.Doc, DocumentoComercial).IDCliente, data.Row("IDArticulo"), Nz(data.Row("QInterna"), 0), data.Row("IDObra"))
        Else
            datos = New DataCalculoRepresentantes(CType(data.Doc, DocumentoComercial).IDCliente, data.Row("IDArticulo"), Nz(data.Row("QInterna"), 0))
        End If
        Return datos
    End Function

#Region " Comisiones por Obra "

    <Task()> Public Shared Sub ComisionesPorObra(ByVal data As DataCalculoRepresentantes, ByVal services As ServiceProvider)
        'TODO
        'Dim f As New Filter
        'If Length(data.IDObra) > 0 AndAlso data.IDObra > 0 Then
        '    f.Add(New NumberFilterItem("IDObra", data.IDObra))
        '    f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
        '    '//OBRA - ARTICULO
        '    Dim ORepresentante As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraRepresentante")
        '    Dim dtObraArtRepr As DataTable = ORepresentante.Filter(f)
        '    If dtObraArtRepr.Rows.Count > 0 Then
        '        data.AddRepresentantes(dtObraArtRepr)
        '    Else
        '        '//OBRA - SUBFAMILIA
        '        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        '        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.IDArticulo)
        '        f.Clear()
        '        f.Add(New NumberFilterItem("IDObra", data.IDObra))
        '        f.Add(New StringFilterItem("IDTipo", ArtInfo.IDTipo))
        '        f.Add(New StringFilterItem("IDFamilia", ArtInfo.IDFamilia))
        '        f.Add(New StringFilterItem("IDSubfamilia", ArtInfo.IDSubFamilia))
        '        Dim dtObraArtSubRepr As DataTable = ORepresentante.Filter(f)
        '        If dtObraArtSubRepr.Rows.Count > 0 Then
        '            data.AddRepresentantes(dtObraArtSubRepr)
        '        Else
        '            '//OBRA - FAMILIA
        '            f.Clear()
        '            f.Add(New NumberFilterItem("IDObra", data.IDObra))
        '            f.Add(New StringFilterItem("IDTipo", ArtInfo.IDTipo))
        '            f.Add(New StringFilterItem("IDFamilia", ArtInfo.IDFamilia))
        '            Dim dtObraArtFamRepr As DataTable = ORepresentante.Filter(f)
        '            If dtObraArtSubRepr.Rows.Count > 0 Then
        '                data.AddRepresentantes(dtObraArtFamRepr)
        '            Else
        '                '//OBRA - TIPO ARTICULO
        '                f.Clear()
        '                f.Add(New NumberFilterItem("IDObra", data.IDObra))
        '                f.Add(New StringFilterItem("IDTipo", ArtInfo.IDTipo))
        '                Dim dtObraTipoRepr As DataTable = ORepresentante.Filter(f)
        '                If dtObraTipoRepr.Rows.Count > 0 Then
        '                    data.AddRepresentantes(dtObraTipoRepr)
        '                Else
        '                    '//OBRA
        '                    f.Clear()
        '                    f.Add(New NumberFilterItem("IDObra", data.IDObra))
        '                    Dim dtObraRepr As DataTable = ORepresentante.Filter(f)
        '                    If dtObraRepr.Rows.Count > 0 Then
        '                        data.AddRepresentantes(dtObraRepr)
        '                    End If
        '                End If
        '            End If
        '        End If
        '    End If
        'End If
    End Sub

#End Region

#Region " Comisiones por Cliente "

    <Task()> Public Shared Sub ComisionesPorCliente(ByVal data As DataCalculoRepresentantes, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataCalculoRepresentantes)(AddressOf ComisionesClienteArticulo, data, services)
        ProcessServer.ExecuteTask(Of DataCalculoRepresentantes)(AddressOf ComisionesCliente, data, services)
        ProcessServer.ExecuteTask(Of DataCalculoRepresentantes)(AddressOf ComisionesClienteFamilia, data, services)
        ProcessServer.ExecuteTask(Of DataCalculoRepresentantes)(AddressOf ComisionesClienteTipo, data, services)
    End Sub

    <Task()> Public Shared Sub ComisionesClienteArticulo(ByVal data As DataCalculoRepresentantes, ByVal services As ServiceProvider)
        If Length(data.IDCliente) > 0 AndAlso Length(data.IDArticulo) > 0 Then
            Dim f As New Filter
            f.Add(New StringFilterItem("IDCliente", data.IDCliente))
            f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
            Dim ClteRepresentante As New ClienteRepresentante
            Dim dtClienteArticulo As DataTable = ClteRepresentante.Filter(f)
            data.AddRepresentantes(dtClienteArticulo)
        End If
    End Sub

    <Task()> Public Shared Sub ComisionesCliente(ByVal data As DataCalculoRepresentantes, ByVal services As ServiceProvider)
        If Length(data.IDCliente) > 0 Then
            Dim f As New Filter
            f.Add(New StringFilterItem("IDCliente", data.IDCliente))
            f.Add(New IsNullFilterItem("IDArticulo"))
            '   f.Add(New NumberFilterItem("Comision", FilterOperator.NotEqual, 0))
            Dim ClteRepresentante As New ClienteRepresentante
            Dim dtCliente As DataTable = ClteRepresentante.Filter(f)
            data.AddRepresentantes(dtCliente)
        End If
    End Sub

    <Task()> Public Shared Sub ComisionesClienteFamilia(ByVal data As DataCalculoRepresentantes, ByVal services As ServiceProvider)
        If Length(data.IDCliente) > 0 AndAlso Length(data.IDArticulo) > 0 Then
            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
            Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.IDArticulo)

            Dim f As New Filter
            f.Add(New StringFilterItem("IDCliente", data.IDCliente))
            f.Add(New StringFilterItem("IDTipo", ArtInfo.IDTipo))
            f.Add(New StringFilterItem("IDFamilia", ArtInfo.IDFamilia))
            f.Add(New NumberFilterItem("QDesde", FilterOperator.LessThanOrEqual, data.Cantidad))

            Dim dtClienteFamilia As DataTable = AdminData.GetData("vNegComisionRepresentante", f, , "IDRepresentante asc, IDTipo asc, IDFamilia asc, QDesde desc")
            data.AddRepresentantes(dtClienteFamilia)
        End If
    End Sub

    <Task()> Public Shared Sub ComisionesClienteTipo(ByVal data As DataCalculoRepresentantes, ByVal services As ServiceProvider)
        If Length(data.IDCliente) > 0 AndAlso Length(data.IDArticulo) > 0 Then
            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
            Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.IDArticulo)

            Dim f As New Filter
            f.Add(New StringFilterItem("IDCliente", data.IDCliente))
            f.Add(New StringFilterItem("IDTipo", ArtInfo.IDTipo))
            f.Add(New NumberFilterItem("QDesde", FilterOperator.LessThanOrEqual, data.Cantidad))

            Dim dtClienteTipo As DataTable = AdminData.GetData("vNegComisionRepresentante", f, , "IDRepresentante asc, IDTipo asc, IDFamilia asc, QDesde desc")
            data.AddRepresentantes(dtClienteTipo)
        End If
    End Sub

#End Region

#Region " Comisiones por Zona "

    <Task()> Public Shared Sub ComisionesPorZona(ByVal data As DataCalculoRepresentantes, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataCalculoRepresentantes)(AddressOf ComisionesZonaFamilia, data, services)
        ProcessServer.ExecuteTask(Of DataCalculoRepresentantes)(AddressOf ComisionesZonaTipo, data, services)
        ProcessServer.ExecuteTask(Of DataCalculoRepresentantes)(AddressOf ComisionesZona, data, services)
    End Sub

    <Task()> Public Shared Sub ComisionesZonaFamilia(ByVal data As DataCalculoRepresentantes, ByVal services As ServiceProvider)
        If Length(data.IDCliente) > 0 AndAlso Length(data.IDArticulo) > 0 Then
            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
            Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.IDArticulo)

            Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
            Dim ClteInfo As ClienteInfo = Clientes.GetEntity(data.IDCliente)

            Dim f As New Filter
            f.Add(New StringFilterItem("IDZona", ClteInfo.Zona))
            f.Add(New StringFilterItem("IDTipo", ArtInfo.IDTipo))
            f.Add(New StringFilterItem("IDFamilia", ArtInfo.IDFamilia))

            Dim ZonaRepres As New ZonaRepresentante
            Dim dtZonaFamilia As DataTable = ZonaRepres.Filter(f)
            data.AddRepresentantes(dtZonaFamilia)
        End If
    End Sub

    <Task()> Public Shared Sub ComisionesZonaTipo(ByVal data As DataCalculoRepresentantes, ByVal services As ServiceProvider)
        If Length(data.IDCliente) > 0 Then
            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
            Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.IDArticulo)

            Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
            Dim ClteInfo As ClienteInfo = Clientes.GetEntity(data.IDCliente)

            Dim f As New Filter
            f.Add(New StringFilterItem("IDZona", ClteInfo.Zona))
            f.Add(New StringFilterItem("IDTipo", ArtInfo.IDTipo))
            f.Add(New IsNullFilterItem("IDFamilia"))
            Dim ZonaRepres As New ZonaRepresentante
            Dim dtZona As DataTable = ZonaRepres.Filter(f)
            data.AddRepresentantes(dtZona)
        End If
    End Sub

    <Task()> Public Shared Sub ComisionesZona(ByVal data As DataCalculoRepresentantes, ByVal services As ServiceProvider)
        If Length(data.IDCliente) > 0 Then
            Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
            Dim ClteInfo As ClienteInfo = Clientes.GetEntity(data.IDCliente)

            Dim f As New Filter
            f.Add(New StringFilterItem("IDZona", ClteInfo.Zona))
            f.Add(New IsNullFilterItem("IDTipo"))
            f.Add(New IsNullFilterItem("IDFamilia"))
            Dim ZonaRepres As New ZonaRepresentante
            Dim dtZona As DataTable = ZonaRepres.Filter(f)
            data.AddRepresentantes(dtZona)
        End If
    End Sub

#End Region

    <Task()> Public Shared Sub ComisionesPorTipoFamilia(ByVal data As DataCalculoRepresentantes, ByVal services As ServiceProvider)
        If Length(data.IDArticulo) > 0 Then
            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
            Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.IDArticulo)

            Dim f As New Filter
            f.Add(New StringFilterItem("IDTipo", ArtInfo.IDTipo))
            f.Add(New StringFilterItem("IDFamilia", ArtInfo.IDFamilia))
            f.Add(New NumberFilterItem("QDesde", FilterOperator.LessThanOrEqual, data.Cantidad))
            Dim dtClienteTipo As DataTable = AdminData.GetData("vNegComisionRepresentanteCantidad", f, , "IDRepresentante asc, IDTipo asc, IDFamilia asc, QDesde desc")
            data.AddRepresentantes(dtClienteTipo)
        End If
    End Sub

#End Region

#Region " Copia de Representantes "

    '//Copia para pasar los representates desde pedido a albarán y de este a la factura.
    <Task()> Public Shared Sub CopiarRepresentantes(ByVal Doc As DocumentoComercial, ByVal services As ServiceProvider)
        Dim PKField As String
        Dim Representantes As DataTable = CType(Doc, DocumentoComercial).dtVentaRepresentante
        Dim blnCopiar As Boolean
        For Each dr As DataRow In Doc.dtLineas.Rows
            blnCopiar = True
            Dim dtRepresentantesOrigen As DataTable : Dim f As New Filter
            Select Case Doc.EntidadLineas
                Case GetType(AlbaranVentaLinea).Name
                    PKField = "IDLineaAlbaran"
                    If Length(dr("IDLineaPedido")) > 0 Then
                        f.Add(New NumberFilterItem("IDLineaPedido", dr("IDLineaPedido")))
                        dtRepresentantesOrigen = New PedidoVentaRepresentante().Filter(f)
                    End If
                    If dr("TipoLineaAlbaran") = enumavlTipoLineaAlbaran.avlComponente Then
                        blnCopiar = False
                    End If
                Case GetType(FacturaVentaLinea).Name
                    PKField = "IDLineaFactura"
                    If Length(dr("IDLineaAlbaran")) > 0 Then
                        f.Add(New NumberFilterItem("IDLineaAlbaran", dr("IDLineaAlbaran")))
                        dtRepresentantesOrigen = New AlbaranVentaRepresentante().Filter(f)
                    End If
            End Select

            If blnCopiar AndAlso Not dtRepresentantesOrigen Is Nothing Then

                For Each drRepresentante As DataRow In dtRepresentantesOrigen.Select
                    Dim drNewRepres As DataRow = Representantes.NewRow
                    drNewRepres(PKField) = dr(PKField)
                    drNewRepres("IDRepresentante") = drRepresentante("IDRepresentante")
                    drNewRepres("Porcentaje") = drRepresentante("Porcentaje")
                    drNewRepres("Comision") = drRepresentante("Comision")
                    Dim dblImporte As Double = 0
                    If drNewRepres("porcentaje") Then '//Comisión por Porcentaje
                        If dr("Importe") <> 0 Then
                            dblImporte = (dr("Importe") * drNewRepres("Comision")) / 100
                        End If
                    Else                                     '//Comisión de Importe fijo
                        dblImporte = drRepresentante("Importe")
                    End If
                    drNewRepres("Importe") = dblImporte

                    If GetType(FacturaVentaRepresentante).Name = Doc.EntidadRepresentantes Then
                        drNewRepres("MarcaRepresentante") = AdminData.GetAutoNumeric
                    End If

                    Dim ValAyB As New ValoresAyB(New DataRowPropertyAccessor(drNewRepres), Doc.IDMoneda, Doc.CambioA, Doc.CambioB)
                    ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)

                    Representantes.Rows.Add(drNewRepres)
                Next
            End If
        Next
    End Sub

    'Public Sub CopiarRepresentantes(ByVal representantesOrigen As DataTable, ByVal representantesDestino As DataTable, ByVal intIDLineaDestino As Integer, ByVal dblImporte As Double)
    '    If Not IsNothing(representantesOrigen) And Not IsNothing(representantesDestino) Then
    '        If representantesOrigen.Rows.Count Then

    '            Dim pkOrigen, pkDestino As String

    '            If representantesOrigen.TableName = GetType(PedidoVentaRepresentante).Name And representantesDestino.TableName = GetType(AlbaranVentaRepresentante).Name Then
    '                pkOrigen = "IDLineaPedido"
    '                pkDestino = "IDLineaAlbaran"
    '            ElseIf representantesOrigen.TableName = GetType(AlbaranVentaRepresentante).Name And representantesDestino.TableName = GetType(FacturaVentaRepresentante).Name Then
    '                pkOrigen = "IDLineaAlbaran"
    '                pkDestino = "IDLineaFactura"
    '                'ElseIf representantesOrigen.TableName = "PedidoVentaRepresentante" And representantesDestino.TableName = "PedidoVentaAnalitica" Then
    '                '    pkOrigen = "IDLineaPedido"
    '                '    pkDestino = "IDLineaPedido"
    '                'ElseIf representantesOrigen.TableName = "AlbaranVentaRepresentante" And representantesDestino.TableName = "AlbaranVentaRepresentante" Then
    '                '    pkOrigen = "IDLineaAlbaran"
    '                '    pkDestino = "IDLineaAlbaran"
    '                'ElseIf representantesOrigen.TableName = "FacturaVentaRepresentante" And representantesDestino.TableName = "FacturaVentaRepresentante" Then
    '                '    pkOrigen = "IDLineaFactura"
    '                '    pkDestino = "IDLineaFactura"
    '            End If

    '            For Each origen As DataRow In representantesOrigen.Rows
    '                Dim destino As DataRow = representantesDestino.NewRow
    '                For Each dc As DataColumn In representantesOrigen.Columns
    '                    If Not (dc.ColumnName = pkOrigen And pkOrigen <> pkDestino) Then
    '                        If dc.ColumnName <> pkDestino Then
    '                            destino(dc.ColumnName) = origen(dc)
    '                            If dc.ColumnName = "Importe" Then
    '                                destino(dc.ColumnName) = (destino(dc.ColumnName) * origen("Porcentaje")) / 100
    '                            End If
    '                        End If
    '                    End If
    '                Next
    '                destino(pkDestino) = intIDLineaDestino
    '                If representantesDestino.TableName = "FacturaVentaRepresentante" Then
    '                    destino("MarcaRepresentante") = AdminData.GetAutoNumeric
    '                End If
    '                representantesDestino.Rows.Add(destino)
    '            Next
    '        End If
    '    End If
    'End Sub

#End Region

#End Region

#Region " Métodos comunes a los procesos de creación de documentos del circuito comercial "


    <Task()> Public Shared Sub AsignarDatosCliente(ByVal Doc As DocumentoComercial, ByVal services As ServiceProvider)
        If Doc.Cliente Is Nothing Then
            Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
            Doc.Cliente = Clientes.GetEntity(Doc.HeaderRow("IDCliente"))
        End If

        If Length(Doc.HeaderRow("IDFormaPago")) = 0 Then Doc.HeaderRow("IDFormaPago") = Doc.Cliente.FormaPago
        If Length(Doc.HeaderRow("IDCondicionPago")) = 0 Then Doc.HeaderRow("IDCondicionPago") = Doc.Cliente.CondicionPago
        If Length(Doc.HeaderRow("IDCondicionPago")) > 0 Then
            If Doc.HeaderRow.Table.Columns.Contains("DtoProntoPago") Then
                Dim CondicionesPago As EntityInfoCache(Of CondicionPagoInfo) = services.GetService(Of EntityInfoCache(Of CondicionPagoInfo))()
                Dim CondPagoInfo As CondicionPagoInfo = CondicionesPago.GetEntity(Doc.HeaderRow("IDCondicionPago"))
                Doc.HeaderRow("DtoProntoPago") = CondPagoInfo.DtoProntoPago
            End If
        End If
        'If length(Doc.HeaderRow("IDDiaPago"))=0 Then Doc.HeaderRow("IDDiaPago") = Doc.Cliente.DiaPago
        If Length(Doc.HeaderRow("IDMoneda")) = 0 Then Doc.HeaderRow("IDMoneda") = Doc.Cliente.Moneda
    End Sub

    <Task()> Public Shared Sub AsignarObservacionesComercial(ByVal Doc As DocumentoComercial, ByVal services As ServiceProvider)
        Dim CampoObsComerciales As String
        Select Case Doc.EntidadCabecera
            Case GetType(PedidoVentaCabecera).Name
                CampoObsComerciales = "TextoComercial"
            Case Else
                CampoObsComerciales = "Texto"
        End Select
        If Not Doc.Cabecera Is Nothing AndAlso Length(Doc.Cabecera.ObsComerciales) > 0 Then Doc.HeaderRow(CampoObsComerciales) = Doc.Cabecera.ObsComerciales
        Dim obs As New DataObservaciones(Doc.EntidadCabecera, CampoObsComerciales, New DataRowPropertyAccessor(Doc.HeaderRow))
        ProcessServer.ExecuteTask(Of DataObservaciones)(AddressOf ProcesoComercial.AsignarObservacionesCliente, obs, services)
    End Sub

    <Task()> Public Shared Sub AsignarEjercicio(ByVal Doc As DocumentoComercial, ByVal services As ServiceProvider)
        Dim AppParamsConta As ParametroContabilidadVenta = services.GetService(Of ParametroContabilidadVenta)()
        If Not AppParamsConta.Contabilidad Then Exit Sub
        Dim DE As New DataEjercicio(New DataRowPropertyAccessor(Doc.HeaderRow), Doc.Fecha)
        ProcessServer.ExecuteTask(Of DataEjercicio)(AddressOf NegocioGeneral.AsignarEjercicioContable, DE, services)
    End Sub

    <Task()> Public Shared Sub AsignarCondicionesEnvio(ByVal Doc As DocumentoComercial, ByVal services As ServiceProvider)
        If Not Doc.Cabecera Is Nothing Then
            If Doc.HeaderRow.Table.Columns.Contains("IDFormaEnvio") AndAlso Length(Doc.Cabecera.IDFormaEnvio) > 0 Then
                Doc.HeaderRow("IDFormaEnvio") = Doc.Cabecera.IDFormaEnvio
            End If
            If Doc.HeaderRow.Table.Columns.Contains("IDCondicionEnvio") AndAlso Length(Doc.Cabecera.IDCondicionEnvio) > 0 Then
                Doc.HeaderRow("IDCondicionEnvio") = Doc.Cabecera.IDCondicionEnvio
            End If
            If Doc.HeaderRow.Table.Columns.Contains("IDModoTransporte") AndAlso Length(Doc.Cabecera.IDModoTransporte) > 0 Then
                Doc.HeaderRow("IDModoTransporte") = Doc.Cabecera.IDModoTransporte
            End If
        End If
    End Sub

#End Region

#Region " Corregir Movimientos "

    '<Task()> Public Shared Function CorregirMovimientos(ByVal Doc As Document, ByVal services As ServiceProvider) As StockUpdateData()
    '    Dim returnData(-1) As StockUpdateData

    '    ProcessServer.ExecuteTask(Of Document)(AddressOf CorreccionMovimientosCambiosCabecera, Doc, services)
    '    ProcessServer.ExecuteTask(Of Document)(AddressOf CorreccionMovimientosCambiosLineas, Doc, services)

    '    Return returnData
    'End Function

    '<Task()> Public Shared Sub CorreccionMovimientosCambiosCabecera(ByVal Doc As Document, ByVal services As ServiceProvider)
    '    Dim returnData(-1) As StockUpdateData
    '    If Doc.HeaderRow.RowState = DataRowState.Modified Then
    '        If Doc.HeaderRow("FechaAlbaran", DataRowVersion.Original) <> Doc.HeaderRow("FechaAlbaran") Then
    '            If Doc.HeaderRow("FechaAlbaran") <> DateTime.MinValue Then
    '                Dim FechaDocumento As Date = Doc.HeaderRow("FechaAlbaran")
    '                Dim f As New Filter
    '                f.Add(New NumberFilterItem("IDAlbaran", Doc.HeaderRow("IDAlbaran")))
    '                f.Add(New NumberFilterItem("EstadoStock", enumavlEstadoStock.avlActualizado))

    '                Dim dtLineasAlbaran As DataTable
    '                If TypeOf Doc Is DocumentoAlbaranVenta Then
    '                    dtLineasAlbaran = CType(Doc, DocumentoAlbaranVenta).dtLineas
    '                    'ElseIf TypeOf Doc Is DocumentoAlbaranCompra Then
    '                    'dtLineasAlbaran = CType(Doc, DocumentoAlbaranCompra).dtCompraLineas
    '                End If
    '                Dim WhereStockActualizado As String = f.Compose(New AdoFilterComposer)
    '                Dim lineasAlbaran() As DataRow = dtLineasAlbaran.Select(WhereStockActualizado)
    '                If Not lineasAlbaran Is Nothing Then
    '                    For Each lineaAlbaran As DataRow In lineasAlbaran
    '                        f.Clear()
    '                        f.Add(New NumberFilterItem("IDLineaAlbaran", lineaAlbaran("IDLineaAlbaran")))
    '                        Dim dtLotes As DataTable
    '                        If TypeOf Doc Is DocumentoAlbaranVenta Then
    '                            dtLotes = CType(Doc, DocumentoAlbaranVenta).dtLote
    '                            'ElseIf TypeOf Doc Is DocumentoAlbaranCompra Then
    '                            'dtLotes = CType(Doc, DocumentoAlbaranCompra).dtLote
    '                        End If
    '                        Dim WhereLotesLinea As String = f.Compose(New AdoFilterComposer)
    '                        Dim lotes() As DataRow = dtLotes.Select(WhereLotesLinea)
    '                        If lotes.Length > 0 Then
    '                            For Each lote As DataRow In lotes
    '                                Dim IDLineaMovimiento As Integer = 0
    '                                '//Movimiento de salida
    '                                If IsNumeric(lote("IDMovimientoSalida")) Then IDLineaMovimiento = lote("IDMovimientoSalida")

    '                                '//Movimiento de entrada (si existe)
    '                                If IsNumeric(lote("IDMovimientoEntrada")) Then IDLineaMovimiento = lote("IDMovimientoEntrada")

    '                                If IDLineaMovimiento <> 0 Then
    '                                    Dim datCorrMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, IDLineaMovimiento, FechaDocumento)
    '                                    Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datCorrMovto, services)
    '                                    ArrayManager.Copy(updateData, returnData)
    '                                End If
    '                            Next
    '                        Else
    '                            Dim IDLineaMovimiento As Integer = 0
    '                            '//Movimiento de salida
    '                            If IsNumeric(lineaAlbaran("IDMovimiento")) Then IDLineaMovimiento = lineaAlbaran("IDMovimiento")
    '                            '//Movimiento de entrada (si existe)
    '                            If IsNumeric(lineaAlbaran("IDMovimientoEntrada")) Then IDLineaMovimiento = lineaAlbaran("IDMovimientoEntrada")

    '                            Dim datCorrMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, IDLineaMovimiento, FechaDocumento)
    '                            Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datCorrMovto, services)
    '                            If updateData.Estado = EstadoStock.Actualizado Then
    '                                ArrayManager.Copy(updateData, returnData)
    '                            Else : Exit Sub
    '                            End If
    '                        End If
    '                    Next
    '                End If
    '            End If
    '        End If
    '    End If
    'End Sub

    '<Task()> Public Shared Sub CorreccionMovimientosCambiosLineas(ByVal Doc As Document, ByVal services As ServiceProvider)
    '    Dim dtLineasAlbaran As DataTable
    '    If TypeOf Doc Is DocumentoAlbaranVenta Then
    '        dtLineasAlbaran = CType(Doc, DocumentoAlbaranVenta).dtLineas
    '        'ElseIf TypeOf Doc Is DocumentoAlbaranCompra Then
    '        'dtLineasAlbaran = CType(Doc, DocumentoAlbaranCompra).dtCompraLineas
    '    End If
    '    Dim f As New Filter
    '    f.Add(New NumberFilterItem("TipoLineaAlbaran", FilterOperator.NotEqual, enumavlTipoLineaAlbaran.avlComponente))
    '    Dim WhereNotComponentes As String = f.Compose(New AdoFilterComposer)
    '    For Each lineaAlbaran As DataRow In dtLineasAlbaran.Select(WhereNotComponentes)
    '        If lineaAlbaran.RowState = DataRowState.Modified Then
    '            If lineaAlbaran("QServida", DataRowVersion.Original) <> lineaAlbaran("QServida") OrElse _
    '               lineaAlbaran("QInterna", DataRowVersion.Original) <> lineaAlbaran("QInterna") OrElse _
    '               lineaAlbaran("ImporteA", DataRowVersion.Original) <> lineaAlbaran("ImporteA") OrElse _
    '               lineaAlbaran("ImporteB", DataRowVersion.Original) <> lineaAlbaran("ImporteB") OrElse _
    '               lineaAlbaran("Precio", DataRowVersion.Original) <> lineaAlbaran("Precio") OrElse _
    '               lineaAlbaran("QEtiContenedor", DataRowVersion.Original) <> lineaAlbaran("QEtiContenedor") Then

    '                If lineaAlbaran("Precio") <> lineaAlbaran("Precio", DataRowVersion.Original) AndAlso _
    '                   lineaAlbaran("QServida") = lineaAlbaran("QServida", DataRowVersion.Original) AndAlso _
    '                   lineaAlbaran("QInterna") = lineaAlbaran("QInterna", DataRowVersion.Original) AndAlso _
    '                   lineaAlbaran("QEtiContenedor") = lineaAlbaran("QEtiContenedor", DataRowVersion.Original) Then

    '                    Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
    '                    Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(lineaAlbaran("IDArticulo"))
    '                    If ArtInfo.RecalcularValoracion = CInt(enumtaValoracionSalidas.taMantenerPrecio) Then
    '                        Dim ctx As New DataDocRow(Doc, lineaAlbaran)
    '                        Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of DataDocRow, StockUpdateData)(AddressOf CorregirMovimiento, ctx, services)
    '                        If updateData Is Nothing Then
    '                            If lineaAlbaran("EstadoStock") = EstadoStock.Actualizado OrElse lineaAlbaran("EstadoStock") = EstadoStock.SinGestion Then
    '                                Dim fComponentes As New Filter
    '                                fComponentes.Add(New NumberFilterItem("TipoLineaAlbaran", enumavlTipoLineaAlbaran.avlComponente))
    '                                fComponentes.Add(New NumberFilterItem("IDLineaPadre", lineaAlbaran("IDLineaAlbaran")))
    '                                Dim WhereComponentes As String = fComponentes.Compose(New AdoFilterComposer)
    '                                For Each componente As DataRow In dtLineasAlbaran.Select(WhereComponentes)
    '                                    ctx = New DataDocRow(Doc, componente)
    '                                    ProcessServer.ExecuteTask(Of DataDocRow, StockUpdateData)(AddressOf CorregirMovimiento, ctx, services)
    '                                Next
    '                            End If
    '                        ElseIf updateData.Estado = EstadoStock.NoActualizado Then
    '                            Throw New Exception(updateData.Detalle)
    '                        End If
    '                    End If
    '                Else
    '                    Dim ctx As New DataDocRow(Doc, lineaAlbaran)
    '                    Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of DataDocRow, StockUpdateData)(AddressOf CorregirMovimiento, ctx, services)
    '                    If updateData Is Nothing Then
    '                        If lineaAlbaran("EstadoStock") = EstadoStock.Actualizado Or lineaAlbaran("EstadoStock") = EstadoStock.SinGestion Then
    '                            Dim fComponentes As New Filter
    '                            fComponentes.Add(New NumberFilterItem("TipoLineaAlbaran", enumavlTipoLineaAlbaran.avlComponente))
    '                            fComponentes.Add(New NumberFilterItem("IDLineaPadre", lineaAlbaran("IDLineaAlbaran")))
    '                            Dim WhereComponentes As String = fComponentes.Compose(New AdoFilterComposer)
    '                            For Each componente As DataRow In dtLineasAlbaran.Select(WhereComponentes)
    '                                ctx = New DataDocRow(Doc, componente)
    '                                ProcessServer.ExecuteTask(Of DataDocRow, StockUpdateData)(AddressOf CorregirMovimiento, ctx, services)
    '                            Next
    '                        End If
    '                    ElseIf updateData.Estado = EstadoStock.NoActualizado Then
    '                        Throw New Exception(updateData.Detalle)
    '                    End If

    '                End If
    '            End If
    '        End If
    '    Next
    'End Sub

    '<Task()> Public Shared Function CorregirMovimiento(ByVal ctx As DataDocRow, ByVal services As ServiceProvider) As StockUpdateData
    '    Dim Cantidad As Double : Dim updateData As StockUpdateData
    '    Dim lineaAlbaran As DataRow = ctx.Row

    '    AdminData.BeginTx()
    '    If lineaAlbaran("QEtiContenedor", DataRowVersion.Original) <> lineaAlbaran("QEtiContenedor") Then
    '        Cantidad = Nz(lineaAlbaran("QEtiContenedor"), 0)
    '        If IsNumeric(lineaAlbaran("IDSalidaContenedor")) Then
    '            '//Correccion movimiento de salida de contenedor
    '            Dim datCorrMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, lineaAlbaran("IDSalidaContenedor"), Cantidad)
    '            updateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datCorrMovto, services)
    '            If Not updateData Is Nothing AndAlso updateData.Estado = EstadoStock.Actualizado Then
    '                lineaAlbaran("IDSalidaContenedor") = updateData.IDLineaMovimiento
    '            Else
    '                AdminData.RollBackTx()
    '                Return updateData
    '            End If
    '        End If

    '        If IsNumeric(lineaAlbaran("IDEntradaContenedor")) Then
    '            '//Correccion movimiento de entrada de contenedor
    '            Dim datCorrMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, lineaAlbaran("IDEntradaContenedor"), Cantidad)
    '            updateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datCorrMovto, services)
    '            If Not updateData Is Nothing AndAlso updateData.Estado = EstadoStock.Actualizado Then
    '                lineaAlbaran("IDEntradaContenedor") = updateData.IDLineaMovimiento
    '            Else
    '                AdminData.RollBackTx()
    '                Return updateData
    '            End If
    '        End If
    '    End If


    '    Dim PrecioA As Double : Dim PrecioB As Double
    '    Cantidad = lineaAlbaran("QInterna")
    '    If (lineaAlbaran("Factor") <> 0 And lineaAlbaran("UdValoracion") <> 0) Then
    '        Dim m As New Moneda
    '        Dim monedaA As MonedaInfo = ProcessServer.ExecuteTask(Of Date, MonedaInfo)(AddressOf Moneda.MonedaA, cnMinDate, services)
    '        Dim monedaB As MonedaInfo = ProcessServer.ExecuteTask(Of Date, MonedaInfo)(AddressOf Moneda.MonedaB, cnMinDate, services)
    '        PrecioA = xRound(lineaAlbaran("PrecioA") / lineaAlbaran("Factor") / lineaAlbaran("UdValoracion") * (1 - lineaAlbaran("Dto1") / 100) * (1 - lineaAlbaran("Dto2") / 100) * (1 - lineaAlbaran("Dto3") / 100) * (1 - lineaAlbaran("Dto") / 100) * (1 - lineaAlbaran("DtoProntoPago") / 100), monedaA.NDecimalesPrecio)
    '        PrecioB = xRound(lineaAlbaran("PrecioB") / lineaAlbaran("Factor") / lineaAlbaran("UdValoracion") * (1 - lineaAlbaran("Dto1") / 100) * (1 - lineaAlbaran("Dto2") / 100) * (1 - lineaAlbaran("Dto3") / 100) * (1 - lineaAlbaran("Dto") / 100) * (1 - lineaAlbaran("DtoProntoPago") / 100), monedaB.NDecimalesPrecio)
    '    End If


    '    Dim dtLoteDoc As DataTable
    '    If TypeOf ctx.Doc Is DocumentoAlbaranVenta Then
    '        dtLoteDoc = CType(ctx.Doc, DocumentoAlbaranVenta).dtLote
    '        'ElseIf TypeOf ctx.Doc Is DocumentoAlbaranCompra Then
    '        'dtLoteDoc = CType(ctx.Doc, DocumentoAlbaranCompra).dtLote
    '    End If

    '    Dim f As New Filter
    '    f.Add(New NumberFilterItem("IDLineaAlbaran", lineaAlbaran("IDLineaAlbaran")))
    '    Dim WhereLotesLinea As String = f.Compose(New AdoFilterComposer)
    '    Dim lote() As DataRow = dtLoteDoc.Select(WhereLotesLinea)
    '    If lote.Length > 0 Then
    '        '//Corregir todos los movimientos asociados a los lotes (solo se corrigen si hay cambio en precio-importe)
    '        If lineaAlbaran("ImporteA", DataRowVersion.Original) <> lineaAlbaran("ImporteA") _
    '        Or lineaAlbaran("ImporteB", DataRowVersion.Original) <> lineaAlbaran("ImporteB") Then
    '            For Each dr As DataRow In lote
    '                If Not dr.IsNull("IDMovimientoSalida") Then
    '                    '//Correccion movimiento de salida
    '                    Dim datCorrMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, dr("IDMovimientoSalida"), PrecioA, PrecioB)
    '                    updateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datCorrMovto, services)
    '                    If updateData Is Nothing OrElse updateData.Estado <> EstadoStock.Actualizado Then
    '                        AdminData.RollBackTx()
    '                        Return updateData
    '                    End If
    '                End If

    '                If Not dr.IsNull("IDMovimientoEntrada") Then
    '                    '//Correccion movimiento de entrada
    '                    Dim datCorrMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, dr("IDMovimientoEntrada"), PrecioA, PrecioB)
    '                    updateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datCorrMovto, services)
    '                    If updateData Is Nothing OrElse updateData.Estado <> EstadoStock.Actualizado Then
    '                        AdminData.RollBackTx()
    '                        Return updateData
    '                    End If
    '                End If
    '            Next
    '        End If
    '    Else
    '        If Not lineaAlbaran.IsNull("IDMovimiento") Then
    '            '//Correccion movimiento de salida
    '            Dim datCorrMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, lineaAlbaran("IDMovimiento"), PrecioA, PrecioB)
    '            updateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datCorrMovto, services)
    '            If Not updateData Is Nothing AndAlso updateData.Estado = EstadoStock.Actualizado Then
    '                lineaAlbaran("IDMovimiento") = updateData.IDLineaMovimiento
    '            Else
    '                AdminData.RollBackTx()
    '                Return updateData
    '            End If
    '        End If

    '        If Not lineaAlbaran.IsNull("IDMovimientoEntrada") Then
    '            '//Correccion movimiento de entrada
    '            Dim datCorrMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, lineaAlbaran("IDMovimientoEntrada"), PrecioA, PrecioB)
    '            updateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datCorrMovto, services)

    '            If Not updateData Is Nothing AndAlso updateData.Estado = EstadoStock.Actualizado Then
    '                lineaAlbaran("IDMovimientoEntrada") = updateData.IDLineaMovimiento
    '            Else
    '                AdminData.RollBackTx()
    '                Return updateData
    '            End If
    '        End If
    '    End If

    '    AdminData.CommitTx()
    '    Return updateData
    'End Function


#End Region

End Class

