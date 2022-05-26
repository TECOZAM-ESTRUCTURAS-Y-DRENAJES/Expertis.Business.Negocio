Public Class CalculoTarifaAlquiler

    <Serializable()> _
    Public Class DataCalculoTarifaAlquiler
        Public IDObra As Integer
        Public IDArticulo, IDCliente As String
        Public IDMoneda As String    ' Moneda del contexto (dato de entrada). Se utiliza entre otras cosas, para devolver la Tarifa en esta Moneda
        Public Cantidad As Double
        Public Fecha As Date
        Public DatosTarifa As dataTarifaAlquiler
        Friend ArticuloSinDto As Boolean
        Friend SeguimientoDtoTarifa As String
        Friend Final As Boolean = False
        Friend Dtos As Boolean = False

        Public Sub New(ByVal IDObra As Integer, ByVal IDMaterial As String, ByVal IDCliente As String, ByVal Cantidad As Double, ByVal Fecha As Date)
            Me.DatosTarifa = New dataTarifaAlquiler
            Me.DatosTarifa.UDValoracion = 1
            Me.IDObra = IDObra
            Me.IDArticulo = IDMaterial
            Me.IDCliente = IDCliente
            Me.Cantidad = Cantidad
            Me.Fecha = Fecha
        End Sub
    End Class

    <Task()> Public Shared Sub TarifaAlquiler(ByVal data As DataCalculoTarifaAlquiler, ByVal services As ServiceProvider)
        If Length(data.IDArticulo) > 0 And Length(data.IDCliente) > 0 And data.Cantidad <> 0 Then
            ProcessServer.ExecuteTask(Of DataCalculoTarifaAlquiler)(AddressOf TarifaArticuloSinDto, data, services)
            ProcessServer.ExecuteTask(Of DataCalculoTarifaAlquiler)(AddressOf TarifaContratoAlquiler, data, services)
            If Not data.Final Then
                ProcessServer.ExecuteTask(Of DataCalculoTarifaAlquiler)(AddressOf DtosContratoAlquiler, data, services)
                ProcessServer.ExecuteTask(Of DataCalculoTarifaAlquiler)(AddressOf TarifaDelArticulo, data, services)
                ProcessServer.ExecuteTask(Of DataCalculoTarifaAlquiler)(AddressOf TarifaCircuitoComercial, data, services)
            End If
            ProcessServer.ExecuteTask(Of DataCalculoTarifaAlquiler)(AddressOf PrecioEnMonedaContexto, data, services)
        End If
    End Sub

    <Task()> Public Shared Sub TarifaArticuloSinDto(ByVal data As DataCalculoTarifaAlquiler, ByVal services As ServiceProvider)
        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        Dim Articulo As ArticuloInfo = Articulos.GetEntity(data.IDArticulo)

        data.ArticuloSinDto = Articulo.SinDtoEnAlquiler
        If data.ArticuloSinDto Then data.SeguimientoDtoTarifa = "RUTA PARA OBTENER LOS DTOS: ARTÍCULO SIN DTOS"
    End Sub

    <Task()> Public Shared Sub TarifaContratoAlquiler(ByVal data As DataCalculoTarifaAlquiler, ByVal services As ServiceProvider)
        Dim f As New Filter
        f.Add(New NumberFilterItem("IDObra", data.IDObra))
        f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
        f.Add(New NumberFilterItem("Tipo", otaTipo.otaPrecioArticulo))
        Dim dt As DataTable = New BE.DataEngine().Filter("tbObraTarifaAlquiler", f, "Precio,Dto1,Dto2,Dto3,UDValoracion")
        'Primero se busca si se está el artículo en el contrato del Alquiler
        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            data.DatosTarifa.Precio = dt.Rows(0)("Precio")
            data.DatosTarifa.Dto1 = IIf(data.ArticuloSinDto, 0, dt.Rows(0)("Dto1"))
            data.DatosTarifa.Dto2 = IIf(data.ArticuloSinDto, 0, dt.Rows(0)("Dto2"))
            data.DatosTarifa.Dto3 = IIf(data.ArticuloSinDto, 0, dt.Rows(0)("Dto3"))
            data.DatosTarifa.UDValoracion = IIf(dt.Rows(0)("UDValoracion") > 0, dt.Rows(0)("UDValoracion"), 1)
            data.SeguimientoDtoTarifa = "RUTA PARA OBTENER LOS DTOS: DTOS DEL ARTICULO DE LAS CONDICIONES DE ALQUILER"
            data.DatosTarifa.SeguimientoTarifa = "RUTA PARA OBTENER EL PRECIO: TARIFA DEL ARTICULO DE LAS CONDICIONES DE ALQUILER" & vbNewLine & data.SeguimientoDtoTarifa
            data.Final = True
        End If
    End Sub

#Region " DtosContratoAlquiler "

    <Task()> Public Shared Sub DtosContratoAlquiler(ByVal data As DataCalculoTarifaAlquiler, ByVal services As ServiceProvider)
        'En el caso de no encontrar el artículo en el contrato de alquiler, se buscan sus descuentos teniendo en cuenta el siguiente orden de prioridad:
        '   1.- Descuentos por Tipo-Familia
        '   2.- Descuentos por Tipo
        '   3.- Descuentos generales
        '   4.- Descuesto línea del Cliente (DtoComercialLinea)
        ProcessServer.ExecuteTask(Of DataCalculoTarifaAlquiler)(AddressOf DtosContratoAlquilerPorTipoFamilia, data, services)
        If Not data.Dtos Then
            ProcessServer.ExecuteTask(Of DataCalculoTarifaAlquiler)(AddressOf DtosContratoAlquilerPorTipo, data, services)
            If Not data.Dtos Then
                ProcessServer.ExecuteTask(Of DataCalculoTarifaAlquiler)(AddressOf DtosContratoAlquilerGenerales, data, services)
                If Not data.Dtos Then
                    ProcessServer.ExecuteTask(Of DataCalculoTarifaAlquiler)(AddressOf DtoComericalCliente, data, services)
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub DtosContratoAlquilerPorTipoFamilia(ByVal data As DataCalculoTarifaAlquiler, ByVal services As ServiceProvider)
        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        Dim Articulo As ArticuloInfo = Articulos.GetEntity(data.IDArticulo)

        Dim f As New Filter
        f.Add(New NumberFilterItem("IDObra", data.IDObra))
        f.Add(New NumberFilterItem("Tipo", otaTipo.otaDescuentoPorFamilia))
        f.Add(New NumberFilterItem("TipoDato", otaTipoDtos.otaDtoPorTipoFamilia))
        f.Add(New StringFilterItem("IDTipo", Articulo.IDTipo))
        f.Add(New StringFilterItem("IDFamilia", Articulo.IDFamilia))
        Dim dt As DataTable = New BE.DataEngine().Filter("tbObraTarifaAlquiler", f, "Dto1,Dto2,Dto3")
        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            If Not data.ArticuloSinDto Then
                data.DatosTarifa.Dto1 = dt.Rows(0)("Dto1")
                data.DatosTarifa.Dto2 = dt.Rows(0)("Dto2")
                data.DatosTarifa.Dto3 = dt.Rows(0)("Dto3")
            End If
            data.DatosTarifa.UDValoracion = IIf(Articulo.UDValoracion > 0, Articulo.UDValoracion, 1)
            data.SeguimientoDtoTarifa = "RUTA PARA OBTENER LOS DTOS: DTOS POR TIPO-FAMILIA DE LAS CONDICIONES DE ALQUILER"
            data.Dtos = True
        End If
    End Sub

    <Task()> Public Shared Sub DtosContratoAlquilerPorTipo(ByVal data As DataCalculoTarifaAlquiler, ByVal services As ServiceProvider)
        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        Dim Articulo As ArticuloInfo = Articulos.GetEntity(data.IDArticulo)

        Dim f As New Filter
        f.Add(New NumberFilterItem("IDObra", data.IDObra))
        f.Add(New NumberFilterItem("Tipo", otaTipo.otaDescuentoPorFamilia))
        f.Add(New NumberFilterItem("TipoDato", otaTipoDtos.otaDtoPorTipo))
        f.Add(New StringFilterItem("IDTipo", Articulo.IDTipo))
        f.Add(New IsNullFilterItem("IDFamilia", True))
        Dim dt As DataTable = New BE.DataEngine().Filter("tbObraTarifaAlquiler", f, "Dto1,Dto2,Dto3")
        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            If Not data.ArticuloSinDto Then
                data.DatosTarifa.Dto1 = dt.Rows(0)("Dto1")
                data.DatosTarifa.Dto2 = dt.Rows(0)("Dto2")
                data.DatosTarifa.Dto3 = dt.Rows(0)("Dto3")
            End If
            data.DatosTarifa.UDValoracion = IIf(Articulo.UDValoracion > 0, Articulo.UDValoracion, 1)
            data.SeguimientoDtoTarifa = "RUTA PARA OBTENER LOS DTOS: DTOS POR TIPO DE LAS CONDICIONES DE ALQUILER"
            data.Dtos = True
        End If
    End Sub

    <Task()> Public Shared Sub DtosContratoAlquilerGenerales(ByVal data As DataCalculoTarifaAlquiler, ByVal services As ServiceProvider)
        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        Dim Articulo As ArticuloInfo = Articulos.GetEntity(data.IDArticulo)

        Dim f As New Filter
        f.Add(New NumberFilterItem("IDObra", data.IDObra))
        f.Add(New NumberFilterItem("Tipo", otaTipo.otaDescuentoPorFamilia))
        f.Add(New NumberFilterItem("TipoDato", otaTipoDtos.otaDtoGeneral))
        f.Add(New IsNullFilterItem("IDTipo", True))
        f.Add(New IsNullFilterItem("IDFamilia", True))
        Dim dt As DataTable = New BE.DataEngine().Filter("tbObraTarifaAlquiler", f, "Dto1,Dto2,Dto3")
        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            If Not data.ArticuloSinDto Then
                data.DatosTarifa.Dto1 = dt.Rows(0)("Dto1")
                data.DatosTarifa.Dto2 = dt.Rows(0)("Dto2")
                data.DatosTarifa.Dto3 = dt.Rows(0)("Dto3")
            End If
            data.DatosTarifa.UDValoracion = IIf(Articulo.UDValoracion > 0, Articulo.UDValoracion, 1)
            data.SeguimientoDtoTarifa = "RUTA PARA OBTENER LOS DTOS: DTOS GENERALES DE LAS CONDICIONES DE ALQUILER"
            data.Dtos = True
        End If
    End Sub

    <Task()> Public Shared Sub DtoComericalCliente(ByVal data As DataCalculoTarifaAlquiler, ByVal services As ServiceProvider)
        Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
        Dim Cliente As ClienteInfo = Clientes.GetEntity(data.IDCliente)
        If Cliente.DtoComercial > 0 Then
            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
            Dim Articulo As ArticuloInfo = Articulos.GetEntity(data.IDArticulo)
            If Not data.ArticuloSinDto Then
                data.DatosTarifa.Dto1 = Cliente.DtoComercial
                data.DatosTarifa.Dto2 = 0
                data.DatosTarifa.Dto3 = 0
            End If
            data.DatosTarifa.UDValoracion = IIf(Articulo.UDValoracion > 0, Articulo.UDValoracion, 1)
            data.SeguimientoDtoTarifa = "RUTA PARA OBTENER LOS DTOS: DTO COMERCIAL DEL CLIENTE"
            data.Dtos = True
        End If
    End Sub

#End Region

    <Task()> Public Shared Sub TarifaDelArticulo(ByVal data As DataCalculoTarifaAlquiler, ByVal services As ServiceProvider)
        Dim f As New Filter
        f.Add(New NumberFilterItem("IDObra", data.IDObra))
        f.Add(New NumberFilterItem("Tipo", otaTipo.otaTarifa))
        Dim dtTarifa As DataTable = New BE.DataEngine().Filter("vNegObraTarifaAlquiler", f, "IDTarifa, IDMoneda", "Orden")
        If Not IsNothing(dtTarifa) AndAlso dtTarifa.Rows.Count > 0 Then

            For Each drTarifa As DataRow In dtTarifa.Rows
                Dim dtArtTarifa As DataTable = New TarifaArticulo().SelOnPrimaryKey(drTarifa("IDTarifa"), data.IDArticulo)
                If Not IsNothing(dtArtTarifa) AndAlso dtArtTarifa.Rows.Count > 0 Then
                    data.DatosTarifa.IDMoneda = drTarifa("IDMoneda")

                    data.Final = True
                    data.DatosTarifa.Precio = dtArtTarifa.Rows(0)("Precio")
                    data.DatosTarifa.SeguimientoTarifa = "RUTA PARA OBTENER EL PRECIO: PRECIO DEL ARTICULO DE LA TARIFA ALQUILER- " & drTarifa("IDTarifa") & vbNewLine & data.SeguimientoDtoTarifa
                    If data.Dtos Then
                        data.DatosTarifa.SeguimientoTarifa = data.DatosTarifa.SeguimientoTarifa & vbNewLine & data.SeguimientoDtoTarifa
                    ElseIf dtArtTarifa.Rows(0)("Dto1") > 0 Or dtArtTarifa.Rows(0)("Dto2") > 0 Or dtArtTarifa.Rows(0)("Dto3") > 0 Then
                        data.DatosTarifa.Dto1 = dtArtTarifa.Rows(0)("Dto1")
                        data.DatosTarifa.Dto2 = dtArtTarifa.Rows(0)("Dto2")
                        data.DatosTarifa.Dto3 = dtArtTarifa.Rows(0)("Dto3")
                        data.DatosTarifa.SeguimientoTarifa = data.DatosTarifa.SeguimientoTarifa & vbNewLine & "RUTA PARA OBTENER LOS DTOS: DTOS DEL ARTICULO DE LA TARIFA- " & drTarifa("IDTarifa")
                        'ElseIf Not data.ArticuloSinDto Then
                        '    Dim dataTarifa As New DataCalculoTarifaComercial(data.IDArticulo, data.IDCliente, data.Cantidad, data.Fecha)
                        '    ProcessServer.ExecuteTask(Of DataCalculoTarifaComercial)(AddressOf ProcesoComercial.TarifaComercial, dataTarifa, services)
                        '    If Not dataTarifa.DatosTarifa Is Nothing Then
                        '        data.DatosTarifa.Dto1 = dataTarifa.DatosTarifa.Dto1
                        '        data.DatosTarifa.Dto2 = dataTarifa.DatosTarifa.Dto2
                        '        data.DatosTarifa.Dto3 = dataTarifa.DatosTarifa.Dto3
                        '    End If
                    Else
                        data.DatosTarifa.Dto1 = 0
                        data.DatosTarifa.Dto2 = 0
                        data.DatosTarifa.Dto3 = 0
                    End If
                    data.DatosTarifa.UDValoracion = IIf(dtArtTarifa.Rows(0)("UDValoracion") > 0, dtArtTarifa.Rows(0)("UDValoracion"), 1)

                    'Aplicar tarifa escalar
                    Dim dataTarifa As New DataCalculoTarifaComercial(data.IDArticulo, data.IDCliente, data.Cantidad, data.Fecha)
                    dataTarifa.DatosTarifa.IDTarifa = dtArtTarifa.Rows(0)("IDTarifa")
                    ProcessServer.ExecuteTask(Of DataCalculoTarifaComercial)(AddressOf ProcesoComercial.RecuperarDatosTarifa, dataTarifa, services)
                    data.DatosTarifa.Precio = dataTarifa.DatosTarifa.Precio
                    data.DatosTarifa.Dto1 = dataTarifa.DatosTarifa.Dto1
                    data.DatosTarifa.Dto2 = dataTarifa.DatosTarifa.Dto2
                    data.DatosTarifa.Dto3 = dataTarifa.DatosTarifa.Dto3

                    Exit For
                End If
            Next

        End If
    End Sub

    <Task()> Public Shared Sub TarifaCircuitoComercial(ByVal data As DataCalculoTarifaAlquiler, ByVal services As ServiceProvider)
        Dim dataTarifa As New DataCalculoTarifaComercial(data.IDArticulo, data.IDCliente, data.Cantidad, data.Fecha)
        ProcessServer.ExecuteTask(Of DataCalculoTarifaComercial)(AddressOf ProcesoComercial.TarifaComercial, dataTarifa, services)
        If Not dataTarifa.DatosTarifa Is Nothing Then
            If dataTarifa.DatosTarifa.Precio = 0 Then
                dataTarifa.DatosTarifa.SeguimientoTarifa = String.Empty
            End If
            If data.DatosTarifa.Precio = 0 Then
                data.DatosTarifa.IDMoneda = dataTarifa.DatosTarifa.IDMoneda
                data.DatosTarifa.Precio = dataTarifa.DatosTarifa.Precio
            End If
            If Not data.ArticuloSinDto AndAlso Not data.Dtos Then
                data.DatosTarifa.Dto1 = dataTarifa.DatosTarifa.Dto1
                data.DatosTarifa.Dto2 = dataTarifa.DatosTarifa.Dto2
                data.DatosTarifa.Dto3 = dataTarifa.DatosTarifa.Dto3
            End If
            If Length(dataTarifa.DatosTarifa.SeguimientoTarifa) > 0 Then
                If data.Dtos Then
                    If InStr(1, dataTarifa.DatosTarifa.SeguimientoTarifa, vbNewLine) Then
                        Dim Pos As Integer = InStr(1, dataTarifa.DatosTarifa.SeguimientoTarifa, vbNewLine)
                        Dim SeguimientoTarifa As String = Left(dataTarifa.DatosTarifa.SeguimientoTarifa, Pos - 1)
                        data.DatosTarifa.SeguimientoTarifa = SeguimientoTarifa & vbNewLine & data.SeguimientoDtoTarifa
                    Else
                        data.DatosTarifa.SeguimientoTarifa = data.SeguimientoDtoTarifa
                    End If
                Else
                    data.DatosTarifa.SeguimientoTarifa = dataTarifa.DatosTarifa.SeguimientoTarifa
                End If
            ElseIf Length(data.SeguimientoDtoTarifa) > 0 Then
                data.DatosTarifa.SeguimientoTarifa = data.DatosTarifa.SeguimientoTarifa & vbNewLine & data.SeguimientoDtoTarifa
            End If
        End If
    End Sub

    <Task()> Public Shared Sub PrecioEnMonedaContexto(ByVal data As CalculoTarifaAlquiler.DataCalculoTarifaAlquiler, ByVal services As ServiceProvider)
        If Length(data.DatosTarifa.IDMoneda) > 0 AndAlso Length(data.IDMoneda) > 0 AndAlso data.IDMoneda <> data.DatosTarifa.IDMoneda Then
            If Length(data.Fecha) = 0 Then data.Fecha = Today
            Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
            Dim MonTarifa As MonedaInfo = Monedas.GetMoneda(data.DatosTarifa.IDMoneda, data.Fecha)
            Dim MonContexto As MonedaInfo = Monedas.GetMoneda(data.IDMoneda, data.Fecha)

            If MonContexto.CambioA <> 0 Then
                data.DatosTarifa.Precio = xRound(data.DatosTarifa.Precio * (MonTarifa.CambioA / MonContexto.CambioA), MonContexto.NDecimalesPrecio)
            End If
        End If
    End Sub

End Class
