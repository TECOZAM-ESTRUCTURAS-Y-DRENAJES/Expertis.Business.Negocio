<Serializable()> _
Public Class fImporte
    Public Importe As Double
    Public ImporteA As Double
    Public ImporteB As Double

    Public Sub New()
        Importe = 0
        ImporteA = 0
        ImporteB = 0
    End Sub

End Class

<Transactional()> _
Public Class NegocioGeneral
    Inherits ContextBoundObject

    'Public Shared cnMAX_DATE As Date = New Date(3000, 1, 1)   '//Se pone este valor como máximo, por que el SQL no admite el Date.MaxValue del Visual

    Public Shared cnLENGTH_NIVELES_ANALITICA As Integer = 3

    'Estados de activo predeterminados de la aplicacion (con marca Sistema=1)
    Public Shared ESTADOACTIVO_DISPONIBLE As String = "0"
    Public Shared ESTADOACTIVO_ENMANTENIMIENTO As String = "1"
    Public Shared ESTADOACTIVO_RESERVADA As String = "2"
    Public Shared ESTADOACTIVO_TRABAJANDO As String = "3"
    Public Shared ESTADOACTIVO_VENDIDO As String = "4"
    Public Shared ESTADOACTIVO_BAJA As String = "5"
    Public Shared ESTADOACTIVO_AVERIADO As String = "6"
    Public Shared ESTADOACTIVO_ENTRANSITO As String = "7"
    Public Shared ESTADOACTIVO_OCUPADOENPORTE As String = "8"
    Public Shared ESTADOACTIVO_PENDIENTEDERETORNAR As String = "14"


    Public Function ContadorB(ByVal IdContador As String) As Boolean
        Return ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf NegocioGeneral.ContadorB, IdContador, New ServiceProvider)
    End Function

    Public Shared Function ConvertirDecimales(ByVal dblCantidad As Double, ByVal blnImporte As Boolean, ByVal strMoneda As String) As Double
        'Si se ha intorducido la var. StrIdMoneda, lo que devuelve esta función es la cantidad
        'introducida pero con los decimales correspondientes, según sea un precio o un importe
        If strMoneda <> vbNullString Then
            Dim DtMoneda As DataTable = New Moneda().SelOnPrimaryKey(strMoneda)
            If blnImporte Then
                ConvertirDecimales = System.Math.Round(dblCantidad, DtMoneda.Rows(0)("NDecimalesImp"))
            Else
                ConvertirDecimales = System.Math.Round(dblCantidad, DtMoneda.Rows(0)("NDecimalesPrec"))
            End If
        Else
            ConvertirDecimales = dblCantidad
        End If
    End Function

    Public Function UserName() As String
        Return DAL.AdminData.GetSessionInfo.UserName
    End Function

    Public Shared Function UserID() As Guid
        Return DAL.AdminData.GetSessionInfo.UserID
    End Function

    Public Shared Sub ValidaCIF(ByVal dr As IPropertyAccessor, Optional ByRef blnCancel As Boolean = False)
        Dim strCIFCopia As String
        Dim strCIF As String

        If dr.ContainsKey("CifProveedor") Then
            strCIFCopia = dr("CifProveedor")
            strCIF = dr("CifProveedor")
        ElseIf dr.ContainsKey("CifCliente") Then
            strCIFCopia = dr("CifCliente")
            strCIF = dr("CifCliente")
        ElseIf dr.ContainsKey("DNI") Then
            strCIFCopia = dr("DNI")
            strCIF = dr("DNI")
        End If

        Dim strLetra As String
        If Not IsNumeric(strCIF) Then
            strCIF = UCase(strCIF) 'ponemos la letra en mayúscula
            Dim strDNI As String = Mid(strCIF, 1, Len(strCIF) - 1) 'quitamos la letra del NIF
            If Len(strDNI) >= 7 And IsNumeric(strDNI) Then
                strLetra = ObtenerLetra(strDNI)
                strCIF = strDNI & strLetra
            Else
                ApplicationService.GenerateError("El dato introducido no corresponde a un NIF.") '12359
            End If
            If strCIFCopia <> strCIF Then
                dr("CifCorrecto") = strCIF
                blnCancel = True
            End If
        Else
            strLetra = ObtenerLetra(strCIF)
            strCIF = strCIF & strLetra
        End If

        If dr.ContainsKey("CifProveedor") Then
            dr("CifProveedor") = strCIF
        ElseIf dr.ContainsKey("CifCliente") Then
            dr("CifCliente") = strCIF
        ElseIf dr.ContainsKey("DNI") Then
            dr("DNI") = strCIF
        End If
    End Sub

    <Task()> Public Shared Sub ValidarCIFObligatorio(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.Table.Columns.Contains("CifProveedor") Then
            If Length(data("CifProveedor")) = 0 Then ApplicationService.GenerateError("El Cíf de Proveedor es obligatorio.")
        ElseIf data.Table.Columns.Contains("CifCliente") Then
            If Length(data("CifCliente")) = 0 Then ApplicationService.GenerateError("El Cif de Cliente es obligatorio.")
        End If
    End Sub

    Public Shared Function ObtenerLetra(ByVal strDNI As String) As String
        Dim intMod As Integer
        Dim strLetra As String

        If IsNumeric(strDNI) Then
            Dim intDNI As Integer = CInt(strDNI)
            intMod = (intDNI - (Int(intDNI / 23) * 23)) + 1
            strLetra = Mid$("TRWAGMYFPDXBNJZSQVHLCKE", intMod, 1)
        End If
        Return strLetra
    End Function

    Public Function GetPeriodString(ByVal lngPeriod As enumcpPeriodo) As String
        Select Case lngPeriod
            Case enumcpPeriodo.cpDia
                GetPeriodString = "d"
            Case enumcpPeriodo.cpSemana
                GetPeriodString = "ww"
            Case enumcpPeriodo.cpMes
                GetPeriodString = "m"
            Case enumcpPeriodo.cpAño
                GetPeriodString = "yyyy"
        End Select
    End Function

#Region " Ficheros "

    Public Function AddSpace(ByVal espacios As Integer, ByVal campo As String) As String
        Dim Cadena As String
        Cadena = Space(espacios - Len(campo))
        Return Cadena
    End Function

#End Region

#Region " CalcularPrecioImporte "

    <Task()> Public Shared Sub CalcularPrecioImporte(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider)
        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
        Dim IDMoneda As String
        If data.Contains("IDMoneda") AndAlso Length(data("IDMoneda")) > 0 Then
            IDMoneda = data("IDMoneda")
        Else
            IDMoneda = Monedas.MonedaA.ID
        End If
        Dim MonInfo As MonedaInfo = Monedas.GetMoneda(IDMoneda)

        Dim dblQ As Double
        Dim dblUnidadValoracion As Double
        Dim dblPrecio As Double : Dim dblPVP As Double
        Dim dbldto1 As Double : Dim dbldto2 As Double : Dim dbldto3 As Double
        Dim dblDto As Double : Dim dblDtoProntoPago As Double

        If Nz(data("UdValoracion"), 0) = 0 Then data("UdValoracion") = 1
        If IsNumeric(data("Cantidad")) Then dblQ = CDbl(data("Cantidad"))
        If IsNumeric(data("UdValoracion")) Then dblUnidadValoracion = CDbl(data("UdValoracion"))
        If IsNumeric(data("Precio")) Then dblPrecio = xRound(CDbl(data("Precio")), MonInfo.NDecimalesPrecio)
        If IsNumeric(data("PVP")) Then dblPVP = xRound(CDbl(data("PVP")), MonInfo.NDecimalesImporte)
        If IsNumeric(data("DTO1")) Then dbldto1 = CDbl(data("DTO1"))
        If IsNumeric(data("DTO2")) Then dbldto2 = CDbl(data("DTO2"))
        If IsNumeric(data("DTO3")) Then dbldto3 = CDbl(data("DTO3"))
        '//Si es un artículo especial, no aplicaremos los descuentos Comercial y Pronto Pago
        If data.ContainsKey("Especial") AndAlso Nz(data("Especial"), False) Then
            data("DTO") = 0
            data("DtoProntoPago") = 0
        End If
        If IsNumeric(data("DTO")) Then dblDto = CDbl(data("DTO"))
        If IsNumeric(data("DtoProntoPago")) Then dblDtoProntoPago = CDbl(data("DtoProntoPago"))

        Dim dblIva As Double = 0
        If dblPVP <> 0 Then
            If data.ContainsKey("IDTipoIVA") AndAlso Length(data("IDTipoIVA")) > 0 Then
                Dim TiposIVA As EntityInfoCache(Of TipoIvaInfo) = services.GetService(Of EntityInfoCache(Of TipoIvaInfo))()
                ' HistoricoTipoIVA
                Dim vFecha As Date
                If data.ContainsKey("Fecha") AndAlso Length(data("Fecha")) > 0 Then
                    vFecha = data("Fecha")
                Else
                    vFecha = Date.Today
                End If
                Dim IVAInfo As TipoIvaInfo = TiposIVA.GetEntity(data("IDTipoIVA"), vFecha)

                If Not IsNothing(IVAInfo) Then
                    dblIva = IVAInfo.Factor
                    If IVAInfo.SinRepercutir Then
                        dblIva = IVAInfo.IVASinRepercutir
                    End If
                End If
            End If
            data("Precio") = xRound((dblPVP * 100) / (dblIva + 100), MonInfo.NDecimalesPrecio)
        Else
            data("Precio") = xRound(dblPrecio, MonInfo.NDecimalesPrecio)
        End If

        Dim dblTotal As Double
        If dblUnidadValoracion <> 0 Then
            dblTotal = ((dblQ / dblUnidadValoracion) * data("Precio") * Nz(data("QTiempo"), 1))
            Dim dblImporte As Double = (dblTotal * dbldto1) / 100
            dblTotal = dblTotal - dblImporte
            dblImporte = (dblTotal * dbldto2) / 100
            dblTotal = dblTotal - dblImporte
            dblImporte = (dblTotal * dbldto3) / 100
            dblTotal = dblTotal - dblImporte
            dblImporte = (dblTotal * dblDto) / 100
            dblTotal = dblTotal - dblImporte
            dblImporte = (dblTotal * dblDtoProntoPago) / 100
            dblTotal = dblTotal - dblImporte
        End If

        Dim dblTotalPVP As Double
        If dblUnidadValoracion <> 0 Then
            dblTotalPVP = ((dblQ / dblUnidadValoracion) * dblPVP * Nz(data("QTiempo"), 1))
            Dim dblPVPDto As Double = (dblTotalPVP * dbldto1) / 100
            dblTotalPVP = dblTotalPVP - dblPVPDto
            dblPVPDto = (dblTotalPVP * dbldto2) / 100
            dblTotalPVP = dblTotalPVP - dblPVPDto
            dblPVPDto = (dblTotalPVP * dbldto3) / 100
            dblTotalPVP = dblTotalPVP - dblPVPDto
            dblPVPDto = (dblTotalPVP * dblDto) / 100
            dblTotalPVP = dblTotalPVP - dblPVPDto
            dblPVPDto = (dblTotalPVP * dblDtoProntoPago) / 100
            dblTotalPVP = dblTotalPVP - dblPVPDto
        End If

        data("Cantidad") = dblQ
        data("UdValoracion") = dblUnidadValoracion
        data("Dto1") = dbldto1
        data("Dto2") = dbldto2
        data("Dto3") = dbldto3
        data("Dto") = dblDto
        data("DtoProntoPago") = dblDtoProntoPago

        data("PVP") = xRound(dblPrecio * (1 - dbldto1 / 100) * (1 - dbldto2 / 100) * (1 - dbldto3 / 100) * (1 - dblDto / 100) * (1 - dblDtoProntoPago / 100) * (1 + dblIva / 100), MonInfo.NDecimalesImporte)

        If Nz(data("Regalo"), False) Then
            data("Importe") = 0
        Else
            'data("Importe") = dblTotal
            'DAVID VELASCO 26/05
            'Si la columna esta marcada, me pone el precio que está en vez de la cuenta que realiza
            Try
                If Nz(data("Modificado"), 0) = 0 Then
                    data("Importe") = dblTotal
                End If
            Catch ex As Exception
                MsgBox(ex.ToString())
            End Try
            
        End If

        If dblPVP <> 0 AndAlso Not Nz(data("Regalo"), False) Then
            data("PVP") = xRound(dblPVP, MonInfo.NDecimalesImporte)
            data("ImportePVP") = xRound(dblTotalPVP, MonInfo.NDecimalesImporte)
            data("Importe") = xRound((data("ImportePVP") * 100) / (dblIva + 100), MonInfo.NDecimalesImporte)
        Else
            data("PVP") = 0
            data("ImportePVP") = 0
        End If

    End Sub

#End Region

#Region " CalcularImportes "

    <Task()> Public Shared Sub CalcularImportes(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Not data.Context.ContainsKey("CambioA") OrElse Length(data.Context("CambioA")) = 0 Then data.Context("CambioA") = 0
        If Not data.Context.ContainsKey("CambioB") OrElse Length(data.Context("CambioB")) = 0 Then data.Context("CambioB") = 0
        If data.Context.ContainsKey("IDMoneda") AndAlso Length(data.Context("IDMoneda")) > 0 Then data.Current("IDMoneda") = data.Context("IDMoneda")
        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf NegocioGeneral.CalcularPrecioImporte, data.Current, services)
        If Length(data.Context("IDMoneda")) > 0 AndAlso Length(data.Context("CambioA")) > 0 AndAlso Length(data.Context("CambioB")) > 0 Then
            Dim ValAyB As New ValoresAyB(data.Current, data.Context("IDMoneda"), data.Context("CambioA"), data.Context("CambioB"))
            ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf MantenimientoValoresAyB, ValAyB, services)
        End If
    End Sub

#End Region

#Region " CALCULOS DE FECHAS "

    <Serializable()> _
    Public Class CrearFechas
        Public FechaDesde As Date
        Public FechaHasta As Date
        Public FechaAlternativa As Date
        Public FechaDivision As Date
        Public Periodo As Boolean

        Public Sub New(ByVal FechaDesde As Date, ByVal FechaHasta As Date, ByVal FechaAlternativa As Date, ByVal FechaDivision As Date)
            Me.FechaDesde = FechaDesde
            Me.FechaHasta = FechaHasta
            Me.FechaAlternativa = FechaAlternativa
            Me.FechaDivision = FechaDivision
        End Sub
        Public Sub New()

        End Sub
    End Class

    'Función que pasada una fecha y un número de días a sumar a la fecha devuelve el día de pago correspondiente.
    <Serializable()> _
    Public Class dataCalculoFechaVencimiento
        Public FechaFactura As Date
        Public Periodo As Integer
        Public Intervalo As String
        Public IDDiaPago As String
        Public Cliente As Boolean
        Public IDCliente As String
        Public IDProveedor As String
        Public FechaVencimiento As Date

        Public Sub New(ByVal FechaFactura As Date, ByVal Periodo As Integer, ByVal Intervalo As String, ByVal IDDiaPago As String, ByVal Cliente As Boolean, ByVal IDCliProv As String)
            Me.FechaFactura = FechaFactura
            Me.Periodo = Periodo
            Me.Intervalo = Intervalo
            Me.IDDiaPago = IDDiaPago
            Me.Cliente = Cliente
            If Cliente Then
                Me.IDCliente = IDCliProv
            Else
                Me.IDProveedor = IDCliProv
            End If
        End Sub
    End Class
    <Task()> Public Shared Sub CalcularFechaVencimiento(ByVal data As dataCalculoFechaVencimiento, ByVal services As ServiceProvider)
        Dim Fecha As Date = DateAdd(data.Intervalo, data.Periodo, data.FechaFactura)
        Dim Dia As Integer = Fecha.Day
        Dim Mes As Integer = Fecha.Month
        Dim Año As Integer = Fecha.Year

        Dim Dia1 As Integer = 0
        Dim Dia2 As Integer = 0
        Dim Dia3 As Integer = 0

        If Length(data.IDDiaPago) > 0 Then
            Dim drDiaPago As DataRow = New DiaPago().GetItemRow(data.IDDiaPago)
            Dia1 = Nz(drDiaPago("Dia1"), 0)
            Dia2 = Nz(drDiaPago("Dia2"), 0)
            Dia3 = Nz(drDiaPago("Dia3"), 0)
        End If

        If Dia1 <> 0 Then
            If (Dia <> Dia1) And (Dia <> Dia2) And (Dia <> Dia3) Then
                If (Dia > Dia3 And Dia3 <> 0) Then
                    Dia = Dia1
                    Mes = IIf(Mes = 12, 1, Mes + 1)
                    Año = IIf(Mes = 1, Año + 1, Año)
                Else
                    If (Dia < Dia3 Or Dia3 = 0) And Dia > Dia2 And Dia2 <> 0 Then
                        If Dia3 = 0 Then
                            Dia = Dia1
                            Mes = IIf(Mes = 12, 1, Mes + 1)
                            Año = IIf(Mes = 1, Año + 1, Año)
                        Else
                            Dia = Dia3
                        End If
                    Else
                        If (Dia < Dia3 Or Dia3 = 0) And (Dia < Dia2 Or Dia2 = 0) And Dia > Dia1 And Dia1 <> 0 Then
                            If Dia2 = 0 Then
                                Dia = Dia1
                                Mes = IIf(Mes = 12, 1, Mes + 1)
                                Año = IIf(Mes = 1, Año + 1, Año)
                            Else
                                Dia = Dia2
                            End If
                        Else
                            If (Dia < Dia3 Or Dia3 = 0) And (Dia < Dia2 Or Dia2 = 0) And Dia < Dia1 Then
                                Dia = Dia1
                            End If
                        End If
                    End If
                End If
            End If
        End If

        ' para controlar el 29,30,31 de febrero
        Do Until IsDate(Año & "/" & Mes & "/" & Dia)
            Dia = Dia - 1
        Loop

        Dim FechaFinal As Date = New Date(Año, Mes, Dia)

        'Miramos si la fecha obtenida está incluida en el periodo vacacional
        Dim dtVacaciones As DataTable = Nothing
        If data.Cliente Then
            dtVacaciones = New ClienteVacaciones().Filter(New StringFilterItem("IDCliente", data.IDCliente), "FechaDesde")
        Else
            dtVacaciones = New ProveedorVacaciones().Filter(New StringFilterItem("IDProveedor", data.IDProveedor), "FechaDesde")
        End If

        For Each drVacacion As DataRow In dtVacaciones.Rows
            Dim dataFecOrigen As New CrearFechas(Nz(drVacacion("FechaDesde")), Nz(drVacacion("FechaHasta")), Nz(drVacacion("FechaAlternativa")), Nz(drVacacion("FechaDivision")))

            Dim dataFechas As New dataPrepararFechas(Fecha, dataFecOrigen)
            Dim dataFecFinal As CrearFechas = ProcessServer.ExecuteTask(Of dataPrepararFechas, CrearFechas)(AddressOf PrepararFechas, dataFechas, services)

            Dim dataFechaAlt As New dataFechaAlternativa(Fecha, FechaFinal, dataFecOrigen, dataFecFinal, data.FechaFactura)
            data.FechaVencimiento = ProcessServer.ExecuteTask(Of dataFechaAlternativa, Date)(AddressOf CalculoFechaAlternativa, dataFechaAlt, services)
            If data.FechaVencimiento = System.DateTime.MinValue And dataFecFinal.Periodo = True Then
                If dataFecFinal.FechaDivision <> System.DateTime.MinValue Then
                    If Fecha < dataFecFinal.FechaDivision Then
                        Dim dataDivMenor As New dataFechaFechaDivisionMenor(Dia1, Dia2, Dia3, dataFecFinal.FechaDesde, Fecha, data.FechaFactura, dataFecFinal)
                        data.FechaVencimiento = ProcessServer.ExecuteTask(Of dataFechaFechaDivisionMenor, Date)(AddressOf CalculoFechaDivisionMenor, dataDivMenor, services)
                    End If
                End If
                If data.FechaVencimiento = System.DateTime.MinValue Then
                    Dim dataDivMayor As New dataFechaFechaDivisionMayor(Dia1, Dia2, Dia3, dataFecFinal.FechaDesde, Fecha)
                    data.FechaVencimiento = ProcessServer.ExecuteTask(Of dataFechaFechaDivisionMayor, Date)(AddressOf CalculoFechaDivisionMayor, dataDivMayor, services)
                End If
            End If
            If data.FechaVencimiento <> System.DateTime.MinValue Then
                Exit For
            End If
        Next

        If data.FechaVencimiento = System.DateTime.MinValue Then
            data.FechaVencimiento = FechaFinal
        End If
    End Sub

    <Serializable()> _
    Public Class dataFechaFechaDivisionMenor
        Public Dia1 As Integer
        Public Dia2 As Integer
        Public Dia3 As Integer
        Public FechaDesde As Date
        Public Fecha As Date
        Public FechaFactura As Date
        Public FechaFinal As CrearFechas

        Public Sub New(ByVal Dia1 As Integer, ByVal Dia2 As Integer, ByVal Dia3 As Integer, ByVal FechaDesde As Date, ByVal Fecha As Date, ByVal FechaFactura As Date, ByVal FechaFinal As CrearFechas)
            Me.Dia1 = Dia1
            Me.Dia2 = Dia2
            Me.Dia3 = Dia3
            Me.FechaDesde = FechaDesde
            Me.Fecha = Fecha
            Me.FechaFactura = FechaFactura
            Me.FechaFinal = FechaFinal
        End Sub
    End Class
    <Task()> Public Shared Function CalculoFechaDivisionMenor(ByVal data As dataFechaFechaDivisionMenor, ByVal services As ServiceProvider) As Date
        Dim Fecha As Date
        Dim FechaAux As Date

        'El primer día de pago del mes anterior empezando por el final.
        If data.Dia3 <> 0 Then
            FechaAux = New Date(data.FechaDesde.Year, data.FechaDesde.Month, data.Dia3)
            If data.FechaDesde > FechaAux Then
                Fecha = New Date(data.FechaDesde.Year, data.FechaDesde.Month, data.Dia3)
                If Fecha <= data.FechaFactura Or (Fecha >= data.FechaFinal.FechaDesde And Fecha <= data.FechaFinal.FechaHasta) Then
                    Fecha = System.DateTime.MinValue
                End If
            End If
        End If
        If Fecha = System.DateTime.MinValue AndAlso data.Dia2 <> 0 Then
            FechaAux = New Date(data.FechaDesde.Year, data.FechaDesde.Month, data.Dia2)
            If data.FechaDesde > FechaAux Then
                Fecha = New Date(data.FechaDesde.Year, data.FechaDesde.Month, data.Dia2)
                If Fecha <= data.FechaFactura Or (Fecha >= data.FechaFinal.FechaDesde And Fecha <= data.FechaFinal.FechaHasta) Then
                    Fecha = System.DateTime.MinValue
                End If
            End If
        End If
        If Fecha = System.DateTime.MinValue AndAlso data.Dia1 <> 0 Then
            FechaAux = New Date(data.FechaDesde.Year, data.FechaDesde.Month, data.Dia1)
            If data.FechaDesde > FechaAux Then
                Fecha = New Date(data.FechaDesde.Year, data.FechaDesde.Month, data.Dia1)
                If Fecha <= data.FechaFactura Or (Fecha >= data.FechaFinal.FechaDesde And Fecha <= data.FechaFinal.FechaHasta) Then
                    Fecha = System.DateTime.MinValue
                End If
            End If
        End If
        If Fecha = System.DateTime.MinValue AndAlso data.Dia3 <> 0 Then
            Fecha = New Date(data.FechaDesde.Year, data.FechaDesde.Month, data.Dia3)
            Fecha = DateAdd(DateInterval.Month, -1, Fecha)
            Fecha = New Date(Fecha.Year, Fecha.Month, data.Dia3)
            If Fecha <= data.FechaFactura Or (Fecha >= data.FechaFinal.FechaDesde And Fecha <= data.FechaFinal.FechaHasta) Then
                Fecha = System.DateTime.MinValue
            End If
        End If
        If Fecha = System.DateTime.MinValue AndAlso data.Dia2 <> 0 Then
            Fecha = New Date(data.FechaDesde.Year, data.FechaDesde.Month, data.Dia2)
            Fecha = DateAdd(DateInterval.Month, -1, Fecha)
            Fecha = New Date(Fecha.Year, Fecha.Month, data.Dia2)
            If Fecha <= data.FechaFactura Or (Fecha >= data.FechaFinal.FechaDesde And Fecha <= data.FechaFinal.FechaHasta) Then
                Fecha = System.DateTime.MinValue
            End If
        End If
        If Fecha = System.DateTime.MinValue AndAlso data.Dia1 <> 0 Then
            Fecha = New Date(data.FechaDesde.Year, data.FechaDesde.Month, data.Dia1)
            Fecha = DateAdd(DateInterval.Month, -1, Fecha)
            Fecha = New Date(Fecha.Year, Fecha.Month, data.Dia1)
            If Fecha <= data.FechaFactura Or (Fecha >= data.FechaFinal.FechaDesde And Fecha <= data.FechaFinal.FechaHasta) Then
                Fecha = System.DateTime.MinValue
            End If
        End If
        If Fecha = System.DateTime.MinValue Then
            Fecha = New Date(data.FechaDesde.Year, data.FechaDesde.Month, data.FechaDesde.Day)
            Fecha = DateAdd(DateInterval.Day, -1, Fecha)
            If Fecha <= data.FechaFactura Or (Fecha >= data.FechaFinal.FechaDesde And Fecha <= data.FechaFinal.FechaHasta) Then
                Fecha = System.DateTime.MinValue
            End If
        End If
        ''Para controlar el 29 , 30 y 31 de febrero

        Do Until IsDate(Fecha)
            Fecha = DateAdd(DateInterval.Day, -1, Fecha)
        Loop
        If Fecha <= data.FechaFactura Or (Fecha >= data.FechaFinal.FechaDesde And Fecha <= data.FechaFinal.FechaHasta) Then
            Fecha = System.DateTime.MinValue
        End If

        Return Fecha
    End Function

    <Serializable()> _
    Public Class dataFechaFechaDivisionMayor
        Public Dia1 As Integer
        Public Dia2 As Integer
        Public Dia3 As Integer
        Public FechaHasta As Date
        Public Fecha As Date

        Public Sub New(ByVal Dia1 As Integer, ByVal Dia2 As Integer, ByVal Dia3 As Integer, ByVal FechaHasta As Date, ByVal Fecha As Date)
            Me.Dia1 = Dia1
            Me.Dia2 = Dia2
            Me.Dia3 = Dia3
            Me.FechaHasta = FechaHasta
            Me.Fecha = Fecha
        End Sub
    End Class
    <Task()> Public Shared Function CalculoFechaDivisionMayor(ByVal data As dataFechaFechaDivisionMayor, ByVal services As ServiceProvider) As Date
        Dim Fecha As Date
        Dim FechaAux As Date

        'El primer día de pago del mes siguiente empezando por el principio.
        If data.Dia1 <> 0 Then
            FechaAux = New Date(data.FechaHasta.Year, data.FechaHasta.Month, data.Dia1)
            If data.FechaHasta < FechaAux Then
                Fecha = New Date(data.FechaHasta.Year, data.FechaHasta.Month, data.Dia1)
            End If
        End If

        If Fecha = System.DateTime.MinValue AndAlso data.Dia2 <> 0 Then
            FechaAux = New Date(data.FechaHasta.Year, data.FechaHasta.Month, data.Dia2)
            If data.FechaHasta < FechaAux Then
                Fecha = New Date(data.FechaHasta.Year, data.FechaHasta.Month, data.Dia2)
            End If
        End If

        If Fecha = System.DateTime.MinValue AndAlso data.Dia3 <> 0 Then
            FechaAux = New Date(data.FechaHasta.Year, data.FechaHasta.Month, data.Dia3)
            If data.FechaHasta < FechaAux Then
                Fecha = New Date(data.FechaHasta.Year, data.FechaHasta.Month, data.Dia3)
            End If
        End If

        If Fecha = System.DateTime.MinValue AndAlso data.Dia1 <> 0 Then
            Fecha = New Date(data.FechaHasta.Year, data.FechaHasta.Month, data.Dia1)
            Fecha = DateAdd(DateInterval.Month, 1, Fecha)
            Fecha = New Date(Fecha.Year, Fecha.Month, data.Dia1)
        End If

        If Fecha = System.DateTime.MinValue AndAlso data.Dia2 <> 0 Then
            Fecha = New Date(data.FechaHasta.Year, data.FechaHasta.Month, data.Dia2)
            Fecha = DateAdd(DateInterval.Month, 1, Fecha)
            Fecha = New Date(Fecha.Year, Fecha.Month, data.Dia2)
        End If

        If Fecha = System.DateTime.MinValue AndAlso data.Dia3 <> 0 Then
            Fecha = New Date(data.FechaHasta.Year, data.FechaHasta.Month, data.Dia3)
            Fecha = DateAdd(DateInterval.Month, 1, Fecha)
            Fecha = New Date(Fecha.Year, Fecha.Month, data.Dia3)
        End If

        If Fecha = System.DateTime.MinValue Then
            Fecha = New Date(data.Fecha.Year, data.FechaHasta.Month, data.FechaHasta.Day)
            Fecha = DateAdd(DateInterval.Day, 1, Fecha)
        End If
        ''Para controlar el 29 , 30 y 31 de febrero
        If Fecha <> System.DateTime.MinValue Then
            Do Until IsDate(Fecha)
                Fecha = DateAdd(DateInterval.Day, 1, Fecha)
            Loop
        End If

        Return Fecha
    End Function

    <Serializable()> _
    Public Class dataPrepararFechas
        Public FechaPropuesta As Date
        Public FechaOrigen As CrearFechas

        Public Sub New(ByVal FechaPropuesta As Date, ByVal FechaOrigen As CrearFechas)
            Me.FechaPropuesta = FechaPropuesta
            Me.FechaOrigen = FechaOrigen
        End Sub
    End Class
    <Task()> Public Shared Function PrepararFechas(ByVal data As dataPrepararFechas, ByVal services As ServiceProvider) As CrearFechas
        Dim dataFechas As New CrearFechas
        If data.FechaOrigen.FechaDesde.Year = data.FechaOrigen.FechaHasta.Year Then
            dataFechas.FechaDesde = New Date(data.FechaPropuesta.Year, data.FechaOrigen.FechaDesde.Month, data.FechaOrigen.FechaDesde.Day)
            dataFechas.FechaHasta = New Date(data.FechaPropuesta.Year, data.FechaOrigen.FechaHasta.Month, data.FechaOrigen.FechaHasta.Day)
        End If

        Dim y As Integer
        If data.FechaOrigen.FechaDesde.Year <> data.FechaOrigen.FechaHasta.Year Then
            Dim DiaDesde As Integer = data.FechaOrigen.FechaDesde.DayOfYear
            Dim DiaHasta As Integer = data.FechaOrigen.FechaHasta.DayOfYear
            Dim DiaPropuesta As Integer = data.FechaPropuesta.DayOfYear
            If DiaPropuesta < DiaDesde Then
                y = DateDiff(DateInterval.Year, data.FechaOrigen.FechaDesde, data.FechaOrigen.FechaHasta)
                dataFechas.FechaDesde = New Date(data.FechaPropuesta.Year - y, data.FechaOrigen.FechaDesde.Month, data.FechaOrigen.FechaDesde.Day)
                dataFechas.FechaHasta = New Date(data.FechaPropuesta.Year, data.FechaOrigen.FechaHasta.Month, data.FechaOrigen.FechaHasta.Day)
            Else
                y = DateDiff(DateInterval.Year, data.FechaOrigen.FechaDesde, data.FechaOrigen.FechaHasta)
                dataFechas.FechaDesde = New Date(data.FechaPropuesta.Year, data.FechaOrigen.FechaDesde.Month, data.FechaOrigen.FechaDesde.Day)
                dataFechas.FechaHasta = New Date(data.FechaPropuesta.Year + y, data.FechaOrigen.FechaHasta.Month, data.FechaOrigen.FechaHasta.Day)
            End If
        End If

        If data.FechaOrigen.FechaAlternativa <> System.DateTime.MinValue Then
            y = DateDiff(DateInterval.Year, data.FechaOrigen.FechaDesde, data.FechaOrigen.FechaAlternativa)

            dataFechas.FechaAlternativa = New Date(dataFechas.FechaDesde.Year + y, data.FechaOrigen.FechaAlternativa.Month, data.FechaOrigen.FechaAlternativa.Day)
        End If
        If data.FechaOrigen.FechaDivision <> System.DateTime.MinValue Then
            y = DateDiff(DateInterval.Year, data.FechaOrigen.FechaDesde, data.FechaOrigen.FechaDivision)

            dataFechas.FechaDivision = New Date(dataFechas.FechaDesde.Year + y, data.FechaOrigen.FechaDivision.Month, data.FechaOrigen.FechaDivision.Day)
        End If

        Return dataFechas
    End Function

    <Serializable()> _
    Public Class dataFechaAlternativa
        Public FechaValidar As Date
        Public FechaCalculada As Date
        Public FechaOrigen As CrearFechas
        Public FechaFinal As CrearFechas
        Public FechaFactura As Date

        Public Sub New(ByVal FechaValidar As Date, ByVal FechaCalculada As Date, ByVal FechaOrigen As CrearFechas, ByVal FechaFinal As CrearFechas, ByVal FechaFactura As Date)
            Me.FechaValidar = FechaValidar
            Me.FechaCalculada = FechaCalculada
            Me.FechaOrigen = FechaOrigen
            Me.FechaFinal = FechaFinal
            Me.FechaFactura = FechaFactura
        End Sub
    End Class
    <Task()> Public Shared Function CalculoFechaAlternativa(ByVal data As dataFechaAlternativa, ByVal services As ServiceProvider) As Date
        'Función que nos devuelve la fecha alternativa o la calculada

        '1º miramos si la fecha calculada con día es correcta
        data.FechaFinal.Periodo = False
        If data.FechaValidar >= data.FechaFinal.FechaDesde And data.FechaValidar <= data.FechaFinal.FechaHasta Then
            '2º miramos si tiene fecha alternativa
            If data.FechaFinal.FechaAlternativa <> System.DateTime.MinValue Then
                If data.FechaValidar >= data.FechaFinal.FechaDesde And data.FechaValidar <= data.FechaFinal.FechaHasta Then
                    data.FechaFinal.Periodo = True
                    'Fecha incorrecta, miramos si la alternativa es o no correcta
                    If data.FechaFinal.FechaAlternativa > data.FechaFactura Then
                        CalculoFechaAlternativa = data.FechaFinal.FechaAlternativa
                        Exit Function
                    End If
                End If
            End If
            If data.FechaFinal.FechaDivision <> System.DateTime.MinValue Then
                data.FechaFinal.Periodo = True
                Exit Function
            End If
        End If

        Dim dataFechas As New dataPrepararFechas(data.FechaCalculada, data.FechaOrigen)
        data.FechaFinal = ProcessServer.ExecuteTask(Of dataPrepararFechas, CrearFechas)(AddressOf PrepararFechas, dataFechas, services)

        If data.FechaCalculada >= data.FechaFinal.FechaDesde And data.FechaCalculada <= data.FechaFinal.FechaHasta Then
            data.FechaFinal.Periodo = True
            If data.FechaFinal.FechaAlternativa <> System.DateTime.MinValue Then
                If data.FechaCalculada >= data.FechaFinal.FechaDesde And data.FechaCalculada <= data.FechaFinal.FechaHasta Then
                    'Fecha incorrecta, miramos si la alternativa es o no correcta
                    If data.FechaFinal.FechaAlternativa > data.FechaFactura Then
                        data.FechaFinal.Periodo = True
                        CalculoFechaAlternativa = data.FechaFinal.FechaAlternativa
                        Exit Function
                    End If
                End If
            End If
        End If

    End Function

#End Region

    <Serializable()> _
    Public Class dataCalculoDigitosControl
        Public Entidad As String
        Public Oficina As String
        Public NCuenta As String

        Public Sub New(ByVal Entidad As String, ByVal Oficina As String, ByVal NCuenta As String)
            Me.Entidad = Entidad
            Me.Oficina = Oficina
            Me.NCuenta = NCuenta
        End Sub
    End Class
    <Task()> Public Shared Function CalculoDigitosControl(ByVal data As dataCalculoDigitosControl, ByVal services As ServiceProvider) As String
        Dim PrimerDigito As String
        Dim SegundoDigito As String

        Dim ii As Integer = CInt(Nz(Mid(data.Entidad, 1, 1), 0)) * 4 + CInt(Nz(Mid(data.Entidad, 2, 1), 0)) * 8 + CInt(Nz(Mid(data.Entidad, 3, 1), 0)) * 5 + CInt(Nz(Mid(data.Entidad, 4, 1), 0)) * 10 + CInt(Nz(Mid(data.Oficina, 1, 1), 0)) * 9 + CInt(Nz(Mid(data.Oficina, 2, 1), 0)) * 7 + CInt(Nz(Mid(data.Oficina, 3, 1), 0)) * 3 + CInt(Nz(Mid(data.Oficina, 4, 1), 0)) * 6
        Dim jj As Integer = CInt(Nz(Mid(data.NCuenta, 1, 1), 0)) * 1 + CInt(Nz(Mid(data.NCuenta, 2, 1), 0)) * 2 + CInt(Nz(Mid(data.NCuenta, 3, 1), 0)) * 4 + CInt(Nz(Mid(data.NCuenta, 4, 1), 0)) * 8 + CInt(Nz(Mid(data.NCuenta, 5, 1), 0)) * 5 + CInt(Nz(Mid(data.NCuenta, 6, 1), 0)) * 10 + CInt(Nz(Mid(data.NCuenta, 7, 1), 0)) * 9 + CInt(Nz(Mid(data.NCuenta, 8, 1), 0)) * 7 + CInt(Nz(Mid(data.NCuenta, 9, 1), 0)) * 3 + CInt(Nz(Mid(data.NCuenta, 10, 1), 0)) * 6
        If (11 - ii Mod 11) = 11 Then
            PrimerDigito = 0
        ElseIf (11 - ii Mod 11) = 10 Then
            PrimerDigito = 1
        Else
            PrimerDigito = CStr(11 - ii Mod 11)
        End If

        If (11 - jj Mod 11) = 11 Then
            SegundoDigito = 0
        ElseIf (11 - jj Mod 11) = 10 Then
            SegundoDigito = 1
        Else
            SegundoDigito = CStr(11 - jj Mod 11)
        End If

        Return PrimerDigito & SegundoDigito
    End Function

#Region " CAMBIO DE MONEDA "

    <Task()> Public Shared Sub CambioMoneda(ByVal data As DataCambioMoneda, ByVal services As ServiceProvider)
        '//Cambia los valores de la moneda old a la moneda new
        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
        Dim MonInfoOld As MonedaInfo = Monedas.GetMoneda(data.IDMonedaOld, data.Fecha)
        Dim MonInfoNew As MonedaInfo = Monedas.GetMoneda(data.IDMonedaNew, data.Fecha)

        If MonInfoNew.CambioA <> 0 Then
            If data.Row.ContainsKey("CambioA") Then
                data.Row("CambioA") = MonInfoNew.CambioA
            End If
            If data.Row.ContainsKey("CambioB") Then
                data.Row("CambioB") = MonInfoNew.CambioB
            End If
            If data.Row.ContainsKey("Precio") Then
                data.Row("Precio") = Nz(data.Row("Precio"), 0) * (MonInfoOld.CambioA / MonInfoNew.CambioA)
            End If
            If data.Row.ContainsKey("Importe") Then
                data.Row("Importe") = Nz(data.Row("Importe"), 0) * (MonInfoOld.CambioA / MonInfoNew.CambioA)
            End If
            If data.Row.ContainsKey("ImpLineas") Then
                data.Row("ImpLineas") = Nz(data.Row("ImpLineas"), 0) * (MonInfoOld.CambioA / MonInfoNew.CambioA)
            End If
            If data.Row.ContainsKey("ImpIVA") Then
                data.Row("ImpIVA") = Nz(data.Row("ImpIVA"), 0) * (MonInfoOld.CambioA / MonInfoNew.CambioA)
            End If
            If data.Row.ContainsKey("ImpSinRepercutir") Then
                data.Row("ImpSinRepercutir") = Nz(data.Row("ImpSinRepercutir"), 0) * (MonInfoOld.CambioA / MonInfoNew.CambioA)
            End If
            If data.Row.ContainsKey("ImpRE") Then
                data.Row("ImpRE") = Nz(data.Row("ImpRE"), 0) * (MonInfoOld.CambioA / MonInfoNew.CambioA)
            End If
            If data.Row.ContainsKey("ImpDto") Then
                data.Row("ImpDto") = Nz(data.Row("ImpDto"), 0) * (MonInfoOld.CambioA / MonInfoNew.CambioA)
            End If
            If data.Row.ContainsKey("ImpTotal") Then
                data.Row("ImpTotal") = Nz(data.Row("ImpTotal"), 0) * (MonInfoOld.CambioA / MonInfoNew.CambioA)
            End If
            If data.Row.ContainsKey("ImportePVP") Then
                data.Row("ImportePVP") = Nz(data.Row("ImportePVP"), 0) * (MonInfoOld.CambioA / MonInfoNew.CambioA)
            End If
            If data.Row.ContainsKey("PVP") Then
                data.Row("PVP") = Nz(data.Row("PVP"), 0) * (MonInfoOld.CambioA / MonInfoNew.CambioA)
            End If
            If data.Row.ContainsKey("BaseImponible") Then
                data.Row("BaseImponible") = Nz(data.Row("BaseImponible"), 0) * (MonInfoOld.CambioA / MonInfoNew.CambioA)
            End If
            If data.Row.ContainsKey("BaseImponibleEspecial") Then
                data.Row("BaseImponibleEspecial") = Nz(data.Row("BaseImponibleEspecial"), 0) * (MonInfoOld.CambioA / MonInfoNew.CambioA)
            End If
            If data.Row.ContainsKey("BaseImponibleNormal") Then
                data.Row("BaseImponibleNormal") = Nz(data.Row("BaseImponibleNormal"), 0) * (MonInfoOld.CambioA / MonInfoNew.CambioA)
            End If
            If data.Row.ContainsKey("ImpIntrastat") Then
                data.Row("ImpIntrastat") = Nz(data.Row("ImpIntrastat"), 0) * (MonInfoOld.CambioA / MonInfoNew.CambioA)
            End If
            If data.Row.ContainsKey("ImpDtoFactura") Then
                data.Row("ImpDtoFactura") = Nz(data.Row("ImpDtoFactura"), 0) * (MonInfoOld.CambioA / MonInfoNew.CambioA)
            End If
            If data.Row.ContainsKey("ImpDpp") Then
                data.Row("ImpDpp") = Nz(data.Row("ImpDpp"), 0) * (MonInfoOld.CambioA / MonInfoNew.CambioA)
            End If
            If data.Row.ContainsKey("ImpRecFinan") Then
                data.Row("ImpRecFinan") = Nz(data.Row("ImpRecFinan"), 0) * (MonInfoOld.CambioA / MonInfoNew.CambioA)
            End If
            If data.Row.ContainsKey("ImpRetencion") Then
                data.Row("ImpRetencion") = Nz(data.Row("ImpRetencion"), 0) * (MonInfoOld.CambioA / MonInfoNew.CambioA)
            End If
            If data.Row.ContainsKey("ImpVencimiento") Then
                data.Row("ImpVencimiento") = Nz(data.Row("ImpVencimiento"), 0) * (MonInfoOld.CambioA / MonInfoNew.CambioA)
            End If
            If data.Row.ContainsKey("RecargoFinanciero") Then
                data.Row("RecargoFinanciero") = Nz(data.Row("RecargoFinanciero"), 0) * (MonInfoOld.CambioA / MonInfoNew.CambioA)
            End If
            If data.Row.ContainsKey("ImpPedido") Then
                data.Row("ImpPedido") = Nz(data.Row("ImpPedido"), 0) * (MonInfoOld.CambioA / MonInfoNew.CambioA)
            End If
            If data.Row.ContainsKey("ImpPuntoVerde") Then
                data.Row("ImpPuntoVerde") = Nz(data.Row("ImpPuntoVerde"), 0) * (MonInfoOld.CambioA / MonInfoNew.CambioA)
            End If
            If data.Row.ContainsKey("ARepercutir") Then
                data.Row("ARepercutir") = Nz(data.Row("ARepercutir"), 0) * (MonInfoOld.CambioA / MonInfoNew.CambioA)
            End If
            If data.Row.ContainsKey("CosteUnitario") Then
                data.Row("CosteUnitario") = Nz(data.Row("CosteUnitario"), 0) * (MonInfoOld.CambioA / MonInfoNew.CambioA)
            End If
            If data.Row.ContainsKey("PrecioVenta") Then
                data.Row("PrecioVenta") = Nz(data.Row("PrecioVenta"), 0) * (MonInfoOld.CambioA / MonInfoNew.CambioA)
            End If
            If data.Row.ContainsKey("ImpOfertaVenta") Then
                data.Row("ImpOfertaVenta") = Nz(data.Row("ImpOfertaVenta"), 0) * (MonInfoOld.CambioA / MonInfoNew.CambioA)
            End If
            If data.Row.ContainsKey("ImpCosteOferta") Then
                data.Row("ImpCosteOferta") = Nz(data.Row("ImpCosteOferta"), 0) * (MonInfoOld.CambioA / MonInfoNew.CambioA)
            End If
        End If
    End Sub

#End Region

#Region " INFORMACION DE SISTEMA: DATABASES "

    Public Function DataBases() As DataTable
        Dim dt As DataTable = New BE.DataEngine().Filter("xDataBase", "*", "", , , True)
        dt.DefaultView.Sort = "IDBaseDatos"
        Return dt
    End Function

    Public Function GetDataBase(ByVal IDBaseDatos As Guid) As DataRow
        Dim dt As DataTable = New BE.DataEngine().Filter("xDataBase", New GuidFilterItem("IdBaseDatos", IDBaseDatos), , , , True)
        If dt.Rows.Count > 0 Then
            Return dt.Rows(0)
        End If
    End Function

    Public Function GetDataBase(ByVal IDBaseDatos As Guid, ByVal sortedDatabases As DataView) As DataRow
        Dim i As Integer
        i = sortedDatabases.Find(IDBaseDatos)
        If i >= 0 Then
            Return sortedDatabases(i).Row
        End If
    End Function

    Public Function GetDataBaseDescription(ByVal IDBaseDatos As Guid) As String
        Dim dataBase As DataRow = New NegocioGeneral().GetDataBase(IDBaseDatos)
        If Not dataBase Is Nothing Then
            Return Nz(dataBase("DescBaseDatos"))
        End If
    End Function

    Public Function GetDataBaseDescription(ByVal IDBaseDatos As Guid, ByVal sortedDatabases As DataView) As String
        Dim dataBase As DataRow = New NegocioGeneral().GetDataBase(IDBaseDatos, sortedDatabases)
        If Not dataBase Is Nothing Then
            Return dataBase("DescBaseDatos") & String.Empty
        End If
    End Function

#End Region

#Region " EXPORTACION FACTURAS "

    Public Function CrearDTExportacionObraTrabajo() As DataTable
        Dim dt As New DataTable
        dt.RemotingFormat = SerializationFormat.Binary
        dt.Columns.Add("IDObra", GetType(Integer))
        dt.Columns.Add("NObra", GetType(String))
        dt.Columns.Add("IDTrabajo", GetType(Integer))
        dt.Columns.Add("CodTrabajo", GetType(String))
        Return dt
    End Function

    Public Function CrearDTExportacionCuentas() As DataTable
        Dim dt As New DataTable
        dt.RemotingFormat = SerializationFormat.Binary
        dt.Columns.Add("IDEjercicio", GetType(String))
        dt.Columns.Add("IDCContable", GetType(String))
        Return dt
    End Function

    Public Function ExportarPlanContableFactura(ByVal strEjercicioDestino As String, ByVal dtPlanContableOrigen As DataTable, ByVal dtCContablesExportar As DataTable) As DataTable
        Dim objFilter As New Filter
        Dim objNegPlanContable As BusinessHelper = BusinessHelper.CreateBusinessObject("PlanContable")
        Dim strIDEjercicioActual As String
        Dim strIDEjercicioAnterior As String
        Dim dtEjercicio As DataTable
        Dim strIN As String = String.Empty
        Dim dtPlanContableDestinoAñadir As DataTable

        dtPlanContableDestinoAñadir = dtPlanContableOrigen.Clone
        If Not IsNothing(dtCContablesExportar) AndAlso dtCContablesExportar.Rows.Count > 0 Then
            '//Comprobar que el Ejercicio existe en la BD Destino.
            Dim objNegEjercicio As BusinessHelper = BusinessHelper.CreateBusinessObject("EjercicioContable")
            dtEjercicio = objNegEjercicio.SelOnPrimaryKey(strEjercicioDestino)


            Dim dtPlanContableDestino As DataTable
            If Not IsNothing(dtEjercicio) AndAlso dtEjercicio.Rows.Count > 0 Then
                '//Recuperamos el Plan Contable Destino completo, para no estar accediendo a la BD continuamente.
                objFilter.Clear()
                objFilter.Add("IDEjercicio", strEjercicioDestino)
                dtPlanContableDestino = objNegPlanContable.Filter(objFilter)
                strIN = String.Empty
            Else
                '//1453: El Ejercicio introducido no existe en la Base de Datos.
                ApplicationService.GenerateError("El Ejercicio introducido no existe en la Base de Datos.")
            End If

            For Each drExportar As DataRow In dtCContablesExportar.Select(Nothing, "IDEjercicio,IDCContable")
                strIDEjercicioActual = drExportar("IDEjercicio")

                If Length(drExportar("IDCContable") & String.Empty) > 0 AndAlso InStr(strIN, drExportar("IDCContable") & String.Empty, CompareMethod.Text) = 0 Then
                    If Len(strIN) > 0 Then strIN = strIN & ","
                    strIN = strIN & drExportar("IDCContable") & String.Empty
                    objFilter.Clear()
                    objFilter.Add(New StringFilterItem("IDEjercicio", strEjercicioDestino))
                    objFilter.Add(New StringFilterItem("IDCContable", drExportar("IDCContable") & String.Empty))
                    '//Comprobar si la C.Contable, está en el Plan Contable de la BD Destino
                    Dim WhereCtaEnDestino As String = objFilter.Compose(New AdoFilterComposer)
                    Dim adr() As DataRow = dtPlanContableDestino.Select(WhereCtaEnDestino)
                    If IsNothing(adr) OrElse adr.Length = 0 Then
                        '//La C.Contable | no existe en el Plan Contable. La añadimos al DataTable de nuevas Cuentas
                        objFilter.Clear()
                        objFilter.Add(New StringFilterItem("IDEjercicio", drExportar("IDEjercicio")))
                        objFilter.Add(New StringFilterItem("IDCContable", drExportar("IDCContable") & String.Empty))
                        Dim WhereCtaEnOrigen As String = objFilter.Compose(New AdoFilterComposer)
                        adr = dtPlanContableOrigen.Select(WhereCtaEnOrigen)
                        If Not IsNothing(adr) AndAlso adr.Length > 0 Then
                            Dim drNuevaCuenta As DataRow = dtPlanContableDestinoAñadir.NewRow
                            drNuevaCuenta.ItemArray = adr(0).ItemArray
                            drNuevaCuenta("IDEjercicio") = strEjercicioDestino
                            drNuevaCuenta("IDContador") = System.DBNull.Value
                            dtPlanContableDestinoAñadir.Rows.Add(drNuevaCuenta)
                            adr = Nothing
                        Else
                            Dim drNuevaCuenta As DataRow = dtPlanContableDestinoAñadir.NewRow
                            drNuevaCuenta("IDEjercicio") = strEjercicioDestino
                            drNuevaCuenta("IDCContable") = drExportar("IDCContable") & String.Empty
                            drNuevaCuenta("DescCContable") = "Cta. Exportada"
                            drNuevaCuenta("Venta") = 0
                            drNuevaCuenta("Gasto") = 0
                            drNuevaCuenta("Compra") = 0
                            drNuevaCuenta("Inversion") = 0
                            ' drNuevaCuenta("Balance") = 0
                            drNuevaCuenta("Auxiliar") = 0
                            drNuevaCuenta("Analitica") = 0
                            drNuevaCuenta("Propietario") = 0
                            drNuevaCuenta("IDContador") = System.DBNull.Value
                            dtPlanContableDestinoAñadir.Rows.Add(drNuevaCuenta)
                        End If
                    End If
                End If
                strIDEjercicioAnterior = strIDEjercicioActual
            Next drExportar

            If Not IsNothing(dtPlanContableDestinoAñadir) AndAlso dtPlanContableDestinoAñadir.Rows.Count > 0 Then
                dtPlanContableDestinoAñadir = objNegPlanContable.Update(dtPlanContableDestinoAñadir)
            End If
        End If

        Return dtPlanContableDestinoAñadir
    End Function


    Public Function ExportarObraTrabajoFactura(ByVal dtObrasTrabajosExp As DataTable, _
                                                ByVal dtFactLinea As DataTable) As DataTable

        Dim objFilter As New Filter
        Dim strNObra, strCodTrabajo As String
        Dim adr() As DataRow

        If Not IsNothing(dtFactLinea) AndAlso dtFactLinea.Rows.Count > 0 Then
            Dim objNegObra, objNegObraTrabajo As BusinessHelper
            objNegObra = BusinessHelper.CreateBusinessObject(AdminData.GetEntityInfo("ObraCabecera"))
            objNegObraTrabajo = BusinessHelper.CreateBusinessObject(AdminData.GetEntityInfo("ObraTrabajo"))

            For Each drRowFactLinea As DataRow In dtFactLinea.Select
                If Length(drRowFactLinea("IDObra") & String.Empty) > 0 Then
                    strNObra = String.Empty
                    strCodTrabajo = String.Empty

                    '//Buscamos en el DataTable del Origen el NObra
                    objFilter.Clear()
                    objFilter.Add(New NumberFilterItem("IDObra", drRowFactLinea("IDObra")))
                    Dim WhereObra As String = objFilter.Compose(New AdoFilterComposer)
                    adr = dtObrasTrabajosExp.Select(WhereObra)
                    If Not IsNothing(adr) AndAlso adr.Length > 0 Then
                        strNObra = adr(0)("NObra")
                    End If
                    objFilter.Clear()
                    objFilter.Add(New StringFilterItem("NObra", strNObra))
                    '// Comprobamos que existe la Obra, si no existe cancelamos.
                    Dim dtObra As DataTable = objNegObra.Filter(objFilter)
                    If IsNothing(dtObra) OrElse dtObra.Rows.Count = 0 Then
                        ''6616,"La Obra no existe."
                        'ApplicationService.GenerateError(6616)
                        ApplicationService.GenerateError("Compruebe que las Obras de las Facturas existen en la BD Destino.|Nº FACTURA: ||OBRA: |", vbNewLine, drRowFactLinea("NFactura"), vbNewLine, strNObra)
                    Else
                        drRowFactLinea("IDObra") = dtObra.Rows(0)("IDObra")
                        dtObra.Rows.Clear()
                        If Length(drRowFactLinea("IDTrabajo") & String.Empty) > 0 Then
                            '//Buscamos en el DataTable del Origen el CodTrabajo
                            objFilter.Clear()
                            objFilter.Add(New NumberFilterItem("IDTrabajo", drRowFactLinea("IDTrabajo")))
                            Dim WhereTrabajo As String = objFilter.Compose(New AdoFilterComposer)
                            adr = dtObrasTrabajosExp.Select(WhereTrabajo)
                            If Not IsNothing(adr) AndAlso adr.Length > 0 Then
                                strCodTrabajo = adr(0)("CodTrabajo")
                            End If

                            If Length(strCodTrabajo) > 0 Then
                                '// Comprobamos que existe el Trabajo, si no existe cancelamos.
                                objFilter.Clear()
                                objFilter.Add(New NumberFilterItem("IDObra", drRowFactLinea("IDObra")))
                                objFilter.Add(New StringFilterItem("CodTrabajo", strCodTrabajo))

                                Dim dtTrabajo As DataTable = objNegObraTrabajo.Filter(objFilter)
                                If IsNothing(dtTrabajo) OrElse dtTrabajo.Rows.Count = 0 Then
                                    '//12509,"El código de trabajo no existe o no corresponde al código de Obra de la Línea."
                                    'ApplicationService.GenerateError(12509)
                                    ApplicationService.GenerateError("El código de trabajo no existe o no corresponde al código de Obra de la Línea.|Nº FACTURA: ||COD.TRABAJO: |", vbNewLine, drRowFactLinea("NFactura"), vbNewLine, strCodTrabajo)
                                Else
                                    drRowFactLinea("IDTrabajo") = dtTrabajo.Rows(0)("IDTrabajo")
                                    dtTrabajo.Rows.Clear()
                                End If
                            End If
                        End If
                    End If
                End If
            Next drRowFactLinea
        End If

        Return dtFactLinea
    End Function

    Public Function PrepararObrasTrabajosExportacion(ByVal dtLineas As DataTable) As DataTable
        Dim objNegObra As BusinessHelper
        Dim objNegObraTrabajo As BusinessHelper
        Dim dtObrasTrabajosExp As DataTable = CrearDTExportacionObraTrabajo()
        Dim drObrasTrabajosExp As DataRow

        If Not IsNothing(dtLineas) AndAlso dtLineas.Rows.Count > 0 Then
            objNegObra = BusinessHelper.CreateBusinessObject(AdminData.GetEntityInfo("ObraCabecera"))
            objNegObraTrabajo = BusinessHelper.CreateBusinessObject(AdminData.GetEntityInfo("ObraTrabajo"))
        End If

        Dim dtObra, dtTrabajo As DataTable
        For Each drRowLineas As DataRow In dtLineas.Rows
            If Length(drRowLineas("IDObra") & String.Empty) > 0 Then
                drObrasTrabajosExp = dtObrasTrabajosExp.NewRow
                dtObra = objNegObra.SelOnPrimaryKey(drRowLineas("IDObra"))
                If Not IsNothing(dtObra) AndAlso dtObra.Rows.Count > 0 Then
                    drObrasTrabajosExp("IDObra") = drRowLineas("IDObra")
                    drObrasTrabajosExp("NObra") = dtObra.Rows(0)("NObra")

                    If Length(drRowLineas("IDTrabajo") & String.Empty) > 0 Then
                        dtTrabajo = objNegObraTrabajo.SelOnPrimaryKey(drRowLineas("IDTrabajo"))
                        If Not IsNothing(dtTrabajo) AndAlso dtTrabajo.Rows.Count > 0 Then
                            drObrasTrabajosExp("IDTrabajo") = drRowLineas("IDTrabajo")
                            drObrasTrabajosExp("CodTrabajo") = dtTrabajo.Rows(0)("CodTrabajo")
                            dtTrabajo.Rows.Clear()
                        End If
                    End If

                    dtObrasTrabajosExp.Rows.Add(drObrasTrabajosExp)
                    dtObra.Rows.Clear()
                End If
            End If
        Next drRowLineas

        objNegObra = Nothing
        objNegObraTrabajo = Nothing

        Return dtObrasTrabajosExp
    End Function

#End Region

#Region "FUNCIONES PARA EL MANTENIMIENTO DE PRECIOS E IMPORTES EN LAS MONEDAS A Y B"

    <Task()> Public Shared Sub AplicarDecimales(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
        Dim MonInfo As MonedaInfo
        If data.Table.Columns.Contains("IDMoneda") AndAlso Length(data("IDMoneda")) > 0 Then
            MonInfo = Monedas.GetMoneda(data("IDMoneda"))
        Else
            MonInfo = Monedas.MonedaA
        End If

        Dim ValAyB As New ValoresAyB(New DataRowPropertyAccessor(data), MonInfo.ID, MonInfo.CambioA, MonInfo.CambioB)
        ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf MantenimientoValoresAyB, ValAyB, services)
    End Sub

    <Task()> Public Shared Sub AplicarDecimalesMoneda(ByVal data As DataAplicarDecimalesMoneda, ByVal services As ServiceProvider)
        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
        Dim MonInfo As MonedaInfo
        If Length(data.IDMoneda) > 0 Then
            MonInfo = Monedas.GetMoneda(data.IDMoneda, data.Fecha)
        Else
            MonInfo = Monedas.MonedaA
        End If

        Dim ValAyB As New ValoresAyB(New DataRowPropertyAccessor(data.Row), MonInfo.ID, MonInfo.CambioA, MonInfo.CambioB)
        ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf MantenimientoValoresAyB, ValAyB, services)
    End Sub

    '//Sólo utilizar desde negocios.
    <Task()> Public Shared Function MantenimientoValoresAyB(ByVal dataAB As ValoresAyB, ByVal services As ServiceProvider) As IPropertyAccessor
        Dim data As IPropertyAccessor = dataAB.Linea
        Dim IDMoneda As String
        If Length(dataAB.IDMoneda) > 0 Then
            IDMoneda = dataAB.IDMoneda
        ElseIf data.ContainsKey("IDMoneda") Then
            IDMoneda = data("IDMoneda")
        Else
            Exit Function
        End If

        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
        Dim Moneda As MonedaInfo = Monedas.GetMoneda(IDMoneda)
        If dataAB.CambioA <> 0 Then
            Moneda.CambioA = dataAB.CambioA
            Moneda.CambioB = dataAB.CambioB
        End If


        Dim MonedaA As MonedaInfo = Monedas.MonedaA
        Dim MonedaB As MonedaInfo = Monedas.MonedaB

        If data.ContainsKey("Precio") Then
            data("Precio") = Nz(data("Precio"), 0)
            data("PrecioA") = xRound(data("Precio") * Moneda.CambioA, MonedaA.NDecimalesPrecio)
            data("PrecioB") = xRound(data("Precio") * Moneda.CambioB, MonedaB.NDecimalesPrecio)
            data("Precio") = xRound(data("Precio"), Moneda.NDecimalesPrecio)
        End If
        If data.ContainsKey("Importe") Then
            data("Importe") = Nz(data("Importe"), 0)
            data("ImporteA") = xRound(data("Importe") * Moneda.CambioA, MonedaA.NDecimalesImporte)
            data("ImporteB") = xRound(data("Importe") * Moneda.CambioB, MonedaB.NDecimalesImporte)
            data("Importe") = xRound(Nz(data("Importe"), 0), Moneda.NDecimalesImporte)
        End If
        If data.ContainsKey("ImporteAmortizado") Then
            data("ImporteAmortizado") = Nz(data("ImporteAmortizado"), 0)
            data("ImporteAmortizadoA") = xRound(data("ImporteAmortizado") * Moneda.CambioA, MonedaA.NDecimalesImporte)
            data("ImporteAmortizadoB") = xRound(data("ImporteAmortizado") * Moneda.CambioB, MonedaB.NDecimalesImporte)
            data("ImporteAmortizado") = xRound(Nz(data("ImporteAmortizado"), 0), Moneda.NDecimalesImporte)
        End If
        If data.ContainsKey("ImpLineas") Then
            data("ImpLineas") = Nz(data("ImpLineas"), 0)
            data("ImpLineasA") = xRound(data("ImpLineas") * Moneda.CambioA, MonedaA.NDecimalesImporte)
            data("ImpLineasB") = xRound(data("ImpLineas") * Moneda.CambioB, MonedaB.NDecimalesImporte)
            data("ImpLineas") = xRound(Nz(data("ImpLineas"), 0), Moneda.NDecimalesImporte)
        End If
        If data.ContainsKey("ImpIVA") Then
            data("ImpIVA") = Nz(data("ImpIVA"), 0)
            data("ImpIVAA") = xRound(data("ImpIVA") * Moneda.CambioA, MonedaA.NDecimalesImporte)
            data("ImpIVAB") = xRound(data("ImpIVA") * Moneda.CambioB, MonedaB.NDecimalesImporte)
            data("ImpIVA") = xRound(Nz(data("ImpIVA"), 0), Moneda.NDecimalesImporte)
        End If
        If data.ContainsKey("ImpSinRepercutir") Then
            data("ImpSinRepercutir") = Nz(data("ImpSinRepercutir"), 0)
            data("ImpSinRepercutirA") = xRound(data("ImpSinRepercutir") * Moneda.CambioA, MonedaA.NDecimalesImporte)
            data("ImpSinRepercutirB") = xRound(data("ImpSinRepercutir") * Moneda.CambioB, MonedaB.NDecimalesImporte)
            data("ImpSinRepercutir") = xRound(Nz(data("ImpSinRepercutir"), 0), Moneda.NDecimalesImporte)
        End If
        If data.ContainsKey("ImpRE") Then
            data("ImpRE") = Nz(data("ImpRE"), 0)
            data("ImpREA") = xRound(data("ImpRE") * Moneda.CambioA, MonedaA.NDecimalesImporte)
            data("ImpREB") = xRound(data("ImpRE") * Moneda.CambioB, MonedaB.NDecimalesImporte)
            data("ImpRE") = xRound(Nz(data("ImpRE"), 0), Moneda.NDecimalesImporte)
        End If
        If data.ContainsKey("ImpDto") Then
            data("ImpDto") = Nz(data("ImpDto"), 0)
            data("ImpDtoA") = xRound(data("ImpDto") * Moneda.CambioA, MonedaA.NDecimalesImporte)
            data("ImpDtoB") = xRound(data("ImpDto") * Moneda.CambioB, MonedaB.NDecimalesImporte)
            data("ImpDto") = xRound(Nz(data("ImpDto"), 0), Moneda.NDecimalesImporte)
        End If
        If data.ContainsKey("ImpTotal") Then
            data("ImpTotal") = Nz(data("ImpTotal"), 0)
            data("ImpTotalA") = xRound(data("ImpTotal") * Moneda.CambioA, MonedaA.NDecimalesImporte)
            data("ImpTotalB") = xRound(data("ImpTotal") * Moneda.CambioB, MonedaB.NDecimalesImporte)
            data("ImpTotal") = xRound(Nz(data("ImpTotal"), 0), Moneda.NDecimalesImporte)
        End If
        If data.ContainsKey("ImportePVP") Then
            data("ImportePVP") = Nz(data("ImportePVP"), 0)
            data("ImportePVPA") = xRound(data("ImportePVP") * Moneda.CambioA, MonedaA.NDecimalesImporte)
            data("ImportePVPB") = xRound(data("ImportePVP") * Moneda.CambioB, MonedaB.NDecimalesImporte)
            data("ImportePVP") = xRound(Nz(data("ImportePVP"), 0), Moneda.NDecimalesImporte)
        End If
        If data.ContainsKey("PVP") Then
            data("PVP") = Nz(data("PVP"), 0)
            data("PVPA") = xRound(data("PVP") * Moneda.CambioA, MonedaA.NDecimalesImporte)
            data("PVPB") = xRound(data("PVP") * Moneda.CambioB, MonedaB.NDecimalesImporte)
            data("PVP") = xRound(Nz(data("PVP"), 0), Moneda.NDecimalesImporte)
        End If
        If data.ContainsKey("BaseImponible") Then
            data("BaseImponible") = Nz(data("BaseImponible"), 0)
            data("BaseImponibleA") = xRound(data("BaseImponible") * Moneda.CambioA, MonedaA.NDecimalesImporte)
            data("BaseImponibleB") = xRound(data("BaseImponible") * Moneda.CambioB, MonedaB.NDecimalesImporte)
            data("BaseImponible") = xRound(Nz(data("BaseImponible"), 0), Moneda.NDecimalesImporte)
        End If
        If data.ContainsKey("ImpAlbaran") Then
            data("ImpAlbaran") = Nz(data("ImpAlbaran"), 0)
            data("ImpAlbaranA") = xRound(data("ImpAlbaran") * Moneda.CambioA, MonedaA.NDecimalesImporte)
            data("ImpAlbaranB") = xRound(data("ImpAlbaran") * Moneda.CambioB, MonedaB.NDecimalesImporte)
            data("ImpAlbaran") = xRound(Nz(data("ImpAlbaran"), 0), Moneda.NDecimalesImporte)
        End If
        If data.ContainsKey("BaseImponibleEspecial") Then
            data("BaseImponibleEspecial") = Nz(data("BaseImponibleEspecial"), 0)
            data("BaseImponibleEspecialA") = xRound(data("BaseImponibleEspecial") * Moneda.CambioA, MonedaA.NDecimalesImporte)
            data("BaseImponibleEspecialB") = xRound(data("BaseImponibleEspecial") * Moneda.CambioB, MonedaB.NDecimalesImporte)
            data("BaseImponibleEspecial") = xRound(Nz(data("BaseImponibleEspecial"), 0), Moneda.NDecimalesImporte)
        End If
        If data.ContainsKey("BaseImponibleNormal") Then
            data("BaseImponibleNormal") = Nz(data("BaseImponibleNormal"), 0)
            data("BaseImponibleNormalA") = xRound(data("BaseImponibleNormal") * Moneda.CambioA, MonedaA.NDecimalesImporte)
            data("BaseImponibleNormalB") = xRound(data("BaseImponibleNormal") * Moneda.CambioB, MonedaB.NDecimalesImporte)
            data("BaseImponibleNormal") = xRound(Nz(data("BaseImponibleNormal"), 0), Moneda.NDecimalesImporte)
        End If
        If data.ContainsKey("ImpIntrastat") Then
            data("ImpIntrastat") = Nz(data("ImpIntrastat"), 0)
            data("ImpIntrastatA") = xRound(data("ImpIntrastat") * Moneda.CambioA, MonedaA.NDecimalesImporte)
            data("ImpIntrastatB") = xRound(data("ImpIntrastat") * Moneda.CambioB, MonedaB.NDecimalesImporte)
            data("ImpIntrastat") = xRound(Nz(data("ImpIntrastat"), 0), Moneda.NDecimalesImporte)
        End If
        If data.ContainsKey("ImpDtoFactura") Then
            data("ImpDtoFactura") = Nz(data("ImpDtoFactura"), 0)
            data("ImpDtoFacturaA") = xRound(data("ImpDtoFactura") * Moneda.CambioA, MonedaA.NDecimalesImporte)
            data("ImpDtoFacturaB") = xRound(data("ImpDtoFactura") * Moneda.CambioB, MonedaB.NDecimalesImporte)
            data("ImpDtoFactura") = xRound(Nz(data("ImpDtoFactura"), 0), Moneda.NDecimalesImporte)
        End If
        If data.ContainsKey("ImpDpp") Then
            data("ImpDpp") = Nz(data("ImpDpp"), 0)
            data("ImpDppA") = xRound(data("ImpDpp") * Moneda.CambioA, MonedaA.NDecimalesImporte)
            data("ImpDppB") = xRound(data("ImpDpp") * Moneda.CambioB, MonedaB.NDecimalesImporte)
            data("ImpDpp") = xRound(Nz(data("ImpDpp"), 0), Moneda.NDecimalesImporte)
        End If
        If data.ContainsKey("ImpRecFinan") Then
            data("ImpRecFinan") = Nz(data("ImpRecFinan"), 0)
            data("ImpRecFinanA") = xRound(data("ImpRecFinan") * Moneda.CambioA, MonedaA.NDecimalesImporte)
            data("ImpRecFinanB") = xRound(data("ImpRecFinan") * Moneda.CambioB, MonedaB.NDecimalesImporte)
            data("ImpRecFinan") = xRound(Nz(data("ImpRecFinan"), 0), Moneda.NDecimalesImporte)
        End If
        If data.ContainsKey("BaseRetencion") Then
            data("BaseRetencion") = Nz(data("BaseRetencion"), 0)
            data("BaseRetencionA") = xRound(data("BaseRetencion") * Moneda.CambioA, MonedaA.NDecimalesImporte)
            data("BaseRetencionB") = xRound(data("BaseRetencion") * Moneda.CambioB, MonedaB.NDecimalesImporte)
            data("BaseRetencion") = xRound(Nz(data("BaseRetencion"), 0), Moneda.NDecimalesImporte)
        End If
        If data.ContainsKey("ImpRetencion") Then
            data("ImpRetencion") = Nz(data("ImpRetencion"), 0)
            data("ImpRetencionA") = xRound(data("ImpRetencion") * Moneda.CambioA, MonedaA.NDecimalesImporte)
            data("ImpRetencionB") = xRound(data("ImpRetencion") * Moneda.CambioB, MonedaB.NDecimalesImporte)
            data("ImpRetencion") = xRound(Nz(data("ImpRetencion"), 0), Moneda.NDecimalesImporte)
        End If
        If data.ContainsKey("ImpVencimiento") Then
            data("ImpVencimiento") = Nz(data("ImpVencimiento"), 0)
            data("ImpVencimientoA") = xRound(data("ImpVencimiento") * Moneda.CambioA, MonedaA.NDecimalesImporte)
            data("ImpVencimientoB") = xRound(data("ImpVencimiento") * Moneda.CambioB, MonedaB.NDecimalesImporte)
            data("ImpVencimiento") = xRound(Nz(data("ImpVencimiento"), 0), Moneda.NDecimalesImporte)
        End If
        If data.ContainsKey("RecargoFinanciero") Then
            data("RecargoFinanciero") = Nz(data("RecargoFinanciero"), 0)
            data("RecargoFinancieroA") = xRound(data("RecargoFinanciero") * Moneda.CambioA, MonedaA.NDecimalesImporte)
            data("RecargoFinancieroB") = xRound(data("RecargoFinanciero") * Moneda.CambioB, MonedaB.NDecimalesImporte)
            data("RecargoFinanciero") = xRound(Nz(data("RecargoFinanciero"), 0), Moneda.NDecimalesImporte)
        End If
        If data.ContainsKey("ImpPedido") Then
            data("ImpPedido") = Nz(data("ImpPedido"), 0)
            data("ImpPedidoA") = xRound(data("ImpPedido") * Moneda.CambioA, MonedaA.NDecimalesImporte)
            data("ImpPedidoB") = xRound(data("ImpPedido") * Moneda.CambioB, MonedaB.NDecimalesImporte)
            data("ImpPedido") = xRound(Nz(data("ImpPedido"), 0), Moneda.NDecimalesImporte)
        End If
        If data.ContainsKey("ImpPuntoVerde") Then
            data("ImpPuntoVerde") = Nz(data("ImpPuntoVerde"), 0)
            data("ImpPuntoVerdeA") = xRound(data("ImpPuntoVerde") * Moneda.CambioA, MonedaA.NDecimalesImporte)
            data("ImpPuntoVerdeB") = xRound(data("ImpPuntoVerde") * Moneda.CambioB, MonedaB.NDecimalesImporte)
            data("ImpPuntoVerde") = xRound(Nz(data("ImpPuntoVerde"), 0), Moneda.NDecimalesImporte)
        End If
        If data.ContainsKey("ARepercutir") Then
            data("ARepercutir") = Nz(data("ARepercutir"), 0)
            data("ARepercutirA") = xRound(data("ARepercutir") * Moneda.CambioA, MonedaA.NDecimalesImporte)
            data("ARepercutirB") = xRound(data("ARepercutir") * Moneda.CambioB, MonedaB.NDecimalesImporte)
            data("ARepercutir") = xRound(Nz(data("ARepercutir"), 0), Moneda.NDecimalesImporte)
        End If
        If data.ContainsKey("CosteUnitario") Then
            data("CosteUnitario") = Nz(data("CosteUnitario"), 0)
            data("CosteUnitarioA") = xRound(data("CosteUnitario") * Moneda.CambioA, MonedaA.NDecimalesPrecio)
            data("CosteUnitarioB") = xRound(data("CosteUnitario") * Moneda.CambioB, MonedaB.NDecimalesPrecio)
            data("CosteUnitario") = xRound(Nz(data("CosteUnitario"), 0), Moneda.NDecimalesPrecio)
        End If
        If data.ContainsKey("PrecioVenta") Then
            data("PrecioVenta") = Nz(data("PrecioVenta"), 0)
            data("PrecioVentaA") = xRound(data("PrecioVenta") * Moneda.CambioA, MonedaA.NDecimalesPrecio)
            data("PrecioVentaB") = xRound(data("PrecioVenta") * Moneda.CambioB, MonedaB.NDecimalesPrecio)
            data("PrecioVenta") = xRound(Nz(data("PrecioVenta"), 0), Moneda.NDecimalesPrecio)
        End If
        If data.ContainsKey("ImpOfertaVenta") Then
            data("ImpOfertaVenta") = Nz(data("ImpOfertaVenta"), 0)
            data("ImpOfertaVentaA") = xRound(data("ImpOfertaVenta") * Moneda.CambioA, MonedaA.NDecimalesImporte)
            data("ImpOfertaVentaB") = xRound(data("ImpOfertaVenta") * Moneda.CambioB, MonedaB.NDecimalesImporte)
            data("ImpOfertaVenta") = xRound(Nz(data("ImpOfertaVenta"), 0), Moneda.NDecimalesImporte)
        End If
        If data.ContainsKey("ImpCosteOferta") Then
            data("ImpCosteOferta") = Nz(data("ImpCosteOferta"), 0)
            data("ImpCosteOfertaA") = xRound(data("ImpCosteOferta") * Moneda.CambioA, MonedaA.NDecimalesImporte)
            data("ImpCosteOfertaB") = xRound(data("ImpCosteOferta") * Moneda.CambioB, MonedaB.NDecimalesImporte)
            data("ImpCosteOferta") = xRound(Nz(data("ImpCosteOferta"), 0), Moneda.NDecimalesImporte)
        End If
        If data.ContainsKey("CosteVariosA") Then
            data("CosteVariosA") = Nz(data("CosteVariosA"), 0)
            data("CosteVariosB") = xRound(data("CosteVariosA") * Moneda.CambioB, MonedaB.NDecimalesPrecio)
            data("CosteVariosA") = xRound(data("CosteVariosA"), MonedaA.NDecimalesPrecio)
        End If
        If data.ContainsKey("PrecioStdA") Then
            data("PrecioStdA") = Nz(data("PrecioStdA"), 0)
            data("PrecioStdB") = xRound(data("PrecioStdA") * Moneda.CambioB, MonedaB.NDecimalesPrecio)
            data("PrecioStdA") = xRound(data("PrecioStdA"), MonedaA.NDecimalesPrecio)
        End If
        If data.ContainsKey("CosteStdA") Then
            data("CosteStdA") = Nz(data("CosteStdA"), 0)
            data("CosteStdB") = xRound(data("CosteStdA") * Moneda.CambioB, MonedaB.NDecimalesPrecio)
            data("CosteStdA") = xRound(data("CosteStdA"), MonedaA.NDecimalesPrecio)
        End If
        If data.ContainsKey("CosteMatStdA") Then
            data("CosteMatStdA") = Nz(data("CosteMatStdA"), 0)
            data("CosteMatStdB") = xRound(data("CosteMatStdA") * Moneda.CambioB, MonedaB.NDecimalesPrecio)
            data("CosteMatStdA") = xRound(data("CosteMatStdA"), MonedaA.NDecimalesPrecio)
        End If
        If data.ContainsKey("CosteOpeStdA") Then
            data("CosteOpeStdA") = Nz(data("CosteOpeStdA"), 0)
            data("CosteOpeStdB") = xRound(data("CosteOpeStdA") * Moneda.CambioB, MonedaB.NDecimalesPrecio)
            data("CosteOpeStdA") = xRound(data("CosteOpeStdA"), MonedaA.NDecimalesPrecio)
        End If
        If data.ContainsKey("CosteExtStdA") Then
            data("CosteExtStdA") = Nz(data("CosteExtStdA"), 0)
            data("CosteExtStdB") = xRound(data("CosteExtStdA") * Moneda.CambioB, MonedaB.NDecimalesPrecio)
            data("CosteExtStdA") = xRound(data("CosteExtStdA"), MonedaA.NDecimalesPrecio)
        End If
        If data.ContainsKey("CosteVarStdA") Then
            data("CosteVarStdA") = Nz(data("CosteVarStdA"), 0)
            data("CosteVarStdB") = xRound(data("CosteVarStdA") * Moneda.CambioB, MonedaB.NDecimalesPrecio)
            data("CosteVarStdA") = xRound(data("CosteVarStdA"), MonedaA.NDecimalesPrecio)
        End If
        If data.ContainsKey("EjecucionValorA") Then
            data("EjecucionValorA") = Nz(data("EjecucionValorA"), 0)
            data("EjecucionValorB") = xRound(data("EjecucionValorA") * Moneda.CambioB, MonedaB.NDecimalesPrecio)
            data("EjecucionValorA") = xRound(data("EjecucionValorA"), MonedaA.NDecimalesPrecio)
        End If
        If data.ContainsKey("PreparacionValorA") Then
            data("PreparacionValorA") = Nz(data("PreparacionValorA"), 0)
            data("PreparacionValorB") = xRound(data("PreparacionValorA") * Moneda.CambioB, MonedaB.NDecimalesPrecio)
            data("PreparacionValorA") = xRound(data("PreparacionValorA"), MonedaA.NDecimalesPrecio)
        End If
        If data.ContainsKey("ManoObraValorA") Then
            data("ManoObraValorA") = Nz(data("ManoObraValorA"), 0)
            data("ManoObraValorB") = xRound(data("ManoObraValorA") * Moneda.CambioB, MonedaB.NDecimalesPrecio)
            data("ManoObraValorA") = xRound(data("ManoObraValorA"), MonedaA.NDecimalesPrecio)
        End If
        If data.ContainsKey("PrecioOfertado") Then
            data("PrecioOfertado") = xRound(data("PrecioOfertado"), MonedaA.NDecimalesPrecio)
        End If
        If data.ContainsKey("ImporteRemesaAnticipo") Then
            data("ImporteRemesaAnticipo") = Nz(data("ImporteRemesaAnticipo"), 0)
            data("ImporteRemesaAnticipoA") = xRound(data("ImporteRemesaAnticipo") * Moneda.CambioA, MonedaA.NDecimalesImporte)
            data("ImporteRemesaAnticipoB") = xRound(data("ImporteRemesaAnticipo") * Moneda.CambioB, MonedaB.NDecimalesImporte)
            data("ImporteRemesaAnticipo") = xRound(Nz(data("ImporteRemesaAnticipo"), 0), Moneda.NDecimalesImporte)
        End If
        Return data
    End Function

    <Task()> Public Shared Function MantenimientoValoresImporteAyB(ByVal dataAB As ValoresAyB, ByVal services As ServiceProvider) As fImporte
        Dim IDMoneda As String
        If Length(dataAB.IDMoneda) > 0 Then
            IDMoneda = dataAB.IDMoneda
        Else
            Exit Function
        End If

        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
        Dim Moneda As MonedaInfo = Monedas.GetMoneda(IDMoneda)
        If dataAB.CambioA <> 0 Then
            Moneda.CambioA = dataAB.CambioA
            Moneda.CambioB = dataAB.CambioB
        End If

        Dim MonedaA As MonedaInfo = Monedas.MonedaA
        Dim MonedaB As MonedaInfo = Monedas.MonedaB

        Dim fImp As New fImporte
        fImp.ImporteA = xRound(dataAB.Importe * Moneda.CambioA, MonedaA.NDecimalesImporte)
        fImp.ImporteB = xRound(dataAB.Importe * Moneda.CambioB, MonedaB.NDecimalesImporte)
        fImp.Importe = xRound(dataAB.Importe, Moneda.NDecimalesImporte)

        Return fImp
    End Function

#End Region

#Region " Funciones genéricas para los circuitos de Ventas y Compras "

    <Task()> Public Shared Sub FactorConversion(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If (Length(data.Current("IDUDMedida")) = 0 Or Length(data.Current("IDUDInterna")) = 0) Then
            data.Current("Factor") = 1
        ElseIf data.Current("IDUDMedida") & String.Empty = data.Current("IDUDInterna") & String.Empty Then
            data.Current("Factor") = 1
        Else
            Dim StDatos As New ArticuloUnidadAB.DatosFactorConversion
            StDatos.IDArticulo = data.Current("IDArticulo")
            StDatos.IDUdMedidaA = data.Current("IDUDMedida")
            StDatos.IDUdMedidaB = data.Current("IDUDInterna")
            StDatos.UnoSiNoExiste = True
            data.Current("Factor") = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, StDatos, services)

        End If
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf NegocioGeneral.CambioFactor, data, services)
    End Sub

    <Task()> Public Shared Sub AsignarArticuloAlmacen(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim AppParamsGen As ParametroGeneral = services.GetService(Of ParametroGeneral)()

        If Length(data.Current("IDAlmacen")) = 0 OrElse Not AppParamsGen.AlmacenCentroGestionActivo Then
            Dim strCentroGestion As String
            If Not data.Context Is Nothing AndAlso data.Context.ContainsKey("IDCentroGestion") AndAlso Length(data.Context("IDCentroGestion")) > 0 Then
                strCentroGestion = data.Context("IDCentroGestion")
            Else
                If data.Current.ContainsKey("IDCentroGestion") AndAlso Length(data.Current("IDCentroGestion")) > 0 Then
                    strCentroGestion = data.Current("IDCentroGestion")
                End If
            End If
            Dim ArtAlm As New DataArtAlm
            ArtAlm.IDArticulo = data.Current("IDArticulo")
            ArtAlm.IDCentroGestion = strCentroGestion
            Dim StrAlmacen As String = ProcessServer.ExecuteTask(Of DataArtAlm, String)(AddressOf ArticuloAlmacen.AlmacenPredeterminadoArticulo, ArtAlm, services)
            If Length(StrAlmacen) > 0 Then
                data.Current("IDAlmacen") = StrAlmacen
            End If

        End If
        If Length(data.Current("IDAlmacen")) > 0 And Length(data.Current("IDArticulo")) > 0 Then
            Dim FilArtAlm As New Filter
            FilArtAlm.Add("IDArticulo", FilterOperator.Equal, data.Current("IDArticulo"))
            FilArtAlm.Add("IDAlmacen", FilterOperator.Equal, data.Current("IDAlmacen"))
            Dim ClsArtAlm As New ArticuloAlmacen
            Dim DtArtAlm As DataTable = ClsArtAlm.Filter(FilArtAlm)
            If Not DtArtAlm Is Nothing AndAlso DtArtAlm.Rows.Count > 0 Then
                data.Current("StockFisico") = DtArtAlm.Rows(0)("StockFisico")
                data.Current("LoteMinimo") = DtArtAlm.Rows(0)("LoteMinimo")
            Else
                data.Current("StockFisico") = 0
                data.Current("LoteMinimo") = 0
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioQInterna(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If IsNumeric(data.Current("QInterna")) Then
            'If Nz(data.Current("Cantidad"), 0) <> 0 Then
            Dim datFactor As New ArticuloUnidadAB.DatosFactorConversion(data.Current("IDArticulo"), data.Current("IDUDMedida"), data.Current("IDUDInterna"), False)
            Dim Factor As Double = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, datFactor, services)
            If Factor = 0 Then
                If Nz(data.Current("Cantidad"), 0) <> 0 Then data.Current("Factor") = data.Current("QInterna") / data.Current("Cantidad")
            Else
                data.Current("Factor") = Factor
                data.Current("Cantidad") = data.Current("QInterna") / data.Current("Factor")
                '// Para evitar modificar la ubicación de las tareas, por si hay sobreescrituras y debido a que no va a ser necesario en todos los casos,
                '// incluimos en el Current una variable para ver si hay que simular el Cambio Manual del campo cantidad, para el Recalculo de Tarifas 
                '// y otros posibles cambios referentes al cambio de Cantidad.
                data.Current("SimularCambioCantidad") = True
            End If
            'End If
        Else
            ApplicationService.GenerateError("Campo no numérico.")
        End If
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoComunes.ValidarFactorDobleUnidad, data, services)
    End Sub

    <Task()> Public Shared Sub CalculoQInterna(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current("QInterna") = Nz(data.Current("Factor"), 1) * Nz(data.Current("Cantidad"), 0)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoComunes.ValidarFactorDobleUnidad, data, services)
    End Sub

    <Task()> Public Shared Sub CambioFactor(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If data.ColumnName = "Factor" Then data.Current(data.ColumnName) = Nz(data.Value, 1)
        If IsNumeric(data.Current("Factor")) Then
            If data.Current("Factor") < 0 Then
                ApplicationService.GenerateError("El factor no es válido.")
            Else
                ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CalculoQInterna, data, services)
            End If
        Else
            ApplicationService.GenerateError("Campo no numérico.")
        End If
    End Sub

    <Task()> Public Shared Sub CambioUDMedida(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDUDMedida")) > 0 Then
            ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf FactorConversion, data, services)
            ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CalculoQInterna, data, services)
        End If
    End Sub

    <Task()> Public Shared Sub CambioPrecio(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("Precio")) > 0 Then
            If IsNumeric(data.Current("Precio")) Then
                data.Current("SeguimientoTarifa") = AdminData.GetMessageText("MANUAL")
                If data.Context.ContainsKey("IDMoneda") AndAlso Length(data.Context("IDMoneda")) > 0 Then
                    Dim ValAyB As New ValoresAyB(data.Current, data.Context("IDMoneda"), data.Context("CambioA"), data.Context("CambioB"))
                    ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)

                    ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf NegocioGeneral.CalcularImportes, data, services)
                    If data.Current.ContainsKey("PrecioNeto") Then
                        If data.Current("Cantidad") > 0 Then
                            data.Current("PrecioNeto") = data.Current("Importe") / data.Current("Cantidad")
                        Else
                            data.Current("PrecioNeto") = 0
                        End If
                    End If

                End If
            Else
                ApplicationService.GenerateError("Campo no numérico.")
            End If
        End If
    End Sub

    '//CambioCondicionPago para las cabeceras
    <Task()> Public Shared Sub CambioCondicionPago(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If data.ColumnName = "IDCondicionPago" Then data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDCondicionPago")) > 0 Then
            Dim CondicionesPago As EntityInfoCache(Of CondicionPagoInfo) = services.GetService(Of EntityInfoCache(Of CondicionPagoInfo))()
            Dim CondPagoInfo As CondicionPagoInfo = CondicionesPago.GetEntity(data.Current("IDCondicionPago"))
            data.Current("DtoProntoPago") = CondPagoInfo.DtoProntoPago
            If data.Current.ContainsKey("RecFinan") Then data.Current("RecFinan") = CondPagoInfo.RecFinan
        Else
            data.Current("DtoProntoPago") = 0
            If data.Current.ContainsKey("RecFinan") Then data.Current("RecFinan") = 0
        End If
    End Sub

    <Task()> Public Shared Sub CambioImporteVencimiento(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDMoneda")) > 0 Then
            Select Case data.ColumnName
                Case "ImpVencimiento"
                    If Nz(data.Current("CambioA"), 0) <> 0 Then
                        data.Current("ImpVencimientoA") = Nz(data.Current("ImpVencimiento"), 0) * data.Current("CambioA")
                    Else
                        data.Current("ImpVencimientoA") = 0
                    End If
                Case "ImpVencimientoA"
                    If Nz(data.Current("CambioA"), 0) <> 0 Then
                        data.Current("ImpVencimiento") = Nz(data.Current("ImpVencimientoA"), 0) / data.Current("CambioA")
                    Else
                        data.Current("ImpVencimiento") = 0
                    End If
                Case "ImpVencimientoB"
                    If Nz(data.Current("CambioB"), 0) <> 0 Then
                        data.Current("ImpVencimiento") = Nz(data.Current("ImpVencimientoB"), 0) / data.Current("CambioB")
                    Else
                        data.Current("ImpVencimiento") = 0
                    End If
            End Select

            Dim ValAyB As New ValoresAyB(data.Current, data.Current("IDMoneda"), Nz(data.Current("CambioA"), 0), Nz(data.Current("CambioB"), 0))
            ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)
        End If
    End Sub

    <Task()> Public Shared Sub CambioImporteRepercutir(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDMoneda")) > 0 Then
            Dim ValAyB As New ValoresAyB(data.Current, data.Current("IDMoneda"), Nz(data.Current("CambioA"), 0), Nz(data.Current("CambioB"), 0))
            ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)
        End If
    End Sub

    <Task()> Public Shared Sub CambioEnCambiosMoneda(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDMoneda")) > 0 Then
            Dim ValAyB As New ValoresAyB(data.Current, data.Current("IDMoneda"), Nz(data.Current("CambioA"), 0), Nz(data.Current("CambioB"), 0))
            ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)
        End If
    End Sub

    <Task()> Public Shared Sub CambioMonedaFechaVto(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            data.Current(data.ColumnName) = data.Value
            If data.ColumnName = "FechaVencimiento" Then
                If Not IsDate(data.Current("FechaVencimiento")) Then
                    ApplicationService.GenerateError("La Fecha de Vencimiento no es correcta.")
                End If
            End If
            If Length(data.Current("IDFactura")) = 0 Then
                If Length(data.Current("IDMoneda")) > 0 AndAlso Length(data.Current("FechaVencimiento")) > 0 Then
                    Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
                    Dim MonInfo As MonedaInfo = Monedas.GetMoneda(data.Current("IDMoneda") & String.Empty, CDate(data.Current("FechaVencimiento")))
                    data.Current("CambioA") = MonInfo.CambioA
                    data.Current("CambioB") = MonInfo.CambioB
                End If
            End If
            Dim ValAyB As New ValoresAyB(data.Current, data.Current("IDMoneda"), Nz(data.Current("CambioA"), 0), Nz(data.Current("CambioB"), 0))
            ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)
        End If
    End Sub

#Region " Cambio de C.Contable "

    '<Task()> Public Sub CambioCContable(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
    '    data.Current(data.ColumnName) = data.Value
    '    If Length(data.Current("CContable")) > 0 Then ComprobarCContable(data, services)
    'End Sub

    '<Task()> Public Function ComprobarCContable(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
    '    Dim IntEnum As enumCuenta = New EjercicioContable().ValidarAuxiliar(data.Current("CContable"), data.Context("IDEjercicio"))
    '    Select Case IntEnum
    '        Case enumCuenta.cAuxiliar, enumCuenta.cNoAuxiliar
    '            FormatoCuentaContable(data, services)
    '        Case enumCuenta.cNoExisteEnPlan, enumCuenta.cVacia
    '            ApplicationService.GenerateError("La Cuenta Contable no existe.")
    '    End Select
    'End Function

    '<Task()> Public Sub FormatoCuentaContable(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
    '    If data.Current.ContainsKey("CContable") AndAlso Length(data.Current("CContable")) > 0 AndAlso (data.Current("CContable").IndexOf(".") >= 0 OrElse data.Current("CContable").IndexOf(",") >= 0) Then
    '        If Not data.Context.ContainsKey("IDEjercicio") Then
    '            Dim ejercicio As New EjercicioContable
    '            Dim strEjercicio As String
    '            If data.Context.ContainsKey("Fecha") Then
    '                strEjercicio = ejercicio.Predeterminado(data.Context("Fecha"))
    '            ElseIf data.Context.ContainsKey("FechaVencimiento") Then
    '                strEjercicio = ejercicio.Predeterminado(data.Context("FechaVencimiento"))
    '            ElseIf data.Context.ContainsKey("FechaApunte") Then
    '                strEjercicio = ejercicio.Predeterminado(data.Context("FechaApunte"))
    '            Else
    '                strEjercicio = ejercicio.Predeterminado()
    '            End If
    '            If Len(strEjercicio) Then
    '                data.Context("IDEjercicio") = strEjercicio
    '            End If
    '        End If

    '        If data.Context.ContainsKey("IDEjercicio") Then
    '            Dim Ejercicios As EntityInfoCache(Of EjercicioContableInfo) = services.GetService(Of EntityInfoCache(Of EjercicioContableInfo))()
    '            Dim EjercicioInfo As EjercicioContableInfo = Ejercicios.GetEntity(data.Context("IDEjercicio"))
    '            If Not IsNothing(EjercicioInfo) Then
    '                data.Current("CContable") = PuntoPorCero(data.Current("CContable"), EjercicioInfo.DigitosAuxiliar)
    '            End If
    '        End If
    '    End If
    'End Sub

    ''///PROVISIONAL
    'Private Function PuntoPorCero(ByVal pCuenta As String, ByVal pNDigitos As Integer) As String
    '    Dim strCeros As String
    '    Dim strC As String

    '    If InStr(pCuenta, ".") Then
    '        strC = "."
    '    ElseIf InStr(pCuenta, ",") Then
    '        strC = ","
    '    End If

    '    If Len(strC) Then
    '        strCeros = New String("0", pNDigitos - Len(pCuenta) + 1)
    '        pCuenta = Replace(pCuenta, strC, strCeros, , 1)
    '    End If
    '    If InStr(pCuenta, ".") Then pCuenta = Replace(pCuenta, ".", "0")
    '    If InStr(pCuenta, ",") Then pCuenta = Replace(pCuenta, ",", "0")

    '    PuntoPorCero = pCuenta
    'End Function

#Region " Asignar Valores Predeterminados "

    <Task()> Public Shared Sub AsignarEjercicioContable(ByVal data As DataEjercicio, ByVal services As ServiceProvider)
        Dim AppPatamsConta As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
        If Not AppPatamsConta.Contabilidad Then Exit Sub

        If data.Fecha = cnMinDate Then data.Fecha = Today
        'If Length(data.Datos("IDEjercicio")) = 0 Then

        data.Datos("IDEjercicio") = ProcessServer.ExecuteTask(Of Date, String)(AddressOf NegocioGeneral.EjercicioPredeterminado, data.Fecha, services)
        If Length(data.Datos("IDContador")) > 0 Then
            Dim EsContadorB As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf NegocioGeneral.ContadorB, data.Datos("IDContador"), services)
            If EsContadorB Then
                data.Datos("IDEjercicio") = ProcessServer.ExecuteTask(Of Date, String)(AddressOf NegocioGeneral.EjercicioPredeterminadoB, data.Fecha, services)
                If data.Datos.ContainsKey("Enviar347") Then data.Datos("Enviar347") = False
            Else
                'If data.Datos.ContainsKey("Enviar347") Then data.Datos("Enviar347") = True
            End If
        End If
        'End If
    End Sub

    <Task()> Public Shared Sub AsignarCentroGestion(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDCentroGestion")) = 0 Then
            Dim cgu As New UsuarioCentroGestion.UsuarioCentroGestionInfo
            cgu = ProcessServer.ExecuteTask(Of UsuarioCentroGestion.UsuarioCentroGestionInfo, UsuarioCentroGestion.UsuarioCentroGestionInfo)(AddressOf UsuarioCentroGestion.ObtenerUsuarioCentroGestion, cgu, services)
            data("IDCentroGestion") = cgu.IDCentroGestion
        End If
    End Sub

    <Task()> Public Shared Sub AsignarAlmacen(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.IsNull("IDAlmacen") Then
            Dim AppParamsGeneral As ParametroGeneral = services.GetService(Of ParametroGeneral)()
            data("IDAlmacen") = AppParamsGeneral.Almacen
            If AppParamsGeneral.AlmacenCentroGestionActivo AndAlso Length(data("IDCentroGestion")) > 0 Then
                data("IDAlmacen") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Almacen.GetAlmacenCentroGestion, data("IDCentroGestion"), services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarTipoAlbaran(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim AppParamsGeneral As ParametroTesoreria = services.GetService(Of ParametroTesoreria)()
        If data.IsNull("IDTipoAlbaran") Then data("IDTipoAlbaran") = AppParamsGeneral.TipoAlbaranPorDefecto()
    End Sub

    <Task()> Public Shared Function ContadorB(ByVal strIDContador As String, ByVal services As ServiceProvider) As Boolean
        ContadorB = False
        If Length(strIDContador) > 0 Then
            Dim dtContador As DataTable = New Contador().SelOnPrimaryKey(strIDContador)
            If Not dtContador Is Nothing AndAlso dtContador.Rows.Count > 0 Then
                ContadorB = Not dtContador.Rows(0)("AIva")
            Else
                '       ApplicationService.GenerateError("El Contador {0} no existe.", strIDContador)
            End If
        End If
    End Function

    <Task()> Public Shared Sub AsignarEnvio347(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf AsignarEnvio347IPropAcc, data.Current, services)
    End Sub

    <Task()> Public Shared Sub AsignarEnvio349(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf AsignarEnvio349IPropAcc, data.Current, services)
    End Sub

    <Task()> Public Shared Sub AsignarEnvio347Doc(ByVal doc As DocumentCabLin, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf AsignarEnvio347IPropAcc, New DataRowPropertyAccessor(doc.HeaderRow), services)
    End Sub

    <Task()> Public Shared Sub AsignarEnvio349Doc(ByVal doc As DocumentCabLin, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf AsignarEnvio349IPropAcc, New DataRowPropertyAccessor(doc.HeaderRow), services)
    End Sub


    <Task()> Public Shared Sub AsignarEnvio347IPropAcc(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider)
        If Length(data("IDPais")) > 0 Then
            data("Enviar347") = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf Pais.NacionalNoCanariasCeutaMelilla, data("IDPais"), services)
        End If
        If Nz(data("Enviar347"), False) AndAlso Length(data("IDContador")) > 0 Then
            Dim EsContadorB As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf NegocioGeneral.ContadorB, data("IDContador"), services)
            data("Enviar347") = Not EsContadorB
            'Else
            '    data("Enviar347") = True
        End If
        If Nz(data("Enviar347"), False) AndAlso Nz(data("RetencionIRPF"), 0) = 0 Then
            data("Enviar347") = True
        Else
            data("Enviar347") = False
        End If
    End Sub

    <Task()> Public Shared Sub AsignarEnvio349IPropAcc(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider)
        If Length(data("IDPais")) > 0 Then
            data("Enviar349") = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf Pais.Intracomunitario, data("IDPais"), services)
        End If
        If Nz(data("Enviar349"), False) AndAlso Length(data("IDContador")) > 0 Then
            Dim EsContadorB As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf NegocioGeneral.ContadorB, data("IDContador"), services)
            data("Enviar349") = Not EsContadorB
            'Else
            '    data("Enviar349") = True
        End If

        If Nz(data("Enviar349"), False) AndAlso Nz(data("RetencionIRPF"), 0) = 0 Then
            data("Enviar349") = True
        Else
            data("Enviar349") = False
        End If
    End Sub

    <Task()> Public Shared Sub ValidarEnvio347Doc(ByVal doc As DocumentCabLin, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ValidarEnvio347IPropAcc, New DataRowPropertyAccessor(doc.HeaderRow), services)
    End Sub

    <Task()> Public Shared Sub ValidarEnvio349Doc(ByVal doc As DocumentCabLin, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ValidarEnvio349IPropAcc, New DataRowPropertyAccessor(doc.HeaderRow), services)
    End Sub

    <Task()> Public Shared Sub ValidarEnvio347(ByVal Row As DataRow, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ValidarEnvio347IPropAcc, New DataRowPropertyAccessor(row), services)
    End Sub

    <Task()> Public Shared Sub ValidarEnvio349(ByVal Row As DataRow, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ValidarEnvio349IPropAcc, New DataRowPropertyAccessor(row), services)
    End Sub


    <Task()> Public Shared Sub ValidarEnvio347IPropAcc(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider)

        Dim Es347 As Boolean = False

        If Length(data("IDPais")) > 0 Then
            Es347 = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf Pais.NacionalNoCanariasCeutaMelilla, data("IDPais"), services)
        End If
        If Es347 AndAlso Length(data("IDContador")) > 0 Then
            Dim EsContadorB As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf NegocioGeneral.ContadorB, data("IDContador"), services)
            Es347 = Not EsContadorB
            'Else
            '    Es347 = True
        End If
        If Es347 AndAlso Nz(data("RetencionIRPF"), 0) = 0 Then
            Es347 = True
        Else
            Es347 = False
        End If

        If Es347 Then
            'data("Enviar347") = True
        Else
            If Nz(data("Enviar347"), False) AndAlso Not Es347 Then
                If data.ContainsKey("SinMensaje") AndAlso Nz(data("SinMensaje"), False) Then
                    data("Enviar347") = False
                Else
                    ApplicationService.GenerateError("La Factura no debe enviarse a 347. No cumple los requisitos necesarios. Revise sus datos.")
                End If
            End If

        End If
    End Sub

    <Task()> Public Shared Sub ValidarEnvio349IPropAcc(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider)
        Dim Es349 As Boolean = False

        If Length(data("IDPais")) > 0 Then
            Es349 = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf Pais.Intracomunitario, data("IDPais"), services)
        End If
        If Es349 AndAlso Length(data("IDContador")) > 0 Then
            Dim EsContadorB As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf NegocioGeneral.ContadorB, data("IDContador"), services)
            Es349 = Not EsContadorB
            'Else
            '    Es349 = True
        End If
        If Es349 AndAlso Nz(data("RetencionIRPF"), 0) = 0 Then
            Es349 = True
        Else
            Es349 = False
        End If

        If Es349 Then
            'data("Enviar349") = True
        Else
            If Nz(data("Enviar349"), False) AndAlso Not Es349 Then
                If data.ContainsKey("SinMensaje") AndAlso Nz(data("SinMensaje"), False) Then
                    data("Enviar349") = False
                Else
                    ApplicationService.GenerateError("La Factura no debe enviarse a 349. No cumple los requisitos necesarios. Revise sus datos.")
                End If
            End If
        End If
        If Es349 Then
            Dim CIF As String = String.Empty
            If data.ContainsKey("CifProveedor") Then
                CIF = data("CifProveedor") & String.Empty
            ElseIf data.ContainsKey("CifCliente") Then
                CIF = data("CifCliente") & String.Empty
            End If

            Dim datDocID As New DataDocIdentificacion(CIF, data("IDPais"), enumTipoDocIdent.NIFOperadorIntra)
            ProcessServer.ExecuteTask(Of DataDocIdentificacion)(AddressOf General.Comunes.ValidarNIFIntracomunitario, datDocID, services)
        End If
    End Sub


    <Task()> Public Shared Sub AsignarMonedaPredeterminada(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
        If Length(data("IDMoneda")) = 0 Then data("IDMoneda") = AppParams.MonedaPredeterminada
    End Sub

    <Task()> Public Shared Sub AsignarMarcaIVAManual(ByVal data As DataRow, ByVal services As ServiceProvider)
        data("IVAManual") = False
    End Sub

    <Task()> Public Shared Sub AsignarMarcaVtosManuales(ByVal data As DataRow, ByVal services As ServiceProvider)
        data("VencimientosManuales") = False
    End Sub

#End Region

#End Region

#Region " Validaciones "

    <Task()> Public Shared Sub ValidarArticuloObligatorio(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDArticulo")) = 0 Then ApplicationService.GenerateError("El Artículo es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarClienteObligatorio(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IdCliente")) = 0 Then ApplicationService.GenerateError("El Cliente es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarClienteBloqueado(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
        Dim ClteInfo As ClienteInfo = Clientes.GetEntity(data("IDCliente"))
        If ClteInfo.Bloqueado Then ApplicationService.GenerateError("El Cliente está Bloqueado.")
    End Sub

    <Task()> Public Shared Sub ValidarProveedorObligatorio(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDProveedor")) = 0 Then ApplicationService.GenerateError("El Proveedor es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarEjercicioContableAlbaran(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim AppParamsConta As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
        If AppParamsConta.Contabilidad AndAlso Length(data("FechaAlbaran")) > 0 And Length(data("IDEjercicio")) = 0 Then
            ApplicationService.GenerateError("No hay un Ejercicio Predeterminado para la Fecha {0}.", Format(data("FechaAlbaran"), "dd/MM/yyyy"))
        End If
    End Sub

    <Task()> Public Shared Sub ValidarEjercicioContablePedido(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim AppParamsConta As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
        If AppParamsConta.Contabilidad AndAlso Length(data("FechaPedido")) > 0 AndAlso data("FechaPedido") <> cnMinDate AndAlso Length(data("IDEjercicio")) = 0 Then
            ApplicationService.GenerateError("No hay un Ejercicio Predeterminado para la Fecha {0}.", Format(data("FechaPedido"), "dd/MM/yyyy"))
        End If
    End Sub

    <Task()> Public Shared Sub ValidarEjercicioContableFactura(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim AppParamsConta As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
        If AppParamsConta.Contabilidad AndAlso Length(data("FechaFactura")) > 0 And Length(data("IDEjercicio")) = 0 Then
            ApplicationService.GenerateError("No hay un Ejercicio Predeterminado para la Fecha {0}.", Format(data("FechaFactura"), "dd/MM/yyyy"))
        End If
    End Sub

    <Task()> Public Shared Sub ValidarFechaAlbaranObligatoria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("FechaAlbaran")) = 0 Then ApplicationService.GenerateError("La Fecha Albarán es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarFechaPedidoObligatoria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("FechaPedido")) = 0 Then ApplicationService.GenerateError("La Fecha Pedido es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarFechaFacturaObligatoria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("FechaFactura")) = 0 Then ApplicationService.GenerateError("La Fecha Factura es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarMonedaObligatoria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDMoneda")) = 0 Then ApplicationService.GenerateError("La Moneda es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarAlbaranObligatorio(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDAlbaran")) = 0 Then ApplicationService.GenerateError("El Albarán es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarPedidoObligatorio(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDPedido")) = 0 Then ApplicationService.GenerateError("El Pedido es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarFacturaObligatoria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDFactura")) = 0 Then ApplicationService.GenerateError("El identificador de la Factura es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarTipoAlbaranObligatorio(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDTipoAlbaran")) = 0 Then ApplicationService.GenerateError("El Tipo de Albarán es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarFormaPagoObligatoria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDFormaPago")) = 0 Then ApplicationService.GenerateError("La Forma Pago es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarCondicionPagoObligatoria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDCondicionPago")) = 0 Then ApplicationService.GenerateError("La Condición de Pago es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarNumeroAlbaranCompra(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim f As New Filter
            f.Add(New StringFilterItem("NAlbaran", data("NAlbaran")))
            If Length(data("IDContador")) > 0 Then
                f.Add(New StringFilterItem("IDContador", data("IDContador")))
            Else
                f.Add(New IsNullFilterItem("IDContador", True))
            End If

            Dim AppParamsConta As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
            If AppParamsConta.Contabilidad Then f.Add(New StringFilterItem("IDEjercicio", data("IDEjercicio")))
            Dim dtACC As DataTable = New AlbaranCompraCabecera().Filter(f)
            If Not dtACC Is Nothing AndAlso dtACC.Rows.Count > 0 Then
                If AppParamsConta.Contabilidad Then
                    ApplicationService.GenerateError("El Albarán {0} ya existe para el Ejercicio {1}.", Quoted(data("NAlbaran")), Quoted(data("IDEjercicio")))
                Else
                    ApplicationService.GenerateError("El Albarán {0} ya existe.", Quoted(data("NAlbaran")))
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarDtoProntoPagoRecFinan(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Nz(data("DtoProntoPago"), 0) <> 0 AndAlso Nz(data("RecFinan"), 0) <> 0 Then
            ApplicationService.GenerateError("No se puede tener Descuento Pronto Pago y Recargo Financiero a la vez.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarClaveOperacion(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim ClvOp As ClaveOperacion?
        If Not IsDBNull(data("IDContador")) Then
            Dim Contadores As EntityInfoCache(Of ContadorInfo) = services.GetService(Of EntityInfoCache(Of ContadorInfo))()
            Dim ContInfo As ContadorInfo = Contadores.GetEntity(data("IDContador"))
            If Length(ContInfo.IDTipoComprobante) > 0 AndAlso Length(ContInfo.ClaveOperacion) > 0 Then
                '//Le asignamos la clave de operación del Tipo de Comprobante asociado al contador.
                ClvOp = ContInfo.ClaveOperacion
            End If
        End If

        If Length(data("ClaveOperacion")) > 0 Then
            If Length(ClvOp) > 0 AndAlso data("ClaveOperacion") <> ClvOp Then
                ApplicationService.GenerateError("La Clave de Operación del Contador no se corresponde con la de la Factura. Revise sus datos.")
            End If
            If data("ClaveOperacion") = ClaveOperacion.FacturaRectificativa Then
                If Length(data("IDFacturaRectificada")) = 0 Then ApplicationService.GenerateError("Debe indicar la Factura que Rectifica.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarCantidadLineaFactura(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Not IsNumeric(data("Cantidad")) Then
            ApplicationService.GenerateError("La Cantidad no es válida.")
        ElseIf Nz(data("Cantidad"), 0) = 0 Then
            ApplicationService.GenerateError("La Cantidad no puede ser cero.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarCantidadLineaAlbaran(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Not IsNumeric(data("QServida")) Then
            ApplicationService.GenerateError("La Cantidad no es válida.")
        ElseIf Nz(data("QServida"), 0) = 0 AndAlso (Not data.Table.Columns.Contains("Regalo") OrElse Not data("Regalo")) Then
            ApplicationService.GenerateError("La Cantidad no puede ser cero.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarCantidadLineaPedido(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Not IsNumeric(data("QPedida")) Then
            ApplicationService.GenerateError("La Cantidad no es válida.")
        ElseIf Nz(data("QPedida"), 0) = 0 Then
            ApplicationService.GenerateError("La Cantidad no puede ser cero.")
        End If
    End Sub
    <Task()> Public Shared Sub ValidarUnidadMedida(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDUdMedida")) = 0 Then
            ApplicationService.GenerateError("La unidad de medida es un dato obligatorio.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarAlmacenObligatorio(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDAlmacen")) = 0 Then ApplicationService.GenerateError("El Almacen es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarAlmacenBloqueado(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDAlmacen")) > 0 Then
            Dim Almacenes As EntityInfoCache(Of AlmacenInfo) = services.GetService(Of EntityInfoCache(Of AlmacenInfo))()
            Dim AlmInfo As AlmacenInfo = Almacenes.GetEntity(data("IDAlmacen"))
            If AlmInfo.Bloqueado Then ApplicationService.GenerateError("El Almacén {0} está bloqueado.", data("IDAlmacen"))
        End If
    End Sub

    <Task()> Public Shared Sub ValidarFechaEntregaObligatoria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("FechaEntrega")) = 0 Then ApplicationService.GenerateError("La Fecha Entrega es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarFechaInicioObligatoria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("FechaInicio")) = 0 Then ApplicationService.GenerateError("La Fecha Inicio es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarFechaFinObligatoria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("FechaFin")) = 0 Then ApplicationService.GenerateError("La Fecha Fin es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarIDCContableObligatoria(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim AppParamsConta As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
        If Not AppParamsConta.Contabilidad Then Exit Sub
        If Length(data("IDCContable")) = 0 Then ApplicationService.GenerateError("La Cuenta Contable es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarCContableObligatoria(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim AppParamsConta As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
        If Not AppParamsConta.Contabilidad Then Exit Sub
        If Length(data("CContable")) = 0 Then ApplicationService.GenerateError("La Cuenta Contable es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarTipoCobroObligatorio(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDTipoCobro")) = 0 Then ApplicationService.GenerateError("El Tipo Cobro es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarTipoPagoObligatorio(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDTipoPago")) = 0 Then ApplicationService.GenerateError("El Tipo Pago es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarUnidadObligatoria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("Unidad")) = 0 Then ApplicationService.GenerateError("La Unidad es obligatoria.")
    End Sub

    <Task()> Public Shared Sub ValidarPeriodoObligatorio(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("Periodo")) = 0 Then ApplicationService.GenerateError("El Periodo es obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarImporteObligatorio(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("Importe")) = 0 Then ApplicationService.GenerateError("El Importe es obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarFechaDesdeObligatoria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("FechaDesde")) = 0 Then ApplicationService.GenerateError("La Fecha Desde es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarFechaHastaObligatoria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("FechaHasta")) = 0 Then ApplicationService.GenerateError("La Fecha Hasta es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarFechaDesdeHasta(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data("FechaDesde") > data("FechaHasta") Then ApplicationService.GenerateError("La Fecha Hasta debe ser mayor que la Fecha Desde.")
    End Sub

    <Task()> Public Shared Sub ValidarCentroObligatorio(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDCentro")) = 0 Then ApplicationService.GenerateError("El Centro es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarCentroCosteObligatorio(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDCentroCoste")) = 0 Then ApplicationService.GenerateError("El Centro Coste es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarTasaObligatoria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDTasa")) = 0 Then ApplicationService.GenerateError("La Tasa es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarFechaVencimientoObligatoria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("FechaVencimiento")) = 0 Then ApplicationService.GenerateError("La Fecha Vencimiento es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarFechaDeclaracion(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Nz(data("FechaParaDeclaracion"), cnMinDate) <> cnMinDate AndAlso Length(data("IDEjercicio")) > 0 AndAlso _
             (data.RowState = DataRowState.Added OrElse _
              (data.RowState = DataRowState.Modified AndAlso _
               (Nz(data("FechaParaDeclaracion"), cnMinDate) <> Nz(data("FechaParaDeclaracion", DataRowVersion.Original), cnMinDate) OrElse _
                data("IDEjercicio") <> data("IDEjercicio", DataRowVersion.Original) & String.Empty))) Then
            Dim Ej As BusinessHelper = BusinessHelper.CreateBusinessObject("EjercicioContable")
            Dim dtEjercicio As DataTable = Ej.SelOnPrimaryKey(data("IDEjercicio"))
            If dtEjercicio.Rows.Count > 0 Then
                Dim UltimaFechaDeclaracion As Date = Nz(dtEjercicio.Rows(0)("UltimaFechaDeclaracion"), cnMinDate)
                If UltimaFechaDeclaracion >= data("FechaParaDeclaracion") Then
                    ApplicationService.GenerateError("La Fecha de Declaración debe ser posterior a la Fecha de Ultima Declaración del Ejercicio {0}.", Quoted(data("IDEjercicio")))
                End If
            End If
        End If
    End Sub
#End Region

#End Region

#Region " Integración con Financiero "


#Region " Ejercicio y Plan contable "


    '<Task()> Public Shared Sub ComprobarCContable(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
    '    Dim AppParamsConta As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
    '    If Not AppParamsConta.Contabilidad Then Exit Sub
    '    Dim FinancieroGeneral As IFinanciero = ProcessServer.ExecuteTask(Of Object, IFinanciero)(AddressOf Comunes.CreateFinancieroGeneral, Nothing, services)
    '    FinancieroGeneral.ComprobarCContable(data, services)
    'End Sub


    <Task()> Public Shared Sub FormatoCuentaContable(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim AppParamsConta As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
        If Not AppParamsConta.Contabilidad Then Exit Sub
        Dim FinancieroGeneral As IFinanciero = ProcessServer.ExecuteTask(Of Object, IFinanciero)(AddressOf Comunes.CreateFinancieroGeneral, Nothing, services)
        FinancieroGeneral.FormatoCuentaContable(data, services)
    End Sub

    <Task()> Public Shared Function CambioCContable(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim AppParamsConta As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
        If Not AppParamsConta.Contabilidad Then Exit Function
        Dim FinancieroGeneral As IFinanciero = ProcessServer.ExecuteTask(Of Object, IFinanciero)(AddressOf Comunes.CreateFinancieroGeneral, Nothing, services)
        FinancieroGeneral.CambioCContable(data, services)
    End Function

    <Task()> Public Shared Function EjercicioPredeterminado(ByVal Fecha As Date, ByVal services As ServiceProvider) As String
        Dim AppParamsConta As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
        If Not AppParamsConta.Contabilidad Then Exit Function

        Dim FinancieroGeneral As IFinanciero = ProcessServer.ExecuteTask(Of Object, IFinanciero)(AddressOf Comunes.CreateFinancieroGeneral, Nothing, services)
        Return FinancieroGeneral.EjercicioPredeterminado(Fecha, services)
    End Function

    <Task()> Public Shared Function EjercicioPredeterminadoB(ByVal Fecha As Date, ByVal services As ServiceProvider) As String
        Dim AppParamsConta As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
        If Not AppParamsConta.Contabilidad Then Exit Function
        Dim FinancieroGeneral As IFinanciero = ProcessServer.ExecuteTask(Of Object, IFinanciero)(AddressOf Comunes.CreateFinancieroGeneral, Nothing, services)
        Return FinancieroGeneral.EjercicioPredeterminadoB(Fecha, services)
    End Function

#End Region

#Region " Diario Contable "

    Public Shared Function CuentaSaldo(ByVal IDEjercicio As String, ByVal IDCContable As String) As DataTable
        '  Dim Diario As Object = BusinessHelper.CreateBusinessObject("DiarioContable")
        'Return Diario.CuentaSaldo(IDEjercicio, IDCContable)
        Dim FinancieroGeneral As IFinanciero = ProcessServer.ExecuteTask(Of Object, IFinanciero)(AddressOf Comunes.CreateFinancieroGeneral, Nothing, New ServiceProvider)
        Return FinancieroGeneral.CuentaSaldo(IDEjercicio, IDCContable)
    End Function

    Public Shared Function ExtractoCuenta(ByVal IDEjercicio As String, ByVal IDCContable As String) As DataTable
        'Dim Diario As Object = BusinessHelper.CreateBusinessObject("DiarioContable")
        ' Return Diario.ExtractoCuenta(IDEjercicio, IDCContable)
        Dim FinancieroGeneral As IFinanciero = ProcessServer.ExecuteTask(Of Object, IFinanciero)(AddressOf Comunes.CreateFinancieroGeneral, Nothing, New ServiceProvider)
        Return FinancieroGeneral.ExtractoCuenta(IDEjercicio, IDCContable)
    End Function

    Public Shared Function DeleteWhere(ByVal IDEjercicio As String, ByVal Filtro As Filter) As Boolean
        Dim FinancieroGeneral As IFinanciero = ProcessServer.ExecuteTask(Of Object, IFinanciero)(AddressOf Comunes.CreateFinancieroGeneral, Nothing, New ServiceProvider)
        Return FinancieroGeneral.DeleteWhere(IDEjercicio, Filtro)
    End Function

#End Region

#Region " Analitica "

    <Task()> Public Shared Sub ActualizarAnalitica(ByVal data As DataDocRow, ByVal services As ServiceProvider)
        Dim FinancieroGeneral As IFinanciero = ProcessServer.ExecuteTask(Of Object, IFinanciero)(AddressOf Comunes.CreateFinancieroGeneral, Nothing, services)
        FinancieroGeneral.ActualizarAnalitica(data, services)
    End Sub

    <Task()> Public Shared Function NuevaAnalitica(ByVal data As DataDocRow, ByVal services As ServiceProvider)
        Dim FinancieroGeneral As IFinanciero = ProcessServer.ExecuteTask(Of Object, IFinanciero)(AddressOf Comunes.CreateFinancieroGeneral, Nothing, services)
        FinancieroGeneral.NuevaAnaliticaLinea(data, services)
    End Function

    <Task()> Public Shared Function AnaliticaCommonBusinessRules(ByVal data As BusinessRuleData, ByVal services As ServiceProvider) As IPropertyAccessor
        Dim FinancieroGeneral As IFinanciero = ProcessServer.ExecuteTask(Of Object, IFinanciero)(AddressOf Comunes.CreateFinancieroGeneral, Nothing, services)
        Return FinancieroGeneral.AnaliticaCommonBusinessRules(data, services)
    End Function

    '<Task()> Public Shared Function AnaliticaCommonValidateRules(ByVal dttSource As DataTable, ByVal services As ServiceProvider) As DataTable
    '    Dim FinancieroGeneral As IFinanciero = ProcessServer.ExecuteTask(Of Object, IFinanciero)(AddressOf Comunes.CreateFinancieroGeneral, Nothing, services)
    '    Return FinancieroGeneral.AnaliticaCommonValidateRules(dttSource, services)
    'End Function

    <Task()> Public Shared Sub AnaliticaCommonValidateRules(ByVal dttSource As DataRow, ByVal services As ServiceProvider)
        Dim FinancieroGeneral As IFinanciero = ProcessServer.ExecuteTask(Of Object, IFinanciero)(AddressOf Comunes.CreateFinancieroGeneral, Nothing, services)
        FinancieroGeneral.AnaliticaCommonValidateRules(dttSource, services)
    End Sub

    <Task()> Public Shared Sub CalcularAnalitica(ByVal doc As DocumentCabLin, ByVal services As ServiceProvider)
        Dim AppParams As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
        If Not AppParams.Contabilidad OrElse Not AppParams.Analitica.AplicarAnalitica Then Exit Sub
        For Each linea As DataRow In doc.dtLineas.Rows
            If linea.RowState = DataRowState.Added Then
                Dim ctx As New DataDocRow(doc, linea)
                ProcessServer.ExecuteTask(Of DataDocRow)(AddressOf NuevaAnalitica, ctx, services)
            End If
            If linea.RowState = DataRowState.Modified Then
                Dim ctx As New DataDocRow(doc, linea)
                ProcessServer.ExecuteTask(Of DataDocRow)(AddressOf ActualizarAnalitica, ctx, services)
            End If
        Next
    End Sub

#Region " Copiar Analitica "

    <Serializable()> _
    Public Class DataCopiarAnalitica
        Public AnaliticaOrigen As DataTable
        Public Doc As DocumentCabLin

        Public Sub New(ByVal AnaliticaOrigen As DataTable, ByVal Doc As DocumentCabLin)
            Me.AnaliticaOrigen = AnaliticaOrigen
            Me.Doc = Doc
        End Sub
    End Class

    <Task()> Public Shared Sub CopiarAnalitica(ByVal data As DataCopiarAnalitica, ByVal services As ServiceProvider)
        If Not data.AnaliticaOrigen Is Nothing AndAlso data.AnaliticaOrigen.Rows.Count > 0 Then
            Dim pkOrigen, pkDestino As String
            Select Case data.AnaliticaOrigen.TableName
                Case GetType(PedidoVentaAnalitica).Name, GetType(PedidoCompraAnalitica).Name
                    pkOrigen = "IDLineaPedido"
                Case GetType(AlbaranVentaAnalitica).Name, GetType(AlbaranCompraAnalitica).Name
                    pkOrigen = "IDLineaAlbaran"
                Case GetType(FacturaVentaAnalitica).Name, GetType(FacturaCompraAnalitica).Name
                    pkOrigen = "IDLineaFactura"
                Case "ObraTrabajoAnalitica"
                    pkOrigen = "IDTrabajo"
                Case "ObraAnalitica"
                    pkOrigen = "IDObra"
            End Select

            Select Case data.Doc.dtAnalitica.TableName
                Case GetType(PedidoVentaAnalitica).Name, GetType(PedidoCompraAnalitica).Name
                    pkDestino = "IDLineaPedido"
                Case GetType(AlbaranVentaAnalitica).Name, GetType(AlbaranCompraAnalitica).Name
                    pkDestino = "IDLineaAlbaran"
                Case GetType(FacturaVentaAnalitica).Name, GetType(FacturaCompraAnalitica).Name
                    pkDestino = "IDLineaFactura"
                Case "ObraTrabajoAnalitica"
                    pkDestino = "IDTrabajo"
                Case "ObraAnalitica"
                    pkDestino = "IDObra"
            End Select

            If Length(pkOrigen) > 0 AndAlso Length(pkDestino) > 0 Then
                Dim f As New Filter
                f.Add(New IsNullFilterItem(pkOrigen, False))
                Dim WhereNotNullOrigenAnalitica As String = f.Compose(New AdoFilterComposer)
                For Each linea As DataRow In data.Doc.dtLineas.Select(WhereNotNullOrigenAnalitica)
                    Dim fIDOrigen As New Filter
                    fIDOrigen.Add(New NumberFilterItem(pkOrigen, linea(pkOrigen)))
                    Dim UltimaRow As DataRow
                    Dim Acum As Double = 0 : Dim AcumA As Double = 0 : Dim AcumB As Double = 0
                    Dim HayAnalitica As Boolean = False
                    Dim WhereOrigenAnalitica As String = fIDOrigen.Compose(New AdoFilterComposer)
                    For Each lineaAnalitica As DataRow In data.AnaliticaOrigen.Select(WhereOrigenAnalitica, pkOrigen)
                        Dim drNewLine As DataRow = data.Doc.dtAnalitica.NewRow
                        drNewLine(pkDestino) = linea(pkDestino)
                        drNewLine("IDCentroCoste") = lineaAnalitica("IDCentroCoste")
                        drNewLine("Porcentaje") = lineaAnalitica("Porcentaje")
                        drNewLine("Importe") = (linea("Importe") * lineaAnalitica("Porcentaje")) / 100
                        Dim ValAyB As New ValoresAyB(New DataRowPropertyAccessor(drNewLine), data.Doc.IDMoneda, data.Doc.CambioA, data.Doc.CambioB)
                        ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf MantenimientoValoresAyB, ValAyB, services)
                        Acum += drNewLine("Importe")
                        AcumA += drNewLine("ImporteA")
                        AcumB += drNewLine("ImporteB")
                        UltimaRow = drNewLine
                        data.Doc.dtAnalitica.Rows.Add(drNewLine)
                        HayAnalitica = True
                    Next
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
            End If
        End If
    End Sub
    'David Velasco Herrero 28/7/22 FACTURA PISO
    <Task()> Public Shared Sub RellenaAnaliticaFactura(ByVal data As DataCopiarAnalitica, ByVal services As ServiceProvider)
        '


        '
        Dim pkDestino As String
        pkDestino = "IDLineaFactura"
        For Each linea As DataRow In data.Doc.dtLineas.Select()
            Dim UltimaRow As DataRow
            Dim Acum As Double = 0 : Dim AcumA As Double = 0 : Dim AcumB As Double = 0
            Dim HayAnalitica As Boolean = False

            Dim lineaAnalitica As DataRow
            Dim drNewLine As DataRow = data.Doc.dtAnalitica.NewRow
            drNewLine(pkDestino) = linea(pkDestino)
            drNewLine("IDCentroCoste") = "T636"
            drNewLine("Porcentaje") = "100"
            drNewLine("Importe") = linea("Importe")
            Dim ValAyB As New ValoresAyB(New DataRowPropertyAccessor(drNewLine), data.Doc.IDMoneda, data.Doc.CambioA, data.Doc.CambioB)
            ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf MantenimientoValoresAyB, ValAyB, services)
            Acum += drNewLine("Importe")
            AcumA += drNewLine("ImporteA")
            AcumB += drNewLine("ImporteB")
            UltimaRow = drNewLine
            data.Doc.dtAnalitica.Rows.Add(drNewLine)
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
    End Sub

#End Region

    <Task()> Public Shared Sub ComprobarAnaliticaOrigen(ByVal data As DataDocRow, ByVal services As ServiceProvider)
        'Dim AppParams As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
        'If AppParams.Contabilidad AndAlso AppParams.Analitica.AplicarAnalitica AndAlso AppParams.Analitica.AnaliticaOrigen Then
        '    Dim Doc As DocumentCabLin = data.Doc

        '    '//Sólo seguir, en el caso, en que la C.Contable sea Analítica.
        '    If Length(data.Row("CContable")) > 0 AndAlso Length(Doc.HeaderRow("IDEjercicio")) > 0 Then
        '        '//Comprobar si la cuenta es analítica
        '        Dim objNegPlanContable As New PlanContable
        '        Dim blnAnalitica As Boolean = objNegPlanContable.CuentaAnalitica(Doc.HeaderRow("IDEjercicio"), data.Row("CContable"))
        '        '//Si NO es analítica, y el parámetro C_ANALIT_T=CCAnalitica(=1), no podemos introducirla.
        '        If Not blnAnalitica Then
        '            If AppParams.Analitica.AnaliticaTipo = enumGestionAnalitica.CCAnalitica Then
        '                Exit Sub
        '            End If
        '        Else
        '            '//Si es analítica, debemos tener un C.Coste en la Cabecera del Pedido/Albaran/Factura.
        '            If Length(data.Doc.HeaderRow("IDCentroCoste")) = 0 Then
        '                ApplicationService.GenerateError("Debe de Seleccionar un Centro de Coste predeterminado para el Desglose Analítico.")
        '            End If
        '        End If
        '    Else
        '        If Length(data.Row("CContable")) > 0 AndAlso Length(Doc.HeaderRow("IDEjercicio")) = 0 Then
        '            ApplicationService.GenerateError("Debe de indicar el Ejercicio Contable.")
        '        End If
        '    End If

        '    Dim newrow As DataRow = Doc.dtAnalitica.NewRow
        '    newrow(Doc.PrimaryKeyLin(0)) = data.Row(Doc.PrimaryKeyLin(0))
        '    newrow("IDCentroCoste") = Doc.HeaderRow("IDCentroCoste")
        '    newrow("Importe") = data.Row("Importe")
        '    newrow("Porcentaje") = 100

        '    Dim ValAyB As New ValoresAyB(newrow, Doc.Moneda.ID, Doc.Moneda.CambioA, Doc.Moneda.CambioB)
        '    ProcessServer.ExecuteTask(Of ValoresAyB)(AddressOf MantenimientoValoresAyB, ValAyB, services)
        '    Doc.dtAnalitica.Rows.Add(newrow.ItemArray)
        'End If
    End Sub

#End Region

#End Region
#Region " Impuestos "

    <Task()> Public Shared Sub CambioImpuesto(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDImpuesto")) > 0 Then
            ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ValidarImpuestoExistente, data.Current, services)
            ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CambioPorcentaje, data, services)
        End If
    End Sub

    <Task()> Public Shared Sub ValidarImpuestoExistente(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider)
        If Length(data("IDImpuesto")) > 0 Then
            Dim dtImpuesto As DataTable = New Impuesto().SelOnPrimaryKey(data("IDImpuesto"))
            If dtImpuesto Is Nothing OrElse dtImpuesto.Rows.Count = 0 Then
                ApplicationService.GenerateError("El Impuesto indicado no existe.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioImporte(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim IncrementoImporte As Double
        If data.ColumnName = "Importe" Then
            If Not IsNumeric(data.Value) Then data.Value = 0
            IncrementoImporte = data.Value - data.Current("Importe")
        End If
        data.Current(data.ColumnName) = data.Value
        If Not IsNumeric(data.Current("valor")) Then data.Current("valor") = 0
        If Not IsNumeric(data.Current("Importe")) Then data.Current("Importe") = 0

        If data.Context("ImporteLinea") < 0 Then
            If data.Context("SumaImporte") + IncrementoImporte > data.Context("ImporteLinea") Then
                'If Nz(data.Context("ImporteLinea"), 0) <> 0 Then
                '    data.Current("Porcentaje") = 100 * (data.Current("Importe") / data.Context("ImporteLinea"))
                'End If
            Else
                ApplicationService.GenerateError("Los importes asignados a los Impuestos superan el importe total de la línea.")
            End If
        Else
            If data.Context("SumaImporte") + IncrementoImporte > data.Context("ImporteLinea") Then
                ApplicationService.GenerateError("Los importes asignados a los Impuestos superan el importe total de la línea.")
                'Else
                '    If Nz(data.Context("ImporteLinea"), 0) <> 0 Then
                '        data.Current("Porcentaje") = 100 * (data.Current("Importe") / data.Context("ImporteLinea"))
                '    End If
            End If
        End If

        If data.Context.ContainsKey("IDMoneda") AndAlso data.Context.ContainsKey("CambioA") AndAlso data.Context.ContainsKey("CambioB") Then
            Dim ValAyB As New ValoresAyB(data.Current, data.Context("IDMoneda"), data.Context("CambioA"), data.Context("CambioB"))
            ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)
        End If
    End Sub

    <Task()> Public Shared Sub CambioPorcentaje(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CambioValor, data, services)
    End Sub
    <Task()> Public Shared Sub CambioValor(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If data.ColumnName = "Valor" Then If Not IsNumeric(data.Value) Then data.Value = 0

        data.Current(data.ColumnName) = data.Value
        If Not IsNumeric(data.Current("Valor")) Then data.Current("Valor") = 0
        If Not IsNumeric(data.Current("Importe")) Then data.Current("Importe") = 0

        If data.Current("Porcentaje") Then
            ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ValidarPorcentajeValido, data.Current, services)
        End If
        Dim importe As Double
        If data.Current("Porcentaje") Then
            importe = data.Context("ImporteLinea") * (data.Current("Valor") / 100)
        Else
            importe = data.Current("Valor") * data.Context("QInterna")
        End If
        Dim IncrementoImporte As Double = importe - data.Current("Importe")
        If data.Context("ImporteLinea") < 0 Then
            If data.Context("SumaImporte") + IncrementoImporte > data.Context("ImporteLinea") Then
                data.Current("Importe") = importe
            Else
                ApplicationService.GenerateError("El importe total asignado a los Impuestos supera el importe total de la línea.")
            End If
        Else
            If data.Context("SumaImporte") + IncrementoImporte > data.Context("ImporteLinea") Then
                ApplicationService.GenerateError("El importe total asignado a los Impuestos supera el importe total de la línea.")
            Else
                data.Current("Importe") = importe
            End If
        End If

        If data.Context.ContainsKey("IDMoneda") AndAlso data.Context.ContainsKey("CambioA") AndAlso data.Context.ContainsKey("CambioB") Then
            Dim ValAyB As New ValoresAyB(data.Current, data.Context("IDMoneda"), data.Context("CambioA"), data.Context("CambioB"))
            ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)
        End If
    End Sub
    <Task()> Public Shared Sub ValidarPorcentajeValido(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider)
        If data("Porcentaje") AndAlso (data("Valor") < 0 OrElse data("Valor") > 100) Then ApplicationService.GenerateError("El Porcentaje debe tener un valor entre 0 y 100.")
    End Sub


#Region " Creación/Actualización de Impuestos de Artículos en Facturas"

    '//AFTERTASK: Después de Negocio.ProcesoFacturacionCompra.CalcularImporteLineasFactura
    '//AFTERTASK: Después de Negocio.ProcesoFacturacionVenta.CalcularImporteLineasFactura
    <Task()> Public Shared Sub CalcularImpuestos(ByVal doc As DocumentCabLin, ByVal services As ServiceProvider)
        For Each linea As DataRow In doc.dtLineas.Rows
            If linea.RowState = DataRowState.Added Then
                Dim ctx As New DataDocRow(doc, linea)
                ProcessServer.ExecuteTask(Of DataDocRow)(AddressOf NuevoImpuesto, ctx, services)
            End If
            If linea.RowState = DataRowState.Modified Then
                Dim ctx As New DataDocRow(doc, linea)
                ProcessServer.ExecuteTask(Of DataDocRow)(AddressOf ActualizarImpuesto, ctx, services)
            End If
        Next
    End Sub

    <Task()> Public Shared Sub NuevoImpuesto(ByVal data As DataDocRow, ByVal services As ServiceProvider)
        If Not IsNothing(data.Row) Then
            If Length(data.Row("IDArticulo")) = 0 Then Exit Sub
            Dim PKLinea As String
            Dim newData As DataTable
            Select Case CType(data.Doc, DocumentCabLin).EntidadLineas
                Case GetType(FacturaVentaLinea).Name, GetType(FacturaCompraLinea).Name
                    PKLinea = "IDLineaFactura"
            End Select

            If TypeOf data.Doc Is DocumentoFacturaVenta Then
                newData = CType(data.Doc, DocumentoFacturaVenta).dtImpuestos
            ElseIf TypeOf data.Doc Is DocumentoFacturaCompra Then
                newData = CType(data.Doc, DocumentoFacturaCompra).dtImpuestos
            End If
            Dim Impuestos As DataTable = New BE.DataEngine().Filter("vNegArticuloImpuesto", New StringFilterItem("IDArticulo", data.Row("IDArticulo")))
            If Not Impuestos Is Nothing AndAlso Impuestos.Rows.Count > 0 Then
                Dim dv As DataView = Impuestos.DefaultView
                dv.RowFilter = "Valor<>0"
                dv.Sort = "IDImpuesto"
                Dim IDImpuesto As String
                For Each r As DataRowView In Impuestos.DefaultView
                    If IDImpuesto <> r("IDImpuesto") Then
                        If r("AplicarSobre") <> CInt(AplicarSobre.Compras) AndAlso TypeOf data.Doc Is DocumentoFacturaVenta OrElse _
                          r("AplicarSobre") <> CInt(AplicarSobre.Ventas) AndAlso TypeOf data.Doc Is DocumentoFacturaCompra Then
                            Dim newrow As DataRow = newData.NewRow()
                            newrow("IDFactura") = data.Doc.HeaderRow("IDFactura")

                            IDImpuesto = r("IDImpuesto")
                            newrow("IDLineaImpuesto") = AdminData.GetAutoNumeric()
                            newrow(PKLinea) = data.Row(PKLinea)
                            newrow("IDImpuesto") = r("IDImpuesto")
                            newrow("Porcentaje") = r("Porcentaje")
                            Dim Valor As Double = Nz(r("Valor"), 0)
                            newrow("Valor") = Valor
                            If r("Porcentaje") Then
                                newrow("Importe") = (data.Row("Importe") * Valor) / 100
                            Else
                                newrow("Importe") = Valor * data.Row("QInterna")
                            End If

                            Dim ValAyB As New ValoresAyB(New DataRowPropertyAccessor(newrow), CType(data.Doc, DocumentCabLin).IDMoneda, CType(data.Doc, DocumentCabLin).CambioA, CType(data.Doc, DocumentCabLin).CambioB)
                            ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)

                            newData.Rows.Add(newrow.ItemArray)
                        End If

                    End If
                Next
            End If

        End If
    End Sub

    <Task()> Public Shared Sub ActualizarImpuesto(ByVal data As DataDocRow, ByVal services As ServiceProvider)
        Dim blnCambioArticulo, blnCambioImporte, blnCambioCantidad As Boolean

        Dim Impuestos As DataTable ' = CType(data.Doc, DocumentoComercial).dtVentaRepresentante
        If TypeOf data.Doc Is DocumentoFacturaVenta Then
            Impuestos = CType(data.Doc, DocumentoFacturaVenta).dtImpuestos
        ElseIf TypeOf data.Doc Is DocumentoFacturaCompra Then
            Impuestos = CType(data.Doc, DocumentoFacturaCompra).dtImpuestos
        End If
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
            If blnCambioArticulo Or blnCambioImporte Or blnCambioCantidad Then
                Dim f As New Filter
                Select Case CType(data.Doc, DocumentCabLin).EntidadLineas
                    Case "FacturaVentaLinea"
                        f.Add(New NumberFilterItem("IDLineaFactura", data.Row("IDLineaFactura")))
                    Case "FacturaCompraLinea"
                        f.Add(New NumberFilterItem("IDLineaFactura", data.Row("IDLineaFactura")))
                        'Case "AlbaranVentaLinea"
                        '    f.Add(New NumberFilterItem("IDLineaAlbaran", data.Row("IDLineaAlbaran")))
                        'Case "PedidoVentaLinea"
                        '    f.Add(New NumberFilterItem("IDLineaPedido", data.Row("IDLineaPedido")))
                End Select

                Dim WhereLinea As String = f.Compose(New AdoFilterComposer)
                Dim ImpuestosLinea() As DataRow = Impuestos.Select(WhereLinea)
                If Not ImpuestosLinea Is Nothing AndAlso ImpuestosLinea.Length > 0 Then
                    If blnCambioImporte And (Not blnCambioArticulo And Not blnCambioCantidad) Then
                        If blnCambioImporte And Not blnCambioArticulo Then

                            If Not data.Doc.HeaderRow Is Nothing Then

                                For Each lineaImpuesto As DataRow In ImpuestosLinea
                                    Dim Valor As Double = lineaImpuesto("Valor")
                                    lineaImpuesto("Valor") = Valor
                                    If lineaImpuesto("Porcentaje") Then
                                        lineaImpuesto("Importe") = (dr("Importe") * Valor) / 100
                                    Else
                                        lineaImpuesto("Importe") = Valor * dr("QInterna")
                                    End If

                                    Dim ValAyB As New ValoresAyB(New DataRowPropertyAccessor(lineaImpuesto), CType(data.Doc, DocumentCabLin).IDMoneda, CType(data.Doc, DocumentCabLin).CambioA, CType(data.Doc, DocumentCabLin).CambioB)
                                    ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)
                                Next
                            End If '
                        End If
                    Else
                        'Si se modifica el artículo o la cantidad, se elimina el desglose anterior y se vuelve a calcular
                        For Each lineaImpuesto As DataRow In ImpuestosLinea
                            lineaImpuesto.Delete()
                        Next
                        ProcessServer.ExecuteTask(Of DataDocRow)(AddressOf NuevoImpuesto, data, services)
                    End If
                Else
                    ProcessServer.ExecuteTask(Of DataDocRow)(AddressOf NuevoImpuesto, data, services)
                End If

            End If
        End If
    End Sub

#End Region
#End Region
End Class
