Public Class ModeloFactoring

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbModeloFactoring"

#End Region

#Region " FACTORING - Generación de ficheros "

    Private Const ERROR_INFO_FACTORING_BANCO_PROPIO As String = "Error en la información de Factoring del Banco Propio."

    Public Const FACTORING_BBVA As String = "A"
    Public Const FACTORING_BSCH As String = "B"
    Public Const FACTORING_BANCO_SABADELL As String = "C"
    Public Const FACTORING_BANCO_POPULAR As String = "D"
    Public Const FACTORING_CAJA_MADRID As String = "E"
    Public Const FACTORING_BANESTO As String = "F"

    <Serializable()> _
    Public Class ExportacionFactoringInfo
        Public DatosExportacion As DataTable
        Public ExportacionCorrecta As Boolean
    End Class

    Public Function ExportacionFicheroFactoring(ByVal strIDProcess As String, ByVal strNFactoring As String, ByVal objBPInfo As BancoPropio.BancoPropioFactoringInfo) As DataTable
        If Not IsNothing(objBPInfo) Then
            '//Generamos la información para el fichero correspondiente. Los ficheros del BBVA y del Banco Popular
            '// son de tipo excel, con lo que su expotación estará implementada en presentación.
            Dim objInfoExportacion As New ExportacionFactoringInfo
            Select Case objBPInfo.TipoFactoring & String.Empty
                Case FACTORING_BSCH
                    objInfoExportacion = ExportacionBSCH(strIDProcess, strNFactoring, objBPInfo)
                Case FACTORING_BANCO_SABADELL
                    objInfoExportacion = ExportacionBancoSabadellAtlantico(strIDProcess, strNFactoring, objBPInfo)
                Case FACTORING_CAJA_MADRID
                    objInfoExportacion = ExportacionCajaMadrid(strIDProcess, strNFactoring, objBPInfo)
                Case FACTORING_BANESTO
                    objInfoExportacion = ExportacionBanesto(strIDProcess, strNFactoring, objBPInfo)
            End Select

            If Not objInfoExportacion.ExportacionCorrecta Then
                ExportacionFicheroFactoring = Nothing
                'ApplicationService.GenerateError("Ha ocurrido un error en el proceso de exportación.")
            Else
                Return objInfoExportacion.DatosExportacion
            End If
        End If
        Return Nothing
    End Function

#Region " BSCH "

    Private Function ExportacionBSCH(ByVal strIDProcess As String, ByVal strNFactoring As String, ByVal objBPInfo As BancoPropio.BancoPropioFactoringInfo) As ExportacionFactoringInfo
        Const LONGITUD_LINEA As Integer = 237

        '//Prefijos
        Const PREFIJO_LINEA_CABECERA_SOPORTE As String = "03"
        Const PREFIJO_LINEA_TOTAL_SOPORTE As String = "99"
        Const PREFIJO_LINEA_CABECERA_REMESA As String = "06"
        Const PREFIJO_LINEA_TOTAL_REMESA As String = "66"
        Const PREFIJO_LINEA_REGISTRO_DETALLE As String = "09"
        Const PREFIJO_FACTURA As String = "F"

        '//Formatos
        Const FORMATO_FECHA As String = "yyMMdd"
        Dim FORMATO_CONTADOR_FACTORING As String = Strings.StrDup(5, "0")
        Dim FORMATO_CONTADOR_FACTURAS As String = Strings.StrDup(4, "0")
        Dim FORMATO_IMPORTE_TOTAL As String = Strings.StrDup(15, "0")

        Dim FORMATO_BANCO As String = Strings.StrDup(4, "0")
        Dim FORMATO_SUCURSAL As String = Strings.StrDup(4, "0")
        Dim FORMATO_DIGITO_CONTROL As String = Strings.StrDup(2, "0")
        Dim FORMATO_NCUENTA As String = Strings.StrDup(10, "0")

        '//Longitudes campos 
        Const LONGITUD_CLIENTE_FACTORING As Integer = 10
        Const LONGITUD_NOMBRE_CEDENTE As Integer = 30
        Const LONGITUD_CODIGO_DEUDOR As Integer = 10
        Const LONGITUD_RAZON_SOCIAL As Integer = 30
        Const LONGITUD_CIFNIF_DEUDOR As Integer = 14
        Const LONGITUD_DIRECCION As Integer = 30
        Const LONGITUD_POBLACION As Integer = 30
        Const LONGITUD_CODIGO_POSTAL As Integer = 5
        Const LONGITUD_PAIS As Integer = 2
        Const LONGITUD_MONEDA As Integer = 3
        Const LONGITUD_NFACTURA As Integer = 10
        Const LONGITUD_NOTA_ABONO As Integer = 10
        Const LONGITUD_COND_VENTA_FACTORING As Integer = 1
        Const LONGITUD_IDCLIENTE As Integer = 15


        Dim objExportacionInfo As New ExportacionFactoringInfo
        objExportacionInfo.ExportacionCorrecta = False

        Dim objFilter As New Filter
        objFilter.Clear()
        objFilter.Add(New StringFilterItem("IDProcess", FilterOperator.Equal, strIDProcess))
        objFilter.Add(New NumberFilterItem("ImpVencimiento", FilterOperator.GreaterThan, 0))
        Dim dtExportacion As DataTable = New BE.DataEngine().Filter("FicheroFactoringBSCH", objFilter)
        If IsNothing(dtExportacion) OrElse dtExportacion.Rows.Count = 0 Then
            objExportacionInfo.DatosExportacion = Nothing
            '//47409: No hay datos a exportar.
            ApplicationService.GenerateError("No hay datos a exportar.")
        End If

        If Not IsNothing(objBPInfo) Then
            Dim strLineaFichero As String

            '//////////////////////  DEFINICION DATOS A RETORNAR  /////////////////////////
            '//Creamos la estructura del DataTable a retornar, con las líneas del fichero.
            Dim dColumn As New DataColumn
            dColumn.ColumnName = "Linea"
            dColumn.DataType = GetType(String)
            dColumn.MaxLength = LONGITUD_LINEA             'Longitud máxima de la línea

            objExportacionInfo.DatosExportacion = New DataTable
            objExportacionInfo.DatosExportacion.Columns.Clear()
            objExportacionInfo.DatosExportacion.Columns.Add(dColumn)

            '/////////////////////// REGISTRO CABECERA DE SOPORTE  ////////////////////////
            strLineaFichero = PREFIJO_LINEA_CABECERA_SOPORTE

            Dim strIDClienteFactoring As String = Left(objBPInfo.IDClienteFactoring & String.Empty, LONGITUD_CLIENTE_FACTORING)
            strIDClienteFactoring = strIDClienteFactoring & Space(LONGITUD_CLIENTE_FACTORING - Length(strIDClienteFactoring))
            strLineaFichero = strLineaFichero & strIDClienteFactoring

            Dim strNombreCedente As String = String.Empty
            Dim objNegDatosEmpresa As New DatosEmpresa
            Dim dtDatosEmpresa As DataTable = objNegDatosEmpresa.Filter("DescEmpresa")
            If Not IsNothing(dtDatosEmpresa) AndAlso dtDatosEmpresa.Rows.Count > 0 Then
                strNombreCedente = Left(dtDatosEmpresa.Rows(0)("DescEmpresa") & String.Empty, LONGITUD_NOMBRE_CEDENTE)
                strNombreCedente = strNombreCedente & Space(LONGITUD_NOMBRE_CEDENTE - Length(strNombreCedente))
            End If
            objNegDatosEmpresa = Nothing
            strLineaFichero = strLineaFichero & strNombreCedente & String.Empty
            strLineaFichero = strLineaFichero & Format(Today, FORMATO_FECHA)
            strLineaFichero = strLineaFichero & Space(LONGITUD_LINEA - Length(strLineaFichero))  'Completamos con blancos la cadena
            objExportacionInfo.DatosExportacion.Rows.Clear()
            Dim drNewLinea As DataRow = objExportacionInfo.DatosExportacion.NewRow
            drNewLinea("Linea") = strLineaFichero
            objExportacionInfo.DatosExportacion.Rows.Add(drNewLinea)

            '/////////////////////// REGISTRO CABECERA DE REMESA  ////////////////////////
            strLineaFichero = PREFIJO_LINEA_CABECERA_REMESA
            'strLineaFichero = strLineaFichero & Format(strNFactoring, FORMATO_CONTADOR_FACTORING)
            Dim strNFactoringAux As String = Left(strNFactoring & String.Empty, Length(FORMATO_CONTADOR_FACTORING))
            strNFactoringAux = strNFactoringAux & Space(Length(FORMATO_CONTADOR_FACTORING) - Length(strNFactoringAux))
            strLineaFichero = strLineaFichero & strNFactoringAux

            strLineaFichero = strLineaFichero & Format(Today, FORMATO_FECHA)
            strLineaFichero = strLineaFichero & Space(LONGITUD_LINEA - Length(strLineaFichero))  'Completamos con blancos la cadena
            drNewLinea = objExportacionInfo.DatosExportacion.NewRow
            drNewLinea("Linea") = strLineaFichero
            objExportacionInfo.DatosExportacion.Rows.Add(drNewLinea)

            '/////////////////////////// REGISTROS DE DETALLE  ///////////////////////////
            Dim intTotalReg As Integer : Dim dblImporteTotal As Double
            If Not IsNothing(dtExportacion) AndAlso dtExportacion.Rows.Count > 0 Then
                For Each drRowExportacion As DataRow In dtExportacion.Rows
                    strLineaFichero = PREFIJO_LINEA_REGISTRO_DETALLE

                    '////////  Datos deudor  //////////
                    'strLineaFichero = strLineaFichero & Left(drRowExportacion("CodigoDeudor") & String.Empty, LONGITUD_CODIGO_DEUDOR) 'Código Deudor
                    strLineaFichero = strLineaFichero & Space(LONGITUD_CODIGO_DEUDOR)                                           'Código Deudor

                    Dim strRazonSocial As String = Left(drRowExportacion("Nombre") & String.Empty, LONGITUD_RAZON_SOCIAL)   'Razón Social Deudor
                    strRazonSocial = strRazonSocial & Space(LONGITUD_RAZON_SOCIAL - Length(strRazonSocial))
                    strLineaFichero = strLineaFichero & strRazonSocial

                    Dim strCIFNIFDeudor As String = drRowExportacion("CifDeudor") & String.Empty
                    strCIFNIFDeudor = Replace(strCIFNIFDeudor, "/", String.Empty)
                    strCIFNIFDeudor = Replace(strCIFNIFDeudor, ".", String.Empty)
                    strCIFNIFDeudor = Replace(strCIFNIFDeudor, "-", String.Empty)
                    strCIFNIFDeudor = Left(strCIFNIFDeudor, LONGITUD_CIFNIF_DEUDOR)
                    strCIFNIFDeudor = strCIFNIFDeudor & Space(LONGITUD_CIFNIF_DEUDOR - Length(strCIFNIFDeudor))
                    strLineaFichero = strLineaFichero & strCIFNIFDeudor  'CIF/NIF Deudor

                    Dim strDireccion As String = Left(drRowExportacion("Direccion") & String.Empty, LONGITUD_DIRECCION)
                    strDireccion = strDireccion & Space(LONGITUD_DIRECCION - Length(strDireccion))
                    strLineaFichero = strLineaFichero & strDireccion

                    Dim strPoblacion As String = Left(drRowExportacion("Poblacion") & String.Empty, LONGITUD_POBLACION)
                    strPoblacion = strPoblacion & Space(LONGITUD_POBLACION - Length(strPoblacion))
                    strLineaFichero = strLineaFichero & strPoblacion

                    Dim strCodPostal As String = Left(drRowExportacion("CodigoPostal") & String.Empty, LONGITUD_CODIGO_POSTAL)
                    strCodPostal = strCodPostal & Space(LONGITUD_CODIGO_POSTAL - Length(strCodPostal))
                    strLineaFichero = strLineaFichero & strCodPostal

                    '///////  Datos Factura  ///////////
                    strLineaFichero = strLineaFichero & PREFIJO_FACTURA

                    Dim strNFactura As String = Left(drRowExportacion("NFactura") & String.Empty, LONGITUD_NFACTURA)
                    strNFactura = strNFactura & Space(LONGITUD_NFACTURA - Length(strNFactura))
                    strLineaFichero = strLineaFichero & strNFactura

                    strLineaFichero = strLineaFichero & Space(LONGITUD_NOTA_ABONO)

                    strLineaFichero = strLineaFichero & Format(Today, FORMATO_FECHA)

                    Dim strFechaVto As String = String.Empty
                    If Length(drRowExportacion("FechaVencimiento") & String.Empty) > 0 Then
                        strFechaVto = Format(CDate(drRowExportacion("FechaVencimiento") & String.Empty), FORMATO_FECHA)
                    End If
                    strFechaVto = strFechaVto & Space(Length(FORMATO_FECHA) - Length(strFechaVto))
                    strLineaFichero = strLineaFichero & strFechaVto

                    dblImporteTotal = dblImporteTotal + (Nz(drRowExportacion("ImporteFactura"), 0) * 100)
                    strLineaFichero = strLineaFichero & Format((Nz(drRowExportacion("ImporteFactura"), 0) * 100), FORMATO_IMPORTE_TOTAL)

                    strLineaFichero = strLineaFichero & Space(6)

                    '//////////  Datos Bancarios  ///////////
                    If Length(drRowExportacion("IDBancoFactura") & String.Empty) > 0 Then
                        strLineaFichero = strLineaFichero & Format(CLng(drRowExportacion("IDBancoFactura") & String.Empty), FORMATO_BANCO)
                    Else
                        strLineaFichero = strLineaFichero & FORMATO_BANCO
                    End If
                    If Length(drRowExportacion("SucursalFactura") & String.Empty) > 0 Then
                        strLineaFichero = strLineaFichero & Format(CLng(drRowExportacion("SucursalFactura") & String.Empty), FORMATO_SUCURSAL)
                    Else
                        strLineaFichero = strLineaFichero & FORMATO_SUCURSAL
                    End If
                    If Length(drRowExportacion("NCuentaFactura") & String.Empty) > 0 Then
                        strLineaFichero = strLineaFichero & Format(CLng(drRowExportacion("NCuentaFactura") & String.Empty), FORMATO_NCUENTA)
                    Else
                        strLineaFichero = strLineaFichero & FORMATO_NCUENTA
                    End If

                    Dim strCVentaFact As String = Left(drRowExportacion("CondicionVentaFactoring") & String.Empty, LONGITUD_COND_VENTA_FACTORING)
                    strCVentaFact = strCVentaFact & Space(LONGITUD_COND_VENTA_FACTORING - Length(strCVentaFact))
                    strLineaFichero = strLineaFichero & strCVentaFact

                    Dim strIDCliente As String = Left(drRowExportacion("IDCliente") & String.Empty, LONGITUD_IDCLIENTE)
                    strIDCliente = strIDCliente & Space(LONGITUD_IDCLIENTE - Length(strIDCliente))
                    strLineaFichero = strLineaFichero & strIDCliente

                    If Length(drRowExportacion("DigitoControlFactura") & String.Empty) > 0 Then
                        strLineaFichero = strLineaFichero & Format(CInt(drRowExportacion("DigitoControlFactura") & String.Empty), FORMATO_DIGITO_CONTROL)
                    Else
                        strLineaFichero = strLineaFichero & FORMATO_DIGITO_CONTROL
                    End If

                    Dim strPais As String = Left(drRowExportacion("CodigoIsoPais") & String.Empty, LONGITUD_PAIS)
                    strPais = strPais & Space(LONGITUD_PAIS - Length(strPais))
                    strLineaFichero = strLineaFichero & strPais

                    Dim strMoneda As String = Left(drRowExportacion("CodigoIsoMoneda") & String.Empty, LONGITUD_MONEDA)
                    strMoneda = strMoneda & Space(LONGITUD_MONEDA - Length(strMoneda))
                    strLineaFichero = strLineaFichero & strMoneda

                    '//Rellenamos con blancos
                    strLineaFichero = strLineaFichero & Space(LONGITUD_LINEA - Length(strLineaFichero))

                    '//Incluimos el registro en el DataTable
                    drNewLinea = objExportacionInfo.DatosExportacion.NewRow
                    drNewLinea("Linea") = strLineaFichero
                    objExportacionInfo.DatosExportacion.Rows.Add(drNewLinea)

                    '//Incrementamos el contador de registros
                    intTotalReg = intTotalReg + 1
                Next drRowExportacion
            Else
                ApplicationService.GenerateError(ERROR_INFO_FACTORING_BANCO_PROPIO)
            End If


            '///////////////////////// REGISTRO TOTAL REMESA /////////////////////////////
            strLineaFichero = PREFIJO_LINEA_TOTAL_REMESA
            strLineaFichero = strLineaFichero & Format(intTotalReg, FORMATO_CONTADOR_FACTURAS)
            strLineaFichero = strLineaFichero & "0000"
            strLineaFichero = strLineaFichero & Format(dblImporteTotal, FORMATO_IMPORTE_TOTAL)
            strLineaFichero = strLineaFichero & FORMATO_IMPORTE_TOTAL
            strLineaFichero = strLineaFichero & Space(LONGITUD_LINEA - Length(strLineaFichero))
            drNewLinea = objExportacionInfo.DatosExportacion.NewRow
            drNewLinea("Linea") = strLineaFichero
            objExportacionInfo.DatosExportacion.Rows.Add(drNewLinea)

            '//////////////////////// REGISTRO TOTAL SOPORTE /////////////////////////////
            strLineaFichero = PREFIJO_LINEA_TOTAL_SOPORTE
            strLineaFichero = strLineaFichero & "0001"
            strLineaFichero = strLineaFichero & Space(LONGITUD_LINEA - Length(strLineaFichero))
            drNewLinea = objExportacionInfo.DatosExportacion.NewRow
            drNewLinea("Linea") = strLineaFichero
            objExportacionInfo.DatosExportacion.Rows.Add(drNewLinea)
        End If
        objExportacionInfo.ExportacionCorrecta = True
        Return objExportacionInfo
    End Function

#End Region

#Region " Banco Sabadell Atlantico "

    Private Function ExportacionBancoSabadellAtlantico(ByVal strIDProcess As String, ByVal strNFactoring As String, ByVal objBPInfo As BancoPropio.BancoPropioFactoringInfo) As ExportacionFactoringInfo
        Const LONGITUD_LINEA As Integer = 200
        Dim NUMERO_REMESA As String = Strings.StrDup(34, "0") & "1"

        '//Prefijos
        Const PREFIJO_LINEA_CABECERA_FICHERO As String = "01"
        Const PREFIJO_LINEA_CABECERA_REMESA As String = "02"
        Const PREFIJO_LINEA_REGISTRO_DETALLE As String = "03"
        Const PREFIJO_LINEA_REGISTRO_DETALLE_PAGO_PARCIAL As String = "04"

        '//Formatos
        Const FORMATO_FECHA As String = "yyyy-MM-dd"
        Dim FORMATO_CONTADOR_FACTURAS As String = Strings.StrDup(5, "0")
        Dim FORMATO_CONTADOR_REGISTROS As String = Strings.StrDup(4, "0")
        Dim FORMATO_IMPORTE_TOTAL As String = Strings.StrDup(15, "0")
        Dim FORMATO_CLIENTE_FACTORING As String = Strings.StrDup(15, "0")

        '//Longitudes campos 
        Const LONGITUD_CIFNIF As Integer = 12
        Const LONGITUD_NFACTURA As Integer = 35
        Const LONGITUD_CODIGO_MEDIO_PAGO_DEUDOR As Integer = 3

        '//Otros 
        Const PAGOS_PARCIALES_FACTURA As String = "0001"
        Const REFERENCIA_INTERNA As String = "01"
        Const TOTAL_REMESAS As String = "0001"
        Const CODIGO_ISO_MONEDA As String = "EUR"

        Dim strCodigoIsoMoneda As String = CODIGO_ISO_MONEDA
        Dim objExportacionInfo As New ExportacionFactoringInfo
        objExportacionInfo.ExportacionCorrecta = False

        Dim objFilter As New Filter
        objFilter.Clear()
        objFilter.Add(New StringFilterItem("IDProcess", FilterOperator.Equal, strIDProcess))
        objFilter.Add(New NumberFilterItem("ImpVencimiento", FilterOperator.GreaterThan, 0))
        Dim dtExportacion As DataTable = New BE.DataEngine().Filter("FicheroFactoringBSabadellAtlantico", objFilter)
        If IsNothing(dtExportacion) OrElse dtExportacion.Rows.Count = 0 Then
            objExportacionInfo.DatosExportacion = Nothing
            '//47409: No hay datos a exportar.
            ApplicationService.GenerateError("No hay datos a exportar.")
        End If

        If Not IsNothing(objBPInfo) Then
            Dim strLineaFichero As String

            '//////////////////////  DEFINICION DATOS A RETORNAR  /////////////////////////
            '//Creamos la estructura del DataTable a retornar, con las líneas del fichero.
            Dim dColumn As New DataColumn
            dColumn.ColumnName = "Linea"
            dColumn.DataType = GetType(String)
            dColumn.MaxLength = LONGITUD_LINEA             'Longitud máxima de la línea

            objExportacionInfo.DatosExportacion = New DataTable
            objExportacionInfo.DatosExportacion.Columns.Clear()
            objExportacionInfo.DatosExportacion.Columns.Add(dColumn)

            '//DataTable auxiliar para ir contando los registros y los totales que deberán aparecer en los registros de la cabecera....
            Dim dtDatosExportLineas As New DataTable
            dtDatosExportLineas = objExportacionInfo.DatosExportacion.Clone

            '/////////////////////////// REGISTROS DE DETALLE  ///////////////////////////
            Dim intTotalReg As Integer : Dim intTotalFact As Integer : Dim dblImporteTotal As Double
            Dim strNFacturaAnt As String = String.Empty
            If Not IsNothing(dtExportacion) AndAlso dtExportacion.Rows.Count > 0 Then
                strCodigoIsoMoneda = dtExportacion.Rows(0)("CodigoISOMoneda") & String.Empty
                For Each drRowExportacion As DataRow In dtExportacion.Rows

                    If strNFacturaAnt <> drRowExportacion("NFActura") Then
                        strLineaFichero = PREFIJO_LINEA_REGISTRO_DETALLE
                        strLineaFichero = strLineaFichero & NUMERO_REMESA

                        Dim strNFactura As String = Left(drRowExportacion("NFactura") & String.Empty, LONGITUD_NFACTURA)
                        strNFactura = strNFactura & Space(LONGITUD_NFACTURA - Length(strNFactura))
                        strLineaFichero = strLineaFichero & strNFactura

                        'Fecha de Emisión de la Factura
                        Dim strFechaFcta As String = Format(CDate(drRowExportacion("FechaFactura") & String.Empty), FORMATO_FECHA)
                        strFechaFcta = strFechaFcta & Space(Length(FORMATO_FECHA) - Length(strFechaFcta))
                        strLineaFichero = strLineaFichero & strFechaFcta

                        dblImporteTotal = dblImporteTotal + Nz(drRowExportacion("ImporteFactura"), 0)
                        strLineaFichero = strLineaFichero & Format((Nz(drRowExportacion("ImporteFactura"), 0) * 100), FORMATO_IMPORTE_TOTAL)

                        Dim strCIFNIFDeudor As String = drRowExportacion("CifDeudor") & String.Empty
                        strCIFNIFDeudor = Replace(strCIFNIFDeudor, "/", String.Empty)
                        strCIFNIFDeudor = Replace(strCIFNIFDeudor, ".", String.Empty)
                        strCIFNIFDeudor = Replace(strCIFNIFDeudor, "-", String.Empty)
                        strCIFNIFDeudor = Left(strCIFNIFDeudor, LONGITUD_CIFNIF)
                        strCIFNIFDeudor = strCIFNIFDeudor & Space(LONGITUD_CIFNIF - Length(strCIFNIFDeudor))
                        strLineaFichero = strLineaFichero & strCIFNIFDeudor  'CIF/NIF Deudor

                        strLineaFichero = strLineaFichero & PAGOS_PARCIALES_FACTURA

                        Dim strFechaVto As String = Format(CDate(drRowExportacion("FechaVencimiento") & String.Empty), FORMATO_FECHA)
                        strFechaVto = strFechaVto & Space(Length(FORMATO_FECHA) - Length(strFechaVto))
                        strLineaFichero = strLineaFichero & strFechaVto

                        '//Rellenamos con blancos
                        strLineaFichero = strLineaFichero & Space(LONGITUD_LINEA - Length(strLineaFichero))

                        '//Incluimos el registro en el DataTable auxiliar
                        Dim drNewLineaAux As DataRow = dtDatosExportLineas.NewRow
                        drNewLineaAux("Linea") = strLineaFichero
                        dtDatosExportLineas.Rows.Add(drNewLineaAux)

                        intTotalReg = intTotalReg + 1
                        intTotalFact = intTotalFact + 1

                        '// A partir de ahora tendremos que controlar si ésta factura tiene más cobros, y si los tiene,
                        '// tendremos que ir a otra categoría de registros que son los de los cobros parciales.
                        '// Estos registros sólo se incluirán si se factorizan varios cobros de una misma factura.
                        Dim objFilterFact As New Filter
                        objFilterFact.Clear()
                        objFilterFact.Add(New StringFilterItem("NumeroFactura", drRowExportacion("NumeroFactura")))
                        Dim dvPagosParciales As DataView = New DataView(dtExportacion, objFilterFact.Compose(New AdoFilterComposer), Nothing, DataViewRowState.CurrentRows)
                        If Not IsNothing(dvPagosParciales) AndAlso dvPagosParciales.Count > 1 Then
                            For Each drvPagoParcial As DataRowView In dvPagosParciales
                                strLineaFichero = PREFIJO_LINEA_REGISTRO_DETALLE_PAGO_PARCIAL

                                strNFactura = Left(drvPagoParcial.Row("NFactura") & String.Empty, LONGITUD_NFACTURA)
                                strNFactura = strNFactura & Space(LONGITUD_NFACTURA - Length(strNFactura))
                                strLineaFichero = strLineaFichero & strNFactura

                                strLineaFichero = strLineaFichero & Format((Nz(drvPagoParcial.Row("ImpVencimiento"), 0) * 100), FORMATO_IMPORTE_TOTAL)

                                strFechaVto = Format(CDate(drvPagoParcial.Row("FechaVencimiento") & String.Empty), FORMATO_FECHA)
                                strFechaVto = strFechaVto & Space(Length(FORMATO_FECHA) - Length(strFechaVto))
                                strLineaFichero = strLineaFichero & strFechaVto

                                'Código del Medio de Pago del deudor.
                                strLineaFichero = strLineaFichero & Space(LONGITUD_CODIGO_MEDIO_PAGO_DEUDOR)

                                '//Rellenamos con blancos
                                strLineaFichero = strLineaFichero & Space(LONGITUD_LINEA - Length(strLineaFichero))

                                '//Incluimos el registro en el DataTable auxiliar
                                drNewLineaAux = dtDatosExportLineas.NewRow
                                drNewLineaAux("Linea") = strLineaFichero
                                dtDatosExportLineas.Rows.Add(drNewLineaAux)

                                intTotalReg = intTotalReg + 1
                            Next drvPagoParcial
                        End If
                    End If
                    strNFacturaAnt = drRowExportacion("NumeroFactura") & String.Empty
                Next drRowExportacion
            Else
                ApplicationService.GenerateError(ERROR_INFO_FACTORING_BANCO_PROPIO)
            End If

            '//////////////////////////////////////////////////////////////////////////////////////////////////////////
            '//NOTA: Los registros de Cabecera y de Remesa deben añadirse después, por que contienen Totales de los detalles.
            '/////////////////////// REGISTRO CABECERA DEL FICHERO  ////////////////////////
            strLineaFichero = PREFIJO_LINEA_CABECERA_FICHERO
            Dim strNifCedente As String ' Dim strNombreCedente As String :
            Dim objNegDatosEmpresa As New DatosEmpresa
            Dim dtDatosEmpresa As DataTable = objNegDatosEmpresa.Filter("DescEmpresa,Cif")
            If Not IsNothing(dtDatosEmpresa) AndAlso dtDatosEmpresa.Rows.Count > 0 Then
                strNifCedente = dtDatosEmpresa.Rows(0)("Cif") & String.Empty
                strNifCedente = Replace(strNifCedente, "/", String.Empty)
                strNifCedente = Replace(strNifCedente, ".", String.Empty)
                strNifCedente = Replace(strNifCedente, "-", String.Empty)
                strNifCedente = Left(strNifCedente, LONGITUD_CIFNIF)
                strNifCedente = strNifCedente & Space(LONGITUD_CIFNIF - Length(strNifCedente))
                strLineaFichero = strLineaFichero & strNifCedente
            End If
            objNegDatosEmpresa = Nothing

            strLineaFichero = strLineaFichero & REFERENCIA_INTERNA
            strLineaFichero = strLineaFichero & Space(3)
            'strLineaFichero = strLineaFichero & Format(objBPInfo.IDClienteFactoring & String.Empty, FORMATO_CLIENTE_FACTORING)
            Dim strIDClienteFactoring As String = Left(objBPInfo.IDClienteFactoring & String.Empty, Length(FORMATO_CLIENTE_FACTORING))
            strIDClienteFactoring = strIDClienteFactoring & Space(Length(FORMATO_CLIENTE_FACTORING) - Length(strIDClienteFactoring))
            strLineaFichero = strLineaFichero & strIDClienteFactoring

            strLineaFichero = strLineaFichero & TOTAL_REMESAS
            strLineaFichero = strLineaFichero & Format(intTotalReg + 2, FORMATO_CONTADOR_REGISTROS)
            strLineaFichero = strLineaFichero & Space(LONGITUD_LINEA - Length(strLineaFichero))  'Completamos con blancos la cadena

            objExportacionInfo.DatosExportacion.Rows.Clear()
            Dim drNewLinea As DataRow = objExportacionInfo.DatosExportacion.NewRow
            drNewLinea("Linea") = strLineaFichero
            objExportacionInfo.DatosExportacion.Rows.Add(drNewLinea)

            '/////////////////////// REGISTRO CABECERA DE REMESA  ////////////////////////
            strLineaFichero = PREFIJO_LINEA_CABECERA_REMESA
            strLineaFichero = strLineaFichero & NUMERO_REMESA
            strLineaFichero = strLineaFichero & Format(Today, FORMATO_FECHA)
            strLineaFichero = strLineaFichero & strCodigoIsoMoneda
            strLineaFichero = strLineaFichero & Format(intTotalFact, FORMATO_CONTADOR_FACTURAS)
            strLineaFichero = strLineaFichero & Format(dblImporteTotal * 100, FORMATO_IMPORTE_TOTAL)
            strLineaFichero = strLineaFichero & Space(LONGITUD_LINEA - Length(strLineaFichero))  'Completamos con blancos la cadena
            drNewLinea = objExportacionInfo.DatosExportacion.NewRow
            drNewLinea("Linea") = strLineaFichero
            objExportacionInfo.DatosExportacion.Rows.Add(drNewLinea)

            '//////////////////////////// REGISTROS DETALLE (TRASPASO) //////////////////////////////
            '//Añadimos los registros del DataTable de los detalles auxiliar, al DataTable a retornar.
            If Not IsNothing(dtDatosExportLineas) Then
                For Each drDatosExportLineas As DataRow In dtDatosExportLineas.Rows
                    drNewLinea = objExportacionInfo.DatosExportacion.NewRow
                    drNewLinea.ItemArray = drDatosExportLineas.ItemArray
                    objExportacionInfo.DatosExportacion.Rows.Add(drNewLinea)
                Next drDatosExportLineas
            End If

        End If
        objExportacionInfo.ExportacionCorrecta = True
        Return objExportacionInfo
    End Function

#End Region

#Region " Caja Madrid "

    Private Function ExportacionCajaMadrid(ByVal strIDProcess As String, ByVal strNFactoring As String, ByVal objBPInfo As BancoPropio.BancoPropioFactoringInfo) As ExportacionFactoringInfo
        Const LONGITUD_LINEA As Integer = 240

        '//Prefijos
        Const PREFIJO_LINEA_CABECERA_REMESA As String = "1"
        Const PREFIJO_LINEA_TOTAL_REMESA As String = "3"
        Const PREFIJO_LINEA_REGISTRO_DETALLE As String = "2"

        '//Formatos
        Const FORMATO_FECHA As String = "yyyyMMdd"
        Dim FORMATO_CONTADOR_FACTURAS As String = Strings.StrDup(5, "0")
        Dim FORMATO_IMPORTE_FACTURA As String = Strings.StrDup(12, "0")
        Dim FORMATO_IMPORTE_TOTAL As String = Strings.StrDup(14, "0")

        '//Longitudes campos 
        Const LONGITUD_CLIENTE_FACTORING As Integer = 11
        Const LONGITUD_CONTADOR_FACTORING As Integer = 15
        Const LONGITUD_IDCLIENTE As Integer = 3
        Const LONGITUD_CODIGO_FACTORING_CAJA_MADRID As Integer = 3
        Const LONGITUD_CIFNIF_DEUDOR As Integer = 10
        Const LONGITUD_TIPO_DOCUMENTO As Integer = 2
        Const LONGITUD_NFACTURA As Integer = 20
        Const LONGITUD_REF_EXTERNA As Integer = 20

        Dim objExportacionInfo As New ExportacionFactoringInfo
        objExportacionInfo.ExportacionCorrecta = False

        Dim objFilter As New Filter
        objFilter.Clear()
        objFilter.Add(New StringFilterItem("IDProcess", FilterOperator.Equal, strIDProcess))
        objFilter.Add(New NumberFilterItem("ImpVencimiento", FilterOperator.GreaterThan, 0))
        Dim dtExportacion As DataTable = New BE.DataEngine().Filter("FicheroFactoringCajaMadrid", objFilter)
        If IsNothing(dtExportacion) OrElse dtExportacion.Rows.Count = 0 Then
            objExportacionInfo.DatosExportacion = Nothing
            '//47409: No hay datos a exportar.
            ApplicationService.GenerateError("No hay datos a exportar.")
        End If

        If Not IsNothing(objBPInfo) Then
            Dim strLineaFichero As String

            '//////////////////////  DEFINICION DATOS A RETORNAR  /////////////////////////
            '//Creamos la estructura del DataTable a retornar, con las líneas del fichero.
            Dim dColumn As New DataColumn
            dColumn.ColumnName = "Linea"
            dColumn.DataType = GetType(String)
            dColumn.MaxLength = LONGITUD_LINEA             'Longitud máxima de la línea

            objExportacionInfo.DatosExportacion = New DataTable
            objExportacionInfo.DatosExportacion.Columns.Clear()
            objExportacionInfo.DatosExportacion.Columns.Add(dColumn)

            '/////////////////////// REGISTRO CABECERA DE REMESA  ////////////////////////
            Dim strIDClienteFactoring As String = Left(objBPInfo.IDClienteFactoring & String.Empty, LONGITUD_CLIENTE_FACTORING)
            strIDClienteFactoring = strIDClienteFactoring & Space(LONGITUD_CLIENTE_FACTORING - Length(strIDClienteFactoring))

            Dim strIDContadorFactoring As String = Left(objBPInfo.IDContadorFactoring & String.Empty, LONGITUD_CONTADOR_FACTORING)
            strIDContadorFactoring = strIDContadorFactoring & Space(LONGITUD_CONTADOR_FACTORING - Length(strIDContadorFactoring))

            Dim strIDCodigoFactoringCajaMadrid As String = String.Empty
            Dim strIDMoneda As String
            If Not IsNothing(dtExportacion) AndAlso dtExportacion.Rows.Count > 0 Then
                strIDCodigoFactoringCajaMadrid = Left(dtExportacion.Rows(0)("CodigoFactoringCajaMadrid") & String.Empty, LONGITUD_CODIGO_FACTORING_CAJA_MADRID)
                strIDMoneda = dtExportacion.Rows(0)("IDMoneda") & String.Empty
                For Each drRowExportacion As DataRow In dtExportacion.Rows
                    If strIDMoneda <> (drRowExportacion("IDMoneda") & String.Empty) Then
                        '//13375: Las facturas de la remesa deben estar expresadas en la misma divisa.
                        ApplicationService.GenerateError("Las facturas de la remesa deben estar expresadas en la misma divisa.")
                    End If
                Next drRowExportacion
            End If
            strIDCodigoFactoringCajaMadrid = strIDCodigoFactoringCajaMadrid & Space(LONGITUD_CODIGO_FACTORING_CAJA_MADRID - Length(strIDCodigoFactoringCajaMadrid))

            strLineaFichero = PREFIJO_LINEA_CABECERA_REMESA
            strLineaFichero = strLineaFichero & "00100"
            strLineaFichero = strLineaFichero & strIDClienteFactoring
            strLineaFichero = strLineaFichero & strIDContadorFactoring
            strLineaFichero = strLineaFichero & "00100"
            strLineaFichero = strLineaFichero & strIDClienteFactoring
            strLineaFichero = strLineaFichero & strIDContadorFactoring
            strLineaFichero = strLineaFichero & Format(Today, FORMATO_FECHA)
            strLineaFichero = strLineaFichero & strIDCodigoFactoringCajaMadrid
            strLineaFichero = strLineaFichero & "00000"
            strLineaFichero = strLineaFichero & Space(LONGITUD_LINEA - Length(strLineaFichero))  'Completamos con blancos la cadena

            Dim drNewLinea As DataRow = objExportacionInfo.DatosExportacion.NewRow
            drNewLinea("Linea") = strLineaFichero
            objExportacionInfo.DatosExportacion.Rows.Add(drNewLinea)

            '/////////////////////////// REGISTROS DE DETALLE  ///////////////////////////
            Dim intTotalReg As Integer : Dim dblImporteTotal As Double
            Dim strNFacturaAnt As String = String.Empty
            If Not IsNothing(dtExportacion) AndAlso dtExportacion.Rows.Count > 0 Then
                For Each drRowExportacion As DataRow In dtExportacion.Rows
                    strLineaFichero = PREFIJO_LINEA_REGISTRO_DETALLE
                    strLineaFichero = strLineaFichero & "00100"
                    strLineaFichero = strLineaFichero & strIDClienteFactoring
                    strLineaFichero = strLineaFichero & strIDContadorFactoring

                    '////////  Datos deudor  //////////
                    Dim strCIFNIFDeudor As String = Left(drRowExportacion("CifDeudor") & String.Empty, LONGITUD_CIFNIF_DEUDOR)
                    strCIFNIFDeudor = strCIFNIFDeudor & Space(LONGITUD_CIFNIF_DEUDOR - Length(strCIFNIFDeudor))
                    strLineaFichero = strLineaFichero & strCIFNIFDeudor  'CIF/NIF Deudor

                    Dim strIDCliente As String = Left(drRowExportacion("IDCliente") & String.Empty, LONGITUD_IDCLIENTE)
                    strIDCliente = strIDCliente & Space(LONGITUD_IDCLIENTE - Length(strIDCliente))
                    strLineaFichero = strLineaFichero & strIDCliente

                    Dim strTipoDocumento As String = String.Empty
                    strTipoDocumento = strTipoDocumento & Space(LONGITUD_TIPO_DOCUMENTO - Length(strTipoDocumento))
                    strLineaFichero = strLineaFichero & strTipoDocumento

                    '///////  Datos Factura  ///////////
                    Dim strNFactura As String = Left(drRowExportacion("NFactura") & String.Empty, LONGITUD_NFACTURA)
                    strNFactura = strNFactura & Space(LONGITUD_NFACTURA - Length(strNFactura))
                    strLineaFichero = strLineaFichero & strNFactura

                    Dim strRefExterna As String = String.Empty
                    strRefExterna = strRefExterna & Space(LONGITUD_REF_EXTERNA - Length(strRefExterna))
                    strLineaFichero = strLineaFichero & strRefExterna

                    strLineaFichero = strLineaFichero & Format((Nz(drRowExportacion("ImporteFactura"), 0) * 100), FORMATO_IMPORTE_FACTURA)
                    If strNFacturaAnt <> drRowExportacion("NFactura") & String.Empty Then
                        dblImporteTotal = dblImporteTotal + (Nz(drRowExportacion("ImporteFactura"), 0) * 100)
                    End If
                    strNFacturaAnt = drRowExportacion("NFactura") & String.Empty

                    Dim strFechaFta As String = Format(CDate(drRowExportacion("FechaFactura") & String.Empty), FORMATO_FECHA)
                    strFechaFta = strFechaFta & Space(Length(FORMATO_FECHA) - Length(strFechaFta))
                    strLineaFichero = strLineaFichero & strFechaFta

                    Dim strFechaVto As String = Format(CDate(drRowExportacion("FechaVencimiento") & String.Empty), FORMATO_FECHA)
                    strFechaVto = strFechaVto & Space(Length(FORMATO_FECHA) - Length(strFechaVto))
                    strLineaFichero = strLineaFichero & strFechaVto

                    '//Rellenamos con blancos
                    strLineaFichero = strLineaFichero & Space(LONGITUD_LINEA - Length(strLineaFichero))

                    '//Incluimos el registro en el DataTable
                    drNewLinea = objExportacionInfo.DatosExportacion.NewRow
                    drNewLinea("Linea") = strLineaFichero
                    objExportacionInfo.DatosExportacion.Rows.Add(drNewLinea)

                    '//Incrementamos el contador de registros
                    intTotalReg = intTotalReg + 1
                Next drRowExportacion
            End If

            '///////////////////////// REGISTRO TOTAL REMESA /////////////////////////////
            strLineaFichero = PREFIJO_LINEA_TOTAL_REMESA
            strLineaFichero = strLineaFichero & "00100"
            strLineaFichero = strLineaFichero & strIDClienteFactoring
            strLineaFichero = strLineaFichero & strIDContadorFactoring
            strLineaFichero = strLineaFichero & Format(intTotalReg, FORMATO_CONTADOR_FACTURAS)
            strLineaFichero = strLineaFichero & Format(dblImporteTotal, FORMATO_IMPORTE_TOTAL)
            strLineaFichero = strLineaFichero & Space(LONGITUD_LINEA - Length(strLineaFichero))
            drNewLinea = objExportacionInfo.DatosExportacion.NewRow
            drNewLinea("Linea") = strLineaFichero
            objExportacionInfo.DatosExportacion.Rows.Add(drNewLinea)
        Else
            ApplicationService.GenerateError(ERROR_INFO_FACTORING_BANCO_PROPIO)
        End If
        objExportacionInfo.ExportacionCorrecta = True
        Return objExportacionInfo
    End Function

#End Region

#Region " Banesto "

    Private Function ExportacionBanesto(ByVal strIDProcess As String, ByVal strNFactoring As String, ByVal objBPInfo As BancoPropio.BancoPropioFactoringInfo) As ExportacionFactoringInfo
        Const LONGITUD_LINEA As Integer = 200

        '//Prefijos
        Const PREFIJO_LINEA_CABECERA_REMESA As String = "01"
        Const PREFIJO_LINEA_TOTAL_REMESA As String = "03"
        Const PREFIJO_LINEA_REGISTRO_DETALLE As String = "02"
        'Const PREFIJO_FACTURA As String = "F"

        '//Formatos
        Const FORMATO_FECHA As String = "yyyyMMdd"
        Dim FORMATO_CONTADOR_FACTORING As String = Strings.StrDup(5, "0")
        Dim FORMATO_CONTADOR_FACTURAS As String = Strings.StrDup(4, "0")
        Dim FORMATO_IMPORTE_TOTAL As String = Strings.StrDup(15, "0")

        Dim FORMATO_BANCO As String = Strings.StrDup(4, "0")
        Dim FORMATO_SUCURSAL As String = Strings.StrDup(4, "0")
        Dim FORMATO_DIGITO_CONTROL As String = Strings.StrDup(2, "0")
        Dim FORMATO_NCUENTA As String = Strings.StrDup(10, "0")

        '//Longitudes campos 
        Const LONGITUD_CLIENTE_FACTORING As Integer = 7
        Const LONGITUD_CIF As Integer = 9
        Const LONGITUD_TIPO_PERSONA As Integer = 1
        Const LONGITUD_NOMBRE_DEUDOR As Integer = 35
        Const LONGITUD_DOMICILIO_DEUDOR As Integer = 35
        Const LONGITUD_POBLACION_DEUDOR As Integer = 30
        Const LONGITUD_CPOSTAL_DEUDOR As Integer = 5
        Const LONGITUD_NFACTURA As Integer = 10
        Const LONGITUD_IMPORTE_NOMINAL As Integer = 15

        Dim objExportacionInfo As New ExportacionFactoringInfo
        objExportacionInfo.ExportacionCorrecta = False

        Dim objFilter As New Filter
        objFilter.Clear()
        objFilter.Add(New StringFilterItem("IDProcess", FilterOperator.Equal, strIDProcess))
        objFilter.Add(New NumberFilterItem("ImpVencimiento", FilterOperator.GreaterThan, 0))
        Dim dtExportacion As DataTable = New BE.DataEngine().Filter("FicheroFactoringBanesto", objFilter)
        If IsNothing(dtExportacion) OrElse dtExportacion.Rows.Count = 0 Then
            objExportacionInfo.DatosExportacion = Nothing
            '//47409: No hay datos a exportar.
            ApplicationService.GenerateError("No hay datos a exportar.")
        End If

        If Not IsNothing(objBPInfo) Then
            Dim strLineaFichero As String = String.Empty

            '//////////////////////  DEFINICION DATOS A RETORNAR  /////////////////////////
            '//Creamos la estructura del DataTable a retornar, con las líneas del fichero.
            Dim dColumn As New DataColumn
            dColumn.ColumnName = "Linea"
            dColumn.DataType = GetType(String)
            dColumn.MaxLength = LONGITUD_LINEA             'Longitud máxima de la línea

            objExportacionInfo.DatosExportacion = New DataTable
            objExportacionInfo.DatosExportacion.Columns.Clear()
            objExportacionInfo.DatosExportacion.Columns.Add(dColumn)

            Dim strIDCodigoISO As String = String.Empty
            Dim strIDMoneda As String = String.Empty
            If Not IsNothing(dtExportacion) AndAlso dtExportacion.Rows.Count > 0 Then
                'strIDCodigoISO = Left(dtExportacion.Rows(0)("CodigoISO") & String.Empty, LONGITUD_CODIGO_ISO)
                strIDMoneda = dtExportacion.Rows(0)("IDMoneda") & String.Empty
                For Each drRowExportacion As DataRow In dtExportacion.Rows
                    If strIDMoneda <> (drRowExportacion("IDMoneda") & String.Empty) Then
                        '//13375: Las facturas de la remesa deben estar expresadas en la misma divisa.
                        ApplicationService.GenerateError("Las facturas de la remesa deben estar expresadas en la misma divisa.")
                    End If
                Next drRowExportacion
            End If


            '/////////////////////// REGISTRO CABECERA DE REMESA  ////////////////////////
            strLineaFichero = PREFIJO_LINEA_CABECERA_REMESA

            Dim strIDClienteFactoring As String = Left(objBPInfo.IDClienteFactoring & String.Empty, LONGITUD_CLIENTE_FACTORING)
            strIDClienteFactoring = strIDClienteFactoring & Space(LONGITUD_CLIENTE_FACTORING - Length(strIDClienteFactoring))
            strLineaFichero = strLineaFichero & strIDClienteFactoring

            strLineaFichero = strLineaFichero & Space(LONGITUD_LINEA - Length(strLineaFichero))  'Completamos con blancos la cadena
            objExportacionInfo.DatosExportacion.Rows.Clear()
            Dim drNewLinea As DataRow = objExportacionInfo.DatosExportacion.NewRow
            drNewLinea("Linea") = strLineaFichero
            objExportacionInfo.DatosExportacion.Rows.Add(drNewLinea)


            '/////////////////////////// REGISTROS DE DETALLE  ///////////////////////////
            Dim intTotalReg As Integer : Dim dblImporteTotal As Double
            If Not IsNothing(dtExportacion) AndAlso dtExportacion.Rows.Count > 0 Then
                For Each drRowExportacion As DataRow In dtExportacion.Rows
                    strLineaFichero = PREFIJO_LINEA_REGISTRO_DETALLE

                    Dim strCif As String = drRowExportacion("CifDeudor") & String.Empty
                    strCif = Left(strCif, LONGITUD_CIF)
                    strCif = strCif & Space(LONGITUD_CIF - Length(strCif))
                    strLineaFichero = strLineaFichero & strCif

                    If IsNumeric(Mid(strCif, 1, 1)) Then strLineaFichero = strLineaFichero & "F" & Space(LONGITUD_TIPO_PERSONA - 1) Else strLineaFichero = strLineaFichero & "J" & Space(LONGITUD_TIPO_PERSONA - 1)

                    Dim strNombreDeudor As String = Left(drRowExportacion("RazonSocial") & String.Empty, LONGITUD_NOMBRE_DEUDOR)
                    strNombreDeudor = strNombreDeudor & Space(LONGITUD_NOMBRE_DEUDOR - Length(strNombreDeudor))
                    strLineaFichero = strLineaFichero & strNombreDeudor

                    Dim strDireccion As String = Left(drRowExportacion("Direccion") & String.Empty, LONGITUD_DOMICILIO_DEUDOR)
                    strDireccion = strDireccion & Space(LONGITUD_DOMICILIO_DEUDOR - Length(strDireccion))
                    strLineaFichero = strLineaFichero & strDireccion

                    Dim strPoblacion As String = Left(drRowExportacion("Poblacion") & String.Empty, LONGITUD_POBLACION_DEUDOR)
                    strPoblacion = strPoblacion & Space(LONGITUD_POBLACION_DEUDOR - Length(strPoblacion))
                    strLineaFichero = strLineaFichero & strPoblacion

                    Dim strCodPostal As String = Left(drRowExportacion("CodigoPostal") & String.Empty, LONGITUD_CPOSTAL_DEUDOR)
                    strCodPostal = strCodPostal & Space(LONGITUD_CPOSTAL_DEUDOR - Length(strCodPostal))
                    strLineaFichero = strLineaFichero & strCodPostal

                    Dim strNFactura As String = Left(drRowExportacion("NFactura") & String.Empty, LONGITUD_NFACTURA)
                    strNFactura = strNFactura & Space(LONGITUD_NFACTURA - Length(strNFactura))
                    strLineaFichero = strLineaFichero & strNFactura

                    Dim strFechaFactura As String = String.Empty
                    If Length(drRowExportacion("FechaFactura")) = 0 Then
                        strLineaFichero = strLineaFichero & "00000000"
                    Else
                        strFechaFactura = Format(CDate(drRowExportacion("FechaFactura")), FORMATO_FECHA)
                        strFechaFactura = strFechaFactura & Space(Length(FORMATO_FECHA) - Length(strFechaFactura))
                        strLineaFichero = strLineaFichero & strFechaFactura
                    End If
                    Dim strFechaVencimiento As String = String.Empty
                    If Length(drRowExportacion("FechaVencimiento")) = 0 Then
                        strLineaFichero = strLineaFichero & "00000000"
                    Else
                        strFechaVencimiento = Format(CDate(drRowExportacion("FechaVencimiento")), FORMATO_FECHA)
                        strFechaVencimiento = strFechaVencimiento & Space(Length(FORMATO_FECHA) - Length(strFechaVencimiento))
                        strLineaFichero = strLineaFichero & strFechaVencimiento
                    End If

                    strLineaFichero = strLineaFichero & "F"

                    dblImporteTotal = dblImporteTotal + Nz(drRowExportacion("ImpVencimiento")) * 100
                    Dim strImporteNominal As String = Left(Format$(Nz(drRowExportacion("ImpVencimiento") * 100), "000000000000000"), LONGITUD_IMPORTE_NOMINAL)
                    strImporteNominal = strImporteNominal & Space(LONGITUD_IMPORTE_NOMINAL - Length(strImporteNominal))
                    strLineaFichero = strLineaFichero & strImporteNominal

                    strIDMoneda = strIDMoneda & Space(2 - Length(strIDMoneda))
                    strLineaFichero = strLineaFichero & strIDMoneda

                    Dim strCodigoBanco As String = "0000"
                    strLineaFichero = strLineaFichero & strCodigoBanco

                    Dim strSucursal As String = "0000"
                    strLineaFichero = strLineaFichero & strSucursal

                    Dim strCuenta As String = "0000000000"
                    strLineaFichero = strLineaFichero & strCuenta

                    strLineaFichero = strLineaFichero & Space(LONGITUD_LINEA - Length(strLineaFichero))  'Completamos con blancos la cadena

                    '//Incluimos el registro en el DataTable
                    drNewLinea = objExportacionInfo.DatosExportacion.NewRow
                    drNewLinea("Linea") = strLineaFichero
                    objExportacionInfo.DatosExportacion.Rows.Add(drNewLinea)

                    '//Incrementamos el contador de registros
                    intTotalReg = intTotalReg + 1
                Next drRowExportacion
            Else
                ApplicationService.GenerateError(ERROR_INFO_FACTORING_BANCO_PROPIO)
            End If


            '///////////////////////// REGISTRO TOTAL REMESA /////////////////////////////
            strLineaFichero = PREFIJO_LINEA_TOTAL_REMESA
            'strLineaFichero = strLineaFichero & Format(intTotalReg, FORMATO_CONTADOR_FACTURAS)
            strLineaFichero = strLineaFichero & Space(LONGITUD_LINEA - Length(strLineaFichero))
            drNewLinea = objExportacionInfo.DatosExportacion.NewRow
            drNewLinea("Linea") = strLineaFichero
            objExportacionInfo.DatosExportacion.Rows.Add(drNewLinea)

        End If
        objExportacionInfo.ExportacionCorrecta = True
        Return objExportacionInfo
    End Function

#End Region

#End Region

End Class