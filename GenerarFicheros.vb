<Transactional()> _
Public Class GenerarFicheros
    Inherits ContextBoundObject

#Region "Variables Privadas / Estructuras"

    Private Enum enumInformeDec347
        decEstatal = 0
        decBizkaia = 1
        decGipuzkoa = 2
        decAlava = 3
        decNavarra = 4
    End Enum

    Private Enum enumDefFichero
        Fich68 = 0
        Fich34 = 1
    End Enum

#End Region

#Region " Ficheros Transferencia"

    <Task()> Public Shared Function AgruparGastos(ByVal Dt As DataTable, ByVal services As ServiceProvider) As DataTable
        Dim DtGastos As DataTable = Dt.Clone
        Dim StrOperario As String = String.Empty
        Dim i As Integer = -1
        For Each Dr As DataRow In Dt.Select("", "IDOperario")
            If StrOperario = Dr("IDOperario") Then
                DtGastos.Rows(i)("ImpVencimientoA") = DtGastos.Rows(i)("ImpVencimientoA") + Dr("ImpVencimientoA")
            Else
                StrOperario = Dr("IDOperario")
                DtGastos.Rows.Add(Dr.ItemArray)
                i += 1
            End If
        Next
        Return DtGastos
    End Function

    <Serializable()> _
    Public Class DatosCabTransfer
        Public IDReg As String
        Public CifEmpresa As String
        Public RefEmpresa As String
        Public NumReg As String
        Public Datos As String
    End Class

    <Serializable()> _
    Public Class DataFicheros
        Public IDProcess As Guid
        Public IDBancoPropio As String
        Public TipoFichero As enumTipoFicheroTrans
        Public FechaEmision As Date
        Public Fichero As String
    End Class

    <Task()> Public Shared Function GenerarFicheroTransferencia(ByVal DatosTransfer As DataFicheros, ByVal services As ServiceProvider) As DataTable
        Dim Long_CIF As Integer = 10
        Dim Long_CIF_Benef As Integer = 12
        Dim Formato_Fecha As String = "ddMMyy"
        Dim Long_Nombre As Integer = 43
        Dim Long_Nombre_Benef As Integer = 36
        Dim Long_Direccion As Integer = 43
        Dim Formato_Importe As String = "000000000000"
        Dim Long_Poblacion As Integer = 38
        Dim Long_Poblacion_Benef As Integer = 31
        Dim Long_Linea As Integer = 72
        If Length(DatosTransfer.FechaEmision) = 0 Then DatosTransfer.FechaEmision = cnMinDate
        Dim DtEmpresa As DataTable = New DatosEmpresa().Filter()
        If Not DtEmpresa Is Nothing AndAlso DtEmpresa.Rows.Count > 0 Then
            Dim dtPagos As DataTable
            If DatosTransfer.TipoFichero = CInt(enumTipoFicheroTrans.tftPago) Then
                dtPagos = New BE.DataEngine().Filter("frmPagoContGenerarFich", New GuidFilterItem("IdProcess", DatosTransfer.IDProcess))
            Else
                dtPagos = New BE.DataEngine().Filter("frmGastoContGenerarFich", New GuidFilterItem("IdProcess", DatosTransfer.IDProcess))
                dtPagos = ProcessServer.ExecuteTask(Of DataTable, DataTable)(AddressOf AgruparGastos, dtPagos, services)
            End If

            If dtPagos Is Nothing OrElse dtPagos.Rows.Count = 0 Then
                ApplicationService.GenerateError("No hay pagos seleccionados. No se generará el fichero.")
            Else
                Dim DtFichero As New DataTable
                DtFichero.RemotingFormat = SerializationFormat.Binary
                DtFichero.Columns.Add("Linea", GetType(String))

                Dim StrPagoAgrupado As String = New Parametro().NFacturaPagoAgupado
                Dim strEmpresaRegistro As String = Strings.Left(CStr(DtEmpresa.Rows(0)("Cif")), Long_CIF)
                If Length(strEmpresaRegistro) > Long_CIF Then
                    strEmpresaRegistro = Left(strEmpresaRegistro, Long_CIF)
                Else
                    strEmpresaRegistro = Space(Long_CIF - Length(strEmpresaRegistro)) & strEmpresaRegistro
                End If

                Dim strDescEmpresa As String = Strings.Left(New GenerarFicheros().TratarSimbolosEspeciales(CStr(DtEmpresa.Rows(0)("DescEmpresa"))), Long_Nombre)
                If Length(strDescEmpresa) > Long_Nombre Then
                    strDescEmpresa = Left(strDescEmpresa, Long_Nombre)
                Else
                    strDescEmpresa = strDescEmpresa & Space(Long_Nombre - Length(strDescEmpresa))
                End If

                Dim strDirEmpresa As String = Strings.Left(New GenerarFicheros().TratarSimbolosEspeciales(CStr(DtEmpresa.Rows(0)("Direccion"))), Long_Nombre)
                If Length(strDirEmpresa) > Long_Nombre Then
                    strDirEmpresa = Left(strDirEmpresa, Long_Nombre)
                Else
                    strDirEmpresa = strDirEmpresa & Space(Long_Nombre - Length(strDirEmpresa))
                End If

                Dim StrRegistro As String = String.Empty

                'Cogemos los datos del BancoPropio
                Dim StrDatosBanco As New EstDatosBancoFicheros
                StrDatosBanco.IDBanco = DatosTransfer.IDBancoPropio
                StrDatosBanco = ProcessServer.ExecuteTask(Of EstDatosBancoFicheros, EstDatosBancoFicheros)(AddressOf DatosBanco, StrDatosBanco, services)

                'Primer Registro de CABECERA Obligatorio
                'Dim RegFirstCab As New DatosCabTransfer
                StrRegistro = "0356"
                StrRegistro &= strEmpresaRegistro
                StrRegistro &= Space(12)
                StrRegistro &= "001"
                StrRegistro &= Format(Today.Date, Formato_Fecha)
                StrRegistro &= IIf(DatosTransfer.FechaEmision <> cnMinDate, Strings.Format(DatosTransfer.FechaEmision, Formato_Fecha), Strings.Format(Today.Date, Formato_Fecha))
                StrRegistro &= StrDatosBanco.Entidad & StrDatosBanco.Sucursal & StrDatosBanco.NCuenta & "0" & Strings.Space(3) & StrDatosBanco.DC & Space(7)
                Dim DrFirstCab As DataRow = DtFichero.NewRow()
                DrFirstCab("Linea") = StrRegistro
                DtFichero.Rows.Add(DrFirstCab)
                Dim IntRegTotales As Integer = 1

                'Segundo Registro CABECERA Obligatorio
                StrRegistro = "0356"
                StrRegistro &= strEmpresaRegistro
                StrRegistro &= Space(12)
                StrRegistro &= "002"
                StrRegistro &= strDescEmpresa
                Dim DrSeconCab As DataRow = DtFichero.NewRow
                DrSeconCab("Linea") = StrRegistro
                DtFichero.Rows.Add(DrSeconCab)
                IntRegTotales += 1

                'Tercer Registro CABECERA Obligatorio
                StrRegistro = "0356"
                StrRegistro &= strEmpresaRegistro
                StrRegistro &= Space(12)
                StrRegistro &= "003"
                StrRegistro &= strDirEmpresa
                Dim DrTerCab As DataRow = DtFichero.NewRow
                DrTerCab("Linea") = StrRegistro
                DtFichero.Rows.Add(DrTerCab)
                IntRegTotales += 1

                'Cuarto Registro CABECERA Obligatorio
                StrRegistro = "0356"
                StrRegistro &= strEmpresaRegistro
                StrRegistro &= Space(12)
                StrRegistro &= "004"
                StrRegistro &= FormatearNumeros(Nz(DtEmpresa.Rows(0)("CodPostal"), 0), 5)
                Dim strPoblacionEmp As String = Left(New GenerarFicheros().TratarSimbolosEspeciales(DtEmpresa.Rows(0)("Poblacion")) & String.Empty, Long_Poblacion)
                If Length(strPoblacionEmp) > Long_Poblacion Then
                    strPoblacionEmp = Left(strPoblacionEmp, Long_Poblacion)
                Else
                    strPoblacionEmp = strPoblacionEmp & Space(Long_Poblacion - Length(strPoblacionEmp))
                End If
                StrRegistro &= strPoblacionEmp
                Dim DrCuarCab As DataRow = DtFichero.NewRow
                DrCuarCab("Linea") = StrRegistro
                DtFichero.Rows.Add(DrCuarCab)
                IntRegTotales += 1

                Dim DblImporteTotal As Double
                Dim StrChequeOTransf, StrConcepto, StrImporteTotal As String
                Dim IntRegistros010 As Integer

                For Each drPagos As DataRow In dtPagos.Select
                    If DatosTransfer.TipoFichero = CInt(enumTipoFicheroTrans.tftPago) Then
                        StrChequeOTransf = IIf(drPagos("ChequeTalon"), "7", "6")
                    Else
                        StrChequeOTransf = IIf(Nz(drPagos("Trasferencia"), 0), "6", "7")
                    End If
                    Dim StrRefBeneficiario As String = Left(drPagos("Cif") & String.Empty, Long_CIF_Benef)
                    If Length(StrRefBeneficiario) > Long_CIF_Benef Then
                        StrRefBeneficiario = Left(StrRefBeneficiario, Long_CIF_Benef)
                    Else
                        StrRefBeneficiario = StrRefBeneficiario & Space(Long_CIF_Benef - Length(StrRefBeneficiario))
                    End If
                    Dim StrImporte As String = Format(drPagos("ImpVencimientoA") * 100, Formato_Importe)

                    'Primer Registro INDIVIDUAL Obligatorio
                    StrRegistro = "065" & StrChequeOTransf
                    StrRegistro &= strEmpresaRegistro
                    StrRegistro &= StrRefBeneficiario
                    StrRegistro &= "010"
                    StrRegistro &= StrImporte
                    StrRegistro &= FormatearNumeros(Nz(drPagos("IDBancoPago"), 0), 4) & FormatearNumeros(Nz(drPagos("SucursalPago"), 0), 4) & FormatearNumeros(Nz(drPagos("NCuentaPago"), 0), 10)
                    StrRegistro &= "1" & "9"
                    StrRegistro &= Space(2)
                    StrRegistro &= FormatearNumeros(Nz(drPagos("DCPago"), 0), 2)
                    StrRegistro &= Space(7)
                    Dim DrFirstInd As DataRow = DtFichero.NewRow
                    DrFirstInd("Linea") = StrRegistro
                    DtFichero.Rows.Add(DrFirstInd)

                    IntRegTotales += 1
                    IntRegistros010 += 1
                    DblImporteTotal += drPagos("ImpVencimientoA")
                    StrImporteTotal = Format(DblImporteTotal * 100, Formato_Importe)

                    'Segundo Registro INDIVIDUAL (Obligatorio)
                    StrRegistro = "065" & StrChequeOTransf
                    StrRegistro &= strEmpresaRegistro
                    StrRegistro &= StrRefBeneficiario
                    StrRegistro &= "011"
                    Dim strDescBeneficiario As String = Left(New GenerarFicheros().TratarSimbolosEspeciales(drPagos("DescBeneficiario")) & String.Empty, Long_Nombre_Benef)
                    If Length(strDescBeneficiario) > Long_Nombre_Benef Then
                        strDescBeneficiario = Left(strDescBeneficiario, Long_Nombre_Benef)
                    Else
                        strDescBeneficiario = strDescBeneficiario & Space(Long_Nombre_Benef - Length(strDescBeneficiario))
                    End If
                    StrRegistro &= strDescBeneficiario
                    StrRegistro &= Space(7)
                    Dim DrSeconInd As DataRow = DtFichero.NewRow
                    DrSeconInd("Linea") = StrRegistro
                    DtFichero.Rows.Add(DrSeconInd)
                    IntRegTotales += 1

                    'Tercer Registro INDIVIDUAL (Obligatorio)
                    StrRegistro = "065" & StrChequeOTransf
                    StrRegistro &= strEmpresaRegistro
                    StrRegistro &= StrRefBeneficiario
                    StrRegistro &= "012"
                    Dim strDirPago As String = String.Empty
                    If Length(drPagos("DireccionPago")) > 0 Then
                        strDirPago = Left(New GenerarFicheros().TratarSimbolosEspeciales(Nz(drPagos("DireccionPago"), String.Empty)) & String.Empty, Long_Nombre_Benef)
                    End If
                    If Length(strDirPago) > Long_Nombre_Benef Then
                        strDirPago = Left(strDirPago, Long_Nombre_Benef)
                    Else
                        strDirPago = strDirPago & Space(Long_Nombre_Benef - Length(strDirPago))
                    End If
                    StrRegistro &= strDirPago
                    StrRegistro &= Space(7)
                    Dim DrTerInd As DataRow = DtFichero.NewRow
                    DrTerInd("Linea") = StrRegistro
                    DtFichero.Rows.Add(DrTerInd)
                    IntRegTotales += 1

                    'Cuarto Registro INDIVIDUAL (Opcional)
                    Dim intRestoDireccion As Integer = Length(New GenerarFicheros().TratarSimbolosEspeciales(Nz(drPagos("DireccionPago"), String.Empty))) - Long_Nombre_Benef
                    If intRestoDireccion > 0 Then
                        StrRegistro = "065" & StrChequeOTransf
                        StrRegistro &= strEmpresaRegistro
                        StrRegistro &= StrRefBeneficiario
                        StrRegistro &= "013"
                        Dim strRestoDireccion As String = Left(Right(New GenerarFicheros().TratarSimbolosEspeciales(drPagos("DireccionPago")) & String.Empty, intRestoDireccion), Long_Nombre_Benef)
                        If Length(strRestoDireccion) > Long_Nombre_Benef Then
                            strRestoDireccion = Left(strRestoDireccion, Long_Nombre_Benef)
                        Else
                            strRestoDireccion = strRestoDireccion & Space(Long_Nombre_Benef - Length(strRestoDireccion))
                        End If
                        StrRegistro &= strRestoDireccion
                        StrRegistro &= Space(7)
                        Dim DrCuarInd As DataRow = DtFichero.NewRow
                        DrCuarInd("Linea") = StrRegistro
                        DtFichero.Rows.Add(DrCuarInd)
                        IntRegTotales += 1
                    End If

                    'Quinto Registro INDIVIDUAL (Obligatorio)
                    StrRegistro = "065" & StrChequeOTransf
                    StrRegistro &= strEmpresaRegistro
                    StrRegistro &= StrRefBeneficiario
                    StrRegistro &= "014"
                    StrRegistro &= FormatearNumeros(Nz(drPagos("CodPostal"), 0), 5)
                    Dim strPoblacionPago As String = String.Empty
                    If Length(drPagos("PoblacionPago")) > 0 Then strPoblacionPago = Left(New GenerarFicheros().TratarSimbolosEspeciales(drPagos("PoblacionPago")) & String.Empty, Long_Poblacion_Benef)
                    If Length(strPoblacionPago) > Long_Poblacion_Benef Then
                        strPoblacionPago = Left(strPoblacionPago, Long_Poblacion_Benef)
                    Else
                        strPoblacionPago = strPoblacionPago & Space(Long_Poblacion_Benef - Length(strPoblacionPago))
                    End If
                    StrRegistro &= strPoblacionPago
                    StrRegistro &= Space(7)
                    Dim DrQuinInd As DataRow = DtFichero.NewRow
                    DrQuinInd("Linea") = StrRegistro
                    DtFichero.Rows.Add(DrQuinInd)
                    IntRegTotales += 1


                    'Sexto Registro INDIVIDUAL (Opcional)
                    StrRegistro = "065" & StrChequeOTransf
                    StrRegistro &= strEmpresaRegistro
                    StrRegistro &= StrRefBeneficiario
                    StrRegistro &= "015"
                    Dim strProvinciaPago As String = String.Empty
                    If Length(drPagos("ProvinciaPago")) > 0 Then strProvinciaPago = Left(New GenerarFicheros().TratarSimbolosEspeciales(drPagos("ProvinciaPago")) & String.Empty, Long_Nombre_Benef)
                    If Length(strProvinciaPago) > Long_Nombre_Benef Then
                        strProvinciaPago = Left(strProvinciaPago, Long_Nombre_Benef)
                    Else
                        strProvinciaPago = strProvinciaPago & Space(Long_Nombre_Benef - Length(strProvinciaPago))
                    End If
                    StrRegistro &= strProvinciaPago
                    StrRegistro &= Space(7)
                    Dim DrLin15 As DataRow = DtFichero.NewRow
                    DrLin15("Linea") = StrRegistro
                    DtFichero.Rows.Add(DrLin15)
                    IntRegTotales += 1

                    'Septimo Registro INDIVIDUAL (Opcional)
                    If DatosTransfer.TipoFichero = enumTipoFicheroTrans.tftPago Then
                        If Length(drPagos("NFactura")) > 0 Then
                            If drPagos("NFactura") = StrPagoAgrupado Then
                                Dim GrupoPago As New EstGrupoPago
                                GrupoPago.IDPago = drPagos("IDPago")
                                GrupoPago.DefFichero = enumDefFichero.Fich34
                                StrConcepto = ProcessServer.ExecuteTask(Of EstGrupoPago, String)(AddressOf FacturasPagoAgrupado, GrupoPago, services)
                            Else : StrConcepto = "Factura Numero " & New GenerarFicheros().TratarSimbolosEspeciales(drPagos("SuFactura") & String.Empty)
                            End If
                        End If
                    Else
                        If Length(drPagos("NObra")) > 0 Then
                            StrConcepto = "Obra Numero " & New GenerarFicheros().TratarSimbolosEspeciales(drPagos("NObra"))
                        Else : StrConcepto = New GenerarFicheros().TratarSimbolosEspeciales(drPagos("Texto") & String.Empty)
                        End If
                    End If
                    If StrConcepto <> String.Empty Then
                        If Len(StrConcepto) < 44 Then
                            StrRegistro = "065" & StrChequeOTransf
                            StrRegistro &= strEmpresaRegistro
                            StrRegistro &= StrRefBeneficiario
                            StrRegistro &= "016"
                            If DatosTransfer.TipoFichero = enumTipoFicheroTrans.tftPago Then
                                StrRegistro &= StrConcepto
                            Else
                                StrRegistro &= "Obra Numero " & New GenerarFicheros().TratarSimbolosEspeciales(drPagos("NObra"))
                            End If
                            If StrRegistro.Length < Long_Linea Then StrRegistro &= Strings.Space(Long_Linea - StrRegistro.Length)
                            Dim DrSeptInd As DataRow = DtFichero.NewRow
                            DrSeptInd("Linea") = StrRegistro
                            DtFichero.Rows.Add(DrSeptInd)
                            IntRegTotales += 1
                        Else
                            Dim strconceptoAux, strconcepto1, strconcepto2 As String
                            Dim pos As Integer = 0
                            Dim posAnt As Integer = 0

                            'el campo concepto sólo puede tener 43 caracteres
                            'cuando se trata de pagos agrupados, las facturas van separadas por comas pero no se debe cortar ningún número de factura
                            'hay que buscar la última coma para partir por ahí el concepto
                            Dim i As Integer = 0
                            Dim LONG_CONCEPTO As Integer = 43
                            Dim IDs() As String = StrConcepto.Split(",")
                            strconceptoAux = String.Empty ' Left(strConcepto, 43)
                            For i = 0 To IDs.Length - 1
                                If Length(strconceptoAux) > 0 AndAlso Length(strconceptoAux) + 1 <= LONG_CONCEPTO Then
                                    If Length(strconceptoAux) + 1 + Length(IDs(i)) <= LONG_CONCEPTO Then
                                        strconceptoAux &= "," & IDs(i)
                                    Else
                                        Exit For
                                    End If
                                ElseIf Length(strconceptoAux) = 0 Then
                                    strconceptoAux &= IDs(i)
                                Else
                                    Exit For
                                End If
                            Next
                            strconcepto1 = Left(strconceptoAux & Space(LONG_CONCEPTO - Len(strconceptoAux)), LONG_CONCEPTO)

                            'lo mismo para el concepto2
                            strconceptoAux = String.Empty
                            For j As Integer = i To IDs.Length - 1
                                If Length(strconceptoAux) > 0 AndAlso Length(strconceptoAux) + 1 <= LONG_CONCEPTO Then
                                    If Length(strconceptoAux) + 1 + Length(IDs(j)) <= LONG_CONCEPTO Then
                                        strconceptoAux &= "," & IDs(j)
                                    Else
                                        Exit For
                                    End If
                                ElseIf Length(strconceptoAux) = 0 Then
                                    strconceptoAux &= IDs(j)
                                Else
                                    Exit For
                                End If
                            Next
                            strconcepto2 = Left(strconceptoAux & Space(LONG_CONCEPTO - Len(strconceptoAux)), LONG_CONCEPTO)

                            'linea 16
                            StrRegistro = "065" & StrChequeOTransf
                            StrRegistro &= strEmpresaRegistro
                            StrRegistro &= StrRefBeneficiario
                            StrRegistro &= "016"
                            If DatosTransfer.TipoFichero = enumTipoFicheroTrans.tftPago Then
                                StrRegistro &= New GenerarFicheros().TratarSimbolosEspeciales(strconcepto1)
                            Else
                                StrRegistro &= "Obra Numero " & New GenerarFicheros().TratarSimbolosEspeciales(drPagos("NObra"))
                            End If
                            If StrRegistro.Length < Long_Linea Then StrRegistro &= Strings.Space(Long_Linea - StrRegistro.Length)
                            Dim DrLin16 As DataRow = DtFichero.NewRow
                            DrLin16("Linea") = StrRegistro
                            DtFichero.Rows.Add(DrLin16)
                            IntRegTotales += 1

                            'linea 17
                            StrRegistro = "065" & StrChequeOTransf
                            StrRegistro &= strEmpresaRegistro
                            StrRegistro &= StrRefBeneficiario
                            StrRegistro &= "017"
                            If DatosTransfer.TipoFichero = enumTipoFicheroTrans.tftPago Then
                                StrRegistro &= strconcepto2
                            Else
                                StrRegistro &= "Obra Numero " & New GenerarFicheros().TratarSimbolosEspeciales(drPagos("NObra"))
                            End If
                            If StrRegistro.Length < Long_Linea Then StrRegistro &= Strings.Space(Long_Linea - StrRegistro.Length)
                            Dim DrLin17 As DataRow = DtFichero.NewRow
                            DrLin17("Linea") = StrRegistro
                            DtFichero.Rows.Add(DrLin17)
                            IntRegTotales += 1
                        End If
                    End If
                Next

                'Registros Totales (Obligatorio)
                StrRegistro = "0856"
                StrRegistro &= strEmpresaRegistro
                StrRegistro &= Space(12)
                StrRegistro &= Space(3)
                StrRegistro &= StrImporteTotal
                StrRegistro &= Format(IntRegistros010, "00000000")
                StrRegistro &= Format(IntRegTotales + 1, "0000000000")
                StrRegistro &= Space(13)
                Dim DrNewTot As DataRow = DtFichero.NewRow
                DrNewTot("Linea") = StrRegistro
                DtFichero.Rows.Add(DrNewTot)
                Return DtFichero
            End If
        End If
    End Function

    <Task()> Public Shared Function GenerarFicheroTransferenciaA(ByVal DatosTransfer As DataFicheros, ByVal services As ServiceProvider) As DataTable
        'Modelo A para Kutxa
        'Dim StrDireccionBanco, StrRefBeneficiario, StrCodPostalProveedor As String
        Dim Long_CIF As Integer = 10
        Dim Long_CIF_Benef As Integer = 12
        Dim Formato_Fecha As String = "ddMMyy"
        Dim Long_Nombre As Integer = 36
        Dim Long_Direccion As Integer = 36
        Dim Formato_Importe As String = "000000000000"
        Dim Long_Poblacion As Integer = 31
        Dim Long_SuFactura As Integer = 12
        Dim Long_Concepto As Integer = 36

        'Creamos el datatable del fichero
        Dim DtFichero As New DataTable

        If Length(DatosTransfer.FechaEmision) = 0 Then DatosTransfer.FechaEmision = cnMinDate

        DtFichero.RemotingFormat = SerializationFormat.Binary
        DtFichero.Columns.Add("Linea", GetType(String))

        'Cogemos los datos de la Empresa
        Dim dtEmpresa As DataTable = New DatosEmpresa().Filter()
        If Not dtEmpresa Is Nothing AndAlso dtEmpresa.Rows.Count > 0 Then
            Dim strEmpresaRegistro As String = Strings.Left(CStr(dtEmpresa.Rows(0)("Cif")), Long_CIF)
            If Length(strEmpresaRegistro) > Long_CIF Then
                strEmpresaRegistro = Left(strEmpresaRegistro, Long_CIF)
            Else
                strEmpresaRegistro = strEmpresaRegistro & Space(Long_CIF - Length(strEmpresaRegistro))
            End If
            Dim strDescEmpresa As String = Strings.Left(New GenerarFicheros().TratarSimbolosEspeciales(CStr(dtEmpresa.Rows(0)("DescEmpresa"))), Long_Nombre)
            If Length(strDescEmpresa) > Long_Nombre Then
                strDescEmpresa = Left(strDescEmpresa, Long_Nombre)
            Else
                strDescEmpresa = strDescEmpresa & Space(Long_Nombre - Length(strDescEmpresa))
            End If
            Dim strDirEmpresa As String = Strings.Left(New GenerarFicheros().TratarSimbolosEspeciales(CStr(dtEmpresa.Rows(0)("Direccion"))), Long_Nombre)
            If Length(strDirEmpresa) > Long_Nombre Then
                strDirEmpresa = Left(strDirEmpresa, Long_Nombre)
            Else
                strDirEmpresa = strDirEmpresa & Space(Long_Nombre - Length(strDirEmpresa))
            End If

            'Cogemos los Pagos seleccionados
            Dim DtPagos As DataTable
            If DatosTransfer.TipoFichero = enumTipoFicheroTrans.tftPago Then
                DtPagos = New BE.DataEngine().Filter("frmPagoContGenerarFich", New GuidFilterItem("IDProcess", DatosTransfer.IDProcess))
            Else
                DtPagos = New BE.DataEngine().Filter("frmGastoContGenerarFich", New GuidFilterItem("IDProcess", DatosTransfer.IDProcess))
                DtPagos = ProcessServer.ExecuteTask(Of DataTable, DataTable)(AddressOf AgruparGastos, DtPagos, services)
            End If

            If IsNothing(DtPagos) OrElse DtPagos.Rows.Count = 0 Then
                ApplicationService.GenerateError("No hay pagos seleccionados. No se generará el fichero.")
            Else
                'Cogemos los datos del BancoPropio

                Dim StrBancoPropio As New EstDatosBancoFicheros
                StrBancoPropio.IDBanco = DatosTransfer.IDBancoPropio
                StrBancoPropio = ProcessServer.ExecuteTask(Of EstDatosBancoFicheros, EstDatosBancoFicheros)(AddressOf DatosBanco, StrBancoPropio, services)

                'Primer Registro de CABECERA Obligatorio (Long = 72)
                Dim strRegistro As String = "0356"
                strRegistro &= strEmpresaRegistro
                strRegistro &= Space(12)
                strRegistro &= "001"
                strRegistro &= Format(Today, Formato_Fecha)
                strRegistro &= IIf(DatosTransfer.FechaEmision <> cnMinDate, Format(DatosTransfer.FechaEmision, Formato_Fecha), Format(Today, Formato_Fecha))
                strRegistro &= StrBancoPropio.Entidad
                strRegistro &= StrBancoPropio.Sucursal
                strRegistro &= StrBancoPropio.NCuenta
                strRegistro &= "0"
                strRegistro &= Space(3)
                strRegistro &= StrBancoPropio.DC
                strRegistro &= Space(7)
                Dim DrNew As DataRow = DtFichero.NewRow
                DrNew("Linea") = strRegistro
                DtFichero.Rows.Add(DrNew)

                Dim intRegTotales As Integer = 1

                'Segundo Registro CABECERA Obligatorio
                strRegistro = "0356"
                strRegistro &= strEmpresaRegistro
                strRegistro &= Space(12)
                strRegistro &= "002"
                strRegistro &= strDescEmpresa
                strRegistro &= Space(7)

                DrNew = DtFichero.NewRow
                DrNew("Linea") = strRegistro
                DtFichero.Rows.Add(DrNew)
                intRegTotales += 1

                'Tercer Registro CABECERA Obligatorio
                strRegistro = "0356"
                strRegistro &= strEmpresaRegistro
                strRegistro &= Space(12)
                strRegistro &= "003"
                strRegistro &= strDirEmpresa
                strRegistro &= Space(7)
                DrNew = DtFichero.NewRow
                DrNew("Linea") = strRegistro
                DtFichero.Rows.Add(DrNew)
                intRegTotales += 1

                'Cuarto Registro CABECERA Obligatorio
                strRegistro = "0356"
                strRegistro &= strEmpresaRegistro
                strRegistro &= Space(12)
                strRegistro &= "004"
                Dim strDireccionBanco As String = New GenerarFicheros().TratarSimbolosEspeciales(Left(StrBancoPropio.Direccion, Long_Direccion))
                If Length(strDireccionBanco) > Long_Direccion Then
                    strDireccionBanco = Left(strDireccionBanco, Long_Direccion)
                Else
                    strDireccionBanco = strDireccionBanco & Space(Long_Direccion - Length(strDireccionBanco))
                End If
                strRegistro &= strDireccionBanco
                strRegistro &= Space(7)
                DrNew = DtFichero.NewRow
                DrNew("Linea") = strRegistro
                DtFichero.Rows.Add(DrNew)
                intRegTotales += 1

                Dim DblImporteTotal As Double
                Dim IntRegistros010 As Integer
                Dim StrChequeOTransf, StrConcepto As String

                For Each DrPagos As DataRow In DtPagos.Select
                    If DatosTransfer.TipoFichero = enumTipoFicheroTrans.tftPago Then
                        StrChequeOTransf = "9"
                    Else : StrChequeOTransf = IIf(DrPagos("Trasferencia"), "7", "6")
                    End If

                    Dim StrRefBeneficiario As String = Left(DrPagos("Cif") & String.Empty, Long_CIF_Benef)
                    If Length(StrRefBeneficiario) > Long_CIF_Benef Then
                        StrRefBeneficiario = Left(StrRefBeneficiario, Long_CIF_Benef)
                    Else
                        StrRefBeneficiario = StrRefBeneficiario & Space(Long_CIF_Benef - Length(StrRefBeneficiario))
                    End If

                    Dim StrImporte As String = Format(DrPagos("ImpVencimientoA") * 100, Formato_Importe)

                    'Primer Registro INDIVIDUAL Obligatorio
                    strRegistro = "065" & StrChequeOTransf
                    strRegistro &= strEmpresaRegistro
                    strRegistro &= StrRefBeneficiario
                    strRegistro &= "010"
                    strRegistro &= StrImporte
                    strRegistro &= FormatearNumeros(Nz(DrPagos("IDBancoPago"), 0), 4) & FormatearNumeros(Nz(DrPagos("SucursalPago"), 0), 4) & FormatearNumeros(Nz(DrPagos("NCuentaPago"), 0), 10)
                    strRegistro &= "1" & "9"
                    strRegistro &= Space(2)
                    strRegistro &= FormatearNumeros(Nz(DrPagos("DCPago"), 0), 2)
                    strRegistro &= Space(7)
                    DrNew = DtFichero.NewRow
                    DrNew("Linea") = strRegistro
                    DtFichero.Rows.Add(DrNew)

                    intRegTotales += 1
                    IntRegistros010 += 1
                    DblImporteTotal += DrPagos("ImpVencimientoA")

                    'Segundo Registro INDIVIDUAL (Obligatorio)
                    strRegistro = "065" & StrChequeOTransf
                    strRegistro &= strEmpresaRegistro
                    strRegistro &= StrRefBeneficiario
                    strRegistro &= "011"
                    Dim strDescBeneficiario As String = Left(New GenerarFicheros().TratarSimbolosEspeciales(DrPagos("DescBeneficiario") & String.Empty), Long_Nombre)
                    If Length(strDescBeneficiario) > Long_Nombre Then
                        strDescBeneficiario = Left(strDescBeneficiario, Long_Nombre)
                    Else
                        strDescBeneficiario = strDescBeneficiario & Space(Long_Nombre - Length(strDescBeneficiario))
                    End If
                    strRegistro &= strDescBeneficiario
                    strRegistro &= Space(7)
                    DrNew = DtFichero.NewRow
                    DrNew("Linea") = strRegistro
                    DtFichero.Rows.Add(DrNew)
                    intRegTotales += 1

                    'Tercer Registro INDIVIDUAL (Obligatorio)
                    strRegistro = "065" & StrChequeOTransf
                    strRegistro &= strEmpresaRegistro
                    strRegistro &= StrRefBeneficiario
                    strRegistro &= "012"
                    Dim strDirPago As String = Left(New GenerarFicheros().TratarSimbolosEspeciales(DrPagos("DireccionPago") & String.Empty), Long_Nombre)
                    If Length(strDirPago) > Long_Nombre Then
                        strDirPago = Left(strDirPago, Long_Nombre)
                    Else
                        strDirPago = strDirPago & Space(Long_Nombre - Length(strDirPago))
                    End If
                    strRegistro &= strDirPago
                    strRegistro &= Space(7)
                    DrNew = DtFichero.NewRow
                    DrNew("Linea") = strRegistro
                    DtFichero.Rows.Add(DrNew)
                    intRegTotales += 1

                    'Cuarto Registro INDIVIDUAL (Opcional)
                    Dim intRestoDireccion As Integer = Length(New GenerarFicheros().TratarSimbolosEspeciales(DrPagos("DireccionPago"))) - Long_Nombre
                    If intRestoDireccion > 0 Then
                        strRegistro = "065" & StrChequeOTransf
                        strRegistro &= strEmpresaRegistro
                        strRegistro &= StrRefBeneficiario
                        strRegistro &= "013"
                        Dim strRestoDireccion As String = Left(Right(New GenerarFicheros().TratarSimbolosEspeciales(DrPagos("DireccionPago") & String.Empty), intRestoDireccion), Long_Nombre)
                        If Length(strRestoDireccion) > Long_Nombre Then
                            strRestoDireccion = Left(strRestoDireccion, Long_Nombre)
                        Else
                            strRestoDireccion = strRestoDireccion & Space(Long_Nombre - Length(strRestoDireccion))
                        End If
                        strRegistro &= strRestoDireccion
                        strRegistro &= Space(7)
                        DrNew = DtFichero.NewRow
                        DrNew("Linea") = strRegistro
                        DtFichero.Rows.Add(DrNew)
                        intRegTotales += 1
                    End If

                    'Quinto Registro INDIVIDUAL (Obligatorio)
                    strRegistro = "065" & StrChequeOTransf
                    strRegistro &= strEmpresaRegistro
                    strRegistro &= StrRefBeneficiario
                    strRegistro &= "014"
                    strRegistro &= FormatearNumeros(Nz(DrPagos("CodPostal"), 0), 5)
                    Dim strPoblacionPago As String = Left(New GenerarFicheros().TratarSimbolosEspeciales(DrPagos("PoblacionPago") & String.Empty), Long_Poblacion)
                    If Length(strPoblacionPago) > Long_Poblacion Then
                        strPoblacionPago = Left(strPoblacionPago, Long_Poblacion)
                    Else
                        strPoblacionPago = strPoblacionPago & Space(Long_Poblacion - Length(strPoblacionPago))
                    End If
                    strRegistro &= strPoblacionPago
                    strRegistro &= Space(7)
                    DrNew = DtFichero.NewRow
                    DrNew("Linea") = strRegistro
                    DtFichero.Rows.Add(DrNew)
                    intRegTotales += 1

                    'Sexto Registro INDIVIDUAL (Opcional)
                    'si no hay provincia, no va el registro
                    If Length(DrPagos("ProvinciaPago")) > 0 Then
                        strRegistro = "065" & StrChequeOTransf
                        strRegistro &= strEmpresaRegistro
                        strRegistro &= StrRefBeneficiario
                        strRegistro &= "015"
                        Dim strProvinciaPago As String = Left(New GenerarFicheros().TratarSimbolosEspeciales(DrPagos("ProvinciaPago") & String.Empty), Long_Nombre)
                        If Length(strProvinciaPago) > Long_Nombre Then
                            strProvinciaPago = Left(strProvinciaPago, Long_Nombre)
                        Else
                            strProvinciaPago = strProvinciaPago & Space(Long_Nombre - Length(strProvinciaPago))
                        End If
                        strRegistro &= strProvinciaPago
                        strRegistro &= Space(7)
                        DrNew = DtFichero.NewRow
                        DrNew("Linea") = strRegistro
                        DtFichero.Rows.Add(DrNew)

                        intRegTotales += 1
                    End If

                    'Septimo Registro INDIVIDUAL (Opcional)
                    Dim strSuFactura As String = New GenerarFicheros().TratarSimbolosEspeciales(DrPagos("SuFactura") & String.Empty)
                    If Length(strSuFactura) > Long_SuFactura Then
                        strSuFactura = Left(strSuFactura, Long_SuFactura)
                    Else
                        strSuFactura = Replace(Space(Long_SuFactura - Length(strSuFactura)), Space(1), "0") & strSuFactura
                    End If
                    StrConcepto = strSuFactura
                    StrConcepto &= Format(DrPagos("FechaFactura"), Formato_Fecha)
                    StrConcepto &= Format(DrPagos("FechaVencimiento"), Formato_Fecha)
                    StrConcepto &= StrImporte
                    If Length(StrConcepto) > Long_Concepto Then
                        StrConcepto = Left(StrConcepto, Long_Concepto)
                    Else
                        StrConcepto = Space(Long_Concepto - Length(StrConcepto)) & StrConcepto
                    End If

                    If StrConcepto <> String.Empty Then
                        strRegistro = "065" & StrChequeOTransf
                        strRegistro &= strEmpresaRegistro
                        strRegistro &= StrRefBeneficiario
                        strRegistro &= "016"
                        strRegistro &= StrConcepto
                        strRegistro &= Space(7)
                        DrNew = DtFichero.NewRow
                        DrNew("Linea") = strRegistro
                        DtFichero.Rows.Add(DrNew)
                        intRegTotales += 1
                    End If
                Next

                'Registros Totales (Obligatorio)
                strRegistro = "0856"
                strRegistro &= strEmpresaRegistro
                strRegistro &= Space(12)
                strRegistro &= Space(3)
                strRegistro &= Format(DblImporteTotal * 100, "000000000000")
                strRegistro &= Format(IntRegistros010, "00000000")
                strRegistro &= Format(intRegTotales + 1, "0000000000")
                strRegistro &= Space(7)
                strRegistro &= Space(6)
                DrNew = DtFichero.NewRow
                DrNew("Linea") = strRegistro
                DtFichero.Rows.Add(DrNew)
                Return DtFichero
            End If
        End If
    End Function

    <Task()> Public Shared Function GenerarFicheroTransferenciaB(ByVal DatosTransfer As DataFicheros, ByVal services As ServiceProvider) As DataTable
        'ByVal StrIDProcess As String, ByVal StrIDBancoPropio As String, _
        '                      ByVal IntTipoFichero As enumTipoFicheroTrans, ByVal StrFichero As String) As DataTable
        'Modelo B para BBV y Caja Laboral
        Dim StrEmpresaRegistro, StrEntidad, StrSucursal, StrNCuenta, StrDC As String
        Dim StrDireccionBanco, StrRefBeneficiario, StrCalleEmpresa, StrCalleProveedor, _
        StrPobProveedor, StrProvProveedor, StrCuentaProveedor, StrCodPostalProveedor As String

        'Creamos el datatable del fichero
        Dim DtFichero As New DataTable
        DtFichero.Columns.Add("Linea", GetType(String))

        If Length(DatosTransfer.FechaEmision) = 0 Then DatosTransfer.FechaEmision = cnMinDate
        'Cogemos los datos de la Empresa
        Dim DtEmpresa As DataTable = AdminData.Filter("tbDatosEmpresa")
        If Not DtEmpresa Is Nothing AndAlso DtEmpresa.Rows.Count > 0 Then
            Dim drEmpresa As DataRow = DtEmpresa.Rows(0)
            StrEmpresaRegistro = Strings.Space(10 - Length(Nz(drEmpresa("Cif"), String.Empty))) & Strings.Left(Nz(drEmpresa("Cif"), String.Empty), 10)
            StrCalleEmpresa = Strings.Left(New GenerarFicheros().TratarSimbolosEspeciales(drEmpresa("Direccion")) & Strings.Space(72), 72)

            'Cogemos los Pagos seleccionados
            Dim dtPagos As DataTable
            If DatosTransfer.TipoFichero = enumTipoFicheroTrans.tftPago Then
                dtPagos = New BE.DataEngine().Filter("frmPagoContGenerarFich", New GuidFilterItem("IDProcess", DatosTransfer.IDProcess))
            Else
                dtPagos = New BE.DataEngine().Filter("frmGastoContGenerarFich", New GuidFilterItem("IDProcess", DatosTransfer.IDProcess))
                dtPagos = ProcessServer.ExecuteTask(Of DataTable, DataTable)(AddressOf AgruparGastos, dtPagos, services)
            End If

            If dtPagos Is Nothing OrElse dtPagos.Rows.Count = 0 Then
                ApplicationService.GenerateError("No hay pagos seleccionados. No se generará el fichero.")
            Else
                Dim StrPagoAgrupado As String = New Parametro().NFacturaPagoAgupado

                'Cogemos los datos del BancoPropio

                Dim strDatosBanco As New EstDatosBancoFicheros
                strDatosBanco.IDBanco = DatosTransfer.IDBancoPropio
                strDatosBanco = ProcessServer.ExecuteTask(Of EstDatosBancoFicheros, EstDatosBancoFicheros)(AddressOf DatosBanco, strDatosBanco, services)

                StrEntidad = FormatearNumeros(strDatosBanco.Entidad, 4)
                StrSucursal = FormatearNumeros(strDatosBanco.Sucursal, 4)
                StrNCuenta = FormatearNumeros(strDatosBanco.NCuenta, 10)
                StrDC = FormatearNumeros(strDatosBanco.DC, 2)
                StrDireccionBanco = strDatosBanco.Direccion

                'Primer Registro de CABECERA Obligatorio
                Dim strRegistro As String = "5101"
                strRegistro &= StrEmpresaRegistro
                strRegistro &= Strings.Format(Today, "ddMMyy")

                Dim Extra As DataRetornoExtraEFichero = ProcessServer.ExecuteTask(Of String, DataRetornoExtraEFichero)(AddressOf ExtraEFichero, DatosTransfer.Fichero, services)
                DatosTransfer.Fichero = Extra.Cadena
                strRegistro &= Strings.Left(Extra.NombreFichero & Strings.Space(8), 8)
                strRegistro &= Strings.Left(New GenerarFicheros().TratarSimbolosEspeciales(drEmpresa("DescEmpresa")) & Strings.Space(40), 40)
                strRegistro &= StrEntidad & StrSucursal & StrNCuenta & Strings.Space(4)
                If Length(strRegistro) < 90 Then
                    strRegistro &= (Strings.Space(90 - Length(strRegistro)))
                Else : strRegistro = Strings.Left(strRegistro, 90)
                End If
                Dim DrNew As DataRow = DtFichero.NewRow
                DrNew("Linea") = strRegistro
                DtFichero.Rows.Add(DrNew)
                Dim IntRegTotales As Integer = 1

                'Segundo Registro CABECERA Obligatorio
                strRegistro = "5102"
                strRegistro &= StrEmpresaRegistro
                strRegistro &= StrCalleEmpresa
                strRegistro &= Strings.Space(4)
                If Length(strRegistro) < 90 Then
                    strRegistro &= (Strings.Space(90 - Length(strRegistro)))
                Else : strRegistro = Strings.Left(strRegistro, 90)
                End If
                Dim DrNew1 As DataRow = DtFichero.NewRow
                DrNew1("Linea") = strRegistro
                DtFichero.Rows.Add(DrNew1)
                IntRegTotales += 1

                'Tercer Registro CABECERA Obligatorio
                strRegistro = "5103"
                strRegistro &= StrEmpresaRegistro
                strRegistro &= Strings.Trim(Nz(drEmpresa("CodPostal"), String.Empty)) & Strings.Space(5 - Length(Nz(drEmpresa("CodPostal"), String.Empty)))
                strRegistro &= Strings.Trim(Strings.Left(New GenerarFicheros().TratarSimbolosEspeciales(Nz(drEmpresa("Poblacion"), String.Empty)), 36)) & Strings.Space(36 - Length(Nz(drEmpresa("Poblacion"), String.Empty)))
                strRegistro &= Strings.Trim(Strings.Left(New GenerarFicheros().TratarSimbolosEspeciales(Nz(drEmpresa("Provincia"), String.Empty)), 25)) & Strings.Space(25 - Length(Nz(drEmpresa("Provincia"), String.Empty)))
                strRegistro &= Strings.Space(10)
                If Len(strRegistro) < 90 Then
                    strRegistro &= (Strings.Space(90 - Length(strRegistro)))
                Else : strRegistro = Strings.Left(strRegistro, 90)
                End If
                Dim DrNew2 As DataRow = DtFichero.NewRow
                DrNew2("Linea") = strRegistro
                IntRegTotales += 1

                Dim DblImporteTotal As Double
                Dim IntRegistros010 As Integer
                Dim StrConcepto As String

                For Each drPagos As DataRow In dtPagos.Select
                    StrRefBeneficiario = Strings.Left(Nz(drPagos("Cif")) & Strings.Space(10), 10)
                    StrCalleProveedor = Strings.Left(Nz(New GenerarFicheros().TratarSimbolosEspeciales(drPagos("DireccionPago"))), 72)
                    StrPobProveedor = Strings.Left(Nz(New GenerarFicheros().TratarSimbolosEspeciales(drPagos("PoblacionPago"))), 36)
                    StrProvProveedor = Strings.Left(Nz(New GenerarFicheros().TratarSimbolosEspeciales(drPagos("ProvinciaPago"))), 15)
                    StrCodPostalProveedor = Strings.Left(Nz(drPagos("CodPostal")), 5)
                    StrCuentaProveedor = FormatearNumeros(drPagos("IDBancoPago"), 4) & FormatearNumeros(drPagos("SucursalPago"), 4) & FormatearNumeros(drPagos("DCPago"), 2) & FormatearNumeros(drPagos("NCuentaPago"), 10)
                    Dim StrImporte As String = Strings.Format(drPagos("ImpVencimientoA") * 100, "0000000000")

                    'Primer Registro INDIVIDUAL Obligatorio
                    strRegistro = "5301"
                    strRegistro &= StrRefBeneficiario
                    strRegistro &= "00000"
                    strRegistro &= New String("0", 7 - Length(Strings.Trim(Strings.Left(Str(Nz(drPagos("IDPago"), 0)), 7)))) & Strings.Trim(Strings.Left(Str(Nz(drPagos("IDPago"), 0)), 7))
                    strRegistro &= Strings.Trim(Strings.Left(New GenerarFicheros().TratarSimbolosEspeciales(Nz(drPagos("DescBeneficiario"), String.Empty)), 40)) & Strings.Space(40 - Length(Strings.Trim(Strings.Left(Nz(drPagos("DescBeneficiario"), String.Empty), 40))))
                    strRegistro &= StrImporte
                    strRegistro &= IIf(drPagos("NFactura") = StrPagoAgrupado, Strings.Format(drPagos("FechaVencimiento"), "ddMMyy"), Strings.Format(drPagos("FechaFactura"), "ddMMyy"))
                    strRegistro &= Strings.Format(drPagos("FechaVencimiento"), "ddMMyy")
                    If DatosTransfer.TipoFichero = enumTipoFicheroTrans.tftPago Then
                        strRegistro &= "1"
                    Else : strRegistro &= "2"
                    End If
                    strRegistro &= " "
                    If Length(strRegistro) < 90 Then
                        strRegistro &= (Strings.Space(90 - Length(strRegistro)))
                    Else
                        strRegistro = Strings.Left(strRegistro, 90)
                    End If
                    Dim DrNew3 As DataRow = DtFichero.NewRow
                    DrNew3("Linea") = strRegistro
                    DtFichero.Rows.Add(DrNew3)

                    IntRegTotales += 1
                    IntRegistros010 += 1
                    DblImporteTotal += drPagos("ImpVencimientoA")

                    'Segundo Registro INDIVIDUAL (Obligatorio)
                    strRegistro = "5302"
                    strRegistro &= StrRefBeneficiario
                    strRegistro &= StrCalleProveedor
                    strRegistro &= Strings.Space(4)
                    If Length(strRegistro) < 90 Then
                        strRegistro &= (Strings.Space(90 - Length(strRegistro)))
                    Else : strRegistro = Strings.Left(strRegistro, 90)
                    End If
                    Dim DrNew4 As DataRow = DtFichero.NewRow
                    DrNew4("Linea") = strRegistro
                    DtFichero.Rows.Add(DrNew4)
                    IntRegTotales += 1

                    'Tercer Registro INDIVIDUAL (Obligatorio)
                    strRegistro = "5303"
                    strRegistro &= StrRefBeneficiario
                    strRegistro &= StrCodPostalProveedor
                    strRegistro &= StrPobProveedor
                    strRegistro &= StrProvProveedor
                    strRegistro &= StrCuentaProveedor
                    If Length(strRegistro) < 90 Then
                        strRegistro &= (Strings.Space(90 - Length(strRegistro)))
                    Else : strRegistro = Strings.Left(strRegistro, 90)
                    End If
                    Dim DrNew5 As DataRow = DtFichero.NewRow
                    DrNew5("Linea") = strRegistro
                    DtFichero.Rows.Add(DrNew5)
                    IntRegTotales += 1

                    'Primer Registro individual (Opcional)
                    If drPagos("NFactura") = StrPagoAgrupado Then
                        Dim Sql As String = "SELECT tbPago.NFactura, tbPago.FechaVencimiento, tbFacturaCompraCabecera.FechaFactura, tbPago.IDPagoAgrupado" & " FROM tbPago LEFT OUTER JOIN tbFacturaCompraCabecera ON tbPago.IDFactura = tbFacturaCompraCabecera.IDFactura" & " where tbPago.IdPagoAgrupado=" & drPagos("IDPago")
                        Dim DtAgrupados As DataTable = AdminData.Execute(Sql, ExecuteCommand.ExecuteReader)

                        If Not DtAgrupados Is Nothing AndAlso DtAgrupados.Rows.Count > 0 Then
                            Dim nreg As Integer = 1
                            For Each drAgrupados As DataRow In DtAgrupados.Select
                                StrConcepto = "Factura " & New GenerarFicheros().TratarSimbolosEspeciales(drAgrupados("NFactura")) & " de Fecha " & drAgrupados("FechaFactura") & " de Fecha Vto " & New GenerarFicheros().TratarSimbolosEspeciales(drAgrupados("FechaVencimiento"))
                                strRegistro = "56" & IIf(nreg < 10, "0" & nreg, nreg)
                                nreg += 1
                                strRegistro &= StrRefBeneficiario
                                strRegistro &= Strings.Left(StrConcepto & Strings.Space(72), 72)
                                If Len(strRegistro) < 90 Then
                                    strRegistro &= (Strings.Space(90 - Length(strRegistro)))
                                Else : strRegistro = Strings.Left(strRegistro, 90)
                                End If
                                Dim DrNew6 As DataRow = DtFichero.NewRow
                                DrNew6("Linea") = strRegistro
                                DtFichero.Rows.Add(DrNew6)
                                IntRegTotales += 1
                            Next
                        End If
                    Else
                        'pago no agrupado
                        If DatosTransfer.TipoFichero = enumTipoFicheroTrans.tftPago Then
                            If Length(drPagos("NFactura")) > 0 Then
                                StrConcepto = "Factura Numero " & New GenerarFicheros().TratarSimbolosEspeciales(drPagos("NFactura"))
                            Else : StrConcepto = New GenerarFicheros().TratarSimbolosEspeciales(drPagos("Texto") & String.Empty)
                            End If
                        Else
                            If Length(drPagos("NObra")) > 0 Then
                                StrConcepto = "Obra Numero " & New GenerarFicheros().TratarSimbolosEspeciales(drPagos("NObra"))
                            Else : StrConcepto = New GenerarFicheros().TratarSimbolosEspeciales(drPagos("Texto") & String.Empty)
                            End If
                        End If
                        strRegistro = "5601"
                        strRegistro &= StrRefBeneficiario
                        strRegistro &= Strings.Left(StrConcepto & Strings.Space(72), 72)
                        If Len(strRegistro) < 90 Then
                            strRegistro &= (Strings.Space(90 - Length(strRegistro)))
                        Else : strRegistro = Strings.Left(strRegistro, 90)
                        End If
                        Dim DrNew7 As DataRow = DtFichero.NewRow
                        DrNew7("Linea") = strRegistro
                        DtFichero.Rows.Add(DrNew7)
                        IntRegTotales += 1
                    End If
                Next

                'Registros Totales (Obligatorio)
                strRegistro = "5901"
                strRegistro &= StrEmpresaRegistro
                strRegistro &= Strings.Format(DblImporteTotal * 100, "000000000000") & Strings.Format(IntRegistros010, "000000000000") & Strings.Format(IntRegTotales + 1, "000000000000")
                If Length(strRegistro) < 90 Then
                    strRegistro &= (Strings.Space(90 - Length(strRegistro)))
                Else : strRegistro = Strings.Left(strRegistro, 90)
                End If
                Dim DrNew8 As DataRow = DtFichero.NewRow
                DrNew8("Linea") = strRegistro
                DtFichero.Rows.Add(DrNew8)
                Return DtFichero
            End If
        End If
    End Function

    <Task()> Public Shared Function GenerarFicheroTransferenciaC(ByVal DatosTransfer As DataFicheros, ByVal services As ServiceProvider) As DataTable
        'Modelo para Banco Guipuzcoano
        Dim StrEmpresaRegistro, StrEntidad, StrSucursal, StrNCuenta, StrDC, StrRefBeneficiario As String

        'Creamos el datatable del fichero
        Dim DtFichero As New DataTable
        DtFichero.Columns.Add("Linea", GetType(String))
        If Length(DatosTransfer.FechaEmision) = 0 Then DatosTransfer.FechaEmision = cnMinDate
        'Cogemos los datos de la Empresa
        Dim DtEmpresa As DataTable = AdminData.Filter("tbDatosEmpresa")
        If Not DtEmpresa Is Nothing AndAlso DtEmpresa.Rows.Count > 0 Then
            Dim DrEmpresa As DataRow = DtEmpresa.Rows(0)
            StrEmpresaRegistro = Strings.Left(Nz(DrEmpresa("Cif")) & "   ", 10)

            'Cogemos los Pagos seleccionados
            Dim DtPagos As DataTable
            If DatosTransfer.TipoFichero = enumTipoFicheroTrans.tftPago Then
                DtPagos = New BE.DataEngine().Filter("frmPagoContGenerarFich", New GuidFilterItem("IDProcess", DatosTransfer.IDProcess))
            Else
                DtPagos = New BE.DataEngine().Filter("frmGastoContGenerarFich", New GuidFilterItem("IDProcess", DatosTransfer.IDProcess))
                DtPagos = ProcessServer.ExecuteTask(Of DataTable, DataTable)(AddressOf AgruparGastos, DtPagos, services)
            End If

            If DtPagos Is Nothing OrElse DtPagos.Rows.Count = 0 Then
                ApplicationService.GenerateError("No hay pagos seleccionados. No se generará el fichero.")
            Else
                Dim strPagoAgrupado As String = New Parametro().NFacturaPagoAgupado

                'Cogemos los datos del BancoPropio

                Dim strDatosBanco As New EstDatosBancoFicheros
                strDatosBanco.IDBanco = DatosTransfer.IDBancoPropio
                strDatosBanco = ProcessServer.ExecuteTask(Of EstDatosBancoFicheros, EstDatosBancoFicheros)(AddressOf DatosBanco, strDatosBanco, services)

                StrEntidad = strDatosBanco.Entidad
                StrSucursal = strDatosBanco.Sucursal
                StrNCuenta = strDatosBanco.NCuenta
                StrDC = strDatosBanco.DC

                'Primer Registro de CABECERA Obligatorio
                Dim StrRegistro As String = "0356"
                StrRegistro &= StrEmpresaRegistro
                StrRegistro &= Strings.Space(12)
                StrRegistro &= "001"
                StrRegistro &= Strings.Format(Today, "ddMMyy") & IIf(DatosTransfer.FechaEmision <> cnMinDate, Strings.Format(DatosTransfer.FechaEmision, "ddMMyy"), Strings.Format(Today, "ddMMyy")) & StrEntidad & StrSucursal & StrNCuenta & "1" & "   " & StrDC
                If Length(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
                Dim DrNew As DataRow = DtFichero.NewRow
                DrNew("Linea") = StrRegistro
                DtFichero.Rows.Add(DrNew)
                Dim IntRegTotales As Integer = 1

                'Segundo Registro CABECERA Obligatorio
                StrRegistro = "0356"
                StrRegistro &= StrEmpresaRegistro
                StrRegistro &= Strings.Space(12)
                StrRegistro &= "002"
                StrRegistro &= Strings.Left(New GenerarFicheros().TratarSimbolosEspeciales(DrEmpresa("DescEmpresa")), 43)
                If Len(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
                Dim DrNew1 As DataRow = DtFichero.NewRow
                DrNew1("Linea") = StrRegistro
                DtFichero.Rows.Add(DrNew1)
                IntRegTotales += 1

                'Tercer Registro CABECERA Obligatorio
                StrRegistro = "0356"
                StrRegistro &= StrEmpresaRegistro
                StrRegistro &= Strings.Space(12)
                StrRegistro &= "003"
                StrRegistro &= Strings.Left(New GenerarFicheros().TratarSimbolosEspeciales(Nz(DrEmpresa("Direccion"))), 43)
                If Length(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
                Dim DrNew2 As DataRow = DtFichero.NewRow
                DrNew2("Linea") = StrRegistro
                DtFichero.Rows.Add(DrNew2)
                IntRegTotales += 1

                'Cuarto Registro CABECERA Obligatorio
                StrRegistro = "0356"
                StrRegistro &= StrEmpresaRegistro
                StrRegistro &= Strings.Space(12)
                StrRegistro &= "004"
                StrRegistro &= Strings.Left(Nz(DrEmpresa("CodPostal")), 5) & Strings.Left(New GenerarFicheros().TratarSimbolosEspeciales(Nz(DrEmpresa("Poblacion"))), 38)
                If Length(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
                Dim DrNew3 As DataRow = DtFichero.NewRow
                DrNew3("Linea") = StrRegistro
                DtFichero.Rows.Add(DrNew3)
                IntRegTotales += 1

                Dim DblImporteTotal As Double
                Dim IntRegistros010 As Integer
                Dim StrChequeOTransf, StrConcepto, StrImporteTotal As String

                For Each DrPagos As DataRow In DtPagos.Select
                    If DatosTransfer.TipoFichero = enumTipoFicheroTrans.tftPago Then
                        StrChequeOTransf = IIf(DrPagos("ChequeTalon"), "7", "6")
                    Else : StrChequeOTransf = IIf(DrPagos("Trasferencia"), "6", "7")
                    End If

                    StrRefBeneficiario = Strings.Left(Nz(DrPagos("Cif")), 12)
                    Dim StrImporte As String = Format(DrPagos("ImpVencimientoA") * 100, "000000000000")

                    'Primer Registro INDIVIDUAL Obligatorio
                    StrRegistro = "0660"
                    StrRegistro &= StrEmpresaRegistro
                    StrRegistro &= StrRefBeneficiario
                    StrRegistro &= "010"
                    StrRegistro &= StrImporte & FormatearNumeros(Nz(DrPagos("IDBancoPago"), 0), 4) & FormatearNumeros(Nz(DrPagos("SucursalPago"), 0), 4) & FormatearNumeros(Nz(DrPagos("NCuentaPago"), 0), 10) & "1" & "9" & "0"
                    If Length(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
                    Dim DrNew4 As DataRow = DtFichero.NewRow
                    DrNew4("Linea") = StrRegistro
                    DtFichero.Rows.Add(DrNew4)

                    IntRegTotales += 1
                    IntRegistros010 += 1
                    DblImporteTotal += DrPagos("ImpVencimientoA")
                    StrImporteTotal = Format(DblImporteTotal * 100, "000000000000")
                    'Segundo Registro INDIVIDUAL (Obligatorio)
                    StrRegistro = "0660"
                    StrRegistro &= StrEmpresaRegistro
                    StrRegistro &= StrRefBeneficiario
                    StrRegistro &= "011"
                    StrRegistro &= Strings.Left(New GenerarFicheros().TratarSimbolosEspeciales(Nz(DrPagos("DescBeneficiario"))), 36)
                    If Length(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
                    Dim DrNew5 As DataRow = DtFichero.NewRow
                    DrNew5("Linea") = StrRegistro
                    DtFichero.Rows.Add(DrNew5)
                    IntRegTotales += 1

                    'Tercer Registro INDIVIDUAL (Obligatorio)
                    StrRegistro = "0660"
                    StrRegistro &= StrEmpresaRegistro
                    StrRegistro &= StrRefBeneficiario
                    StrRegistro &= "012"
                    StrRegistro &= Strings.Left(New GenerarFicheros().TratarSimbolosEspeciales(Nz(DrPagos("DireccionPago"))), 36)
                    If Length(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
                    Dim DrNew6 As DataRow = DtFichero.NewRow
                    DrNew6("Linea") = StrRegistro
                    DtFichero.Rows.Add(DrNew6)
                    IntRegTotales += 1

                    'Cuarto Registro INDIVIDUAL (Opcional)
                    Dim IntRestoDireccion As Integer = Length(New GenerarFicheros().TratarSimbolosEspeciales(Nz(DrPagos("DireccionPago")))) - 36
                    If IntRestoDireccion > 0 Then
                        StrRegistro = "0660"
                        StrRegistro &= StrEmpresaRegistro
                        StrRegistro &= StrRefBeneficiario
                        StrRegistro &= "013"
                        StrRegistro &= Strings.Left(Strings.Right(New GenerarFicheros().TratarSimbolosEspeciales(Nz(DrPagos("DireccionPago"))), IntRestoDireccion), 36)
                        If Length(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
                        Dim DrNew7 As DataRow = DtFichero.NewRow
                        DrNew7("Linea") = StrRegistro
                        DtFichero.Rows.Add(DrNew7)
                        IntRegTotales += 1
                    End If

                    'Quinto Registro INDIVIDUAL (Obligatorio)
                    StrRegistro = "0660"
                    StrRegistro &= StrEmpresaRegistro
                    StrRegistro &= StrRefBeneficiario
                    StrRegistro &= "014"
                    StrRegistro &= Strings.Left(DrPagos("CodPostal") & Strings.Space(5), 5) & Strings.Left(New GenerarFicheros().TratarSimbolosEspeciales(Nz(DrPagos("PoblacionPago")), String.Empty), 36)
                    If Length(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
                    Dim DrNew8 As DataRow = DtFichero.NewRow
                    DrNew8("Linea") = StrRegistro
                    DtFichero.Rows.Add(DrNew8)
                    IntRegTotales += 1

                    'Sexto Registro INDIVIDUAL (Opcional)
                    StrRegistro = "0660"
                    StrRegistro &= StrEmpresaRegistro
                    StrRegistro &= StrRefBeneficiario
                    StrRegistro &= "015"
                    StrRegistro &= Strings.Left(New GenerarFicheros().TratarSimbolosEspeciales(Nz(DrPagos("ProvinciaPago"))), 36)
                    If Length(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
                    Dim DrNew9 As DataRow = DtFichero.NewRow
                    DrNew9("Linea") = StrRegistro
                    DtFichero.Rows.Add(DrNew9)
                    IntRegTotales += 1

                    'Septimo Registro INDIVIDUAL (Opcional)
                    If DatosTransfer.TipoFichero = enumTipoFicheroTrans.tftPago Then
                        If Length(DrPagos("NFactura")) > 0 Then
                            If DrPagos("NFactura") = strPagoAgrupado Then
                                Dim GrupoPago As New EstGrupoPago
                                GrupoPago.IDPago = DrPagos("IDPago")
                                GrupoPago.DefFichero = enumDefFichero.Fich34
                                StrConcepto &= ProcessServer.ExecuteTask(Of EstGrupoPago, String)(AddressOf FacturasPagoAgrupado, GrupoPago, services)
                            Else : StrConcepto = "Factura Numero " & DrPagos("NFactura")
                            End If
                        Else : StrConcepto = New GenerarFicheros().TratarSimbolosEspeciales(DrPagos("Texto") & String.Empty)
                        End If
                    Else
                        If Length(DrPagos("NObra")) > 0 Then
                            StrConcepto = "Obra Numero " & DrPagos("NObra")
                        Else : StrConcepto = New GenerarFicheros().TratarSimbolosEspeciales(DrPagos("Texto") & String.Empty)
                        End If
                    End If

                    If StrConcepto <> String.Empty Then
                        StrRegistro = "0660"
                        StrRegistro &= StrEmpresaRegistro
                        StrRegistro &= StrRefBeneficiario
                        StrRegistro &= "016"
                        If DatosTransfer.TipoFichero = enumTipoFicheroTrans.tftPago Then
                            StrRegistro &= StrConcepto
                        Else : StrRegistro &= "Obra Numero " & DrPagos("NObra")
                        End If
                        If Length(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
                        Dim DrNew10 As DataRow = DtFichero.NewRow
                        DrNew10("Linea") = StrRegistro
                        DtFichero.Rows.Add(DrNew10)
                        IntRegTotales += 1
                    End If

                    'fecha vencimiento
                    StrRegistro = "0660"
                    StrRegistro &= StrEmpresaRegistro
                    StrRegistro &= StrRefBeneficiario
                    StrRegistro &= "910"
                    StrRegistro &= Strings.Format(DrPagos("FechaVencimiento"), "ddMMyyyy")
                    If Len(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
                    Dim DrNew11 As DataRow = DtFichero.NewRow
                    DrNew11("Linea") = StrRegistro
                    DtFichero.Rows.Add(DrNew11)
                    IntRegTotales += 1
                Next

                'Registros Totales (Obligatorio)
                StrRegistro = "0856"
                StrRegistro &= StrEmpresaRegistro
                StrRegistro &= Strings.Space(12)
                StrRegistro &= Strings.Space(3)
                StrRegistro &= StrImporteTotal & Strings.Format(IntRegistros010, "00000000") & Strings.Format(IntRegTotales + 1, "0000000000")
                If Len(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
                Dim DrNew12 As DataRow = DtFichero.NewRow
                DrNew12("Linea") = StrRegistro
                DtFichero.Rows.Add(DrNew12)
                Return DtFichero
            End If
        End If
    End Function

#End Region

#Region "Ficheros Confirming - Remesas de Pagos "

    <Serializable()> _
    Public Class DatosConfirmingCab
        Public CIF As String
        Public Entidad As String
        Public Sucursal As String
        Public NCuenta As String
        Public DC As String
        Public CodClie As String
        Public TotalRemesa As String
        Public Fecha As String
        Public Moneda As String
        Public Espacios As String
    End Class

    <Serializable()> _
    Public Class DatosConfirmingLin
        Public CIF As String
        Public DescBeneficiario As String
        Public DireccionPago As String
        Public Espacio As String
        Public CodPostal As String
        Public PoblacionPago As String
        Public Espacios As String
        Public NFactura As String
        Public FechaFactura As String
        Public ImpVencimiento As String
        Public FechaVencimiento As String
        Public EspaciosFinales As String
        Public Telefono As String
        Public Fax As String
    End Class

    <Task()> Public Shared Function GenerarFicheroConfirming(ByVal DatosConf As DataFicheros, ByVal services As ServiceProvider) As DataTable
        '*******************************
        '*** Formato confirming BBVA ***
        '*******************************
        Dim Long_CIF As Integer = 9
        Dim Long_Nombre As Integer = 30
        Dim StDatosCab As New DatosConfirmingCab
        Dim StDatosLin As New DatosConfirmingLin
        Dim StDatosBanco As New EstDatosBancoFicheros
        Dim DblImp As Double
        Dim DtFichero As New DataTable
        DtFichero.RemotingFormat = SerializationFormat.Binary
        DtFichero.Columns.Add("Linea", GetType(String))
        If Length(DatosConf.FechaEmision) = 0 Then DatosConf.FechaEmision = cnMinDate
        Dim DtBancos As DataTable = New BancoPropio().Filter(New FilterItem("IDBancoPropio", FilterOperator.Equal, DatosConf.IDBancoPropio))
        If Not DtBancos Is Nothing AndAlso DtBancos.Rows.Count > 0 Then
            Select Case Nz(DtBancos.Rows(0)("TipoConfirming"), -1)
                Case 1
                    Return ProcessServer.ExecuteTask(Of DataFicheros, DataTable)(AddressOf GenerarFicheroConfirming1, DatosConf, services)
                Case 2
                    Return ProcessServer.ExecuteTask(Of DataFicheros, DataTable)(AddressOf GenerarFicheroConfirming2, DatosConf, services)
            End Select
        End If

        'Cogemos los datos de la Empresa
        Dim DtEmpresa As DataTable = New DatosEmpresa().Filter()
        Dim strEmpresaRegistro As String = Strings.Left(CStr(DtEmpresa.Rows(0)("Cif")), Long_CIF)
        If Length(strEmpresaRegistro) > Long_CIF Then
            strEmpresaRegistro = Left(strEmpresaRegistro, Long_CIF)
        Else
            strEmpresaRegistro = strEmpresaRegistro & Space(Long_CIF - Length(strEmpresaRegistro))
        End If

        'Cogemos los Pagos seleccionados
        Dim DtPagos As DataTable = New BE.DataEngine().Filter("frmPagoContGenerarFich", New GuidFilterItem("IDProcess", DatosConf.IDProcess))
        If DtPagos Is Nothing OrElse DtPagos.Rows.Count = 0 Then
            ApplicationService.GenerateError("No hay pagos seleccionados. No se generará el fichero.")
        End If

        'Cogemos los datos del BancoPropio
        Dim Long_CodClteBanco As Integer = 6
        StDatosBanco.IDBanco = DatosConf.IDBancoPropio
        StDatosBanco = ProcessServer.ExecuteTask(Of EstDatosBancoFicheros, EstDatosBancoFicheros)(AddressOf DatosBanco, StDatosBanco, services)

        If StDatosBanco.CodClie = Strings.Space(Long_CodClteBanco) Then
            ApplicationService.GenerateError("Debe introducir el código de cliente que somos para el banco en su ficha")
            Exit Function
        Else
            StDatosCab.CodClie = Strings.Space(Long_CodClteBanco - Strings.Len(Strings.Trim(StDatosBanco.CodClie))) & Strings.Trim(StDatosBanco.CodClie)
        End If


        Dim SqlTot As String = "SELECT SUM(tbPago.ImpVencimientoA) AS ImporteVto " & _
        "FROM xNumericCheck RIGHT OUTER JOIN tbPago ON xNumericCheck.IdEnlace = tbPago.IDPago " & _
        "WHERE xNumericCheck.IdProcess = '" & (DatosConf.IDProcess).ToString & "'"
        Dim DblTotRemesa As Double = AdminData.Execute(SqlTot, ExecuteCommand.ExecuteScalar)

        'Primer Registro de CABECERA Obligatorio
        StDatosCab.CIF = strEmpresaRegistro
        StDatosCab.Fecha = Strings.Format(Today, "ddMMyy")
        StDatosCab.TotalRemesa = Strings.Format(DblTotRemesa * 100, "0000000000000") 'Importe remesa
        StDatosCab.Moneda = "EUR"  'Moneda remesa
        StDatosCab.Espacios = Space(171)
        Dim DrNewCab As DataRow = DtFichero.NewRow
        DrNewCab("Linea") = StDatosCab.CodClie & StDatosCab.CIF & StDatosCab.Fecha & StDatosCab.TotalRemesa & StDatosCab.Moneda & StDatosCab.Espacios
        DtFichero.Rows.Add(DrNewCab)

        'Registro Detalle
        For Each Dr As DataRow In DtPagos.Select
            StDatosLin.CIF = Strings.Left(Nz(Dr("Cif"), ""), Long_CIF) & Strings.Space(Long_CIF - Strings.Len(Strings.Left(Nz(Dr("Cif"), ""), Long_CIF)))
            StDatosLin.DescBeneficiario = Strings.Left(Strings.Trim(Nz(Dr("DescBeneficiario"), "")), Long_Nombre) & Strings.Space(Long_Nombre - Len(Strings.Left(Strings.Trim(Nz(Dr("DescBeneficiario"), "")), Long_Nombre)))
            StDatosLin.DireccionPago = Strings.Left(Strings.Trim(Nz(Dr("DireccionPago"), "")), Long_Nombre) & Strings.Space(Long_Nombre - Len(Strings.Left(Strings.Trim(Nz(Dr("DireccionPago"), "")), Long_Nombre)))
            StDatosLin.Espacio = Strings.Space(7)
            StDatosLin.CodPostal = Strings.Left(Strings.Trim(Nz(Dr("CodPostal"), "")), 5) & Strings.Space(5 - Strings.Len(Strings.Left(Strings.Trim(Nz(Dr("CodPostal"), "")), 5)))
            StDatosLin.PoblacionPago = Strings.Left(Strings.Trim(Nz(Dr("PoblacionPago"), "")), 20) & Strings.Space(20 - Strings.Len(Strings.Left(Strings.Trim(Nz(Dr("PoblacionPago"), "")), 20)))
            StDatosLin.Espacios = Strings.Space(38)
            StDatosLin.NFactura = Strings.Left(Strings.Trim(Nz(Dr("NFactura"), "")), 11) & Strings.Space(11 - Strings.Len(Strings.Left(Strings.Trim(Nz(Dr("NFactura"), "")), 11)))
            StDatosLin.FechaFactura = Strings.Format(Dr("FechaFactura"), "ddMMyy")
            DblImp = Dr("ImpVencimientoA")
            StDatosLin.ImpVencimiento = Strings.Space(13 - Strings.Len(Strings.Trim(CStr(DblImp * 100)))) & Strings.Trim(CStr(DblImp * 100))
            StDatosLin.FechaVencimiento = Strings.Format(Dr("FechaVencimiento"), "ddMMyy")
            StDatosLin.EspaciosFinales = Strings.Space(31)
            Dim DrNewLin As DataRow = DtFichero.NewRow
            DrNewLin("Linea") = StDatosLin.CIF & StDatosLin.DescBeneficiario & StDatosLin.DireccionPago & _
            StDatosLin.Espacio & StDatosLin.CodPostal & StDatosLin.PoblacionPago & StDatosLin.Espacios & _
            StDatosLin.NFactura & StDatosLin.FechaFactura & StDatosLin.ImpVencimiento & StDatosLin.FechaVencimiento & _
            StDatosLin.EspaciosFinales
            DtFichero.Rows.Add(DrNewLin)
        Next
        Return DtFichero
    End Function

    <Task()> Public Shared Function GenerarFicheroConfirming1(ByVal DatosConf As DataFicheros, ByVal services As ServiceProvider) As DataTable
        '*******************************
        '*** Formato confirming BBVA ***
        '*******************************
        Dim StDatosCab As New DatosConfirmingCab
        Dim StDatosLin As New DatosConfirmingLin
        Dim StDatosBanco As New EstDatosBancoFicheros
        Dim StrPagoAgrupado As String = New Parametro().ObtenerPredeterminado("FRAPAGOAGR")
        Dim DtPagosAgrupados As DataTable
        Dim DtFichero As New DataTable
        DtFichero.Columns.Add("Linea", GetType(String))

        'Cogemos los datos de la Empresa
        Dim DtEmpresa As DataTable = AdminData.Filter("tbDatosEmpresa")
        Dim StrEmpReg As String = Strings.Left(Nz(DtEmpresa.Rows(0)("Cif")), 10)

        Dim DtPagos As DataTable = New BE.DataEngine().Filter("frmPagoContGenerarFich", New GuidFilterItem("IDProcess", DatosConf.IDProcess))
        If DtPagos Is Nothing OrElse DtPagos.Rows.Count = 0 Then
            ApplicationService.GenerateError("No hay pagos seleccionados. No se generará el fichero.")
            Exit Function
        End If

        'Cogemos los datos del BancoPropio
        StDatosBanco.IDBanco = DatosConf.IDBancoPropio
        StDatosBanco = ProcessServer.ExecuteTask(Of EstDatosBancoFicheros, EstDatosBancoFicheros)(AddressOf DatosBanco, StDatosBanco, services)

        If StDatosBanco.CodClie = Strings.Space(6) Then
            ApplicationService.GenerateError("Debe introducir el código de cliente que somos para el banco en su ficha")
            Exit Function
        Else
            StDatosCab.CodClie = Strings.Right(Strings.Space(6) & Strings.Trim(StDatosBanco.CodClie), 6)
        End If

        Dim SqlTot As String = "SELECT SUM(tbPago.ImpVencimientoA) AS ImporteVto " & _
        "FROM xNumericCheck RIGHT OUTER JOIN tbPago ON xNumericCheck.IdEnlace = tbPago.IDPago " & _
        "WHERE xNumericCheck.IdProcess = '" & (DatosConf.IDProcess).ToString & "'"
        Dim DblTotRemesa As Double = AdminData.Execute(SqlTot, ExecuteCommand.ExecuteScalar)

        'Primer Registro de CABECERA Obligatorio
        StDatosCab.CIF = Strings.Left(Nz(Strings.Trim(StrEmpReg), "") & Strings.Space(9), 9)
        StDatosCab.Fecha = Strings.Format(Today, "yyyyMMdd")
        StDatosCab.TotalRemesa = Strings.Format(DblTotRemesa * 100, "0000000000000") 'Importe remesa
        StDatosCab.Moneda = "EUR" 'Moneda remesa
        StDatosCab.Espacios = Strings.Space(171)
        Dim DrNewCab As DataRow = DtFichero.NewRow
        DrNewCab("Linea") = StDatosCab.CodClie & StDatosCab.CIF & StDatosCab.Fecha & StDatosCab.TotalRemesa & _
        StDatosCab.Moneda & StDatosCab.Espacios
        DtFichero.Rows.Add(DrNewCab)

        'Registro Detalle
        For Each Dr As DataRow In DtPagos.Select
            StDatosLin.CIF = Strings.Left(Nz(Dr("Cif"), "") & Strings.Space(9), 9)
            StDatosLin.DescBeneficiario = Strings.Left(Strings.Trim(Nz(Dr("DescBeneficiario"), "")) & Strings.Space(30), 30)
            StDatosLin.DireccionPago = Strings.Left(Strings.Trim(Nz(Dr("DireccionPago"), "")) & Strings.Space(30), 30)
            StDatosLin.Espacio = Strings.Space(7)
            StDatosLin.CodPostal = Strings.Left(Strings.Trim(Nz(Dr("CodPostal"), "")) & Strings.Space(5), 5)
            StDatosLin.PoblacionPago = Strings.Left(Strings.Trim(Nz(Dr("PoblacionPago"), "")) & Strings.Space(20), 20)
            StDatosLin.Telefono = Strings.Left(Strings.Trim(Nz(Dr("Telefono"), "")) & Strings.Space(10), 10)
            StDatosLin.Fax = Strings.Left(Strings.Trim(Nz(Dr("Fax"), "")) & Strings.Space(10), 10)
            StDatosLin.Espacios = Strings.Space(18)
            StDatosLin.NFactura = Strings.Left(Strings.Trim(Nz(Dr("NFactura"), "")) & Strings.Space(11), 11)
            If Dr("NFactura") = StrPagoAgrupado Then
                'pago agrupado: ponemos las fecha de factura de la primera factura
                Dim Sql As String = "SELECT tbPago.NFactura, tbPago.FechaVencimiento, " & _
                "tbFacturaCompraCabecera.FechaFactura, tbPago.IDPagoAgrupado " & _
                "FROM tbPago LEFT OUTER JOIN tbFacturaCompraCabecera ON tbPago.IDFactura = tbFacturaCompraCabecera.IDFactura " & _
                "WHERE tbPago.IdPagoAgrupado=" & Dr("IDPago")
                DtPagosAgrupados = AdminData.Execute(Sql, ExecuteCommand.ExecuteReader)
            End If
            If Dr("NFactura") = StrPagoAgrupado Then
                StDatosLin.FechaFactura = Strings.Format(DtPagosAgrupados.Rows(0)("FechaFactura"), "yyyyMMdd")
            Else
                StDatosLin.FechaFactura = Strings.Format(Dr("FechaFactura"), "yyyyMMdd")
            End If
            StDatosLin.ImpVencimiento = Strings.Right(Strings.Space(13) & Strings.Trim(Str(Dr("ImpVencimientoA") * 100)), 13)
            StDatosLin.FechaVencimiento = Strings.Format(Dr("FechaVencimiento"), "yyyyMMdd")
            StDatosLin.Espacios = Strings.Space(31)
            Dim DrNewLin As DataRow = DtFichero.NewRow
            DrNewLin("Linea") = StDatosLin.CIF & StDatosLin.DescBeneficiario & StDatosLin.DireccionPago & _
            StDatosLin.Espacio & StDatosLin.CodPostal & StDatosLin.PoblacionPago & StDatosLin.Telefono & _
            StDatosLin.Fax & StDatosLin.Espacios & StDatosLin.NFactura & StDatosLin.FechaFactura & _
            StDatosLin.ImpVencimiento & StDatosLin.FechaVencimiento & StDatosLin.Espacios
            DtFichero.Rows.Add(DrNewLin)
        Next
        Return DtFichero
    End Function

    <Task()> Public Shared Function GenerarFicheroConfirmingDeustche(ByVal DatosConf As DataFicheros, ByVal services As ServiceProvider) As DataTable
        '****************************************
        '*** Formato confirming Deustche Bank ***
        '****************************************
        Dim NumReg, NumReg010, NReg As Integer
        Dim ImpTotal As Double
        Dim StrRegistro As String = String.Empty
        Dim StDatosBanco As New EstDatosBancoFicheros
        Dim DtPagosAgrupados As DataTable
        Dim DtEmpresa As DataTable = New BE.DataEngine().Filter("tbDatosEmpresa", "*", "")
        Dim StrEmpReg As String = Strings.Left(Nz(DtEmpresa.Rows(0)("Cif")), 10)
        Dim StrNomEmpresa As String = (Strings.Left(Nz(DtEmpresa.Rows(0)("DescEmpresa"), " ") & Strings.Space(40), 40))
        Dim StrCIF As String = (Strings.Left(Nz(DtEmpresa.Rows(0)("Cif"), " ") & Strings.Space(10), 10))
        Dim DtFichero As New DataTable
        DtFichero.Columns.Add("Linea", GetType(String))

        'Cogemos los Pagos seleccionados
        Dim DtPagos As DataTable = New BE.DataEngine().Filter("frmPagoContGenerarFich", New GuidFilterItem("IDProcess", DatosConf.IDProcess))

        If DtPagos Is Nothing OrElse DtPagos.Rows.Count = 0 Then
            ApplicationService.GenerateError("No hay pagos seleccionados. No se generará el fichero.")
            Exit Function
        End If

        'Cogemos los datos del BancoPropio
        StDatosBanco.IDBanco = DatosConf.IDBancoPropio
        StDatosBanco = ProcessServer.ExecuteTask(Of EstDatosBancoFicheros, EstDatosBancoFicheros)(AddressOf DatosBanco, StDatosBanco, services)

        StDatosBanco.Entidad = StDatosBanco.Entidad
        StDatosBanco.Sucursal = StDatosBanco.Sucursal
        StDatosBanco.NCuenta = StDatosBanco.NCuenta
        StDatosBanco.DC = StDatosBanco.DC
        If StDatosBanco.CodClie = Strings.Space(6) Then
            ApplicationService.GenerateError("Debe introducir el código de cliente que somos para el banco en su ficha")
            Exit Function
        Else
            StDatosBanco.CodClie = Strings.Right(Space(6) & Strings.Trim(StDatosBanco.CodClie), 6)
        End If

        'Primer Registro de CABECERA Obligatorio
        NumReg += 1
        StrRegistro = "0356"
        StrRegistro &= StrCIF  'NIF propio
        StrRegistro &= "0000000"
        StrRegistro &= "O"
        StrRegistro &= "02"
        StrRegistro &= "X"
        StrRegistro &= "R"
        StrRegistro &= "001"
        StrRegistro &= Strings.Format(Today, "ddMMyy")
        StrRegistro &= Strings.Format(Today, "ddMMyy")
        StrRegistro &= StDatosBanco.Entidad
        StrRegistro &= StDatosBanco.Sucursal
        StrRegistro &= StDatosBanco.NCuenta & Strings.Space(4)
        StrRegistro &= StDatosBanco.DC & Strings.Space(4)
        StrRegistro &= Strings.Space(181)

        Dim DrNewCab As DataRow = DtFichero.NewRow
        DrNewCab("Linea") = StrRegistro
        DtFichero.Rows.Add(DrNewCab)

        'Registro Detalle
        NumReg += 1
        StrRegistro = "0356"
        StrRegistro &= StrCIF  'NIF propio
        StrRegistro &= Strings.Space(12)
        StrRegistro &= "002"
        StrRegistro &= StrNomEmpresa
        StrRegistro &= Strings.Space(181)

        Dim DrNewDet As DataRow = DtFichero.NewRow
        DrNewDet("Linea") = StrRegistro
        DtFichero.Rows.Add(DrNewDet)

        Dim StrPagoAgrupado As String = New Parametro().ObtenerPredeterminado("FRAPAGOAGR") & String.Empty
        For Each Dr As DataRow In DtPagos.Select
            'Primer reg.Detalle
            NumReg += 1
            ImpTotal += (Dr("ImpVencimientoA") * 100)
            NumReg010 += 1
            StrRegistro = "0657"
            StrRegistro &= StrCIF 'NIF propio
            StrRegistro &= Strings.Space(12)
            StrRegistro &= "010"
            StrRegistro &= Strings.Format(Dr("ImpVencimientoA") * 100, "000000000000")
            StrRegistro &= Strings.Space(24)
            StrRegistro &= "0"
            StrRegistro &= Strings.Format(Dr("FechaVencimiento"), "ddMMyy")
            StrRegistro &= Strings.Space(178)
            Dim DrNewDet1 As DataRow = DtFichero.NewRow
            DrNewDet1("Linea") = StrRegistro
            DtFichero.Rows.Add(DrNewDet1)

            'Segundo reg.Detalle
            NumReg += 1
            StrRegistro = "0657"
            StrRegistro &= StrCIF  'NIF propio
            StrRegistro &= Strings.Space(12)
            StrRegistro &= "011"
            StrRegistro &= (Strings.Left(Strings.Trim(Nz(Dr("DescBeneficiario"), " ")) & Space(40), 40))
            StrRegistro &= Strings.Space(181)
            Dim DrNewDet2 As DataRow = DtFichero.NewRow
            DrNewDet2("Linea") = StrRegistro
            DtFichero.Rows.Add(DrNewDet2)

            'Tercer reg.Detalle
            NumReg += 1
            StrRegistro = "0657"
            StrRegistro &= StrCIF  'NIF propio
            StrRegistro &= Strings.Space(12)
            StrRegistro &= "012"
            StrRegistro &= (Strings.Left(Strings.Trim(Nz(Dr("DireccionPago"), " ")) & Space(40), 40))
            StrRegistro &= Strings.Space(181)
            Dim DrNewDet3 As DataRow = DtFichero.NewRow
            DrNewDet3("Linea") = StrRegistro
            DtFichero.Rows.Add(DrNewDet3)
            'Cuarto reg.Detalle
            If Len(Trim(Dr("DireccionPago") & String.Empty)) > 40 Then
                NumReg += 1
                StrRegistro = "0657"
                StrRegistro &= StrCIF  'NIF propio
                StrRegistro &= Strings.Space(12)
                StrRegistro &= "013"
                StrRegistro &= (Strings.Mid(Trim(Nz(Dr("DireccionPago"), String.Empty)), 41, 40))
                StrRegistro &= Strings.Space(181)
                Dim DrNewDet4 As DataRow = DtFichero.NewRow
                DrNewDet4("Linea") = StrRegistro
                DtFichero.Rows.Add(DrNewDet4)
            End If
            'Quinto reg.Detalle
            NumReg += 1
            StrRegistro = "0657"
            StrRegistro &= StrCIF  'NIF propio
            StrRegistro &= Strings.Space(12)
            StrRegistro &= "014"
            StrRegistro &= Strings.Left(Strings.Trim(Nz(Dr("CodPostal"), String.Empty)) & Strings.Space(5), 5)
            StrRegistro &= (Strings.Left(Strings.Trim(Nz(Dr("PoblacionPago"), String.Empty)) & Strings.Space(31), 31))
            StrRegistro &= Strings.Space(185)
            Dim DrNewDet5 As DataRow = DtFichero.NewRow
            DrNewDet5("Linea") = StrRegistro
            DtFichero.Rows.Add(DrNewDet5)
            'Sexto reg.Detalle
            NumReg += 1
            StrRegistro = "0657"
            StrRegistro &= StrCIF  'NIF propio
            StrRegistro &= Strings.Space(12)
            StrRegistro &= "015"
            StrRegistro &= (Strings.Left(Strings.Trim(Nz(Dr("ProvinciaPago"), String.Empty)) & Strings.Space(40), 40))
            StrRegistro &= Strings.Space(181)
            Dim DrNewDet6 As DataRow = DtFichero.NewRow
            DrNewDet6("Linea") = StrRegistro
            DtFichero.Rows.Add(DrNewDet6)
            'Septimo reg.Detalle
            NumReg += 1
            StrRegistro = "0657"
            StrRegistro &= StrCIF  'NIF propio
            StrRegistro &= Strings.Space(12)
            StrRegistro &= "018"
            StrRegistro &= (Strings.Left(Strings.Trim(Nz(Dr("Cif"), String.Empty)) & Strings.Space(12), 12))
            StrRegistro &= Strings.Space(209)
            Dim DrNewDet7 As DataRow = DtFichero.NewRow
            DrNewDet7("Linea") = StrRegistro
            DtFichero.Rows.Add(DrNewDet7)

            'para Pagos Agrupados
            If Dr("NFactura") & String.Empty = StrPagoAgrupado & String.Empty Then
                Dim Sql As String = "SELECT tbPago.NFactura, tbPago.FechaVencimiento, " & _
                "tbFacturaCompraCabecera.FechaFactura,tbPago.ImpVencimientoA, " & _
                "tbPago.IDPagoAgrupado,tbFacturaCompraCabecera.NFactura " & _
                "FROM tbPago LEFT OUTER JOIN tbFacturaCompraCabecera ON " & _
                "tbPago.IDFactura = tbFacturaCompraCabecera.IDFactura " & _
                "where tbPago.IdPagoAgrupado= " & Dr("IDPago")
                DtPagosAgrupados = AdminData.Execute(Sql, ExecuteCommand.ExecuteReader)

                If Not DtPagosAgrupados Is Nothing AndAlso DtPagosAgrupados.Rows.Count > 0 Then
                    NReg = 19
                    For Each DrAgp As DataRow In DtPagosAgrupados.Select
                        StrRegistro = "0657"
                        StrRegistro &= StrCIF  'NIF propio
                        StrRegistro &= Strings.Space(12)
                        StrRegistro &= "0" & NReg
                        NumReg += 1
                        StrRegistro &= Strings.Left(DrAgp("NFactura") & Strings.Space(17), 17)
                        StrRegistro &= Strings.Right("00" & CDate(DrAgp("FechaFactura")).Day, 2) & "." & Strings.Right("00" & Month(DrAgp("FechaFactura")), 2) & "•" & Strings.Right(CStr(Year(DrAgp("FechaFactura"))), 2)
                        StrRegistro &= Strings.Format(System.Math.Abs(DrAgp("ImpVencimientoA") * 100), "0000000000")
                        StrRegistro &= IIf(DrAgp("ImpVencimientoA") < 0, "-", Strings.Space(1))
                        StrRegistro &= Strings.Space(15)
                        StrRegistro &= Strings.Space(15)
                        StrRegistro &= Strings.Space(45)
                        StrRegistro &= Strings.Space(2)
                        StrRegistro &= Strings.Space(40)
                        StrRegistro &= Strings.Space(68)
                        Dim DrNewAgp As DataRow = DtFichero.NewRow
                        DrNewAgp("Linea") = StrRegistro
                        DtFichero.Rows.Add(DrNewAgp)
                        NReg += 1
                    Next
                End If
            End If
        Next
        'Reg.Totales
        NumReg += 1
        StrRegistro = "0856"
        StrRegistro &= StrCIF   'NIF propio
        StrRegistro &= Strings.Space(15)
        StrRegistro &= Strings.Format(ImpTotal, "000000000000")
        StrRegistro &= Strings.Format(NumReg010, "00000000")
        StrRegistro &= Strings.Format(NumReg, "0000000000")
        StrRegistro &= Strings.Space(191)
        Dim DrNewTot As DataRow = DtFichero.NewRow
        DrNewTot("Linea") = StrRegistro
        DtFichero.Rows.Add(DrNewTot)
        Return DtFichero
    End Function
    '<Task()> Public Shared Function GenerarFicheroConfirming2(ByVal DatosConf As DataFicheros, ByVal services As ServiceProvider) As DataTable
    '    '****************************************
    '    '*** Formato confirming Bankinter ***IDClienteConfirming
    '    '****************************************
    '    Dim NumReg, NumReg010, NReg As Integer
    '    Dim ImpTotal As Double
    '    Dim StrRegistro As String = String.Empty
    '    Dim StDatosBanco As New EstDatosBancoFicheros
    '    Dim DtPagosAgrupados As DataTable
    '    Dim DtEmpresa As DataTable = New BE.DataEngine().Filter("tbDatosEmpresa", "*", "")
    '    Dim StrEmpReg As String = Strings.Left(Nz(DtEmpresa.Rows(0)("Cif")), 10)
    '    Dim StrNomEmpresa As String = (Strings.Left(Nz(DtEmpresa.Rows(0)("DescEmpresa"), " ") & Strings.Space(40), 40))
    '    Dim StrCIF As String = (Strings.Left(Nz(DtEmpresa.Rows(0)("Cif"), " ") & Strings.Space(10), 10))
    '    Dim DtFichero As New DataTable
    '    DtFichero.Columns.Add("Linea", GetType(String))

    '    'Cogemos los Pagos seleccionados
    '    Dim DtPagos As DataTable = New BE.DataEngine().Filter("frmPagoContGenerarFich", New GuidFilterItem("IDProcess", DatosConf.IDProcess))

    '    If DtPagos Is Nothing OrElse DtPagos.Rows.Count = 0 Then
    '        ApplicationService.GenerateError("No hay pagos seleccionados. No se generará el fichero.")
    '        Exit Function
    '    End If

    '    'Cogemos los datos del BancoPropio
    '    StDatosBanco.IDBanco = DatosConf.IDBancoPropio
    '    StDatosBanco = ProcessServer.ExecuteTask(Of EstDatosBancoFicheros, EstDatosBancoFicheros)(AddressOf DatosBanco, StDatosBanco, services)

    '    StDatosBanco.Entidad = StDatosBanco.Entidad
    '    StDatosBanco.Sucursal = StDatosBanco.Sucursal
    '    StDatosBanco.NCuenta = StDatosBanco.NCuenta
    '    StDatosBanco.DC = StDatosBanco.DC
    '    If StDatosBanco.CodClie = Strings.Space(6) Then
    '        ApplicationService.GenerateError("Debe introducir el código de cliente que somos para el banco en su ficha")
    '        Exit Function
    '        'Else
    '        '    StDatosBanco.CodClie = Strings.Right(Space(6) & Strings.Trim(StDatosBanco.CodClie), 6)
    '    End If

    '    'Primer Registro de CABECERA Obligatorio
    '    NumReg += 1
    '    StrRegistro = "0356"
    '    StrRegistro &= StrCIF  'NIF propio
    '    StrRegistro &= Strings.Space(12)
    '    StrRegistro &= "001"
    '    StrRegistro &= Strings.Format(Today, "yyMMdd")
    '    StrRegistro &= Strings.Format(Today, "yyMMdd")
    '    StrRegistro &= Left(StDatosBanco.CodClie, 4)  'StDatosBanco.Entidad
    '    StrRegistro &= Mid(StDatosBanco.CodClie, 5, 4) 'StDatosBanco.Sucursal
    '    StrRegistro &= Mid(StDatosBanco.CodClie, 11, 10) & Strings.Space(4) 'StDatosBanco.NCuenta & Strings.Space(4)
    '    StrRegistro &= Mid(StDatosBanco.CodClie, 9, 2) 'StDatosBanco.DC
    '    StrRegistro &= Strings.Space(181)

    '    Dim DrNewCab As DataRow = DtFichero.NewRow
    '    DrNewCab("Linea") = StrRegistro
    '    DtFichero.Rows.Add(DrNewCab)



    '    Dim strRegistroInicial

    '    Dim StrPagoAgrupado As String = New Parametro().ObtenerPredeterminado("FRAPAGOAGR") & String.Empty
    '    For Each Dr As DataRow In DtPagos.Select
    '        'Datps comunes
    '        strRegistroInicial = "06"
    '        If Dr("ChequeTalon") Then
    '            strRegistroInicial &= "57"
    '        Else
    '            strRegistroInicial &= "56"
    '        End If
    '        strRegistroInicial &= StrCIF 'NIF propio
    '        strRegistroInicial &= Dr("Cif") & Strings.Space(12 - Length(Dr("Cif")))


    '        'Primer reg.Detalle
    '        NumReg += 1
    '        ImpTotal += (Dr("ImpVencimientoA") * 100)
    '        NumReg010 += 1
    '        StrRegistro = strRegistroInicial
    '        StrRegistro &= "010"

    '        StrRegistro &= Strings.Format(Dr("ImpVencimientoA") * 100, "000000000000")
    '        If Dr("ChequeTalon") Then
    '            StrRegistro &= Strings.Space(24)
    '        Else
    '            StrRegistro &= Dr("IdBancoPago")
    '            StrRegistro &= Dr("SucursalPago")
    '            StrRegistro &= Dr("NCuentaPago") & Strings.Space(4)
    '            StrRegistro &= Dr("DCPago") & Strings.Space(4)
    '        End If
    '        'StrRegistro &= "0"
    '        'StrRegistro &= Strings.Format(Dr("FechaVencimiento"), "ddMMyy")
    '        'StrRegistro &= Strings.Space(178)
    '        Dim DrNewDet1 As DataRow = DtFichero.NewRow
    '        DrNewDet1("Linea") = StrRegistro
    '        DtFichero.Rows.Add(DrNewDet1)

    '        'Segundo reg.Detalle
    '        NumReg += 1
    '        StrRegistro = strRegistroInicial
    '        StrRegistro &= "011"
    '        StrRegistro &= (Strings.Left(Strings.Trim(Nz(Dr("DescBeneficiario"), " ")) & Space(40), 40))
    '        StrRegistro &= Strings.Space(181)
    '        Dim DrNewDet2 As DataRow = DtFichero.NewRow
    '        DrNewDet2("Linea") = StrRegistro
    '        DtFichero.Rows.Add(DrNewDet2)

    '        'Tercer reg.Detalle
    '        NumReg += 1
    '        StrRegistro = strRegistroInicial
    '        StrRegistro &= "012"
    '        StrRegistro &= (Strings.Left(Strings.Trim(Nz(Dr("DireccionPago"), " ")) & Space(40), 40))
    '        StrRegistro &= Strings.Space(181)
    '        Dim DrNewDet3 As DataRow = DtFichero.NewRow
    '        DrNewDet3("Linea") = StrRegistro
    '        DtFichero.Rows.Add(DrNewDet3)
    '        'Cuarto reg.Detalle
    '        If Len(Trim(Dr("DireccionPago") & String.Empty)) > 40 Then
    '            NumReg += 1
    '            StrRegistro = strRegistroInicial
    '            StrRegistro &= "013"
    '            StrRegistro &= (Strings.Mid(Trim(Nz(Dr("DireccionPago"), String.Empty)), 41, 40))
    '            StrRegistro &= Strings.Space(181)
    '            Dim DrNewDet4 As DataRow = DtFichero.NewRow
    '            DrNewDet4("Linea") = StrRegistro
    '            DtFichero.Rows.Add(DrNewDet4)
    '        End If
    '        'Quinto reg.Detalle
    '        NumReg += 1
    '        StrRegistro = strRegistroInicial
    '        StrRegistro &= "014"
    '        StrRegistro &= Strings.Left(Strings.Trim(Nz(Dr("CodPostal"), String.Empty)) & Strings.Space(5), 5)
    '        StrRegistro &= (Strings.Left(Strings.Trim(Nz(Dr("PoblacionPago"), String.Empty)) & Strings.Space(31), 31))
    '        StrRegistro &= Strings.Space(185)
    '        Dim DrNewDet5 As DataRow = DtFichero.NewRow
    '        DrNewDet5("Linea") = StrRegistro
    '        DtFichero.Rows.Add(DrNewDet5)
    '        'Sexto reg.Detalle
    '        NumReg += 1
    '        StrRegistro = strRegistroInicial
    '        StrRegistro &= "015"
    '        StrRegistro &= (Strings.Left(Strings.Trim(Nz(Dr("ProvinciaPago"), String.Empty)) & Strings.Space(40), 40))
    '        StrRegistro &= Strings.Space(181)
    '        Dim DrNewDet6 As DataRow = DtFichero.NewRow
    '        DrNewDet6("Linea") = StrRegistro
    '        DtFichero.Rows.Add(DrNewDet6)
    '        'Septimo reg.Detalle
    '        NumReg += 1
    '        StrRegistro = strRegistroInicial
    '        StrRegistro &= "018"
    '        StrRegistro &= Strings.Format(Dr("FechaVencimiento"), "yyMMdd")
    '        StrRegistro &= (Strings.Left(Strings.Trim(Nz(Dr("SuFactura"), String.Empty)) & Strings.Space(16), 16))
    '        StrRegistro &= Strings.Space(14)
    '        Dim DrNewDet7 As DataRow = DtFichero.NewRow
    '        DrNewDet7("Linea") = StrRegistro
    '        DtFichero.Rows.Add(DrNewDet7)

    '        'para Pagos Agrupados
    '        If Dr("NFactura") = StrPagoAgrupado Then
    '            Dim Sql As String = "SELECT tbPago.NFactura, tbPago.FechaVencimiento, " & _
    '            "tbFacturaCompraCabecera.FechaFactura,tbPago.ImpVencimientoA, " & _
    '            "tbPago.IDPagoAgrupado,tbFacturaCompraCabecera.NFactura " & _
    '            "FROM tbPago LEFT OUTER JOIN tbFacturaCompraCabecera ON " & _
    '            "tbPago.IDFactura = tbFacturaCompraCabecera.IDFactura " & _
    '            "where tbPago.IdPagoAgrupado= " & Dr("IDPago")
    '            DtPagosAgrupados = AdminData.Execute(Sql, ExecuteCommand.ExecuteReader)

    '            If Not DtPagosAgrupados Is Nothing AndAlso DtPagosAgrupados.Rows.Count > 0 Then
    '                NReg = 19
    '                For Each DrAgp As DataRow In DtPagosAgrupados.Select
    '                    StrRegistro = "06"
    '                    StrRegistro = strRegistroInicial
    '                    StrRegistro &= "0" & NReg
    '                    NumReg += 1
    '                    StrRegistro &= Strings.Left(DrAgp("NFactura") & Strings.Space(17), 17)
    '                    StrRegistro &= Strings.Right("00" & CDate(DrAgp("FechaFactura")).Day, 2) & "." & Strings.Right("00" & Month(DrAgp("FechaFactura")), 2) & "•" & Strings.Right(CStr(Year(DrAgp("FechaFactura"))), 2)
    '                    StrRegistro &= Strings.Format(System.Math.Abs(DrAgp("ImpVencimientoA") * 100), "0000000000")
    '                    StrRegistro &= IIf(DrAgp("ImpVencimientoA") < 0, "-", Strings.Space(1))
    '                    StrRegistro &= Strings.Space(15)
    '                    StrRegistro &= Strings.Space(15)
    '                    StrRegistro &= Strings.Space(45)
    '                    StrRegistro &= Strings.Space(2)
    '                    StrRegistro &= Strings.Space(40)
    '                    StrRegistro &= Strings.Space(68)
    '                    Dim DrNewAgp As DataRow = DtFichero.NewRow
    '                    DrNewAgp("Linea") = StrRegistro
    '                    DtFichero.Rows.Add(DrNewAgp)
    '                    NReg += 1
    '                Next
    '            End If
    '        End If
    '    Next
    '    'Reg.Totales
    '    NumReg += 1
    '    StrRegistro = "0856"
    '    StrRegistro &= StrCIF   'NIF propio
    '    StrRegistro &= Strings.Space(15)
    '    StrRegistro &= Strings.Format(ImpTotal, "000000000000")
    '    StrRegistro &= Strings.Format(NumReg010, "00000000")
    '    StrRegistro &= Strings.Format(NumReg, "0000000000")
    '    StrRegistro &= Strings.Space(191)
    '    Dim DrNewTot As DataRow = DtFichero.NewRow
    '    DrNewTot("Linea") = StrRegistro
    '    DtFichero.Rows.Add(DrNewTot)
    '    Return DtFichero
    'End Function
    <Task()> Public Shared Function GenerarFicheroConfirming2(ByVal DatosConf As Expertis.Business.Negocio.GenerarFicheros.DataFicheros, ByVal services As ServiceProvider) As DataTable
        '****************************************
        '*** Formato confirming Bankinter ***
        '****************************************
        Dim NumReg, NumReg010, NReg As Integer
        Dim ImpTotal As Double
        Dim StrRegistro As String = String.Empty
        Dim StDatosBanco As New Expertis.Business.Negocio.GenerarFicheros.EstDatosBancoFicheros
        Dim DtPagosAgrupados As DataTable
        Dim DtEmpresa As DataTable = New BE.DataEngine().Filter("tbDatosEmpresa", "*", "")
        Dim StrEmpReg As String = Strings.Left(Nz(DtEmpresa.Rows(0)("Cif")), 10)
        Dim StrNomEmpresa As String = (Strings.Left(Nz(DtEmpresa.Rows(0)("DescEmpresa"), " ") & Strings.Space(40), 40))
        Dim StrCIF As String = (Strings.Left(Nz(DtEmpresa.Rows(0)("Cif"), " ") & Strings.Space(10), 10))
        Dim DtFichero As New DataTable
        DtFichero.Columns.Add("Linea", GetType(String))

        'Cogemos los Pagos seleccionados
        Dim DtPagos As DataTable = New BE.DataEngine().Filter("frmPagoContGenerarFich", New GuidFilterItem("IDProcess", DatosConf.IDProcess))

        If DtPagos Is Nothing OrElse DtPagos.Rows.Count = 0 Then
            ApplicationService.GenerateError("No hay pagos seleccionados. No se generará el fichero.")
            Exit Function
        End If

        'Cogemos los datos del BancoPropio
        StDatosBanco.IDBanco = DatosConf.IDBancoPropio
        StDatosBanco = ProcessServer.ExecuteTask(Of Expertis.Business.Negocio.GenerarFicheros.EstDatosBancoFicheros, Expertis.Business.Negocio.GenerarFicheros.EstDatosBancoFicheros)(AddressOf Expertis.Business.Negocio.GenerarFicheros.DatosBanco, StDatosBanco, services)
        If Length(StDatosBanco.CodClie) = 20 Then
            StDatosBanco.Entidad = Left(StDatosBanco.CodClie, 4)
            StDatosBanco.Sucursal = Mid(StDatosBanco.CodClie, 5, 4)
            StDatosBanco.NCuenta = Right(StDatosBanco.CodClie, 10)
            StDatosBanco.DC = Mid(StDatosBanco.CodClie, 9, 2)
        Else
            StDatosBanco.Entidad = StDatosBanco.Entidad
            StDatosBanco.Sucursal = StDatosBanco.Sucursal
            StDatosBanco.NCuenta = StDatosBanco.NCuenta
            StDatosBanco.DC = StDatosBanco.DC
        End If
        If StDatosBanco.CodClie = Strings.Space(6) Then
            ApplicationService.GenerateError("Debe introducir el código de cliente que somos para el banco en su ficha")
            Exit Function
        Else
            StDatosBanco.CodClie = Strings.Right(Space(6) & Strings.Trim(StDatosBanco.CodClie), 6)
        End If

        'Primer Registro de CABECERA Obligatorio
        NumReg += 1
        StrRegistro = "0360"
        StrRegistro &= StrCIF  'NIF propio
        StrRegistro &= Strings.Space(12)
        StrRegistro &= "001"
        StrRegistro &= Strings.Format(Today, "yyMMdd")
        StrRegistro &= Strings.Format(Today, "yyMMdd")
        StrRegistro &= StDatosBanco.Entidad
        StrRegistro &= StDatosBanco.Sucursal
        StrRegistro &= StDatosBanco.NCuenta & Strings.Space(4)
        StrRegistro &= StDatosBanco.DC & Strings.Space(4)
        StrRegistro &= Strings.Space(3)

        Dim DrNewCab As DataRow = DtFichero.NewRow
        DrNewCab("Linea") = StrRegistro
        DtFichero.Rows.Add(DrNewCab)



        Dim strRegistroInicial

        Dim StrPagoAgrupado As String = New Parametro().ObtenerPredeterminado("FRAPAGOAGR") & String.Empty
        For Each Dr As DataRow In DtPagos.Select
            'Datps comunes
            strRegistroInicial = "06"
            If Dr("ChequeTalon") Then
                strRegistroInicial &= "57"
            Else
                strRegistroInicial &= "56"
            End If
            strRegistroInicial &= StrCIF 'NIF propio
            strRegistroInicial &= Dr("Cif") & Strings.Space(12 - Length(Dr("Cif")))


            'Primer reg.Detalle
            NumReg += 1
            ImpTotal += (Math.Abs(Dr("ImpVencimientoA")) * 100)
            NumReg010 += 1
            StrRegistro = strRegistroInicial
            StrRegistro &= "010"

            StrRegistro &= Strings.Format(Math.Abs(Dr("ImpVencimientoA")) * 100, "000000000000")
            'If Dr("ChequeTalon") Then
            StrRegistro &= Strings.Space(19)
            If Dr("ImpVencimientoA") < 0 Then
                StrRegistro &= "-"
            Else
                StrRegistro &= " "
            End If
            'Else
            'StrRegistro &= Dr("IdBancoPago")
            'StrRegistro &= Dr("SucursalPago")
            'StrRegistro &= Dr("NCuentaPago") & Strings.Space(4)
            'StrRegistro &= Dr("DCPago") & Strings.Space(4)
            'End If
            'StrRegistro &= "0"
            'StrRegistro &= Strings.Format(Dr("FechaVencimiento"), "ddMMyy")
            'StrRegistro &= Strings.Space(178)

            Dim DrNewDet1 As DataRow = DtFichero.NewRow

            If Length(StrRegistro) < 72 Then
                StrRegistro &= Strings.Space(72 - Length(StrRegistro))
            End If

            DrNewDet1("Linea") = StrRegistro
            DtFichero.Rows.Add(DrNewDet1)

            'Segundo reg.Detalle
            NumReg += 1
            StrRegistro = strRegistroInicial
            StrRegistro &= "011"
            StrRegistro &= (Strings.Left(Strings.Trim(Nz(Dr("DescBeneficiario"), " ")) & Space(36), 36))
            ' StrRegistro &= Strings.Space(3)
            Dim DrNewDet2 As DataRow = DtFichero.NewRow

            If Length(StrRegistro) < 72 Then
                StrRegistro &= Strings.Space(72 - Length(StrRegistro))
            End If

            DrNewDet2("Linea") = StrRegistro
            DtFichero.Rows.Add(DrNewDet2)

            'Tercer reg.Detalle
            NumReg += 1
            StrRegistro = strRegistroInicial
            StrRegistro &= "012"
            StrRegistro &= (Strings.Left(Strings.Trim(Nz(Dr("DireccionPago"), " ")) & Space(36), 36))
            ' StrRegistro &= Strings.Space(3)
            Dim DrNewDet3 As DataRow = DtFichero.NewRow

            If Length(StrRegistro) < 72 Then
                StrRegistro &= Strings.Space(72 - Length(StrRegistro))
            End If

            DrNewDet3("Linea") = StrRegistro
            DtFichero.Rows.Add(DrNewDet3)
            'Cuarto reg.Detalle
            'If Len(Trim(Dr("DireccionPago") & String.Empty)) > 40 Then
            '    NumReg += 1
            '    StrRegistro = strRegistroInicial
            '    StrRegistro &= "013"
            '    StrRegistro &= (Strings.Mid(Trim(Nz(Dr("DireccionPago"), String.Empty)), 41, 40))
            '    StrRegistro &= Strings.Space(3)
            '    Dim DrNewDet4 As DataRow = DtFichero.NewRow
            '    DrNewDet4("Linea") = StrRegistro
            '    DtFichero.Rows.Add(DrNewDet4)
            'End If
            'Quinto reg.Detalle
            NumReg += 1
            StrRegistro = strRegistroInicial
            StrRegistro &= "014"
            StrRegistro &= Strings.Left(Strings.Trim(Nz(Dr("CodPostal"), String.Empty)) & Strings.Space(5), 5)
            StrRegistro &= (Strings.Left(Strings.Trim(Nz(Dr("PoblacionPago"), String.Empty)) & Strings.Space(32), 32))
            '  StrRegistro &= Strings.Space(7)
            Dim DrNewDet5 As DataRow = DtFichero.NewRow

            If Length(StrRegistro) < 72 Then
                StrRegistro &= Strings.Space(72 - Length(StrRegistro))
            End If

            DrNewDet5("Linea") = StrRegistro
            DtFichero.Rows.Add(DrNewDet5)
            'Sexto reg.Detalle
            'NumReg += 1
            'StrRegistro = strRegistroInicial
            'StrRegistro &= "015"
            'StrRegistro &= (Strings.Left(Strings.Trim(Nz(Dr("ProvinciaPago"), String.Empty)) & Strings.Space(40), 40))
            'StrRegistro &= Strings.Space(3)
            'Dim DrNewDet6 As DataRow = DtFichero.NewRow

            'If Length(StrRegistro) < 72 Then
            '    StrRegistro &= Strings.Space(72 - Length(StrRegistro))
            'End If

            'DrNewDet6("Linea") = StrRegistro
            'DtFichero.Rows.Add(DrNewDet6)

            '173 reg.Detalle
            NumReg += 1
            StrRegistro = strRegistroInicial
            StrRegistro &= "173"
            If Dr("ChequeTalon") Then
                StrRegistro &= Strings.Space(34)
            Else
                StrRegistro &= (Strings.Left(Strings.Trim(Nz(Dr("CodigoIBAN"), " ")) & Space(34), 34))
                StrRegistro &= Strings.Space(1)
                StrRegistro &= (Left(Strings.Trim(Nz(Dr("CodigoIBAN"), " ")) & Space(2), 2))
                '     StrRegistro &= (Left(Strings.Trim(Nz(Dr("Swift"), " ")) & Space(11), 11))
            End If


            Dim DrNewDet173 As DataRow = DtFichero.NewRow

            If Length(StrRegistro) < 72 Then
                StrRegistro &= Strings.Space(72 - Length(StrRegistro))
            End If

            DrNewDet173("Linea") = StrRegistro
            DtFichero.Rows.Add(DrNewDet173)

            '174 reg.Detalle
            NumReg += 1
            StrRegistro = strRegistroInicial
            StrRegistro &= "174"
            If Dr("ChequeTalon") Then
                StrRegistro &= Strings.Space(11)
            Else

                'CAMBIOS RELLENAR CON CEROS
                'StrRegistro &= (Left(Strings.Trim(Nz(Dr("Swift"), " ")) & Space(11), 11))
                Dim StrSwift As String = Strings.Trim(Nz(Dr("Swift"), String.Empty))
                StrRegistro &= (Strings.Left(StrSwift & New String("0", 11 - StrSwift.Length), 11))

            End If

            Dim DrNewDet174 As DataRow = DtFichero.NewRow

            If Length(StrRegistro) < 72 Then
                StrRegistro &= Strings.Space(72 - Length(StrRegistro))
            End If

            DrNewDet174("Linea") = StrRegistro
            DtFichero.Rows.Add(DrNewDet174)

            '175 reg.Detalle
            NumReg += 1
            StrRegistro = strRegistroInicial
            StrRegistro &= "175"

            Dim DrNewDet175 As DataRow = DtFichero.NewRow

            StrRegistro &= "E"
            If Length(StrRegistro) < 72 Then
                StrRegistro &= Strings.Space(72 - Length(StrRegistro))
            End If

            DrNewDet175("Linea") = StrRegistro
            DtFichero.Rows.Add(DrNewDet175)


            '182 reg.Detalle
            NumReg += 1
            StrRegistro = strRegistroInicial
            StrRegistro &= "182"

            Dim DrNewDet182 As DataRow = DtFichero.NewRow

            StrRegistro &= Dr("Cif") & Strings.Space(12 - Length(Dr("Cif")))
            If Length(StrRegistro) < 72 Then
                StrRegistro &= Strings.Space(72 - Length(StrRegistro))
            End If

            DrNewDet182("Linea") = StrRegistro
            DtFichero.Rows.Add(DrNewDet182)

            'Septimo reg.Detalle

            NumReg += 1
            StrRegistro = strRegistroInicial
            StrRegistro &= "018"
            StrRegistro &= Strings.Format(Dr("FechaVencimiento"), "yyMMdd")
            ' StrRegistro &= (Strings.Left(Strings.Trim(Nz(Dr("SuFactura"), String.Empty)) & Strings.Space(16), 16))


            'para Pagos Agrupados
            Dim SuFactura As String
            If Dr("NFactura") & String.Empty = StrPagoAgrupado & String.Empty Then
                Dim Sql As String = "SELECT tbFacturaCompraCabecera.SuFactura, tbPago.NFactura, tbPago.FechaVencimiento, " & _
                "tbFacturaCompraCabecera.FechaFactura,tbPago.ImpVencimientoA, " & _
                "tbPago.IDPagoAgrupado,tbFacturaCompraCabecera.NFactura " & _
                "FROM tbPago LEFT OUTER JOIN tbFacturaCompraCabecera ON " & _
                "tbPago.IDFactura = tbFacturaCompraCabecera.IDFactura " & _
                "where tbPago.IdPagoAgrupado= " & Dr("IDPago")
                DtPagosAgrupados = AdminData.Execute(Sql, ExecuteCommand.ExecuteReader)

                If Not DtPagosAgrupados Is Nothing AndAlso DtPagosAgrupados.Rows.Count > 0 Then

                    For Each DrAgp As DataRow In DtPagosAgrupados.Select
                        SuFactura = SuFactura & DrAgp("SuFactura") & ","
                        If Length(SuFactura) > 16 Then
                            SuFactura = Left(SuFactura, 16)
                            Exit For
                        End If
                    Next
                End If
                StrRegistro &= (Strings.Left(Strings.Trim(Nz(SuFactura, String.Empty)) & Strings.Space(14), 14))
                SuFactura = String.Empty
            Else
                StrRegistro &= (Strings.Left(Strings.Trim(Nz(Dr("SuFactura"), String.Empty)) & Strings.Space(14), 14))
            End If

            ' StrRegistro &= Strings.Space(14)
            If Length(StrRegistro) < 72 Then
                StrRegistro &= Strings.Space(72 - Length(StrRegistro))
            End If
            Dim DrNewDet7 As DataRow = DtFichero.NewRow
            DrNewDet7("Linea") = StrRegistro
            DtFichero.Rows.Add(DrNewDet7)

            'CAMBIOS AÑADIR EL 019
            '019 reg.Detalle
            NumReg += 1
            StrRegistro = strRegistroInicial
            StrRegistro &= "019"

            Dim DrNewDet019 As DataRow = DtFichero.NewRow

            If Length(StrRegistro) < 72 Then
                StrRegistro &= Strings.Space(72 - Length(StrRegistro))
            End If

            DrNewDet019("Linea") = StrRegistro
            DtFichero.Rows.Add(DrNewDet019)

        Next


        'Reg.Totales
        NumReg += 1
        StrRegistro = "0860"
        StrRegistro &= StrCIF   'NIF propio
        StrRegistro &= Strings.Space(15)
        StrRegistro &= Strings.Format(ImpTotal, "000000000000")
        StrRegistro &= Strings.Format(NumReg010, "00000000")
        StrRegistro &= Strings.Format(NumReg, "0000000000")
        StrRegistro &= Strings.Space(13)
        Dim DrNewTot As DataRow = DtFichero.NewRow
        DrNewTot("Linea") = StrRegistro
        DtFichero.Rows.Add(DrNewTot)
        Return DtFichero
    End Function
    <Task()> Public Shared Function GenerarFicheroConfirmingBankia(ByVal DatosConf As Expertis.Business.Negocio.GenerarFicheros.DataFicheros, ByVal services As ServiceProvider) As DataTable
        '****************************************
        '*** Formato confirming Bankia ***
        '****************************************
        Dim NumReg, NumRegDivisa, NumReg010, NumReg010Divisa, NReg, NRegDivisa As Integer
        Dim ImpTotal, ImpTotalDivisa As Double
        Dim StrRegistro As String = String.Empty
        Dim StDatosBanco As New Expertis.Business.Negocio.GenerarFicheros.EstDatosBancoFicheros
        Dim DtPagosAgrupados As DataTable
        Dim DtEmpresa As DataTable = New BE.DataEngine().Filter("tbDatosEmpresa", "*", "")
        Dim StrEmpReg As String = Strings.Left(Nz(DtEmpresa.Rows(0)("Cif")), 9)
        Dim StrNomEmpresa As String = (Strings.Left(Nz(DtEmpresa.Rows(0)("DescEmpresa"), " ") & Strings.Space(36), 36))
        Dim StrDireccion As String = (Strings.Left(Nz(DtEmpresa.Rows(0)("Direccion"), " ") & Strings.Space(36), 36))
        Dim StrPoblacion As String = (Strings.Left(Nz(DtEmpresa.Rows(0)("POBLACION"), " ") & Strings.Space(36), 36))
        Dim StrCIF As String = (Strings.Left(Nz(DtEmpresa.Rows(0)("Cif"), " ") & Strings.Space(9), 9))
        Dim DtFichero As New DataTable
        DtFichero.Columns.Add("Linea", GetType(String))

        'Cogemos los Pagos seleccionados
        Dim Monedas As New MonedaCache

        Dim f As New Filter
        f.Add(New GuidFilterItem("IDProcess", DatosConf.IDProcess))


        Dim DtPagos As DataTable = New BE.DataEngine().Filter("frmPagoContGenerarFichBankia", f)

        If DtPagos Is Nothing OrElse DtPagos.Rows.Count = 0 Then
            ApplicationService.GenerateError("No hay pagos seleccionados. No se generará el fichero.")
            Exit Function
        End If

        'Cogemos los datos del BancoPropio
        StDatosBanco.IDBanco = DatosConf.IDBancoPropio
        StDatosBanco = ProcessServer.ExecuteTask(Of Expertis.Business.Negocio.GenerarFicheros.EstDatosBancoFicheros, Expertis.Business.Negocio.GenerarFicheros.EstDatosBancoFicheros)(AddressOf Expertis.Business.Negocio.GenerarFicheros.DatosBanco, StDatosBanco, services)

        StDatosBanco.Entidad = StDatosBanco.Entidad
        StDatosBanco.Sucursal = StDatosBanco.Sucursal
        StDatosBanco.NCuenta = StDatosBanco.NCuenta
        StDatosBanco.DC = StDatosBanco.DC
        If StDatosBanco.CodClie = Strings.Space(6) Then
            ApplicationService.GenerateError("Debe introducir el código de cliente que somos para el banco en su ficha")
            Exit Function
        Else
            StDatosBanco.CodClie = Strings.Right(Space(6) & Strings.Trim(StDatosBanco.CodClie), 6)
        End If

        'Tipo 1º de CABECERA Obligatorio
        Dim strRegistroInicial


        NumReg += 1

        strRegistroInicial = "0362"
        strRegistroInicial &= StrCIF  'NIF propio
        strRegistroInicial &= "000"
        StrRegistro = strRegistroInicial & (Strings.Left(Nz(DtPagos.Rows(0)("ContratoOIE")) & Strings.Space(11), 11))   'TODO NCONTRATO CONFIRMING 40000063700 
        StrRegistro &= Strings.Space(1)
        StrRegistro &= "001"
        StrRegistro &= Strings.Format(Today, "ddMMyy")
        StrRegistro &= Strings.Format(Today, "ddMMyy")
        StrRegistro &= StDatosBanco.Entidad
        StrRegistro &= StDatosBanco.Sucursal
        StrRegistro &= StDatosBanco.DC
        StrRegistro &= StDatosBanco.NCuenta
        StrRegistro &= "0"
        StrRegistro &= Strings.Space(8)

        Dim DrNewCab As DataRow = DtFichero.NewRow
        DrNewCab("Linea") = StrRegistro
        DtFichero.Rows.Add(DrNewCab)

        'Tipo 2º de CABECERA Obligatorio

        StrRegistro = strRegistroInicial & (Strings.Left(Nz(DtPagos.Rows(0)("IDRemesa")) & Strings.Space(12), 12))
        StrRegistro &= "002"
        StrRegistro &= StrNomEmpresa
        StrRegistro &= Strings.Space(5)
        Dim DrNewCab2 As DataRow = DtFichero.NewRow
        DrNewCab2("Linea") = StrRegistro
        DtFichero.Rows.Add(DrNewCab2)

        'Tipo 3º de CABECERA Obligatorio
        StrRegistro = strRegistroInicial & Strings.Space(12)
        StrRegistro &= "003"
        StrRegistro &= StrDireccion
        StrRegistro &= Strings.Space(5)
        Dim DrNewCab3 As DataRow = DtFichero.NewRow
        DrNewCab3("Linea") = StrRegistro
        DtFichero.Rows.Add(DrNewCab3)

        'Tipo 4º de CABECERA Obligatorio
        StrRegistro = strRegistroInicial & Strings.Space(12)
        StrRegistro &= "004"
        StrRegistro &= StrPoblacion
        StrRegistro &= Strings.Space(5)
        Dim DrNewCab4 As DataRow = DtFichero.NewRow
        DrNewCab4("Linea") = StrRegistro
        DtFichero.Rows.Add(DrNewCab4)

        'Tipo 5º de CABECERA Obligatorio
        StrRegistro = strRegistroInicial & Strings.Space(12)
        StrRegistro &= "005"
        StrRegistro &= (Strings.Left(Nz(DtPagos.Rows(0)("LineaComercioExterior")) & Strings.Space(36), 36))    'TODO '" Numero de línea de confirming INTERNACIONAL"
        StrRegistro &= Strings.Space(5)
        Dim DrNewCab5 As DataRow = DtFichero.NewRow
        DrNewCab5("Linea") = StrRegistro
        DtFichero.Rows.Add(DrNewCab5)


        ' CABECERA EN EUROS Obligatorio



        'Cabecera de euros
        StrRegistro = "0460"
        StrRegistro &= StrCIF
        StrRegistro &= "000"
        StrRegistro &= Strings.Space(72 - Length(StrRegistro))
        Dim DrNewCabEuro As DataRow = DtFichero.NewRow
        DrNewCabEuro("Linea") = StrRegistro
        DtFichero.Rows.Add(DrNewCabEuro)

        strRegistroInicial = "0660"
        strRegistroInicial &= StrCIF  'NIF propio
        strRegistroInicial &= "000"


        Dim StrPagoAgrupado As String = New Parametro().ObtenerPredeterminado("FRAPAGOAGR") & String.Empty
        Dim DrPagosMonedaA() As DataRow = DtPagos.Select("IDMoneda = '" & Monedas.MonedaA.ID & "'")
        If DrPagosMonedaA.Length > 0 Then
            For Each Dr As DataRow In DrPagosMonedaA


                'Tipo 1 Detalle euros
                NumReg += 1
                NumRegDivisa += 1
                ImpTotal += (Dr("ImpVencimientoA") * 100)
                ImpTotalDivisa += (Dr("ImpVencimientoA") * 100)
                NumReg010 += 1
                NumReg010Divisa += 1

                StrRegistro = strRegistroInicial

                StrRegistro &= Dr("Cif") & Strings.Space(12 - Length(Dr("Cif")))
                StrRegistro &= "033"
                StrRegistro &= Left(Dr("CodigoIban"), 2)
                StrRegistro &= Mid(Dr("CodigoIban"), 3, 2)
                StrRegistro &= Mid(Dr("CodigoIban"), 5) & Strings.Space(30 - Length(Mid(Dr("CodigoIban"), 5)))
                StrRegistro &= Dr("Concepto")
                StrRegistro &= Strings.Space(6)

                Dim DrNewDet1 As DataRow = DtFichero.NewRow

                DrNewDet1("Linea") = StrRegistro
                DtFichero.Rows.Add(DrNewDet1)

                'Tipo 2 Detalle euros
                NumReg += 1
                StrRegistro = strRegistroInicial
                StrRegistro &= Dr("Cif") & Strings.Space(12 - Length(Dr("Cif")))
                StrRegistro &= "034"
                StrRegistro &= Strings.Format(Dr("ImpVencimientoA") * 100, "000000000000")
                StrRegistro &= "3"
                StrRegistro &= Left(Dr("CodigoIso"), 2)
                StrRegistro &= Strings.Space(6)
                StrRegistro &= (Strings.Left(Strings.Trim(Nz(Dr("swift"), " ")) & Space(11), 11))
                StrRegistro &= Strings.Space(9)

                Dim DrNewDet2 As DataRow = DtFichero.NewRow
                DrNewDet2("Linea") = StrRegistro
                DtFichero.Rows.Add(DrNewDet2)

                'Tipo 3 Detalle euros
                NumReg += 1
                StrRegistro = strRegistroInicial
                StrRegistro &= Dr("Cif") & Strings.Space(12 - Length(Dr("Cif")))
                StrRegistro &= "035"
                StrRegistro &= (Strings.Left(Strings.Trim(Nz(Dr("DescBeneficiario"), " ")) & Space(36), 36))
                StrRegistro &= Strings.Space(5)
                Dim DrNewDet3 As DataRow = DtFichero.NewRow

                If Length(StrRegistro) < 72 Then
                    StrRegistro &= Strings.Space(72 - Length(StrRegistro))
                End If

                DrNewDet3("Linea") = StrRegistro
                DtFichero.Rows.Add(DrNewDet3)

                'Tipo 4 Detalle euros
                NumReg += 1
                StrRegistro = strRegistroInicial
                StrRegistro &= Dr("Cif") & Strings.Space(12 - Length(Dr("Cif")))
                StrRegistro &= "036"
                StrRegistro &= (Strings.Left(Strings.Trim(Nz(Dr("DireccionPago"), " ")) & Space(36), 36))
                StrRegistro &= Strings.Space(5)
                Dim DrNewDet4 As DataRow = DtFichero.NewRow

                DrNewDet4("Linea") = StrRegistro
                DtFichero.Rows.Add(DrNewDet4)

                'Tipo 6 Detalle euros
                NumReg += 1
                StrRegistro = strRegistroInicial
                StrRegistro &= Dr("Cif") & Strings.Space(12 - Length(Dr("Cif")))
                StrRegistro &= "038"
                StrRegistro &= (Strings.Left(Strings.Trim(Nz(Dr("CodPostal") & Dr("PoblacionPago"), " ")) & Space(36), 36))
                StrRegistro &= Strings.Space(5)
                Dim DrNewDet6 As DataRow = DtFichero.NewRow

                DrNewDet6("Linea") = StrRegistro
                DtFichero.Rows.Add(DrNewDet6)

                'Tipo 7 Detalle euros
                NumReg += 1
                StrRegistro = strRegistroInicial
                StrRegistro &= Dr("Cif") & Strings.Space(12 - Length(Dr("Cif")))
                StrRegistro &= "039"
                StrRegistro &= (Strings.Left(Strings.Trim(Nz(Dr("DescPais"), " ")) & Space(36), 36))
                StrRegistro &= Strings.Space(5)
                Dim DrNewDet7 As DataRow = DtFichero.NewRow

                DrNewDet7("Linea") = StrRegistro
                DtFichero.Rows.Add(DrNewDet7)

                'Tipo 10 Detalle euros
                NumReg += 1
                StrRegistro = strRegistroInicial
                StrRegistro &= Strings.Format(Dr("FechaVencimiento"), "ddMMyyyy")
                StrRegistro &= Strings.Space(4)
                StrRegistro &= "042"
                StrRegistro &= Strings.Format(Dr("FechaFactura"), "ddMMyyyy")
                StrRegistro &= Strings.Space(1)
                StrRegistro &= Dr("Cif") & Strings.Space(12 - Length(Dr("Cif")))
                StrRegistro &= Strings.Space(1)

                'para Pagos Agrupados
                Dim SuFactura As String
                If Dr("NFactura") = StrPagoAgrupado Then
                    Dim Sql As String = "SELECT tbFacturaCompraCabecera.SuFactura, tbPago.NFactura, tbPago.FechaVencimiento, " & _
                    "tbFacturaCompraCabecera.FechaFactura,tbPago.ImpVencimientoA, " & _
                    "tbPago.IDPagoAgrupado,tbFacturaCompraCabecera.NFactura " & _
                    "FROM tbPago LEFT OUTER JOIN tbFacturaCompraCabecera ON " & _
                    "tbPago.IDFactura = tbFacturaCompraCabecera.IDFactura " & _
                    "where tbPago.IdPagoAgrupado= " & Dr("IDPago")
                    DtPagosAgrupados = AdminData.Execute(Sql, ExecuteCommand.ExecuteReader)

                    If Not DtPagosAgrupados Is Nothing AndAlso DtPagosAgrupados.Rows.Count > 0 Then

                        For Each DrAgp As DataRow In DtPagosAgrupados.Select
                            SuFactura = SuFactura & DrAgp("SuFactura") & ","
                            If Length(SuFactura) > 16 Then
                                SuFactura = Left(SuFactura, 18)
                                Exit For
                            End If
                        Next
                    End If
                    StrRegistro &= (Strings.Left(Strings.Trim(Nz(SuFactura, String.Empty)) & Strings.Space(18), 18))
                    SuFactura = String.Empty
                Else
                    StrRegistro &= (Strings.Left(Strings.Trim(Nz(Dr("SuFactura"), String.Empty)) & Strings.Space(18), 18))
                End If
                StrRegistro &= Strings.Space(1)
                Dim DrNewDet10 As DataRow = DtFichero.NewRow

                DrNewDet10("Linea") = StrRegistro
                DtFichero.Rows.Add(DrNewDet10)

                'Tipo 13 Detalle euros
                NumReg += 1
                StrRegistro = strRegistroInicial
                StrRegistro &= Dr("Cif") & Strings.Space(12 - Length(Dr("Cif")))
                StrRegistro &= "045"
                StrRegistro &= Dr("email") & Strings.Space(36 - Length(Dr("email")))
                StrRegistro &= Strings.Space(5)
                Dim DrNewDet13 As DataRow = DtFichero.NewRow

                DrNewDet13("Linea") = StrRegistro
                DtFichero.Rows.Add(DrNewDet13)

                'Tipo 16 Detalle euros
                NumReg += 1
                StrRegistro = strRegistroInicial
                StrRegistro &= Dr("Cif") & Strings.Space(12 - Length(Dr("Cif")))
                StrRegistro &= "085"
                StrRegistro &= Dr("ClavePago") '  01-Mercancías 02-No Mercancías
                StrRegistro &= New String("0", 38) ' ceros
                StrRegistro &= Strings.Space(1)
                Dim DrNewDet16 As DataRow = DtFichero.NewRow

                DrNewDet16("Linea") = StrRegistro
                DtFichero.Rows.Add(DrNewDet16)
            Next


            'Reg.Totales euros
            NumReg += 1
            StrRegistro = "0860"
            StrRegistro &= StrCIF   'NIF propio
            StrRegistro &= "000"
            StrRegistro &= Strings.Space(15)
            StrRegistro &= Strings.Format(ImpTotalDivisa, "000000000000")
            StrRegistro &= Strings.Format(NumReg010Divisa, "00000000")
            StrRegistro &= Strings.Format(NumReg, "0000000000")
            StrRegistro &= Strings.Space(11)
            Dim DrNewTotEURO As DataRow = DtFichero.NewRow
            DrNewTotEURO("Linea") = StrRegistro
            DtFichero.Rows.Add(DrNewTotEURO)
        End If

        'REGISTROS EN DIVISAS
        ImpTotalDivisa = 0
        NumReg010Divisa = 0
        NumRegDivisa = 0

        Dim LstDivisas As List(Of String) = (From DrDivisa As DataRow In DtPagos.Rows Where CStr(DrDivisa("IDMoneda")) <> Monedas.MonedaA.ID Select CStr(DrDivisa("IDMoneda"))).Distinct.ToList

        If LstDivisas.Count > 0 Then
            ' CABECERA EN DIVISA Obligatorio
            For Each IDDivisa As String In LstDivisas
                'TODO FOR EACH DISTINTA DIVISA
                Dim StrISODivisa As String = New Moneda().GetItemRow(IDDivisa)("CodigoISO")


                'Cabecera de divisa
                StrRegistro = "0471"
                StrRegistro &= StrCIF
                StrRegistro &= "000"
                StrRegistro &= StrISODivisa
                StrRegistro &= Strings.Space(72 - Length(StrRegistro))
                NumReg += 1
                NumRegDivisa += 1

                Dim DrNewCabDivisa As DataRow = DtFichero.NewRow
                DrNewCabDivisa("Linea") = StrRegistro
                DtFichero.Rows.Add(DrNewCabDivisa)

                strRegistroInicial = "0671"
                strRegistroInicial &= StrCIF  'NIF propio
                strRegistroInicial &= "000"

                For Each Dr As DataRow In DtPagos.Select("IDMoneda = '" & IDDivisa & "'")
                    'Tipo 1 cabecera DIVISA
                    NumRegDivisa += 1
                    NumReg += 1
                    ImpTotalDivisa += (Dr("ImpVencimiento") * 100)
                    ImpTotal += (Dr("ImpVencimiento") * 100)
                    NumReg010Divisa += 1
                    NumReg010 += 1
                    StrRegistro = strRegistroInicial

                    StrRegistro &= Dr("Cif") & Strings.Space(12 - Length(Dr("Cif")))
                    StrRegistro &= "073"
                    StrRegistro &= Left(Dr("CodigoIban"), 2)
                    StrRegistro &= Mid(Dr("CodigoIban"), 3, 2)
                    StrRegistro &= Mid(Dr("CodigoIban"), 5) & Strings.Space(30 - Length(Mid(Dr("CodigoIban"), 5)))
                    StrRegistro &= Dr("Concepto")
                    StrRegistro &= Strings.Space(6)

                    Dim DrNewDet1 As DataRow = DtFichero.NewRow
                    DrNewDet1("Linea") = StrRegistro
                    DtFichero.Rows.Add(DrNewDet1)

                    'Tipo 2 Detalle DIVISA
                    NumRegDivisa += 1
                    NumReg += 1
                    StrRegistro = strRegistroInicial
                    StrRegistro &= Dr("Cif") & Strings.Space(12 - Length(Dr("Cif")))
                    StrRegistro &= "074"
                    StrRegistro &= Strings.Format(Dr("ImpVencimiento") * 100, "000000000000")
                    StrRegistro &= "3"
                    StrRegistro &= Left(Dr("CodigoISO"), 2)
                    StrRegistro &= Strings.Space(6)
                    StrRegistro &= (Strings.Left(Strings.Trim(Nz(Dr("swift"), " ")) & Space(11), 11))
                    StrRegistro &= Strings.Space(9)

                    Dim DrNewDet2 As DataRow = DtFichero.NewRow
                    DrNewDet2("Linea") = StrRegistro
                    DtFichero.Rows.Add(DrNewDet2)

                    'Tipo 3 Detalle DIVISA
                    NumRegDivisa += 1
                    NumReg += 1
                    StrRegistro = strRegistroInicial
                    StrRegistro &= Dr("Cif") & Strings.Space(12 - Length(Dr("Cif")))
                    StrRegistro &= "075"
                    StrRegistro &= (Strings.Left(Strings.Trim(Nz(Dr("DescBeneficiario"), " ")) & Space(36), 36))
                    StrRegistro &= Strings.Space(5)
                    Dim DrNewDet3 As DataRow = DtFichero.NewRow

                    If Length(StrRegistro) < 72 Then
                        StrRegistro &= Strings.Space(72 - Length(StrRegistro))
                    End If

                    DrNewDet3("Linea") = StrRegistro
                    DtFichero.Rows.Add(DrNewDet3)

                    'Tipo 4 Detalle DIVISA
                    NumRegDivisa += 1
                    NumReg += 1
                    StrRegistro = strRegistroInicial
                    StrRegistro &= Dr("Cif") & Strings.Space(12 - Length(Dr("Cif")))
                    StrRegistro &= "076"
                    StrRegistro &= (Strings.Left(Strings.Trim(Nz(Dr("DireccionPago"), " ")) & Space(36), 36))
                    StrRegistro &= Strings.Space(5)
                    Dim DrNewDet4 As DataRow = DtFichero.NewRow

                    DrNewDet4("Linea") = StrRegistro
                    DtFichero.Rows.Add(DrNewDet4)

                    'Tipo 6 Detalle DIVISA
                    NumRegDivisa += 1
                    NumReg += 1
                    StrRegistro = strRegistroInicial
                    StrRegistro &= Dr("Cif") & Strings.Space(12 - Length(Dr("Cif")))
                    StrRegistro &= "078"
                    StrRegistro &= (Strings.Left(Strings.Trim(Nz(Dr("CodPostal") & Dr("PoblacionPago"), " ")) & Space(36), 36))
                    StrRegistro &= Strings.Space(5)
                    Dim DrNewDet6 As DataRow = DtFichero.NewRow

                    DrNewDet6("Linea") = StrRegistro
                    DtFichero.Rows.Add(DrNewDet6)

                    'Tipo 7 Detalle DIVISA
                    NumRegDivisa += 1
                    NumReg += 1
                    StrRegistro = strRegistroInicial
                    StrRegistro &= Dr("Cif") & Strings.Space(12 - Length(Dr("Cif")))
                    StrRegistro &= "079"
                    StrRegistro &= (Strings.Left(Strings.Trim(Nz(Dr("DescPais"), " ")) & Space(36), 36))
                    StrRegistro &= Strings.Space(5)
                    Dim DrNewDet7 As DataRow = DtFichero.NewRow

                    DrNewDet7("Linea") = StrRegistro
                    DtFichero.Rows.Add(DrNewDet7)

                    'Tipo 10 Detalle DIVISA
                    NumRegDivisa += 1
                    NumReg += 1
                    StrRegistro = strRegistroInicial
                    StrRegistro &= Strings.Format(Dr("FechaVencimiento"), "ddMMyyyy")
                    StrRegistro &= Strings.Space(4)
                    StrRegistro &= "082"
                    StrRegistro &= Strings.Format(Dr("FechaFactura"), "ddMMyyyy")
                    StrRegistro &= Strings.Space(1)
                    StrRegistro &= Dr("Cif") & Strings.Space(12 - Length(Dr("Cif")))
                    StrRegistro &= Strings.Space(1)

                    'para Pagos Agrupados
                    Dim SuFactura As String
                    If Dr("NFactura") = StrPagoAgrupado Then
                        Dim Sql As String = "SELECT tbFacturaCompraCabecera.SuFactura, tbPago.NFactura, tbPago.FechaVencimiento, " & _
                        "tbFacturaCompraCabecera.FechaFactura,tbPago.ImpVencimientoA, " & _
                        "tbPago.IDPagoAgrupado,tbFacturaCompraCabecera.NFactura " & _
                        "FROM tbPago LEFT OUTER JOIN tbFacturaCompraCabecera ON " & _
                        "tbPago.IDFactura = tbFacturaCompraCabecera.IDFactura " & _
                        "where tbPago.IdPagoAgrupado= " & Dr("IDPago")
                        DtPagosAgrupados = AdminData.Execute(Sql, ExecuteCommand.ExecuteReader)

                        If Not DtPagosAgrupados Is Nothing AndAlso DtPagosAgrupados.Rows.Count > 0 Then

                            For Each DrAgp As DataRow In DtPagosAgrupados.Select
                                SuFactura = SuFactura & DrAgp("SuFactura") & ","
                                If Length(SuFactura) > 16 Then
                                    SuFactura = Left(SuFactura, 18)
                                    Exit For
                                End If
                            Next
                        End If
                        StrRegistro &= (Strings.Left(Strings.Trim(Nz(SuFactura, String.Empty)) & Strings.Space(18), 18))
                        SuFactura = String.Empty
                    Else
                        StrRegistro &= (Strings.Left(Strings.Trim(Nz(Dr("SuFactura"), String.Empty)) & Strings.Space(18), 18))
                    End If
                    StrRegistro &= Strings.Space(1)

                    Dim DrNewDet10 As DataRow = DtFichero.NewRow

                    DrNewDet10("Linea") = StrRegistro
                    DtFichero.Rows.Add(DrNewDet10)


                    'Tipo 11 Detalle DIVISA
                    NumRegDivisa += 1
                    NumReg += 1
                    StrRegistro = strRegistroInicial
                    StrRegistro &= Dr("Cif") & Strings.Space(12 - Length(Dr("Cif")))
                    StrRegistro &= "085"
                    StrRegistro &= Dr("ClavePago")
                    StrRegistro &= New String("0", 38) ' ceros
                    StrRegistro &= Strings.Space(1)
                    Dim DrNewDet11 As DataRow = DtFichero.NewRow

                    DrNewDet11("Linea") = StrRegistro
                    DtFichero.Rows.Add(DrNewDet11)


                    'Tipo 14 Detalle DIVISA
                    NumRegDivisa += 1
                    NumReg += 1
                    StrRegistro = strRegistroInicial
                    StrRegistro &= Dr("Cif") & Strings.Space(12 - Length(Dr("Cif")))
                    StrRegistro &= "088"
                    StrRegistro &= Dr("email") & Strings.Space(36 - Length(Dr("email")))
                    StrRegistro &= Strings.Space(5)
                    Dim DrNewDet14 As DataRow = DtFichero.NewRow

                    DrNewDet14("Linea") = StrRegistro
                    DtFichero.Rows.Add(DrNewDet14)


                Next


                'Reg.Totales DIVISA
                NumRegDivisa += 1
                NumReg += 1
                StrRegistro = "0871"
                StrRegistro &= StrCIF   'NIF propio
                StrRegistro &= "000"
                StrRegistro &= StrISODivisa
                StrRegistro &= Strings.Space(12)
                StrRegistro &= Strings.Format(ImpTotalDivisa, "000000000000")
                StrRegistro &= Strings.Format(NumReg010Divisa, "00000000")
                StrRegistro &= Strings.Format(NumRegDivisa, "0000000000")
                StrRegistro &= Strings.Space(11)
                Dim DrNewTotDivisa As DataRow = DtFichero.NewRow
                DrNewTotDivisa("Linea") = StrRegistro
                DtFichero.Rows.Add(DrNewTotDivisa)
                NumRegDivisa = 0
                NumReg010Divisa = 0
                ImpTotalDivisa = 0
            Next
        End If
        StrRegistro = "0962"
        StrRegistro &= StrCIF   'NIF propio
        StrRegistro &= "000"
        StrRegistro &= Strings.Space(15)
        StrRegistro &= Strings.Format(ImpTotal, "000000000000")
        StrRegistro &= Strings.Format(NumReg010, "00000000")
        NumReg = DtFichero.Rows.Count + 1
        StrRegistro &= Strings.Format(NumReg, "0000000000")
        StrRegistro &= Strings.Space(11)
        Dim DrNewTot As DataRow = DtFichero.NewRow
        DrNewTot("Linea") = StrRegistro
        DtFichero.Rows.Add(DrNewTot)
        Return DtFichero
    End Function

#End Region

#Region "Ficheros 347 "

    Public Function GenerarFichero347(ByVal IntAño As Integer) As DataTable
        Dim DtFichero As DataTable
        Dim Informe347 As String = New Parametro().InformeDeclaracion347
        If Length(Informe347) > 0 Then
            Select Case Informe347
                Case enumInformeDec347.decEstatal
                    DtFichero = GenerarFichero347Estatal(IntAño)
                Case enumInformeDec347.decBizkaia
                    DtFichero = GenerarFichero347Bizkaia(IntAño)
                Case enumInformeDec347.decGipuzkoa
                    DtFichero = GenerarFichero347Gipuzkoa(IntAño)
                Case enumInformeDec347.decAlava
                    DtFichero = GenerarFichero347Alava(IntAño)
                Case enumInformeDec347.decNavarra
                    DtFichero = GenerarFichero347Navarra(IntAño)
            End Select
        End If
        Return DtFichero
    End Function

    Private Function GenerarFichero347Alava(ByVal IntAño As Integer) As DataTable
        Dim StrCifDet, StrDenominacionDet, StrCPProvinciaDet, StrCodPaisDet, _
        StrDenominacionEmpresa, StrCifCab, StrCPProvinciaCab, StrCodPaisParam, _
        StrSql, StrWhere, StrRegistro As String

        Dim IntNumRegs, IntNumRegClientes, IntNumRegProveedores As Integer
        Dim DblImpClientes, DblImpProveedores, DblImpRegs As Double

        'Creamos el datatable que contendrá las líneas del fichero
        Dim DtFichero As New DataTable
        DtFichero.Columns.Add("Linea", GetType(String))

        'Obtenemos los parametros para crear el fichero
        Dim DtParametros As DataTable = New Parametro().Filter(, "left(IDParametro,3)='347'")
        If DtParametros Is Nothing OrElse DtParametros.Rows.Count = 0 Then ApplicationService.GenerateError("No se encuentran datos en parámetros para la generación del fichero, por favor, actualícelos.")

        'Obtenemos los Importes Totales para Cientes y Proveedores, en las dos monedas
        Dim DtEmpresa As DataTable = AdminData.Filter("tbDatosEmpresa")
        StrSql = "Sum(ImpTotalA) as ImporteTotalA,Count(*) as NumRegistros"
        Dim DtClientes347Tot As DataTable = AdminData.Filter("VRptInformeTesoreria347Clientes", StrSql, "Año=" & IntAño, , False)
        StrSql = "Sum(ImpTotalA) as ImporteTotalA,Count(*)as NumRegistros"
        Dim DtProveedores347Tot As DataTable = AdminData.Filter("VRptInformeTesoreria347Proveedores", StrSql, "Año=" & IntAño, , False)

        Dim DrFil() As DataRow = DtParametros.Select("IDParametro=347CODPAIS")
        If DrFil.Length > 0 Then StrCodPaisParam = Nz(DrFil(0)("Valor"), "   ")

        DrFil = DtParametros.Select("IDParametro=347MONEDA")
        If DrFil.Length > 0 Then
            IntNumRegClientes = Nz(DtClientes347Tot.Rows(0)("NumRegistros"), 0)
            DblImpClientes = Nz(DtClientes347Tot.Rows(0)("ImporteTotalA"), 0)
        End If

        DrFil = DtParametros.Select("IDParametro=347MONEDA")
        If DrFil.Length > 0 Then
            IntNumRegProveedores = Nz(DtProveedores347Tot.Rows(0)("NumRegistros"), 0)
            DblImpProveedores = Nz(DtProveedores347Tot.Rows(0)("ImporteTotalA"), 0)
        End If

        IntNumRegs = IntNumRegProveedores + IntNumRegClientes
        DblImpRegs = DblImpClientes + DblImpProveedores

        StrCPProvinciaCab = Strings.Left(Nz(DtEmpresa.Rows(0)("CodPostal"), ""), 2)
        If Length(Nz(DtEmpresa.Rows(0)("Cif"), "")) > 0 Then
            StrCifCab = Strings.Left(Nz(DtEmpresa.Rows(0)("Cif"), ""), 9)
            StrCifCab = New String(" ", 9 - Length(StrCifCab)) & StrCifCab
        Else : StrCifCab = Strings.Space(9)
        End If
        StrDenominacionEmpresa = Left(CStr(Nz(TratarSimbolosEspeciales(DtEmpresa.Rows(0)("DescEmpresa")), "")).ToUpper, 40)

        'Registro 1: Declarante
        StrRegistro = "1"
        StrRegistro &= "347"
        StrRegistro &= Strings.Format(IntAño, "0000")
        StrRegistro &= StrCifCab
        StrRegistro &= StrDenominacionEmpresa

        DrFil = DtParametros.Select("IDParametro=347TSOPORT")
        If DrFil.Length > 0 Then
            StrRegistro &= Nz(Strings.Left(DrFil(0)("Valor"), 1), "D")
        End If

        DrFil = DtParametros.Select("IDParametro=347TLFCONT")
        If DrFil.Length > 0 Then
            StrRegistro &= Nz(Strings.Format(DrFil("Valor"), "000000000"), "999999999")
        End If

        DrFil = DtParametros.Select("IDParametro=347NOMCONT")
        If DrFil.Length > 0 Then
            StrRegistro &= Nz(Strings.Left(DrFil(0)("Valor"), 40), "SIN DEFINIR                             ")
            If Length(DrFil(0)("Valor") & String.Empty) < 40 Then StrRegistro &= New String(" ", 40 - Length(DrFil(0)("Valor") & String.Empty))
        End If

        DrFil = DtParametros.Select("IDParametro=347MONEDA")
        If DrFil.Length > 0 Then
            StrRegistro &= IIf(DrFil(0)("Valor") = "Pesetas", "3470000000001", "3480000000001")
        End If

        DrFil = DtParametros.Select("IDParametro=347DECSUST")
        If DrFil.Length > 0 Then 'no ofrecemos posibilidad de declaracion complementaria
            StrRegistro &= IIf(DrFil(0)("Valor") = True, " S", "  ")
        End If
        StrRegistro &= New String(" ", 13) ' posiciones 123-135
        StrRegistro &= Strings.Format(IntNumRegs, "000000000")
        StrRegistro &= Format(xRound(DblImpRegs, 2) * 100, "000000000000000")
        StrRegistro &= Strings.Space(54)
        StrRegistro &= Strings.Space(13)

        Dim DrNew As DataRow = DtFichero.NewRow
        DrNew("Linea") = StrRegistro
        DtFichero.Rows.Add(DrNew)

        'Registro 2: Operacion
        Dim DtClientes347Det As DataTable = AdminData.Filter("VRptInformeTesoreria347Clientes", , "Año=" & IntAño, , False)
        Dim DtProveedores347Det As DataTable = AdminData.Filter("VRptInformeTesoreria347Proveedores", , "Año=" & IntAño, , False)

        If Not DtProveedores347Det Is Nothing AndAlso DtProveedores347Det.Rows.Count > 0 Then
            For Each Dr As DataRow In DtProveedores347Det.Select
                If Length(Nz(Dr("CifProveedor"), "")) > 0 Then
                    StrCifDet = Strings.Left(Nz(Dr("CifProveedor"), ""), 9)
                    StrCifDet = New String(" ", 9 - Length(StrCifDet)) & StrCifDet
                Else : StrCifDet = Strings.Space(9)
                End If
                StrDenominacionDet = Strings.Left(CStr(Nz(TratarSimbolosEspeciales(Dr("RazonSocial")), "")).ToUpper, 40)

                If StrCodPaisParam = "011" Then
                    StrCodPaisDet = "   "
                    StrCPProvinciaDet = Strings.Left(Nz(Dr("CodPostal"), "00"), 2)
                Else
                    StrCodPaisDet = StrCodPaisParam
                    StrCPProvinciaDet = "99"
                End If

                StrRegistro = "2"
                StrRegistro &= "347"
                StrRegistro &= Strings.Format(IntAño, "0000")
                StrRegistro &= StrCifCab
                StrRegistro &= StrCifDet
                StrRegistro &= New String(" ", 9) 'nuevo 12/03/01
                StrRegistro &= StrDenominacionDet
                StrRegistro &= "D" 'tipo de hoja
                StrRegistro &= StrCPProvinciaDet '& "000"
                StrRegistro &= StrCodPaisDet
                StrRegistro &= "A" 'proveedor
                StrRegistro &= Strings.Format(xRound(Dr("ImpTotalA") * 100, 0), "000000000000000") ' centimos de Euro
                StrRegistro &= Strings.Space(1)
                StrRegistro &= Strings.Space(1)
                StrRegistro &= Strings.Space(151)

                Dim DrNew1 As DataRow = DtFichero.NewRow
                DrNew1("Linea") = StrRegistro
                DtFichero.Rows.Add(DrNew1)
            Next
        End If

        If Not DtClientes347Det Is Nothing AndAlso DtClientes347Det.Rows.Count <> 0 Then
            For Each Dr As DataRow In DtClientes347Det.Select
                StrCifDet = Strings.Left(Nz(Dr("CifCliente"), ""), 9)
                StrCifDet = New String(" ", 9 - Len(StrCifDet)) & StrCifDet
                StrDenominacionDet = Strings.Left(CStr(Nz(TratarSimbolosEspeciales(Dr("RazonSocial")), "")).ToUpper, 40)
                If StrCodPaisParam = "011" Then
                    StrCodPaisDet = "   "
                    StrCPProvinciaDet = Strings.Left(Nz(Dr("CodPostal"), "00"), 2)
                Else
                    StrCodPaisDet = StrCodPaisParam
                    StrCPProvinciaDet = "99"
                End If
                StrRegistro = "2"
                StrRegistro &= "347"
                StrRegistro &= Strings.Format(IntAño, "0000")
                StrRegistro &= StrCifCab
                StrRegistro &= StrCifDet
                StrRegistro &= New String(" ", 9)
                StrRegistro &= StrDenominacionDet
                StrRegistro &= "D" 'tipo de hoja
                StrRegistro &= StrCPProvinciaDet
                StrRegistro &= StrCodPaisDet
                StrRegistro &= "B" 'Cliente
                StrRegistro &= Strings.Format(xRound(Dr("ImpTotalA") * 100, 0), "000000000000000") ' centimos de Euro
                StrRegistro &= Strings.Space(1)
                StrRegistro &= Strings.Space(1)
                StrRegistro &= Strings.Space(151)

                Dim DrNew2 As DataRow = DtFichero.NewRow
                DrNew2("Linea") = StrRegistro
                DtFichero.Rows.Add(DrNew2)
            Next
        End If
        Return DtFichero
    End Function

    Private Function GenerarFichero347Bizkaia(ByVal IntAño As Integer) As DataTable
        Dim StrRegistro, StrWhere, StrCifDet, StrDenominacionDet, StrCPProvinciaDet, _
        StrCodPaisDet, StrDenominacionEmpresa, StrCifCab, StrCPProvinciaCab, StrCodPaisParam As String
        Dim DblImpClientes, DblImpRegs, DblImpProveedores As Double
        Dim IntNumRegs, IntNumRegClientes, IntNumRegProveedores As Integer

        Dim DtFichero As New DataTable
        DtFichero.Columns.Add("Linea", GetType(String))

        'Obtenemos los parametros para crear el fichero
        Dim DtParametros As DataTable = New Parametro().Filter(, "left(IDParametro,3)='347'")
        If DtParametros Is Nothing OrElse DtParametros.Rows.Count = 0 Then ApplicationService.GenerateError("No se encuentran datos en parámetros para la generación del fichero, por favor, actualícelos.")

        'Obtenemos los Importes Totales para Cientes y Proveedores, en las dos monedas
        Dim DtEmpresa As DataTable = AdminData.Filter("tbDatosEmpresa")
        Dim StrSql As String = "Sum(ImpTotalA) as ImporteTotalA,Count(*) as NumRegistros"
        Dim DtClientes347Tot As DataTable = AdminData.Filter("VRptInformeTesoreria347Clientes", StrSql, "Año=" & IntAño, , False)
        StrSql = "Sum(ImpTotalA) as ImporteTotalA,Count(*)as NumRegistros"
        Dim DtProveedores347Tot As DataTable = AdminData.Filter("VRptInformeTesoreria347Proveedores", StrSql, "Año=" & IntAño, , False)

        Dim DrFil() As DataRow = DtParametros.Select("IDParametro=347CODPAIS")
        If DrFil.Length > 0 Then StrCodPaisParam = Nz(DrFil(0)("Valor"), "   ")

        DrFil = DtParametros.Select("IDParametro=347MONEDA")
        If DrFil.Length > 0 Then
            IntNumRegClientes = Nz(DtClientes347Tot.Rows(0)("NumRegistros"), 0)
            DblImpClientes = Nz(DtClientes347Tot.Rows(0)("ImporteTotalA"), 0)
        End If

        DrFil = DtParametros.Select("IDParametro=347MONEDA")
        If DrFil.Length > 0 Then
            IntNumRegProveedores = Nz(DtProveedores347Tot.Rows(0)("NumRegistros"), 0)
            DblImpProveedores = Nz(DtProveedores347Tot.Rows(0)("ImporteTotalA"), 0)
        End If

        IntNumRegs = IntNumRegProveedores + IntNumRegClientes
        DblImpRegs = DblImpClientes + DblImpProveedores

        StrCPProvinciaCab = Strings.Left(Nz(DtEmpresa.Rows(0)("CodPostal"), ""), 2)
        If Length(Nz(DtEmpresa.Rows(0)("Cif"), "")) > 0 Then
            StrCifCab = Strings.Left(Nz(DtEmpresa.Rows(0)("Cif"), ""), 9)
            StrCifCab = New String(" ", 9 - Length(StrCifCab)) & StrCifCab
        Else : StrCifCab = Strings.Space(9)
        End If
        StrDenominacionEmpresa = Strings.Left(CStr(Nz(TratarSimbolosEspeciales(DtEmpresa.Rows(0)("DescEmpresa")), "")).ToUpper, 40)

        'Registro 1: Declarante
        StrRegistro = "1"
        StrRegistro &= "347"
        StrRegistro &= Strings.Format(IntAño, "0000")
        StrRegistro &= StrCifCab
        StrRegistro &= StrDenominacionEmpresa

        DrFil = DtParametros.Select("IDParametro=347TSOPORT")
        If DrFil.Length > 0 Then
            StrRegistro &= Nz(Strings.Left(DrFil(0)("Valor"), 1), "D")
        End If

        DrFil = DtParametros.Select("IDParametro=347TLFCONT")
        If DrFil.Length > 0 Then
            StrRegistro &= Nz(Strings.Format(DrFil(0)("Valor"), "000000000"), "999999999")
        End If

        DrFil = DtParametros.Select("IDParametro=347NOMCONT")
        If DrFil.Length > 0 Then
            StrRegistro &= Nz(Strings.Left(DrFil(0)("Valor"), 40), "SIN DEFINIR                             ")
            If Length(DrFil(0)("Valor") & String.Empty) < 40 Then StrRegistro &= New String(" ", 40 - Length(DrFil(0)("Valor") & String.Empty))
        End If

        DrFil = DtParametros.Select("IDParametro=347MONEDA")
        If DrFil.Length > 0 Then
            StrRegistro &= IIf(DrFil(0)("Valor") = "Pesetas", "3470000000001", "3480000000001")
        End If

        DrFil = DtParametros.Select("IDParametro=347DECSUST")
        If DrFil.Length > 0 Then 'no ofrecemos posibilidad de declaracion complementaria
            StrRegistro &= IIf(DrFil(0)("Valor") = True, " S", "  ")
        End If

        StrRegistro &= New String("0", 13) ' posiciones 123-135
        StrRegistro &= Strings.Format(IntNumRegs, "000000000")
        StrRegistro &= Format(xRound(DblImpRegs, 2) * 100, "000000000000000")
        StrRegistro &= New String("0", 9)
        StrRegistro &= New String("0", 15)
        StrRegistro &= Strings.Space(54)
        StrRegistro &= Strings.Space(13)

        Dim DrNew As DataRow = DtFichero.NewRow
        DrNew("Linea") = StrRegistro
        DtFichero.Rows.Add(DrNew)

        'Registro 2: Operacion
        Dim DtClientes347Det As DataTable = AdminData.Filter("VRptInformeTesoreria347Clientes", , "Año=" & IntAño, , False)
        Dim DtProveedores347Det As DataTable = AdminData.Filter("VRptInformeTesoreria347Proveedores", , "Año=" & IntAño, , False)

        If Not DtProveedores347Det Is Nothing AndAlso DtProveedores347Det.Rows.Count > 0 Then
            For Each Dr As DataRow In DtProveedores347Det.Select
                If Len(Nz(Dr("CifProveedor"), "")) > 0 Then
                    StrCifDet = Strings.Left(Nz(Dr("CifProveedor"), ""), 9)
                    StrCifDet = New String(" ", 9 - Length(StrCifDet)) & StrCifDet
                Else
                    StrCifDet = Strings.Space(9)
                End If
                StrDenominacionDet = Strings.Left(CStr(Nz(TratarSimbolosEspeciales(Dr("RazonSocial")), "")).ToUpper, 40)

                If StrCodPaisParam = "011" Then
                    StrCodPaisDet = "   "
                    StrCPProvinciaDet = Strings.Left(Nz(Dr("CodPostal"), "00"), 2)
                Else
                    StrCodPaisDet = StrCodPaisParam
                    StrCPProvinciaDet = "99"
                End If

                StrRegistro = "2"
                StrRegistro &= "347"
                StrRegistro &= Strings.Format(IntAño, "0000")
                StrRegistro &= StrCifCab
                StrRegistro &= StrCifDet
                StrRegistro &= New String(" ", 9)
                StrRegistro &= StrDenominacionDet
                StrRegistro &= "D" 'tipo de hoja
                StrRegistro &= StrCPProvinciaDet '& "000"
                StrRegistro &= StrCodPaisDet
                StrRegistro &= "A" 'proveedor
                StrRegistro &= Strings.Format(xRound(Dr("ImpTotalA") * 100, 0), "000000000000000") ' centimos de Euro
                StrRegistro &= Strings.Space(1)
                StrRegistro &= Strings.Space(1)
                StrRegistro &= Strings.Space(151)

                Dim DrNew1 As DataRow = DtFichero.NewRow
                DrNew1("Linea") = StrRegistro
                DtFichero.Rows.Add(DrNew1)
            Next
        End If

        If Not DtClientes347Det Is Nothing AndAlso DtClientes347Det.Rows.Count <> 0 Then
            For Each DrClie As DataRow In DtClientes347Det.Select
                StrCifDet = Strings.Left(Nz(DrClie("CifCliente"), ""), 9)
                StrCifDet = New String(" ", 9 - Len(StrCifDet)) & StrCifDet
                StrDenominacionDet = Strings.Left(CStr(Nz(TratarSimbolosEspeciales(DrClie("RazonSocial")), "")).ToUpper, 40)

                If StrCodPaisParam = "011" Then
                    StrCodPaisDet = "   "
                    StrCPProvinciaDet = Strings.Left(Nz(DrClie("CodPostal"), "00"), 2)
                Else
                    StrCodPaisDet = StrCodPaisParam
                    StrCPProvinciaDet = "99"
                End If

                StrRegistro = "2"
                StrRegistro &= "347"
                StrRegistro &= Strings.Format(IntAño, "0000") 'Right(intAño, 2)
                StrRegistro &= StrCifCab
                StrRegistro &= StrCifDet
                StrRegistro &= New String(" ", 9) 'nuevo 12/03/01
                StrRegistro &= StrDenominacionDet
                StrRegistro &= "D" 'tipo de hoja
                StrRegistro &= StrCPProvinciaDet '& "000"
                StrRegistro &= StrCodPaisDet
                StrRegistro &= "B" 'Cliente
                StrRegistro &= Strings.Format(xRound(DrClie("ImpTotalA") * 100, 0), "000000000000000") ' centimos de Euro
                StrRegistro &= Space(1)
                StrRegistro &= Space(1)
                StrRegistro &= StrRegistro & Space(151)
                Dim DrNew2 As DataRow = DtFichero.NewRow
                DrNew2("linea") = StrRegistro
                DtFichero.Rows.Add(DrNew2)
            Next
        End If
        Return DtFichero
    End Function

    Private Function GenerarFichero347Estatal(ByVal IntAño As Integer) As DataTable
        Dim StrRegistro, StrDenominacionEmpresa, StrCifCab, StrCPProvinciaCab, _
        StrCifDet, StrDenominacionDet, StrCPProvinciaDet, StrCodPaisDet, _
        StrCodPaisParam, StrSql, StrWhere As String
        Dim IntNumRegs, IntNumRegClientes, IntNumRegProveedores As Short
        Dim DblImpClientes, DblImpRegs, DblImpProveedores As Double

        Dim DtFichero As New DataTable
        DtFichero.Columns.Add("Linea", GetType(String))

        'Obtenemos los parametros para crear el fichero
        Dim DtParametros As DataTable = New Parametro().Filter(, "left(IDParametro,3)='347'")
        If DtParametros.Rows.Count = 0 Then ApplicationService.GenerateError("No se encuentran datos en parámetros para la generación del fichero, por favor, actualícelos.")

        'Obtenemos los Importes Totales para Cientes y Proveedores, en las dos monedas
        Dim DtEmpresa As DataTable = AdminData.Filter("tbDatosEmpresa")
        StrSql = "Sum(ImpTotalA) as ImporteTotalA, Count(*) as NumRegistros"
        Dim DtClientes347Tot As DataTable = AdminData.Filter("VRptInformeTesoreria347Clientes", StrSql, "Año=" & IntAño, , False)
        StrSql = "Sum(ImpTotalA) as ImporteTotalA, Count(*)as NumRegistros"
        Dim DtProveedores347Tot As DataTable = AdminData.Filter("VRptInformeTesoreria347Proveedores", StrSql, "Año=" & IntAño, , False)

        Dim DrFil() As DataRow = DtParametros.Select("IDParametro=347CODPAIS")
        If DrFil.Length > 0 Then
            StrCodPaisParam = Nz(DrFil(0)("Valor"), "   ")
        End If

        DrFil = DtParametros.Select("IDParametro=347MONEDA")
        If DrFil.Length > 0 Then
            IntNumRegClientes = Nz(DtClientes347Tot.Rows(0)("NumRegistros"), 0)
            DblImpClientes = Nz(DtClientes347Tot.Rows(0)("ImporteTotalA"), 0)
        End If

        DrFil = DtParametros.Select("IDParametro=347MONEDA")
        If DrFil.Length > 0 Then
            IntNumRegProveedores = Nz(DtProveedores347Tot.Rows(0)("NumRegistros"), 0)
            DblImpProveedores = Nz(DtProveedores347Tot.Rows(0)("ImporteTotalA"), 0)
        End If
        IntNumRegs = IntNumRegProveedores + IntNumRegClientes
        'Registro 1: Declarante
        StrCPProvinciaCab = Strings.Left(Nz(DtEmpresa.Rows(0)("CodPostal"), ""), 2)
        If Length(Nz(DtEmpresa.Rows(0)("Cif"), "")) > 0 Then
            StrCifCab = Strings.Left(Nz(DtEmpresa.Rows(0)("Cif"), ""), 9)
            StrCifCab = New String(" ", 9 - Length(StrCifCab)) & StrCifCab
        Else : StrCifCab = Strings.Space(9)
        End If
        StrDenominacionEmpresa = Strings.Left(CStr(Nz(TratarSimbolosEspeciales(DtEmpresa.Rows(0)("DescEmpresa")), "")).ToUpper, 40)

        StrRegistro = "1" 'Tipo de Registro
        StrRegistro &= "347" 'Modelo Declaracion
        StrRegistro &= Strings.Format(IntAño, "0000") 'Ejercicio
        StrRegistro &= StrCifCab 'NIF Declarante
        StrRegistro &= StrDenominacionEmpresa 'Razon Social del Declarante

        DrFil = DtParametros.Select("IDParametro=347TSOPORT")
        If DrFil.Length > 0 Then
            StrRegistro &= Nz(Strings.Left(DrFil(0)("Valor"), 1), "D")
        End If

        DrFil = DtParametros.Select("IDParametro=347TLFCONT")
        If DrFil.Length > 0 Then
            StrRegistro &= Nz(Strings.Format(DrFil(0)("Valor"), "000000000"), "999999999")
        End If

        DrFil = DtParametros.Select("IDParametro=347NOMCONT")
        If DrFil.Length > 0 Then
            StrRegistro &= Nz(Strings.Left(DrFil(0)("Valor"), 40), "SIN DEFINIR                             ")
            If Length(DrFil(0)("Valor") & String.Empty) < 40 Then StrRegistro &= New String(" ", 40 - Length(DrFil(0)("Valor") & String.Empty))
        End If

        StrRegistro &= New String("0", 13) 'Num de Declaracion
        DrFil = DtParametros.Select("IDParametro=347DECSUST")
        If DrFil.Length > 0 Then
            StrRegistro &= IIf(DrFil(0)("Valor"), " S", "  ")
        End If
        StrRegistro &= New String("0", 13) 'Num de Declaracion anterior
        StrRegistro &= Strings.Format(IntNumRegs, "000000000") 'Num total de entidades
        StrRegistro &= Format((DblImpClientes + DblImpProveedores) * 100, "000000000000000") 'Importe total de las operaciones
        StrRegistro &= New String("0", 9) 'Num total de inmuebles
        StrRegistro &= New String("0", 15) 'Importe total de las op de arrendamiento
        StrRegistro &= Strings.Space(54) 'Blancos
        StrRegistro &= Strings.Space(13) 'Sello electronico

        Dim DrNew As DataRow = DtFichero.NewRow
        DrNew("Linea") = StrRegistro
        DtFichero.Rows.Add(DrNew)

        'Registro 2: Operacion
        Dim DtClientes347Det As DataTable = AdminData.Filter("VRptInformeTesoreria347Clientes", , "Año=" & IntAño, , False)
        Dim DtProveedores347Det As DataTable = AdminData.Filter("VRptInformeTesoreria347Proveedores", , "Año=" & IntAño, , False)

        If Not DtProveedores347Det Is Nothing AndAlso DtProveedores347Det.Rows.Count > 0 Then
            For Each Dr As DataRow In DtProveedores347Det.Select
                If Length(Nz(Dr("CifProveedor"), "")) > 0 Then
                    StrCifDet = Strings.Left(Nz(Dr("CifProveedor"), ""), 9)
                    StrCifDet = New String(" ", 9 - Length(StrCifDet)) & StrCifDet
                Else : StrCifDet = Strings.Space(9)
                End If
                StrDenominacionDet = Strings.Left(CStr(Nz(TratarSimbolosEspeciales(Dr("RazonSocial")), "")).ToUpper, 40)

                If StrCodPaisParam = "011" Then
                    StrCodPaisDet = "   "
                    StrCPProvinciaDet = Strings.Left(Nz(Dr("CodPostal"), "00"), 2)
                Else
                    StrCodPaisDet = StrCodPaisParam
                    StrCPProvinciaDet = "99"
                End If

                StrRegistro = "2" 'Tipo Registro
                StrRegistro &= "347" 'Modelo Declaracion
                StrRegistro &= Strings.Format(IntAño, "0000") 'Ejercicio
                StrRegistro &= StrCifCab 'NIF Declarante
                StrRegistro &= StrCifDet 'NIF Declarado
                StrRegistro &= New String(" ", 9) 'NIF Representante Legal
                StrRegistro &= StrDenominacionDet 'Razon Social
                StrRegistro &= "D" 'Tipo de hoja
                StrRegistro &= StrCPProvinciaDet '& "000"  'Codigo Provincia
                StrRegistro &= StrCodPaisDet 'Codigo Pais
                StrRegistro &= "A" 'proveedor              'Clave Codigo(Adquisiciones de bienes y servicios superiores a 500.000pts)
                StrRegistro &= Format(Dr("ImpTotalA") * 100, "000000000000000") 'Importe Operaciones
                StrRegistro &= Strings.Space(1) 'Operacion seguro
                StrRegistro &= Strings.Space(1) 'Arrendamiento local negocio
                StrRegistro &= Strings.Space(151) 'Blancos  100-250

                Dim DrNew1 As DataRow = DtFichero.NewRow
                DrNew1("Linea") = StrRegistro
                DtFichero.Rows.Add(DrNew1)
            Next
        End If

        If Not DtClientes347Det Is Nothing AndAlso DtClientes347Det.Rows.Count > 0 Then
            For Each Dr As DataRow In DtClientes347Det.Select
                If Length(Nz(Dr("CifCliente"))) > 0 Then
                    StrCifDet = Strings.Left(Nz(Dr("CifCliente"), ""), 9)
                    StrCifDet = New String(" ", 9 - Length(StrCifDet)) & StrCifDet
                Else
                    StrCifDet = Space(9)
                End If
                StrDenominacionDet = Strings.Left(CStr(Nz(TratarSimbolosEspeciales(Dr("RazonSocial")), "")).ToUpper, 40)
                If StrCodPaisParam = "011" Then
                    StrCodPaisDet = "   "
                    StrCPProvinciaDet = Strings.Left(Nz(Dr("CodPostal"), "00"), 2)
                Else
                    StrCodPaisDet = StrCodPaisParam
                    StrCPProvinciaDet = "99"
                End If
                StrRegistro = "2" 'Tipo Registro
                StrRegistro &= "347" 'Modelo Declaracion
                StrRegistro &= Strings.Format(IntAño, "0000") 'Ejercicio
                StrRegistro &= StrCifCab 'NIF Declarante
                StrRegistro &= StrCifDet 'NIF Declarado
                StrRegistro &= New String(" ", 9) 'NIF Representante legal
                StrRegistro &= StrDenominacionDet 'Razon Social
                StrRegistro &= "D" 'Tipo de hoja
                StrRegistro &= StrCPProvinciaDet '& "000"  'Codigo Provincia
                StrRegistro &= StrCodPaisDet
                StrRegistro &= "B" 'Clave Codigo (Entregas de bienes y servicios superiores a 500.000 pts)
                StrRegistro &= Format(Dr("ImpTotalA") * 100, "000000000000000") 'Importe Total
                StrRegistro &= Strings.Space(1) 'Operacion seguro
                StrRegistro &= Strings.Space(1) 'Arrendamiento local negocio
                StrRegistro &= Strings.Space(151) 'Blancos  100-250

                Dim DrNew2 As DataRow = DtFichero.NewRow
                DrNew2("Linea") = StrRegistro
                DtFichero.Rows.Add(DrNew2)
            Next
        End If
        Return DtFichero
    End Function

    Private Function GenerarFichero347Gipuzkoa(ByVal IntAño As Integer) As DataTable
        Dim StrRegistro, StrDenominacionEmpresa, StrCifCab, StrCPProvinciaCab, StrCifDet, StrDenominacionDet, StrCPProvinciaDet, StrCodPaisDet, StrCodPaisParam As String
        Dim DblImpClientes, DblImpProveedores As Double
        Dim IntNumRegProveedores, IntNumRegClientes As Integer

        'Creamos el datatable que contendrá las líneas del fichero
        Dim DtFichero As New DataTable
        DtFichero.Columns.Add("Linea", GetType(String))

        'Obtenemos los parametros para crear el fichero
        Dim DtParametros As DataTable = New Parametro().Filter(, "left(IDParametro,3)='347'")
        If DtParametros Is Nothing OrElse DtParametros.Rows.Count = 0 Then ApplicationService.GenerateError("No se encuentran datos en parámetros para la generación del fichero, por favor, actualícelos.")

        'Obtenemos los Importes Totales para Cientes y Proveedores, en las dos monedas
        Dim DtEmpresas As DataTable = AdminData.Filter("tbDatosEmpresa")
        Dim StrSql As String = "Sum(ImpTotalA) as ImporteTotalA, Count(*) as NumRegistros"
        Dim DtClientes347Tot As DataTable = AdminData.Filter("VRptInformeTesoreria347Clientes", StrSql, "Año=" & IntAño, , False)
        StrSql = "Sum(ImpTotalA) as ImporteTotalA, Count(*)as NumRegistros"
        Dim DtProveedores347Tot As DataTable = AdminData.Filter("VRptInformeTesoreria347Proveedores", StrSql, "Año=" & IntAño, , False)

        Dim DrSel() As DataRow = DtParametros.Select("IDParametro='347CODPAIS'")
        If DrSel.Length > 0 Then StrCodPaisParam = Nz(DrSel(0)("Valor"), "   ")

        DrSel = DtParametros.Select("IDParametro='347MONEDA'")
        If DrSel.Length > 0 Then
            IntNumRegClientes = Nz(DtClientes347Tot.Rows(0)("NumRegistros"), 0)
            DblImpClientes = Nz(DtClientes347Tot.Rows(0)("ImporteTotalA"), 0)
        End If

        DrSel = DtParametros.Select("IDParametro='347MONEDA'")
        If DrSel.Length > 0 Then
            IntNumRegProveedores = Nz(DtProveedores347Tot.Rows(0)("NumRegistros"), 0)
            DblImpProveedores = Nz(DtProveedores347Tot.Rows(0)("ImporteTotalA"), 0)
        End If

        'Registro 1: Declarante

        StrCPProvinciaCab = Strings.Left(Nz(DtEmpresas.Rows(0)("CodPostal"), ""), 2)
        If Length(Nz(DtEmpresas.Rows(0)("Cif"), "")) > 0 Then
            StrCifCab = Strings.Left(Nz(DtEmpresas.Rows(0)("Cif"), ""), 9)
            StrCifCab = New String(" ", 9 - Length(StrCifCab)) & StrCifCab
        Else : StrCifCab = Strings.Space(9)
        End If
        StrDenominacionEmpresa = Strings.Left(CStr(Nz(TratarSimbolosEspeciales(DtEmpresas.Rows(0)("DescEmpresa")), "")).ToUpper, 40)

        StrRegistro = "1" 'Tipo de Registro
        StrRegistro &= "347" 'Modelo Declaracion
        StrRegistro &= Strings.Format(IntAño, "0000") 'Ejercicio
        StrRegistro &= StrCPProvinciaCab 'Codigo provincia
        StrRegistro &= StrCifCab 'NIF Declarante
        StrRegistro &= StrDenominacionEmpresa 'Razon Social del Declarante
        StrRegistro &= Strings.Format(IntNumRegProveedores, "000000000") 'Num de personas (Adquisiciones)
        StrRegistro &= Format(DblImpProveedores * 100, "000000000000000") 'Importe de adquisiciones
        StrRegistro &= Strings.Format(IntNumRegClientes, "000000000") 'Num de personas (Adquisiciones)
        StrRegistro &= Format(DblImpClientes * 100, "000000000000000") 'Importe de adquisiciones
        StrRegistro &= New String("0", 9) 'Num Pagos
        StrRegistro &= New String("0", 15) 'Importe Pagos
        StrRegistro &= New String("0", 9) 'Num Compras al margen
        StrRegistro &= New String("0", 15) 'Importe Compras al margen
        StrRegistro &= New String("0", 9) 'Num Subvenciones
        StrRegistro &= New String("0", 15) 'Importe Subvenciones
        StrRegistro &= "E" 'Unidad monetaria

        Dim DrNew As DataRow = DtFichero.NewRow
        DrNew("Linea") = StrRegistro
        DtFichero.Rows.Add(DrNew)

        'Registro 2: Operacion
        Dim DtClientes347Det As DataTable = AdminData.Filter("VRptInformeTesoreria347Clientes", , "Año=" & IntAño, , False)
        Dim DtProveedores347Det As DataTable = AdminData.Filter("VRptInformeTesoreria347Proveedores", , "Año=" & IntAño, , False)

        If Not DtProveedores347Det Is Nothing AndAlso DtProveedores347Det.Rows.Count > 0 Then
            For Each Dr As DataRow In DtProveedores347Det.Select
                If Length(Nz(Dr("CifProveedor"), "")) > 0 Then
                    StrCifDet = Strings.Left(Nz(Dr("CifProveedor"), ""), 9)
                    StrCifDet = New String(" ", 9 - Length(StrCifDet)) & StrCifDet
                Else : StrCifDet = Strings.Space(9)
                End If
                StrDenominacionDet = Strings.Left(CStr(Nz(TratarSimbolosEspeciales(Dr("RazonSocial")), "")).ToUpper, 40)

                If StrCodPaisParam = "011" Then
                    StrCodPaisDet = "   " 'Código para residentes
                    StrCPProvinciaDet = Strings.Left(Nz(Dr("CodPostal"), "00"), 2)
                Else
                    StrCodPaisDet = FormatearNumeros(StrCodPaisParam, 3)
                    StrCPProvinciaDet = "99"
                End If

                StrRegistro = "2" 'Tipo Registro
                StrRegistro &= StrRegistro & "347" 'Modelo Declaracion
                StrRegistro &= Format(IntAño, "0000") 'Ejercicio
                StrRegistro &= StrCPProvinciaCab 'Codigo provincia
                StrRegistro &= StrCifCab 'NIF Declarante
                StrRegistro &= "A" 'proveedor              'Clave Codigo(Adquisiciones de bienes y servicios superiores a 500.000pts)
                StrRegistro &= StrCifDet 'NIF Declarado
                StrRegistro &= StrDenominacionDet 'Razon Social
                StrRegistro &= New String(" ", 1) 'Representante Legal
                StrRegistro &= StrCPProvinciaDet & StrCodPaisDet 'Codigo Provincia/Pais
                StrRegistro &= Format(Dr("ImpTotalA") * 100, "000000000000000") 'Importe Operaciones
                StrRegistro &= Strings.Space(1) 'Operacion seguro
                StrRegistro &= Strings.Space(1) 'Arrendamiento local negocio
                StrRegistro &= Strings.Space(88) 'Blancos  93-180

                Dim DrNew1 As DataRow = DtFichero.NewRow
                DrNew1("Linea") = StrRegistro
                DtFichero.Rows.Add(DrNew1)
            Next
        End If

        If Not DtClientes347Det Is Nothing AndAlso DtClientes347Det.Rows.Count > 0 Then
            For Each Dr As DataRow In DtClientes347Det.Select
                If Length(Nz(StrCifDet)) > 0 Then
                    StrCifDet = Strings.Left(Nz(Dr("CifCliente"), ""), 9)
                    StrCifDet = New String(" ", 9 - Length(StrCifDet)) & StrCifDet
                Else : StrCifDet = Strings.Space(9)
                End If
                StrDenominacionDet = Strings.Left(CStr(Nz(TratarSimbolosEspeciales(Dr("RazonSocial")), "")).ToUpper, 40)

                If StrCodPaisParam = "011" Then
                    StrCodPaisDet = "   " 'Código para residentes
                    StrCPProvinciaDet = Strings.Left(Nz(Dr("CodPostal"), "00"), 2)
                Else
                    StrCodPaisDet = FormatearNumeros(StrCodPaisParam, 3)
                    StrCPProvinciaDet = "99"
                End If

                StrRegistro = "2" 'Tipo Registro
                StrRegistro &= "347" 'Modelo Declaracion
                StrRegistro &= Strings.Format(IntAño, "0000") 'Ejercicio
                StrRegistro &= StrCPProvinciaCab 'Codigo provincia
                StrRegistro &= StrCifCab 'NIF Declarante
                StrRegistro &= "B" 'cliente                'Clave Codigo (Entregas de bienes y servicios superiores a 500.000 pts)
                StrRegistro &= StrCifDet 'NIF Declarado
                StrRegistro &= StrDenominacionDet 'Razon Social
                StrRegistro &= Strings.Space(1) 'Representante Legal
                StrRegistro &= StrCPProvinciaDet & StrCodPaisDet 'Codigo Provincia/Pais
                StrRegistro &= Format(Dr("ImpTotalA") * 100, "000000000000000") 'Importe Operaciones
                StrRegistro &= Strings.Space(1) 'Operacion seguro
                StrRegistro &= Strings.Space(1) 'Arrendamiento local negocio
                StrRegistro &= Strings.Space(88) 'Blancos  93-180

                Dim DrNew2 As DataRow = DtFichero.NewRow
                DrNew2("Linea") = StrRegistro
                DtFichero.Rows.Add(DrNew2)
            Next
        End If
        Return DtFichero
    End Function

    Private Function GenerarFichero347Navarra(ByVal IntAño As Integer) As DataTable
        Dim StrRegistro, StrDenominacionEmpresa, StrCifCab, StrCPProvinciaCab, StrCifDet, StrDenominacionDet, StrCPProvinciaDet, StrCodPaisDet, StrCodPaisParam As String
        Dim IntNumRegClientes, IntNumRegs, IntNumRegProveedores As Integer
        Dim DblImpClientes, DblImpProveedores As Double

        'Creamos el datatable que contendrá las líneas del fichero
        Dim DtFichero As New DataTable
        DtFichero.Columns.Add("Linea", GetType(String))

        'Obtenemos los parametros para crear el fichero
        Dim DtParametros As DataTable = New Parametro().Filter(, "left(IDParametro,3)='347'")
        If DtParametros Is Nothing OrElse DtParametros.Rows.Count = 0 Then ApplicationService.GenerateError("No se encuentran datos en parámetros para la generación del fichero, por favor, actualícelos.")

        'Obtenemos los Importes Totales para Cientes y Proveedores, en las dos monedas
        Dim DtEmpresas As DataTable = AdminData.Filter("tbDatosEmpresa")
        Dim StrSql As String = "Sum(ImpTotalA) as ImporteTotalA, Count(*) as NumRegistros"
        Dim DtClientes347Tot As DataTable = AdminData.Filter("VRptInformeTesoreria347Clientes", StrSql, "Año=" & IntAño, , False)
        StrSql = "Sum(ImpTotalA) as ImporteTotalA, Count(*)as NumRegistros"
        Dim DtProveedores347Tot As DataTable = AdminData.Filter("VRptInformeTesoreria347Proveedores", StrSql, "Año=" & IntAño, , False)

        Dim DrSel() As DataRow = DtParametros.Select("IDParametro='347CODPAIS'")
        If DrSel.Length > 0 Then StrCodPaisParam = Nz(DrSel(0)("Valor"), "   ")

        DrSel = DtParametros.Select("IDParametro='347Moneda'")
        If DrSel.Length > 0 Then
            IntNumRegClientes = Nz(DrSel(0)("NumRegistros"), 0)
            DblImpClientes = Nz(DrSel(0)("ImporteTotalA"), 0)
        End If

        DrSel = DtParametros.Select("IDParametro='347Moneda'")
        If DrSel.Length > 0 Then
            IntNumRegProveedores = Nz(DrSel(0)("NumRegistros"), 0)
            DblImpProveedores = Nz(DrSel(0)("ImporteTotalA"), 0)
        End If

        IntNumRegs = IntNumRegProveedores + IntNumRegClientes

        'Registro 1: Declarante

        StrCPProvinciaCab = Strings.Left(Nz(DtEmpresas.Rows(0)("CodPostal"), ""), 2)
        If Length(Nz(DtEmpresas.Rows(0)("Cif"), "")) > 0 Then
            StrCifCab = Strings.Left(Nz(DtEmpresas.Rows(0)("Cif"), ""), 9)
            StrCifCab = New String(" ", 9 - Length(StrCifCab)) & StrCifCab
        Else : StrCifCab = Strings.Space(9)
        End If
        StrDenominacionEmpresa = Strings.Left(CStr(Nz(DtEmpresas.Rows(0)("DescEmpresa"), "")).ToUpper, 40)
        StrDenominacionEmpresa = TratarSimbolosEspeciales(StrDenominacionEmpresa)

        StrRegistro = "1" 'Tipo de Registro
        StrRegistro &= "347" 'Modelo Declaracion
        StrRegistro &= Strings.Format(IntAño, "0000") 'Ejercicio
        StrRegistro &= StrCifCab 'NIF Declarante
        StrRegistro &= StrDenominacionEmpresa 'Razon Social del Declarante

        DrSel = DtParametros.Select("IDParametro='347TSOPORT'")
        If DrSel.Length > 0 Then
            StrRegistro &= Nz(Strings.Left(DrSel(0)("Valor"), 1), "D")
        End If

        DrSel = DtParametros.Select("IDParametro='347TLFCONT'")
        If DrSel.Length > 0 Then  'Persona con quien relacionarse:Telefono
            StrRegistro &= Nz(Strings.Format(DrSel(0)("Valor"), "000000000"), "000000000")
        End If

        DrSel = DtParametros.Select("IDParametro='347NOMCONT'")
        If DrSel.Length > 0 Then 'Persona con quien relacionarse:Nombre y apellidos
            StrRegistro &= Nz(Strings.Left(DrSel(0)("Valor"), 40), "SIN DEFINIR                             ")
            If Length(DrSel(0)("Valor") & String.Empty) < 40 Then StrRegistro &= New String(" ", 40 - Length(DrSel(0)("Valor") & String.Empty))
        End If

        StrRegistro &= "348" & New String("0", 10) 'Num de Declaracion

        DrSel = DtParametros.Select("IDParametro='347DECSUST'")  'Complementaria o sustitutiva
        If DrSel.Length > 0 Then 'no ofrecemos posibilidad de declaracion complementaria
            StrRegistro &= IIf(DrSel(0)("Valor"), " S", "  ")
        End If

        StrRegistro &= New String("0", 13) 'Num de Declaracion anterior
        StrRegistro &= Strings.Format(IntNumRegs, "000000000") 'Num total de entidades
        StrRegistro &= Format((DblImpClientes + DblImpProveedores) * 100, "000000000000000") 'Importe total de las operaciones
        StrRegistro &= New String("0", 9) 'Num total de inmuebles
        StrRegistro &= New String("0", 15) 'Importe total de las op de arrendamiento
        StrRegistro &= Strings.Space(54) 'Blancos
        StrRegistro &= Strings.Space(13) 'Sello electronico

        Dim DrNew As DataRow = DtFichero.NewRow
        DrNew("Linea") = StrRegistro
        DtFichero.Rows.Add(DrNew)

        'Registro 2: Operacion
        Dim DtClientes347Det As DataTable = AdminData.Filter("VRptInformeTesoreria347Clientes", , "Año=" & IntAño, , False)
        Dim DtProveedores347Det As DataTable = AdminData.Filter("VRptInformeTesoreria347Proveedores", , "Año=" & IntAño, , False)

        If Not DtProveedores347Det Is Nothing AndAlso DtProveedores347Det.Rows.Count > 0 Then
            For Each Dr As DataRow In DtProveedores347Det.Select
                If Length(Nz(Dr("CifProveedor"), "")) > 0 Then
                    StrCifDet = Strings.Left(Nz(Dr("CifProveedor"), ""), 9)
                    StrCifDet = New String(" ", 9 - Length(StrCifDet)) & StrCifDet
                Else : StrCifDet = Strings.Space(9)
                End If
                StrDenominacionDet = Strings.Left(CStr(Nz(Dr("RazonSocial"), "")).ToUpper, 40)
                StrDenominacionDet = TratarSimbolosEspeciales(StrDenominacionDet)

                If StrCodPaisParam = "011" Then
                    StrCodPaisDet = "   "
                    StrCPProvinciaDet = Strings.Left(Nz(Dr("CodPostal"), "00"), 2)
                Else
                    StrCodPaisDet = StrCodPaisParam
                    StrCPProvinciaDet = "99"
                End If

                StrRegistro = "2" 'Tipo Registro
                StrRegistro &= "347" 'Modelo Declaracion
                StrRegistro &= Strings.Format(IntAño, "0000") 'Ejercicio
                StrRegistro &= StrCifCab 'NIF Declarante
                StrRegistro &= StrCifDet 'NIF Declarado
                StrRegistro &= New String(" ", 9) 'NIF Representante Legal
                StrRegistro &= StrDenominacionDet 'Razon Social
                StrRegistro &= "D" 'Tipo de hoja
                StrRegistro &= StrCPProvinciaDet '& "000"  'Codigo Provincia
                StrRegistro &= StrCodPaisDet 'Codigo Pais
                StrRegistro &= "A" 'proveedor              'Clave Codigo(Adquisiciones de bienes y servicios superiores a 500.000pts)
                StrRegistro &= Format(Dr("ImpTotalA") * 100, "000000000000000") 'Importe Operaciones
                StrRegistro &= Strings.Space(1) 'Operacion seguro
                StrRegistro &= Strings.Space(1) 'Arrendamiento local negocio
                StrRegistro &= Strings.Space(151) 'Blancos  100-250

                Dim DrNew1 As DataRow = DtFichero.NewRow
                DrNew1("Linea") = StrRegistro
                DtFichero.Rows.Add(DrNew1)
            Next
        End If

        If Not DtClientes347Det Is Nothing AndAlso DtClientes347Det.Rows.Count > 0 Then
            For Each Dr As DataRow In DtClientes347Det.Select
                If Length(Nz(StrCifDet)) > 0 Then
                    StrCifDet = Strings.Left(Nz(Dr("CifCliente"), ""), 9)
                    StrCifDet = New String(" ", 9 - Length(StrCifDet)) & StrCifDet
                Else : StrCifDet = Strings.Space(9)
                End If
                StrDenominacionDet = Strings.Left(CStr(Nz(Dr("RazonSocial"), "")).ToUpper, 40)
                StrDenominacionDet = TratarSimbolosEspeciales(StrDenominacionDet)
                If StrCodPaisParam = "011" Then
                    StrCodPaisDet = "   "
                    StrCPProvinciaDet = Strings.Left(Nz(Dr("CodPostal"), "00"), 2)
                Else
                    StrCodPaisDet = StrCodPaisParam
                    StrCPProvinciaDet = "99"
                End If

                StrRegistro = "2" 'Tipo Registro
                StrRegistro &= "347" 'Modelo Declaracion
                StrRegistro &= Strings.Format(IntAño, "0000") 'Ejercicio
                StrRegistro &= StrCifCab 'NIF Declarante
                StrRegistro &= StrCifDet 'NIF Declarado
                StrRegistro &= New String(" ", 9) 'NIF Representante legal
                StrRegistro &= StrDenominacionDet 'Razon Social
                StrRegistro &= "D" 'Tipo de hoja
                StrRegistro &= StrCPProvinciaDet '& "000"  'Codigo Provincia
                StrRegistro &= StrCodPaisDet
                StrRegistro &= "B" 'Clave Codigo (Entregas de bienes y servicios superiores a 500.000 pts)
                StrRegistro &= Format(Dr("ImpTotalA") * 100, "000000000000000") 'Importe Total
                StrRegistro &= Strings.Space(1) 'Operacion seguro
                StrRegistro &= Strings.Space(1) 'Arrendamiento local negocio
                StrRegistro &= Strings.Space(151) 'Blancos  100-250

                Dim DrNew2 As DataRow = DtFichero.NewRow
                DrNew2("Linea") = StrRegistro
                DtFichero.Rows.Add(DrNew2)
            Next
        End If
        Return DtFichero
    End Function

#End Region

#Region "Ficheros 349 "

    Public Function GenerarFichero349(ByVal IntAño As Integer, ByVal IntTrimestre As Integer, ByVal StrNumDesde As String, ByVal StrNumHasta As String) As DataTable
        Dim StrRegistro, StrDenominacionCab, StrCifCab, StrCPProvinciaCab, StrCifDet, _
        StrDenominacionDet, StrCPProvinciaDet, StrTelContacto, StrPersonaContacto, StrPeriodo, StrDecSust As String
        Dim DblImporteTotal As Double
        Dim IntFile, IntNumRegistros As Integer

        'Creamos el datatable que contendrá las líneas del fichero
        Dim DtFichero As New DataTable
        DtFichero.Columns.Add("Linea", GetType(String))

        'Saco los datos de la empresa para la Cabecera
        Dim DtEmpresas As DataTable = AdminData.Filter("tbDatosEmpresa")

        'Saco el número de registros y el importe total de los clientes para la Cabecera
        Dim StrSql As String = "Sum(BaseImponibleA) as ImporteTotalA,Count(CifCliente) as NumRegistros"
        Dim StrWhere As String = "NDeclaracionIVA >='" & StrNumDesde & "' and NDeclaracionIVA <='" & StrNumHasta & "' and Año=" & IntAño
        Dim DtClientes349Tot As DataTable = AdminData.Filter("VRptInformeTesoreria349Clientes", StrSql, StrWhere, , False)

        'Saco el número de registros y el importe total de los proveedores para a Cabecera
        StrWhere = "NDeclaracionIVA >='" & StrNumDesde & "' and NDeclaracionIVA <='" & StrNumHasta & "' and Año=" & IntAño & ""
        StrSql = "Sum(BaseImponibleA) as ImporteTotalA,Count(CifProveedor) as NumRegistros"
        Dim DtProveedores349Tot As DataTable = AdminData.Filter("VRptInformeTesoreria349Proveedores", StrSql, StrWhere, , False)

        If Not DtClientes349Tot Is Nothing AndAlso DtClientes349Tot.Rows.Count > 0 Then
            IntNumRegistros = Nz(DtClientes349Tot.Rows(0)("NumRegistros"), 0)
            DblImporteTotal = Nz(DtClientes349Tot.Rows(0)("ImporteTotalA"), 0)
        End If

        If Not DtProveedores349Tot Is Nothing AndAlso DtProveedores349Tot.Rows.Count > 0 Then
            IntNumRegistros += Nz(DtProveedores349Tot.Rows(0)("NumRegistros"), 0)
            DblImporteTotal += Nz(DtProveedores349Tot.Rows(0)("ImporteTotalA"), 0)
        End If

        'Obtenemos los parametros para crear el fichero
        Dim DtParam As DataTable = New Parametro().Filter(, "left(IDParametro,3)='347'")

        If DtParam Is Nothing OrElse DtParam.Rows.Count = 0 Then
            ApplicationService.GenerateError("No se encuentran datos en parámetros para la generación del fichero, por favor, actualícelos.")
        Else
            Dim DrSel() As DataRow = DtParam.Select("IDParametro='347TLFCONT'")
            If DrSel.Length > 0 Then StrTelContacto = Nz(Strings.Format(DrSel(0)("Valor"), "000000000"), "000000000")
            DrSel = DtParam.Select("IDParametro='347NOMCONT'")
            If DrSel.Length > 0 Then StrPersonaContacto = Nz(Strings.Left(DrSel(0)("Valor"), 40), "SIN DEFINIR                             ")
            DrSel = DtParam.Select("IDParametro='347DECSUST'")
            If DrSel.Length > 0 Then StrDecSust = IIf(DrSel(0)("Valor"), " S", "  ") 'no ofrecemos posibilidad de declaracion complementaria
        End If

        StrPeriodo = IIf(IntTrimestre = 0, IntTrimestre & "A", IntTrimestre & "T")

        'Registro 1: Declarante(Cabecera)
        If Not DtEmpresas Is Nothing AndAlso DtEmpresas.Rows.Count > 0 Then
            StrCPProvinciaCab = Strings.Left(Nz(DtEmpresas.Rows(0)("CodPostal")), 2)
            StrCifCab = Strings.Left(Nz(DtEmpresas.Rows(0)("Cif"), ""), 9)
            StrCifCab = New String(" ", 9 - Length(StrCifCab)) & StrCifCab
            StrDenominacionCab = Strings.Left(CStr(Nz(DtEmpresas.Rows(0)("DescEmpresa"))).ToUpper, 40)
            StrDenominacionCab = TratarSimbolosEspeciales(StrDenominacionCab)
        End If

        'Registro 1: Registro Declarante
        StrRegistro = "1"
        StrRegistro &= "349"
        StrRegistro &= Strings.Format(IntAño, "0000") 'Ejercicio
        StrRegistro &= StrCifCab 'Nif Declarante
        StrRegistro &= StrDenominacionCab 'Razon Social Declarante
        StrRegistro &= "D" 'Tipo de Soporte
        StrRegistro &= StrTelContacto & StrPersonaContacto 'Persona con quien relacioarese
        StrRegistro &= "343" & New String("0", 10) 'Num justificante
        StrRegistro &= StrDecSust 'Declaracion Complementaria
        StrRegistro &= New String("0", 13) 'Num Justificante de la decl anterior
        StrRegistro &= StrPeriodo 'Periodo
        StrRegistro &= Strings.Format(IntNumRegistros, "000000000") 'Num de operadores intracomunitarios
        StrRegistro &= Strings.Format(xRound(DblImporteTotal * 100, 0), "000000000000000") 'Importe de operaciones intracomunitarios
        StrRegistro &= New String("0", 9) 'Num de operadores con rectificacines
        StrRegistro &= New String("0", 15) 'Importe de las rectificacines
        StrRegistro &= Strings.Space(52) 'Blancos
        StrRegistro &= Strings.Space(13) 'Sello electronico

        Dim DrNew As DataRow = DtFichero.NewRow
        DrNew("Linea") = StrRegistro
        DtFichero.Rows.Add(DrNew)

        'Registro 2: Registro de operador intracomunitario
        StrSql = "Min(Cifproveedor) as CifProveedor,Sum(BaseImponibleA) as ImporteTotalA,Min(RazonSocial)as RazonSocial,Min(CodPostal)as CodPostal"
        StrWhere = "NDeclaracionIVA >='" & StrNumDesde & "' and NDeclaracionIVA <='" & StrNumHasta & "' and Año=" & IntAño & " GROUP BY CifProveedor"
        Dim DtProveedores349Det As DataTable = AdminData.Filter("VRptInformeTesoreria349Proveedores", StrSql, StrWhere, , False)
        If Not DtProveedores349Det Is Nothing AndAlso DtProveedores349Det.Rows.Count > 0 Then
            For Each Dr As DataRow In DtProveedores349Det.Select
                StrCifDet = Left(Nz(Dr("CifProveedor")), 17)
                StrCifDet = StrCifDet & New String(" ", 17 - Len(StrCifDet))
                StrDenominacionDet = Strings.Left(CStr(Nz(Dr("RazonSocial"))).ToUpper, 40)
                StrDenominacionDet = TratarSimbolosEspeciales(StrDenominacionDet)

                StrRegistro = "2" 'Tipo de Registro
                StrRegistro &= "349" 'Modelo de
                StrRegistro &= Strings.Format(IntAño, "0000") 'Ejercicio
                StrRegistro &= StrCifCab 'Nif declarante
                StrRegistro &= Strings.Space(58) 'Blancos
                StrRegistro &= StrCifDet 'Nif operador comunitario
                StrRegistro &= StrDenominacionDet 'Razon social
                StrRegistro &= "A" 'Clave operacion: Adquisicion
                StrRegistro &= Strings.Format(xRound(Dr("ImporteTotalA") * 100, 0), "0000000000000") 'Base Imponible
                StrRegistro &= Strings.Space(104) 'Blancos

                Dim DrNew1 As DataRow = DtFichero.NewRow
                DrNew1("Linea") = StrRegistro
                DtFichero.Rows.Add(DrNew1)
            Next
        End If

        StrSql = "Min(CifCliente) as CifCliente,Sum(BaseImponibleA) as ImporteTotalA,Min(RazonSocial)as RazonSocial,Min(CodPostal)as CodPostal"
        StrWhere = "NDeclaracionIVA >='" & StrNumDesde & "' and NDeclaracionIVA <='" & StrNumHasta & "' and Año=" & IntAño & " GROUP BY CifCliente"
        Dim DtClientes349Det As DataTable = AdminData.Filter("VRptInformeTesoreria349Clientes", StrSql, StrWhere, , False)

        If Not DtClientes349Det Is Nothing AndAlso DtClientes349Det.Rows.Count > 0 Then
            For Each Dr As DataRow In DtClientes349Det.Select
                StrCifDet = Strings.Left(Nz(Dr("CifCliente")), 17)
                StrCifDet = StrCifDet & New String(" ", 17 - Length(StrCifDet))
                StrDenominacionDet = Strings.Left(CStr(Nz(Dr("RazonSocial"))).ToUpper, 40)
                StrDenominacionDet = TratarSimbolosEspeciales(StrDenominacionDet)

                StrRegistro = "2" 'Tipo de Registro
                StrRegistro &= "349" 'Modelo de
                StrRegistro &= Strings.Format(IntAño, "0000") 'Ejercicio
                StrRegistro &= StrCifCab 'Nif declarante
                StrRegistro &= Strings.Space(58) 'Blancos
                StrRegistro &= StrCifDet 'Nif operador comunitario
                StrRegistro &= StrDenominacionDet 'Razon social
                StrRegistro &= "E" 'Clave operacion: Adquisicion
                StrRegistro &= Strings.Format(xRound(Dr("ImporteTotalA") * 100, 0), "0000000000000") 'Base Imponible
                StrRegistro &= Strings.Space(104) 'Blancos

                Dim DrNew2 As DataRow = DtFichero.NewRow
                DrNew2("Linea") = StrRegistro
                DtFichero.Rows.Add(DrNew2)
            Next
        End If
        Return DtFichero
    End Function

#End Region

#Region "Otras Declaraciones"

    <Task()> Public Shared Function GenerarFichero68(ByVal DatosConf As DataFicheros, ByVal services As ServiceProvider) As DataTable
        Dim DtPagos As DataTable = New Parametro().SelOnPrimaryKey("CONT_PC")
        If Not DtPagos Is Nothing AndAlso DtPagos.Rows.Count > 0 Then
            Dim strContador As String = DtPagos.Rows(0)("Valor")
            If Length(strContador) > 0 Then
                Dim DatosCont As Contador.CounterTx = ProcessServer.ExecuteTask(Of String, Contador.CounterTx)(AddressOf Contador.CounterValueTx, strContador, services)
                Dim strContPagosCert As String = DatosCont.strCounterValue 'Se guarda el valor del contador
                Dim DtCounter As DataTable = DatosCont.DtCounter  'datatable para actualizar el contador
                Dim DtPagosMarca As DataTable = New BE.DataEngine().Filter("frmPagoContGenerarFich68", New GuidFilterItem("IDProcess", DatosConf.IDProcess))
                If Not DtPagosMarca Is Nothing AndAlso DtPagosMarca.Rows.Count > 0 Then
                    Dim strIN As String
                    Dim strInDesglose As String
                    For Each Dr As DataRow In DtPagosMarca.Select
                        If Length(strIN) > 0 Then
                            strIN &= ", " & Dr("IdPago")
                        Else : strIN = Dr("IdPago")
                        End If
                    Next
                    If Length(strIN) > 0 Then
                        strInDesglose = "IDPagoAgrupado IN (" & strIN & ")"
                        strIN = "IdPago IN (" & strIN & ")"
                        DtPagos = New Pago().Filter(, strIN)
                        If Not DtPagos Is Nothing AndAlso DtPagos.Rows.Count > 0 Then
                            'Se recorren los pagos marcados y se actualiza cn el valor del contador
                            For Each DrP As DataRow In DtPagos.Select
                                DrP("NOperacion") = strContPagosCert
                            Next
                            Dim F As New GenerarFicheros
                            GenerarFichero68 = F.GenerarDt68(DtPagosMarca, DatosConf.IDBancoPropio, strInDesglose, CInt(strContPagosCert))
                            If Not GenerarFichero68 Is Nothing Then BusinessHelper.UpdateTable(DtPagos)
                        End If
                    Else
                        ApplicationService.GenerateError("No hay pagos seleccionados. No se generará el fichero.")
                    End If
                End If
            End If
        End If
    End Function

    <Task()> Public Shared Function GenerarFichero341(ByVal Datos341 As DataFicheros, ByVal services As ServiceProvider) As DataTable
        Dim StrChequeOTransf, StrRefBeneficiario, StrConcepto, StrImporte, StrImporteTotalGeneral As String
        Dim IntRestoDireccion, IntRegistros033, IntRegTotalesTrans, IntRegistros010, IntRegTotalesNacional, _
        IntRegistros043, IntRegTotalesEspecial, IntRegTotalGeneral, IntRegistrosNumeroTotal As Integer
        Dim DblImporteTotalNacional, DblImporteTotalTrans, DblImporteTotalEspecial, DblImporteTotalGeneral As Double

        Dim Long_CIF As Integer = 9
        Dim Formato_Fecha As String = "ddMMyy"
        Dim Long_Nombre As Integer = 36
        Dim Long_Linea As Integer = 72
        Dim Long_Poblacion As Integer = 31

        'Creamos el datatable del fichero Final
        Dim DtFichero As New DataTable
        If Length(Datos341.FechaEmision) = 0 Then Datos341.FechaEmision = cnMinDate
        DtFichero.RemotingFormat = SerializationFormat.Binary
        DtFichero.Columns.Add("Linea", GetType(String))
        'Creamos el datatable del fichero Nacional
        Dim DtFichNacional As DataTable = DtFichero.Clone
        'Creamos el datatable del fichero Transfronterizo
        Dim DtFichTrans As DataTable = DtFichero.Clone
        'Creamos el datatable del fichero Especial
        Dim DtFichEspecial As DataTable = DtFichero.Clone

        'Cogemos los datos de la Empresa
        Dim DtEmpresa As DataTable = AdminData.Filter("tbDatosEmpresa")
        Dim strEmpresaRegistro As String = Strings.Left(CStr(DtEmpresa.Rows(0)("Cif")), Long_CIF)
        If Length(strEmpresaRegistro) > Long_CIF Then
            strEmpresaRegistro = Left(strEmpresaRegistro, Long_CIF)
        Else
            strEmpresaRegistro = strEmpresaRegistro & Space(Long_CIF - Length(strEmpresaRegistro))
        End If
        Dim strDescEmpresa As String = Strings.Left(CStr(DtEmpresa.Rows(0)("DescEmpresa")), Long_Nombre)
        If Length(strDescEmpresa) > Long_Nombre Then
            strDescEmpresa = Left(strDescEmpresa, Long_Nombre)
        Else
            strDescEmpresa = strDescEmpresa & Space(Long_Nombre - Length(strDescEmpresa))
        End If
        Dim strDirEmpresa As String = Strings.Left(CStr(DtEmpresa.Rows(0)("Direccion")), Long_Nombre)
        If Length(strDirEmpresa) > Long_Nombre Then
            strDirEmpresa = Left(strDirEmpresa, Long_Nombre)
        Else
            strDirEmpresa = strDirEmpresa & Space(Long_Nombre - Length(strDirEmpresa))
        End If

        'Cogemos los Pagos seleccionados
        Dim DtPagos As DataTable = New BE.DataEngine().Filter("frmPagosGenerarFichero341", New GuidFilterItem("IDProcess", Datos341.IDProcess))
        If Not DtPagos Is Nothing AndAlso DtPagos.Rows.Count = 0 Then
            ApplicationService.GenerateError("No hay pagos seleccionados. No se generará el fichero.")
            Exit Function
        End If

        'Cogemos los datos del BancoPropio
        Dim StDatosBanco As New EstDatosBancoFicheros
        StDatosBanco.IDBanco = Datos341.IDBancoPropio
        StDatosBanco = ProcessServer.ExecuteTask(Of EstDatosBancoFicheros, EstDatosBancoFicheros)(AddressOf DatosBanco, StDatosBanco, services)

        'Primer Registro de CABECERA Obligatorio
        Dim StrRegistro As String = "0362"
        StrRegistro &= strEmpresaRegistro
        StrRegistro &= StDatosBanco.Sufijo
        StrRegistro &= Space(12)
        StrRegistro &= "001"
        StrRegistro &= Format(Today, Formato_Fecha) & IIf(Datos341.FechaEmision <> cnMinDate, Format(Datos341.FechaEmision, Formato_Fecha), Format(Today, Formato_Fecha))
        StrRegistro &= StDatosBanco.Entidad & StDatosBanco.Sucursal & StDatosBanco.DC & StDatosBanco.NCuenta & "0" & Space(8)
        If Length(StrRegistro) < Long_Linea Then StrRegistro = StrRegistro & (Space(Long_Linea - Len(StrRegistro)))
        Dim DrNew As DataRow = DtFichero.NewRow
        DrNew("Linea") = StrRegistro
        DtFichero.Rows.Add(DrNew)
        IntRegTotalGeneral += 1

        'Segundo Registro CABECERA Obligatorio
        StrRegistro = "0362"
        StrRegistro &= strEmpresaRegistro
        StrRegistro &= StDatosBanco.Sufijo
        StrRegistro &= Space(12)
        StrRegistro &= "002"
        StrRegistro &= strDescEmpresa & Strings.Space(5)
        If Len(StrRegistro) < Long_Linea Then StrRegistro &= (Strings.Space(Long_Linea - Length(StrRegistro)))
        Dim DrNew2 As DataRow = DtFichero.NewRow
        DrNew2("Linea") = StrRegistro
        DtFichero.Rows.Add(DrNew2)
        IntRegTotalGeneral += 1

        'Tercer Registro CABECERA Obligatorio
        StrRegistro = "0362"
        StrRegistro &= strEmpresaRegistro
        StrRegistro &= StDatosBanco.Sufijo
        StrRegistro &= Strings.Space(12)
        StrRegistro &= "003"
        StrRegistro &= strDirEmpresa & Strings.Space(5)
        If Length(StrRegistro) < Long_Linea Then StrRegistro &= (Strings.Space(Long_Linea - Length(StrRegistro)))
        Dim DrNew3 As DataRow = DtFichero.NewRow
        DrNew3("Linea") = StrRegistro
        DtFichero.Rows.Add(DrNew3)
        IntRegTotalGeneral += 1

        'Cuarto Registro CABECERA Obligatorio
        StrRegistro = "0362"
        StrRegistro &= strEmpresaRegistro
        StrRegistro &= StDatosBanco.Sufijo
        StrRegistro &= Strings.Space(12)
        StrRegistro &= "004"
        StrRegistro &= Format(CInt(Nz(DtEmpresa.Rows(0)("CodPostal"), 0)), "00000")
        Dim strPoblacionEmp As String = Left(DtEmpresa.Rows(0)("Poblacion") & String.Empty, Long_Poblacion)
        If Length(strPoblacionEmp) > Long_Poblacion Then
            strPoblacionEmp = Left(strPoblacionEmp, Long_Poblacion)
        Else
            strPoblacionEmp = strPoblacionEmp & Space(Long_Poblacion - Length(strPoblacionEmp))
        End If
        StrRegistro &= strPoblacionEmp
        StrRegistro &= Space(5)
        If Len(StrRegistro) < Long_Linea Then StrRegistro &= (Strings.Space(Long_Linea - Length(StrRegistro)))
        Dim DrNew4 As DataRow = DtFichero.NewRow
        DrNew4("Linea") = StrRegistro
        DtFichero.Rows.Add(DrNew4)
        IntRegTotalGeneral += 1

        Dim f As New GenerarFicheros
        For Each Dr As DataRow In DtPagos.Select
            If Nz(Dr("Extranjero"), False) Then
                If Dr("ImpVencimientoA") >= 12500 Then
                    'Registro Especial
                    f.GenerarRegistrosEspeciales(DblImporteTotalEspecial, DtFichEspecial, IntRegistros043, IntRegTotalesEspecial, Dr, strEmpresaRegistro, StDatosBanco.Sufijo)
                Else
                    'Registro Transfronterizo
                    f.GenerarRegistrosTransfronterizos(DblImporteTotalTrans, DtFichTrans, IntRegistros033, IntRegTotalesTrans, Dr, strEmpresaRegistro, StDatosBanco.Sufijo)
                End If
            Else
                'Registro Nacional
                f.GenerarRegistrosNacionales(DblImporteTotalNacional, DtFichNacional, IntRegistros010, IntRegTotalesNacional, Dr, strEmpresaRegistro, StDatosBanco.Sufijo)
            End If
        Next

        If DtFichNacional.Rows.Count <> 0 Then
            f.GenerarBloqueNacional(DtFichero, DtFichNacional, IntRegTotalesNacional, DblImporteTotalNacional, strEmpresaRegistro, StDatosBanco.Sufijo, IntRegistros010)
        End If
        If DtFichTrans.Rows.Count <> 0 Then
            f.GenerarBloqueTransfronterizo(DtFichero, DtFichTrans, IntRegTotalesTrans, DblImporteTotalTrans, strEmpresaRegistro, StDatosBanco.Sufijo, IntRegistros033)
        End If
        If DtFichEspecial.Rows.Count <> 0 Then
            f.GenerarBloqueEspecial(DtFichero, DtFichEspecial, IntRegTotalesEspecial, DblImporteTotalEspecial, strEmpresaRegistro, StDatosBanco.Sufijo, IntRegistros043)
        End If

        'Registros Totales General (Obligatorio)
        StrRegistro = "0962"
        StrRegistro &= strEmpresaRegistro
        StrRegistro &= StDatosBanco.Sufijo
        StrRegistro &= Space(15)
        DblImporteTotalGeneral = DblImporteTotalNacional + DblImporteTotalEspecial + DblImporteTotalTrans
        StrImporteTotalGeneral = Format(DblImporteTotalGeneral * 100, "000000000000")
        IntRegTotalGeneral += IntRegTotalesNacional + IntRegTotalesTrans + IntRegTotalesEspecial
        IntRegistrosNumeroTotal = IntRegistros010 + IntRegistros033 + IntRegistros043
        StrRegistro &= StrImporteTotalGeneral & Strings.Format(IntRegistrosNumeroTotal, "00000000") & Strings.Format(IntRegTotalGeneral + 1, "0000000000")
        StrRegistro &= Strings.Space(11)
        If Len(StrRegistro) < Long_Linea Then StrRegistro &= (Strings.Space(Long_Linea - Length(StrRegistro)))
        Dim DrNew5 As DataRow = DtFichero.NewRow
        DrNew5("Linea") = StrRegistro
        DtFichero.Rows.Add(DrNew5)
        Return DtFichero
    End Function

    Private Function GenerarDt68(ByVal DTPagos As DataTable, ByVal StrIDBancoPropio As String, ByVal StrINWhere As String, ByVal IntNOperacion As Integer) As DataTable
        Dim services As New ServiceProvider
        Dim DtDesglose As DataTable
        Dim DblImporte, DblImporteDesglose, DblImporteTotal As Double
        Dim IntRegistros5301, IntContPagos, IntRefNumCertificado, IntDigitoControl As Integer
        Dim strPagoAgrupado As String = New Parametro().ObtenerPredeterminado("FRAPAGOAGR") & String.Empty
        strPagoAgrupado = TratarSimbolosEspeciales(strPagoAgrupado)

        'Obtenemos los decimales de la moneda
        Dim strFormato As New String("0", 12)
        Dim Long_CIF As Integer = 9
        Dim Long_CIF_Benef As Integer = 12

        'Creamos el datatable del fichero
        Dim DtFichero As New DataTable
        DtFichero.RemotingFormat = SerializationFormat.Binary
        DtFichero.Columns.Add("Linea", GetType(String))

        'Cogemos los datos de la Empresa
        Dim DtEmpresa As DataTable = AdminData.Filter("vDatosEmpresaPais")
        Dim strEmpresaRegistro As String = Strings.Left(CStr(DtEmpresa.Rows(0)("Cif")), Long_CIF)
        If Length(strEmpresaRegistro) > Long_CIF Then
            strEmpresaRegistro = Left(strEmpresaRegistro, Long_CIF)
        Else
            strEmpresaRegistro = strEmpresaRegistro & Space(Long_CIF - Length(strEmpresaRegistro))
        End If

        If Length(StrINWhere) > 0 Then
            Dim DrPagos() As DataRow = DTPagos.Select("NFactura = '" & strPagoAgrupado & "'")
            If DrPagos.Length > 0 Then DtDesglose = AdminData.Filter("frmPagoContGenerarFich68", , StrINWhere)
        End If

        'Cogemos los datos del BancoPropio
        Dim StDatosBancos As New EstDatosBancoFicheros
        StDatosBancos.IDBanco = StrIDBancoPropio
        StDatosBancos = ProcessServer.ExecuteTask(Of EstDatosBancoFicheros, EstDatosBancoFicheros)(AddressOf DatosBanco, StDatosBancos, services)

        'Primer Registro de CABECERA Obligatorio
        Dim Long_Nombre As Integer = 40
        Dim strDescEmpresa As String = Strings.Left(CStr(DtEmpresa.Rows(0)("DescEmpresa")), Long_Nombre)
        strDescEmpresa = TratarSimbolosEspeciales(strDescEmpresa)
        If Length(strDescEmpresa) > Long_Nombre Then
            strDescEmpresa = Left(strDescEmpresa, Long_Nombre)
        Else
            strDescEmpresa = strDescEmpresa & Space(Long_Nombre - Length(strDescEmpresa))
        End If

        Dim StrSufijo As String = String.Empty
        If Not StDatosBancos.Sufijo Is Nothing AndAlso Length(Trim(StDatosBancos.Sufijo)) > 0 Then
            StrSufijo = Strings.Format(CInt(StDatosBancos.Sufijo), "000")
        Else : StrSufijo = "000"
        End If

        'Código del Registro y de la Operación
        Dim StrRegistro As String = "0359"
        StrRegistro &= strEmpresaRegistro
        StrRegistro &= StrSufijo
        StrRegistro &= Strings.Space(12)
        StrRegistro &= "001"
        StrRegistro &= Strings.Format(Today, "ddMMyy")
        StrRegistro &= Strings.Space(9)
        StrRegistro &= StDatosBancos.PrefijoIBAN
        StrRegistro &= StDatosBancos.Entidad & StDatosBancos.Sucursal & StDatosBancos.DC & StDatosBancos.NCuenta
        StrRegistro &= Strings.Space(30)
        Dim DrNew As DataRow = DtFichero.NewRow
        DrNew("Linea") = StrRegistro
        DtFichero.Rows.Add(DrNew)
        Dim IntRegTotales As Integer = 1
        Dim strDH As String
        Dim Long_Nombre_Benef As Integer = 40
        Dim Long_Dir_Benef As Integer = 45
        Dim Long_Pob_Benef As Integer = 40
        Dim Long_Prov_Benef As Integer = 30
        Dim Long_Pais_Benef As Integer = 20
        Dim Long_Linea As Integer = 100

        For Each Dr As DataRow In DTPagos.Select
            Dim StrRefBeneficiario As String = Left(Dr("Cif") & String.Empty, Long_CIF_Benef)
            If Length(StrRefBeneficiario) > Long_CIF_Benef Then
                StrRefBeneficiario = Left(StrRefBeneficiario, Long_CIF_Benef)
            Else
                StrRefBeneficiario = StrRefBeneficiario & Space(Long_CIF_Benef - Length(StrRefBeneficiario))
            End If

            If Dr("ImpVencimientoA") < 0 Then
                strDH = "D"
                strFormato = "0" & Strings.Right(strFormato, Length(strFormato) - 1)
                DblImporte = -Dr("ImpVencimientoA")
            Else
                strDH = "H"
                strFormato = "0" & Strings.Right(strFormato, Length(strFormato) - 1)
                DblImporte = Dr("ImpVencimientoA")
            End If
            Dim StrImporte As String = Format(DblImporte * 100, strFormato)

            'Primer Registro Individual Obligatorio del Beneficiario
            Dim StrDescBeneficiario As String = Left(Dr("DescBeneficiario") & String.Empty, Long_Nombre_Benef)
            StrDescBeneficiario = TratarSimbolosEspeciales(StrDescBeneficiario)
            If Length(StrDescBeneficiario) > Long_Nombre_Benef Then
                StrDescBeneficiario = Left(StrDescBeneficiario, Long_Nombre_Benef)
            Else
                StrDescBeneficiario = StrDescBeneficiario & Space(Long_Nombre_Benef - Length(StrDescBeneficiario))
            End If

            IntRefNumCertificado = Right(Dr("IDPago"), 7)
            IntDigitoControl = ("9000" & IntRefNumCertificado) Mod 7
            StrRegistro = "0659"
            StrRegistro &= strEmpresaRegistro
            StrRegistro &= StrSufijo
            StrRegistro &= StrRefBeneficiario
            StrRegistro &= "010"
            StrRegistro &= StrDescBeneficiario
            StrRegistro &= Strings.Space(29)
            Dim DrNew2 As DataRow = DtFichero.NewRow
            DrNew2("Linea") = StrRegistro
            DtFichero.Rows.Add(DrNew2)
            IntRegTotales += 1

            'Segundo Registro Individual Obligatorio del Beneficiario
            Dim StrDirBeneficiario As String = Left(Dr("DireccionPago") & String.Empty, Long_Dir_Benef)
            StrDirBeneficiario = TratarSimbolosEspeciales(StrDirBeneficiario, False, True)
            If Length(StrDirBeneficiario) > Long_Dir_Benef Then
                StrDirBeneficiario = Left(StrDirBeneficiario, Long_Dir_Benef)
            Else
                StrDirBeneficiario = StrDirBeneficiario & Space(Long_Dir_Benef - Length(StrDirBeneficiario))
            End If
            Dim StrCPBeneficiario As String = FormatearNumeros(Dr("CodPostalPago"), 5)
            Dim StrPobBeneficiario As String = Left(Dr("PoblacionPagoSinCP") & String.Empty, Long_Pob_Benef)
            StrPobBeneficiario = TratarSimbolosEspeciales(StrPobBeneficiario)
            If Length(StrPobBeneficiario) > Long_Pob_Benef Then
                StrPobBeneficiario = Left(StrPobBeneficiario, Long_Pob_Benef)
            Else
                StrPobBeneficiario = StrPobBeneficiario & Space(Long_Pob_Benef - Length(StrPobBeneficiario))
            End If
            Dim StrProvBeneficiario As String = Left(Dr("ProvinciaPago") & String.Empty, Long_Prov_Benef)
            StrProvBeneficiario = TratarSimbolosEspeciales(StrProvBeneficiario)
            If Length(StrProvBeneficiario) > Long_Prov_Benef Then
                StrProvBeneficiario = Left(StrProvBeneficiario, Long_Prov_Benef)
            Else
                StrProvBeneficiario = StrProvBeneficiario & Space(Long_Prov_Benef - Length(StrProvBeneficiario))
            End If
            Dim StrPaisBeneficiario As String = Left(Dr("DescPaisPago") & String.Empty, Long_Pais_Benef)
            StrPaisBeneficiario = TratarSimbolosEspeciales(StrPaisBeneficiario)
            If Length(StrPaisBeneficiario) > Long_Pais_Benef Then
                StrPaisBeneficiario = Left(StrPaisBeneficiario, Long_Pais_Benef)
            Else
                StrPaisBeneficiario = StrPaisBeneficiario & Space(Long_Pais_Benef - Length(StrPaisBeneficiario))
            End If

            'Código de Registro y de Operación
            StrRegistro = "0659"
            StrRegistro &= strEmpresaRegistro
            StrRegistro &= StrSufijo
            StrRegistro &= StrRefBeneficiario
            StrRegistro &= "011"
            StrRegistro &= StrDirBeneficiario
            StrRegistro &= Strings.Space(24)
            Dim DrNew3 As DataRow = DtFichero.NewRow
            DrNew3("Linea") = StrRegistro
            DtFichero.Rows.Add(DrNew3)
            IntRegTotales += 1

            'Tercer Registro Individual Obligatorio del Beneficiario
            StrRegistro = "0659"
            StrRegistro &= strEmpresaRegistro
            StrRegistro &= StrSufijo
            StrRegistro &= StrRefBeneficiario
            StrRegistro &= "012"
            StrRegistro &= StrCPBeneficiario
            StrRegistro &= StrPobBeneficiario
            StrRegistro &= Strings.Space(24)
            Dim DrNew4 As DataRow = DtFichero.NewRow
            DrNew4("Linea") = StrRegistro
            DtFichero.Rows.Add(DrNew4)
            IntRegTotales += 1

            'Cuarto Registro Individual Obligatorio del Beneficiario
            StrRegistro = "0659"
            StrRegistro &= strEmpresaRegistro
            StrRegistro &= StrSufijo
            StrRegistro &= StrRefBeneficiario
            StrRegistro &= "013"
            StrRegistro &= StrCPBeneficiario
            StrRegistro &= Strings.Space(4)
            StrRegistro &= StrProvBeneficiario
            StrRegistro &= StrPaisBeneficiario
            StrRegistro &= Strings.Space(10)
            Dim DrNew5 As DataRow = DtFichero.NewRow
            DrNew5("Linea") = StrRegistro
            DtFichero.Rows.Add(DrNew5)
            IntRegTotales += 1

            'Quinto Registro Individual Obligatorio del Beneficiario
            StrRegistro = "0659"
            StrRegistro &= strEmpresaRegistro
            StrRegistro &= StrSufijo
            StrRegistro &= StrRefBeneficiario
            StrRegistro &= "014"
            StrRegistro &= FormatearNumeros(IntRefNumCertificado, 7) & CStr(IntDigitoControl)
            StrRegistro &= Strings.Format(Dr("FechaVencimiento"), "ddMMyyyy")
            StrRegistro &= StrImporte
            StrRegistro &= "0"
            StrRegistro &= Strings.Space(2)
            StrRegistro &= Strings.Space(6)
            StrRegistro &= Strings.Space(32)
            Dim DrNew6 As DataRow = DtFichero.NewRow
            DrNew6("Linea") = StrRegistro
            DtFichero.Rows.Add(DrNew6)
            IntRegTotales += 1

            'Sexto Registro Individual Obligatorio del Beneficiario
            Dim Long_Factura As Integer = 12
            Dim IntNDato As Integer = 15
            Dim StrSuFactura As String = Strings.Left(Dr("SuFactura") & String.Empty, Long_Factura)
            StrSuFactura = TratarSimbolosEspeciales(StrSuFactura)
            If Length(StrSuFactura) > Long_Factura Then
                StrSuFactura = Left(StrSuFactura, Long_Factura)
            Else
                StrSuFactura = StrSuFactura & Space(Long_Factura - Length(StrSuFactura))
            End If
            StrRegistro = "0659"
            StrRegistro &= strEmpresaRegistro
            StrRegistro &= StrSufijo
            StrRegistro &= StrRefBeneficiario
            StrRegistro &= "015"
            StrRegistro &= FormatearNumeros(IntRefNumCertificado, 7) & CStr(IntDigitoControl)
            Dim StrNFacturaReg As String = Strings.Left(Dr("Factura") & String.Empty, Long_Factura)
            If Length(StrNFacturaReg) > Long_Factura Then
                StrNFacturaReg = Left(StrNFacturaReg, Long_Factura)
            Else
                StrNFacturaReg = StrNFacturaReg & Space(Long_Factura - Length(StrNFacturaReg))
            End If
            StrRegistro &= StrNFacturaReg
            StrRegistro &= Strings.Format(Nz(Dr("FechaFactura"), Dr("FechaVencimiento")), "ddMMyyyy")
            StrRegistro &= StrImporte
            StrRegistro &= strDH
            Dim strNFactura As String
            If Length(Dr("NFactura")) > 0 Then
                If Dr("NFactura") = strPagoAgrupado Then
                    Dim GrupoPago As New EstGrupoPago
                    GrupoPago.IDPago = Dr("IDPago")
                    GrupoPago.DefFichero = enumDefFichero.Fich68
                    strNFactura = ProcessServer.ExecuteTask(Of EstGrupoPago, String)(AddressOf FacturasPagoAgrupado, GrupoPago, services)
                    strNFactura = TratarSimbolosEspeciales(strNFactura)
                Else
                    strNFactura = "S Fra " & Strings.Left(TratarSimbolosEspeciales(Dr("NFactura")), 21)
                End If
            ElseIf Nz(Dr("Texto")) <> String.Empty Then
                strNFactura = TratarSimbolosEspeciales(Nz(Dr("Texto")))
                strNFactura = Strings.Left(strNFactura, 28)
            Else
                strNFactura = String.Empty
            End If
            StrRegistro &= strNFactura
            If StrRegistro.Length < Long_Linea Then StrRegistro &= Strings.Space(Long_Linea - StrRegistro.Length)
            Dim DrNew7 As DataRow = DtFichero.NewRow
            DrNew7("Linea") = StrRegistro
            DtFichero.Rows.Add(DrNew7)
            IntRegTotales += 1
            IntRegistros5301 += 1
            DblImporteTotal += Dr("ImpVencimientoA")
        Next

        'Registro de TOTALES
        If DblImporteTotal < 0 Then
            strFormato = "-" & Right(strFormato, Len(strFormato) - 1)
            DblImporteTotal = -DblImporteTotal
        Else : strFormato = "0" & Strings.Right(strFormato, Length(strFormato) - 1)
        End If
        Dim StrImporteTotal As String = Format(DblImporteTotal * 100, strFormato)

        'Código de Registro y de Operación
        StrRegistro = "0859"
        StrRegistro &= strEmpresaRegistro
        StrRegistro &= StrSufijo
        StrRegistro &= Strings.Space(12)
        StrRegistro &= Strings.Space(3)
        StrRegistro &= StrImporteTotal
        StrRegistro &= Strings.Format(IntRegTotales + 1, New String("0", 10))
        StrRegistro &= Strings.Space(47)
        Dim DrNew8 As DataRow = DtFichero.NewRow
        DrNew8("Linea") = StrRegistro
        DtFichero.Rows.Add(DrNew8)
        Return DtFichero
    End Function

#Region " Remesas Cobros"

    <Serializable()> _
    Public Class DataRemesa
        Public IDBanco As String
        Public FechaEmision As Date
        Public FechaCargo As Date
        Public Sufijo As String
        Public Remesa As Integer
    End Class

    <Task()> Public Shared Function GenerarFichero19(ByVal DatosRemesa As DataRemesa, ByVal services As ServiceProvider) As DataTable
        Dim intRegSoporte As Integer 'Registros del Ordenante: Cabecera, obligatorio y opcionales
        Dim intRegOrdenante As Integer 'Registros de todo el soporte
        Dim intTotalReg As Integer 'Registros del 0680
        Dim dblImporteTotal As Double

        Dim StrNAgrupFact As String = New Parametro().NFacturaCobroAgupado
        Dim Formato_Fecha As String = "ddMMyy"
        Dim Long_CIF As Integer = 12
        Dim Long_Nombre As Integer = 40
        Dim Long_Importe As Integer = 10

        Dim strFormatoImporte As String = New String("0", Long_Importe)

        'Creamos el datatable
        Dim dtFichero As New DataTable
        dtFichero.RemotingFormat = SerializationFormat.Binary
        dtFichero.Columns.Add("Linea", GetType(String))

        'Cogemos los datos de la Empresa
        Dim dtEmpresa As DataTable = AdminData.Filter("tbDatosEmpresa", , , , False)
        If dtEmpresa.Rows.Count = 0 Then
            ApplicationService.GenerateError("No hay Datos de Empresa.")
        End If

        If Length(DatosRemesa.Sufijo) = 0 Then
            DatosRemesa.Sufijo = FormatearNumeros(Nz(DatosRemesa.Sufijo, 0), 3)
        End If
        Dim strEmpresaRegistro As String = dtEmpresa.Rows(0)("Cif") & DatosRemesa.Sufijo
        If Length(strEmpresaRegistro) > Long_CIF Then
            strEmpresaRegistro = Left(strEmpresaRegistro, Long_CIF)
        Else
            strEmpresaRegistro = strEmpresaRegistro & Space(Long_CIF - Length(strEmpresaRegistro))
        End If

        'Cogemos los datos del BancoPropio
        Dim StrDatosBanco As New EstDatosBancoFicheros
        StrDatosBanco.IDBanco = DatosRemesa.IDBanco
        StrDatosBanco = ProcessServer.ExecuteTask(Of EstDatosBancoFicheros, EstDatosBancoFicheros)(AddressOf DatosBanco, StrDatosBanco, services)

        Dim strEntidad As String = StrDatosBanco.Entidad
        Dim strSucursal As String = StrDatosBanco.Sucursal
        Dim strDC As String = StrDatosBanco.DC
        Dim strNCuenta As String = StrDatosBanco.NCuenta

        intRegSoporte = 0

        'REGISTRO CABECERA DEL PRESENTADOR
        Dim strDescEmpresa As String = dtEmpresa.Rows(0)("DescEmpresa") & String.Empty
        If Length(strDescEmpresa) > Long_Nombre Then
            strDescEmpresa = Left(strDescEmpresa, Long_Nombre)
        Else
            strDescEmpresa = strDescEmpresa & Space(Long_Nombre - Length(strDescEmpresa))
        End If


        Dim strRegistro As String = "5180"
        strRegistro = strRegistro & strEmpresaRegistro  'Codigo de Presentador
        strRegistro = strRegistro & Format(DatosRemesa.FechaEmision, Formato_Fecha) 'Fecha Confeccion del soporte
        strRegistro = strRegistro & Space(6) 'Libre
        strRegistro = strRegistro & strDescEmpresa 'Nombre cliente presentador
        strRegistro = strRegistro & Space(20) 'Libre
        strRegistro = strRegistro & strEntidad & strSucursal 'Entidad y sucursal receptora del soporte
        strRegistro = strRegistro & Space(66) 'Libre

        Dim drFichero As DataRow = dtFichero.NewRow
        drFichero("Linea") = strRegistro
        dtFichero.Rows.Add(drFichero)


        intRegSoporte = intRegSoporte + 1
        intRegOrdenante = 0

        'REGISTRO CABECERA EMISOR
        strRegistro = "5380"
        strRegistro = strRegistro & strEmpresaRegistro 'Codigo del Ordenante
        strRegistro = strRegistro & Format(DatosRemesa.FechaEmision, Formato_Fecha) 'Fecha confeccion del soporte
        strRegistro = strRegistro & Format(DatosRemesa.FechaCargo, Formato_Fecha) 'Fecha Negociacion
        strRegistro = strRegistro & strDescEmpresa 'Nombre del cliente ordenante
        strRegistro = strRegistro & strEntidad & strSucursal & strDC & strNCuenta 'CCC del cliente oredenante
        strRegistro = strRegistro & Space(8) & "01" 'Tipo procedimiento adeudo
        strRegistro = strRegistro & Space(64) 'Libre

        drFichero = dtFichero.NewRow
        drFichero("Linea") = strRegistro
        dtFichero.Rows.Add(drFichero)

        intRegSoporte = intRegSoporte + 1
        intRegOrdenante = intRegOrdenante + 1

        'REGISTROS INDIVIDUALES OBLIGATORIOS
        '    strWhere = "IdPrograma='" & strPrograma & "' AND IdUsuario='" & FwnAplicacion.GetSessionInfo.UserName & "'"
        '    Set rcsCobrosMarcados = FwnConsulta.Filter("vfrmRemesaDefinitivaGenFichero", , strWhere)
        Dim dtCobrosMarcados As DataTable = New BE.DataEngine().Filter("vfrmRemesaDefinitivaGenFichero", New NumberFilterItem("IDRemesa", DatosRemesa.Remesa), , "CifCliente")
        If dtCobrosMarcados.Rows.Count > 0 Then
            dblImporteTotal = 0
            For Each drCobro As DataRow In dtCobrosMarcados.Rows
                If Length(drCobro("CifCliente")) < 9 Then
                    ApplicationService.GenerateError("El CIF del cliente tiene menos de nueve dígitos para el cobro {0} - {1}", drCobro("NFactura"), drCobro("RazonSocial"))
                End If
                If Length(drCobro("IDBanco")) = 0 OrElse Len(drCobro("IDBanco") & drCobro("Sucursal") & drCobro("DigitoControl") & drCobro("NCuenta") & String.Empty) <> 20 Then
                    ApplicationService.GenerateError("La cuenta bancaria no tiene 20 digitos para el cobro {0}-{1}", drCobro("NFactura"), drCobro("RazonSocial"))
                End If
                Dim strIDCliente As String = Left(Nz(drCobro("IdClienteBanco"), String.Empty), 12)
                If Length(strIDCliente) > Long_CIF Then
                    strIDCliente = Left(strIDCliente, Long_CIF)
                Else
                    strIDCliente = strIDCliente & Space(Long_CIF - Length(strIDCliente))
                End If

                strDescEmpresa = Left(Nz(drCobro("RazonSocial"), String.Empty), Long_Nombre)
                If Length(strDescEmpresa) > Long_Nombre Then
                    strDescEmpresa = Left(strDescEmpresa, Long_Nombre)
                Else
                    strDescEmpresa = strDescEmpresa & Space(Long_Nombre - Length(strDescEmpresa))
                End If
                Dim strCuentaCorriente As String = FormatearNumeros(Nz(drCobro("IDBanco"), 0), 4)
                strCuentaCorriente &= FormatearNumeros(Nz(drCobro("Sucursal"), 0), 4)
                strCuentaCorriente &= FormatearNumeros(Nz(drCobro("DigitoControl"), 0), 2)
                strCuentaCorriente &= FormatearNumeros(Nz(drCobro("NCuenta"), 0), 10)

                strRegistro = "5680"
                strRegistro = strRegistro & strEmpresaRegistro 'Codigo de Ordenanate
                strRegistro = strRegistro & strIDCliente 'Codigo de Referencia
                strRegistro = strRegistro & strDescEmpresa 'Titular de la domiciliacion
                strRegistro = strRegistro & strCuentaCorriente 'CCC de la Domiciliacion
                strRegistro = strRegistro & Format((drCobro("ImpVencimientoA") + drCobro("ARepercutirA")) * 100, strFormatoImporte) 'Importe de la domiciliacion
                '                strRegistro = strRegistro & Replace(Format$(.Fields("ImpVencimientoA").Value, "00000000.00"), ",", string.empty) 'Importe de la domiciliacion
                strRegistro = strRegistro & Space(6) 'Codigo para Devoluciones
                strRegistro = strRegistro & Space(10) 'Codigo referencia interna
                Dim strFacturas As String
                If Length(drCobro("NFactura")) > 0 Then
                    If drCobro("NFactura") = StrNAgrupFact Then 'Concepto
                        strFacturas = ProcessServer.ExecuteTask(Of String, String)(AddressOf FacturasCobroAgrupado, drCobro("IdCobro"), services)
                    Else
                        strFacturas = Left("FRA Nº " & drCobro("NFactura") & Space(40), 40)
                    End If
                Else
                    strFacturas = Space(40)
                End If
                strRegistro = strRegistro & strFacturas
                strRegistro = strRegistro & Space(8) 'Libre

                drFichero = dtFichero.NewRow
                drFichero("Linea") = strRegistro
                dtFichero.Rows.Add(drFichero)

                dblImporteTotal = dblImporteTotal + drCobro("ImpVencimientoA") + drCobro("ARepercutirA")

                intTotalReg = intTotalReg + 1
                intRegSoporte = intRegSoporte + 1
                intRegOrdenante = intRegOrdenante + 1
            Next
        End If


        'REGISTRO DE TOTAL ORDENANTE
        strRegistro = "5880"
        strRegistro = strRegistro & strEmpresaRegistro 'Codigo de Ordenante
        strRegistro = strRegistro & Space(12) 'Libre
        strRegistro = strRegistro & Space(40) 'Libre
        strRegistro = strRegistro & Space(20) 'Libre
        strRegistro = strRegistro & Format(dblImporteTotal * 100, strFormatoImporte) 'Total importes
        strRegistro = strRegistro & Space(6) 'Libre
        strRegistro = strRegistro & Format(intTotalReg, "0000000000") 'Numero de adeudos
        strRegistro = strRegistro & Format(intRegOrdenante + 1, "0000000000") 'Numero de registros
        strRegistro = strRegistro & Space(38) 'Libre
        drFichero = dtFichero.NewRow
        drFichero("Linea") = strRegistro
        dtFichero.Rows.Add(drFichero)

        intRegSoporte = intRegSoporte + 1

        'REGISTRO DE TOTAL GENERAL
        strRegistro = "5980"
        strRegistro = strRegistro & strEmpresaRegistro 'Codigo del Presentador
        strRegistro = strRegistro & Space(12) 'Libre
        strRegistro = strRegistro & Space(40) 'Libre
        strRegistro = strRegistro & "0001" 'Numero de Ordenantes
        strRegistro = strRegistro & Space(16) 'Libre
        strRegistro = strRegistro & Format(dblImporteTotal * 100, strFormatoImporte) 'Total importes
        strRegistro = strRegistro & Space(6) 'Libre
        strRegistro = strRegistro & Format(intTotalReg, "0000000000") 'Numero de adeudos
        strRegistro = strRegistro & Format(intRegSoporte + 1, "0000000000") 'Numero de registros
        strRegistro = strRegistro & Space(38) 'Libre
        drFichero = dtFichero.NewRow
        drFichero("Linea") = strRegistro
        dtFichero.Rows.Add(drFichero)

        Return dtFichero
    End Function

    <Task()> Public Shared Function GenerarFichero58(ByVal DatosRemesa As DataRemesa, ByVal services As ServiceProvider) As DataTable
        Dim IntTotalReg As Integer
        Dim Formato_Fecha As String = "ddMMyy"
        Dim Long_CIF As Integer = 12
        Dim Long_Nombre As Integer = 40
        Dim Long_Importe As Integer = 10
        Dim Long_IDCobro As Integer = 6 : Dim Long_IDCliente As Integer = 10
        Dim Long_Pob_Cliente As Integer = 35 : Dim Long_Pob_Ordenante As Integer = 38
        Dim Long_Cod_Postal As Integer = 5 : Dim Long_Cod_Provincia As Integer = 2


        'Obtenemos los decimales de la moneda
        Dim strFormatoImporte As String = New String("0", Long_Importe)

        'Cogemos los datos de la Empresa
        Dim dtEmpresa As DataTable = AdminData.Filter("tbDatosEmpresa", , , , False)
        If dtEmpresa.Rows.Count = 0 Then
            ApplicationService.GenerateError("No hay Datos de Empresa.")
        End If

        If Length(DatosRemesa.Sufijo) = 0 Then
            DatosRemesa.Sufijo = FormatearNumeros(Nz(DatosRemesa.Sufijo, 0), 3)
        End If
        Dim strEmpresaRegistro As String = dtEmpresa.Rows(0)("Cif") & DatosRemesa.Sufijo
        If Length(strEmpresaRegistro) > Long_CIF Then
            strEmpresaRegistro = Left(strEmpresaRegistro, Long_CIF)
        Else
            strEmpresaRegistro = strEmpresaRegistro & Space(Long_CIF - Length(strEmpresaRegistro))
        End If

        'Cogemos los datos del BancoPropio
        Dim StrDatosBanco As New EstDatosBancoFicheros
        StrDatosBanco.IDBanco = DatosRemesa.IDBanco
        StrDatosBanco = ProcessServer.ExecuteTask(Of EstDatosBancoFicheros, EstDatosBancoFicheros)(AddressOf DatosBanco, StrDatosBanco, services)

        Dim strEntidad As String = StrDatosBanco.Entidad
        Dim strSucursal As String = StrDatosBanco.Sucursal
        Dim strDC As String = StrDatosBanco.DC
        Dim strNCuenta As String = StrDatosBanco.NCuenta

        'Creamos el datatable donde se almacenan los datos del fichero
        Dim dtFichero As New DataTable
        dtFichero.Columns.Add("Linea", GetType(String))

        'REGISTRO CABECERA PRESENTADOR
        Dim strDescEmpresa As String = dtEmpresa.Rows(0)("DescEmpresa") & String.Empty
        If Length(strDescEmpresa) > Long_Nombre Then
            strDescEmpresa = Left(strDescEmpresa, Long_Nombre)
        Else
            strDescEmpresa = strDescEmpresa & Space(Long_Nombre - Length(strDescEmpresa))
        End If
        Dim strRegistro As String = "5170"
        strRegistro &= strEmpresaRegistro       'CIF + Sufijo
        strRegistro &= Strings.Format(DatosRemesa.FechaEmision, Formato_Fecha)
        strRegistro &= Strings.Space(6)  'Libre
        strRegistro &= strDescEmpresa
        strRegistro &= Strings.Space(20)  'Libre
        strRegistro &= strEntidad & strSucursal
        strRegistro &= Space(66)  'Libre

        Dim drFichero As DataRow = dtFichero.NewRow
        drFichero("Linea") = strRegistro
        dtFichero.Rows.Add(drFichero)

        'REGISTRO DE CABECERA DE ORDENANTE
        strRegistro = "5370"
        strRegistro &= strEmpresaRegistro   'Codigo del Ordenante
        'strRegistro &= Format(Today, Formato_Fecha) 'Fecha confeccion del soporte
        strRegistro &= Format(DatosRemesa.FechaEmision, Formato_Fecha) 'Fecha Emisión
        strRegistro &= Space(6)   'Libre
        strRegistro &= strDescEmpresa 'Nombre del cliente ordenante
        strRegistro &= strEntidad & strSucursal & strDC & strNCuenta 'CCC del cliente oredenante
        strRegistro &= Strings.Space(8) & "06"
        strRegistro &= Space(52) 'Libre
        strRegistro &= New String("0", 9) 'Código INE de la Plaza de Emisión
        strRegistro &= Space(3) 'Libre

        drFichero = dtFichero.NewRow
        drFichero("Linea") = strRegistro
        dtFichero.Rows.Add(drFichero)


        'REGISTROS INDIVIDUALES OBLIGATORIOS
        Dim strImporte As String
        Dim dtCobrosMarcados As DataTable = New BE.DataEngine().Filter("vfrmRemesaDefinitivaGenFichero", New StringFilterItem("IdRemesa", DatosRemesa.Remesa))
        If Not dtCobrosMarcados Is Nothing Then
            IntTotalReg = 0
            For Each drCobrosMarcados As DataRow In dtCobrosMarcados.Select(Nothing, "CifCliente")
                'If Length(drCobrosMarcados("CifCliente")) < 9 Then
                '    ApplicationService.GenerateError("El CIF del cliente no tiene nueve dígitos para el cobro {0}-{1}", drCobrosMarcados("NFactura"), drCobrosMarcados("RazonSocial"))
                'End If
                Dim strCIF As String = Left(drCobrosMarcados("CifCliente") & String.Empty, Long_CIF)
                If Length(strCIF) > Long_CIF Then
                    strCIF = Left(strCIF, Long_CIF)
                Else
                    strCIF = strCIF & Space(Long_CIF - Length(strCIF))
                End If

                strDescEmpresa = Strings.Left(Nz(drCobrosMarcados("RazonSocial")), Long_Nombre)
                If Length(strDescEmpresa) > Long_Nombre Then
                    strDescEmpresa = Left(strDescEmpresa, Long_Nombre)
                Else
                    strDescEmpresa = strDescEmpresa & Space(Long_Nombre - Length(strDescEmpresa))
                End If


                Dim StrCuentaCorriente As String
                If Length(drCobrosMarcados("IDBanco")) > 0 AndAlso Length(drCobrosMarcados("IDBanco") & drCobrosMarcados("Sucursal") & drCobrosMarcados("DigitoControl") & drCobrosMarcados("NCuenta") & String.Empty) <> 20 Then
                    ApplicationService.GenerateError("La cuenta bancaria no tiene 20 digitos para el cobro |-|", drCobrosMarcados("NFactura"), drCobrosMarcados("RazonSocial"))
                Else
                    strEntidad = FormatearNumeros(Nz(drCobrosMarcados("IDBanco"), 0), 4)
                    strSucursal = FormatearNumeros(Nz(drCobrosMarcados("Sucursal"), 0), 4)
                    strDC = FormatearNumeros(Nz(drCobrosMarcados("DigitoControl"), 0), 2)
                    strNCuenta = FormatearNumeros(Nz(drCobrosMarcados("NCuenta"), 0), 10)
                    StrCuentaCorriente = strEntidad & strSucursal & strDC & strNCuenta
                End If

                strImporte = Format((Nz(drCobrosMarcados("ImpVencimientoA"), 0) + Nz(drCobrosMarcados("ARepercutirA"), 0)) * 100, strFormatoImporte)

                Dim strDatosCobro As String = Right(drCobrosMarcados("IdCobro"), Long_IDCobro)
                If Length(strDatosCobro) > Long_IDCobro Then
                    strDatosCobro = Left(strDatosCobro, Long_IDCobro)
                Else
                    strDatosCobro = strDatosCobro & Space(Long_IDCobro - Length(strDatosCobro))
                End If

                Dim StrDatosCliente As String = Right(drCobrosMarcados("IdCliente"), Long_IDCliente)
                If Length(StrDatosCliente) > Long_IDCliente Then
                    StrDatosCliente = Left(StrDatosCliente, Long_IDCliente)
                Else
                    StrDatosCliente = StrDatosCliente & Space(Long_IDCliente - Length(StrDatosCliente))
                End If

                Dim datZonaGrupo As New EstGrupo(drCobrosMarcados("IdCobro"), drCobrosMarcados("NFactura") & String.Empty)
                Dim strPrimerCampoConcepto As String = Format(ProcessServer.ExecuteTask(Of EstGrupo, String)(AddressOf ZonaG, datZonaGrupo, services))
                Dim strFechaVto As String = Format(drCobrosMarcados("FechaVencimiento"), Formato_Fecha)

                strRegistro = "5670"
                strRegistro &= strEmpresaRegistro
                strRegistro &= strCIF
                strRegistro &= strDescEmpresa
                strRegistro &= StrCuentaCorriente
                strRegistro &= strImporte
                strRegistro &= strDatosCobro
                strRegistro &= StrDatosCliente
                strRegistro &= strPrimerCampoConcepto
                strRegistro &= strFechaVto
                strRegistro &= Space(2) 'Libre

                drFichero = dtFichero.NewRow
                drFichero("Linea") = strRegistro
                dtFichero.Rows.Add(drFichero)

                IntTotalReg += 1

                Dim strDomicilio As String = Strings.Left(Nz(drCobrosMarcados("DomicilioDeudor")), Long_Nombre)
                If Length(strDomicilio) > Long_Nombre Then
                    strDomicilio = Left(strDomicilio, Long_Nombre)
                Else
                    strDomicilio = strDomicilio & Space(Long_Nombre - Length(strDomicilio))
                End If
                Dim strPoblacionClte As String = Strings.Left(Nz(drCobrosMarcados("Poblacion")), Long_Pob_Cliente)
                If Length(strPoblacionClte) > Long_Pob_Cliente Then
                    strPoblacionClte = Left(strPoblacionClte, Long_Pob_Cliente)
                Else
                    strPoblacionClte = strPoblacionClte & Space(Long_Pob_Cliente - Length(strPoblacionClte))
                End If

                Dim strCodPostalClte As String = Format(CInt(Nz(drCobrosMarcados("CodPostal"), 0)), "00000")
                Dim strPoblEmpresa As String = Left(dtEmpresa.Rows(0)("Poblacion") & String.Empty, Long_Pob_Ordenante)
                If Length(strPoblEmpresa) > Long_Pob_Ordenante Then
                    strPoblEmpresa = Left(strPoblEmpresa, Long_Pob_Ordenante)
                Else
                    strPoblEmpresa = strPoblEmpresa & Space(Long_Pob_Ordenante - Length(strPoblEmpresa))
                End If
                Dim strCPEmpresa As String = Left(dtEmpresa.Rows(0)("CodPostal") & String.Empty, Long_Cod_Provincia)
                If Length(strCPEmpresa) > Long_Cod_Provincia Then
                    strCPEmpresa = Left(strCPEmpresa, Long_Cod_Provincia)
                Else
                    strCPEmpresa = strCPEmpresa & Space(Long_Cod_Provincia - Length(strCPEmpresa))
                End If

                'REGISTRO OBLIGATORIO DE DOMICILIO (PARA NO DOMICILIADOS)
                strRegistro = "5676"
                strRegistro &= strEmpresaRegistro
                strRegistro &= strCIF
                strRegistro &= strDomicilio
                strRegistro &= strPoblacionClte
                strRegistro &= strCodPostalClte
                strRegistro &= strPoblEmpresa
                strRegistro &= strCPEmpresa
                strRegistro &= Format(DatosRemesa.FechaEmision, Formato_Fecha)
                strRegistro &= Space(8) 'Libre

                drFichero = dtFichero.NewRow
                drFichero("Linea") = strRegistro
                dtFichero.Rows.Add(drFichero)

                IntTotalReg += 1
            Next
        End If

        'REGISTRO DE TOTAL REORDENANTE
        Dim DblTotalRemesa, DblTotalRecibos As Double
        Dim dtTotales As DataTable = New BE.DataEngine().Filter("vfrmRemesaDefinitivaGenFichero", New NumberFilterItem("IDRemesa", DatosRemesa.Remesa), "Count(*) as Cuenta, Sum(ImpVencimientoA) as SumaImpVtoA, sum(ARepercutirA) as SumaARepercutirA")
        If Not dtTotales Is Nothing AndAlso dtTotales.Rows.Count > 0 Then
            DblTotalRemesa = Nz(dtTotales.Rows(0)("SumaImpVtoA"), 0) + Nz(dtTotales.Rows(0)("SumaARepercutirA"), 0)
            DblTotalRecibos = Nz(dtTotales.Rows(0)("Cuenta"), 0)
        End If

        strImporte = Format(DblTotalRemesa * 100, strFormatoImporte)

        strRegistro = "5870"
        strRegistro &= strEmpresaRegistro
        strRegistro &= Strings.Space(12)
        strRegistro &= Strings.Space(40)
        strRegistro &= Strings.Space(20)
        strRegistro &= strImporte
        strRegistro &= Strings.Space(6)
        strRegistro &= Strings.Format(DblTotalRecibos, "0000000000")
        strRegistro &= Strings.Format(IntTotalReg + 2, "0000000000")
        strRegistro &= Strings.Space(38)
        drFichero = dtFichero.NewRow
        drFichero("Linea") = strRegistro
        dtFichero.Rows.Add(drFichero)

        'REGISTRO DE TOTAL GENERAL
        strRegistro = "5970"
        strRegistro &= strEmpresaRegistro
        strRegistro &= Strings.Space(12)
        strRegistro &= Strings.Space(40)
        strRegistro &= "0001"
        strRegistro &= Strings.Space(16)
        strRegistro &= strImporte
        strRegistro &= Strings.Space(6)
        strRegistro &= Strings.Format(DblTotalRecibos, "0000000000")
        strRegistro &= Strings.Format(IntTotalReg + 4, "0000000000")
        strRegistro &= Strings.Space(38)

        drFichero = dtFichero.NewRow
        drFichero("Linea") = strRegistro
        dtFichero.Rows.Add(drFichero)
        Return dtFichero
    End Function

#End Region

#End Region

#Region "Funciones Privadas"

    <Serializable()> _
    Public Class EstDatosBancoFicheros
        Public Entidad As String
        Public Sucursal As String
        Public DC As String
        Public NCuenta As String
        Public Direccion As String
        Public CodClie As String
        Public PrefijoIBAN As String
        Public Sufijo As String
        Public IDBanco As String
        Public CodigoIBAN As String
        Public SWIFT As String
    End Class

    <Task()> Public Shared Function DatosBanco(ByVal EstDatosBanco As EstDatosBancoFicheros, ByVal services As ServiceProvider) As EstDatosBancoFicheros
        Dim DtBancos As DataTable = New BancoPropio().SelOnPrimaryKey(EstDatosBanco.IDBanco)
        If Not DtBancos Is Nothing AndAlso DtBancos.Rows.Count > 0 Then
            Dim DrBanco As DataRow = DtBancos.Rows(0)
            If Length(DrBanco("IDBanco")) = 0 OrElse Length(DrBanco("Sucursal")) = 0 OrElse Length(DrBanco("DigitoControl")) = 0 OrElse Length(DrBanco("NCuenta")) = 0 Then
                ApplicationService.GenerateError("Revise los datos de la Cuenta del Banco Propio {0}.", Quoted(DrBanco("DescBancoPropio")))
            End If
            EstDatosBanco.Entidad = Nz(DrBanco("IDBanco"), 0)
            EstDatosBanco.Sucursal = Nz(DrBanco("Sucursal"), 0)
            EstDatosBanco.DC = Nz(DrBanco("DigitoControl"), 0)
            EstDatosBanco.NCuenta = Nz(DrBanco("NCuenta"), 0)
            EstDatosBanco.Direccion = Nz(DrBanco("Domicilio"), "Sin Direccion.")
            EstDatosBanco.CodClie = Nz(DrBanco("IDClienteConfirming"), Strings.Space(6))
            EstDatosBanco.PrefijoIBAN = Left(Nz(DrBanco("PrefijoIBAN"), "SinIBAN"), 4)
            EstDatosBanco.Sufijo = Nz(DrBanco("SufijoRemesas"), Strings.Space(3))
            EstDatosBanco.CodigoIBAN = FormatearNumeros(DrBanco("CodigoIBAN") & String.Empty, 34)
            EstDatosBanco.SWIFT = FormatearNumeros(DrBanco("SWIFT") & String.Empty, 11)
        Else
            EstDatosBanco.Entidad = "ENTI"
            EstDatosBanco.Sucursal = "SUCU"
            EstDatosBanco.DC = "DC"
            EstDatosBanco.NCuenta = "Sin Cuenta Def."
            EstDatosBanco.Direccion = "Sin Direccion."
            EstDatosBanco.CodClie = Strings.Space(6)
            EstDatosBanco.PrefijoIBAN = Left("SinIBAN", 4)
            EstDatosBanco.Sufijo = Strings.Space(3)
            EstDatosBanco.CodigoIBAN = "Sin CodigoIBAN"
            EstDatosBanco.SWIFT = "Sin SWIFT"
        End If
        EstDatosBanco.Entidad = FormatearNumeros(EstDatosBanco.Entidad, 4)
        EstDatosBanco.Sucursal = FormatearNumeros(EstDatosBanco.Sucursal, 4)
        EstDatosBanco.DC = FormatearNumeros(EstDatosBanco.DC, 2)
        EstDatosBanco.NCuenta = FormatearNumeros(EstDatosBanco.NCuenta, 10)
        'EstDatosBanco.PrefijoIBAN = Format(CInt(EstDatosBanco.PrefijoIBAN), "0000")
        Return EstDatosBanco
    End Function

    'Private Function DatosBanco(ByVal IDBanco As String) As EstDatosBancoFicheros
    '    Dim EstDatosBanco As EstDatosBancoFicheros
    '    Dim DtBancos As DataTable = New BancoPropio().SelOnPrimaryKey(IDBanco)
    '    If Not DtBancos Is Nothing AndAlso DtBancos.Rows.Count > 0 Then
    '        Dim DrBanco As DataRow = DtBancos.Rows(0)
    '        EstDatosBanco.Entidad = Nz(DrBanco("IDBanco"), "ENTI")
    '        EstDatosBanco.Sucursal = Nz(DrBanco("Sucursal"), "SUCU")
    '        EstDatosBanco.DC = Nz(DrBanco("DigitoControl"), "DC")
    '        EstDatosBanco.NCuenta = Nz(DrBanco("NCuenta"), "Sin Cuenta Def.")
    '        EstDatosBanco.Direccion = Nz(DrBanco("Domicilio"), "Sin Direccion.")
    '        EstDatosBanco.CodClie = Nz(DrBanco("IDClienteConfirming"), Strings.Space(6))
    '        EstDatosBanco.PrefijoIBAN = Nz(DrBanco("PrefijoIBAN"), "SinIBAN")
    '        EstDatosBanco.Sufijo = Nz(DrBanco("SufijoRemesas"), Strings.Space(3))
    '    Else
    '        EstDatosBanco.Entidad = "ENTI"
    '        EstDatosBanco.Sucursal = "SUCU"
    '        EstDatosBanco.DC = "DC"
    '        EstDatosBanco.NCuenta = "Sin Cuenta Def."
    '        EstDatosBanco.Direccion = "Sin Direccion."
    '        EstDatosBanco.CodClie = Strings.Space(6)
    '        EstDatosBanco.PrefijoIBAN = "SinIBAN"
    '        EstDatosBanco.Sufijo = Strings.Space(3)
    '    End If
    '    Return EstDatosBanco
    'End Function
    <Serializable()> _
    Public Class EstGrupo
        Public NFactura As String
        Public IDCobro As Integer

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDCobro As Integer, ByVal NFactura As String)
            Me.IDCobro = IDCobro
            Me.NFactura = NFactura
        End Sub
    End Class

    <Task()> Public Shared Function ZonaG(ByVal Grupo As EstGrupo, ByVal services As ServiceProvider) As String
        Dim strConcepto As String
        If Grupo.NFactura = "0" Then
            strConcepto = Strings.Left(ProcessServer.ExecuteTask(Of Integer, String)(AddressOf FacturasCobroAgrupado, Grupo.IDCobro, services), 40)
        Else
            strConcepto = Strings.Left("Nº Factura: " & Grupo.NFactura, 40)
        End If
        Return strConcepto & Space(40 - strConcepto.Length)
    End Function

    <Task()> Public Shared Function FacturasCobroAgrupado(ByVal IDCobro As Integer, ByVal services As ServiceProvider) As String
        Dim strLista As String
        Dim dtCobro As DataTable = New Cobro().Filter(New NumberFilterItem("IDCobroAgrupado", IDCobro), , "NFactura")
        If Not dtCobro Is Nothing Then
            If dtCobro.Rows.Count <> 0 Then
                For Each drCobro As DataRow In dtCobro.Rows
                    If Len(strLista) Then
                        strLista = strLista & "," & drCobro("NFactura") & String.Empty
                    Else
                        strLista = drCobro("NFactura") & String.Empty
                    End If
                Next
            End If
        End If
        FacturasCobroAgrupado = Left("Fras:" & strLista, 40)
    End Function

    <Serializable()> _
    Public Class EstGrupoPago
        Public DefFichero As Short
        Public IDPago As Integer

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDPago As Integer, ByVal DefFichero As Short)
            Me.IDPago = IDPago
            Me.DefFichero = DefFichero
        End Sub
    End Class

    <Task()> Public Shared Function FacturasPagoAgrupado(ByVal GrupoPago As EstGrupoPago, ByVal services As ServiceProvider) As String
        Dim strLista As String
        Dim dtPago As DataTable = New BE.DataEngine().Filter("ViewSuFacturaGenerarFicheros", New NumberFilterItem("IDPagoAgrupado", GrupoPago.IDPago), "SuFactura")
        If Not dtPago Is Nothing Then
            If dtPago.Rows.Count <> 0 Then
                strLista = String.Empty
                For Each drPago As DataRow In dtPago.Rows
                    If Length(drPago("SuFactura")) > 0 Then
                        If Len(strLista) Then
                            strLista = strLista & "," & drPago("SuFactura")
                        Else
                            strLista = drPago("SuFactura")
                        End If
                    End If
                Next
            End If
        End If
        Select Case GrupoPago.DefFichero
            Case enumDefFichero.Fich34
                FacturasPagoAgrupado = "Factura Numero " & strLista
            Case enumDefFichero.Fich68
                FacturasPagoAgrupado = Strings.Left("Fras " & strLista, 28)
        End Select
    End Function

    <Serializable()> _
    Public Class DataRetornoExtraEFichero
        Public Cadena As String
        Public NombreFichero As String

        Public Sub New(ByVal Cadena As String, ByVal NombreFichero As String)
            Me.Cadena = Cadena
            Me.NombreFichero = NombreFichero
        End Sub
    End Class

    <Task()> Public Shared Function ExtraEFichero(ByVal cADENA As String, ByVal services As ServiceProvider) As DataRetornoExtraEFichero
        Dim nPos As Short
        Dim nombrefichero As Object
        Dim ccaracter As String
        Dim i As Short

        cADENA = ProcessServer.ExecuteTask(Of String, String)(AddressOf INVIERTECADENA, cADENA, services)
        nPos = InStr(1, cADENA, "\")
        cADENA = Mid(cADENA, 1, nPos - 1)
        cADENA = ProcessServer.ExecuteTask(Of String, String)(AddressOf INVIERTECADENA, cADENA, services)
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto nombrefichero. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        nombrefichero = ""
        For i = 1 To Len(cADENA)
            ccaracter = Mid(cADENA, i, 1)
            If ccaracter <> "." Then
                'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto nombrefichero. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                nombrefichero = nombrefichero & ccaracter
            Else
                Exit For
            End If
        Next
        If Len(nombrefichero) > 8 Then
            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto nombrefichero. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            nombrefichero = Left(nombrefichero, 8)
        End If
        If Len(nombrefichero) < 8 Then
            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto nombrefichero. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            nombrefichero = nombrefichero & Space(8 - Len(nombrefichero))
        End If
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto nombrefichero. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        Return New DataRetornoExtraEFichero(cADENA, nombrefichero)
    End Function

    <Task()> Public Shared Function INVIERTECADENA(ByVal cADENA As String, ByVal services As ServiceProvider) As String
        Dim CADENAINV As String
        Dim i As Short
        CADENAINV = ""
        For i = 1 To Len(cADENA)
            CADENAINV = CADENAINV & Mid(cADENA, Len(cADENA) + 1 - i, 1)
        Next
        INVIERTECADENA = CADENAINV
    End Function

    Private Sub GenerarRegistrosNacionales(ByRef DblImporteTotalNacional As Double, ByRef DtFichNacional As DataTable, ByRef IntRegistros010 As Integer, ByRef IntRegTotalesNacional As Integer, ByRef Dr As DataRow, ByVal StrEmpresaRegistro As String, ByVal StrSufijo As String)
        Dim services As New ServiceProvider
        Dim StrRegistro, StrImporte, StrConcepto As String
        Dim Long_CIF_Benef As Integer = 12
        Dim StrRefBeneficiario As String = Left(Dr("Cif") & String.Empty, Long_CIF_Benef)
        If Length(StrRefBeneficiario) > Long_CIF_Benef Then
            StrRefBeneficiario = Left(StrRefBeneficiario, Long_CIF_Benef)
        Else
            StrRefBeneficiario = StrRefBeneficiario & Space(Long_CIF_Benef - Length(StrRefBeneficiario))
        End If

        'Primer Registro de Detalle Nacional(Obligatorio)
        StrRegistro = "06"
        If Dr("ChequeTalon") Then
            StrRegistro &= "57"
        ElseIf Dr("Trasferencia") Then
            StrRegistro &= "56"
        Else
            StrRegistro &= Strings.Space(2)
        End If
        StrRegistro &= StrEmpresaRegistro
        StrRegistro &= StrSufijo
        StrImporte = Format(Nz(Dr("ImpVencimientoA"), 0) * 100, "000000000000")
        StrRegistro &= StrRefBeneficiario & "010"
        StrRegistro &= StrImporte
        StrRegistro &= FormatearNumeros(Nz(Dr("IDBancoPago"), 0), 4) & FormatearNumeros(Nz(Dr("SucursalPago"), 0), 4) & FormatearNumeros(Nz(Dr("DCPago"), 0), 2) & FormatearNumeros(Nz(Dr("NCuentaPago"), 0), 10)
        StrRegistro &= "191      "
        If Len(StrRegistro) < 72 Then StrRegistro = StrRegistro & (Space(72 - Len(StrRegistro)))
        Dim DrNew As DataRow = DtFichNacional.NewRow
        DrNew("Linea") = StrRegistro
        DtFichNacional.Rows.Add(DrNew)
        IntRegTotalesNacional += 1
        IntRegistros010 += 1
        DblImporteTotalNacional += Nz(Dr("ImpVencimientoA"), 0)

        'Segundo Registro de Detalle Nacional(Obligatorio)
        StrRegistro = "06"
        If Dr("ChequeTalon") Then
            StrRegistro &= "57"
        ElseIf Dr("Trasferencia") Then
            StrRegistro &= "56"
        Else
            StrRegistro &= Strings.Space(2)
        End If
        StrRegistro &= StrEmpresaRegistro
        StrRegistro &= StrSufijo
        StrRegistro &= StrRefBeneficiario & "011"
        If Length(Dr("DescBeneficiario") & String.Empty) > 36 Then
            StrRegistro &= Strings.Left(Nz(Dr("DescBeneficiario")), 36) & Strings.Space(5)
        Else
            StrRegistro &= Strings.Left(Nz(Dr("DescBeneficiario")), 36) & Strings.Space(36 - Length(Nz(Dr("DescBeneficiario")))) & Strings.Space(5)
        End If
        If Len(StrRegistro) < 72 Then StrRegistro = StrRegistro & (Space(72 - Len(StrRegistro)))
        Dim DrNew2 As DataRow = DtFichNacional.NewRow
        DrNew2("Linea") = StrRegistro
        DtFichNacional.Rows.Add(DrNew2)
        IntRegTotalesNacional += 1

        'Tercer Registro de Detalle Nacional(Obligatorio sin abono directo)
        StrRegistro = "06"
        If Dr("ChequeTalon") Then
            StrRegistro &= "57"
        ElseIf Dr("Trasferencia") Then
            StrRegistro &= "56"
        Else
            StrRegistro &= Strings.Space(2)
        End If
        StrRegistro &= StrEmpresaRegistro
        StrRegistro &= StrSufijo
        StrRegistro &= StrRefBeneficiario & "012"
        If Length(Dr("DireccionPago") & String.Empty) > 36 Then
            StrRegistro &= Left(Nz(Dr("DireccionPago"), String.Empty), 36) & Strings.Space(5)
        Else
            StrRegistro &= Left(Nz(Dr("DireccionPago"), String.Empty), 36) & Strings.Space(36 - Length(Nz(Dr("DireccionPago")))) & Strings.Space(5)
        End If
        If Len(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
        Dim DrNew3 As DataRow = DtFichNacional.NewRow
        DrNew3("Linea") = StrRegistro
        DtFichNacional.Rows.Add(DrNew3)
        IntRegTotalesNacional += 1

        'Quinto Registro de Detalle Nacional(Obligatorio sin abono directo)
        StrRegistro = "06"
        If Dr("ChequeTalon") Then
            StrRegistro &= "57"
        ElseIf Dr("Trasferencia") Then
            StrRegistro &= "56"
        Else : StrRegistro &= Strings.Space(2)
        End If
        StrRegistro &= StrEmpresaRegistro
        StrRegistro &= StrSufijo
        StrRegistro &= StrRefBeneficiario & "014"

        If Length(Dr("CodPostal") & String.Empty) > 5 Then
            StrRegistro &= Strings.Left(Nz(Dr("CodPostal"), String.Empty), 5)
        Else : StrRegistro &= Strings.Left(Nz(Dr("CodPostal"), String.Empty), 5) & Space(5 - Length(Nz(Dr("CodPostal"))))
        End If
        If Length(Dr("PoblacionPago") & String.Empty) > 31 Then
            StrRegistro &= Strings.Left(Nz(Dr("PoblacionPago"), String.Empty), 31) & Strings.Space(5)
        Else : StrRegistro &= Strings.Left(Nz(Dr("PoblacionPago"), String.Empty), 31) & Strings.Space(31 - Length(Nz(Dr("PoblacionPago")))) & Strings.Space(5)
        End If
        If Length(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
        Dim DrNew4 As DataRow = DtFichNacional.NewRow
        DrNew4("Linea") = StrRegistro
        DtFichNacional.Rows.Add(DrNew4)
        IntRegTotalesNacional += 1

        'Séptimo Registro de Detalle Nacional(Obligatorio sin abono directo)
        StrRegistro = "06"
        If Dr("ChequeTalon") Then
            StrRegistro &= "57"
        ElseIf Dr("Trasferencia") Then
            StrRegistro &= "56"
        Else
            StrRegistro &= Strings.Space(2)
        End If
        StrRegistro &= StrEmpresaRegistro
        StrRegistro &= StrSufijo
        StrRegistro &= StrRefBeneficiario & "016"

        'Obtenemos el número de SuFactura o de varias si está agrupado
        Dim StrPagoAgrupado As String = New Parametro().ObtenerPredeterminado("FRAPAGOAGR") & String.Empty

        If Length(Dr("NFactura")) > 0 Then
            If Dr("NFactura") = StrPagoAgrupado Then
                Dim GrupoPago As New EstGrupoPago
                GrupoPago.IDPago = Dr("IDPago")
                GrupoPago.DefFichero = enumDefFichero.Fich34
                StrConcepto &= ProcessServer.ExecuteTask(Of EstGrupoPago, String)(AddressOf FacturasPagoAgrupado, GrupoPago, services)
            Else
                StrConcepto = "S/FRA. " & Dr("SuFactura")
            End If
        ElseIf Nz(Dr("Texto").Value) <> String.Empty Then
            StrConcepto = Dr("Texto")
        Else : StrConcepto = String.Empty
        End If

        If Length(StrConcepto) > 36 Then
            StrRegistro &= Strings.Left(StrConcepto, 36) & Strings.Space(5)
        Else
            StrRegistro &= Strings.Left(StrConcepto, 36) & Strings.Space(36 - Length(StrConcepto)) & Strings.Space(5)
        End If
        If Length(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
        Dim DrNew5 As DataRow = DtFichNacional.NewRow
        DrNew5("Linea") = StrRegistro
        DtFichNacional.Rows.Add(DrNew5)
        IntRegTotalesNacional += 1

        'Séptimo Registro de Detalle Nacional(Obligatorio sin abono directo)
        If Len(StrConcepto) > 36 Then
            StrRegistro = "06"
            If Dr("ChequeTalon") Then
                StrRegistro &= "57"
            ElseIf Dr("Trasferencia") Then
                StrRegistro &= "56"
            Else : StrRegistro &= Strings.Space(2)
            End If
            StrRegistro &= StrEmpresaRegistro
            StrRegistro &= StrSufijo
            StrRegistro &= StrRefBeneficiario & "017"
            If Len(Mid(StrConcepto, 36)) > 36 Then
                StrRegistro &= Strings.Left(Strings.Mid(StrConcepto, 36), 36) & Strings.Space(5)
            Else
                StrRegistro &= Strings.Left(Strings.Mid(StrConcepto, 36), 36) & Strings.Space(36 - Length(Strings.Mid(StrConcepto, 36))) & Strings.Space(5)
            End If
            If Length(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
            Dim DrNew6 As DataRow = DtFichNacional.NewRow
            DrNew6("Linea") = StrRegistro
            DtFichNacional.Rows.Add(DrNew6)
            IntRegTotalesNacional += 1
        End If
    End Sub

    Private Sub GenerarRegistrosTransfronterizos(ByRef DblImporteTotalTrans As Double, ByRef DtFichTrans As DataTable, ByRef IntRegistros033 As Integer, ByRef IntRegTotalesTrans As Integer, ByRef Dr As DataRow, ByVal StrEmpresaRegistro As String, ByVal StrSufijo As String)
        Dim services As New ServiceProvider
        Dim StrRegistro, StrImporte, StrCodigoIBAN, StrPagoAgrupado, StrConcepto As String

        Dim Long_CIF_Benef As Integer = 12
        Dim StrRefBeneficiario As String = Left(Dr("Cif") & String.Empty, Long_CIF_Benef)
        If Length(StrRefBeneficiario) > Long_CIF_Benef Then
            StrRefBeneficiario = Left(StrRefBeneficiario, Long_CIF_Benef)
        Else
            StrRefBeneficiario = StrRefBeneficiario & Space(Long_CIF_Benef - Length(StrRefBeneficiario))
        End If

        'Primer Registro de Detalle Transfronterizo
        StrRegistro = "0660"
        StrRegistro &= StrEmpresaRegistro
        StrRegistro &= StrSufijo
        StrRegistro &= StrRefBeneficiario & "033"
        StrCodigoIBAN = Nz(Dr("CodigoIBAN"))
        StrRegistro &= StrCodigoIBAN & Strings.Space(34 - Length(StrCodigoIBAN)) & "7"
        StrRegistro &= Strings.Space(6)
        If Length(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
        Dim DrNew As DataRow = DtFichTrans.NewRow
        DrNew("Linea") = StrRegistro
        DtFichTrans.Rows.Add(DrNew)
        IntRegTotalesTrans += 1
        IntRegistros033 += 1

        'Segundo Registro de Detalle Transfronterizo
        StrRegistro = "0660"
        StrRegistro &= StrEmpresaRegistro
        StrRegistro &= StrSufijo
        StrRegistro &= StrRefBeneficiario & "034"
        StrImporte = Format(Nz(Dr("ImpVencimientoA"), 0) * 100, "000000000000")
        StrRegistro &= StrImporte
        StrRegistro &= Nz(Dr("ClaveGastos"), 3)
        If Length(Dr("CodigoIBAN") & String.Empty) > 0 Then
            StrRegistro &= Strings.Left(Nz(Dr("CodigoIBAN")), 2)
        Else : StrRegistro &= Strings.Space(2)
        End If
        StrRegistro &= Strings.Space(6)
        If Length(Dr("Swift") & String.Empty) > 0 Then
            StrRegistro &= Strings.Left(Nz(Dr("Swift")), 11)
        Else : StrRegistro &= Strings.Space(11)
        End If
        StrRegistro &= Strings.Space(9)
        If Length(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
        Dim DrNew2 As DataRow = DtFichTrans.NewRow
        DrNew2("Linea") = StrRegistro
        DtFichTrans.Rows.Add(DrNew2)
        IntRegTotalesTrans += 1
        DblImporteTotalTrans += Nz(Dr("ImpVencimientoA"), 0)

        'Tercer Registro de Detalle Transfronterizo
        StrRegistro = "0660"
        StrRegistro &= StrEmpresaRegistro
        StrRegistro &= StrSufijo
        StrRegistro = StrRegistro & StrRefBeneficiario & "035"
        If Length(Dr("DescBeneficiario") & String.Empty) > 36 Then
            StrRegistro &= Strings.Left(Nz(Dr("DescBeneficiario")), 36) & Strings.Space(5)
        Else
            StrRegistro &= Strings.Left(Nz(Dr("DescBeneficiario")), 36) & Strings.Space(36 - Length(Nz(Dr("DescBeneficiario")))) & Strings.Space(5)
        End If
        If Length(StrRegistro) < 72 Then StrRegistro = StrRegistro & (Space(72 - Len(StrRegistro)))
        Dim DrNew3 As DataRow = DtFichTrans.NewRow
        DrNew3("Linea") = StrRegistro
        DtFichTrans.Rows.Add(DrNew3)
        IntRegTotalesTrans += 1

        'Cuarto Registro de Detalle Especial
        StrRegistro = "0660"
        StrRegistro &= StrEmpresaRegistro
        StrRegistro &= StrSufijo
        StrRegistro &= StrRefBeneficiario & "036"
        If Length(Dr("DireccionPago") & String.Empty) > 36 Then
            StrRegistro &= Strings.Left(Nz(Dr("DireccionPago")), 36) & Strings.Space(5)
        Else
            StrRegistro &= Strings.Left(Nz(Dr("DireccionPago")), 36) & Strings.Space(36 - Strings.Len(Nz(Dr("DireccionPago")))) & Strings.Space(5)
        End If
        If Length(StrRegistro) < 72 Then StrRegistro = StrRegistro & (Strings.Space(72 - Strings.Len(StrRegistro)))
        Dim DrNew4 As DataRow = DtFichTrans.NewRow
        DrNew4("Linea") = StrRegistro
        DtFichTrans.Rows.Add(DrNew4)
        IntRegTotalesTrans += 1

        'Quinto Registro de Detalle Especial
        If Length(Dr("DireccionPago") & String.Empty) > 36 Then
            StrRegistro = "0660"
            StrRegistro &= StrEmpresaRegistro
            StrRegistro &= StrSufijo
            StrRegistro &= StrRefBeneficiario & "37"
            If Length(Strings.Mid(Dr("DireccionPago"), 36)) > 36 Then
                StrRegistro &= Strings.Left(Strings.Mid(Nz(Dr("DireccionPago")), 36), 36) & Strings.Space(5)
            Else
                StrRegistro &= Strings.Left(Strings.Mid(Nz(Dr("DireccionPago")), 36), 36) & Strings.Space(36 - Length(Strings.Mid(Nz(Dr("DireccionPago")), 36))) & Strings.Space(5)
            End If
            If Length(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
            Dim DrNew5 As DataRow = DtFichTrans.NewRow
            DrNew5("Linea") = StrRegistro
            DtFichTrans.Rows.Add(DrNew5)
            IntRegTotalesTrans += 1
        End If

        'Sexto Registro de Detalle Especial
        StrRegistro = "0660"
        StrRegistro &= StrEmpresaRegistro
        StrRegistro &= StrSufijo
        StrRegistro &= StrRefBeneficiario & "038"
        If Length(Dr("CodPostal") & String.Empty) > 5 Then
            StrRegistro &= Strings.Left(Nz(Dr("CodPostal")), 5)
        Else : StrRegistro &= Strings.Left(Nz(Dr("CodPostal")), 5) & Strings.Space(5 - Length(Nz(Dr("CodPostal"))))
        End If
        If Length(Dr("PoblacionPago") & String.Empty) > 31 Then
            StrRegistro &= Strings.Left(Nz(Dr("PoblacionPago")), 31) & Strings.Space(5)
        Else
            StrRegistro &= Strings.Left(Nz(Dr("PoblacionPago")), 31) & Strings.Space(31 - Length(Nz(Dr("PoblacionPago")))) & Strings.Space(5)
        End If
        If Len(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
        Dim DrNew6 As DataRow = DtFichTrans.NewRow
        DrNew6("Linea") = StrRegistro
        DtFichTrans.Rows.Add(DrNew6)
        IntRegTotalesTrans += 1

        'Séptimo Registro de Detalle Especial
        StrRegistro = "0660"
        StrRegistro &= StrEmpresaRegistro
        StrRegistro &= StrSufijo
        StrRegistro &= StrRefBeneficiario & "039"
        If Length(Dr("DescPais") & String.Empty) > 36 Then
            StrRegistro &= Strings.Left(Nz(Dr("DescPais")), 36) & Strings.Space(5)
        Else
            StrRegistro &= Strings.Left(Nz(Dr("DescPais")), 36) & Strings.Space(36 - Length(Nz(Dr("DescPais")))) & Strings.Space(5)
        End If
        If Len(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
        Dim DrNew7 As DataRow = DtFichTrans.NewRow
        DrNew7("Linea") = StrRegistro
        DtFichTrans.Rows.Add(DrNew7)
        IntRegTotalesTrans += 1

        'Octavo Registro de Detalle Especial
        StrRegistro = "0660"
        StrRegistro &= StrEmpresaRegistro
        StrRegistro &= StrSufijo
        StrRegistro &= StrRefBeneficiario & "040"

        'Obtenemos el número de SuFactura o de varias si está agrupado
        StrPagoAgrupado = New Parametro().ObtenerPredeterminado("FRAPAGOAGR") & String.Empty

        If Length(Dr("NFactura")) > 0 Then
            If Dr("NFactura") = StrPagoAgrupado Then
                Dim GrupoPago As New EstGrupoPago
                GrupoPago.IDPago = Dr("IDPago")
                GrupoPago.DefFichero = enumDefFichero.Fich34
                StrConcepto &= ProcessServer.ExecuteTask(Of EstGrupoPago, String)(AddressOf FacturasPagoAgrupado, GrupoPago, services)

            Else : StrConcepto = "S/FRA. " & Dr("NFactura")
            End If
        ElseIf Nz(Dr("Texto")) <> String.Empty Then
            StrConcepto = Dr("Texto")
        Else : StrConcepto = String.Empty
        End If

        If Length(StrConcepto) > 36 Then
            StrRegistro &= Strings.Left(StrConcepto, 36) & Strings.Space(5)
        Else
            StrRegistro &= Strings.Left(StrConcepto, 36) & Strings.Space(36 - Length(StrConcepto)) & Strings.Space(5)
        End If
        If Len(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
        Dim DrNew8 As DataRow = DtFichTrans.NewRow
        DrNew8("Linea") = StrRegistro
        DtFichTrans.Rows.Add(DrNew8)
        IntRegTotalesTrans += 1

        'Noveno Registro de Detalle Especial
        If Length(StrConcepto) Then
            StrRegistro = "0660"
            StrRegistro &= StrEmpresaRegistro
            StrRegistro &= StrSufijo
            StrRegistro &= StrRefBeneficiario & "041"
            If Len(Mid(StrConcepto, 36)) > 36 Then
                StrRegistro &= Strings.Left(Strings.Mid(StrConcepto, 36), 36) & Strings.Space(5)
            Else : StrRegistro &= Strings.Left(Strings.Mid(StrConcepto, 36), 36) & Strings.Space(36 - Length(Strings.Mid(StrConcepto, 36))) & Strings.Space(5)
            End If
            If Length(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
            Dim DrNew9 As DataRow = DtFichTrans.NewRow
            DrNew9("Linea") = StrRegistro
            DtFichTrans.Rows.Add(DrNew9)
            IntRegTotalesTrans += 1
        End If
    End Sub

    Private Sub GenerarRegistrosEspeciales(ByRef DblImporteTotalEspecial As Double, ByRef DtFichEspecial As DataTable, ByRef IntRegistros043 As Integer, ByRef IntRegTotalesEspecial As Integer, ByRef Dr As DataRow, ByVal StrEmpresaRegistro As String, ByVal StrSufijo As String)
        Dim StrRegistro, StrImporte, StrCodigoIBAN, StrPagoAgrupado, StrConcepto As String

        Dim Long_CIF_Benef As Integer = 12
        Dim StrRefBeneficiario As String = Left(Dr("Cif") & String.Empty, Long_CIF_Benef)
        If Length(StrRefBeneficiario) > Long_CIF_Benef Then
            StrRefBeneficiario = Left(StrRefBeneficiario, Long_CIF_Benef)
        Else
            StrRefBeneficiario = StrRefBeneficiario & Space(Long_CIF_Benef - Length(StrRefBeneficiario))
        End If

        'Primer Registro de Detalle Especial
        StrRegistro = "0661"
        StrRegistro &= StrEmpresaRegistro
        StrRegistro &= StrSufijo
        StrRegistro &= StrRefBeneficiario & "043"
        StrCodigoIBAN = Nz(Dr("CodigoIBAN"))
        StrRegistro &= StrCodigoIBAN & "7"
        StrRegistro &= Strings.Space(6)
        If Length(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
        Dim DrNew As DataRow = DtFichEspecial.NewRow
        DrNew("Linea") = StrRegistro
        DtFichEspecial.Rows.Add(DrNew)
        IntRegTotalesEspecial += 1
        IntRegistros043 += 1

        'Segundo Registro de Detalle Especial
        StrRegistro = "0661"
        StrRegistro &= StrEmpresaRegistro
        StrRegistro &= StrSufijo
        StrRegistro &= StrRefBeneficiario & "044"
        StrImporte = Format(Nz(Dr("ImpVencimientoA"), 0) * 100, "000000000000")
        StrRegistro &= StrImporte
        StrRegistro &= Nz(Dr("ClaveGastos"), 3)
        Dim Long_IBAN As Integer = 2
        Dim StrCodIBAN As String = Left(Dr("CodigoIBAN") & String.Empty, Long_IBAN)
        If Length(StrCodIBAN) > Long_IBAN Then
            StrCodIBAN = Left(StrCodIBAN, Long_IBAN)
        Else
            StrCodIBAN = StrCodIBAN & Space(Long_IBAN - Length(StrCodIBAN))
        End If
        StrRegistro &= StrCodIBAN
        StrRegistro &= Strings.Space(6)
        Dim Long_Swift As Integer = 11
        Dim StrCodSwift As String = Left(Dr("Swift") & String.Empty, Long_Swift)
        If Length(StrCodSwift) > Long_Swift Then
            StrCodSwift = Left(StrCodSwift, Long_Swift)
        Else
            StrCodSwift = StrCodSwift & Space(Long_Swift - Length(StrCodSwift))
        End If
        StrRegistro &= StrCodSwift
        StrRegistro &= Strings.Space(9)
        If Length(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
        Dim DrNew2 As DataRow = DtFichEspecial.NewRow
        DrNew2("Linea") = StrRegistro
        DtFichEspecial.Rows.Add(DrNew2)
        IntRegTotalesEspecial += 1
        DblImporteTotalEspecial += Nz(Dr("ImpVencimientoA"), 0)

        'Tercer Registro de Detalle Especial
        StrRegistro = "0661"
        StrRegistro &= StrEmpresaRegistro
        StrRegistro &= StrSufijo
        StrRegistro &= StrRefBeneficiario & "045"
        If Length(Dr("DescBeneficiario") & String.Empty) > 36 Then
            StrRegistro &= Strings.Left(Nz(Dr("DescBeneficiario")), 36) & Strings.Space(5)
        Else : StrRegistro &= Strings.Left(Nz(Dr("DescBeneficiario")), 36) & Strings.Space(36 - Length(Nz(Dr("DescBeneficiario")))) & Strings.Space(5)
        End If
        If Length(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
        Dim DrNew3 As DataRow = DtFichEspecial.NewRow
        DrNew3("Linea") = StrRegistro
        DtFichEspecial.Rows.Add(DrNew3)
        IntRegTotalesEspecial += 1

        'Cuarto Registro de Detalle Especial
        StrRegistro = "0661"
        StrRegistro &= StrEmpresaRegistro
        StrRegistro &= StrSufijo
        StrRegistro &= StrRefBeneficiario & "046"
        If Length(Dr("DireccionPago") & String.Empty) > 36 Then
            StrRegistro &= Strings.Left(Nz(Dr("DireccionPago"), String.Empty), 36) & Strings.Space(5)
        Else : StrRegistro &= Strings.Left(Nz(Dr("DireccionPago"), String.Empty), 36) & Strings.Space(36 - Length(Nz(Dr("DireccionPago")))) & Strings.Space(5)
        End If
        If Length(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
        Dim DrNew4 As DataRow = DtFichEspecial.NewRow
        DrNew4("Linea") = StrRegistro
        DtFichEspecial.Rows.Add(DrNew4)
        IntRegTotalesEspecial += 1

        'Quinto Registro de Detalle Especial
        If Length(Dr("DireccionPago") & String.Empty) > 36 Then
            StrRegistro = "0661"
            StrRegistro &= StrEmpresaRegistro
            StrRegistro &= StrSufijo
            StrRegistro &= StrRefBeneficiario & "047"
            If Length(Strings.Mid(Dr("DireccionPago"), 36)) > 36 Then
                StrRegistro &= Strings.Left(Strings.Mid(Dr("DireccionPago"), 36), 36) & Strings.Space(5)
            Else : StrRegistro &= Strings.Left(Strings.Mid(Dr("DireccionPago"), 36), 36) & Strings.Space(36 - Length(Strings.Mid(Nz(Dr("DireccionPago")), 36))) & Strings.Space(5)
            End If
            If Length(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
            Dim DrNew5 As DataRow = DtFichEspecial.NewRow
            DrNew5("Linea") = StrRegistro
            DtFichEspecial.Rows.Add(DrNew5)
            IntRegTotalesEspecial += 1
        End If

        'Sexto Registro de Detalle Especial
        StrRegistro = "0661"
        StrRegistro &= StrEmpresaRegistro
        StrRegistro &= StrSufijo
        StrRegistro &= StrRefBeneficiario & "048"
        If Length(Dr("CodPostal") & String.Empty) > 5 Then
            StrRegistro &= Strings.Left(Nz(Dr("CodPostal"), String.Empty), 5)
        Else : StrRegistro &= Strings.Left(Nz(Dr("CodPostal"), String.Empty), 5) & Strings.Space(5 - Length(Nz(Dr("CodPostal"))))
        End If
        If Length(Dr("PoblacionPago") & String.Empty) > 31 Then
            StrRegistro &= Strings.Left(Nz(Dr("PoblacionPago"), String.Empty), 31) & Strings.Space(5)
        Else : StrRegistro &= Strings.Left(Nz(Dr("PoblacionPago"), String.Empty), 31) & Strings.Space(31 - Length(Nz(Dr("PoblacionPago")))) & Strings.Space(5)
        End If
        If Length(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
        Dim DrNew6 As DataRow = DtFichEspecial.NewRow
        DrNew6("Linea") = StrRegistro
        DtFichEspecial.Rows.Add(DrNew6)
        IntRegTotalesEspecial += 1

        'Séptimo Registro de Detalle Especial
        StrRegistro = "0661"
        StrRegistro &= StrEmpresaRegistro
        StrRegistro &= StrSufijo
        StrRegistro &= StrRefBeneficiario & "049"
        If Length(Dr("DescPais") & String.Empty) > 36 Then
            StrRegistro &= Strings.Left(Nz(Dr("DescPais"), String.Empty), 36) & Strings.Space(5)
        Else : StrRegistro &= Strings.Left(Nz(Dr("DescPais"), String.Empty), 36) & Strings.Space(36 - Length(Nz(Dr("DescPais")))) & Strings.Space(5)
        End If
        If Length(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
        Dim DrNew7 As DataRow = DtFichEspecial.NewRow
        DrNew7("Linea") = StrRegistro
        DtFichEspecial.Rows.Add(DrNew7)
        IntRegTotalesEspecial += 1

        'Octavo Registro de Detalle Especial
        StrRegistro = "0661"
        StrRegistro &= StrEmpresaRegistro
        StrRegistro &= StrSufijo
        StrRegistro &= StrRefBeneficiario & "050"

        'Obtenemos el número de SuFactura o de varias si está agrupado
        StrPagoAgrupado = New Parametro().ObtenerPredeterminado("FRAPAGOAGR") & String.Empty

        If Length(Dr("NFactura")) > 0 Then
            If Dr("NFactura") = StrPagoAgrupado Then
                Dim GrupoPago As New EstGrupoPago
                GrupoPago.IDPago = Dr("IDPago")
                GrupoPago.DefFichero = enumDefFichero.Fich34
                StrConcepto = ProcessServer.ExecuteTask(Of EstGrupoPago, String)(AddressOf FacturasPagoAgrupado, GrupoPago, New ServiceProvider)
            Else : StrConcepto = "S/FRA. " & Dr("SuFactura")
            End If
        ElseIf Nz(Dr("Texto")) <> String.Empty Then
            StrConcepto = Dr("Texto")
        Else : StrConcepto = String.Empty
        End If

        If Length(StrConcepto) > 36 Then
            StrRegistro &= Strings.Left(StrConcepto, 36) & Strings.Space(5)
        Else : StrRegistro &= Strings.Left(StrConcepto, 36) & Strings.Space(36 - Length(StrConcepto)) & Strings.Space(5)
        End If
        If Len(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
        Dim DrNew8 As DataRow = DtFichEspecial.NewRow
        DrNew8("Linea") = StrRegistro
        DtFichEspecial.Rows.Add(DrNew8)
        IntRegTotalesEspecial += 1

        'Noveno Registro de Detalle Especial
        If Length(StrConcepto) > 36 Then
            StrRegistro = "0661"
            StrRegistro &= StrEmpresaRegistro
            StrRegistro &= StrSufijo
            StrRegistro &= StrRefBeneficiario & "051"
            If Length(Strings.Mid(StrConcepto, 36)) > 36 Then
                StrRegistro &= Strings.Left(Strings.Mid(StrConcepto, 36), 36) & Strings.Space(5)
            Else : StrRegistro &= Strings.Left(Strings.Mid(StrConcepto, 36), 36) & Strings.Space(36 - Length(Strings.Mid(StrConcepto, 36))) & Strings.Space(5)
            End If
            If Length(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
            Dim DrNew9 As DataRow = DtFichEspecial.NewRow
            DrNew9("Linea") = StrRegistro
            DtFichEspecial.Rows.Add(DrNew9)
            IntRegTotalesEspecial += 1
        End If

        'Undécimo Registro de Detalle Especial
        StrRegistro = "0661"
        StrRegistro &= StrEmpresaRegistro
        StrRegistro &= StrSufijo
        StrRegistro &= StrRefBeneficiario & "053"
        StrRegistro &= Space(41)
        If Length(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
        Dim DrNew10 As DataRow = DtFichEspecial.NewRow
        DrNew10("Linea") = StrRegistro
        DtFichEspecial.Rows.Add(DrNew10)
        IntRegTotalesEspecial += 1

        'Duodécimo Registro de Detalle Especial
        StrRegistro = "0661"
        StrRegistro &= StrEmpresaRegistro
        StrRegistro &= StrSufijo
        StrRegistro &= StrRefBeneficiario & "054"
        StrRegistro &= Strings.Space(41)
        If Length(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
        Dim DrNew11 As DataRow = DtFichEspecial.NewRow
        DrNew11("Linea") = StrRegistro
        DtFichEspecial.Rows.Add(DrNew11)
        IntRegTotalesEspecial += 1

        'Decimotercero Registro de Detalle Especial
        StrRegistro = "0661"
        StrRegistro &= StrEmpresaRegistro
        StrRegistro &= StrSufijo
        StrRegistro &= StrRefBeneficiario & "055"
        StrRegistro &= "01"
        If Length(Dr("IDPartidaEstadistica") & String.Empty) > 0 Then
            StrRegistro &= Strings.Left(Nz(Dr("IDPartidaEstadistica")), 6)
        Else : StrRegistro &= Strings.Space(6)
        End If
        If Length(Dr("CodigoIBAN") & String.Empty) > 0 Then
            StrRegistro &= Strings.Left(Nz(Dr("CodigoIBAN")), 2)
        Else : StrRegistro &= Strings.Space(2)
        End If
        StrRegistro &= StrEmpresaRegistro 'Creemos que es el mismo
        StrRegistro &= Strings.Space(8) 'No se sabe que código
        StrRegistro &= Strings.Space(12) 'No se sabe que código
        StrRegistro &= Strings.Space(2)
        If Length(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
        Dim DrNew12 As DataRow = DtFichEspecial.NewRow
        DrNew12("Linea") = StrRegistro
        DtFichEspecial.Rows.Add(DrNew12)
        IntRegTotalesEspecial += 1
    End Sub

    Private Sub GenerarBloqueNacional(ByRef DtFichero As DataTable, ByVal DtFichNacional As DataTable, ByRef IntRegTotalesNacional As Integer, ByRef DblImporteTotalNacional As Double, ByVal StrEmpresaRegistro As String, ByVal StrSufijo As String, ByRef IntRegistros010 As Integer)
        Dim StrRegistro, StrImporteTotalNacional As String

        'Registro de Cabecera Nacional
        StrRegistro = "0456"
        StrRegistro &= StrEmpresaRegistro
        StrRegistro &= StrSufijo
        StrRegistro &= Strings.Space(56)
        If Length(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
        Dim DrNew As DataRow = DtFichero.NewRow
        DrNew("Linea") = StrRegistro
        DtFichero.Rows.Add(DrNew)
        IntRegTotalesNacional += 1

        'Inserto los registros de Detalle
        For Each Dr As DataRow In DtFichNacional.Select
            DtFichero.Rows.Add(Dr.ItemArray)
        Next

        'Registros Totales Nacionales (Obligatorio)
        StrRegistro = "0856"
        StrRegistro &= StrEmpresaRegistro
        StrRegistro &= StrSufijo
        StrRegistro &= Strings.Space(15)
        StrImporteTotalNacional = Format(DblImporteTotalNacional * 100, "000000000000")
        StrRegistro &= StrImporteTotalNacional & Strings.Format(IntRegistros010, "00000000") & Strings.Format(IntRegTotalesNacional + 1, "0000000000")
        StrRegistro &= Strings.Space(11)
        If Len(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
        Dim DrNew1 As DataRow = DtFichero.NewRow
        DrNew1("Linea") = StrRegistro
        DtFichero.Rows.Add(DrNew1)
        IntRegTotalesNacional += 1
    End Sub

    Private Sub GenerarBloqueTransfronterizo(ByRef DtFichero As DataTable, ByVal DtFichTrans As DataTable, ByRef IntRegTotalesTrans As Integer, ByRef DblImporteTotalTrans As Double, ByVal StrEmpresaRegistro As String, ByVal StrSufijo As String, ByRef IntRegistros033 As Integer)
        Dim StrRegistro, StrImporteTotalTrans As String

        'Registro de Cabecera Transfronterizo
        StrRegistro = "0460"
        StrRegistro &= StrEmpresaRegistro
        StrRegistro &= StrSufijo
        StrRegistro &= Strings.Space(56)
        If Length(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
        Dim DrNew As DataRow = DtFichero.NewRow
        DrNew("Linea") = StrRegistro
        DtFichero.Rows.Add(DrNew)
        IntRegTotalesTrans += 1

        'Inserto los registros de Detalle
        For Each Dr As DataRow In DtFichTrans.Select
            DtFichero.Rows.Add(Dr.ItemArray)
        Next

        'Registros Totales Transfronterizos (Obligatorio)
        StrRegistro = "0860"
        StrRegistro &= StrEmpresaRegistro
        StrRegistro &= StrSufijo
        StrRegistro &= Strings.Space(15)
        StrImporteTotalTrans = Format(DblImporteTotalTrans * 100, "000000000000")
        StrRegistro &= StrImporteTotalTrans & Strings.Format(IntRegistros033, "00000000") & Strings.Format(IntRegTotalesTrans + 1, "0000000000")
        StrRegistro &= Strings.Space(11)
        If Len(StrRegistro) < 72 Then StrRegistro &= (Strings.Space(72 - Length(StrRegistro)))
        Dim DrNew1 As DataRow = DtFichero.NewRow
        DrNew1("Linea") = StrRegistro
        IntRegTotalesTrans += 1
    End Sub

    Private Sub GenerarBloqueEspecial(ByRef DtFichero As DataTable, ByVal DtFichEspecial As DataTable, _
                                      ByRef IntRegTotalesEspecial As Short, ByRef DblImporteTotalEspecial As Double, _
                                      ByVal StrEmpresaRegistro As String, ByVal StrSufijo As String, ByRef IntRegistros043 As Short)
        Dim strRegistro As String
        Dim strImporteTotalEspecial As String

        'Registro de Cabecera Especial
        Dim RegEsp As New DatosCabTransfer
        RegEsp.IDReg = "0461"
        RegEsp.CifEmpresa = StrEmpresaRegistro
        RegEsp.RefEmpresa = StrSufijo
        RegEsp.Datos = Strings.Space(56)
        strRegistro = RegEsp.IDReg & RegEsp.CifEmpresa & RegEsp.RefEmpresa & RegEsp.Datos
        If strRegistro.Length < 72 Then strRegistro &= Strings.Space(72 - strRegistro.Length)
        Dim DrRegCab As DataRow = DtFichero.NewRow
        DrRegCab("Linea") = strRegistro
        DtFichero.Rows.Add(DrRegCab)
        IntRegTotalesEspecial += 1

        'Inserto los registros de Detalle
        For Each Dr As DataRow In DtFichEspecial.Select
            DtFichero.Rows.Add(Dr.ItemArray)
        Next

        'Registros Totales Especiales (Obligatorio)
        Dim RegTot As New DatosCabTransfer
        RegTot.IDReg = "0861"
        RegTot.CifEmpresa = StrEmpresaRegistro
        RegTot.RefEmpresa = StrSufijo
        RegTot.NumReg = Strings.Space(15)
        strImporteTotalEspecial = Format(DblImporteTotalEspecial * 100, "000000000000")
        RegTot.Datos = strImporteTotalEspecial & Strings.Format(IntRegistros043, "00000000") & Strings.Format(IntRegTotalesEspecial + 1, "0000000000") & Strings.Space(11)
        strRegistro = RegTot.IDReg & RegTot.CifEmpresa & RegTot.RefEmpresa & RegTot.NumReg & RegTot.Datos
        If strRegistro.Length < 72 Then strRegistro &= Strings.Space(72 - strRegistro.Length)
        Dim DrTot As DataRow = DtFichero.NewRow
        DrTot("Linea") = strRegistro
        DtFichero.Rows.Add(DrTot)
        IntRegTotalesEspecial += 1
    End Sub

    Private Function TratarSimbolosEspeciales(ByVal StrCadena As String, Optional ByVal BlnExcluirN As Boolean = True, Optional ByVal blnSimbolosExtendidos As Boolean = False) As String
        Dim SimbolosExtendidos() As Integer = {45}
        '// 45: guión

        '1º Grupo de Simbolos Especiales
        For i As Integer = 33 To 47
            If Not blnSimbolosExtendidos OrElse Not SimbolosExtendidos.Contains(i) Then
                StrCadena = Strings.Replace(StrCadena, Chr(i), "", , , CompareMethod.Binary)
            End If
        Next
        '2º Grupo de Simbolos Especiales
        For i As Integer = 58 To 64
            StrCadena = Strings.Replace(StrCadena, Chr(i), "", , , CompareMethod.Binary)
        Next
        '3º Grupo de Simbolos Especiales
        For i As Integer = 91 To 96
            StrCadena = Strings.Replace(StrCadena, Chr(i), "", , , CompareMethod.Binary)
        Next
        '4º Grupo de Simbolos Especiales
        For i As Integer = 123 To 126
            StrCadena = Strings.Replace(StrCadena, Chr(i), "", , , CompareMethod.Binary)
        Next

        'Simbolo de la Ñ (164-165)
        If BlnExcluirN Then
            StrCadena = Strings.Replace(StrCadena, "ñ", "N", , , CompareMethod.Binary)
            StrCadena = Strings.Replace(StrCadena, "Ñ", "N", , , CompareMethod.Binary)
        End If

        'Caracteres con Acentos en mayusculas
        StrCadena = Strings.Replace(StrCadena, "Á", "A", , , CompareMethod.Binary) : StrCadena = Strings.Replace(StrCadena, "á", "A", , , CompareMethod.Binary)
        StrCadena = Strings.Replace(StrCadena, "É", "E", , , CompareMethod.Binary) : StrCadena = Strings.Replace(StrCadena, "é", "E", , , CompareMethod.Binary)
        StrCadena = Strings.Replace(StrCadena, "Í", "I", , , CompareMethod.Binary) : StrCadena = Strings.Replace(StrCadena, "í", "I", , , CompareMethod.Binary)
        StrCadena = Strings.Replace(StrCadena, "Ó", "O", , , CompareMethod.Binary) : StrCadena = Strings.Replace(StrCadena, "ó", "O", , , CompareMethod.Binary)
        StrCadena = Strings.Replace(StrCadena, "Ú", "U", , , CompareMethod.Binary) : StrCadena = Strings.Replace(StrCadena, "ú", "U", , , CompareMethod.Binary)

        'Últimos símbolos sueltos de la tabla Ascii
        For i As Integer = 128 To 255
            StrCadena = Strings.Replace(StrCadena, Chr(i), "", , , CompareMethod.Binary)
        Next

        Return StrCadena
    End Function

    Public Shared Function FormatearNumeros(ByVal StrNum As String, ByVal IntRelleno As Integer) As String
        If IntRelleno - StrNum.Length < 0 Then
            Return New String("0", IntRelleno)
        Else
            Dim StrRelleno As String = New String("0", IntRelleno - StrNum.Length)
            If StrRelleno.Length > 0 Then StrNum = StrRelleno & StrNum
            Return StrNum
        End If
    End Function

#End Region

End Class