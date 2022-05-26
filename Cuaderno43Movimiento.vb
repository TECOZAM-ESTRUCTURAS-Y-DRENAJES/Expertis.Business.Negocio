Option Strict Off
Option Explicit On
Public Class Cuaderno43Movimiento
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
	Public Sub New
		MyBase.New(cnEntidad)
	End Sub

    Private Const cnEntidad As String = "tbCuaderno43Movimiento"
    Public Structure typCabeceraCuenta
        <VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=4)> Public strBanco As String
        <VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=4)> Public strOficina As String
        <VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=10)> Public strCuenta As String
        <VBFixedString(6), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=6)> Public strFechaInicial As String
        <VBFixedString(6), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=6)> Public strFechaFinal As String
        <VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=1)> Public strCodigoSaldo As String
        <VBFixedString(14), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=14)> Public strImporteSaldo As String
        <VBFixedString(3), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=3)> Public strClaveDivisa As String
        <VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=1)> Public strInformacion As String
        <VBFixedString(26), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=26)> Public strNombre As String
        <VBFixedString(3), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=3)> Public strLibre As String
    End Structure

    Public Structure typMovimiento
        <VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=4)> Public strLibre1 As String
        <VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=4)> Public strOficina As String
        <VBFixedString(6), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=6)> Public strFechaoperacion As String
        <VBFixedString(6), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=6)> Public strFechavalor As String
        <VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=2)> Public strConceptoComun As String
        <VBFixedString(3), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=3)> Public strConceptoPropio As String
        <VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=1)> Public strdebehaber As String
        <VBFixedString(14), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=14)> Public strImporte As String
        <VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=10)> Public strdocumento As String
        <VBFixedString(28), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=28)> Public strLibre2 As String
        <VBFixedString(78), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=78)> Public strDescMovimiento As String
        Dim strFichero As String
    End Structure

    Private mvcabecera As typCabeceraCuenta
    Private Declare Function MoveFileEx Lib "kernel32" Alias "MoveFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal dwFlags As Integer) As Integer
    Private Const MOVEFILE_COPY_ALLOWED As Short = &H2S
    Private VectorRutas() As String
    Private mrcsBancosPropios As DataTable
    Private mstrIDBDInicial As String
    Private mblnPrimera As Boolean
    Private mblnGrabar As Boolean

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarIdentificador)
    End Sub

    <Task()> Public Shared Sub AsignarIdentificador(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added AndAlso Length(data("ID")) = 0 Then data("ID") = AdminData.GetAutoNumeric
    End Sub

    Public Function Leer_Ficheros43() As Integer
        Dim FwnCuaderno43Dir As Cuaderno43Directorio
     
        Dim strSystemDB As String

        Leer_Ficheros43 = True

        '1º Obtenemos todas las rutas de los ficheros 43
        FwnCuaderno43Dir = New Cuaderno43Directorio
        Dim dtRutas As DataTable = FwnCuaderno43Dir.Filter("*")
      
        If dtRutas.Rows.Count <> 0 Then
            Dim dtFichero As DataTable = CrearRcsFichero(dtRutas)
            If dtFichero.Rows.Count <> 0 Then
                'Obtenemos la Base de Datos de Sistema
                strSystemDB = GetSystemDataBase()
                If Len(strSystemDB) > 0 Then
                    'Obtenemos un rs con todos los BancosPropios de todas la BD's
                    mrcsBancosPropios = AdminData.Execute("getbancosmultiempresa", False, strSystemDB)
                    'Guardamos la BD con la que hemos arrancado
                    LeerMovimientos(dtFichero)
                End If
            Else
                Leer_Ficheros43 = False
            End If
        Else
            ApplicationService.GenerateError("No se han creado directorios para leer estos ficheros.", "Leer_Ficheros43")
        End If


        Leer_Ficheros43 = False
        MoverFicheros(True)
    End Function
    Private Function CrearRcsFichero(ByVal dtRutas As DataTable) As DataTable
        'Dim dtFichero As DataTable
        'Dim intPosicion As Short
        'Dim lngPos As Integer
        'Dim strLinea As New VB6.FixedLengthString(80)
        'Dim lngPosicion As Integer
        'Dim intFile As Short
        'Dim strDirectorio As String
        'Dim strFichero As String
        'Dim strRuta As String
        'Dim strRutaCompleta As String
        'Dim strExtension As String
        'Dim strFicheroCopia As String
        'Dim strRutaCopia As String
        'Dim strRutaCompletaCopia As String
        'Dim strExtensionCopia As String
        'Dim strDirectorioCopia As String
        'Dim i As Short
        'Dim CadenaRutas As String


        ''Creamos el DT que contendrá las líneas del fichero
        'If dtRutas.Rows.Count <> 0 Then
        '    dtFichero.Columns.Add("Linea", GetType(String))
        '    dtFichero.Columns.Add("NomFichero", GetType(String))
        '    For Each drRuta As DataRow In dtRutas.Rows
        '        strDirectorio = drRuta("DirectorioFichero")
        '        strExtension = drRuta("ExtensionFichero")
        '        strDirectorioCopia = drRuta("DirectorioCopia")
        '        strExtensionCopia = drRuta("ExtensionCopia")
        '        strRuta = strDirectorio & "\*." & strExtension
        '        strFichero = Dir(strRuta)
        '        If strFichero <> vbNullString Then
        '            Do While strFichero <> vbNullString
        '                i = i + 1
        '                strRutaCompleta = strDirectorio & "\" & strFichero
        '                strFicheroCopia = Replace(strFichero, strExtension, strExtensionCopia, , , CompareMethod.Text)
        '                strRutaCompletaCopia = strDirectorioCopia & "\" & strFicheroCopia
        '                CadenaRutas = strRutaCompleta & ";" & strRutaCompletaCopia
        '                ReDim Preserve VectorRutas(i)
        '                VectorRutas(i) = CadenaRutas
        '                'Abrimos el fichero para lectura
        '                intFile = FreeFile()
        '                FileOpen(intFile, strRutaCompleta, OpenMode.Binary, OpenAccess.Read)
        '                lngPosicion = 1
        '                Do While Not EOF(intFile)
        '                    FileGet(intFile, strLinea.Value, lngPosicion)
        '                    dtFichero.AddNew()
        '                    dtFichero.Rows(0)("Linea") = strLinea.Value
        '                    lngPosicion = lngPosicion + 82
        '                   dtFichero.Rows(0)("NomFichero") = strFichero
        '                Loop
        '                FileClose(intFile)
        '                strFichero = Dir()
        '            Loop
        '        End If
        '    Next
        '    CrearRcsFichero = dtFichero
        'End If


        'Exit Function

    End Function
    Private Function ExtraerCabecera(ByVal strCabecera As String) As typCabeceraCuenta


        mvcabecera.strBanco = Mid(strCabecera, 3, 4)
        mvcabecera.strOficina = Mid(strCabecera, 7, 4)
        mvcabecera.strCuenta = Mid(strCabecera, 11, 10)
        mvcabecera.strFechaInicial = Mid(strCabecera, 21, 6)
        mvcabecera.strFechaFinal = Mid(strCabecera, 27, 6)
        mvcabecera.strCodigoSaldo = Mid(strCabecera, 33, 1)
        mvcabecera.strImporteSaldo = Mid(strCabecera, 34, 14)
        mvcabecera.strClaveDivisa = Mid(strCabecera, 48, 3)
        mvcabecera.strInformacion = Mid(strCabecera, 51, 1)
        mvcabecera.strNombre = Mid(strCabecera, 52, 26)
        mvcabecera.strLibre = Mid(strCabecera, 78, 3)

        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto ExtraerCabecera. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        ExtraerCabecera = mvcabecera


        Exit Function

    End Function
    Private Sub LeerMovimientos(ByRef dtFichero As DataTable)
        'Dim strCodigoRegistro As String
        'Dim vMovimientos As typMovimiento
        'Dim strDescMovimiento As New VB6.FixedLengthString(78)
        'Dim strCabecera As String
        'Dim strIdBaseDatos As String
        'Dim strDBActual As String
        'Dim FwnAplicacion As AdminData
        'Dim rcsMov As Recordset
        'Dim nvdAct As AdminData


        'Do While Not rcsFichero.EOF
        '    strCodigoRegistro = Mid(rcsFichero.Fields("Linea").Value, 1, 2)
        '    Select Case strCodigoRegistro

        '        Case CStr(11)
        '            If mblnGrabar Then
        '                GrabarMovimientos(rcsMov)
        '            End If
        '            'Obtengo la cabecera del fichero
        '            strCabecera = rcsFichero.Fields("Linea").Value
        '            'Extraigo los campos de la cabecera
        '            'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto mvcabecera. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        '            mvcabecera = ExtraerCabecera(strCabecera)
        '            'Extraigo la BD donde grabar los movimientos
        '            strIdBaseDatos = ExtraerBaseDatos()
        '            If Len(strIdBaseDatos) > 0 Then
        '                CambiarBD(strIdBaseDatos)
        '                rcsMov = AddNewForm()
        '            End If
        '        Case CStr(22), CStr(23), CStr(24)
        '            If Len(strIdBaseDatos) > 0 Then
        '                If strCodigoRegistro = "22" And mblnPrimera Then
        '                    rcsMov.AddNew()
        '                End If
        '                mblnPrimera = True
        '                'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto vMovimientos. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        '                vMovimientos = ExtraerMovimientos(rcsFichero, strCodigoRegistro)
        '                TratarMvto(vMovimientos, rcsMov, strCodigoRegistro)
        '                mblnGrabar = True
        '            End If
        '    End Select
        '    rcsFichero.MoveNext()
        'Loop

        'If mblnGrabar Then
        '    GrabarMovimientos(rcsMov)
        'End If

        'nvdAct = New AdminData
        'nvdAct.SetData(udtRcs)
        ''UPGRADE_NOTE: El objeto nvdAct no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
        'nvdAct = Nothing

        ''movemos los ficheros ya leidos
        'MoverFicheros(False)

        ''Nos situamos de nuevo en la Base de Datos inicial
        'strDBActual = GetPropertyValue(mstrIDBDInicial, "DataBase")
        'FwnAplicacion = New AdminData
        ''REVISAR(11/10/05)
        ''FwnAplicacion.SetDataBase(strDBActual)
        ''UPGRADE_NOTE: El objeto FwnAplicacion no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
        'FwnAplicacion = Nothing


        ''UPGRADE_NOTE: El objeto rcsMov no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
        'If Not rcsMov Is Nothing Then rcsMov = Nothing
        ''UPGRADE_NOTE: El objeto FwnAplicacion no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
        'FwnAplicacion = Nothing
        'Exit Sub

        'Resume
    End Sub
    Private Function ExtraerBaseDatos() As String


        'mrcsBancosPropios.Filter = "IdBanco='" & mvcabecera.strBanco & "' AND NCuenta='" & mvcabecera.strCuenta & "'"
        'If mrcsBancosPropios.RecordCount <> 0 Then
        '    ExtraerBaseDatos = mrcsBancosPropios.Fields("IdBaseDatos").Value
        '    mrcsBancosPropios.Filter = FilterGroupEnum.adFilterNone
        'Else
        '    ApplicationService.GenerateError("El Banco Propio no existe.", vbInformation, "ExtraerBaseDatos")
        '    ExtraerBaseDatos = vbNullString
        'End If


        'Exit Function

    End Function
    Private Function ExtraerMovimientos(ByRef dtFichero As DataTable, ByVal strCodigoRegistro As String) As typMovimiento
        Dim vMovimientos As typMovimiento


        Select Case strCodigoRegistro
            Case CStr(22)
                vMovimientos.strLibre1 = Mid(dtFichero.Rows(0)("Linea"), 3, 4)
                vMovimientos.strOficina = Mid(dtFichero.Rows(0)("Linea"), 7, 4)
                vMovimientos.strFechaoperacion = Mid(dtFichero.Rows(0)("Linea"), 11, 6)
                vMovimientos.strFechavalor = Mid(dtFichero.Rows(0)("Linea"), 17, 6)
                vMovimientos.strConceptoComun = Mid(dtFichero.Rows(0)("Linea"), 23, 2)
                vMovimientos.strConceptoPropio = Mid(dtFichero.Rows(0)("Linea"), 25, 3)
                vMovimientos.strdebehaber = Mid(dtFichero.Rows(0)("Linea"), 28, 1)
                vMovimientos.strImporte = Mid(dtFichero.Rows(0)("Linea"), 29, 14)
                vMovimientos.strdocumento = Mid(dtFichero.Rows(0)("Linea"), 43, 10)
                vMovimientos.strLibre2 = Mid(dtFichero.Rows(0)("Linea"), 53, 28)
                vMovimientos.strFichero = dtFichero.Rows(0)("NomFichero")
            Case CStr(23), CStr(24)
                vMovimientos.strDescMovimiento = Mid(dtFichero.Rows(0)("Linea"), 3, 78)
        End Select

        ExtraerMovimientos = vMovimientos


        Exit Function

    End Function
    Private Sub TratarMvto(ByRef vMovimiento As typMovimiento, ByRef dtMov As DataTable, ByVal strCodigoRegistro As String)
        Dim clsAdmin As AdminData


        If Not dtMov Is Nothing Then
            With dtMov
                Select Case strCodigoRegistro
                    Case CStr(22)
                        clsAdmin = New AdminData
                        .Rows(0)("ID") = clsAdmin.GetAutoNumeric
                        'UPGRADE_NOTE: El objeto clsAdmin no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
                        clsAdmin = Nothing
                        .Rows(0)("FechaIntegracion") = Today
                        .Rows(0)("nombrefichero") = vMovimiento.strFichero
                        .Rows(0)("IDBanco") = mvcabecera.strBanco
                        .Rows(0)("Sucursal") = mvcabecera.strOficina
                        .Rows(0)("NCuenta") = mvcabecera.strCuenta
                        .Rows(0)("ConceptoComun") = vMovimiento.strConceptoComun
                        .Rows(0)("ConceptoPropio") = vMovimiento.strConceptoPropio
                        .Rows(0)("ClaveDH") = vMovimiento.strdebehaber
                        .Rows(0)("FechaMovimiento") = Formato_Fecha(vMovimiento.strFechaoperacion)
                        .Rows(0)("FechaValor") = Formato_Fecha(vMovimiento.strFechavalor)
                        .Rows(0)("ImporteA") = CDbl(Left(vMovimiento.strImporte, 12)) + (CDbl(Right(vMovimiento.strImporte, 2)) / 100)
                        .Rows(0)("NDocumento") = vMovimiento.strdocumento
                    Case CStr(23), CStr(24)
                        .Rows(0)("DescripcionMovimiento") = vMovimiento.strDescMovimiento
                End Select
            End With

        End If


        Exit Sub

    End Sub

    Private Function MoverFicheros(ByVal blnAtras As Boolean) As Integer
        Dim lngResult As Integer
        Dim strFicheroOrigen As String
        Dim strFicheroDestino As String
        Dim intPos As Short
        Dim i As Short
        Dim lngerror As Integer


        MoverFicheros = False

        On Error Resume Next

        If UBound(VectorRutas) > 0 Then
            If Err.Number = 0 Then
                For i = 1 To UBound(VectorRutas)
                    intPos = InStr(1, VectorRutas(i), ";")
                    If Not blnAtras Then
                        strFicheroOrigen = Mid(VectorRutas(i), 1, intPos - 1)
                        strFicheroDestino = Mid(VectorRutas(i), intPos + 1)
                    Else
                        strFicheroOrigen = Mid(VectorRutas(i), intPos + 1)
                        strFicheroDestino = Mid(VectorRutas(i), 1, intPos - 1)
                    End If
                    Rename(strFicheroOrigen, strFicheroDestino)
                Next i
            Else
            End If
        End If

        MoverFicheros = True


        Exit Function

    End Function
    Private Function Formato_Fecha(ByVal strFecha As String) As String


        Formato_Fecha = New Date(CInt("20" & Left(strFecha, 2)), CInt(Mid(strFecha, 3, 2)), CInt(Right(strFecha, 2)))
        'Formato_Fecha = Right(strFecha, 2) & "/" & Mid(strFecha, 3, 2) & "/" & Left(strFecha, 2)


        Exit Function

    End Function
    Public Sub GrabarMovimientos(ByRef dtMov As DataTable)


        'mblnGrabar = False
        'If rcsMov.RecordCount <> 0 Then
        '    CarrierCargarRcs(udtRcs, rcsMov)
        '    'UPGRADE_NOTE: El objeto rcsMov no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
        '    rcsMov = Nothing
        '    mblnPrimera = False
        'End If


        'On Error Resume Next
        ''Cerrar los objetos abiertos en el procedimiento actual
        'Exit Sub

    End Sub
    Public Sub CambiarBD(ByVal strIdBaseDatos As String)
        Dim FwnAplicacion As AdminData


        FwnAplicacion = New AdminData
        'REVISAR(11/10/05)
        'FwnAplicacion.SetDataBase(strIdBaseDatos)
        'UPGRADE_NOTE: El objeto FwnAplicacion no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
        FwnAplicacion = Nothing


        On Error Resume Next
        'Cerrar los objetos abiertos en el procedimiento actual
        Exit Sub

    End Sub

    Private Function GetSystemDataBase() As String
        Dim strFichero As String
        Dim strRutaCompleta As String
        Dim strLinea As String
        strFichero = Dir(VB6.GetPath & "\*.udl")
        strRutaCompleta = VB6.GetPath & "\" & strFichero
        If strFichero <> vbNullString Then
            FileOpen(1, strRutaCompleta, OpenMode.Input)
            Do While Not EOF(1)
                strLinea = LineInput(1)
            Loop
            FileClose(1)
        End If

        If Len(strLinea) > 0 Then
            'TODO GetPropertyValue
            'GetSystemDataBase = GetPropertyValue(strLinea, "Initial Catalog")
        End If


        On Error Resume Next
        'Cerrar los objetos abiertos en el procedimiento actual
        Exit Function


    End Function

    Public Sub Marcar_Procesados(ByVal strWhere As String)
        Dim dtMov As DataTable = AdminData.Filter("VCtlCILecturaCSB43", , strWhere)
        If dtMov.Rows.Count <> 0 Then
            For Each drMov As DataRow In dtMov.rows
                drMov("Procesado") = 1
            Next
            Update(dtMov)
        End If
        Exit Sub
    End Sub
End Class