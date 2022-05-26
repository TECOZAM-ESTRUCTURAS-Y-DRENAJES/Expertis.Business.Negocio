Public Class CierreInventario

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbCierreInventario"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data("Contabilizado") AndAlso Not data("Cerrado") Then ApplicationService.GenerateError("Los estados del cierre, Cerrado y Contabilizado, son incompatibles entre sí ")
    End Sub

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarFechaCierre)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarValoresDeshacerCierre)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarValoresDescontabilizarCierre)
    End Sub

    <Task()> Public Shared Sub AsignarFechaCierre(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data("Cerrado") AndAlso IsDBNull(data("FechaCierre")) Then data("FechaCierre") = Today.Date
    End Sub

    <Task()> Public Shared Sub AsignarValoresDeshacerCierre(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data("Cerrado", DataRowVersion.Original) AndAlso Not data("Contabilizado", DataRowVersion.Original) Then
            If Not data("Cerrado") AndAlso Not data("Contabilizado") Then
                Dim Cierre As New DataCierre
                Cierre.IDEjercicio = data("IDEjercicio")
                Cierre.Periodo = data("idmescierre")
                Dim blnValidar As Boolean = ProcessServer.ExecuteTask(Of DataCierre, Boolean)(AddressOf ValidarDeshacerCierre, Cierre, services)

                If blnValidar Then
                    'data("PropuestaCorrecta") = False
                    data("FechaCierre") = System.DBNull.Value
                    data("ValorA") = 0 : data("ValorB") = 0
                ElseIf data("Cerrado") AndAlso data("Contabilizado") Then
                    For Each Dc As DataColumn In data.Table.Columns
                        If Dc.ColumnName <> "Observaciones" Then
                            If data(Dc.ColumnName, DataRowVersion.Original) <> data(Dc.ColumnName) Then
                                ApplicationService.GenerateError("No se permite actualizar un Cierre de Inventario Cerrado")
                            End If
                        End If
                    Next
                End If
            End If
        End If
    End Sub
    <Task()> Public Shared Sub AsignarValoresDescontabilizarCierre(ByVal data As DataRow, ByVal services As ServiceProvider)

        If data("Cerrado", DataRowVersion.Original) AndAlso data("Contabilizado", DataRowVersion.Original) Then
            If Not data("Cerrado") AndAlso data("Contabilizado") Then
                ApplicationService.GenerateError("El Cierre está Contabilizado")
            ElseIf Not data("Cerrado") And Not data("Contabilizado") Then
                ApplicationService.GenerateError("Los estados del cierre, Cerrado y Contabilizado, son incompatibles entre sí ")
            ElseIf data("Cerrado") And Not data("Contabilizado") Then
                Dim Cierre As New DataCierre(data("IDEjercicio"), data("IDMesCierre"))
                Dim blnValidar As Boolean = ProcessServer.ExecuteTask(Of DataCierre, Boolean)(AddressOf ValidarDescontabilizarCierre, Cierre, services)
                'Se quiere hacer una descontabilizacion
                If blnValidar Then
                    If Not IsDBNull(data("NAsiento")) Then
                        Dim f As New Filter
                        f.Add(New NumberFilterItem("NAsiento", data("NAsiento")))
                        If NegocioGeneral.DeleteWhere(data("IDEjercicio"), f) Then
                            data("NAsiento") = System.DBNull.Value
                        End If
                    Else
                        ApplicationService.GenerateError("No se puede determinar el Asiento a eliminar")
                    End If
                End If
            Else
                ApplicationService.GenerateError("No se permite actualizar un Cierre de Inventario Contabilizado")
            End If
        End If
    End Sub
    'Public Overloads Overrides Function Update(ByVal dttSource As System.Data.DataTable) As System.Data.DataTable
    '    Dim BlnCancel As Boolean
    '    If Not dttSource Is Nothing AndAlso dttSource.Rows.Count > 0 Then
    '        For Each Dr As DataRow In dttSource.Select
    '            If Dr.RowState = DataRowState.Added OrElse Dr.RowState = DataRowState.Modified Then
    '                If Dr.RowState = DataRowState.Modified Then
    '                    If Not Dr("Cerrado", DataRowVersion.Original) AndAlso Not Dr("Contabilizado", DataRowVersion.Original) Then
    '                        If Dr("Cerrado") AndAlso Dr("Contabilizado") Then ApplicationService.GenerateError("Los estados del cierre, Cerrado y Contabilizado, son incompatibles entre sí ")
    '                    ElseIf Dr("Cerrado", DataRowVersion.Original) AndAlso Not Dr("Contabilizado", DataRowVersion.Original) Then
    '                        If Not Dr("Cerrado") AndAlso Not Dr("Contabilizado") Then
    '                            If ValidarDeshacerCierre(Dr("IDEjercicio"), Dr("idmescierre")) = True Then
    '                                Dr("PropuestaCorrecta") = False
    '                                Dr("FechaCierre") = System.DBNull.Value
    '                                Dr("ValorA") = 0 : Dr("ValorB") = 0
    '                            ElseIf Dr("Cerrado") AndAlso Dr("Contabilizado") Then
    '                                For Each Dc As DataColumn In dttSource.Columns
    '                                    If Dc.ColumnName <> "Observaciones" Then
    '                                        If Dr(Dc.ColumnName, DataRowVersion.Original) <> Dr(Dc.ColumnName) Then
    '                                            ApplicationService.GenerateError("No se permite actualizar un Cierre de Inventario Cerrado")
    '                                        End If
    '                                    End If
    '                                Next
    '                            End If
    '                        ElseIf Dr("Cerrado", DataRowVersion.Original) AndAlso Dr("Contabilizado", DataRowVersion.Original) Then
    '                            If Not Dr("Cerrado") AndAlso Dr("Contabilizado") Then
    '                                ApplicationService.GenerateError("El Cierre está Contabilizado")
    '                            ElseIf Not Dr("Cerrado") And Not Dr("Contabilizado") Then
    '                                ApplicationService.GenerateError("Los estados del cierre, Cerrado y Contabilizado, son incompatibles entre sí ")
    '                            ElseIf Dr("Cerrado") And Not Dr("Contabilizado") Then
    '                                'Se quiere hacer una descontabilizacion
    '                                If ValidarDescontabilizarCierre(Dr("IDEjercicio"), Dr("idmescierre")) = True Then
    '                                    If Not IsDBNull(Dr("NAsiento")) Then
    '                                        If NegocioGeneral.DeleteWhere(Dr("IDEjercicio"), "NAsiento=" & Dr("NAsiento")) = True Then
    '                                            Dr("NAsiento") = System.DBNull.Value
    '                                        End If
    '                                    Else
    '                                        ApplicationService.GenerateError("No se puede determinar el Asiento a eliminar")
    '                                    End If
    '                                End If
    '                            Else
    '                                ApplicationService.GenerateError("No se permite actualizar un Cierre de Inventario Contabilizado")
    '                            End If
    '                        End If
    '                    End If
    '                End If
    '            End If
    '        Next
    '        MyBase.Update(dttSource)
    '    End If
    '    Return dttSource
    'End Function

#End Region

#Region "Definiciones"
    <Serializable()> _
    Public Class DataCierre
        Public IDEjercicio As String
        Public Periodo As Integer
        Public FechaHastaPeriodo As Date
        Public MostrarObsoletos As Boolean

        Public Sub New(ByVal IDEjercicio As String, ByVal Periodo As Integer)
            Me.IDEjercicio = IDEjercicio
            Me.Periodo = Periodo
            Me.FechaHastaPeriodo = Today
            Me.MostrarObsoletos = True
        End Sub
        Public Sub New()
            Me.FechaHastaPeriodo = Today
            Me.MostrarObsoletos = True
        End Sub
    End Class

    <Serializable()> _
    Public Class DataCierreFechas
        Public FechaDesde As Date
        Public FechaHasta As Date
        Public DtResultado As DataTable
        Public Resultado As enumstkResultadoCierre
        Public DtCierre As DataTable
    End Class
#End Region

#Region "Funciones Publicas"


    <Task()> Public Shared Function ValidacionPropuesta(ByVal DatosCierre As DataCierre, ByVal services As ServiceProvider) As DataTable
        Dim BlnExisteCabecera, BlnCrearCabecera As Boolean
        Dim DteFechaDesde, DteFechaHasta As Date
        Dim DtDatos As DataTable
        Dim DrCierre() As DataRow
        Const VIEWNAME As String = "vNegCierreInventario"
        If Length(DatosCierre.IDEjercicio) = 0 Then
            ApplicationService.GenerateError("El Ejercicio Contable es obligatorio.")
        Else
            Dim DtCierre As DataTable = New BE.DataEngine().Filter(VIEWNAME, New StringFilterItem("IDEjercicio", DatosCierre.IDEjercicio), , "FechaCierre DESC, FechaHasta DESC")
            If Not DtCierre Is Nothing AndAlso DtCierre.Rows.Count > 0 Then
                If AreEquals(DtCierre.Rows(0)("EjercicioCerrado"), 1) Then
                    'El ejercicio contable esta cerrado
                    ApplicationService.GenerateError("El Ejercicio Contable está Cerrado. No es posible lanzar el proceso.")
                Else
                    DrCierre = DtCierre.Select("IDMesCierre=" & DatosCierre.Periodo)
                    If DrCierre.Length > 0 Then
                        If Not DrCierre(0)("Cerrado") Then
                            BlnCrearCabecera = True
                        Else
                            If AreEquals(DrCierre(0)("Cerrado"), 1) Then
                                'El periodo esta cerrado
                                ApplicationService.GenerateError("El Cierre está Cerrado ")
                            Else
                                BlnExisteCabecera = True
                            End If
                        End If

                        If Length(DrCierre(0)("FechaDesde")) > 0 Then DteFechaDesde = DrCierre(0)("FechaDesde")
                        If Length(DrCierre(0)("FechaHasta")) > 0 Then DteFechaHasta = DrCierre(0)("FechaHasta")
                    Else
                        'El periodo no es valido
                        ApplicationService.GenerateError("El Período seleccionado no es válido. La Fecha de Cierre es obligatoria.")
                    End If
                End If
            Else
                ApplicationService.GenerateError("El Ejercicio introducido no existe en la Base de Datos.")
            End If
        End If

        Dim DtCabecera As DataTable
        If BlnCrearCabecera Then
            Dim CI As New CierreInventario
            DtCabecera = CI.AddNew()
            If Not DtCabecera Is Nothing Then
                Dim DrNew As DataRow = DtCabecera.NewRow
                DrNew("IDEjercicio") = DatosCierre.IDEjercicio : DrNew("idmescierre") = DatosCierre.Periodo
                DrNew("PropuestaCorrecta") = False
                DrNew("Cerrado") = False
                DrNew("FechaCierre") = System.DBNull.Value
                DrNew("Contabilizado") = False
                DrNew("ValorA") = 0 : DrNew("ValorB") = 0
                DrNew("Observaciones") = System.DBNull.Value
                DtCabecera.Rows.Add(DrNew)
            End If
        ElseIf BlnExisteCabecera Then
            If DrCierre.Length > 0 Then
                DtCabecera.ImportRow(DrCierre(0))
            End If
        End If
        If Not DtCabecera Is Nothing Then
            If DtCabecera.Rows.Count > 0 Then
                If BlnCrearCabecera Then
                    DtDatos = DtCabecera.Clone
                    If IsNothing(DtDatos) Then DtDatos = New DataTable
                    DtDatos.Columns.Add("FechaDesde", GetType(Date))
                    DtDatos.Columns.Add("FechaHasta", GetType(Date))
                End If
                Dim DrNew As DataRow = DtDatos.NewRow
                For Each dc As DataColumn In DtCabecera.Columns
                    DrNew(dc.ColumnName) = DtCabecera.Rows(0)(dc.ColumnName)
                Next
                If BlnCrearCabecera Then
                    DrNew("FechaDesde") = DteFechaDesde
                    DrNew("FechaHasta") = DteFechaHasta
                    DtDatos.Rows.Add(DrNew)
                End If
                Return DtDatos
            End If
        End If
    End Function

    <Task()> Public Shared Function ValidacionCierre(ByVal DatosCierre As DataCierre, ByVal services As ServiceProvider) As Boolean
        Dim LngTotal As Integer
        Dim DtCierre As DataTable
        Dim DtNoCerrados As DataTable
        Dim DtAnterior As DataTable
        Const VIEWNAME As String = "vNegCierreInventario"

        If Length(DatosCierre.IDEjercicio) = 0 Then
            ApplicationService.GenerateError("El Ejercicio Contable es obligatorio.")
        Else
            DtCierre = New BE.DataEngine().Filter(VIEWNAME, New StringFilterItem("IDEjercicio", DatosCierre.IDEjercicio), , "FechaCierre DESC, FechaHasta DESC")
            If Not DtCierre Is Nothing AndAlso DtCierre.Rows.Count > 0 Then
                If AreEquals(DtCierre.Rows(0)("EjercicioCerrado"), 1) Then
                    'El ejercicio contable esta cerrado
                    ApplicationService.GenerateError("El Ejercicio Contable está Cerrado. No es posible lanzar el proceso.")
                Else
                    Dim DrMesCier() As DataRow = DtCierre.Select("IDMesCierre=" & DatosCierre.Periodo)
                    If CDbl(DrMesCier.Length) > 0 Then
                        If DrMesCier(0)("PropuestaCorrecta") Then
                            Dim DrCierre() As DataRow = DtCierre.Select("Cerrado=1 AND IDMesCierre=" & DatosCierre.Periodo)
                            If DrCierre.Length <= 0 Then
                                Dim IDEjercicio As String
                                If DatosCierre.Periodo = 1 Then
                                    If Length(DrMesCier(0)("IDEjercicioAnterior")) > 0 Then IDEjercicio = DrMesCier(0)("IDEjercicioAnterior")
                                Else
                                    IDEjercicio = DrMesCier(0)("IDEjercicio")
                                End If
                                If Length(IDEjercicio) > 0 Then
                                    Dim fCerrados As New Filter
                                    fCerrados.Add(New StringFilterItem("IDEjercicio", IDEjercicio))
                                    If DatosCierre.Periodo > 1 Then fCerrados.Add(New NumberFilterItem("IDMesCierre", FilterOperator.LessThan, DatosCierre.Periodo))
                                    DtAnterior = New BE.DataEngine().Filter(VIEWNAME, fCerrados)
                                    If Not DtAnterior Is Nothing AndAlso DtAnterior.Rows.Count > 0 Then
                                        LngTotal = DtAnterior.Rows.Count
                                        Dim DrAnt() As DataRow = DtAnterior.Select("Cerrado=0 OR Cerrado=NULL")
                                        If DrAnt.Length > 0 Then
                                            If LngTotal = DrAnt.Length Then
                                                'Es muy posible que sea el primer cierre de inventario
                                                ApplicationService.GenerateError("No es posible cerrar este período. Es necesario cerrar todos los períodos anteriores. Si es la primera vez que realiza un cierre, no establezca ningún ejercicio anterior para el ejercicio actual.")
                                            Else
                                                'Hay periodos anteriores no cerrados
                                                ApplicationService.GenerateError("No es posible cerrar este período. Es necesario cerrar todos los períodos anteriores.")
                                            End If
                                        End If
                                    Else
                                        ApplicationService.GenerateError("Los periodos del Ejercicio introducido no existen en la Base de Datos.")
                                    End If
                                End If
                                Dim DrDatos() As DataRow = DtCierre.Select("IDMesCierre=" & DatosCierre.Periodo, "Secuencia")
                                If CDbl(DrDatos.Length) <> 1 Then
                                    'Hay periodos anteriores no cerrados
                                    ApplicationService.GenerateError("No es posible cerrar este período. Es necesario cerrar todos los períodos anteriores.")
                                Else
                                    Return True
                                End If
                            Else
                                Dim DrCier2() As DataRow = DtCierre.Select("Cerrado=1 AND IDMesCierre=" & DatosCierre.Periodo)
                                If CDbl(DrCier2.Length) > 0 Then
                                    'El periodo esta cerrado
                                    ApplicationService.GenerateError("El Cierre está Cerrado ")
                                Else
                                    DtNoCerrados = DtCierre.Copy
                                    Dim DrNoCer() As DataRow = DtNoCerrados.Select("(Cerrado= 0 OR Cerrado=NULL) AND IDMesCierre=" & DatosCierre.Periodo, "Secuencia Desc")
                                    If CDbl(DrNoCer.Length) > 0 Then
                                        If DrNoCer(0)("Secuencia") > DtCierre.Rows(0)("Secuencia") Then
                                            If AreEquals(DrNoCer(0)("Secuencia"), DrNoCer(0)("Secuencia") + 1) Then
                                                ValidacionCierre = True
                                            Else
                                                'Hay periodos anteriores no cerrados
                                                ApplicationService.GenerateError("No es posible cerrar este período. Es necesario cerrar todos los períodos anteriores.")
                                            End If
                                        Else
                                            'El periodo esta cerrado
                                            ApplicationService.GenerateError("El Período está Cerrado ")
                                        End If
                                    Else
                                        'El periodo no es valido
                                        ApplicationService.GenerateError("El Período seleccionado no es válido. La Fecha de Cierre es obligatoria.")
                                    End If
                                End If
                            End If
                        Else
                            'La propuesta no es correcta
                            ApplicationService.GenerateError("La propuesta de Cierre no es correcta")
                        End If
                    Else
                        'El periodo no es valido
                        ApplicationService.GenerateError("El Período seleccionado no es válido. La Fecha de Cierre es obligatoria.")
                    End If
                End If
            Else
                ApplicationService.GenerateError("El Ejercicio introducido no existe en la Base de Datos.")
            End If
        End If
    End Function

    <Task()> Public Shared Function ConfirmacionCierre(ByVal DatosCierre As DataCierre, ByVal services As ServiceProvider) As Boolean
        If ProcessServer.ExecuteTask(Of DataCierre, Boolean)(AddressOf ValidacionCierre, DatosCierre, services) Then
            Dim DtCierre As DataTable = New CierreInventario().SelOnPrimaryKey(DatosCierre.IDEjercicio, DatosCierre.Periodo)
            If Not DtCierre Is Nothing AndAlso DtCierre.Rows.Count > 0 Then
                If DtCierre.Rows(0)("PropuestaCorrecta") Then
                    If Not DtCierre.Rows(0)("Cerrado") Then
                        Dim dFecha As Date = Today
                        Dim ClsCierre As Object = BusinessHelper.CreateBusinessObject("Cierre")
                        Dim dtCierrePeriodo As DataTable = ClsCierre.SelOnPrimaryKey(DatosCierre.IDEjercicio, DatosCierre.Periodo)
                        If Not dtCierrePeriodo Is Nothing AndAlso dtCierrePeriodo.Rows.Count > 0 Then
                            If Not IsDBNull(dtCierrePeriodo.Rows(0)("FechaHasta")) Then dFecha = dtCierrePeriodo.Rows(0)("FechaHasta")
                        End If

                        DtCierre.Rows(0)("Cerrado") = True
                        DtCierre.Rows(0)("FechaCierre") = dFecha
                        DatosCierre.FechaHastaPeriodo = dFecha
                        Dim DtArtAlm As DataTable = ProcessServer.ExecuteTask(Of DataCierre, DataTable)(AddressOf PrepararArticuloAlmacen, DatosCierre, services)

                        AdminData.BeginTx()
                        BusinessHelper.UpdateTable(DtArtAlm)
                        BusinessHelper.UpdateTable(DtCierre)
                        Return True
                    End If
                Else
                    ApplicationService.GenerateError("La propuesta de Cierre no es correcta")
                End If
            Else
                ApplicationService.GenerateError("El registro se ha eliminado o no existe.")
            End If
        End If
    End Function

    <Task()> Public Shared Function DeshacerCierre(ByVal DatosCierre As DataCierre, ByVal services As ServiceProvider) As Boolean
        Dim CI As New CierreInventario
        Dim Dt As DataTable = CI.SelOnPrimaryKey(DatosCierre.IDEjercicio, DatosCierre.Periodo)
        If Not Dt Is Nothing AndAlso Dt.Rows.Count > 0 Then
            If Not Dt.Rows(0)("Contabilizado") Then
                If Dt.Rows(0)("Cerrado") Then
                    Dt.Rows(0)("Cerrado") = False
                    CI.Update(Dt)
                    Return True
                End If
            Else
                ApplicationService.GenerateError("El Cierre está Contabilizado")
            End If
        Else
            ApplicationService.GenerateError("El registro se ha eliminado o no existe.")
        End If
    End Function

    <Task()> Public Shared Function DescontabilizarCierre(ByVal DatosCierre As DataCierre, ByVal services As ServiceProvider) As Boolean
        Dim CI As New CierreInventario
        Dim Dt As DataTable = CI.SelOnPrimaryKey(DatosCierre.IDEjercicio, DatosCierre.Periodo)
        If Not Dt Is Nothing AndAlso Dt.Rows.Count > 0 Then
            If Dt.Rows(0)("Contabilizado") Then
                If ProcessServer.ExecuteTask(Of DataCierre, Boolean)(AddressOf ValidarDescontabilizarCierre, DatosCierre, services) Then
                    If Dt.Rows(0)("Cerrado") Then
                        Dt.Rows(0)("Contabilizado") = False
                        'CI.Update(Dt)
                        'Return True
                        Try
                            AdminData.BeginTx()
                            '//Borramos el apunte.
                            Dim f As New Filter
                            f.Add(New StringFilterItem("IDEjercicio", DatosCierre.IDEjercicio))
                            f.Add(New NumberFilterItem("IDTipoApunte", enumDiarioTipoApunte.RegularizacionEx))
                            f.Add(New NumberFilterItem("Mes", DatosCierre.Periodo))
                            Dim dtAsiento As DataTable = AdminData.GetData("tbDiarioContable", f)

                            Dim FinancieroGeneral As IFinanciero = ProcessServer.ExecuteTask(Of Object, IFinanciero)(AddressOf Comunes.CreateFinancieroGeneral, Nothing, services)
                            FinancieroGeneral.DeleteWhere(DatosCierre.IDEjercicio, f)

                            '//Cambiamos el Contabilizado a False.
                            BusinessHelper.UpdateTable(Dt)
                            AdminData.CommitTx(True)
                            Return True
                        Catch ex As Exception
                            AdminData.RollBackTx(True)
                            ApplicationService.GenerateError(ex.Message)
                        End Try
                    End If
                End If
            End If

        Else
            ApplicationService.GenerateError("El registro se ha eliminado o no existe.")
        End If
    End Function

    <Task()> Public Shared Function PropuestaCierre(ByVal DatosCierre As DataCierre, ByVal services As ServiceProvider) As Integer
        Dim DtResultado As DataTable
        Dim BlnExisteCierre As Boolean
        Dim CierreFechas As New DataCierreFechas
        '///Validacion de los datos de entrada
        If Length(DatosCierre.IDEjercicio) = 0 Then
            ApplicationService.GenerateError("El Ejercicio Contable es obligatorio.")
        Else
            '///Inicio proceso de cierre
            '///Validacion del cierre

            CierreFechas.DtCierre = ProcessServer.ExecuteTask(Of DataCierre, DataTable)(AddressOf ValidacionPropuesta, DatosCierre, services)
            If Not CierreFechas.DtCierre Is Nothing Then
                BlnExisteCierre = Not (CierreFechas.DtCierre.Rows.Count = 0)
            End If
            If BlnExisteCierre Then
                CierreFechas.FechaDesde = CierreFechas.DtCierre.Rows(0)("FechaDesde")
                CierreFechas.FechaHasta = CierreFechas.DtCierre.Rows(0)("FechaHasta")
                CierreFechas.DtResultado = DtResultado
                '///Inicializar Detalle del cierre (si existen lineas de detalle)
                CierreFechas.Resultado = ProcessServer.ExecuteTask(Of DataCierre, Integer)(AddressOf PropuestaInicializar, DatosCierre, services)
                If CierreFechas.Resultado = enumstkResultadoCierre.stkRCError Then
                    ApplicationService.GenerateError("Error en el proceso de inicialización del detalle del cierre.")
                Else
                    '///Calcular el stock a fecha
                    CierreFechas = ProcessServer.ExecuteTask(Of DataCierreFechas, DataCierreFechas)(AddressOf PropuestaStockAFecha, CierreFechas, services)
                    If CierreFechas.Resultado = enumstkResultadoCierre.stkRCPasoTerminado Then
                        '///Valoracion del stock
                        CierreFechas.DtResultado = ProcessServer.ExecuteTask(Of DataCierreFechas, DataTable)(AddressOf PropuestaValoracion, CierreFechas, services)
                        '///Actualizar Detalle y Cabecera segun el resultado de cada paso
                        CierreFechas.Resultado = ProcessServer.ExecuteTask(Of DataCierreFechas, Integer)(AddressOf PropuestaActualizar, CierreFechas, services)
                    Else
                        CierreFechas.Resultado = ProcessServer.ExecuteTask(Of DataCierreFechas, Integer)(AddressOf PropuestaActualizar, CierreFechas, services)
                    End If
                End If
            End If
        End If
        Return CierreFechas.Resultado
    End Function

    <Task()> Public Shared Function PropuestaCierreInventario(ByVal data As DataCierre, ByVal services As ServiceProvider) As Integer
        Dim Esquema As String = ProcessServer.ExecuteTask(Of Object, String)(AddressOf Comunes.GetEsquemaBD, Nothing, services)
        Dim Resul As New enumstkResultadoCierre
        'Datos de Cabecera
        Dim DtCierre As DataTable = ProcessServer.ExecuteTask(Of DataCierre, DataTable)(AddressOf ValidacionPropuesta, data, services)
        If Not DtCierre Is Nothing AndAlso DtCierre.Rows.Count > 0 Then
            AdminData.BeginTx()
            '//Eliminamos la propuesta anterior
            If ProcessServer.ExecuteTask(Of DataCierre, Integer)(AddressOf PropuestaInicializar, data, services) = enumstkResultadoCierre.stkRCError Then
                ApplicationService.GenerateError("Error en el proceso de inicialización del detalle del cierre.")
            End If

            Dim DtCabecera As DataTable = New CierreInventario().SelOnPrimaryKey(data.IDEjercicio, data.Periodo)
            If DtCabecera Is Nothing OrElse DtCabecera.Rows.Count = 0 Then
                Dim DrNew As DataRow = DtCabecera.NewRow
                DrNew("IDEjercicio") = data.IDEjercicio
                DrNew("IDMesCierre") = data.Periodo
                DrNew("Cerrado") = False
                DrNew("FechaCierre") = data.FechaHastaPeriodo
                DrNew("Contabilizado") = False
                DrNew("ValorA") = 0
                DrNew("ValorB") = 0
                DrNew("PropuestaCorrecta") = False
                DtCabecera.Rows.Add(DrNew)
            End If
            BusinessHelper.UpdateTable(DtCabecera)

            'TODO HAY QUE BUSCAR EL AÑO CORRESPONDIENTE DEL EJERCICIO SELECCIONADO PARA LA PROPUESTA DE CIERRE
            Dim ClsEjer As BusinessHelper = BusinessHelper.CreateBusinessObject("EjercicioContable")
            Dim DtEjer As DataTable = ClsEjer.SelOnPrimaryKey(data.IDEjercicio)
            If Not DtEjer Is Nothing AndAlso DtEjer.Rows.Count > 0 Then
                Dim DteFechaDesde As Date = DtEjer.Rows(0)("FechaDesde")
                Dim DteFechaHasta As Date = DtEjer.Rows(0)("FechaHasta")
                Dim DteFechaFinal As New Date(DteFechaDesde.Year, data.Periodo, 1)
                data.FechaHastaPeriodo = New Date(DteFechaFinal.Year, DteFechaFinal.Month, DteFechaFinal.AddMonths(1).AddDays(-1).Day)
            End If

            'Datos de Detalle
            Dim CierreInv As New CierreInventarioDetalle
            Dim strSelect As String = "'" & data.IDEjercicio & "' AS IDEjercicio," & data.Periodo & " AS Periodo, "
            
            Dim vStr As String = "SELECT " + strSelect + " vFrmCIValoracionAlmacenFecha.* ,(vFrmCIValoracionAlmacenFecha.Precio * vFrmCIValoracionAlmacenFecha.Acumulado) AS ValorA " + _
                            "FROM vFrmCIValoracionAlmacenFecha RIGHT OUTER JOIN " + _
                           "(SELECT   tbMaestroArticuloAlmacen.IDArticulo, tbMaestroArticuloAlmacen.IDAlmacen," & Esquema & ".fMovimientoValArticulo('" & Format(data.FechaHastaPeriodo, "yyyyMMdd") & "', IDArticulo, IDAlmacen) AS IDLineaMovimiento " + _
                           "  FROM            tbMaestroArticuloAlmacen WHERE " & Esquema & ".fMovimientoValArticulo('" & Format(data.FechaHastaPeriodo, "yyyyMMdd") & "', IDArticulo, IDAlmacen)<>0 ) AS MovimientoArticuloAlmacen ON " + _
                           " vFrmCIValoracionAlmacenFecha.IDLineaMovimiento = MovimientoArticuloAlmacen.IDLineaMovimiento " + _
                           " AND vFrmCIValoracionAlmacenFecha.IDArticulo = MovimientoArticuloAlmacen.IDArticulo " + _
                           " AND vFrmCIValoracionAlmacenFecha.IDAlmacen = MovimientoArticuloAlmacen.IDAlmacen " + _
                           " WHERE vFrmCIValoracionAlmacenFecha.IDLineaMovimiento IS NOT NULL" & _
                            " AND vFrmCIValoracionAlmacenFecha.Empresa = 1"
            If Not data.MostrarObsoletos Then
                vStr &= " AND vFrmCIValoracionAlmacenFecha.Activo = 1"
            End If

            Dim dtDatosValoracionDetalle As DataTable = AdminData.GetData(vStr, String.Empty)
            Dim DetalleCierreInv As List(Of DataRow) = (From c In dtDatosValoracionDetalle Select c).ToList
            For Each detalle As DataRow In DetalleCierreInv

                Dim dtCierreInv As DataTable = CierreInv.AddNew
                Dim drNew As DataRow = dtCierreInv.NewRow
                drNew("IDEjercicio") = detalle("IDEjercicio")
                drNew("IDMesCierre") = detalle("Periodo")
                drNew("IDDetalle") = 0
                drNew("IDArticulo") = detalle("IDArticulo")
                drNew("IDAlmacen") = detalle("IDAlmacen")
                drNew("StockFisico") = detalle("Acumulado")
                drNew("IDUDInterna") = detalle("IDUDInterna")
                drNew("PrecioAlmacenA") = detalle("Precio")
                drNew("ValorA") = detalle("ValorA")
                drNew("PrecioEstandarA") = detalle("PrecioEstandar")
                drNew("PrecioFIFOFechaA") = detalle("FIFOF")
                drNew("PrecioFIFOMvtoA") = detalle("FIFOFD")
                drNew("PrecioMedioA") = detalle("PrecioMedio")
                drNew("PrecioUltimoA") = detalle("PrecioUltimaCompra")
                drNew("FechaCalculo") = detalle("FechaDocumento")
                drNew("PrecioAlmacenB") = 0
                drNew("ValorB") = 0
                drNew("PrecioEstandarB") = 0
                drNew("PrecioFIFOFechaB") = 0
                drNew("PrecioFIFOMvtoB") = 0
                drNew("PrecioMedioB") = 0
                drNew("PrecioUltimoB") = 0
                dtCierreInv.Rows.Add(drNew)

                BusinessHelper.UpdateTable(dtCierreInv)
            Next

    

            'Actualizar los datos de cabecera
            DtCabecera.AcceptChanges()
            If AdminData.Execute("SELECT COUNT(IDDetalle) AS Total From vNegCierreInventarioDetalle WHERE StockFisico < 0 AND IDEjercicio = '" & data.IDEjercicio & "' AND IDMesCierre = " & data.Periodo, ExecuteCommand.ExecuteScalar, False) = 0 Then
                If AdminData.Execute("SELECT COUNT(IDDetalle) AS Total From vNegCierreInventarioDetalle WHERE PrecioAlmacenA < 0 AND IDEjercicio = '" & data.IDEjercicio & "' AND IDMesCierre = " & data.Periodo, ExecuteCommand.ExecuteScalar, False) = 0 Then
                    DtCabecera.Rows(0)("PropuestaCorrecta") = True
                    Resul = enumstkResultadoCierre.stkRCPasoTerminado
                Else : Resul = enumstkResultadoCierre.stkRCPrecioCeroNegativo
                End If
            Else : Resul = enumstkResultadoCierre.stkRCStockNegativo
            End If
            DtCabecera.Rows(0)("ValorA") = Nz(AdminData.Execute("SELECT SUM(ValorA) AS ValorA From tbCierreInventarioDetalle WHERE IDEjercicio = '" & data.IDEjercicio & "' AND IDMesCierre = " & data.Periodo, ExecuteCommand.ExecuteScalar, False), 0)
            DtCabecera.Rows(0)("ValorB") = 0
            BusinessHelper.UpdateTable(DtCabecera)
            AdminData.CommitTx(True)
        End If
        'Controlar el posible resultado a presentación
        Return Resul
    End Function

#End Region

#Region "Funciones Privadas"

    <Task()> Public Shared Function ValidarDeshacerCierre(ByVal DatosCierre As DataCierre, ByVal services As ServiceProvider) As Boolean
        Dim DtCierre As DataTable

        If Length(DatosCierre.IDEjercicio) > 0 Then
            DtCierre = New BE.DataEngine().Filter("vNegCierreInventario", New StringFilterItem("IDEjercicio", DatosCierre.IDEjercicio), , "FechaCierre DESC, FechaHasta DESC")
            If Not DtCierre Is Nothing AndAlso DtCierre.Rows.Count > 0 Then
                Dim DrCier() As DataRow = DtCierre.Select("IDMesCierre=" & DatosCierre.Periodo, "Secuencia")
                If CDbl(DrCier.Length) > 0 Then
                    If Not DrCier(0)("Contabilizado") Then
                        If DrCier(0)("Cerrado") Then
                            DrCier = DtCierre.Select("IDMesCierre>" & DatosCierre.Periodo & " AND Cerrado=1", "Secuencia")
                            If CDbl(DrCier.Length) > 0 Then
                                ApplicationService.GenerateError("No es posible deshacer el Cierre. Existen cierres posteriores en estado Cerrado")
                            Else
                                Return True
                            End If
                        End If
                    Else
                        ApplicationService.GenerateError("El Cierre está Contabilizado")
                    End If
                End If
            End If
        End If
    End Function

    <Task()> Public Shared Function PrepararArticuloAlmacen(ByVal DatosCierre As DataCierre, ByVal services As ServiceProvider) As DataTable
        Dim ClsDetalle As New CierreInventarioDetalle
        Dim ClsArtAlm As New ArticuloAlmacen
        Dim DtDetalle As DataTable
        Dim DtArtAlm As DataTable
        Dim DtCierre As DataTable

        If Length(DatosCierre.IDEjercicio) > 0 Then
            Dim f As New Filter
            f.Add(New StringFilterItem("IDEjercicio", DatosCierre.IDEjercicio))
            f.Add(New NumberFilterItem("IDMesCierre", DatosCierre.Periodo))
            DtDetalle = ClsDetalle.Filter(f)
            If Not DtDetalle Is Nothing AndAlso DtDetalle.Rows.Count > 0 Then
                Dim fArtAlm As New Filter(FilterUnionOperator.Or)
                fArtAlm.Add(New IsNullFilterItem("FechaCalculo", True))
                fArtAlm.Add(New FilterItem("FechaCalculo", FilterOperator.LessThanOrEqual, DatosCierre.FechaHastaPeriodo))
                DtArtAlm = ClsArtAlm.Filter(fArtAlm)
                If Not DtArtAlm Is Nothing AndAlso DtArtAlm.Rows.Count > 0 Then
                    Dim ArtAlmEnCierreDetalle As List(Of AACI) = (From ArtAlm In DtArtAlm Join Det In DtDetalle On _
                                                                        UCase(ArtAlm("IDArticulo")) Equals UCase(Det("IDArticulo")) And _
                                                                         UCase(ArtAlm("IDAlmacen")) Equals UCase(Det("IDAlmacen")) _
                                                                  Order By ArtAlm("IDArticulo"), ArtAlm("IDAlmacen") _
                                                  Select New AACI() With {.IDArticulo = ArtAlm("IDArticulo"), .IDAlmacen = ArtAlm("IDAlmacen"), _
                                                                                      .PrecioMedioAOriginal = ArtAlm("PrecioMedioA"), .PrecioMedioBOriginal = ArtAlm("PrecioMedioB"), _
                                                                                      .PrecioFIFOFechaAOriginal = ArtAlm("PrecioFIFOFechaA"), .PrecioFIFOFechaBOriginal = ArtAlm("PrecioFIFOFechaB"), _
                                                                                      .PrecioFIFOMvtoAOriginal = ArtAlm("PrecioFIFOMvtoA"), .PrecioFIFOMvtoBOriginal = ArtAlm("PrecioFIFOMvtoB"), _
                                                                                      .FechaCalculoOriginal = IIf(Nz(ArtAlm("FechaCalculo"), cnMinDate) = cnMinDate, Nothing, ArtAlm("FechaCalculo")), .StockFechaCalculoOriginal = ArtAlm("StockFechaCalculo"), _
                                                                                      .PrecioMedioA = Det("PrecioMedioA"), .PrecioMedioB = Det("PrecioMedioB"), _
                                                                                      .PrecioFIFOFechaA = Det("PrecioFIFOFechaA"), .PrecioFIFOFechaB = Det("PrecioFIFOFechaB"), _
                                                                                      .PrecioFIFOMvtoA = Det("PrecioFIFOMvtoA"), .PrecioFIFOMvtoB = Det("PrecioFIFOMvtoB"), _
                                                                                      .FechaCalculo = DatosCierre.FechaHastaPeriodo, .StockFechaCalculo = Det("StockFisico")}).ToList
                    If Not ArtAlmEnCierreDetalle Is Nothing AndAlso ArtAlmEnCierreDetalle.Count > 0 Then
                        Dim dtArtAlmAct As New DataTable
                        dtArtAlmAct.TableName = DtArtAlm.TableName

                        For Each field As System.Reflection.PropertyInfo In GetType(AACI).GetProperties
                            If Not Nullable.GetUnderlyingType(field.PropertyType) Is Nothing Then
                                dtArtAlmAct.Columns.Add(field.Name, Nullable.GetUnderlyingType(field.PropertyType))
                            Else
                                dtArtAlmAct.Columns.Add(field.Name, field.PropertyType)
                            End If
                        Next

                        For Each ArtAlm As AACI In ArtAlmEnCierreDetalle
                            Dim row As DataRow = dtArtAlmAct.NewRow()
                            row("IDArticulo") = ArtAlm.IDArticulo
                            row("IDAlmacen") = ArtAlm.IDAlmacen

                            row("PrecioMedioA") = ArtAlm.PrecioMedioAOriginal
                            row("PrecioMedioB") = ArtAlm.PrecioMedioBOriginal

                            row("PrecioFIFOFechaA") = ArtAlm.PrecioFIFOFechaAOriginal
                            row("PrecioFIFOFechaB") = ArtAlm.PrecioFIFOFechaBOriginal

                            row("PrecioFIFOMvtoA") = ArtAlm.PrecioFIFOMvtoAOriginal
                            row("PrecioFIFOMvtoB") = ArtAlm.PrecioFIFOMvtoBOriginal

                            If Nz(ArtAlm.FechaCalculoOriginal, cnMinDate) <> cnMinDate Then row("FechaCalculo") = ArtAlm.FechaCalculoOriginal
                            row("StockFechaCalculo") = ArtAlm.StockFechaCalculoOriginal

                            dtArtAlmAct.Rows.Add(row)
                            row.AcceptChanges()


                            row("PrecioMedioA") = ArtAlm.PrecioMedioA
                            row("PrecioMedioB") = ArtAlm.PrecioMedioB

                            row("PrecioFIFOFechaA") = ArtAlm.PrecioFIFOFechaA
                            row("PrecioFIFOFechaB") = ArtAlm.PrecioFIFOFechaB

                            row("PrecioFIFOMvtoA") = ArtAlm.PrecioFIFOMvtoA
                            row("PrecioFIFOMvtoB") = ArtAlm.PrecioFIFOMvtoB

                            If Nz(ArtAlm.FechaCalculo, cnMinDate) <> cnMinDate Then row("FechaCalculo") = ArtAlm.FechaCalculo
                            row("StockFechaCalculo") = ArtAlm.StockFechaCalculo
                        Next

                        Return dtArtAlmAct
                    End If

                End If
            End If
        End If
    End Function

    Public Class AACI
        Implements IEquatable(Of AACI)


        Public mIDArticulo As String
        Public mIDAlmacen As String

        Public mPrecioMedioAOriginal As Double
        Public mPrecioMedioBOriginal As Double

        Public mPrecioFIFOFechaAOriginal As Double
        Public mPrecioFIFOFechaBOriginal As Double

        Public mPrecioFIFOMvtoAOriginal As Double
        Public mPrecioFIFOMvtoBOriginal As Double

        Public mFechaCalculoOriginal As Date?
        Public mStockFechaCalculoOriginal As Double


        Public mPrecioMedioA As Double
        Public mPrecioMedioB As Double

        Public mPrecioFIFOFechaA As Double
        Public mPrecioFIFOFechaB As Double

        Public mPrecioFIFOMvtoA As Double
        Public mPrecioFIFOMvtoB As Double

        Public mFechaCalculo As Date
        Public mStockFechaCalculo As Double

        Public Property IDArticulo() As String
            Get
                Return mIDArticulo
            End Get
            Set(ByVal value As String)
                mIDArticulo = value
            End Set
        End Property

        Public Property IDAlmacen() As String
            Get
                Return mIDAlmacen
            End Get
            Set(ByVal value As String)
                mIDAlmacen = value
            End Set
        End Property

        Public Property PrecioMedioA() As Double
            Get
                Return mPrecioMedioA
            End Get
            Set(ByVal value As Double)
                mPrecioMedioA = value
            End Set
        End Property
        Public Property PrecioMedioB() As Double
            Get
                Return mPrecioMedioB
            End Get
            Set(ByVal value As Double)
                mPrecioMedioB = value
            End Set
        End Property

        Public Property PrecioFIFOFechaA() As Double
            Get
                Return mPrecioFIFOFechaA
            End Get
            Set(ByVal value As Double)
                mPrecioFIFOFechaA = value
            End Set
        End Property
        Public Property PrecioFIFOFechaB() As Double
            Get
                Return mPrecioFIFOFechaB
            End Get
            Set(ByVal value As Double)
                mPrecioFIFOFechaB = value
            End Set
        End Property

        Public Property PrecioFIFOMvtoA() As Double
            Get
                Return mPrecioFIFOMvtoA
            End Get
            Set(ByVal value As Double)
                mPrecioFIFOMvtoA = value
            End Set
        End Property
        Public Property PrecioFIFOMvtoB() As Double
            Get
                Return mPrecioFIFOMvtoB
            End Get
            Set(ByVal value As Double)
                mPrecioFIFOMvtoB = value
            End Set
        End Property

        Public Property FechaCalculo() As Date
            Get
                Return mFechaCalculo
            End Get
            Set(ByVal value As Date)
                mFechaCalculo = value
            End Set
        End Property

        Public Property StockFechaCalculo() As Double
            Get
                Return mStockFechaCalculo
            End Get
            Set(ByVal value As Double)
                mStockFechaCalculo = value
            End Set
        End Property


        Public Property PrecioMedioAOriginal() As Double
            Get
                Return mPrecioMedioAOriginal
            End Get
            Set(ByVal value As Double)
                mPrecioMedioAOriginal = value
            End Set
        End Property
        Public Property PrecioMedioBOriginal() As Double
            Get
                Return mPrecioMedioBOriginal
            End Get
            Set(ByVal value As Double)
                mPrecioMedioBOriginal = value
            End Set
        End Property

        Public Property PrecioFIFOFechaAOriginal() As Double
            Get
                Return mPrecioFIFOFechaAOriginal
            End Get
            Set(ByVal value As Double)
                mPrecioFIFOFechaAOriginal = value
            End Set
        End Property
        Public Property PrecioFIFOFechaBOriginal() As Double
            Get
                Return mPrecioFIFOFechaBOriginal
            End Get
            Set(ByVal value As Double)
                mPrecioFIFOFechaBOriginal = value
            End Set
        End Property

        Public Property PrecioFIFOMvtoAOriginal() As Double
            Get
                Return mPrecioFIFOMvtoAOriginal
            End Get
            Set(ByVal value As Double)
                mPrecioFIFOMvtoAOriginal = value
            End Set
        End Property
        Public Property PrecioFIFOMvtoBOriginal() As Double
            Get
                Return mPrecioFIFOMvtoBOriginal
            End Get
            Set(ByVal value As Double)
                mPrecioFIFOMvtoBOriginal = value
            End Set
        End Property

        Public Property FechaCalculoOriginal() As Date?
            Get
                Return mFechaCalculoOriginal
            End Get
            Set(ByVal value As Date?)
                mFechaCalculoOriginal = value
            End Set
        End Property

        Public Property StockFechaCalculoOriginal() As Double
            Get
                Return mStockFechaCalculoOriginal
            End Get
            Set(ByVal value As Double)
                mStockFechaCalculoOriginal = value
            End Set
        End Property



        Public Overloads Function Equals(ByVal other As AACI) As Boolean Implements System.IEquatable(Of AACI).Equals
            If other Is Nothing Then
                Return False
            End If
            Return (Me.IDArticulo.Equals(other.IDArticulo)) AndAlso (Me.IDAlmacen.Equals(other.IDAlmacen))
        End Function
    End Class


    <Task()> Public Shared Function ValidarEliminarCierre(ByVal DatosCierre As DataCierre, ByVal services As ServiceProvider) As Boolean
        Dim DtCierre As DataTable

        If Length(DatosCierre.IDEjercicio) > 0 Then
            DtCierre = New BE.DataEngine().Filter("vNegCierreInventario", New StringFilterItem("IDEjercicio", DatosCierre.IDEjercicio), , "FechaCierre DESC, FechaHasta DESC")
            If Not DtCierre Is Nothing AndAlso DtCierre.Rows.Count > 0 Then
                Dim DrDatos() As DataRow = DtCierre.Select("IDMesCierre=" & DatosCierre.Periodo, "Secuencia")
                If CDbl(DrDatos.Length) > 0 Then
                    If Not IsDBNull(DtCierre.Rows(0)("Contabilizado")) AndAlso Not IsDBNull(DtCierre.Rows(0)("Cerrado").Value) Then
                        If AreEquals(DtCierre.Rows(0)("Contabilizado"), 1) Then
                            ApplicationService.GenerateError("El Cierre está Contabilizado")
                        Else
                            Dim DrCier() As DataRow = DtCierre.Select("Cerrado=1")
                            If CDbl(DrCier.Length) > 0 Then
                                ApplicationService.GenerateError("No es posible eliminar el Cierre. Existen cierres posteriores en curso o en estado Cerrado")
                            Else
                                Return True
                            End If
                        End If
                    Else
                        'No hay nada que eliminar:no hay cebecera de cierre (tampoco habra lineas de detalle)
                        Return False
                    End If
                Else
                    Return False
                End If
            Else
                ApplicationService.GenerateError("El registro se ha eliminado o no existe.")
            End If
        Else
            ApplicationService.GenerateError("El Ejercicio Contable es obligatorio.")
        End If

        Exit Function
    End Function

    <Task()> Public Shared Function ValidarDescontabilizarCierre(ByVal DatosCierre As DataCierre, ByVal services As ServiceProvider) As Boolean
        Dim DtCierre As DataTable

        If Length(DatosCierre.IDEjercicio) > 0 Then
            DtCierre = New BE.DataEngine().Filter("vNegCierreInventario", New StringFilterItem("IDEjercicio", DatosCierre.IDEjercicio), , "FechaCierre DESC, FechaHasta DESC")
            If Not DtCierre Is Nothing AndAlso DtCierre.Rows.Count > 0 Then
                If AreEquals(DtCierre.Rows(0)("EjercicioCerrado"), 0) Then
                    Dim DrCier() As DataRow = DtCierre.Select("Cerrado=1")
                    If DrCier.Length > 0 Then
                        Dim DrDatos() As DataRow = DtCierre.Select("Cerrado=1 AND IDMesCierre=" & DatosCierre.Periodo, "FechaHasta DESC")
                        If CDbl(DrDatos.Length) = 0 Then
                            ApplicationService.GenerateError("El Cierre no está cerrado")
                        End If
                        DrDatos = DtCierre.Select("Cerrado=1 AND IDMesCierre>" & DatosCierre.Periodo & " AND Contabilizado=1", "FechaHasta DESC")
                        If CDbl(DrDatos.Length) > 0 Then
                            ApplicationService.GenerateError("No es posible descontabilizar el cierre. Existen períodos posteriores contabilizados")
                        Else
                            Return True
                        End If
                    End If
                Else
                    ApplicationService.GenerateError("El ejercicio está cerrado.No se pueden realizar cambios.")
                End If
            End If
        End If
    End Function

    <Task()> Public Shared Function PropuestaInicializar(ByVal DatosCierre As DataCierre, ByVal services As ServiceProvider) As Integer
        Dim StrWhere As String
        StrWhere = " WHERE IDEjercicio='" & DatosCierre.IDEjercicio & "' AND IDMesCierre=" & DatosCierre.Periodo
        Dim sql As String = "Delete from tbCierreInventarioDetalle" & StrWhere
        AdminData.Execute(sql)
        Return enumstkResultadoCierre.stkRCPasoTerminado
    End Function

    'Private Function PropuestaStockAFecha(ByVal DteFechaHasta As Date, ByRef DtResultado As DataTable) As Integer
    '    Const VIEW_NAME As String = "vNegCierreInventarioDatos"
    '    Dim strCommand As String

    '    Dim DtAux As DataTable
    '    Dim services As New ServiceProvider
    '    Dim DtDatos As DataTable = AdminData.GetData(VIEW_NAME)
    '    If Not DtDatos Is Nothing AndAlso DtDatos.Rows.Count > 0 Then
    '        Dim stockAFecha() As StockAFechaInfo = ProcesoStocks.StockAFecha(DteFechaHasta, services)
    '        If Not stockAFecha Is Nothing AndAlso stockAFecha.Length > 0 Then
    '            If IsNothing(DtResultado) Then DtResultado = New DataTable
    '            DtResultado.Columns.Add("IDArticulo", GetType(String))
    '            DtResultado.Columns.Add("IDAlmacen", GetType(String))
    '            DtResultado.Columns.Add("StockFisico", GetType(Double))
    '            For Each stock As StockAFechaInfo In stockAFecha
    '                Dim nr As DataRow = DtResultado.NewRow()
    '                nr("IDArticulo") = stock.IDArticulo
    '                nr("IDAlmacen") = stock.IDAlmacen
    '                nr("StockFisico") = stock.StockAFecha
    '                DtResultado.Rows.Add(nr)
    '            Next
    '        End If
    '        If DtResultado Is Nothing Then
    '            ApplicationService.GenerateError("No hay resultados en el cálculo del stock a fecha.")
    '        Else

    '            DtAux = DtDatos.Clone
    '            DtAux.Columns.Add("StockFisico", GetType(Double))
    '            DtAux.Columns.Add("Cantidad", GetType(Double))

    '            For Each Dr As DataRow In DtDatos.Select
    '                Dim DrNew As DataRow = DtAux.NewRow()
    '                For Each Dc As DataColumn In DtDatos.Columns
    '                    If Not IsDBNull(Dr(Dc.ColumnName)) Then
    '                        DrNew(Dc.ColumnName) = Dr(Dc.ColumnName)
    '                    Else
    '                        If DtAux.Columns(Dc.ColumnName).DataType Is GetType(Integer) Then
    '                            DrNew(Dc.ColumnName) = Dr(Dc.ColumnName)
    '                        ElseIf DtAux.Columns(Dc.ColumnName).DataType Is GetType(String) Then
    '                            If DtAux.Columns(Dc.ColumnName).AllowDBNull Then
    '                                DrNew(Dc.ColumnName) = System.DBNull.Value
    '                            End If
    '                        End If
    '                    End If
    '                Next
    '                DtAux.Rows.Add(DrNew)
    '                If DtResultado.Rows.Count > 0 Then
    '                    Dim DrResul() As DataRow = DtResultado.Select("IDArticulo='" & Dr("IDArticulo") & "' AND IDAlmacen='" & Dr("IDAlmacen") & "'")
    '                    If DrResul.Length > 0 Then
    '                        DrNew("StockFisico") = DrResul(0)("StockFisico")
    '                        DrNew("Cantidad") = DrResul(0)("StockFisico")
    '                    Else
    '                        DrNew("StockFisico") = 0
    '                        DrNew("Cantidad") = 0
    '                    End If
    '                Else
    '                    DrNew("StockFisico") = 0
    '                End If
    '            Next
    '            If DtAux.Rows.Count > 0 Then
    '                Dim DrAux() As DataRow = DtAux.Select("StockFisico < 0 AND Activo = 1")
    '                DtResultado = DtAux
    '                If DrAux.Length > 0 Then
    '                    Return enumstkResultadoCierre.stkRCStockNegativo
    '                Else
    '                    Return enumstkResultadoCierre.stkRCPasoTerminado
    '                End If
    '            End If
    '        End If
    '    End If
    'End Function

    <Task()> Public Shared Function PropuestaStockAFecha(ByVal CierreFecha As DataCierreFechas, ByVal services As ServiceProvider) As DataCierreFechas
        Const VIEW_NAME As String = "vNegCierreInventarioDatos"
        Dim dtDatos As DataTable = New BE.DataEngine().Filter(VIEW_NAME, "*", "")
        If Not dtDatos Is Nothing AndAlso dtDatos.Rows.Count > 0 Then
            Dim dtDatosStockCantidad As DataTable = dtDatos.Clone
            dtDatosStockCantidad.Columns.Add("StockFisico", GetType(Double))
            dtDatosStockCantidad.Columns.Add("Cantidad", GetType(Double))
            dtDatosStockCantidad.Columns.Add("StockFisico2", GetType(Double))
            dtDatosStockCantidad.Columns.Add("Cantidad2", GetType(Double))

            For Each Dr As DataRow In dtDatos.Select
                Dim DrNew As DataRow = dtDatosStockCantidad.NewRow()
                For Each Dc As DataColumn In dtDatos.Columns
                    If Not IsDBNull(Dr(Dc.ColumnName)) Then
                        DrNew(Dc.ColumnName) = Dr(Dc.ColumnName)
                    Else
                        If dtDatosStockCantidad.Columns(Dc.ColumnName).DataType Is GetType(Integer) Then
                            DrNew(Dc.ColumnName) = Dr(Dc.ColumnName)
                        ElseIf dtDatosStockCantidad.Columns(Dc.ColumnName).DataType Is GetType(String) Then
                            If dtDatosStockCantidad.Columns(Dc.ColumnName).AllowDBNull Then
                                DrNew(Dc.ColumnName) = System.DBNull.Value
                            End If
                        End If
                    End If
                Next
                dtDatosStockCantidad.Rows.Add(DrNew)

                Dim datosStock As New DataArticuloAlmacenFecha(Dr("IDArticulo"), Dr("IDAlmacen"), CierreFecha.FechaHasta)
                Dim stk As StockAFechaInfo = ProcessServer.ExecuteTask(Of DataArticuloAlmacenFecha, StockAFechaInfo)(AddressOf ProcesoStocks.GetStockAcumuladoAFecha, datosStock, services)
                If Not stk Is Nothing Then
                    DrNew("StockFisico") = stk.StockAFecha
                    DrNew("Cantidad") = stk.StockAFecha
                    DrNew("Cantidad2") = 0
                    DrNew("StockFisico2") = 0
                    If ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, Dr("IDArticulo"), services) Then
                        DrNew("StockFisico2") = stk.StockAFecha2
                        DrNew("Cantidad2") = stk.StockAFecha2
                    End If
                Else
                    DrNew("Cantidad") = 0
                    DrNew("StockFisico") = 0
                    DrNew("Cantidad2") = 0
                    DrNew("StockFisico2") = 0
                End If
            Next
            If dtDatosStockCantidad.Rows.Count > 0 Then
                Dim AppParams As ParametroStocks = services.GetService(Of ParametroStocks)()
                Dim f As New Filter
                f.Add(New BooleanFilterItem("Activo", True))
                Dim fStockNegativo As New Filter(FilterUnionOperator.Or)
                fStockNegativo.Add(New NumberFilterItem("StockFisico", FilterOperator.LessThan, 0))
                If AppParams.GestionDobleUnidad Then
                    Dim f2UD As New Filter
                    f2UD.Add(New IsNullFilterItem("IDUDInterna2", False))
                    f2UD.Add(New NumberFilterItem("StockFisico2", FilterOperator.LessThan, 0))

                    fStockNegativo.Add(f2UD)
                End If
                f.Add(fStockNegativo)
                Dim WhereValidacion As String = f.Compose(New AdoFilterComposer)

                Dim DrAux() As DataRow = dtDatosStockCantidad.Select(WhereValidacion)
                CierreFecha.DtResultado = dtDatosStockCantidad
                If DrAux.Length > 0 Then
                    CierreFecha.Resultado = enumstkResultadoCierre.stkRCStockNegativo
                Else
                    CierreFecha.Resultado = enumstkResultadoCierre.stkRCPasoTerminado
                End If
                Return CierreFecha
            End If
        End If
    End Function

    <Task()> Public Shared Function PropuestaValoracion(ByVal CierreFecha As DataCierreFechas, ByVal services As ServiceProvider) As DataTable
        Dim precios As ValoracionPreciosInfo
        Dim BlnPrecNoValido, BlnActivo As Boolean
        Dim DblFIFOFechaA, DblFIFOFechaB, DblFIFOMvtoA, DblFIFOMvtoB As Double
        Dim DblPrecioMedioA, DblPrecioMedioB As Double
        Dim LngUdValoracion As Integer

        Dim DtDetalle As New DataTable
        DtDetalle.RemotingFormat = SerializationFormat.Binary
        DtDetalle.Columns.Add("IDArticulo", GetType(String))
        DtDetalle.Columns.Add("IDAlmacen", GetType(String))
        DtDetalle.Columns.Add("IDUDInterna", GetType(String))
        DtDetalle.Columns.Add("IDUDInterna2", GetType(String))
        DtDetalle.Columns.Add("StockFisico", GetType(Double))
        DtDetalle.Columns.Add("StockFisico2", GetType(Double))
        DtDetalle.Columns.Add("PrecioAlmacenA", GetType(Double))
        DtDetalle.Columns.Add("PrecioAlmacenB", GetType(Double))
        DtDetalle.Columns.Add("ValorA", GetType(Double))
        DtDetalle.Columns.Add("ValorB", GetType(Double))
        DtDetalle.Columns.Add("PrecioEstandarA", GetType(Double))
        DtDetalle.Columns.Add("PrecioEstandarB", GetType(Double))
        DtDetalle.Columns.Add("PrecioFIFOFechaA", GetType(Double))
        DtDetalle.Columns.Add("PrecioFIFOFechaB", GetType(Double))
        DtDetalle.Columns.Add("PrecioFIFOMvtoA", GetType(Double))
        DtDetalle.Columns.Add("PrecioFIFOMvtoB", GetType(Double))
        DtDetalle.Columns.Add("PrecioMedioA", GetType(Double))
        DtDetalle.Columns.Add("PrecioMedioB", GetType(Double))
        DtDetalle.Columns.Add("PrecioUltimoA", GetType(Double))
        DtDetalle.Columns.Add("PrecioUltimoB", GetType(Double))
        DtDetalle.Columns.Add("FechaCalculo", GetType(Date))
        For Each Dr As DataRow In CierreFecha.DtResultado.Select
            Dim DrNew As DataRow = DtDetalle.NewRow
            DrNew("IDArticulo") = Dr("IDArticulo")
            DrNew("IDAlmacen") = Dr("IDAlmacen")
            DrNew("StockFisico") = Dr("StockFisico")
            If Dr.Table.Columns.Contains("StockFisico2") AndAlso Length(Dr("StockFisico2")) > 0 Then DrNew("StockFisico2") = Dr("StockFisico2")
            DrNew("IDUDInterna") = Dr("IDUDInterna")
            If Dr.Table.Columns.Contains("IDUDInterna2") AndAlso Length(Dr("IDUDInterna2")) > 0 Then DrNew("IDUDInterna2") = Dr("IDUDInterna2")
            DrNew("FechaCalculo") = CierreFecha.FechaHasta
            DblFIFOFechaA = 0 : DblFIFOFechaB = 0
            DblFIFOMvtoA = 0 : DblFIFOMvtoB = 0
            DblPrecioMedioA = 0 : DblPrecioMedioB = 0
            If IsDBNull(Dr("Activo")) Then
                BlnActivo = False
            Else
                BlnActivo = Dr("Activo")
            End If
            If BlnActivo Then
                LngUdValoracion = IIf(Dr("UdValoracion") > 0, Dr("UdValoracion"), 1)
                DrNew("PrecioEstandarA") = Dr("PrecioEstandarA") / LngUdValoracion
                DrNew("PrecioEstandarB") = Dr("PrecioEstandarB") / LngUdValoracion
                Dim datosPrecio As New ProcesoStocks.DataValoracionFIFO(DrNew("IDArticulo"), DrNew("IDAlmacen"), DrNew("StockFisico"), CierreFecha.FechaHasta, enumstkValoracionFIFO.stkVFOrdenarPorFecha)
                precios = ProcessServer.ExecuteTask(Of ProcesoStocks.DataValoracionFIFO, ValoracionPreciosInfo)(AddressOf ProcesoStocks.ValoracionFIFO, datosPrecio, services)
                If Not precios Is Nothing Then
                    DblFIFOFechaA = precios.PrecioA : DblFIFOFechaB = precios.PrecioB
                End If
                DrNew("PrecioFIFOFechaA") = DblFIFOFechaA
                DrNew("PrecioFIFOFechaB") = DblFIFOFechaB
                datosPrecio = New ProcesoStocks.DataValoracionFIFO(DrNew("IDArticulo"), DrNew("IDAlmacen"), DrNew("StockFisico"), CierreFecha.FechaHasta, enumstkValoracionFIFO.stkVFOrdenarPorMvto)
                precios = ProcessServer.ExecuteTask(Of ProcesoStocks.DataValoracionFIFO, ValoracionPreciosInfo)(AddressOf ProcesoStocks.ValoracionFIFO, datosPrecio, services)

                If Not precios Is Nothing Then
                    DblFIFOMvtoA = precios.PrecioA : DblFIFOMvtoB = precios.PrecioB
                End If
                DrNew("PrecioFIFOMvtoA") = DblFIFOMvtoA
                DrNew("PrecioFIFOMvtoB") = DblFIFOMvtoB
                'ValoresInicialesPrecioMedio(DrNew("IDArticulo"), DrNew("IDAlmacen"), DteFechaDesde, DblStockInicial, DblPrecioInicialA, DblPrecioInicialB)
                Dim ArtAlmFechaHasta As New DataArticuloAlmacenFecha(DrNew("IDArticulO"), DrNew("IDAlmacen"), CierreFecha.FechaHasta)
                precios = ProcessServer.ExecuteTask(Of DataArticuloAlmacenFecha, ValoracionPreciosInfo)(AddressOf ProcesoStocks.ValoracionPrecioMedioAFecha, ArtAlmFechaHasta, services)

                If Not precios Is Nothing Then
                    DblPrecioMedioA = precios.PrecioA : DblPrecioMedioB = precios.PrecioB
                End If
                DrNew("PrecioMedioA") = DblPrecioMedioA
                DrNew("PrecioMedioB") = DblPrecioMedioB
                DrNew("PrecioUltimoA") = Dr("PrecioUltimaCompraA")
                DrNew("PrecioUltimoB") = Dr("PrecioUltimaCompraB")

                ' ///Asignar el precio de almacen de acuerdo al criterio de valoracion del articulo
                Select Case CType(Dr("CriterioValoracion"), enumtaValoracion)
                    Case enumtaValoracion.taPrecioEstandar
                        DrNew("PrecioAlmacenA") = DrNew("PrecioEstandarA")
                        DrNew("PrecioAlmacenB") = DrNew("PrecioEstandarB")
                    Case enumtaValoracion.taPrecioFIFOFecha
                        DrNew("PrecioAlmacenA") = DrNew("PrecioFIFOFechaA")
                        DrNew("PrecioAlmacenB") = DrNew("PrecioFIFOFechaB")
                    Case enumtaValoracion.taPrecioFIFOMvto
                        DrNew("PrecioAlmacenA") = DrNew("PrecioFIFOMvtoA")
                        DrNew("PrecioAlmacenB") = DrNew("PrecioFIFOMvtoB")
                    Case enumtaValoracion.taPrecioMedio
                        DrNew("PrecioAlmacenA") = DrNew("PrecioMedioA")
                        DrNew("PrecioAlmacenB") = DrNew("PrecioMedioB")
                    Case enumtaValoracion.taPrecioUltCompra
                        DrNew("PrecioAlmacenA") = DrNew("PrecioUltimoA")
                        DrNew("PrecioAlmacenB") = DrNew("PrecioUltimoB")
                End Select

                If Not BlnPrecNoValido Then BlnPrecNoValido = (DrNew("PrecioAlmacenA") < 0)
                If DrNew("StockFisico") > 0 Then
                    DrNew("ValorA") = DrNew("PrecioAlmacenA") * DrNew("StockFisico")
                    DrNew("ValorB") = DrNew("PrecioAlmacenB") * DrNew("StockFisico")
                Else
                    DrNew("ValorA") = 0 : DrNew("ValorB") = 0
                End If
            Else
                DrNew("PrecioEstandarA") = 0 : DrNew("PrecioEstandarB") = 0
                DrNew("PrecioFIFOFechaA") = 0 : DrNew("PrecioFIFOFechaB") = 0
                DrNew("PrecioFIFOMvtoA") = 0 : DrNew("PrecioFIFOMvtoB") = 0
                DrNew("PrecioMedioA") = 0 : DrNew("PrecioMedioB") = 0
                DrNew("PrecioUltimoA") = 0 : DrNew("PrecioUltimoB") = 0
                DrNew("PrecioAlmacenA") = 0 : DrNew("PrecioAlmacenB") = 0
                DrNew("ValorA") = 0 : DrNew("ValorB") = 0
            End If
            DtDetalle.Rows.Add(DrNew)
        Next
        If BlnPrecNoValido Then
            CierreFecha.Resultado = enumstkResultadoCierre.stkRCPrecioCeroNegativo
        Else
            CierreFecha.Resultado = enumstkResultadoCierre.stkRCPasoTerminado
        End If
        Return DtDetalle
    End Function

    <Task()> Public Shared Function PropuestaActualizar(ByVal CierreFecha As DataCierreFechas, ByVal services As ServiceProvider) As Integer
        Dim DblValorTotalA, DblValorTotalB As Double
        Dim DtDetalle As DataTable
        Dim ClsCID As New CierreInventarioDetalle
        Dim CI As New CierreInventario
        If Not CierreFecha.DtCierre Is Nothing AndAlso CierreFecha.DtCierre.Rows.Count > 0 Then
            Dim DtCabecera As DataTable = CI.SelOnPrimaryKey(CierreFecha.DtCierre.Rows(0)("IDEjercicio"), CierreFecha.DtCierre.Rows(0)("idmescierre"))
            If Not DtCabecera Is Nothing Then
                If DtCabecera.Rows.Count = 0 Then
                    DtCabecera = CI.AddNew()
                    Dim DrNew As DataRow = DtCabecera.NewRow
                    For Each Dc As DataColumn In DtCabecera.Columns
                        DrNew(Dc.ColumnName) = CierreFecha.DtCierre.Rows(0)(Dc.ColumnName)
                    Next
                    DtCabecera.Rows.Add(DrNew)
                End If
            End If
            If Not CierreFecha.DtResultado Is Nothing Then
                If CierreFecha.DtResultado.Rows.Count > 0 Then
                    DtDetalle = ClsCID.AddNew()
                    Select Case CierreFecha.Resultado
                        Case enumstkResultadoCierre.stkRCStockNegativo
                            For Each Dr As DataRow In CierreFecha.DtResultado.Select
                                Dim DrNew As DataRow = DtDetalle.NewRow()
                                DrNew("IDDetalle") = AdminData.GetAutoNumeric
                                DrNew("IDEjercicio") = CierreFecha.DtCierre.Rows(0)("IDEjercicio")
                                DrNew("idmescierre") = CierreFecha.DtCierre.Rows(0)("idmescierre")
                                DrNew("IDArticulo") = Dr("IDArticulo")
                                DrNew("IDAlmacen") = Dr("IDAlmacen")
                                DrNew("StockFisico") = Dr("StockFisico")
                                DrNew("IDUDInterna") = Dr("IDUDInterna")
                                If DrNew.Table.Columns.Contains("StockFisico2") AndAlso Length(Dr("StockFisico2")) > 0 Then DrNew("StockFisico2") = Dr("StockFisico2")
                                If DrNew.Table.Columns.Contains("IDUDInterna2") AndAlso Length(Dr("IDUDInterna2")) > 0 Then DrNew("IDUDInterna2") = Dr("IDUDInterna2")
                                DrNew("FechaCalculo") = CierreFecha.FechaHasta
                                DtDetalle.Rows.Add(DrNew)
                            Next

                        Case enumstkResultadoCierre.stkRCPrecioCeroNegativo, enumstkResultadoCierre.stkRCPasoTerminado
                            For Each Dr As DataRow In CierreFecha.DtResultado.Select
                                Dim DrNew As DataRow = DtDetalle.NewRow()
                                DrNew("IDDetalle") = AdminData.GetAutoNumeric
                                DrNew("IDEjercicio") = CierreFecha.DtCierre.Rows(0)("IDEjercicio")
                                DrNew("idmescierre") = CierreFecha.DtCierre.Rows(0)("idmescierre")
                                For Each Dc As DataColumn In CierreFecha.DtResultado.Columns
                                    If DrNew.Table.Columns.Contains(Dc.ColumnName) Then
                                        DrNew(Dc.ColumnName) = Dr(Dc.ColumnName)
                                    End If
                                    If CierreFecha.Resultado = enumstkResultadoCierre.stkRCPasoTerminado Then
                                        If Dc.ColumnName = "ValorA" Then
                                            DblValorTotalA += Dr(Dc.ColumnName)
                                        ElseIf Dc.ColumnName = "ValorB" Then
                                            DblValorTotalB += Dr(Dc.ColumnName)
                                        End If
                                    End If
                                Next
                                DrNew("FechaCalculo") = CierreFecha.FechaHasta
                                DtDetalle.Rows.Add(DrNew)
                            Next
                            If CierreFecha.Resultado = enumstkResultadoCierre.stkRCPasoTerminado Then
                                DtCabecera.Rows(0)("PropuestaCorrecta") = True
                                DtCabecera.Rows(0)("ValorA") = DblValorTotalA
                                DtCabecera.Rows(0)("ValorB") = DblValorTotalB
                            End If
                    End Select
                End If
            End If
            BusinessHelper.UpdateTable(DtCabecera)
            BusinessHelper.UpdateTable(DtDetalle)
        End If
        If CierreFecha.Resultado = enumstkResultadoCierre.stkRCError Then
            ApplicationService.GenerateError("Error en el proceso.")
        Else
            Return enumstkResultadoCierre.stkRCPasoTerminado
        End If
    End Function

    'Private Function ValoresInicialesPrecioMedio(ByVal StrArticulo As String, _
    '                                             ByVal StrAlmacen As String, _
    '                                             ByVal DteFechaDesde As Date, _
    '                                             ByRef DblStockInicial As Double, _
    '                                             ByRef DblPrecioInicialA As Double, _
    '                                             ByRef DblPrecioInicialB As Double) As Date
    '    Dim DtAux As DataTable

    '    If (Length(StrArticulo) * Length(StrAlmacen)) > 0 Then
    '        Dim f As New Filter
    '        f.Add(New StringFilterItem("IDArticulo", StrArticulo))
    '        f.Add(New StringFilterItem("IDAlmacen", StrAlmacen))
    '        If DteFechaDesde <> System.DateTime.MinValue Then f.Add(New DateFilterItem("FechaHasta", FilterOperator.LessThan, DteFechaDesde))
    '        DtAux = AdminData.GetData("vNegPrecioUltimoCierre", f, "TOP 1 *", "FechaDesde DESC, FechaHasta DESC, Secuencia DESC")
    '        If Not DtAux Is Nothing AndAlso DtAux.Rows.Count > 0 Then
    '            DblStockInicial = DtAux.Rows(0)("StockFisico")
    '            DblPrecioInicialA = DtAux.Rows(0)("PrecioAlmacenA")
    '            DblPrecioInicialB = DtAux.Rows(0)("PrecioAlmacenB")
    '            If Not IsDBNull(DtAux.Rows(0)("FechaHasta")) Then ValoresInicialesPrecioMedio = DtAux.Rows(0)("FechaHasta")
    '        End If
    '    End If
    'End Function

#End Region

End Class
