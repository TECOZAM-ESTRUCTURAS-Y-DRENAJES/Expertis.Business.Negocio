Public Class PresupuestoCosteEstandar
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbPresupuestoCosteEstandar"

#Region " RegisterAddNewTasks "

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarIdentificador, data, services)
    End Sub

    <Task()> Public Shared Sub AsignarIdentificador(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDPresupuesto")) = 0 Then data("IDPresupuesto") = AdminData.GetAutoNumeric
    End Sub

#End Region

#Region " BusinessRules "

    ''' <summary>
    ''' Reglas de negocio. Lista de tareas asociadas a cambios.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>Solo se establece la lista en este punto no se ejecutan</remarks>
    Public Overrides Function GetBusinessRules() As Solmicro.Expertis.Engine.BE.BusinessRules
        Dim oBRL As New BusinessRules
        oBRL.Add("IDEmpresa", AddressOf CambioEmpresa)
        oBRL.Add("IDCliente", AddressOf CambioCliente)
        oBRL.Add("CosteMatStdA", AddressOf RecalcularImportes)
        oBRL.Add("CosteOpeStdA", AddressOf RecalcularImportes)
        oBRL.Add("CosteExtStdA", AddressOf RecalcularImportes)
        oBRL.Add("CosteVarStdA", AddressOf RecalcularImportes)
        oBRL.Add("PorcentajeMat", AddressOf RecalcularImportes)
        oBRL.Add("PorcentajeOpe", AddressOf RecalcularImportes)
        oBRL.Add("PorcentajeExt", AddressOf RecalcularImportes)
        oBRL.Add("PorcentajeVar", AddressOf RecalcularImportes)
        oBRL.Add("PorcentajeBeneficio", AddressOf RecalcularImportes)
        oBRL.Add("PVPA", AddressOf RecalcularImportes)

        Return oBRL
    End Function

    <Task()> Public Shared Sub CambioEmpresa(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim f As New Filter
            f.Add(New StringFilterItem("IDEmpresa", data.Value))
            Dim dtEmpresa As DataTable = New BE.DataEngine().Filter("tbMaestroEmpresa", f)
            If Not dtEmpresa Is Nothing AndAlso dtEmpresa.Rows.Count > 0 Then
                data.Current("IDCliente") = System.DBNull.Value
                data.Current("IDEmpresa") = data.Value

                data.Current("Direccion") = dtEmpresa.Rows(0)("Direccion")
                data.Current("CodPostal") = dtEmpresa.Rows(0)("CodPostal")
                data.Current("Poblacion") = dtEmpresa.Rows(0)("Poblacion")
                data.Current("Provincia") = dtEmpresa.Rows(0)("Provincia")
                data.Current("IDPais") = dtEmpresa.Rows(0)("IDPais")
                data.Current("Telefono") = dtEmpresa.Rows(0)("Telefono1")
                data.Current("Fax") = dtEmpresa.Rows(0)("Fax")
                data.Current("Email") = dtEmpresa.Rows(0)("Email")

                Dim dtCliente As DataTable = New Cliente().Filter(f)
                If dtCliente.Rows.Count > 0 Then
                    data.Current("IDCliente") = dtCliente.Rows(0)("IDCliente")
                    data.Current = New PresupuestoCosteEstandar().ApplyBusinessRule("IDCliente", data.Current("IDCliente"), data.Current, data.Context)
                End If
            End If
        Else
            data.Current("Direccion") = DBNull.Value
            data.Current("CodPostal") = DBNull.Value
            data.Current("Poblacion") = DBNull.Value
            data.Current("Provincia") = DBNull.Value
            data.Current("IDPais") = DBNull.Value
            data.Current("Telefono") = DBNull.Value
            data.Current("Fax") = DBNull.Value
            data.Current("Email") = DBNull.Value
            data.Current("IDPersona") = DBNull.Value
        End If
    End Sub

    <Task()> Public Shared Sub CambioCliente(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
            Dim Cliente As ClienteInfo = Clientes.GetEntity(data.Value)

            data.Current("IDEmpresa") = Cliente.IDEmpresa

            Dim IDCliente As String = data.Value
            If Cliente.GrupoDireccion Then IDCliente = Cliente.GrupoCliente
            Dim StDatos As New ClienteDireccion.DataDirecEnvio(IDCliente, enumcdTipoDireccion.cdDireccionEnvio)
            Dim dtDireccion As DataTable = ProcessServer.ExecuteTask(Of ClienteDireccion.DataDirecEnvio, DataTable)(AddressOf ClienteDireccion.ObtenerDireccionEnvio, StDatos, services)
            If Not dtDireccion Is Nothing AndAlso dtDireccion.Rows.Count > 0 Then
                data.Current("Direccion") = dtDireccion.Rows(0)("Direccion")
                data.Current("CodPostal") = dtDireccion.Rows(0)("CodPostal")
                data.Current("Poblacion") = dtDireccion.Rows(0)("Poblacion")
                data.Current("Provincia") = dtDireccion.Rows(0)("Provincia")
                data.Current("IDPais") = dtDireccion.Rows(0)("IDPais")
                data.Current("Telefono") = dtDireccion.Rows(0)("Telefono")
                data.Current("Fax") = dtDireccion.Rows(0)("Fax")
                data.Current("Email") = dtDireccion.Rows(0)("Email")
                data.Current("IDPersona") = DBNull.Value
            Else
                data.Current("Direccion") = DBNull.Value
                data.Current("CodPostal") = DBNull.Value
                data.Current("Poblacion") = DBNull.Value
                data.Current("Provincia") = DBNull.Value
                data.Current("IDPais") = DBNull.Value
                data.Current("Telefono") = DBNull.Value
                data.Current("Fax") = DBNull.Value
                data.Current("Email") = DBNull.Value
                data.Current("IDPersona") = DBNull.Value
                ApplicationService.GenerateError("Este Cliente no tiene una direccion predeterminada. Debe de crear una en el mantenimiento de Clientes.")
            End If
        Else
            data.Current("Direccion") = DBNull.Value
            data.Current("CodPostal") = DBNull.Value
            data.Current("Poblacion") = DBNull.Value
            data.Current("Provincia") = DBNull.Value
            data.Current("IDPais") = DBNull.Value
            data.Current("Telefono") = DBNull.Value
            data.Current("Fax") = DBNull.Value
            data.Current("Email") = DBNull.Value
            data.Current("IDPersona") = DBNull.Value
        End If
    End Sub

    <Task()> Public Shared Sub RecalcularImportes(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        Dim dblCosteTotal As Double = Nz(data.Current("CosteMatStdA"), 0) + Nz(data.Current("CosteOpeStdA"), 0) + Nz(data.Current("CosteExtStdA"), 0) + Nz(data.Current("CosteVarStdA"), 0)
        data.Current("CosteStdA") = dblCosteTotal

        Dim dataAplicarPorcMat As New Comunes.datosAplicarMargen(Nz(data.Current("CosteMatStdA"), 0), Nz(data.Current("PorcentajeMat"), 0))
        ProcessServer.ExecuteTask(Of Comunes.datosAplicarMargen)(AddressOf Comunes.AplicarMargen, dataAplicarPorcMat, services)

        Dim dataAplicarPorcOpe As New Comunes.datosAplicarMargen(Nz(data.Current("CosteOpeStdA"), 0), Nz(data.Current("PorcentajeOpe"), 0))
        ProcessServer.ExecuteTask(Of Comunes.datosAplicarMargen)(AddressOf Comunes.AplicarMargen, dataAplicarPorcOpe, services)

        Dim dataAplicarPorcExt As New Comunes.datosAplicarMargen(Nz(data.Current("CosteExtStdA"), 0), Nz(data.Current("PorcentajeExt"), 0))
        ProcessServer.ExecuteTask(Of Comunes.datosAplicarMargen)(AddressOf Comunes.AplicarMargen, dataAplicarPorcExt, services)

        Dim dataAplicarPorcVar As New Comunes.datosAplicarMargen(Nz(data.Current("CosteVarStdA"), 0), Nz(data.Current("PorcentajeVar"), 0))
        ProcessServer.ExecuteTask(Of Comunes.datosAplicarMargen)(AddressOf Comunes.AplicarMargen, dataAplicarPorcVar, services)

        Dim TotalVenta As Double = dataAplicarPorcMat.Venta + dataAplicarPorcOpe.Venta + dataAplicarPorcExt.Venta + dataAplicarPorcVar.Venta
        Select Case data.ColumnName
            Case "PVPA"
                Dim infoCalculoMargen As New Comunes.DatosCalculoMargen(Nz(data.Current("PVPA"), 0), TotalVenta)
                data.Current("PorcentajeBeneficio") = ProcessServer.ExecuteTask(Of Comunes.DatosCalculoMargen, Double)(AddressOf Comunes.CalcularMargen, infoCalculoMargen, services)
            Case Else
                Dim dataAplicarPorcBeneficio As New Comunes.datosAplicarMargen(TotalVenta, Nz(data.Current("PorcentajeBeneficio"), 0))
                ProcessServer.ExecuteTask(Of Comunes.datosAplicarMargen)(AddressOf Comunes.AplicarMargen, dataAplicarPorcBeneficio, services)

                data.Current("PVPA") = dataAplicarPorcBeneficio.Venta
        End Select

        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf CalcularImportesAyB, data.Current, services)
    End Sub

#End Region

#Region " RegisterValidateTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidaDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidaDatosObligatorios(ByVal dr As DataRow, ByVal services As ServiceProvider)
        If Length(dr("IDArticulo")) = 0 Then ApplicationService.GenerateError("El Artículo es un dato obligatorio.")
    End Sub

#End Region

#Region " RegisterUpdateTasks "

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarIdentificador)
        updateProcess.AddTask(Of DataRow)(AddressOf ActualizarImportesAyB)
    End Sub

    <Task()> Public Shared Sub ActualizarImportesAyB(ByVal dr As DataRow, ByVal services As ServiceProvider)
        Dim dataImportesAB As IPropertyAccessor = New DataRowPropertyAccessor(dr)
        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf CalcularImportesAyB, dataImportesAB, services)
    End Sub

    <Task()> Public Shared Sub CalcularImportesAyB(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider)
        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
        Dim MonedaA As MonedaInfo = Monedas.MonedaA
        Dim MonedaB As MonedaInfo = Monedas.MonedaB

        Dim ValAyB As New ValoresAyB(data, MonedaA.ID, MonedaA.CambioA, MonedaA.CambioB)
        ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)

        If data.ContainsKey("PVPA") AndAlso data("PVPA") <> 0 Then
            data("PVPB") = xRound(data("PVPA") * MonedaA.CambioB, MonedaB.NDecimalesImporte)
            data("PVPA") = xRound(data("PVPA"), MonedaA.NDecimalesImporte)
        End If

        If data.ContainsKey("PorcentajeMat") Then data("PorcentajeMat") = xRound(data("PorcentajeMat"), 2)
        If data.ContainsKey("PorcentajeOpe") Then data("PorcentajeOpe") = xRound(data("PorcentajeOpe"), 2)
        If data.ContainsKey("PorcentajeExt") Then data("PorcentajeExt") = xRound(data("PorcentajeExt"), 2)
        If data.ContainsKey("PorcentajeVar") Then data("PorcentajeVar") = xRound(data("PorcentajeVar"), 2)
        If data.ContainsKey("PorcentajeBeneficio") Then data("PorcentajeBeneficio") = xRound(data("PorcentajeBeneficio"), 2)

        If data.ContainsKey("CosteOperacionA") AndAlso data("CosteOperacionA") <> 0 Then
            data("CosteOperacionB") = xRound(data("CosteOperacionA") * MonedaA.CambioB, MonedaB.NDecimalesPrecio)
            data("CosteOperacionA") = xRound(data("CosteOperacionA"), MonedaA.NDecimalesPrecio)
        End If
        If data.ContainsKey("TasaEjecucionA") AndAlso Nz(data("TasaEjecucionA"), 0) <> 0 Then
            data("TasaEjecucionB") = xRound(data("TasaEjecucionA") * MonedaA.CambioB, MonedaB.NDecimalesPrecio)
            data("TasaEjecucionA") = xRound(data("TasaEjecucionA"), MonedaA.NDecimalesPrecio)
        End If
        If data.ContainsKey("TasaPreparacionA") AndAlso Nz(data("TasaPreparacionA"), 0) <> 0 Then
            data("TasaEjecucionB") = xRound(data("TasaPreparacionA") * MonedaA.CambioB, MonedaB.NDecimalesPrecio)
            data("TasaPreparacionA") = xRound(data("TasaPreparacionA"), MonedaA.NDecimalesPrecio)
        End If
        If data.ContainsKey("TasaMODA") AndAlso Nz(data("TasaMODA"), 0) <> 0 Then
            data("TasaMODB") = xRound(data("TasaMODA") * MonedaA.CambioB, MonedaB.NDecimalesPrecio)
            data("TasaMODA") = xRound(data("TasaMODA"), MonedaA.NDecimalesPrecio)
        End If
    End Sub

#End Region

#Region " NuevoPresupuesto "

#Region " Versión 'Taskeada' "

    <Serializable()> _
    Public Class dataNuevoPresupuesto
        Public IDFuente As String
        Public TipoFuente As enumpceFuente
        Public IDDestino As String
        Public IDRuta As String
        Public IDEstructura As String
        Public TipoCoste As enumpceTipoCoste
        Public IDContador As String
        Public IDCliente As String
        Public IDEmpresa As String

        Public Sub New(ByVal TipoFuente As enumpceFuente, ByVal IDFuente As String, ByVal IDDestino As String, ByVal IDRuta As String, _
                       ByVal IDEstructura As String, ByVal TipoCoste As enumpceTipoCoste, ByVal IDContador As String)

            Me.TipoFuente = TipoFuente
            Me.IDFuente = IDFuente
            Me.IDDestino = IDDestino
            Me.IDRuta = IDRuta
            Me.IDEstructura = IDEstructura
            Me.TipoCoste = TipoCoste
            Me.IDContador = IDContador
        End Sub
    End Class
    <Task()> Public Shared Sub GenerarNuevoPresupuesto(ByVal data As dataNuevoPresupuesto, ByVal services As ServiceProvider)
        If Length(data.IDFuente) > 0 Then
            If Length(data.IDDestino) > 0 Then
                ProcessServer.ExecuteTask(Of dataNuevoPresupuesto)(AddressOf GenerarPresupuestoCosteEstandar, data, services)
            Else
                ApplicationService.GenerateError("El Artículo de Destino es obligatorio.")
            End If
        Else
            ProcessServer.ExecuteTask(Of dataNuevoPresupuesto)(AddressOf GenerarPresupuestoCosteEstandarSinOrigen, data, services)
        End If
    End Sub

#Region " GenerarPresupuestoCosteEstandar "

    <Task()> Public Shared Sub GenerarPresupuestoCosteEstandar(ByVal data As dataNuevoPresupuesto, ByVal services As ServiceProvider)
        Dim dtPresup As DataTable = ProcessServer.ExecuteTask(Of dataNuevoPresupuesto, DataTable)(AddressOf GetOrigenPresupuesto, data, services)
        If Not dtPresup Is Nothing AndAlso dtPresup.Rows.Count > 0 Then
            Dim NPresupuesto As String = String.Empty
            If data.TipoCoste = enumpceTipoCoste.pcePresupuesto Then
                ProcessServer.ExecuteTask(Of String, String)(AddressOf GetNPresupuesto, data.IDContador, services)
            End If
            Select Case data.TipoFuente
                Case enumpceFuente.pceArticulo
                Case enumpceFuente.pcePreSim
            End Select
        End If
    End Sub

    <Task()> Public Shared Function GetOrigenPresupuesto(ByVal data As dataNuevoPresupuesto, ByVal services As ServiceProvider) As DataTable
        Dim dt As DataTable = Nothing
        Select Case data.TipoFuente
            Case enumpceFuente.pceArticulo
                Dim StDataCoste As New ArticuloCosteEstandar.DataCosteEstandarPresupuesto(data.IDFuente, data.IDRuta, data.IDEstructura)
                Dim dsCoste As DataSet = ProcessServer.ExecuteTask(Of ArticuloCosteEstandar.DataCosteEstandarPresupuesto, DataSet)(AddressOf ArticuloCosteEstandar.CosteEstandarPresupuesto, StDataCoste, services)
                If Not dsCoste Is Nothing Then
                    ' If Not dsPresupuesto.Tables("ArticuloCosteEstandar") Is Nothing Then
                    dt = dsCoste.Tables("ArticuloCosteEstandar")
                    ' End If
                End If
            Case enumpceFuente.pcePreSim
                dt = New ArticuloCosteEstandar().SelOnPrimaryKey(data.IDFuente)
        End Select

        Return dt
    End Function

#End Region

    <Task()> Public Shared Sub GenerarPresupuestoCosteEstandarSinOrigen(ByVal data As dataNuevoPresupuesto, ByVal services As ServiceProvider)
        Dim NPresupuesto As String = String.Empty
        If data.TipoCoste = enumpceTipoCoste.pcePresupuesto Then
            ProcessServer.ExecuteTask(Of String, String)(AddressOf GetNPresupuesto, data.IDContador, services)
        End If

        Dim PCE As New PresupuestoCosteEstandar
        Dim dtPCE As DataTable = PCE.AddNewForm
        Select Case data.TipoFuente
            Case enumpceFuente.pceArticulo
                dtPCE.Rows(0)("IDArticulo") = data.IDDestino
                dtPCE.Rows(0)("DescArticulo") = Left(data.IDDestino & " (" & data.IDRuta & " -- " & data.IDEstructura & ")", 300)
                dtPCE.Rows(0)("Fecha") = Date.Today
                dtPCE.Rows(0)("CosteStdA") = 0
                dtPCE.Rows(0)("PVPA") = 0
                dtPCE.Rows(0)("Estado") = enumpceEstado.pcePresupuestado
                If Length(NPresupuesto) > 0 Then dtPCE.Rows(0)("NPresupuesto") = NPresupuesto
                If Length(data.IDContador) > 0 Then dtPCE.Rows(0)("IDContador") = data.IDContador
            Case enumpceFuente.pcePreSim
                dtPCE.Rows(0)("IDArticulo") = data.IDDestino
                If data.TipoCoste = enumpceTipoCoste.pcePresupuesto Then
                    dtPCE.Rows(0)("IDContador") = data.IDContador
                    dtPCE.Rows(0)("NPresupuesto") = NPresupuesto
                End If
        End Select
        PCE.Update(dtPCE)
    End Sub

    <Task()> Public Shared Function GetNPresupuesto(ByVal IDContador As String, ByVal services As ServiceProvider) As String
        If Length(IDContador) > 0 Then
            Dim NPresupuesto As String = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, IDContador, services)
            If Length(NPresupuesto) = 0 Then
                ApplicationService.GenerateError("El contador no existe o no está correctamente configurado.")
            End If
            Return NPresupuesto
        Else
            ApplicationService.GenerateError("No se ha configurado contador para Presupuestos.")
        End If
    End Function





#End Region

    Public Function NuevoPresupuesto(ByVal lngTipoFuente As enumpceFuente, ByVal strIDFuente As String, ByVal strIDDestino As String, _
                                     ByVal strIDRuta As String, ByVal strIDEstructura As String, ByVal lngTipoCoste As enumpceTipoCoste, _
                                     ByVal strIDContador As String, Optional ByVal IDCliente As String = Nothing, _
                                     Optional ByVal IDEmpresa As String = Nothing) As DataTable
        Dim lngIDPresupuesto As Integer
        Dim dsPresupuesto As DataSet
        Dim dtSource As DataTable
        Dim dtAux As DataTable
        Dim FwnCosteStd As ArticuloCosteEstandar

        If Length(strIDFuente) > 0 Then
            If Length(strIDDestino) > 0 Then
                Select Case lngTipoFuente
                    'Establecer la fuente del presupuesto.Calcular si es necesario
                    Case enumpceFuente.pceArticulo
                        Dim services As New ServiceProvider
                        Dim StDataCoste As New ArticuloCosteEstandar.DataCosteEstandarPresupuesto(strIDFuente, strIDRuta, strIDEstructura)
                        FwnCosteStd = New ArticuloCosteEstandar
                        dsPresupuesto = ProcessServer.ExecuteTask(Of ArticuloCosteEstandar.DataCosteEstandarPresupuesto, DataSet)(AddressOf ArticuloCosteEstandar.CosteEstandarPresupuesto, StDataCoste, services)

                        If Not dsPresupuesto Is Nothing Then
                            If Not dsPresupuesto.Tables("ArticuloCosteEstandar") Is Nothing Then
                                dtSource = dsPresupuesto.Tables("ArticuloCosteEstandar")
                            End If
                        End If
                    Case enumpceFuente.pcePreSim
                        dtSource = SelOnPrimaryKey(CInt(strIDFuente))
                End Select

                'Completar datos
                If Not dtSource Is Nothing AndAlso dtSource.Rows.Count > 0 Then

                    dtAux = LlenarPresupuestoCosteEstandar(dtSource, lngTipoFuente, strIDFuente, strIDDestino, strIDRuta, strIDEstructura, lngTipoCoste, strIDContador, IDCliente, IDEmpresa)
                    If Not dtAux Is Nothing AndAlso dtAux.Rows.Count > 0 Then
                        If IsNumeric(dtAux.Rows(0)("IDPresupuesto")) Then
                            lngIDPresupuesto = dtAux.Rows(0)("IDPresupuesto")
                        End If
                        BusinessHelper.UpdateTable(dtAux)
                        Select Case lngTipoFuente
                            Case enumpceFuente.pceArticulo
                                BusinessHelper.UpdateTable(LlenarPresupuestoCosteMaterial(dsPresupuesto.Tables("HistoricoCosteMaterial"), lngIDPresupuesto))
                                BusinessHelper.UpdateTable(LlenarPresupuestoCosteOperacion(dsPresupuesto.Tables("HistoricoCosteOperacion"), lngIDPresupuesto))
                                BusinessHelper.UpdateTable(LlenarPresupuestoCosteVarios(dsPresupuesto.Tables("HistoricoCosteVarios"), lngIDPresupuesto))
                            Case enumpceFuente.pcePreSim
                                BusinessHelper.UpdateTable(LlenarPresupuesto(CInt(strIDFuente), "PresupuestoCosteMaterial", lngIDPresupuesto))
                                BusinessHelper.UpdateTable(LlenarPresupuesto(CInt(strIDFuente), "PresupuestoCosteOperacion", lngIDPresupuesto))
                                BusinessHelper.UpdateTable(LlenarPresupuesto(CInt(strIDFuente), "PresupuestoCosteVarios", lngIDPresupuesto))
                        End Select

                    End If
                End If
            Else
                ApplicationService.GenerateError("El Artículo de Destino es obligatorio")
            End If
        Else
            dtAux = CrearPresupuestoCosteEstandarSinOrigen(lngTipoFuente, strIDDestino, strIDRuta, strIDEstructura, lngTipoCoste, strIDContador)
            If Not dtAux Is Nothing AndAlso dtAux.Rows.Count > 0 Then
                BusinessHelper.UpdateTable(dtAux)
            End If
        End If
        Return (dtAux)
    End Function

    Public Function LlenarPresupuestoCosteEstandar(ByVal dtData As DataTable, ByVal lngTipoFuente As enumpceFuente, ByVal strIDFuente As String, ByVal strIDDestino As String, _
                                                        ByVal strIDRuta As String, ByVal strIDEstructura As String, ByVal lngTipoCoste As enumpceTipoCoste, _
                                                        ByVal strIDContador As String, Optional ByVal IDCliente As String = Nothing, _
                                                       Optional ByVal IDEmpresa As String = Nothing) As DataTable
        Dim services As New ServiceProvider
        Dim blnCancel As Boolean
        Dim lngIDPresupuesto As Integer
        Dim strNPresupuesto As String
        Dim dtNew As DataTable


        If Not dtData Is Nothing AndAlso dtData.Rows.Count > 0 Then

            If lngTipoCoste = enumpceTipoCoste.pcePresupuesto Then
                If Length(strIDContador) > 0 Then
                    Dim clsAdmin As AdminData
                    strNPresupuesto = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, strIDContador, services)
                    If Length(strNPresupuesto) = 0 Then
                        blnCancel = True
                        ApplicationService.GenerateError("El contador | no existe o no está correctamente configurado")
                    End If
                Else
                    blnCancel = True
                    ApplicationService.GenerateError("El contador | no existe o no está correctamente configurado")
                End If
            End If

            If Not blnCancel Then
                dtNew = AddNewForm()
                If Not dtNew Is Nothing Then
                    Select Case lngTipoFuente
                        Case enumpceFuente.pceArticulo
                            With dtNew
                                .Rows(0)("IDArticulo") = strIDFuente
                                .Rows(0)("IDArticuloPresupuesto") = strIDDestino
                                .Rows(0)("DescArticulo") = Left(strIDDestino & " (" & strIDRuta & " -- " & strIDEstructura & ")", 300)
                                .Rows(0)("Fecha") = Today
                                .Rows(0)("CosteMatStdA") = dtData.Rows(0)("CosteAcuMatUltA")
                                .Rows(0)("CosteOpeStdA") = dtData.Rows(0)("CosteAcuOpeUltA")
                                .Rows(0)("CosteExtStdA") = dtData.Rows(0)("CosteAcuExtUltA")
                                .Rows(0)("CosteVarStdA") = dtData.Rows(0)("CosteAcuVarUltA")
                                .Rows(0)("CosteStdA") = .Rows(0)("CosteMatStdA") + .Rows(0)("CosteOpeStdA") + .Rows(0)("CosteExtStdA") + .Rows(0)("CosteVarStdA")
                                .Rows(0)("PVPA") = .Rows(0)("CosteMatStdA") + .Rows(0)("CosteOpeStdA") + .Rows(0)("CosteExtStdA") + .Rows(0)("CosteVarStdA")
                                .Rows(0)("PVPB") = 0
                                .Rows(0)("estado") = enumpceEstado.pcePresupuestado
                                If Not IDCliente Is Nothing Then
                                    .Rows(0)("IDCliente") = IDCliente
                                    .Rows(0)("DescCliente") = New Cliente().GetItemRow(IDCliente)("DescCliente")
                                    'Datos Dirección
                                    Dim FilDirec As New Filter
                                    FilDirec.Add("IDCliente", FilterOperator.Equal, IDCliente)
                                    FilDirec.Add("Predeterminada", FilterOperator.Equal, 1)
                                    Dim DtDirec As DataTable = New ClienteDireccion().Filter(FilDirec)
                                    If Not DtDirec Is Nothing AndAlso DtDirec.Rows.Count > 0 Then
                                        .Rows(0)("Direccion") = DtDirec.Rows(0)("Direccion")
                                        .Rows(0)("CodPostal") = DtDirec.Rows(0)("CodPostal") & String.Empty
                                        .Rows(0)("Poblacion") = DtDirec.Rows(0)("Poblacion") & String.Empty
                                        .Rows(0)("Provincia") = DtDirec.Rows(0)("Provincia") & String.Empty
                                        .Rows(0)("IDPais") = DtDirec.Rows(0)("IDPais") & String.Empty
                                        .Rows(0)("Telefono") = DtDirec.Rows(0)("Telefono") & String.Empty
                                        .Rows(0)("Fax") = DtDirec.Rows(0)("Fax") & String.Empty
                                        .Rows(0)("EMail") = DtDirec.Rows(0)("Email") & String.Empty
                                    End If
                                    'Datos Contacto
                                    Dim FilContac As New Filter
                                    FilContac.Add("IDCliente", FilterOperator.Equal, IDCliente)
                                    FilContac.Add("Predeterminada", FilterOperator.Equal, 1)
                                    Dim DtContac As DataTable = New ClientePersonaContacto().Filter(FilContac)
                                    If Not DtContac Is Nothing AndAlso DtContac.Rows.Count > 0 Then
                                        .Rows(0)("Interlocutor") = DtContac.Rows(0)("IDPersona")
                                    End If
                                End If
                                If Not IDEmpresa Is Nothing Then
                                    .Rows(0)("IDEmpresa") = IDEmpresa
                                    'Datos Dirección
                                    Dim Control As BE.BusinessHelper = CreateBusinessObject("Empresa")
                                    Dim FilDirec As New Filter
                                    FilDirec.Add("IDEmpresa", FilterOperator.Equal, IDEmpresa)
                                    Dim DtDirec As DataTable = Control.Filter(FilDirec)
                                    .Rows(0)("DescCliente") = Control.GetItemRow(IDEmpresa)("DescEmpresa")
                                    If Not DtDirec Is Nothing AndAlso DtDirec.Rows.Count > 0 Then
                                        .Rows(0)("Direccion") = DtDirec.Rows(0)("Direccion")
                                        .Rows(0)("CodPostal") = DtDirec.Rows(0)("CodPostal") & String.Empty
                                        .Rows(0)("Poblacion") = DtDirec.Rows(0)("Poblacion") & String.Empty
                                        .Rows(0)("Provincia") = DtDirec.Rows(0)("Provincia") & String.Empty
                                        .Rows(0)("IDPais") = DtDirec.Rows(0)("IDPais") & String.Empty
                                        .Rows(0)("Telefono") = DtDirec.Rows(0)("Telefono1") & String.Empty
                                        .Rows(0)("Fax") = DtDirec.Rows(0)("Fax") & String.Empty
                                        .Rows(0)("EMail") = DtDirec.Rows(0)("Email") & String.Empty
                                    End If
                                    'Datos Contacto
                                    Dim Filcontacto As New Filter
                                    Filcontacto.Add("IDEmpresa", FilterOperator.Equal, IDEmpresa)
                                    Filcontacto.Add("Predeterminada", FilterOperator.Equal, 1)
                                    Control = CreateBusinessObject("EmpresaPersona")
                                    Dim DtContac As DataTable = Control.Filter(Filcontacto)
                                    If Not DtContac Is Nothing AndAlso DtContac.Rows.Count > 0 Then
                                        .Rows(0)("Interlocutor") = DtContac.Rows(0)("IDPersona")
                                    End If
                                End If
                                If Length(strNPresupuesto) > 0 Then .Rows(0)("NPresupuesto") = strNPresupuesto
                                If Length(strIDContador) > 0 Then .Rows(0)("IDContador") = strIDContador
                            End With

                        Case enumpceFuente.pcePreSim
                            lngIDPresupuesto = dtNew.Rows(0)("IDPresupuesto")
                            For Each col As DataColumn In dtNew.Columns
                                If col.ColumnName <> "IDPresupuesto" And col.ColumnName <> "IDContador" And col.ColumnName <> "Fecha" And col.ColumnName <> "NPresupuesto" And col.ColumnName <> "IDArticuloPresupuesto" And InStr(1, col.ColumnName, "Audi", CompareMethod.Text) = 0 Then
                                    dtNew.Rows(0)(col.ColumnName) = dtData.Rows(0)(col.ColumnName)
                                End If
                            Next

                            dtNew.Rows(0)("IDArticuloPresupuesto") = strIDDestino
                            dtNew.Rows(0)("Fecha") = Today
                            If lngTipoCoste = enumpceTipoCoste.pcePresupuesto Then
                                dtNew.Rows(0)("IDContador") = strIDContador
                                dtNew.Rows(0)("NPresupuesto") = strNPresupuesto
                            End If

                    End Select
                End If
            End If
        End If
        Return (dtNew)
    End Function

    Public Function CrearPresupuestoCosteEstandarSinOrigen(ByVal lngTipoFuente As enumpceFuente, ByVal strIDDestino As String, ByVal strIDRuta As String, ByVal strIDEstructura As String, ByVal lngTipoCoste As enumpceTipoCoste, ByVal strIDContador As String) As DataTable
        Dim services As New ServiceProvider
        Dim blnCancel As Boolean
        Dim lngIDPresupuesto As Integer
        Dim strNPresupuesto As String
        Dim dtNew As DataTable

        If lngTipoCoste = enumpceTipoCoste.pcePresupuesto Then
            If Length(strIDContador) > 0 Then
                strNPresupuesto = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, strIDContador, services)
                If Length(strNPresupuesto) = 0 Then
                    blnCancel = True
                    ApplicationService.GenerateError("El contador | no existe o no está correctamente configurado")
                End If
            Else
                blnCancel = True
                ApplicationService.GenerateError("El contador | no existe o no está correctamente configurado")
            End If
        End If

        If Not blnCancel Then
            dtNew = AddNewForm()
            If Not dtNew Is Nothing Then
                Select Case lngTipoFuente
                    Case enumpceFuente.pceArticulo
                        With dtNew
                            .Rows(0)("IDArticulo") = strIDDestino
                            .Rows(0)("DescArticulo") = Left(strIDDestino & " (" & strIDRuta & " -- " & strIDEstructura & ")", 300)
                            .Rows(0)("Fecha") = Today
                            .Rows(0)("CosteStdA") = Nz(.Rows(0)("CosteMatStdA"), 0) + Nz(.Rows(0)("CosteOpeStdA"), 0) + Nz(.Rows(0)("CosteExtStdA"), 0) + Nz(.Rows(0)("CosteVarStdA"), 0)
                            .Rows(0)("PVPA") = 0
                            .Rows(0)("estado") = enumpceEstado.pcePresupuestado
                            If Len(strNPresupuesto) > 0 Then .Rows(0)("NPresupuesto") = strNPresupuesto
                            If Len(strIDContador) > 0 Then .Rows(0)("IDContador") = strIDContador
                        End With

                    Case enumpceFuente.pcePreSim
                        lngIDPresupuesto = dtNew.Rows(0)("IDPresupuesto")

                        dtNew.Rows(0)("IDArticulo") = strIDDestino
                        If lngTipoCoste = enumpceTipoCoste.pcePresupuesto Then
                            dtNew.Rows(0)("IDContador") = strIDContador
                            dtNew.Rows(0)("NPresupuesto") = strNPresupuesto
                        End If
                End Select
            End If
        End If

        Return dtNew
    End Function

    Public Function LlenarPresupuestoCosteMaterial(ByVal dtData As DataTable, ByVal lngIDPresupuesto As Integer) As DataTable
        Dim dtArticulo As DataTable
        Dim dtNew As DataTable
        Dim FwnArticulo As Articulo
        Dim FwnPCMat As PresupuestoCosteMaterial

        If Not dtData Is Nothing AndAlso dtData.Rows.Count > 0 Then
            FwnPCMat = New PresupuestoCosteMaterial
            dtNew = FwnPCMat.AddNew()
            If Not dtNew Is Nothing Then
                FwnArticulo = New Articulo
                For Each dr As DataRow In dtData.Rows
                    If Not AreEquals(dr("IDArticulo"), dr("IDComponente")) Then
                        If AreEquals(CInt(dr("Tipo")), enumacsTipoArticulo.acsCompra) Then
                            With dtNew
                                Dim rw As DataRow = .NewRow
                                rw("IDPresupMaterial") = AdminData.GetAutoNumeric
                                rw("IDPresupuesto") = lngIDPresupuesto
                                rw("IDArticulo") = dr("IDArticulo")
                                rw("IDComponente") = dr("IDComponente")
                                dtArticulo = FwnArticulo.Filter("DescArticulo", "IDArticulo='" & dr("IDComponente") & "'")
                                If Not dtArticulo Is Nothing AndAlso dtArticulo.Rows.Count > 0 Then
                                    rw("DescComponente") = dtArticulo.Rows(0)("DescArticulo")
                                End If
                                rw("Cantidad") = dr("CantidadAcumulada")
                                rw("Merma") = dr("Merma")
                                rw("PrecioStdA") = dr("CosteMatStdA")
                                rw("PorcentajeMat") = 0
                                rw("CosteStdA") = rw("Cantidad") * (1 + (rw("Merma") / 100)) * rw("PrecioStdA")
                                .Rows.Add(rw)
                            End With
                        End If
                    End If
                Next
            End If
        End If
        Return (dtNew)
    End Function

    Public Function LlenarPresupuestoCosteOperacion(ByVal dtData As DataTable, ByVal lngIDPresupuesto As Integer) As DataTable
        Dim dtNew As DataTable
        Dim FwnPCOpe As PresupuestoCosteOperacion

        If Not dtData Is Nothing AndAlso dtData.Rows.Count > 0 Then
            FwnPCOpe = New PresupuestoCosteOperacion
            dtNew = FwnPCOpe.AddNew()
            If Not dtNew Is Nothing Then
                For Each dr As DataRow In dtData.Rows
                    With dtNew
                        Dim rw As DataRow = .NewRow
                        rw("IDPresupOperacion") = AdminData.GetAutoNumeric
                        rw("IDPresupuesto") = lngIDPresupuesto
                        rw("IDArticulo") = dr("IDArticulo")
                        rw("Secuencia") = dr("Secuencia")
                        rw("TipoOperacion") = dr("TipoOperacion")
                        rw("IDOperacion") = dr("IDOperacion")
                        rw("DescOperacion") = dr("DescOperacion")
                        rw("IDCentro") = dr("IDCentro")
                        rw("FactorHombre") = dr("FactorHombre")
                        rw("TiempoPrep") = dr("TiempoPrep")
                        rw("UdTiempoPrep") = dr("UdTiempoPrep")
                        rw("TiempoEjecUnit") = dr("TiempoEjecUnit")
                        rw("UdTiempoEjec") = dr("UdTiempoEjec")
                        rw("FactorProduccion") = dr("FactorProduccion")
                        rw("TasaEjecucionA") = dr("TasaEjecucionA")
                        rw("TasaPreparacionA") = dr("TasaPreparacionA")
                        rw("TasaMODA") = dr("TasaMODA")
                        rw("TasaEjecucionB") = dr("TasaEjecucionB")
                        rw("TasaPreparacionB") = dr("TasaPreparacionB")
                        rw("TasaMODB") = dr("TasaMODB")
                        rw("Nivel") = dr("Nivel")
                        rw("Orden") = dr("Orden")
                        rw("LoteMinimo") = dr("LoteMinimo")
                        rw("CosteOperacionA") = dr("CosteOperacionA")
                        rw("CosteOperacionB") = dr("CosteOperacionB")
                        rw("IDProveedor") = dr("IDProveedor")
                        rw("CantidadAcumulada") = dr("CantidadAcumulada")
                        .Rows.Add(rw)
                    End With
                Next
            End If
        End If
        Return (dtNew)
    End Function

    Public Function LlenarPresupuestoCosteVarios(ByVal dtData As DataTable, ByVal lngIDPresupuesto As Integer) As DataTable
        Dim dtNew As DataTable
        Dim FwnPCVar As PresupuestoCosteVarios

        If Not dtData Is Nothing AndAlso dtData.Rows.Count > 0 Then
            FwnPCVar = New PresupuestoCosteVarios
            dtNew = FwnPCVar.AddNew()
            If Not dtNew Is Nothing Then
                For Each dr As DataRow In dtData.Rows
                    With dtNew
                        Dim rw As DataRow = .NewRow
                        rw("IDPresupVarios") = AdminData.GetAutoNumeric
                        rw("IDPresupuesto") = lngIDPresupuesto
                        rw("IDArticulo") = dr("IDArticulo")
                        rw("IDVarios") = dr("IDVarios")
                        rw("DescVarios") = dr("DescVarios")
                        rw("Nivel") = dr("Nivel")
                        rw("Orden") = dr("Orden")
                        rw("Valor") = dr("Valor")
                        rw("Tipo") = dr("Tipo")
                        rw("CosteVariosA") = dr("CosteVariosA")
                        rw("CosteVariosB") = dr("CosteVariosB")
                        .Rows.Add(rw)
                    End With
                Next
            End If
        End If
        Return (dtNew)
    End Function

    Public Function LlenarPresupuesto(ByVal lngIDFuente As Integer, ByVal strEntidad As String, ByVal lngIDPresupuesto As Integer) As DataTable
        Dim strPK As String
        Dim dtPk As DataTable
        Dim dtPresupuesto As DataTable
        Dim dtAux As DataTable
        Dim fwn = BusinessHelper.CreateBusinessObject(AdminData.GetEntityInfo(strEntidad))

        Dim ofilter As New Filter
        ofilter.Add(New NoRowsFilterItem)
        dtPk = fwn.filter(ofilter)
        Dim dtPk1 As DataTable = fwn.PrimaryKeyTable()

        If Not dtPk Is Nothing Then
            strPK = dtPk.Columns(0).ColumnName
        End If
        If Length(strPK) > 0 Then dtAux = fwn.Filter(, "IDPresupuesto=" & lngIDFuente)

        If Not dtAux Is Nothing Then
            dtPresupuesto = dtPk.Clone
            If Not dtPresupuesto Is Nothing Then
                For Each dr As DataRow In dtAux.Rows
                    Dim rw As DataRow = dtPresupuesto.NewRow
                    rw(strPK) = AdminData.GetAutoNumeric
                    For Each col As DataColumn In dtPresupuesto.Columns
                        If col.ColumnName <> strPK And InStr(1, col.ColumnName, "Audi", CompareMethod.Text) = 0 Then
                            If col.ColumnName = "IDPresupuesto" Then
                                rw("IDPresupuesto") = lngIDPresupuesto
                            Else
                                rw(col.ColumnName) = dr(col.ColumnName)
                            End If
                        End If
                    Next
                    dtPresupuesto.Rows.Add(rw)
                Next
            End If
        End If
        Return (dtPresupuesto)
    End Function

#End Region

End Class