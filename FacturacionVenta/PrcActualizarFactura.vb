'Proceso de Creación de la factura establecido de forma estándar, relación de tareas
Imports Solmicro.Expertis.Engine.BE.BusinessProcesses
Public Class PrcActualizarFactura
    Inherits Process(Of DataPrcActualizarFactura, ResultFacturacion)

    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of DataPrcActualizarFactura, ArrayList)(AddressOf PrepararInformacionProceso)
        Me.AddForEachTask(Of DocumentoFacturaVenta)(AddressOf ActualizarDocumentoFactura, OnExceptionBehaviour.NextLoop)
        Me.AddTask(Of Object, ResultFacturacion)(AddressOf Resultado)
    End Sub

    <Task()> Public Shared Function PrepararInformacionProceso(ByVal data As DataPrcActualizarFactura, ByVal services As ServiceProvider) As ArrayList
        Dim arDocsFra As ArrayList = AdminData.GetSessionData("__frax__")
        If Not arDocsFra Is Nothing Then
            If data.RstFacturacion.Log Is Nothing Then data.RstFacturacion.Log = New LogProcess
            services.RegisterService(data, GetType(DataPrcActualizarFactura))
        End If
        Return arDocsFra
    End Function

    <Task()> Public Shared Sub ActualizarDocumentoFactura(ByVal doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If doc.dtLineas.Rows.Count = 0 Then Exit Sub

        AdminData.BeginTx()
        Dim PrcInfo As ProcessInfoFra = services.GetService(Of ProcessInfoFra)()
        Dim TipoFact As enumTipoFactura = enumTipoFactura.tfNormal
        If PrcInfo.ConPropuesta Then
            Dim FactVenta As New ResultFacturacion
            Dim DataFact As DataPrcActualizarFactura = services.GetService(Of DataPrcActualizarFactura)()
            FactVenta = DataFact.RstFacturacion
            TipoFact = DataFact.TipoFactura

            Dim dv As New DataView(FactVenta.PropuestaFacturas, Nothing, "IDFactura", DataViewRowState.CurrentRows)
            Dim idx As Integer = dv.Find(doc.HeaderRow("IDFactura"))
            If idx >= 0 Then
                doc.HeaderRow("IDContador") = dv(idx)("IDcontador")
                doc.HeaderRow("FechaFactura") = dv(idx)("FechaFactura")
                doc.HeaderRow("IDEjercicio") = dv(idx)("IDEjercicio")
                doc.HeaderRow("FechaParaDeclaracion") = dv(idx)("FechaParaDeclaracion")
                ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ProcesoFacturacionVenta.FechaParaDeclaracionComoProveedor, New DataRowPropertyAccessor(doc.HeaderRow), services)
            End If
        End If

        ProcessServer.ExecuteTask(Of DataRow)(AddressOf NegocioGeneral.ValidarFechaDeclaracion, doc.HeaderRow, services)
        Dim AppParams As ParametroVenta = services.GetService(Of ParametroVenta)()

        Dim TPVFactura As Boolean = AppParams.TPVFactura
        If TypeOf doc.Cabecera Is FraCabAlbaran AndAlso Length(CType(doc.Cabecera, FraCabAlbaran).IDTPV) > 0 AndAlso (TPVFactura OrElse CType(doc.Cabecera, FraCabAlbaran).AgrupFactura = enummcAgrupFactura.mcAlbaran) Then
            doc.HeaderRow("NFactura") = CType(doc.Cabecera, FraCabAlbaran).NAlbaran
        Else
            Dim StDatos As New Contador.DatosCounterValue
            StDatos.IDCounter = doc.HeaderRow("IDContador")
            StDatos.TargetClass = New FacturaVentaCabecera
            StDatos.TargetField = "NFactura"
            StDatos.DateField = "FechaFactura"
            StDatos.DateValue = doc.HeaderRow("FechaFactura")
            StDatos.IDEjercicio = doc.HeaderRow("IDEjercicio") & String.Empty
            doc.HeaderRow("NFactura") = ProcessServer.ExecuteTask(Of Contador.DatosCounterValue, String)(AddressOf Contador.CounterValue, StDatos, services)
        End If

        'TODO Orden de ejecución de la actualización
        If doc.HeaderRow.Table.Columns.Contains("IDMandato") Then
            If Length(doc.HeaderRow("IDMandato")) = 0 Then ProcessServer.ExecuteTask(Of DataRow)(AddressOf GetMandatoPredeterminado, doc.HeaderRow, services)
        End If

		ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf FacturaVentaCabecera.AsignarMotivoNoAseguradoIProp, New DataRowPropertyAccessor(doc.HeaderRow), services)

        ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf NegocioGeneral.AsignarEnvio347Doc, doc, services)
        ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf NegocioGeneral.AsignarEnvio349Doc, doc, services)
        ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularBasesImponibles, doc, services)
        ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularTotales, doc, services)
        ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularPuntoVerde, doc, services)
        ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularVencimientos, doc, services)
        ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarClaveOperacion, doc, services)
        'ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoComunes.ValidarIVASDocFV, doc, services)
        ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.GrabarDocumento, doc, services)
        '///Añadimos los cobros de compensación
        ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.NuevosCobrosCompensacionOSFacturasCerradasConFianza, doc, services)

        Select Case TipoFact
            Case enumTipoFactura.tfNormal
                ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.ActualizarAlbaranEnProceso, doc, services)
                For Each Linea As DataRow In doc.dtLineas.Rows
                    Dim datosObra As New ProcesoFacturacionObras.DataActualizarRowConceptosObras(Linea)
                    ProcessServer.ExecuteTask(Of ProcesoFacturacionObras.DataActualizarRowConceptosObras)(AddressOf ProcesoFacturacionObras.ActualizarObraMateriales, datosObra, services)
                    ProcessServer.ExecuteTask(Of ProcesoFacturacionObras.DataActualizarRowConceptosObras)(AddressOf ProcesoFacturacionObras.ActualizarObraTrabajo, datosObra, services)
                    ProcessServer.ExecuteTask(Of ProcesoFacturacionObras.DataActualizarRowConceptosObras)(AddressOf ProcesoFacturacionObras.ActualizarObraCabecera, datosObra, services)
                    Dim datosCtrlOT As New ProcesoFacturacionVenta.DataActualizarRowControlOT(Linea)
                    ProcessServer.ExecuteTask(Of ProcesoFacturacionVenta.DataActualizarRowControlOT)(AddressOf ProcesoFacturacionVenta.ActualizarControlOT, datosCtrlOT, services)
                Next
            Case enumTipoFactura.tfObra
                Dim datosObra As New ProcesoFacturacionObras.DataActualizarConceptosObras(doc)
                ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionObras.ActualizarObraTrabajoFacturacion, doc, services)
                ProcessServer.ExecuteTask(Of ProcesoFacturacionObras.DataActualizarConceptosObras)(AddressOf ProcesoFacturacionObras.ActualizarConceptosObras, datosObra, services)
            Case enumTipoFactura.tfCertificacion
                For Each Linea As DataRow In doc.dtLineas.Rows
                    Dim datosObra As New ProcesoFacturacionObras.DataActualizarRowConceptosObras(Linea)
                    ProcessServer.ExecuteTask(Of ProcesoFacturacionObras.DataActualizarRowConceptosObras)(AddressOf ProcesoFacturacionObras.ActualizarObraTrabajo, datosObra, services)
                    ProcessServer.ExecuteTask(Of ProcesoFacturacionObras.DataActualizarRowConceptosObras)(AddressOf ProcesoFacturacionObras.ActualizarObraCabecera, datosObra, services)
                    ProcessServer.ExecuteTask(Of ProcesoFacturacionObras.DataActualizarRowConceptosObras)(AddressOf ProcesoFacturacionObras.ActualizarObraCertificaciones, datosObra, services)
                Next
            Case enumTipoFactura.tfPromocionObra
                ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionObras.ActualizaObraPromoLocalVencimiento, doc, services)
            Case enumTipoFactura.tfPromocionObraFinal
                ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionObras.ActualizaObraPromoLocalVencimientoFinal, doc, services)
        End Select
        AdminData.CommitTx(True)
        ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf AgregarFacturaResultado, doc, services)
    End Sub

    <Task()> Public Shared Sub GetMandatoPredeterminado(ByVal HeaderRow As DataRow, ByVal services As ServiceProvider)
        If Nz(HeaderRow("IDClienteBanco"), 0) <> 0 AndAlso HeaderRow.Table.Columns.Contains("IDMandato") Then
            Dim fMandatoPred As New Filter
            fMandatoPred.Add(New BooleanFilterItem("Predeterminado", True))
            fMandatoPred.Add(New NumberFilterItem("IDClienteBanco", HeaderRow("IDClienteBanco")))
            Dim dtMandato As DataTable = AdminData.GetData("tbMaestroMandato", fMandatoPred)
            If dtMandato.Rows.Count > 0 Then
                HeaderRow("IDMandato") = dtMandato.Rows(0)("IDMandato")
            End If
        End If
    End Sub
    <Task()> Public Shared Sub AgregarFacturaResultado(ByVal data As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        Dim PrcInfo As ProcessInfoFra = services.GetService(Of ProcessInfoFra)()
        Dim FactVenta As New ResultFacturacion
        If PrcInfo.ConPropuesta Then
            Dim DataFact As DataPrcActualizarFactura = services.GetService(Of DataPrcActualizarFactura)()
            FactVenta = DataFact.RstFacturacion
        Else : FactVenta = services.GetService(Of ResultFacturacion)()
        End If
        ReDim Preserve FactVenta.Log.CreatedElements(UBound(FactVenta.Log.CreatedElements) + 1)
        FactVenta.Log.CreatedElements(UBound(FactVenta.Log.CreatedElements)) = New CreateElement
        FactVenta.Log.CreatedElements(UBound(FactVenta.Log.CreatedElements)).IDElement = data.HeaderRow("IDFactura")
        FactVenta.Log.CreatedElements(UBound(FactVenta.Log.CreatedElements)).NElement = data.HeaderRow("NFactura")
        data.ClearDoc()
    End Sub

    <Task()> Public Shared Function Resultado(ByVal data As Object, ByVal services As ServiceProvider) As ResultFacturacion
        Dim InfoPrc As ProcessInfoFra = services.GetService(Of ProcessInfoFra)()
        Dim Result As New ResultFacturacion
        If InfoPrc.ConPropuesta Then
            Dim DataFactVenta As DataPrcActualizarFactura = services.GetService(Of DataPrcActualizarFactura)()
            Result = DataFactVenta.RstFacturacion
        Else : Result = services.GetService(Of ResultFacturacion)()
        End If
        Return Result
    End Function

    Protected Overrides Function OnException(ByVal exceptionArgs As Engine.BE.BusinessProcesses.ProcessExceptionArgs) As Engine.BE.BusinessProcesses.OnExceptionBehaviour
        If AdminData.ExistsTX Then ProcessServer.ExecuteTask(Of Boolean)(AddressOf Comunes.RollbackTransaction, True, New ServiceProvider())

        Dim Datafvr As DataPrcActualizarFactura = exceptionArgs.Services.GetService(Of DataPrcActualizarFactura)()
        Dim fvr As ResultFacturacion = Datafvr.RstFacturacion
        Dim log As LogProcess = fvr.Log
        ReDim Preserve log.Errors(log.Errors.Length)

        If TypeOf exceptionArgs.TaskData Is DocumentoFacturaVenta Then
            Dim doc As DocumentoFacturaVenta = CType(exceptionArgs.TaskData, DocumentoFacturaVenta)
            Dim fra As FraCab = CType(exceptionArgs.TaskData, DocumentoFacturaVenta).Cabecera
            If TypeOf fra Is FraCabAlbaran Then
                Dim FraAlb As FraCabAlbaran = CType(fra, FraCabAlbaran)
                Select Case FraAlb.Agrupacion
                    Case enummcAgrupFactura.mcAlbaran
                        log.Errors(log.Errors.Length - 1) = New ClassErrors(" Albarán: " & FraAlb.NAlbaran, exceptionArgs.Exception.Message)
                    Case enummcAgrupFactura.mcCliente
                        log.Errors(log.Errors.Length - 1) = New ClassErrors(" Cliente: " & FraAlb.IDCliente, exceptionArgs.Exception.Message)
                End Select
                doc.ClearDoc()
            ElseIf TypeOf fra Is FraCabObra Then
                Dim FraObra As FraCabObra = CType(fra, FraCabObra)
                Select Case FraObra.AgrupacionObra
                    Case enummcAgrupFacturaObra.mcCliente
                        log.Errors(log.Errors.Length - 1) = New ClassErrors(" Cliente: " & FraObra.IDCliente, exceptionArgs.Exception.Message)
                    Case enummcAgrupFacturaObra.mcObra
                        log.Errors(log.Errors.Length - 1) = New ClassErrors(" Obra: " & FraObra.NObra, exceptionArgs.Exception.Message)
                    Case enummcAgrupFacturaObra.mcObraPedidoClte
                        log.Errors(log.Errors.Length - 1) = New ClassErrors(" Obra/PedidoCliente: " & FraObra.NObra & "/" & FraObra.NumeroPedido, exceptionArgs.Exception.Message)
                    Case enummcAgrupFacturaObra.mcObraTrabajo
                        log.Errors(log.Errors.Length - 1) = New ClassErrors(" Obra/Trabajo: " & FraObra.NObra & "/" & FraObra.IDTrabajo, exceptionArgs.Exception.Message)
                End Select
            ElseIf TypeOf fra Is FraCabMnto Then
                Dim FraOT As FraCabMnto = CType(fra, FraCabMnto)
                Select Case FraOT.Agrupacion
                    Case enummcAgrupOT.Cliente
                        log.Errors(log.Errors.Length - 1) = New ClassErrors(" Cliente: " & FraOT.IDCliente, exceptionArgs.Exception.Message)
                    Case enummcAgrupOT.OT
                        log.Errors(log.Errors.Length - 1) = New ClassErrors(" OT: " & FraOT.NROT, exceptionArgs.Exception.Message)
                End Select
            End If
        Else
            log.Errors(log.Errors.Length - 1) = New ClassErrors(String.Empty, exceptionArgs.Exception.Message)
        End If

        Return MyBase.OnException(exceptionArgs)
    End Function

End Class
