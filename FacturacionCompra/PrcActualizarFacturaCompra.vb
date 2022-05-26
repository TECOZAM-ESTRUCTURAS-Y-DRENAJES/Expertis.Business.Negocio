
'Proceso de Creación de la factura establecido de forma estándar, relación de tareas
Imports Solmicro.Expertis.Engine.BE.BusinessProcesses
Public Class PrcActualizarFacturaCompra
    Inherits Process(Of DataPrcActualizarFactura, ResultFacturacion)

    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of DataPrcActualizarFactura, ArrayList)(AddressOf PrepararInformacionProceso)
        Me.AddForEachTask(Of DocumentoFacturaCompra)(AddressOf ActualizarDocumentoFactura, OnExceptionBehaviour.NextLoop)
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

    <Task()> Public Shared Sub ActualizarDocumentoFactura(ByVal doc As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        AdminData.BeginTx()
        Dim DataFact As DataPrcActualizarFactura = services.GetService(Of DataPrcActualizarFactura)()
        Dim Fact As ResultFacturacion = DataFact.RstFacturacion
        Dim dv As New DataView(Fact.PropuestaFacturas, Nothing, "IDFactura", DataViewRowState.CurrentRows)
        Dim idx As Integer = dv.Find(doc.HeaderRow("IDFactura"))
        If idx >= 0 Then
            doc.HeaderRow("IDContador") = dv(idx)("IDcontador")
            doc.HeaderRow("FechaFactura") = dv(idx)("FechaFactura")
            doc.HeaderRow("IDEjercicio") = dv(idx)("IDEjercicio")
            doc.HeaderRow("SuFactura") = dv(idx)("SuFactura")
            doc.HeaderRow("SuFechaFactura") = dv(idx)("SuFechaFactura")
            doc.HeaderRow("FechaParaDeclaracion") = dv(idx)("FechaParaDeclaracion")
            ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ProcesoFacturacionCompra.FechaParaDeclaracionPorProveedor, New DataRowPropertyAccessor(doc.HeaderRow), services)

            Dim StDatos As New Contador.DatosCounterValue
            StDatos.IDCounter = doc.HeaderRow("IDContador")
            StDatos.TargetClass = New FacturaCompraCabecera
            StDatos.TargetField = "NFactura"
            StDatos.DateField = "FechaFactura"
            StDatos.DateValue = doc.HeaderRow("FechaFactura")
            StDatos.IDEjercicio = doc.HeaderRow("IDEjercicio") & String.Empty
            doc.HeaderRow("NFactura") = ProcessServer.ExecuteTask(Of Contador.DatosCounterValue, String)(AddressOf Contador.CounterValue, StDatos, services)
        End If

        ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoFacturacionCompra.ValidarSuFactura, doc.HeaderRow, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf NegocioGeneral.ValidarFechaDeclaracion, doc.HeaderRow, services)

        ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf NegocioGeneral.AsignarEnvio347Doc, doc, services)
        ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf NegocioGeneral.AsignarEnvio349Doc, doc, services)

        'TODO Orden de ejecución de la actualización
        ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularBasesImponibles, doc, services)
        ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularTotales, doc, services)
        ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularVencimientos, doc, services)
        ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarClaveOperacion, doc, services)
        'ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf ProcesoComunes.ValidarIVASDocFC, doc, services)
        ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.GrabarDocumento, doc, services)

        ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.ActualizarQFacturadaAlbaranEnProceso, doc, services)
        ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.ActualizarObras, doc, services)
        ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.ActualizarOTs, doc, services)

        AdminData.CommitTx(True)
        ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf AgregarFacturaResultado, doc, services)
    End Sub

    <Task()> Public Shared Sub AgregarFacturaResultado(ByVal data As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        Dim DataFact As DataPrcActualizarFactura = services.GetService(Of DataPrcActualizarFactura)()
        Dim Fact As ResultFacturacion = DataFact.RstFacturacion
        ReDim Preserve Fact.Log.CreatedElements(UBound(Fact.Log.CreatedElements) + 1)
        Fact.Log.CreatedElements(UBound(Fact.Log.CreatedElements)) = New CreateElement
        Fact.Log.CreatedElements(UBound(Fact.Log.CreatedElements)).IDElement = data.HeaderRow("IDFactura")
        Fact.Log.CreatedElements(UBound(Fact.Log.CreatedElements)).NElement = data.HeaderRow("NFactura")
    End Sub

    <Task()> Public Shared Function Resultado(ByVal data As Object, ByVal services As ServiceProvider) As ResultFacturacion
        Dim DataFactVenta As DataPrcActualizarFactura = services.GetService(Of DataPrcActualizarFactura)()
        Return DataFactVenta.RstFacturacion
    End Function

    Protected Overrides Function OnException(ByVal exceptionArgs As Engine.BE.BusinessProcesses.ProcessExceptionArgs) As Engine.BE.BusinessProcesses.OnExceptionBehaviour
        ProcessServer.ExecuteTask(Of Boolean)(AddressOf Comunes.RollbackTransaction, True, New ServiceProvider())

        Dim Datafr As DataPrcActualizarFactura = exceptionArgs.Services.GetService(Of DataPrcActualizarFactura)()
        Dim fr As ResultFacturacion = Datafr.RstFacturacion
        Dim log As LogProcess = fr.Log
        ReDim Preserve log.Errors(log.Errors.Length)
        If TypeOf exceptionArgs.TaskData Is DocumentoFacturaCompra Then
            Dim fra As DocumentoFacturaCompra = exceptionArgs.TaskData
            Select Case fra.Cabecera.Agrupacion
                Case enummpAgrupFactura.mpAlbaran
                    log.Errors(log.Errors.Length - 1) = New ClassErrors(" Albarán: " & CType(fra.Cabecera, FraCabCompraAlbaran).NAlbaran, exceptionArgs.Exception.Message)
                Case enummpAgrupFactura.mpProveedor
                    log.Errors(log.Errors.Length - 1) = New ClassErrors(" Proveedor: " & fra.IDProveedor, exceptionArgs.Exception.Message)
            End Select
        Else
            log.Errors(log.Errors.Length - 1) = New ClassErrors(String.Empty, exceptionArgs.Exception.Message)
        End If
        Return MyBase.OnException(exceptionArgs)
    End Function

End Class
