Imports Solmicro.Expertis.Engine.BE.BusinessProcesses

Public Class PrcCrearFraCompraLeasing
    Inherits Process

    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of FraCabCompra, DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CrearDocumentoFactura)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarValoresPredeterminadosGenerales)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarProveedorGrupo)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarDatosProveedor)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarDireccion)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarBanco)

        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoCompra.AsignarObservacionesCompra)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoComunes.AsignarCentroGestion)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarContador)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarNumeroFacturaPropuesta)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoCompra.AsignarEjercicio)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarSuFactura)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarSuFechaFactura)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoComunes.AsignarCambiosMoneda)
        'Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoComunes.AsignarRetencionPorGarantiaObra)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarEstadoFactura)

        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf NegocioGeneral.AsignarEnvio347Doc)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf NegocioGeneral.AsignarEnvio349Doc)

        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CrearLineasDesdePagosLeasing)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularImporteLineasFacturas)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf NegocioGeneral.CalcularImpuestos)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CopiarAnalitica)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularBasesImponibles)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularTotales)
        'Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.NuevaLineaFacturaObraEntregaCuenta)
        ' Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarClaveOperacion)
        ' Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.ValidarDocumento)

        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf AñadirAResultado)
    End Sub

    <Task()> Public Shared Sub AñadirAResultado(ByVal oDocFra As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        Dim Facturas As DataTable = services.GetService(Of ResultFacturacion)().PropuestaFacturas
        Facturas.Rows.Add(oDocFra.HeaderRow.ItemArray)

        Dim arDocFras As ArrayList = services.GetService(Of ArrayList)()

        arDocFras.Add(oDocFra)
    End Sub

    Protected Overrides Function OnException(ByVal exceptionArgs As Engine.BE.BusinessProcesses.ProcessExceptionArgs) As Engine.BE.BusinessProcesses.OnExceptionBehaviour
        ProcessServer.ExecuteTask(Of Boolean)(AddressOf Comunes.RollbackTransaction, True, New ServiceProvider())

        Dim fvr As ResultFacturacion = exceptionArgs.Services.GetService(Of ResultFacturacion)()
        Dim log As LogProcess = fvr.Log
        ReDim Preserve log.Errors(log.Errors.Length)
        If TypeOf exceptionArgs.TaskData Is DocumentoFacturaCompra Then
            Dim fra As FraCabCompra = CType(exceptionArgs.TaskData, DocumentoFacturaCompra).Cabecera
            Select Case fra.Agrupacion
                Case enummpAgrupFactura.mpAlbaran
                    log.Errors(log.Errors.Length - 1) = New ClassErrors(" Albarán: " & CType(fra, FraCabCompraAlbaran).NAlbaran, exceptionArgs.Exception.Message)
                Case enummpAgrupFactura.mpProveedor
                    log.Errors(log.Errors.Length - 1) = New ClassErrors(" Proveedor: " & fra.IDProveedor, exceptionArgs.Exception.Message)
            End Select
        Else
            log.Errors(log.Errors.Length - 1) = New ClassErrors(String.Empty, exceptionArgs.Exception.Message)
        End If
        Return MyBase.OnException(exceptionArgs)
    End Function

End Class


