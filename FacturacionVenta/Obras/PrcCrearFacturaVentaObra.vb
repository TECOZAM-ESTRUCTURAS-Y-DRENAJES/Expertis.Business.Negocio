Public Class PrcCrearFacturaVentaObra
    Inherits Process

    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of FraCab, DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CrearDocumentoFactura)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarValoresPredeterminadosGenerales)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionObras.AsignarTipoFactura)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarClienteGrupo)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarDatosCliente)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionObras.AsignarDireccion)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionObras.AsignarDiaPago)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarDatosFiscales)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarBanco)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoComercial.AsignarObservacionesComercial)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoComunes.AsignarCentroGestion)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarContador)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarNumeroFacturaPropuesta)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoComercial.AsignarEjercicio)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionObras.AsignarCambiosMoneda)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoComunes.AsignarRetencionPorGarantiaObra)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf NegocioGeneral.AsignarEnvio347Doc)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf NegocioGeneral.AsignarEnvio349Doc)

        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionObras.CrearLineasDesdeObras)

        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularImporteLineasFacturas)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf NegocioGeneral.CalcularImpuestos)

        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionObras.CalcularSegurosTasasAlquiler)
        'Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CopiarRepresentantes)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionObras.CopiarAnalitica)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularBasesImponibles)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularTotales)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionObras.NuevaLineaFacturaObraEntregaCuenta)

        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionObras.RecalcularDireccion)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf AñadirAResultado)
    End Sub

    <Task()> Public Shared Sub AñadirAResultado(ByVal oDocFra As DocumentoFacturaVenta, ByVal services As ServiceProvider)
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
        If TypeOf exceptionArgs.TaskData Is DocumentoFacturaVenta Then

            Dim fra As FraCabObra = CType(exceptionArgs.TaskData, DocumentoFacturaVenta).Cabecera
            Select Case fra.AgrupacionObra
                Case enummcAgrupFacturaObra.mcCliente
                    log.Errors(log.Errors.Length - 1) = New ClassErrors(" Cliente: " & fra.IDCliente, exceptionArgs.Exception.Message)
                Case enummcAgrupFacturaObra.mcObra
                    log.Errors(log.Errors.Length - 1) = New ClassErrors(" Obra: " & fra.NObra, exceptionArgs.Exception.Message)
                Case enummcAgrupFacturaObra.mcObraPedidoClte
                    log.Errors(log.Errors.Length - 1) = New ClassErrors(" Obra/PedidoCliente: " & fra.NObra & "/" & fra.NumeroPedido, exceptionArgs.Exception.Message)
                Case enummcAgrupFacturaObra.mcObraTrabajo
                    log.Errors(log.Errors.Length - 1) = New ClassErrors(" Obra/Trabajo: " & fra.NObra & "/" & fra.IDTrabajo, exceptionArgs.Exception.Message)
            End Select
        Else
            log.Errors(log.Errors.Length - 1) = New ClassErrors(String.Empty, exceptionArgs.Exception.Message)
        End If
        Return MyBase.OnException(exceptionArgs)
    End Function

End Class