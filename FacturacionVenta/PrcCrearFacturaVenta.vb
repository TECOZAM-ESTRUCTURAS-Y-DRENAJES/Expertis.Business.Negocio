'Proceso de Creación de la factura establecido de forma estándar, relación de tareas
Imports Solmicro.Expertis.Engine.BE.BusinessProcesses
Public Class PrcCrearFacturaVenta
    Inherits Process

    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of FraCab, DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CrearDocumentoFactura)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarValoresPredeterminadosGenerales)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarDescuentoFactura)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarClienteGrupo)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarDatosCliente)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarDireccion)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarDatosFiscales)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarBanco)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoComercial.AsignarObservacionesComercial)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoComunes.AsignarCentroGestion)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarContador)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarNumeroFacturaPropuesta)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoComercial.AsignarEjercicio)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoComunes.AsignarCambiosMoneda)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoComunes.AsignarRetencionPorGarantiaObra)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf NegocioGeneral.AsignarEnvio347Doc)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf NegocioGeneral.AsignarEnvio349Doc)

        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CrearLineasDesdeAlbaran)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularImporteLineasFacturas)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf NegocioGeneral.CalcularImpuestos)


        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CopiarRepresentantes)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CopiarAnalitica)

        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularBasesImponibles)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularTotales)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionObras.NuevaLineaFacturaObraEntregaCuenta)

        '      Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularPuntoVerde)
        'Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarClaveOperacion)

        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf AñadirAResultado)

    End Sub

    <Task()> Public Shared Sub AñadirAResultado(ByVal oDocFra As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        Dim InfoFra As ProcessInfoFra = services.GetService(Of ProcessInfoFra)()
        If InfoFra.ConPropuesta Then
            Dim arDocFras As ArrayList = services.GetService(Of ArrayList)()
            arDocFras.Add(oDocFra)

            Dim Facturas As DataTable = services.GetService(Of ResultFacturacion)().PropuestaFacturas
            Facturas.Rows.Add(oDocFra.HeaderRow.ItemArray)
        Else
            ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf PrcActualizarFactura.ActualizarDocumentoFactura, oDocFra, services)
        End If
    End Sub

    Protected Overrides Function OnException(ByVal exceptionArgs As Engine.BE.BusinessProcesses.ProcessExceptionArgs) As Engine.BE.BusinessProcesses.OnExceptionBehaviour
        If AdminData.ExistsTX Then ProcessServer.ExecuteTask(Of Boolean)(AddressOf Comunes.RollbackTransaction, True, New ServiceProvider())

        Dim fvr As ResultFacturacion = exceptionArgs.Services.GetService(Of ResultFacturacion)()
        Dim log As LogProcess = fvr.Log
        ReDim Preserve log.Errors(log.Errors.Length)
        If TypeOf exceptionArgs.TaskData Is DocumentoFacturaVenta Then
            Dim fra As FraCab = CType(exceptionArgs.TaskData, DocumentoFacturaVenta).Cabecera
            If TypeOf fra Is FraCabAlbaran Then
                Dim FraAlb As FraCabAlbaran = CType(fra, FraCabAlbaran)
                Select Case FraAlb.Agrupacion
                    Case enummcAgrupFactura.mcAlbaran
                        log.Errors(log.Errors.Length - 1) = New ClassErrors(" Albarán: " & FraAlb.NAlbaran, exceptionArgs.Exception.Message)
                    Case enummcAgrupFactura.mcCliente
                        log.Errors(log.Errors.Length - 1) = New ClassErrors(" Cliente: " & FraAlb.IDCliente, exceptionArgs.Exception.Message)
                End Select
            End If
        Else
            log.Errors(log.Errors.Length - 1) = New ClassErrors(String.Empty, exceptionArgs.Exception.Message)
        End If

        Return MyBase.OnException(exceptionArgs)
    End Function

End Class