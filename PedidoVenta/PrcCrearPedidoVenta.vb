Imports Solmicro.Expertis.Engine.BE.BusinessProcesses

Public Class PrcCrearPedidoVenta
    Inherits Process


    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of PedCab, DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.CrearDocumentoPedidoVenta)
        Me.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.AsignarValoresPredeterminadosGenerales)
        Me.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.AsignarDatosCliente)
        Me.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.AsignarPedidoCliente)
        Me.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoComunes.ActualizarCambiosMoneda)
        Me.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoComunes.AsignarCentroGestion)
        Me.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoComunes.AsignarAlmacen)
        Me.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoComunes.AsignarContador)
        Me.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.AsignarNumeroPedido)
        Me.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.AsignarDatosCabeceraEDI)
        Me.AddTask(Of DocumentoPedidoVenta)(AddressOf AsignarObservacionesInternas)
        Me.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoComercial.AsignarObservacionesComercial)
        Me.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.AsignarDireccionEnvio)
        Me.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoComercial.AsignarCondicionesEnvio)
        Me.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoComercial.AsignarEjercicio)
        Me.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.CrearLineasDesdeOrigen)
        Me.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.CalcularImporteLineasPedido)
        Me.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.LineasDeRegalo)
        Me.AddTask(Of DocumentoPedidoVenta)(AddressOf NegocioGeneral.CalcularAnalitica)
        Me.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoComercial.CalcularRepresentantes)
        Me.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.CalcularBasesImponibles)
        Me.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoComunes.TotalDocumento)
        Me.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.GrabarDocumento)
        Me.AddTask(Of DocumentoPedidoVenta)(AddressOf AñadirAResultado)
    End Sub

    <Task()> Public Shared Sub AsignarObservacionesInternas(ByVal Doc As DocumentoPedidoVenta, ByVal services As ServiceProvider)
        If Not Doc.Cabecera Is Nothing Then
            Select Case Doc.Cabecera.Origen
                Case enumOrigenPedido.PedidoCompra
                    Dim cab As PedCabVentaPedidoCompra = CType(Doc.Cabecera, PedCabVentaPedidoCompra)
                    Doc.HeaderRow("Texto") = cab.Texto
                Case enumOrigenPedido.Programa
                    Doc.HeaderRow("Texto") = CType(Doc.Cabecera, PedCabPrograma).Texto
            End Select
        End If
    End Sub

    <Task()> Public Shared Sub AñadirAResultado(ByVal Doc As DocumentoPedidoVenta, ByVal services As ServiceProvider)
        Dim log As LogProcess = services.GetService(Of LogProcess)()
        If log Is Nothing Then log = New LogProcess
        ReDim Preserve log.CreatedElements(UBound(log.CreatedElements) + 1)
        log.CreatedElements(UBound(log.CreatedElements)) = New CreateElement
        log.CreatedElements(UBound(log.CreatedElements)).IDElement = Doc.HeaderRow("IDPedido")
        log.CreatedElements(UBound(log.CreatedElements)).NElement = Doc.HeaderRow("NPedido")

        ProcessServer.ExecuteTask(Of DocumentoPedidoVenta)(AddressOf PrepararResultadoMultiempresa, Doc, services)
    End Sub

    <Task()> Public Shared Sub PrepararResultadoMultiempresa(ByVal Doc As DocumentoPedidoVenta, ByVal services As ServiceProvider)
        If Not Doc.Cabecera Is Nothing Then

            Select Case Doc.Cabecera.Origen
                Case enumOrigenPedido.PedidoCompra
                    Dim cab As PedCabVentaPedidoCompra = CType(Doc.Cabecera, PedCabVentaPedidoCompra)
                    Dim PedidosCompra As DataResultadoMultiempresaPC = services.GetService(Of DataResultadoMultiempresaPC)()
                    Dim BDInfo As DataBasesDatosMultiempresa = services.GetService(Of DataBasesDatosMultiempresa)()
                    Dim PCInfo As GeneracionPedidosCompraInfo = PedidosCompra.Item(cab.IDPedido)
                    If Not PCInfo Is Nothing Then
                        '//Estamos en el Pedido de Venta de la Base de Datos secundaria
                        PCInfo.IDPedidoVenta2 = Doc.HeaderRow("IDPedido")
                        PCInfo.NPedidoVenta2 = Doc.HeaderRow("NPedido")
                        PCInfo.BaseDatos2 = BDInfo.DescBaseDatosSecundaria
                        Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
                        Dim ClteInfo As ClienteInfo = Clientes.GetEntity(Doc.HeaderRow("IDCliente"))
                        PCInfo.Cliente = Doc.HeaderRow("IDCliente") & " - " & ClteInfo.DescCliente
                    End If
            End Select
        End If
    End Sub

    Protected Overrides Function OnException(ByVal exceptionArgs As Engine.BE.BusinessProcesses.ProcessExceptionArgs) As Engine.BE.BusinessProcesses.OnExceptionBehaviour
        AdminData.RollBackTx(True)

        Dim log As LogProcess = exceptionArgs.Services.GetService(Of LogProcess)()
        ReDim Preserve log.Errors(log.Errors.Length)
        If TypeOf exceptionArgs.TaskData Is DocumentoPedidoVenta Then
            Dim ped As DocumentoPedidoVenta = exceptionArgs.TaskData
            Select Case ped.Cabecera.Origen
                Case enumOrigenPedido.Programa
                    Select Case CType(ped.Cabecera, PedCabPrograma).Agrupacion
                        Case enummcAgrupPedido.mcPrograma
                            log.Errors(log.Errors.Length - 1) = New ClassErrors(" Programa: " & CType(ped.Cabecera, PedCabPrograma).IDPrograma, exceptionArgs.Exception.Message)
                        Case enummcAgrupPedido.mcCliente
                            log.Errors(log.Errors.Length - 1) = New ClassErrors(" Cliente: " & ped.IDCliente, exceptionArgs.Exception.Message)
                    End Select
                Case enumOrigenPedido.Oferta
                    log.Errors(log.Errors.Length - 1) = New ClassErrors(" Oferta: " & CType(ped.Cabecera, PedCabOfertaComercial).NOferta, exceptionArgs.Exception.Message)
                Case enumOrigenPedido.PedidoCompra
                    log.Errors(log.Errors.Length - 1) = New ClassErrors(" Pedido: " & CType(ped.Cabecera, PedCabVentaPedidoCompra).NPedido, exceptionArgs.Exception.Message)
                    Dim PedidosCompra As DataResultadoMultiempresaPC = exceptionArgs.Services.GetService(Of DataResultadoMultiempresaPC)()
                    Dim BDInfo As DataBasesDatosMultiempresa = exceptionArgs.Services.GetService(Of DataBasesDatosMultiempresa)()
                    Dim PCInfo As GeneracionPedidosCompraInfo = PedidosCompra.Item(CType(ped.Cabecera, PedCabVentaPedidoCompra).IDPedido)
                    If Not PCInfo Is Nothing Then
                        '//Estamos en el Pedido de Venta de la Base de Datos secundaria
                        PCInfo.BaseDatos2 = BDInfo.DescBaseDatosSecundaria
                        PCInfo.StrError = log.Errors(log.Errors.Length - 1).MessageError
                    End If
                Case enumOrigenPedido.EDI
                    log.Errors(log.Errors.Length - 1) = New ClassErrors(" Pedido EDI: " & CType(ped.Cabecera, PedCabEDI).IDPedidoEDI, exceptionArgs.Exception.Message)
            End Select
        Else
            log.Errors(log.Errors.Length - 1) = New ClassErrors(String.Empty, exceptionArgs.Exception.Message)
        End If

        Return MyBase.OnException(exceptionArgs)
    End Function

End Class
