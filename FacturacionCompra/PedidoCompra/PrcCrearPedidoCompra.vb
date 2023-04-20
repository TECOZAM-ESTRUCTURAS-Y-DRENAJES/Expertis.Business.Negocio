Imports Solmicro.Expertis.Engine.BE.BusinessProcesses

Public Class PrcCrearPedidoCompra
    Inherits Process(Of PedCabCompra, DocumentoPedidoCompra)


    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of PedCabCompra, DocumentoPedidoCompra)(AddressOf ProcesoPedidoCompra.CrearDocumentoPedidoCompra)
        Me.AddTask(Of DocumentoPedidoCompra)(AddressOf ProcesoPedidoCompra.AsignarValoresPredeterminadosGenerales)
        Me.AddTask(Of DocumentoPedidoCompra)(AddressOf ProcesoPedidoCompra.AsignarDatosProveedor)
        Me.AddTask(Of DocumentoPedidoCompra)(AddressOf ProcesoComunes.AsignarContador)
        Me.AddTask(Of DocumentoPedidoCompra)(AddressOf ProcesoPedidoCompra.AsignarNumeroPedido)
        Me.AddTask(Of DocumentoPedidoCompra)(AddressOf ProcesoPedidoCompra.AsignarOperario)
        Me.AddTask(Of DocumentoPedidoCompra)(AddressOf ProcesoPedidoCompra.AsignarDiaPago)
        Me.AddTask(Of DocumentoPedidoCompra)(AddressOf ProcesoComunes.AsignarCentroGestion)
        Me.AddTask(Of DocumentoPedidoCompra)(AddressOf ProcesoComunes.ActualizarCambiosMoneda)
        Me.AddTask(Of DocumentoPedidoCompra)(AddressOf ProcesoPedidoCompra.AsignarTipoCompra)
        Me.AddTask(Of DocumentoPedidoCompra)(AddressOf AsignarAlmacen)
        Me.AddTask(Of DocumentoPedidoCompra)(AddressOf ProcesoPedidoCompra.AsignarEmpresaGrupo)
        Me.AddTask(Of DocumentoPedidoCompra)(AddressOf ProcesoPedidoCompra.AsignarEntregaProveedor)
        Me.AddTask(Of DocumentoPedidoCompra)(AddressOf ProcesoPedidoCompra.AsignarPedidoVenta)

        Me.AddTask(Of DocumentoPedidoCompra)(AddressOf ProcesoCompra.AsignarObservacionesCompra)
        Me.AddTask(Of DocumentoPedidoCompra)(AddressOf ProcesoPedidoCompra.AsignarDireccionEnvio)

        Me.AddTask(Of DocumentoPedidoCompra)(AddressOf ProcesoPedidoCompra.CrearLineasPedidoDesdeOrigen)
        Me.AddTask(Of DocumentoPedidoCompra)(AddressOf ProcesoPedidoCompra.CalcularImporteLineasPedido)
        Me.AddTask(Of DocumentoPedidoCompra)(AddressOf NegocioGeneral.CalcularAnalitica)
        Me.AddTask(Of DocumentoPedidoCompra)(AddressOf ProcesoPedidoCompra.CalcularBasesImponibles)
        Me.AddTask(Of DocumentoPedidoCompra)(AddressOf ProcesoComunes.TotalDocumento)
        Me.AddTask(Of DocumentoPedidoCompra)(AddressOf Comunes.BeginTransaction)
        Me.AddTask(Of DocumentoPedidoCompra)(AddressOf ProcesoPedidoCompra.ActualizarEntidadesDependientes)
        Me.AddTask(Of DocumentoPedidoCompra)(AddressOf Comunes.UpdateDocument)
        Me.AddTask(Of DocumentoPedidoCompra)(AddressOf AñadirAResultado)
        Me.AddTask(Of DocumentoPedidoCompra)(AddressOf CrearPedidoVentaEnBDSecundaria)
    End Sub

    <Task()> Public Shared Sub AsignarAlmacen(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        If Length(Doc.HeaderRow("IDAlmacen")) > 0 Then Exit Sub
        Dim AppParamsCompra As ParametroCompra = services.GetService(Of ParametroCompra)()
        If Length(Doc.HeaderRow("IDTipoCompra")) > 0 Then
            Select Case Doc.HeaderRow("IDTipoCompra")
                Case AppParamsCompra.TipoCompraNormal
                    ProcessServer.ExecuteTask(Of DocumentoPedidoCompra)(AddressOf ProcesoComunes.AsignarAlmacen, Doc, services)
                Case AppParamsCompra.TipoCompraSubcontratacion
                    Doc.HeaderRow("IDAlmacen") = CType(Doc.Cabecera, PedCabCompraSubcontratacion).IDAlmacen
            End Select
        Else
            ProcessServer.ExecuteTask(Of DocumentoPedidoCompra)(AddressOf ProcesoComunes.AsignarAlmacen, Doc, services)
        End If
    End Sub

    <Task()> Public Shared Sub AñadirAResultado(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        Dim log As LogProcess = services.GetService(Of LogProcess)()
        If log Is Nothing Then log = New LogProcess
        ReDim Preserve log.CreatedElements(UBound(log.CreatedElements) + 1)
        log.CreatedElements(UBound(log.CreatedElements)) = New CreateElement
        log.CreatedElements(UBound(log.CreatedElements)).IDElement = Doc.HeaderRow("IDPedido")
        log.CreatedElements(UBound(log.CreatedElements)).NElement = Doc.HeaderRow("NPedido")

        ProcessServer.ExecuteTask(Of DocumentoPedidoCompra)(AddressOf PrepararResultadoMultiempresa, Doc, services)
    End Sub

    <Task()> Public Shared Sub PrepararResultadoMultiempresa(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        If Not Doc Is Nothing AndAlso Not Doc.Cabecera Is Nothing Then

            Select Case Doc.Cabecera.Origen
                Case enumOrigenPedidoCompra.PedidoVenta
                    Dim OrigenPedido As PedCabCompraPedidoVenta = CType(Doc.Cabecera, PedCabCompraPedidoVenta)
                    '//Preparamos la información del Pedido Compra creado y del Pedido Venta Origen.
                    Dim PCInfo As New GeneracionPedidosCompraInfo
                    PCInfo.IDPedidoVenta1 = OrigenPedido.IDOrigen
                    PCInfo.NPedidoVenta1 = OrigenPedido.NOrigen

                    PCInfo.IDPedidoCompra = Doc.HeaderRow("IDPedido")
                    PCInfo.NPedidoCompra = Doc.HeaderRow("NPedido")

                    Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
                    Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(Doc.HeaderRow("IDProveedor"))
                    PCInfo.Proveedor = ProvInfo.IDProveedor & " - " & ProvInfo.DescProveedor
                    PCInfo.EmpresaGrupo = OrigenPedido.Multiempresa

                    PCInfo.EntregaProveedor = OrigenPedido.EntregaProveedor
                    PCInfo.BaseDatos1 = AdminData.GetSessionInfo.DataBase.DataBaseDescription

                    Dim PedidosCompra As DataResultadoMultiempresaPC = services.GetService(Of DataResultadoMultiempresaPC)()
                    PedidosCompra.Add(PCInfo)
            End Select
        End If
    End Sub

    <Task()> Public Shared Sub CrearPedidoVentaEnBDSecundaria(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        If Doc Is Nothing OrElse Doc.Cabecera Is Nothing Then Exit Sub
        If New Parametro().EmpresaGrupo Then
            Select Case Doc.Cabecera.Origen
                Case enumOrigenPedidoCompra.PedidoVenta
                    'If Length(CType(Doc.Cabecera, PedCabCompraPedidoVenta).IDBaseDatosSecundaria) > 0 Then
                    If CType(Doc.Cabecera, PedCabCompraPedidoVenta).Multiempresa Then
                        '//Necesitamos volver a cargar el documento, ya que lo hemos guardado ya en el sistema previamente.
                        Dim DocPC As New DocumentoPedidoCompra(Doc.HeaderRow("IDPedido"))
                        Dim dat As New DataPrcCrearPedidoVentaEnBDSecundaria(AdminData.GetSessionInfo.DataBase.DataBaseID, DocPC)
                        Dim PedidosCompra As DataResultadoMultiempresaPC = services.GetService(Of DataResultadoMultiempresaPC)()
                        Dim BDInfo As DataBasesDatosMultiempresa = services.GetService(Of DataBasesDatosMultiempresa)()
                        services = New ServiceProvider
                        services.RegisterService(PedidosCompra)
                        services.RegisterService(BDInfo)
                        ProcessServer.RunProcess(GetType(PrcCrearPedidoVentaEnBDSecundaria), dat, services)
                        'Controlar que si ha dado error no actualizar el pedido de compra
                        'Dim PedidosCompra As DataResultadoMultiempresaPC = services.GetService(Of DataResultadoMultiempresaPC)()
                        'Dim BDInfo As DataBasesDatosMultiempresa = services.GetService(Of DataBasesDatosMultiempresa)()
                        PedidosCompra = services.GetService(Of DataResultadoMultiempresaPC)()
                        Dim PCInfo As GeneracionPedidosCompraInfo = PedidosCompra.Item(Doc.HeaderRow("IDPedido"))
                        If Not PCInfo Is Nothing AndAlso Length(PCInfo.StrError) = 0 Then
                            '//Estamos en el Pedido de Venta de la Base de Datos secundaria
                            ProcessServer.ExecuteTask(Of DocumentoPedidoCompra)(AddressOf Comunes.UpdateDocument, DocPC, services)
                        End If
                    End If
            End Select
        End If
    End Sub

    Protected Overrides Function OnException(ByVal exceptionArgs As Engine.BE.BusinessProcesses.ProcessExceptionArgs) As Engine.BE.BusinessProcesses.OnExceptionBehaviour
        ProcessServer.ExecuteTask(Of Boolean)(AddressOf Comunes.RollbackTransaction, True, New ServiceProvider())

        Dim log As LogProcess = exceptionArgs.Services.GetService(Of LogProcess)()
        ReDim Preserve log.Errors(log.Errors.Length)
        If TypeOf exceptionArgs.TaskData Is DocumentoPedidoCompra Then
            Dim ped As DocumentoPedidoCompra = exceptionArgs.TaskData
            Select Case ped.Cabecera.Origen
                Case enumOrigenPedidoCompra.Programa
                    log.Errors(log.Errors.Length - 1) = New ClassErrors(" Programa: " & CType(ped.Cabecera, PedCabCompra).IDOrigen, exceptionArgs.Exception.Message)
                Case enumOrigenPedidoCompra.OfertaCompra
                    log.Errors(log.Errors.Length - 1) = New ClassErrors(" Oferta: " & CType(ped.Cabecera, PedCabCompra).IDOrigen, exceptionArgs.Exception.Message)
                Case enumOrigenPedidoCompra.Subcontratacion
                    Dim ProcInfo As ProcessInfoSubcontratacion = exceptionArgs.Services.GetService(Of ProcessInfoSubcontratacion)()
                    If ProcInfo.AgruparPorProveedor Then
                        log.Errors(log.Errors.Length - 1) = New ClassErrors(" Proveedor: " & ped.IDProveedor, exceptionArgs.Exception.Message)
                    Else
                        log.Errors(log.Errors.Length - 1) = New ClassErrors(" Orden: " & ped.Cabecera.NOrigen, exceptionArgs.Exception.Message)
                    End If
                Case enumOrigenPedidoCompra.PedidoVenta
                    Dim OrigenPedido As PedCabCompraPedidoVenta = CType(ped.Cabecera, PedCabCompraPedidoVenta)
                    '//Preparamos la información del Pedido Compra creado y del Pedido Venta Origen.
                    Dim PCInfo As New GeneracionPedidosCompraInfo
                    PCInfo.IDPedidoVenta1 = OrigenPedido.IDOrigen
                    PCInfo.NPedidoVenta1 = OrigenPedido.NOrigen
                    PCInfo.BaseDatos1 = AdminData.GetSessionInfo.DataBase.DataBaseDescription
                    PCInfo.StrError = " Pedido: " & ped.Cabecera.NOrigen & " : " & exceptionArgs.Exception.Message

                    Dim PedidosCompra As DataResultadoMultiempresaPC = exceptionArgs.Services.GetService(Of DataResultadoMultiempresaPC)()
                    PedidosCompra.Add(PCInfo)
                Case Else
                    Dim origen As String
                    If Length(CType(ped.Cabecera, PedCabCompra).NOrigen) > 0 Then origen = CType(ped.Cabecera, PedCabCompra).NOrigen
                    If Length(origen) = 0 Then origen = CType(ped.Cabecera, PedCabCompra).IDOrigen

                    log.Errors(log.Errors.Length - 1) = New ClassErrors(" Origen: " & origen, exceptionArgs.Exception.Message)
            End Select
        Else
            log.Errors(log.Errors.Length - 1) = New ClassErrors(String.Empty, exceptionArgs.Exception.Message)
        End If

        Return MyBase.OnException(exceptionArgs)
    End Function

End Class
