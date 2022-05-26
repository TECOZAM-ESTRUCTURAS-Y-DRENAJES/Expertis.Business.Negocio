Imports Solmicro.Expertis.Engine.BE.BusinessProcesses
Public Class PrcCrearPedidoCompraOfertaCompra
    Inherits Process(Of DataPrcCrearPedidoCompraOfertaCompra, LogProcess)

    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of DataPrcCrearPedidoCompraOfertaCompra)(AddressOf DatosIniciales)
        Me.AddTask(Of DataPrcCrearPedidoCompraOfertaCompra, PedCabCompraOfertaCompra())(AddressOf ProcesoPedidoCompra.AgruparOfertasCompra)
        Me.AddTask(Of PedCabCompraOfertaCompra())(AddressOf ProcesoPedidoCompra.Ordenar)
        'Bucle para recorrer todos los documentos a factura a generar
        Me.AddForEachTask(Of PrcCrearPedidoCompra)(OnExceptionBehaviour.NextLoop)
        Me.AddTask(Of Object, LogProcess)(AddressOf ProcesoComunes.ResultadoLogProcess)
    End Sub

    <Task()> Public Shared Sub DatosIniciales(ByVal data As DataPrcCrearPedidoCompraOfertaCompra, ByVal services As ServiceProvider)
        services.RegisterService(New ProcessInfo(Nothing))
        '//Prepara en el service lo que va a retornar el proceso, para tenerlo disponible en todo el proceso.
        services.RegisterService(New LogProcess)
    End Sub

End Class
