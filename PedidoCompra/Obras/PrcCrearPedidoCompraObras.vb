Public Class PrcCrearPedidoCompraObras
    Inherits Process(Of DataPrcCrearPedidoCompraObras, LogProcess)

    'Lista de tareas a ejecutar
    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of DataPrcCrearPedidoCompraObras)(AddressOf ValidacionesPrevias)
        Me.AddTask(Of DataPrcCrearPedidoCompraObras)(AddressOf DatosIniciales)
        Me.AddTask(Of DataPrcCrearPedidoCompraObras, PedCabCompra())(AddressOf ProcesoPedidoCompra.AgruparObras)
        Me.AddTask(Of PedCabCompra())(AddressOf ProcesoPedidoCompra.Ordenar)
        'Bucle para recorrer todos los documentos a factura a generar
        Me.AddForEachTask(Of PrcCrearPedidoCompra)(OnExceptionBehaviour.NextLoop)
        Me.AddTask(Of Object, LogProcess)(AddressOf ProcesoComunes.ResultadoLogProcess)
    End Sub

    <Task()> Public Shared Sub ValidacionesPrevias(ByVal data As DataPrcCrearPedidoCompraObras, ByVal services As ServiceProvider)
        'ProcessServer.ExecuteTask(Of String)(AddressOf ProcesoComunes.ValidarContadorObligatorio, data.IDContador & String.Empty, services)

        Dim datosValCont As New ProcesoComunes.DataValidarContadorEntidad(data.IDContador, GetType(PedidoCompraCabecera).Name)
        ProcessServer.ExecuteTask(Of ProcesoComunes.DataValidarContadorEntidad)(AddressOf ProcesoComunes.ValidarContadorEntidad, datosValCont, services)
    End Sub

    <Task()> Public Shared Sub DatosIniciales(ByVal data As DataPrcCrearPedidoCompraObras, ByVal services As ServiceProvider)
        '//Prepara en el service información del proceso.
        services.RegisterService(New ProcessInfoPC(data.IDContador, data.IDOperario, data.FechaEntrega))

        '//Prepara en el service lo que va a retornar el proceso, para tenerlo disponible en todo el proceso.
        services.RegisterService(New LogProcess)
    End Sub

End Class
