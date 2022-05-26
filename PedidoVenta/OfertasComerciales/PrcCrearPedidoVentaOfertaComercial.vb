Public Class PrcCrearPedidoVentaOfertaComercial
    Inherits Process(Of DataPrcCrearPedidoOfertaComercial, LogProcess)

    'Lista de tareas a ejecutar
    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of DataPrcCrearPedidoOfertaComercial)(AddressOf ValidacionesPrevias)
        Me.AddTask(Of DataPrcCrearPedidoOfertaComercial)(AddressOf DatosIniciales)
        Me.AddTask(Of DataPrcCrearPedidoOfertaComercial, PedCabOfertaComercial())(AddressOf ProcesoPedidoVenta.AgruparOfertas)
        'Me.AddTask(Of PedCabPrograma())(AddressOf ProcesoPedidoVenta.Ordenar)
        'Bucle para recorrer todos los documentos a factura a generar
        Me.AddForEachTask(Of PrcCrearPedidoVenta)(OnExceptionBehaviour.NextLoop)
        Me.AddTask(Of Object, LogProcess)(AddressOf ProcesoComunes.ResultadoLogProcess)
    End Sub

    <Task()> Public Shared Sub ValidacionesPrevias(ByVal data As DataPrcCrearPedidoOfertaComercial, ByVal services As ServiceProvider)
        ' ProcessServer.ExecuteTask(Of String)(AddressOf ProcesoComunes.ValidarContadorObligatorio, data.IDContador & String.Empty, services)

        Dim datosValCont As New ProcesoComunes.DataValidarContadorEntidad(data.IDContador, GetType(PedidoVentaCabecera).Name)
        ProcessServer.ExecuteTask(Of ProcesoComunes.DataValidarContadorEntidad)(AddressOf ProcesoComunes.ValidarContadorEntidad, datosValCont, services)
    End Sub

    <Task()> Public Shared Sub DatosIniciales(ByVal data As DataPrcCrearPedidoOfertaComercial, ByVal services As ServiceProvider)
        '//Prepara en el service información del proceso.
        services.RegisterService(New ProcessInfoPV(data.IDContador, data.Detalle))

        '//Prepara en el service lo que va a retornar el proceso, para tenerlo disponible en todo el proceso.
        services.RegisterService(New LogProcess)
    End Sub

End Class
