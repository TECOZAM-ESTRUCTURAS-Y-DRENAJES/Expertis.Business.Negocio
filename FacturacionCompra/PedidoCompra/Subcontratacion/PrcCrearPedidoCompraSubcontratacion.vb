Imports Solmicro.Expertis.Engine.BE.BusinessProcesses
Public Class PrcCrearPedidoCompraSubcontratacion
    Inherits Process(Of DataPrcCrearPedidoCompraSubcontratacion, LogProcess)

    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of DataPrcCrearPedidoCompraSubcontratacion)(AddressOf ValidacionesPrevias)
        Me.AddTask(Of DataPrcCrearPedidoCompraSubcontratacion)(AddressOf DatosIniciales)
        Me.AddTask(Of DataPrcCrearPedidoCompraSubcontratacion, PedCabCompraSubcontratacion())(AddressOf ProcesoPedidoCompra.AgruparSubcontrataciones)
        Me.AddTask(Of PedCabCompraSubcontratacion())(AddressOf ProcesoPedidoCompra.Ordenar)
        'Bucle para recorrer todos los documentos a factura a generar
        Me.AddForEachTask(Of PrcCrearPedidoCompra)(OnExceptionBehaviour.NextLoop)
        Me.AddTask(Of Object, LogProcess)(AddressOf ProcesoComunes.ResultadoLogProcess)
    End Sub

    <Task()> Public Shared Sub DatosIniciales(ByVal data As DataPrcCrearPedidoCompraSubcontratacion, ByVal services As ServiceProvider)
        If Length(data.IDContador) > 0 Then
            services.RegisterService(New ProcessInfoSubcontratacion(data.IDContador, data.AgruparPorProveedor))
        End If

        '//Prepara en el service lo que va a retornar el proceso, para tenerlo disponible en todo el proceso.
        services.RegisterService(New LogProcess)
    End Sub

    <Task()> Public Shared Sub ValidacionesPrevias(ByVal data As DataPrcCrearPedidoCompraSubcontratacion, ByVal services As ServiceProvider)
        ' ProcessServer.ExecuteTask(Of String)(AddressOf ProcesoComunes.ValidarContadorObligatorio, data.IDContador & String.Empty, services)

        Dim datosValCont As New ProcesoComunes.DataValidarContadorEntidad(data.IDContador, GetType(PedidoCompraCabecera).Name)
        ProcessServer.ExecuteTask(Of ProcesoComunes.DataValidarContadorEntidad)(AddressOf ProcesoComunes.ValidarContadorEntidad, datosValCont, services)
    End Sub

End Class



