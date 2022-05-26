Imports Solmicro.Expertis.Engine.BE.BusinessProcesses
Public Class PrcCrearPedidoCompraSolicitudCompra
    Inherits Process(Of DataPrcCrearPedidoCompraSolicitudCompra, LogProcess)

    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of DataPrcCrearPedidoCompraSolicitudCompra)(AddressOf ValidacionesPrevias)
        Me.AddTask(Of DataPrcCrearPedidoCompraSolicitudCompra)(AddressOf DatosIniciales)
        Me.AddTask(Of DataPrcCrearPedidoCompraSolicitudCompra, PedCabCompraSolicitudCompra())(AddressOf ProcesoPedidoCompra.AgruparSolicitudesCompra)
        Me.AddTask(Of PedCabCompraSolicitudCompra())(AddressOf ProcesoPedidoCompra.Ordenar)
        'Bucle para recorrer todos los documentos a factura a generar
        Me.AddForEachTask(Of PrcCrearPedidoCompra)(OnExceptionBehaviour.NextLoop)
        Me.AddTask(Of Object, LogProcess)(AddressOf ProcesoComunes.ResultadoLogProcess)
    End Sub

    <Task()> Public Shared Sub DatosIniciales(ByVal data As DataPrcCrearPedidoCompraSolicitudCompra, ByVal services As ServiceProvider)
        If Length(data.IDContador) > 0 Then
            services.RegisterService(New ProcessInfo(data.IDContador))
        End If

        '//Prepara en el service lo que va a retornar el proceso, para tenerlo disponible en todo el proceso.
        services.RegisterService(New LogProcess)
    End Sub

    <Task()> Public Shared Sub ValidacionesPrevias(ByVal data As DataPrcCrearPedidoCompraSolicitudCompra, ByVal services As ServiceProvider)
        'ProcessServer.ExecuteTask(Of String)(AddressOf ProcesoComunes.ValidarContadorObligatorio, data.IDContador & String.Empty, services)

        Dim datosValCont As New ProcesoComunes.DataValidarContadorEntidad(data.IDContador, GetType(PedidoCompraCabecera).Name)
        ProcessServer.ExecuteTask(Of ProcesoComunes.DataValidarContadorEntidad)(AddressOf ProcesoComunes.ValidarContadorEntidad, datosValCont, services)
    End Sub


End Class

