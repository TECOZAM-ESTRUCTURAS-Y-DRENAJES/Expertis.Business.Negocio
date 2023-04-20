Imports Solmicro.Expertis.Engine.BE.BusinessProcesses

Public Class PrcCrearPedidoCompraPlanificacion
    Inherits Process(Of DataPrcCrearPedidoCompraPlanificacion, LogProcess)

    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of DataPrcCrearPedidoCompraPlanificacion)(AddressOf ValidacionesPrevias)
        Me.AddTask(Of DataPrcCrearPedidoCompraPlanificacion)(AddressOf DatosIniciales)
        Me.AddTask(Of DataPrcCrearPedidoCompraPlanificacion, PedCabCompraPlanif())(AddressOf ProcesoPedidoCompra.AgruparPlanificaciones)
        Me.AddTask(Of PedCabCompraPlanif())(AddressOf ProcesoPedidoCompra.Ordenar)
        Me.AddForEachTask(Of PrcCrearPedidoCompra)(OnExceptionBehaviour.NextLoop)
        Me.AddTask(Of Object, LogProcess)(AddressOf ProcesoComunes.ResultadoLogProcess)
    End Sub

    <Task()> Public Shared Sub DatosIniciales(ByVal data As DataPrcCrearPedidoCompraPlanificacion, ByVal services As ServiceProvider)
        '//Prepara en el service información del proceso.
        services.RegisterService(New ProcessInfoPlanif(data.IDContador, data.AgruparPorProveedor))

        '//Prepara en el service lo que va a retornar el proceso, para tenerlo disponible en todo el proceso.
        services.RegisterService(New LogProcess)
    End Sub

    <Task()> Public Shared Sub ValidacionesPrevias(ByVal data As DataPrcCrearPedidoCompraPlanificacion, ByVal services As ServiceProvider)
        'ProcessServer.ExecuteTask(Of String)(AddressOf ProcesoComunes.ValidarContadorObligatorio, data.IDContador & String.Empty, services)

        Dim datosValCont As New ProcesoComunes.DataValidarContadorEntidad(data.IDContador, GetType(PedidoCompraCabecera).Name)
        ProcessServer.ExecuteTask(Of ProcesoComunes.DataValidarContadorEntidad)(AddressOf ProcesoComunes.ValidarContadorEntidad, datosValCont, services)
    End Sub

End Class
