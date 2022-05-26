Public Class PrcCrearPedidoCompraDesdePedidoVenta
    Inherits Process(Of DataPrcCrearPedidoCompraDesdePedidoVenta, DataResultadoMultiempresaPC)

    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of DataPrcCrearPedidoCompraDesdePedidoVenta)(AddressOf Comunes.BeginTransaction)
        Me.AddTask(Of DataPrcCrearPedidoCompraDesdePedidoVenta)(AddressOf ValidacionesPrevias)
        Me.AddTask(Of DataPrcCrearPedidoCompraDesdePedidoVenta)(AddressOf DatosIniciales)
        Me.AddTask(Of DataPrcCrearPedidoCompraDesdePedidoVenta, PedCabCompraPedidoVenta())(AddressOf ProcesoPedidoCompra.AgruparPedidosVenta)
        Me.AddForEachTask(Of PrcCrearPedidoCompra)()
        Me.AddTask(Of Object, DataResultadoMultiempresaPC)(AddressOf GetResultadoMultiempresa)
    End Sub

    <Task()> Public Shared Sub DatosIniciales(ByVal data As DataPrcCrearPedidoCompraDesdePedidoVenta, ByVal services As ServiceProvider)
        services.RegisterService(New ProcessInfo(data.IDContador))

        '//Prepara en el service lo que va a retornar el proceso, para tenerlo disponible en todo el proceso.
        services.RegisterService(New LogProcess)

        '//Preparamos el retorno de los resultados.
        services.RegisterService(New DataResultadoMultiempresaPC)

    End Sub

    <Task()> Public Shared Sub ValidacionesPrevias(ByVal data As DataPrcCrearPedidoCompraDesdePedidoVenta, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataPrcCrearPedidoCompraDesdePedidoVenta)(AddressOf ValidarContadorObligatorio, data, services)
        ProcessServer.ExecuteTask(Of DataPrcCrearPedidoCompraDesdePedidoVenta)(AddressOf ValidarProveedorObligatorio, data, services)
        ProcessServer.ExecuteTask(Of DataPrcCrearPedidoCompraDesdePedidoVenta)(AddressOf ValidarFechaEntregaObligatoria, data, services)
        ProcessServer.ExecuteTask(Of DataPrcCrearPedidoCompraDesdePedidoVenta)(AddressOf ValidarBaseDatosMultiempresa, data, services)
    End Sub
    <Task()> Public Shared Sub ValidarContadorObligatorio(ByVal data As DataPrcCrearPedidoCompraDesdePedidoVenta, ByVal services As ServiceProvider)
        If Length(data.IDContador) = 0 Then ApplicationService.GenerateError("El contador es obligatorio.")
    End Sub
    <Task()> Public Shared Sub ValidarProveedorObligatorio(ByVal data As DataPrcCrearPedidoCompraDesdePedidoVenta, ByVal services As ServiceProvider)
        Dim f As New Filter
        f.Add(New IsNullFilterItem("IDProveedor"))
        Dim Where As String = f.Compose(New AdoFilterComposer)
        Dim adr() As DataRow = data.Propuestas.Select(Where)
        If Not adr Is Nothing AndAlso adr.Length > 0 Then
            ApplicationService.GenerateError("El Proveedor es obligatorio para todos los registros que inician el proceso.")
        End If
    End Sub
    <Task()> Public Shared Sub ValidarFechaEntregaObligatoria(ByVal data As DataPrcCrearPedidoCompraDesdePedidoVenta, ByVal services As ServiceProvider)
        Dim f As New Filter
        f.Add(New IsNullFilterItem("FechaEntrega"))
        Dim Where As String = f.Compose(New AdoFilterComposer)
        Dim adr() As DataRow = data.Propuestas.Select(Where)
        If Not adr Is Nothing AndAlso adr.Length > 0 Then
            ApplicationService.GenerateError("La Fecha de Entrega es obligatoria para todos los registros que inician el proceso.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarBaseDatosMultiempresa(ByVal data As DataPrcCrearPedidoCompraDesdePedidoVenta, ByVal services As ServiceProvider)
        If New Parametro().EmpresaGrupo Then
            Dim detallesView As New DataView(data.Propuestas)
            Dim f As New Filter
            f.Add(New BooleanFilterItem("EmpresaGrupo", True))
            f.Add(New IsNullFilterItem("BaseDatos"))
            Dim Where As String = f.Compose(New AdoFilterComposer)
            Dim adr() As DataRow = data.Propuestas.Select(Where)
            If Not adr Is Nothing AndAlso adr.Length > 0 Then
                ApplicationService.GenerateError("El Proveedor {0} no tiene asignada una Base de Datos válida.", Quoted(detallesView(0).Row("IDProveedor")))
            End If
        End If
    End Sub

    <Task()> Public Shared Function GetResultadoMultiempresa(ByVal data As Object, ByVal services As ServiceProvider) As DataResultadoMultiempresaPC
        Return services.GetService(Of DataResultadoMultiempresaPC)()
    End Function

End Class


