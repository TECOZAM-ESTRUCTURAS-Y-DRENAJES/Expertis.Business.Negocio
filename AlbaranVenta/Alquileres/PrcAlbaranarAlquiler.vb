Public Class PrcAlbaranarAlquiler
    Inherits Process(Of DataPrcAlbaranar, ResultAlbaranAlquiler)

    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of DataPrcAlbaranar)(AddressOf ValidacionesPrevias)
        Me.AddTask(Of DataPrcAlbaranar)(AddressOf DatosIniciales)
        Me.AddTask(Of DataPrcAlbaranar, AlbCabVentaAlquiler())(AddressOf ProcesoAlbaranVentaAlquiler.AgruparAlquiler)
        'Bucle para recorrer todos los documentos a expedir a generar
        Me.AddForEachTask(Of PrcCrearAlbaranVentaAlquiler)(OnExceptionBehaviour.NextLoop)
        Me.AddTask(Of Object, ResultAlbaranAlquiler)(AddressOf ProcesoAlbaranVentaAlquiler.Resultado)
    End Sub

    <Task()> Public Shared Sub ValidacionesPrevias(ByVal data As DataPrcAlbaranar, ByVal services As ServiceProvider)
        Dim datosValCont As New ProcesoComunes.DataValidarContadorEntidad(data.IDContador, GetType(AlbaranVentaCabecera).Name)
        ProcessServer.ExecuteTask(Of ProcesoComunes.DataValidarContadorEntidad)(AddressOf ProcesoComunes.ValidarContadorEntidad, datosValCont, services)
    End Sub

    <Task()> Public Shared Sub DatosIniciales(ByVal data As DataPrcAlbaranar, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataPrcAlbaranar)(AddressOf ProcesoAlbaranVenta.DatosIniciales, data, services)
        Dim Albaranes As DataTable = New AlbaranVentaCabecera().AddNew
        services.RegisterService(New ResultAlbaranAlquiler(Albaranes))
    End Sub

End Class
