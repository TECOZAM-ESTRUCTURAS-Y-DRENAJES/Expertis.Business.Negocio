'David Velasco 3/3/22
Public Class PrcAlbaranarAlquilerDeposito
    Inherits Process(Of PrcAlbaranarAlquilerDeposito, ResultAlbaranAlquiler)

    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of DataPrcAlbaranarDeposito)(AddressOf ValidacionesPrevias)
        Me.AddTask(Of DataPrcAlbaranarDeposito)(AddressOf DatosIniciales)
        Me.AddTask(Of DataPrcAlbaranarDeposito, AlbCabVentaAlquiler())(AddressOf ProcesoAlbaranVentaAlquiler.AgruparAlquiler2)
        'Bucle para recorrer todos los documentos a expedir a generar
        Me.AddForEachTask(Of PrcCrearAlbaranVentaAlquiler)(OnExceptionBehaviour.NextLoop)
        Me.AddTask(Of Object, ResultAlbaranAlquiler)(AddressOf ProcesoAlbaranVentaAlquiler.Resultado)
    End Sub

    <Task()> Public Shared Sub ValidacionesPrevias(ByVal data As DataPrcAlbaranarDeposito, ByVal services As ServiceProvider)
        Dim datosValCont As New ProcesoComunes.DataValidarContadorEntidad(data.IDContador, GetType(AlbaranVentaCabecera).Name)
        ProcessServer.ExecuteTask(Of ProcesoComunes.DataValidarContadorEntidad)(AddressOf ProcesoComunes.ValidarContadorEntidad, datosValCont, services)
    End Sub

    <Task()> Public Shared Sub DatosIniciales(ByVal data As DataPrcAlbaranarDeposito, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataPrcAlbaranarDeposito)(AddressOf ProcesoAlbaranVenta.DatosIniciales2, data, services)
        Dim Albaranes As DataTable = New AlbaranVentaCabecera().AddNew
        services.RegisterService(New ResultAlbaranAlquiler(Albaranes))
    End Sub

End Class
