Public Class PrcFacturacionOTs
    Inherits Process(Of DataPrcFacturacionOTs, ResultFacturacion)

    'Lista de tareas a ejecutar
    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of DataPrcFacturacionOTs)(AddressOf ValidacionesPrevias)
        Me.AddTask(Of DataPrcFacturacionOTs)(AddressOf DatosIniciales)
        Me.AddTask(Of DataPrcFacturacionOTs, FraCabMnto())(AddressOf ProcesoFacturacionVenta.AgruparOTs)
        'Bucle para recorrer todos los documentos a factura a generar
        Me.AddForEachTask(Of PrcCrearFacturaVentaOT)(OnExceptionBehaviour.NextLoop)
        Me.AddTask(Of Object, ResultFacturacion)(AddressOf ProcesoFacturacionVenta.Resultado)
    End Sub

    <Task()> Public Shared Sub ValidacionesPrevias(ByVal data As DataPrcFacturacionOTs, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataPrcFacturacionOTs)(AddressOf ValidarExistenDatos, data, services)
        ProcessServer.ExecuteTask(Of DataPrcFacturacionOTs)(AddressOf ValidarContadorObligatorio, data, services)
        ProcessServer.ExecuteTask(Of DataPrcFacturacionOTs)(AddressOf ValidarContadorDeEntidad, data, services)
    End Sub

    <Task()> Public Shared Sub ValidarExistenDatos(ByVal data As DataPrcFacturacionOTs, ByVal services As ServiceProvider)
        If data.IDMntoOTControl Is Nothing OrElse data.IDMntoOTControl.Length = 0 Then
            ApplicationService.GenerateError("No exiten datos a Facturar.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarContadorObligatorio(ByVal data As DataPrcFacturacionOTs, ByVal services As ServiceProvider)
        If Length(data.IDContador) = 0 Then
            Dim f As New Filter
            f.Add(New StringFilterItem("Entidad", GetType(FacturaVentaCabecera).Name))
            f.Add(New BooleanFilterItem("Predeterminado", True))
            Dim dtCont As DataTable = New EntidadContador().Filter(f)
            If Not dtCont Is Nothing AndAlso dtCont.Rows.Count > 0 Then
                data.IDContador = dtCont.Rows(0)("IDContador")
            Else
                ApplicationService.GenerateError("Debe indicar un Contador para la entidad {0}.", GetType(FacturaVentaCabecera).Name)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarContadorDeEntidad(ByVal data As DataPrcFacturacionOTs, ByVal services As ServiceProvider)
        If Length(data.IDContador) > 0 Then
            Dim datosValCont As New ProcesoComunes.DataValidarContadorEntidad(data.IDContador, GetType(FacturaVentaCabecera).Name)
            ProcessServer.ExecuteTask(Of ProcesoComunes.DataValidarContadorEntidad)(AddressOf ProcesoComunes.ValidarContadorEntidad, datosValCont, services)
        End If
    End Sub

    <Task()> Public Shared Sub DatosIniciales(ByVal data As DataPrcFacturacionOTs, ByVal services As ServiceProvider)
        '//Prepara en el service información del proceso.
        services.RegisterService(New ProcessInfoFra(data.IDContador, String.Empty))

        '//Prepara en el service lo que va a retornar el proceso, para tenerlo disponible en todo el proceso.
        services.RegisterService(New LogProcess)

        Dim Facturas As DataTable = New FacturaVentaCabecera().AddNew
        services.RegisterService(New ResultFacturacion(Facturas))
    End Sub
End Class
