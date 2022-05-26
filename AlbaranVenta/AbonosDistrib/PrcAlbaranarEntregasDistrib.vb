Public Class PrcAlbaranarEntregasDistrib
    Inherits Process(Of DataPrcCrearAlbaranVentaAbono, AlbaranLogProcess)

    '//Crea la secuencia de Tareas a realizar en el proceso de Albaranar los Pedidos
    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of DataPrcCrearAlbaranVentaAbono, AlbCabVenta())(AddressOf DatosIniciales)
        '//Bucle para recorrer todos los documentos de Albarán de Venta a generar
        Me.AddForEachTask(Of PrcCrearAlbaranVentaAbono)(OnExceptionBehaviour.NextLoop)
        Me.AddTask(Of Object, AlbaranLogProcess)(AddressOf ProcesoComunes.ResultadoAlbaran)
    End Sub


    <Task()> Public Shared Function DatosIniciales(ByVal data As DataPrcCrearAlbaranVentaAbono, ByVal services As ServiceProvider) As AlbCabVenta()

        Dim ProcInfo As New ProcessInfoAVDistrib
        If Not data.IDContador Is Nothing Then ProcInfo.IDContador = data.IDContador
        If Not data.FechaAlbaran Is Nothing Then ProcInfo.FechaAlbaran = data.FechaAlbaran
        services.RegisterService(ProcInfo, GetType(ProcessInfoAVDistrib))

        Dim IDAlbaranes(-1) As Object
        ReDim IDAlbaranes(data.IDAlbaranCliente.Length - 1)
        data.IDAlbaranCliente.CopyTo(IDAlbaranes, 0)

        Dim f As New Filter
        f.Add(New InListFilterItem("IDAlbaran", IDAlbaranes, FilterType.Numeric))
        Dim dtAlbOrigen As DataTable = New AlbaranVentaCabecera().Filter(f)
        If dtAlbOrigen.Rows.Count > 0 Then
            Dim AlbCab(-1) As AlbCabVentaAlbaran

            For Each dr As DataRow In dtAlbOrigen.Rows
                ReDim Preserve AlbCab(AlbCab.Length)
                AlbCab(AlbCab.Length - 1) = New AlbCabVentaAlbaran(dr)
            Next

            Return AlbCab
        Else
            ApplicationService.GenerateError("No hay datos para procesar.")
        End If
    End Function

End Class
