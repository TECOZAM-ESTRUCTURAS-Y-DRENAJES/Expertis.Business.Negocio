Public Class PrcCrearFacturaCompraEntregaCta
    Inherits Process(Of FraCabCompraEntregaCta)

    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of FraCabCompraEntregaCta, DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CrearDocumentoFactura)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarValoresPredeterminadosGenerales)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarProveedorGrupo)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarDatosProveedor)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarDireccion)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarBanco)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoCompra.AsignarObservacionesCompra)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoComunes.AsignarCentroGestion)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarContador)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarNumeroFactura)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoCompra.AsignarEjercicio)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarSuFactura)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoComunes.AsignarCambiosMoneda)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf AsignarCondicionesEccasEntegasCta)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf NegocioGeneral.AsignarEnvio347Doc)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf NegocioGeneral.AsignarEnvio349Doc)

        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf CrearLineasDesdeEntregas)

        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularImporteLineasFacturas)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf NegocioGeneral.CalcularImpuestos)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularAnaliticaFacturas)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularBasesImponibles)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularTotales)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularVencimientos)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf AsociarVencimientosEntregaCta)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarClaveOperacion)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf Business.General.Comunes.BeginTransaction)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf Business.General.Comunes.UpdateDocument)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoComunes.AddFacturaCreadaResultadoFacturacion)
        Me.AddTask(Of DocumentoFacturaCompra)(AddressOf ActualizarEntregaCuenta)
    End Sub

    <Task()> Public Shared Sub AsignarCondicionesEccasEntegasCta(ByVal Doc As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        Doc.HeaderRow("IDDiaPago") = System.DBNull.Value
        If Length(Doc.Cabecera.IDBancoPropio) Then
            Doc.HeaderRow("IDBancoPropio") = Doc.Cabecera.IDBancoPropio
        Else
            Doc.HeaderRow("IDBancoPropio") = System.DBNull.Value
        End If
    End Sub


    <Task()> Public Shared Sub CrearLineasDesdeEntregas(ByVal Doc As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        Dim lineas As DataTable = Doc.dtLineas
        Dim oFCL As New FacturaCompraLinea
        If lineas Is Nothing Then
            lineas = oFCL.AddNew
            Doc.Add(GetType(FacturaCompraLinea).Name, lineas)
        End If

        Dim fraCabEntCta As FraCabCompraEntregaCta = Doc.Cabecera
        Dim context As New BusinessData(Doc.HeaderRow)

        For Each fraLin As FraLinEntregaCta In fraCabEntCta.Lineas
            Dim linea As DataRow = lineas.NewRow
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoFacturacionCompra.AsignarValoresPredeterminadosLinea, linea, services)
            linea("IDFactura") = Doc.HeaderRow("IDFactura")
            oFCL.ApplyBusinessRule("IDCentroGestion", Doc.HeaderRow("IDCentroGestion"), linea, context)
            oFCL.ApplyBusinessRule("IDArticulo", fraLin.IDArticulo, linea, context)
            oFCL.ApplyBusinessRule("CContable", fraLin.CContable, linea, context)
            oFCL.ApplyBusinessRule("Cantidad", fraLin.Cantidad, linea, context)
            oFCL.ApplyBusinessRule("Dto1", 0, linea, context)
            oFCL.ApplyBusinessRule("Dto2", 0, linea, context)
            oFCL.ApplyBusinessRule("Dto3", 0, linea, context)
            oFCL.ApplyBusinessRule("Dto", 0, linea, context)
            oFCL.ApplyBusinessRule("DtoProntoPago", 0, linea, context)
            oFCL.ApplyBusinessRule("Precio", fraLin.Precio, linea, context)
            If Nz(fraCabEntCta.IDObra, 0) > 0 Then oFCL.ApplyBusinessRule("IDObra", fraCabEntCta.IDObra, linea, context)
            lineas.Rows.Add(linea)
        Next
    End Sub

    <Task()> Public Shared Sub AsociarVencimientosEntregaCta(ByVal Doc As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        If Not Doc.dtPagos Is Nothing AndAlso Doc.dtPagos.Rows.Count > 0 Then
            Doc.HeaderRow("VencimientosManuales") = True
            Dim FraCab As FraCabCompraEntregaCta = CType(Doc.Cabecera, FraCabCompraEntregaCta)
            For Each drVencimiento As DataRow In Doc.dtPagos.Rows
                drVencimiento("IDTipoPago") = FraCab.IDTipoPago
                drVencimiento("CContable") = FraCab.CContableProveedor
            Next
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarEntregaCuenta(ByVal Doc As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        Dim drRowEntrega As DataRow = New EntregasACuenta().GetItemRow(CType(Doc.Cabecera, FraCabCompraEntregaCta).IDEntrega)
        drRowEntrega("IDFactura") = Doc.HeaderRow("IDFactura")
        drRowEntrega("Generado") = True
        BusinessHelper.UpdateTable(drRowEntrega.Table)
    End Sub

    Protected Overrides Function OnException(ByVal exceptionArgs As Engine.BE.BusinessProcesses.ProcessExceptionArgs) As Engine.BE.BusinessProcesses.OnExceptionBehaviour
        ProcessServer.ExecuteTask(Of Boolean)(AddressOf Comunes.RollbackTransaction, True, New ServiceProvider())

        Dim resultFra As ResultFacturacion = exceptionArgs.Services.GetService(Of ResultFacturacion)()
        Dim log As LogProcess = resultFra.Log
        ReDim Preserve log.Errors(log.Errors.Length)
        If TypeOf exceptionArgs.TaskData Is DocumentoFacturaCompra Then
            Dim fra As FraCabCompra = CType(exceptionArgs.TaskData, DocumentoFacturaCompra).Cabecera
            If TypeOf fra Is FraCabCompraEntregaCta Then
                Dim FraEntCta As FraCabCompraEntregaCta = CType(fra, FraCabCompraEntregaCta)
                log.Errors(log.Errors.Length - 1) = New ClassErrors(FraEntCta.IDEntrega, exceptionArgs.Exception.Message)
            End If
        Else
            log.Errors(log.Errors.Length - 1) = New ClassErrors(String.Empty, exceptionArgs.Exception.Message)
        End If
        Return MyBase.OnException(exceptionArgs)
    End Function

End Class