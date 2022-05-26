Public Class PrcCrearFacturaVentaEntregaCta
    Inherits Process(Of FraCabEntregaCta)

    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of FraCabEntregaCta, DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CrearDocumentoFactura)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarValoresPredeterminadosGenerales)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarClienteGrupo)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarDatosCliente)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarDireccion)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarDatosFiscales)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoComercial.AsignarObservacionesComercial)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoComunes.AsignarCentroGestion)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarContador)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarNumeroFactura)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoComercial.AsignarEjercicio)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoComunes.AsignarCambiosMoneda)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf AsignarCondicionesEccasEntegasCta)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf NegocioGeneral.AsignarEnvio347Doc)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf NegocioGeneral.AsignarEnvio349Doc)

        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf CrearLineasDesdeEntregas)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularImporteLineasFacturas)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf NegocioGeneral.CalcularImpuestos)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularRepresentantes)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularAnalitica)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularBasesImponibles)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularTotales)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularPuntoVerde)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularVencimientos)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf AsociarVencimientosEntregaCta)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarClaveOperacion)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf Business.General.Comunes.BeginTransaction)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf Business.General.Comunes.UpdateDocument)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoComunes.AddFacturaCreadaResultadoFacturacion)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ActualizarEntregaCuenta)
    End Sub

    <Task()> Public Shared Sub AsignarCondicionesEccasEntegasCta(ByVal Doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        Doc.HeaderRow("IDDiaPago") = System.DBNull.Value
        If Length(Doc.Cabecera.IDBancoPropio) Then
            Doc.HeaderRow("IDBancoPropio") = Doc.Cabecera.IDBancoPropio
        Else
            Doc.HeaderRow("IDBancoPropio") = System.DBNull.Value
        End If
    End Sub

    <Task()> Public Shared Sub AsociarVencimientosEntregaCta(ByVal Doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If Not Doc.dtCobros Is Nothing AndAlso Doc.dtCobros.Rows.Count > 0 Then
            Doc.HeaderRow("VencimientosManuales") = True
            Dim FraCab As FraCabEntregaCta = CType(Doc.Cabecera, FraCabEntregaCta)
            For Each drVencimiento As DataRow In Doc.dtCobros.Rows
                drVencimiento("IDTipoCobro") = FraCab.IDTipoCobro
                drVencimiento("CContable") = FraCab.CContableCliente
            Next
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarEntregaCuenta(ByVal Doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        Dim drRowEntrega As DataRow = New EntregasACuenta().GetItemRow(CType(Doc.Cabecera, FraCabEntregaCta).IDEntrega)
        drRowEntrega("IDFactura") = Doc.HeaderRow("IDFactura")
        drRowEntrega("Generado") = True
        BusinessHelper.UpdateTable(drRowEntrega.Table)
    End Sub

    <Task()> Public Shared Sub CrearLineasDesdeEntregas(ByVal Doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        Dim lineas As DataTable = Doc.dtLineas
        Dim oFVL As New FacturaVentaLinea
        If lineas Is Nothing Then
            lineas = oFVL.AddNew
            Doc.Add(GetType(FacturaVentaLinea).Name, lineas)
        End If

        Dim fraCabEntCta As FraCabEntregaCta = Doc.Cabecera
        Dim context As New BusinessData(Doc.HeaderRow)

        For Each fraLin As FraLinEntregaCta In fraCabEntCta.Lineas
            Dim linea As DataRow = lineas.NewRow
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoFacturacionVenta.AsignarValoresPredeterminadosLinea, linea, services)
            linea("IDFactura") = Doc.HeaderRow("IDFactura")
            oFVL.ApplyBusinessRule("IDCentroGestion", Doc.HeaderRow("IDCentroGestion"), linea, context)
            oFVL.ApplyBusinessRule("IDArticulo", fraLin.IDArticulo, linea, context)
            oFVL.ApplyBusinessRule("CContable", fraLin.CContable, linea, context)
            oFVL.ApplyBusinessRule("Cantidad", fraLin.Cantidad, linea, context)
            oFVL.ApplyBusinessRule("Dto1", 0, linea, context)
            oFVL.ApplyBusinessRule("Dto2", 0, linea, context)
            oFVL.ApplyBusinessRule("Dto3", 0, linea, context)
            oFVL.ApplyBusinessRule("Dto", 0, linea, context)
            oFVL.ApplyBusinessRule("DtoProntoPago", 0, linea, context)
            oFVL.ApplyBusinessRule("Precio", fraLin.Precio, linea, context)
            If Nz(fraCabEntCta.IDObra, 0) > 0 Then oFVL.ApplyBusinessRule("IDObra", fraCabEntCta.IDObra, linea, context)

            lineas.Rows.Add(linea)
        Next
    End Sub

    Protected Overrides Function OnException(ByVal exceptionArgs As Engine.BE.BusinessProcesses.ProcessExceptionArgs) As Engine.BE.BusinessProcesses.OnExceptionBehaviour
        ProcessServer.ExecuteTask(Of Boolean)(AddressOf Comunes.RollbackTransaction, True, New ServiceProvider())

        Dim resultFra As ResultFacturacion = exceptionArgs.Services.GetService(Of ResultFacturacion)()
        Dim log As LogProcess = resultFra.Log
        ReDim Preserve log.Errors(log.Errors.Length)
        If TypeOf exceptionArgs.TaskData Is DocumentoFacturaVenta Then
            Dim fra As FraCab = CType(exceptionArgs.TaskData, DocumentoFacturaVenta).Cabecera
            If TypeOf fra Is FraCabEntregaCta Then
                Dim FraEntCta As FraCabEntregaCta = CType(fra, FraCabEntregaCta)
                log.Errors(log.Errors.Length - 1) = New ClassErrors(FraEntCta.IDEntrega, exceptionArgs.Exception.Message)
            End If
        Else
            log.Errors(log.Errors.Length - 1) = New ClassErrors(String.Empty, exceptionArgs.Exception.Message)
        End If
        Return MyBase.OnException(exceptionArgs)
    End Function

End Class