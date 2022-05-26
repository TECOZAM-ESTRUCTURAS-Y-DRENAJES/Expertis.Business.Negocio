Public Class PrcCrearFacturaVentaOT
    Inherits Process

    'Lista de tareas a ejecutar
    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of FraCabMnto, DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CrearDocumentoFactura)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarValoresPredeterminadosGenerales)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarDescuentoFactura)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarClienteGrupo)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarDatosCliente)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf AsignarDiaPago)

        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarDireccion)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarDatosFiscales)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarBanco)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoComercial.AsignarObservacionesComercial)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoComunes.AsignarCentroGestion)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarContador)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarNumeroFacturaPropuesta)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoComercial.AsignarEjercicio)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoComunes.AsignarCambiosMoneda)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf NegocioGeneral.AsignarEnvio347Doc)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf NegocioGeneral.AsignarEnvio349Doc)

        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf CrearLineasDesdeOT)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularImporteLineasFacturas)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf NegocioGeneral.CalcularImpuestos)

        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularBasesImponibles)
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularTotales)
        
        Me.AddTask(Of DocumentoFacturaVenta)(AddressOf AñadirAResultado)
    End Sub

    <Task()> Public Shared Sub AsignarDiaPago(ByVal Doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If Not Doc.Cabecera Is Nothing AndAlso Length(CType(Doc.Cabecera, FraCabMnto).IDDiaPago) > 0 Then
            Doc.HeaderRow("IDDiaPago") = CType(Doc.Cabecera, FraCabMnto).IDDiaPago
        End If
    End Sub

    <Task()> Public Shared Sub CrearLineasDesdeOT(ByVal Doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        Dim fra As FraCabMnto = Doc.Cabecera

        Dim ids(fra.Lineas.Length - 1) As Object
        For i As Integer = 0 To ids.Length - 1
            ids(i) = fra.Lineas(i).IDMntoOTControl
        Next

        Dim f As New Filter
        f.Add(New InListFilterItem("IDMntoOTControl", ids, FilterType.Numeric))

        Dim oFVL As New FacturaVentaLinea
        Dim dtControlOT As DataTable = ProcessServer.ExecuteTask(Of Filter, DataTable)(AddressOf ProcesoFacturacionVenta.GetDatosFacturacionOT, f, services)
        Dim lineas As DataTable = Doc.dtLineas
        If lineas Is Nothing Then
            lineas = oFVL.AddNew
            Doc.Add(GetType(FacturaVentaLinea).Name, lineas)
        End If
        Dim context As New BusinessData(Doc.HeaderRow)
        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
        Dim fraCabAlb As FraCabMnto = Doc.Cabecera
        For Each control As DataRow In dtControlOT.Rows
            Dim fralin As FraLinMnto = Nothing
            For i As Integer = 0 To fraCabAlb.Lineas.Length - 1
                If control("IDMntoOTControl") = fraCabAlb.Lineas(i).IDMntoOTControl Then
                    fralin = fraCabAlb.Lineas(i)
                    Exit For
                End If
            Next
            If Not fralin Is Nothing Then
                If fralin.QaFacturar <> 0 Then
                    Dim linea As DataRow = lineas.NewRow

                    ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoFacturacionVenta.AsignarValoresPredeterminadosLinea, linea, services)

                    linea("IDFactura") = Doc.HeaderRow("IDFactura")
                    linea = oFVL.ApplyBusinessRule("IDArticulo", control("IDArticulo"), linea, context)
                    linea("DescArticulo") = control("DescArticulo")
                    linea("lote") = control("lote")
                    linea("IDCentroGestion") = Doc.HeaderRow("IDCentroGestion")
                    linea("Cantidad") = fralin.QaFacturar
                    linea("QInterna") = fralin.QaFacturar * linea("Factor")
                    If Length(control("IDUDMedida")) > 0 Then linea("IDUDMedida") = control("IDUDMedida")
                    If Length(control("UdValoracion")) > 0 Then linea("UdValoracion") = control("UdValoracion")

                    linea("Precio") = Nz(control("PrecioVentaA"), 0)
                    If Doc.HeaderRow("IDMoneda") <> Monedas.MonedaA.ID Then
                        Dim datCamMon As New DataCambioMoneda(New DataRowPropertyAccessor(linea), Monedas.MonedaA.ID, Doc.HeaderRow("IDMoneda"), Doc.HeaderRow("FechaFactura"))
                        ProcessServer.ExecuteTask(Of DataCambioMoneda)(AddressOf NegocioGeneral.CambioMoneda, datCamMon, services)
                    End If

                    linea("Dto1") = control("Dto1")
                    linea("Dto2") = control("Dto2")
                    linea("Dto3") = control("Dto3")

                    linea("IDOT") = control("IDOT")
                    linea("IDMntoOTControl") = control("IDMntoOTControl")
                    lineas.Rows.Add(linea)

                End If
            End If
        Next
    End Sub

    <Task()> Public Shared Sub AñadirAResultado(ByVal Doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        Dim Facturas As DataTable = services.GetService(Of ResultFacturacion)().PropuestaFacturas
        Facturas.Rows.Add(Doc.HeaderRow.ItemArray)

        Dim arDocFras As ArrayList = services.GetService(Of ArrayList)()

        arDocFras.Add(Doc)
    End Sub

    Protected Overrides Function OnException(ByVal exceptionArgs As Engine.BE.BusinessProcesses.ProcessExceptionArgs) As Engine.BE.BusinessProcesses.OnExceptionBehaviour
        ProcessServer.ExecuteTask(Of Boolean)(AddressOf Comunes.RollbackTransaction, True, New ServiceProvider())

        Dim fvr As ResultFacturacion = exceptionArgs.Services.GetService(Of ResultFacturacion)()
        Dim log As LogProcess = fvr.Log
        ReDim Preserve log.Errors(log.Errors.Length)
        If TypeOf exceptionArgs.TaskData Is DocumentoFacturaVenta Then
            Dim fra As FraCab = CType(exceptionArgs.TaskData, DocumentoFacturaVenta).Cabecera
            If TypeOf fra Is FraCabMnto Then
                Dim FraAlb As FraCabMnto = CType(fra, FraCabMnto)
                log.Errors(log.Errors.Length - 1) = New ClassErrors(" Cliente: " & FraAlb.IDCliente, exceptionArgs.Exception.Message)
            End If
        Else : log.Errors(log.Errors.Length - 1) = New ClassErrors(String.Empty, exceptionArgs.Exception.Message)
        End If

        Return MyBase.OnException(exceptionArgs)
    End Function

End Class
