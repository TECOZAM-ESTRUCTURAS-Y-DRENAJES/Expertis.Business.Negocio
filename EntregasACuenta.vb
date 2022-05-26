Public Class EntregasACuenta

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbEntregasACuenta"

#End Region

#Region "Eventos RegisterAddNewTasks"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarIdentidicador, data, services)
    End Sub

#End Region

#Region "Eventos RegisterDeleteTasks"

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf ComprobarGenSald)
        'deleteProcess.AddTask(Of DataRow)(AddressOf BorrarCobrosPagos)
    End Sub

    <Task()> Public Shared Sub ComprobarGenSald(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data("Generado") OrElse data("Saldado") Then ApplicationService.GenerateError("Una Entrega Generada o Saldada no puede eliminarse.")
    End Sub

    '<Task()> Public Shared Sub BorrarCobrosPagos(ByVal data As DataRow, ByVal services As ServiceProvider)
    '    If Not data("Generado") AndAlso Not data("Saldado") Then
    '        Dim objNegCobroPago As BusinessHelper
    '        If Length(data("IDCliente")) > 0 Then
    '            objNegCobroPago = BusinessHelper.CreateBusinessObject("Cobro")
    '        ElseIf Length(data("IDProveedor")) > 0 Then
    '            objNegCobroPago = BusinessHelper.CreateBusinessObject("Pago")
    '        End If
    '        Dim dtCobrosPagos As DataTable = objNegCobroPago.Filter(New FilterItem("IDEntrega", FilterOperator.Equal, data("IDEntrega")))
    '        For Each drCP As DataRow In dtCobrosPagos.Rows
    '            drCP.Delete()
    '        Next
    '        BusinessHelper.UpdateTable(dtCobrosPagos)
    '    End If
    'End Sub

    <Task()> Public Shared Sub EliminarEntregasRetencionSinFactura(ByVal data As DataTable, ByVal services As ServiceProvider)
        '//Eliminamos las Entregas de TipoRetención que no tengan vinculada ninguna factura.
        For Each drRow As DataRow In data.Select
            If drRow("TipoEntrega") = enumTipoEntrega.Retencion AndAlso Not CBool(drRow("Generado")) AndAlso Not CBool(drRow("Saldado")) Then
                drRow.Delete()
            End If
        Next
        BusinessHelper.UpdateTable(data)
    End Sub

#End Region

#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("IDCliente", AddressOf CambioCliente)
        oBrl.Add("IDProveedor", AddressOf CambioProveedor)
        oBrl.Add("IDArticulo", AddressOf CambioArticulo)
        oBrl.Add("TipoEntrega", AddressOf CambioTipoEntrega)
        oBrl.Add("IDTipoCobroPago", AddressOf CambioTipoCobroPago)
        oBrl.Add("CCArticulo", AddressOf CambioCContables)
        oBrl.Add("CCClienteProveedor", AddressOf CambioCContables)
        oBrl.Add("IDObra", AddressOf CambioObra)
        oBrl.Add("FechaEntrega", AddressOf CambioFechaEntrega)
        oBrl.Add("GenerarFactura", AddressOf CambioGenerarFactura)
        Return oBrl
    End Function

    <Task()> Public Shared Sub CambioCliente(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Value) > 0 Then
            ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf TipoCobroPagoEntrega, data.Current, services)
            Dim drRowCliente As DataRow = New Cliente().GetItemRow(data.Value)
            data.Current("IDBancoPropio") = drRowCliente("IDBancoPropio")
            data.Current("IDFormaPago") = drRowCliente("IDFormaPago")
            data.Current("IDMoneda") = drRowCliente("IDMoneda")
            data.Current("Titulo") = drRowCliente("DescCliente")
            If Length(data.Current("TipoEntrega")) > 0 AndAlso Length(data.Current("IDArticulo")) > 0 AndAlso _
               (Length(data.Current("IDCliente")) > 0 OrElse Length(data.Current("IDProveedor")) > 0) Then
                ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf EntregasACuenta.CuentasContablesEntrega, data.Current, services)
            End If
        Else
            If Length(data.Current("IDProveedor")) = 0 Then ApplicationService.GenerateError("Introduzca el Cliente o el Proveedor de la Entrega.")
        End If
    End Sub

    <Task()> Public Shared Sub CambioProveedor(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Value) > 0 Then
            ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf TipoCobroPagoEntrega, data.Current, services)
            Dim drRowProveedor As DataRow = New Proveedor().GetItemRow(data.Value)
            data.Current("IDBancoPropio") = drRowProveedor("IDBancoPropio")
            data.Current("IDFormaPago") = drRowProveedor("IDFormaPago")
            data.Current("IDMoneda") = drRowProveedor("IDMoneda")
            data.Current("Titulo") = drRowProveedor("DescProveedor")
            If Length(data.Current("TipoEntrega")) > 0 AndAlso Length(data.Current("IDArticulo")) > 0 AndAlso _
               (Length(data.Current("IDCliente")) > 0 OrElse Length(data.Current("IDProveedor")) > 0) Then
                ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf EntregasACuenta.CuentasContablesEntrega, data.Current, services)
            End If
        Else
            If Length(data.Current("IDProveedor")) = 0 Then ApplicationService.GenerateError("Introduzca el Cliente o el Proveedor de la Entrega.")
        End If
    End Sub

    <Task()> Public Shared Sub CambioArticulo(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Value) > 0 Then
            If Length(data.Current("TipoEntrega")) > 0 AndAlso Length(data.Current("IDArticulo")) > 0 AndAlso _
               (Length(data.Current("IDCliente")) > 0 OrElse Length(data.Current("IDProveedor")) > 0) Then
                ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf EntregasACuenta.CuentasContablesEntrega, data.Current, services)
            End If
        Else : ApplicationService.GenerateError("Introduzca el Artículo.")
        End If
    End Sub

    <Task()> Public Shared Sub CambioTipoEntrega(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Value) > 0 AndAlso IsNumeric(data.Value) Then
            '//Establecemos el Tipo de Cobro, en función del Tipo de Entrega.
            ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf TipoCobroPagoEntrega, data.Current, services)
            If Length(data.Current("TipoEntrega")) > 0 AndAlso Length(data.Current("IDArticulo")) > 0 AndAlso _
                 (Length(data.Current("IDCliente")) > 0 OrElse Length(data.Current("IDProveedor")) > 0) Then
                ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf EntregasACuenta.CuentasContablesEntrega, data.Current, services)
            End If
        Else
            ApplicationService.GenerateError("Introduzca el Tipo de Entrega.")
        End If
    End Sub

    <Task()> Public Shared Sub CambioTipoCobroPago(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Not IsNothing(data.Value) AndAlso Nz(data.Value, -1) <> -1 Then
            Dim objNegTipoCobroPago As BusinessHelper
            If data.Current.ContainsKey("IDCliente") AndAlso Length(data.Current("IDCliente")) > 0 Then
                objNegTipoCobroPago = New TipoCobro
            ElseIf data.Current.ContainsKey("IDProveedor") AndAlso Length(data.Current("IDProveedor")) > 0 Then
                objNegTipoCobroPago = New TipoPago
            End If
            Dim drRowTipoCobroPago As DataRow = objNegTipoCobroPago.GetItemRow(data.Value)
            If Not IsNothing(drRowTipoCobroPago) AndAlso data.Current.ContainsKey("DescTipoCobroPago") Then
                data.Current("DescTipoCobroPago") = drRowTipoCobroPago("DescTipo")
            End If
        Else : ApplicationService.GenerateError("Introduzca el Tipo de Cobro/Pago.")
        End If
    End Sub

    <Task()> Public Shared Sub CambioCContables(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Value) > 0 AndAlso Length(data.Current("FechaEntrega") & String.Empty) > 0 Then
            Dim strEjercicio As String = ProcessServer.ExecuteTask(Of Date, String)(AddressOf NegocioGeneral.EjercicioPredeterminado, data.Current("FechaEntrega"), services)
            Dim objNegPlanContable As BusinessHelper = BusinessHelper.CreateBusinessObject("PlanContable")
            Dim drRowPlanContable As DataRow = objNegPlanContable.GetItemRow(strEjercicio, data.Value)
        End If
    End Sub

    <Task()> Public Shared Sub CambioObra(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Not IsNothing(data.Value) AndAlso data.Value <> 0 Then
            Dim objNegObra As New Object
            objNegObra = BusinessHelper.CreateBusinessObject(AdminData.GetEntityInfo("ObraCabecera"))
            Dim drRowObra As DataRow = objNegObra.GetItemRow(data.Value)
            If Not IsNothing(drRowObra) Then
                If data.Current.ContainsKey("NObra") Then data.Current("NObra") = drRowObra("NObra")
                If data.Current.ContainsKey("DescObra") Then data.Current("DescObra") = drRowObra("DescObra")
            End If
            objNegObra = Nothing
        End If
    End Sub

    <Task()> Public Shared Sub CambioFechaEntrega(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Value) = 0 Then ApplicationService.GenerateError("Introduzca la Fecha de la Entrega.")
    End Sub

    <Task()> Public Shared Sub CambioGenerarFactura(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If data.Value = False AndAlso data.Current.ContainsKey("TipoEntrega") AndAlso data.Current("TipoEntrega") = enumTipoEntrega.Retencion Then
            ApplicationService.GenerateError("La Entrega es de tipo Retención. Este tipo de Entrega debe Generar Factura.")
        Else : data.Current(data.ColumnName) = data.Value
        End If
    End Sub

    <Task()> Public Shared Sub CuentasContablesEntrega(ByVal current As IPropertyAccessor, ByVal services As ServiceProvider)
        If (Length(current("IDCliente")) > 0 OrElse Length(current("IDProveedor")) > 0) AndAlso Length(current("IDArticulo")) > 0 Then
            Dim dtClienteProveedor As DataTable
            Dim strCampoCCClienteProveedor As String
            Dim intCircuito As Circuito
            '//Recogemos la información del Cliente/Proveedor
            Dim ctx As New BusinessData

            If Length(current("IDCliente")) > 0 Then
                Dim objNegCliente As New Cliente
                dtClienteProveedor = objNegCliente.SelOnPrimaryKey(current("IDCliente"))
                objNegCliente = Nothing
                strCampoCCClienteProveedor = "CCCliente"
                intCircuito = Circuito.Ventas
                ctx("IDCliente") = current("IDCliente")
            ElseIf Length(current("IDProveedor")) > 0 Then
                Dim objNegProveedor As New Proveedor
                dtClienteProveedor = objNegProveedor.SelOnPrimaryKey(current("IDProveedor"))
                objNegProveedor = Nothing
                strCampoCCClienteProveedor = "CCProveedor"
                intCircuito = Circuito.Compras
                ctx("IDProveedor") = current("IDProveedor")
            End If

            '//Establecemos las cuentas en función del Tipo de Entrega
            If Not IsNothing(dtClienteProveedor) AndAlso dtClienteProveedor.Rows.Count > 0 Then
                Select Case current("TipoEntrega")
                    Case enumTipoEntrega.Anticipo
                        If Length(dtClienteProveedor.Rows(0)("CCAnticipo")) > 0 Then
                            current("CCArticulo") = dtClienteProveedor.Rows(0)("CCAnticipo")
                            current("CCClienteProveedor") = dtClienteProveedor.Rows(0)(strCampoCCClienteProveedor)
                        Else
                            current("CCClienteProveedor") = dtClienteProveedor.Rows(0)(strCampoCCClienteProveedor)
                            If intCircuito = Circuito.Compras Then
                                Dim brd As New BusinessRuleData("IDArticulo", current("IDArticulo"), current, ctx)
                                ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoCompra.CambioArticulo, brd, services)
                                current("CCArticulo") = current("CContable")
                            Else
                                Dim brd As New BusinessRuleData("IDArticulo", current("IDArticulo"), current, ctx)
                                ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoComercial.CambioArticulo, brd, services)
                                current("CCArticulo") = current("CContable")
                            End If
                        End If
                    Case enumTipoEntrega.Fianza
                        If Length(dtClienteProveedor.Rows(0)("CCFianza")) > 0 Then
                            current("CCArticulo") = dtClienteProveedor.Rows(0)("CCFianza")
                            current("CCClienteProveedor") = dtClienteProveedor.Rows(0)(strCampoCCClienteProveedor)
                        Else
                            current("CCClienteProveedor") = dtClienteProveedor.Rows(0)(strCampoCCClienteProveedor)
                            If intCircuito = Circuito.Compras Then
                                Dim brd As New BusinessRuleData("IDArticulo", current("IDArticulo"), current, ctx)
                                ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoCompra.CambioArticulo, brd, services)
                                current("CCArticulo") = current("CContable")
                            Else
                                Dim brd As New BusinessRuleData("IDArticulo", current("IDArticulo"), current, ctx)
                                ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoComercial.CambioArticulo, brd, services)
                                current("CCArticulo") = current("CContable")
                            End If
                        End If
                    Case enumTipoEntrega.Retencion
                        current("CCClienteProveedor") = dtClienteProveedor.Rows(0)(strCampoCCClienteProveedor)
                        If Length(dtClienteProveedor.Rows(0)("CCRetencion")) > 0 Then
                            current("CCArticulo") = dtClienteProveedor.Rows(0)("CCRetencion")
                        Else
                            If intCircuito = Circuito.Compras Then
                                'Dim objCompra As New Compra
                                'current = objCompra.DetailBusinessRules("IDArticulo", current("IDArticulo"), current, services, ctx)
                                'TODO: Revisar esto, cambiará cuando tengamos las B.Rules
                                Dim brd As New BusinessRuleData("IDArticulo", current("IDArticulo"), current, ctx)
                                ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoCompra.CambioArticulo, brd, services)
                                current("CCArticulo") = current("CContable")
                            Else
                                Dim brd As New BusinessRuleData("IDArticulo", current("IDArticulo"), current, ctx)
                                ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoComercial.CambioArticulo, brd, services)
                                current("CCArticulo") = current("CContable")
                            End If
                        End If
                End Select
            End If
        End If
    End Sub

    <Task()> Public Shared Sub TipoCobroPagoEntrega(ByVal current As IPropertyAccessor, ByVal services As ServiceProvider)
        If (Length(current("IDCliente")) > 0 OrElse Length(current("IDProveedor")) > 0) AndAlso Length(current("TipoEntrega")) > 0 Then
            Dim p As New Parametro
            Select Case current("TipoEntrega")
                Case enumTipoEntrega.Anticipo
                    If (Length(current("IDCliente")) > 0) Then
                        current("IDTipoCobroPago") = p.TipoCobroAnticipo()
                    Else
                        current("IDTipoCobroPago") = p.TipoPagoAnticipo()
                    End If
                Case enumTipoEntrega.Fianza
                    If (Length(current("IDCliente")) > 0) Then
                        current("IDTipoCobroPago") = p.TipoCobroFianza()
                    Else
                        current("IDTipoCobroPago") = p.TipoPagoFianza()
                    End If
                Case enumTipoEntrega.Retencion
                    If (Length(current("IDCliente")) > 0) Then
                        current("IDTipoCobroPago") = p.TipoCobroRetencion()
                    Else
                        current("IDTipoCobroPago") = p.TipoPagoRetencion()
                    End If
            End Select

            If current.ContainsKey("DescTipoCobroPago") Then
                Dim objNegTPC As BusinessHelper
                If (Length(current("IDCliente")) > 0) Then
                    objNegTPC = BusinessHelper.CreateBusinessObject(AdminData.GetEntityInfo("TipoCobro"))
                Else
                    objNegTPC = BusinessHelper.CreateBusinessObject(AdminData.GetEntityInfo("TipoPago"))
                End If
                Dim drRowTipoCobroPago As DataRow = objNegTPC.GetItemRow(current("IDTipoCobroPago"))
                objNegTPC = Nothing
                If Not IsNothing(drRowTipoCobroPago) Then
                    current("DescTipoCobroPago") = drRowTipoCobroPago("DescTipo")
                End If
            End If

            p = Nothing
        End If
    End Sub

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("TipoEntrega")) = 0 Then ApplicationService.GenerateError("Introduzca el Tipo de Entrega.")
        If Length(data("IDCliente")) = 0 AndAlso Length(data("IDProveedor")) = 0 Then ApplicationService.GenerateError("Introduzca el Cliente o el Proveedor de la Entrega.")
        If Length(data("IDCliente")) > 0 AndAlso Length(data("IDProveedor")) > 0 Then ApplicationService.GenerateError("Introduzca el Cliente o el Proveedor de la Entrega.")
        If Length(data("FechaEntrega")) = 0 Then ApplicationService.GenerateError("Introduzca la Fecha de la Entrega.")
        If Length(data("IDTipoCobroPago")) = 0 Then ApplicationService.GenerateError("Debe introducir el Tipo de Cobro/Pago.")
        If data("TipoEntrega") = enumTipoEntrega.Fianza AndAlso data("Importe") > 0 Then ApplicationService.GenerateError("El importe, en las Entregas de Tipo Fianza, debe de ser negativo.")
        If Length(data("IDCliente")) > 0 Then
            Dim DtClie As DataTable = New Cliente().SelOnPrimaryKey(data("IDCliente"))
            If DtClie Is Nothing OrElse DtClie.Rows.Count = 0 Then ApplicationService.GenerateError("El Cliente introducido no existe.")
        End If
        If Length(data("IDProveedor")) > 0 Then
            Dim DtProv As DataTable = New Proveedor().SelOnPrimaryKey(data("IDProveedor"))
            If IsNothing(DtProv) OrElse DtProv.Rows.Count = 0 Then ApplicationService.GenerateError("El Proveedor introducido no existe.")
        End If
        If Length(data("IDArticulo")) = 0 Then
            ApplicationService.GenerateError("Introduzca el Artículo.")
        Else
            Dim DtArt As DataTable = New Articulo().SelOnPrimaryKey(data("IDArticulo"))
            If DtArt Is Nothing OrElse DtArt.Rows.Count = 0 Then ApplicationService.GenerateError("El Artículo introducido no existe.")
        End If
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarIdentidicador)
        updateProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.UpdateEntityRow)
        updateProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.MarcarComoActualizado)
        updateProcess.AddTask(Of DataRow)(AddressOf ComprobarRetencion)
    End Sub

    <Task()> Public Shared Sub AsignarIdentidicador(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added AndAlso Length(data("IDEntrega")) = 0 Then data("IDEntrega") = AdminData.GetAutoNumeric
    End Sub

    <Task()> Public Shared Sub ComprobarRetencion(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data("TipoEntrega") = enumTipoEntrega.Retencion AndAlso Not CBool(data("Generado")) AndAlso Not CBool(data("Saldado")) Then
            Dim ClsEnt As New EntregasACuenta
            ClsEnt.Delete(data)
        End If
    End Sub

#End Region

#Region " Eliminar Entregas a Cuenta "

    <Task()> Public Shared Function EliminarCobro(ByVal IDCobro As Integer, ByVal services As ServiceProvider) As Boolean
        Dim c As New Cobro
        Dim dtCobro As DataTable = c.SelOnPrimaryKey(IDCobro)
        If dtCobro.Rows(0)("Contabilizado") <> enumContabilizado.NoContabilizado OrElse dtCobro.Rows(0)("Liquidado") <> enumContabilizado.NoContabilizado OrElse dtCobro.Rows(0)("RecibidoEfecto") <> enumContabilizado.NoContabilizado Then
            ApplicationService.GenerateError("El Cobro está gestionado, no se puede eliminar.")
        Else
            c.Delete(dtCobro)
        End If
    End Function

    <Task()> Public Shared Function EliminarPago(ByVal IDPago As Integer, ByVal services As ServiceProvider) As Boolean
        Dim p As New Pago
        Dim dtPago As DataTable = p.SelOnPrimaryKey(IDPago)
        If dtPago.Rows.Count > 0 Then
            If dtPago.Rows(0)("Contabilizado") <> enumContabilizado.NoContabilizado OrElse dtPago.Rows(0)("GeneradoAsientoRemesa") <> enumContabilizado.NoContabilizado OrElse dtPago.Rows(0)("GeneradoAsientoTalon") <> enumContabilizado.NoContabilizado Then
                ApplicationService.GenerateError("El Pago está gestionado, no se puede eliminar.")
            Else
                p.Delete(dtPago)
            End If
        End If
    End Function


    <Serializable()> _
    Public Class DatosElimRestricEntFn
        Public IDEntrega As Integer
        Public IDFactura As Integer
        Public Circuito As Circuito

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDEntrega As Integer, ByVal IDFactura As Integer, ByVal Circuito As Circuito)
            Me.IDEntrega = IDEntrega
            Me.IDFactura = IDFactura
            Me.Circuito = Circuito
        End Sub
    End Class

    <Serializable()> _
    Public Class DatosElimRestricEnt
        Public IDFactura As Integer
        Public Circuito As Circuito
    End Class

    <Serializable()> _
    Public Class DatosElimRestricEntCobro
        Public IDEntrega As Integer
        Public IDCobroPago As Integer
    End Class

    <Serializable()> _
    Public Class DatosNuevaEntrega
        Public DtCab As DataTable
        Public DtLineas As DataTable
        Public DtEntregas As DataTable
        Public Circuito As Integer

        Public Sub New()
        End Sub

        Public Sub New(ByVal DtCab As DataTable, ByVal dtLineas As DataTable, ByVal dtEntregas As DataTable, ByVal Circuito As Circuito)
            Me.DtCab = DtCab
            Me.DtLineas = dtLineas
            Me.DtEntregas = dtEntregas
            Me.Circuito = Circuito
        End Sub
    End Class

    <Serializable()> _
    Public Class DatosGenFactEntrega
        Public DtEntregas As DataTable
        Public Circuito As Circuito
    End Class

    <Serializable()> _
    Public Class DatosNuevaLinFact
        Public HeaderRow As DataRow
        Public LineasFra As DataTable
        Public Entrega As DataTable
        Public Circuito As Circuito
        Public FacturaEntrega As Boolean

        Public Sub New(ByVal HeaderRow As DataRow, ByVal Entrega As DataTable, ByVal Circuito As Circuito, ByVal FacturaEntrega As Boolean, Optional ByVal LineasFra As DataTable = Nothing)
            Me.HeaderRow = HeaderRow
            Me.Entrega = Entrega
            Me.Circuito = Circuito
            Me.FacturaEntrega = FacturaEntrega
            If Not LineasFra Is Nothing Then Me.LineasFra = LineasFra
        End Sub
    End Class

    <Task()> Public Shared Function EliminarRestriccionesDeleteEntregaCuentaFn(ByVal data As DatosElimRestricEntFn, ByVal services As ServiceProvider) As Boolean
        Dim objFilter As New Filter
        Dim blnEliminarEntregaRetencion As Boolean

        '//Comprobamos si la línea a borrar es de una factura de ENTREGA SALDADA y/o GENERADA.
        objFilter.Clear()
        objFilter.Add(New NumberFilterItem("IDEntrega", data.IDEntrega))
        Dim dtEC As DataTable = New EntregasACuenta().Filter(objFilter)
        For Each drRowEC As DataRow In dtEC.Select
            If drRowEC("TipoEntrega") <> enumTipoEntrega.Retencion Then
                '//Entrega de Tipo FIANZA o ANTICIPO
                If Length(drRowEC("IDFacturaDestino")) > 0 AndAlso drRowEC("IDFacturaDestino") = data.IDFactura Then
                    drRowEC("IDFacturaDestino") = System.DBNull.Value
                    drRowEC("Saldado") = False
                End If

                If (Length(drRowEC("IDFactura")) > 0 AndAlso drRowEC("IDFactura") = data.IDFactura) Then
                    If Not CBool(drRowEC("Saldado")) AndAlso Length(drRowEC("IDFacturaDestino")) = 0 Then
                        If drRowEC("GenerarFactura") Then
                            drRowEC("IDFactura") = System.DBNull.Value
                            drRowEC("Generado") = False
                        End If
                    Else
                        If drRowEC("GenerarFactura") Then
                            ApplicationService.GenerateError("La Entrega está saldada, no se puede eliminar la Factura.")
                        Else
                            ApplicationService.GenerateError("La Entrega está saldada, no se puede eliminar el Cobro/Pago.")
                        End If
                    End If
                End If
            Else  '//Entrega de Tipo RETENCION
                If Length(drRowEC("IDFactura")) > 0 AndAlso drRowEC("IDFactura") = data.IDFactura Then
                    drRowEC("IDFactura") = System.DBNull.Value
                    drRowEC("Generado") = False
                End If

                If (Length(drRowEC("IDFacturaDestino")) > 0 AndAlso drRowEC("IDFacturaDestino") = data.IDFactura) Then
                    If Not CBool(drRowEC("Generado")) AndAlso Length(drRowEC("IDFactura")) = 0 Then
                        If drRowEC("GenerarFactura") Then
                            drRowEC("IDFacturaDestino") = System.DBNull.Value
                            drRowEC("Saldado") = False
                            blnEliminarEntregaRetencion = True
                        End If
                    Else
                        If drRowEC("GenerarFactura") Then
                            ApplicationService.GenerateError("Se ha realizado la factura de la retención. No se puede eliminar la Factura.")
                        End If
                    End If
                Else
                    If Length(drRowEC("IDFacturaDestino")) = 0 Then blnEliminarEntregaRetencion = True
                End If

                '//En este caso, desvinculamos los cobros/pagos de la factura de venta/compra de la entrega.
                objFilter.Clear()
                objFilter.Add(New NumberFilterItem("IDFactura", data.IDFactura))
                Select Case data.Circuito
                    Case Circuito.Ventas
                        Dim objNegCobro As New Cobro
                        Dim dtCobroFactura As DataTable = objNegCobro.Filter(objFilter)
                        For Each drCobroFactura As DataRow In dtCobroFactura.Select
                            drCobroFactura("IDEntrega") = System.DBNull.Value
                        Next
                        objNegCobro.Update(dtCobroFactura)
                        objNegCobro = Nothing
                    Case Circuito.Compras
                        Dim objNegPago As New Pago
                        Dim dtPagoFactura As DataTable = objNegPago.Filter(objFilter)
                        For Each drCobroFactura As DataRow In dtPagoFactura.Select
                            drCobroFactura("IDEntrega") = System.DBNull.Value
                        Next
                        objNegPago.Update(dtPagoFactura)
                        objNegPago = Nothing
                End Select
            End If
        Next
        BusinessHelper.UpdateTable(dtEC)
        Return blnEliminarEntregaRetencion
    End Function

    <Task()> Public Shared Sub EliminarRestriccionesDeleteEntregaCuenta(ByVal data As DatosElimRestricEnt, ByVal services As ServiceProvider)
        Dim objFilter As New Filter

        '//Comprobamos si la factura a borrar es una factura de ENTREGA SALDADA.
        objFilter.Clear()
        objFilter.Add(New NumberFilterItem("IDFacturaDestino", data.IDFactura))
        Select Case data.Circuito
            Case Circuito.Ventas
                objFilter.Add(New IsNullFilterItem("IDCliente", False))
            Case Circuito.Compras
                objFilter.Add(New IsNullFilterItem("IDProveedor", False))
        End Select

        Dim dtECSaldada As DataTable = New EntregasACuenta().Filter(objFilter)
        For Each drRowECSaldada As DataRow In dtECSaldada.Select
            If drRowECSaldada("TipoEntrega") <> enumTipoEntrega.Retencion Then
                drRowECSaldada("IDFacturaDestino") = System.DBNull.Value
                drRowECSaldada("Saldado") = False
            Else
                If Not CBool(drRowECSaldada("Generado")) AndAlso Length(drRowECSaldada("IDFactura")) = 0 Then
                    drRowECSaldada("IDFacturaDestino") = System.DBNull.Value
                    drRowECSaldada("Saldado") = False
                Else
                    If drRowECSaldada("GenerarFactura") Then
                        ApplicationService.GenerateError("Se ha realizado la factura de la retención. No se puede eliminar la Factura.")
                    End If
                End If
            End If
        Next
        BusinessHelper.UpdateTable(dtECSaldada)

        '//Comprobamos si la factura a borrar es una factura de ENTREGA GENERADA pero no SALDADA.
        objFilter.Clear()
        objFilter.Add(New NumberFilterItem("IDFactura", data.IDFactura))
        Select Case data.Circuito
            Case Circuito.Ventas
                objFilter.Add(New IsNullFilterItem("IDCliente", False))
            Case Circuito.Compras
                objFilter.Add(New IsNullFilterItem("IDProveedor", False))
        End Select
        Dim dtECGenerada As DataTable = New EntregasACuenta().Filter(objFilter)
        For Each drRowECGenerada As DataRow In dtECGenerada.Select
            If drRowECGenerada("TipoEntrega") <> enumTipoEntrega.Retencion Then
                If Not CBool(drRowECGenerada("Saldado")) AndAlso Length(drRowECGenerada("IDFacturaDestino")) = 0 Then
                    drRowECGenerada("IDFactura") = System.DBNull.Value
                    drRowECGenerada("Generado") = False
                Else
                    If drRowECGenerada("GenerarFactura") Then
                        ApplicationService.GenerateError("La Entrega está saldada, no se puede eliminar la Factura.")
                    Else
                        ApplicationService.GenerateError("La Entrega está saldada, no se puede eliminar el Cobro/Pago.")
                    End If
                End If
            Else
                drRowECGenerada("IDFactura") = System.DBNull.Value
                drRowECGenerada("Generado") = False
            End If
        Next
        BusinessHelper.UpdateTable(dtECGenerada)
    End Sub

    <Task()> Public Shared Sub EliminarRestriccionesDeleteEntregaCuentaCobroPago(ByVal data As DatosElimRestricEntCobro, ByVal services As ServiceProvider)
        Dim objFilter As New Filter
        objFilter.Add(New NumberFilterItem("IDEntrega", data.IDEntrega))
        objFilter.Add(New NumberFilterItem("IDCobroPago", data.IDCobroPago))
        Dim dtECCobroPago As DataTable = New EntregasACuenta().Filter(objFilter)
        For Each drRowECCobroPago As DataRow In dtECCobroPago.Select
            If Length(drRowECCobroPago("IDFacturaDestino")) > 0 OrElse CBool(drRowECCobroPago("Saldado")) Then
                ApplicationService.GenerateError("El Cobro/Pago está vinculado a una Factura.")
            Else
                drRowECCobroPago("IDCobroPago") = System.DBNull.Value
                drRowECCobroPago("Generado") = False
            End If
        Next
        BusinessHelper.UpdateTable(dtECCobroPago)
    End Sub

#End Region

#Region " Generar Nueva Entrega (Retención) desde Obras - Maestro de Entregas a Cuenta "

    <Task()> Public Shared Function NuevaEntregaTipoRetencionFacturaObra(ByVal data As DatosNuevaEntrega, ByVal services As ServiceProvider) As DataRow
        If Not IsNothing(data.DtCab.Rows(0)) AndAlso Length(data.DtCab.Rows(0)("IDObra")) > 0 Then
            Dim f As New Filter
            Dim dtInfoRetencion As DataTable
            Dim strCampoIDClienteProveedor As String
            Select Case data.Circuito
                Case Circuito.Ventas
                    '//Buscamos en la Obra la información de retención.
                    f.Clear()
                    f.Add(New NumberFilterItem("IDObra", data.DtCab.Rows(0)("IDObra")))
                    f.Add(New NumberFilterItem("Impuestos", TipoRetencionImpuestos.AntesImpuestos))
                    Dim OC As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraCabecera")
                    dtInfoRetencion = OC.Filter(f)
                    strCampoIDClienteProveedor = "IDCliente"
                Case Circuito.Compras
                    ' //Buscamos en la Obra si está el IDProveedor, en la lista de retenciones. Sólo traemos los 
                    ' //de antes de impuestos.
                    f.Clear()
                    f.Add(New StringFilterItem("IDProveedor", data.DtCab.Rows(0)("IDProveedor")))
                    f.Add(New NumberFilterItem("IDObra", data.DtCab.Rows(0)("IDObra")))
                    f.Add(New NumberFilterItem("Impuestos", TipoRetencionImpuestos.AntesImpuestos))
                    Dim OP As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraProveedor")
                    dtInfoRetencion = OP.Filter(f)
                    strCampoIDClienteProveedor = "IDProveedor"
            End Select

            If Not IsNothing(dtInfoRetencion) AndAlso dtInfoRetencion.Rows.Count > 0 Then
                Dim EC As New EntregasACuenta

                If IsNothing(data.DtEntregas) Then data.DtEntregas = EC.AddNew
                Dim drNewEntrega As DataRow = data.DtEntregas.NewRow
                drNewEntrega("IDEntrega") = AdminData.GetAutoNumeric
                EC.ApplyBusinessRule("TipoEntrega", CInt(enumTipoEntrega.Retencion), drNewEntrega)
                Dim dFechaEntrega As Date
                If Length(dtInfoRetencion.Rows(0)("FechaRetencion")) > 0 Then
                    dFechaEntrega = dtInfoRetencion.Rows(0)("FechaRetencion")
                Else
                    dFechaEntrega = DateAdd(New NegocioGeneral().GetPeriodString(dtInfoRetencion.Rows(0)("TipoPeriodo")), dtInfoRetencion.Rows(0)("Periodo"), data.DtCab.Rows(0)("FechaFactura"))
                End If
                EC.ApplyBusinessRule("FechaEntrega", dFechaEntrega, drNewEntrega)
                EC.ApplyBusinessRule(strCampoIDClienteProveedor, data.DtCab.Rows(0)(strCampoIDClienteProveedor), drNewEntrega)

                Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
                EC.ApplyBusinessRule("IDArticulo", AppParams.ArticuloRetencion, drNewEntrega)  'El genérico de la retención.

                Dim dblImporte As Double
                Select Case dtInfoRetencion.Rows(0)("TipoRetencion")
                    Case enumTipoRetencion.troSobreBI       '//Sobre Base Imponible
                        For Each drLinea As DataRow In data.DtLineas.Rows
                            dblImporte = dblImporte + Nz(drLinea("Importe"), 0)
                        Next
                    Case enumTipoRetencion.troSobreTotal    '//Sobre el Total
                        For Each drLinea As DataRow In data.DtLineas.Rows
                            Dim datImpIVA As New TipoIva.DataCalcularImporteIVA(drLinea("IDTipoIVA"), Nz(drLinea("Importe"), 0))
                            dblImporte = dblImporte + Nz(drLinea("Importe"), 0) + ProcessServer.ExecuteTask(Of TipoIva.DataCalcularImporteIVA, Double)(AddressOf TipoIva.CalcularImporteIVA, datImpIVA, services)
                        Next
                End Select
                Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
                Dim MonInfoA As MonedaInfo = Monedas.MonedaA
                dblImporte = xRound((dblImporte * Nz(dtInfoRetencion.Rows(0)("Retencion"), 0) / 100), MonInfoA.NDecimalesImporte)
                EC.ApplyBusinessRule("Importe", dblImporte, drNewEntrega)
                EC.ApplyBusinessRule("GenerarFactura", True, drNewEntrega)
                EC.ApplyBusinessRule("IDObra", data.DtCab.Rows(0)("IDObra"), drNewEntrega)
                drNewEntrega("Generado") = False
                drNewEntrega("Saldado") = True
                drNewEntrega("IDFacturaDestino") = data.DtCab.Rows(0)("IDFactura")  'Aquí irá el IDFactura que se generará
                data.DtEntregas.Rows.Add(drNewEntrega)
                Return drNewEntrega
            End If
        End If
    End Function

#End Region

#Region " Generar Cobro/Pago de Entregas a Cuenta "

    <Task()> Public Shared Function GenerarCobroPagoEntrega(ByVal data As DatosGenFactEntrega, ByVal services As ServiceProvider) As DataTable
        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
        Dim MonInfoA As MonedaInfo = Monedas.MonedaA

        Dim dtErrores As DataTable
        Dim strCampoCobroPago As String
        Dim strCampoTipoCobroPago As String
        Dim strCampoIDClienteProveedor As String

        Dim objNegCobroPago As BusinessHelper
        Select Case data.Circuito
            Case Circuito.Ventas
                objNegCobroPago = BusinessHelper.CreateBusinessObject(GetType(Cobro).Name)
                strCampoCobroPago = "IDCobro"
                strCampoTipoCobroPago = "IDTipoCobro"
                strCampoIDClienteProveedor = "IDCliente"
            Case Circuito.Compras
                objNegCobroPago = BusinessHelper.CreateBusinessObject(GetType(Pago).Name)
                strCampoCobroPago = "IDPago"
                strCampoTipoCobroPago = "IDTipoPago"
                strCampoIDClienteProveedor = "IDProveedor"
        End Select

        Dim dtCobroPago As DataTable
        For Each drRowEntrega As DataRow In data.DtEntregas.Select
            Try
                '//Generar el cobro/pago
                dtCobroPago = objNegCobroPago.AddNewForm
                If Not IsNothing(dtCobroPago) AndAlso dtCobroPago.Rows.Count > 0 Then
                    dtCobroPago.Rows(0)("IDEntrega") = drRowEntrega("IDEntrega")
                    objNegCobroPago.ApplyBusinessRule(strCampoIDClienteProveedor, drRowEntrega(strCampoIDClienteProveedor), dtCobroPago.Rows(0))
                    objNegCobroPago.ApplyBusinessRule("IDMoneda", Nz(drRowEntrega("IDMoneda"), MonInfoA.ID), dtCobroPago.Rows(0))
                    objNegCobroPago.ApplyBusinessRule("FechaVencimiento", drRowEntrega("FechaEntrega"), dtCobroPago.Rows(0))
                    objNegCobroPago.ApplyBusinessRule("ImpVencimiento", drRowEntrega("Importe"), dtCobroPago.Rows(0))
                    objNegCobroPago.ApplyBusinessRule(strCampoTipoCobroPago, drRowEntrega("IDTipoCobroPago"), dtCobroPago.Rows(0))
                    If Length(drRowEntrega("IDBancoPropio")) > 0 Then
                        objNegCobroPago.ApplyBusinessRule("IDBancoPropio", drRowEntrega("IDBancoPropio"), dtCobroPago.Rows(0))
                    Else
                        objNegCobroPago.ApplyBusinessRule("IDBancoPropio", System.DBNull.Value, dtCobroPago.Rows(0))
                    End If
                    objNegCobroPago.ApplyBusinessRule("IDFormaPago", drRowEntrega("IDFormaPago"), dtCobroPago.Rows(0))
                    objNegCobroPago.ApplyBusinessRule("IDObra", drRowEntrega("IDObra"), dtCobroPago.Rows(0))

                    ' En el caso de generar cobro o pago las cuentas contables van a revés que en la factura. Lo cambiamos aquí.
                    If Length(drRowEntrega("CCArticulo")) > 0 Then
                        dtCobroPago.Rows(0)("CContable") = drRowEntrega("CCArticulo")
                    End If
                End If
                dtCobroPago = objNegCobroPago.Update(dtCobroPago)
                If Not IsNothing(dtCobroPago) AndAlso dtCobroPago.Rows.Count > 0 Then
                    drRowEntrega("IDCobroPago") = dtCobroPago.Rows(0)(strCampoCobroPago)
                    drRowEntrega("Generado") = True
                End If
            Catch ex As Exception
                '//Si se ha producido algún error al generar alguna factura, indicamos el error y pasamos 
                '//a tratar la siguiente entrega.
                If IsNothing(dtErrores) Then
                    dtErrores = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf EntregasACuenta.CrearDTListaErrores, Nothing, services)
                End If
                Dim drRowError As DataRow = dtErrores.NewRow
                drRowError("IDEntrega") = drRowEntrega("IDEntrega")
                drRowError("TipoEntrega") = drRowEntrega("TipoEntrega")
                drRowError("FechaEntrega") = drRowEntrega("FechaEntrega")
                drRowError(strCampoIDClienteProveedor) = drRowEntrega(strCampoIDClienteProveedor)
                drRowError("IDArticulo") = drRowEntrega("IDArticulo")
                drRowError("Importe") = drRowEntrega("Importe")
                drRowError("Error") = ex.Message
                dtErrores.Rows.Add(drRowError)
            End Try
        Next drRowEntrega
        Dim ClsEnt As New EntregasACuenta
        ClsEnt.Update(data.DtEntregas)
        Return dtErrores
    End Function

#End Region

#Region " Añadir Entregas a una Factura  "
    <Serializable()> _
    Public Class DataEntregas
        Public IDEntregas() As Integer
        Public IDFactura As Integer

        Public Sub New(ByVal IDFactura As Integer, ByVal IDEntregas() As Integer)
            Me.IDFactura = IDFactura
            Me.IDEntregas = IDEntregas
        End Sub
    End Class

#Region " Añadir Entregas a una Factura de Compra "

    <Task()> Public Shared Sub AñadirEntregasAFacturaCompra(ByVal data As DataEntregas, ByVal services As ServiceProvider)
        If data.IDFactura > 0 Then
            '//Recuperamos la cabecera de la Factura Destino.
            Dim Doc As DocumentoFacturaCompra = New DocumentoFacturaCompra(data.IDFactura)

            '//Recuperamos las entregas seleccionadas.
            Dim IDEntregasCopy(data.IDEntregas.Length - 1) As Object
            data.IDEntregas.CopyTo(IDEntregasCopy, 0)
            Dim dtEntregas As DataTable = New EntregasACuenta().Filter(New InListFilterItem("IDEntrega", IDEntregasCopy, FilterType.Numeric))


            '/////////////////////////  ENTREGAS DE TIPO FACTURA  ///////////////////////////
            Dim datEntregas As New DataFacturaCompraEntregas(Doc, dtEntregas)
            ProcessServer.ExecuteTask(Of DataFacturaCompraEntregas)(AddressOf AddEntregasTipoFacturaCompras, datEntregas, services)

            '/////////////////////////  ENTREGAS DE TIPO COBRO/PAGO  ///////////////////////////
            ProcessServer.ExecuteTask(Of DataFacturaCompraEntregas)(AddressOf AddEntregasTipoPago, datEntregas, services)

            ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularAnaliticaFacturas, Doc, services)
            ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularVencimientos, Doc, services)

            Dim f As New Filter
            f.Add(New BooleanFilterItem("GenerarFactura", False))
            f.Add(New IsNullFilterItem("IDProveedor", False))
            Dim WhereEntregasTipoCobroPago As String = AdminData.ComposeFilter(f)
            Dim adrPagos As DataRow() = dtEntregas.Select(WhereEntregasTipoCobroPago)
            If Not adrPagos Is Nothing AndAlso adrPagos.Length > 0 Then
                ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularTotales, Doc, services)
            End If

            '//Guardamos la factura y las entregas a cuenta
            ProcessServer.ExecuteTask(Of Object)(AddressOf Business.General.Comunes.BeginTransaction, Nothing, services)
            ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf Business.General.Comunes.UpdateDocument, Doc, services)
            Dim ClsEnt As New EntregasACuenta
            ClsEnt.Update(dtEntregas)
            ProcessServer.ExecuteTask(Of Boolean)(AddressOf Business.General.Comunes.CommitTransaction, True, services)
        Else
            ApplicationService.GenerateError("Debe indicar una factura.")
        End If
    End Sub
    Public Class DataFacturaCompraEntregas
        Public Factura As DocumentoFacturaCompra
        Public Entregas As DataTable

        Public Sub New(ByVal Factura As DocumentoFacturaCompra, ByVal Entregas As DataTable)
            Me.Factura = Factura
            Me.Entregas = Entregas
        End Sub
    End Class
    <Task()> Public Shared Sub AddEntregasTipoFacturaCompras(ByVal data As DataFacturaCompraEntregas, ByVal services As ServiceProvider)
        If data.Entregas Is Nothing OrElse data.Entregas.Rows.Count = 0 Then Exit Sub

        Dim datValMoneda As New DataValidarMonedaEntregaEnFactura(data.Factura.HeaderRow("IDMoneda") & String.Empty, data.Entregas)
        ProcessServer.ExecuteTask(Of DataValidarMonedaEntregaEnFactura)(AddressOf ValidarMonedaEntregaEnFactura, datValMoneda, services)

        ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf ProcesoComunes.DescuentosCeroCabeceraFactura, data.Factura, services)

        Dim EntregasTipoFra As List(Of DataRow) = (From c In data.Entregas Where Not c.IsNull("IDProveedor") AndAlso Not c.IsNull("GenerarFactura") AndAlso CBool(c("GenerarFactura")) = True Select c).ToList
        If Not EntregasTipoFra Is Nothing AndAlso EntregasTipoFra.Count > 0 Then
            Dim context As New BusinessData(data.Factura.HeaderRow)
            Dim FL As New FacturaCompraLinea

            For Each drRowEntrega As DataRow In EntregasTipoFra
                '//Generar líneas de la factura 
                Dim drNewRow As DataRow = data.Factura.dtLineas.NewRow
                ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoFacturacionCompra.AsignarValoresPredeterminadosLinea, drNewRow, services)
                drNewRow("IDOrdenLinea") = 1
                drNewRow("IDCentroGestion") = data.Factura.HeaderRow("IDCentroGestion")
                drNewRow("IDEntrega") = drRowEntrega("IDEntrega")
                drNewRow("IDFactura") = data.Factura.HeaderRow("IDFactura")
                FL.ApplyBusinessRule("IDArticulo", drRowEntrega("IDArticulo"), drNewRow, context)
                ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoComunes.DescuentosCeroLinea, drNewRow, services)
                FL.ApplyBusinessRule("CContable", drRowEntrega("CCArticulo"), drNewRow, context)
                FL.ApplyBusinessRule("Cantidad", 1, drNewRow, context)
                FL.ApplyBusinessRule("Precio", (-1) * drRowEntrega("Importe"), drNewRow, context)

                Dim DataInfo As New TipoIva.DataCalcularImporteIVA(drNewRow("IDTipoIva"), drRowEntrega("Importe"))
                Dim dblImpIVA As Double = ProcessServer.ExecuteTask(Of TipoIva.DataCalcularImporteIVA, Double)(AddressOf TipoIva.CalcularImporteIVA, DataInfo, services)
                If Length(drRowEntrega("IDObra")) > 0 Then FL.ApplyBusinessRule("IDObra", drRowEntrega("IDObra"), drNewRow, context)
                data.Factura.dtLineas.Rows.Add(drNewRow)

                drRowEntrega("IDFacturaDestino") = data.Factura.HeaderRow("IDFactura")
                drRowEntrega("Saldado") = True
            Next drRowEntrega

            ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularImporteLineasFacturas, data.Factura, services)
            ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf NegocioGeneral.CalcularImpuestos, data.Factura, services)
            ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularBasesImponibles, data.Factura, services)
            ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularTotales, data.Factura, services)
        End If
    End Sub

    <Task()> Public Shared Sub AddEntregasTipoPago(ByVal data As DataFacturaCompraEntregas, ByVal services As ServiceProvider)
        If data.Entregas Is Nothing OrElse data.Entregas.Rows.Count = 0 Then Exit Sub

        Dim datValMoneda As New DataValidarMonedaEntregaEnFactura(data.Factura.HeaderRow("IDMoneda") & String.Empty, data.Entregas)
        ProcessServer.ExecuteTask(Of DataValidarMonedaEntregaEnFactura)(AddressOf ValidarMonedaEntregaEnFactura, datValMoneda, services)

        Dim EntregasTipoVto As List(Of DataRow) = (From c In data.Entregas Where Not c.IsNull("IDProveedor") AndAlso Not c.IsNull("GenerarFactura") AndAlso CBool(c("GenerarFactura")) = False Select c).ToList
        If Not EntregasTipoVto Is Nothing AndAlso EntregasTipoVto.Count > 0 Then
            Dim blnVtosAñadidos As Boolean = False
            Dim fPagos As New Filter(FilterUnionOperator.Or)
            For Each drRowEntrega As DataRow In EntregasTipoVto
                blnVtosAñadidos = True

                fPagos.Add(New NumberFilterItem("IDPago", drRowEntrega("IDCobroPago")))
                '//Actualizamos la Entrega
                drRowEntrega("IDFacturaDestino") = data.Factura.HeaderRow("IDFactura")
                drRowEntrega("Saldado") = True
            Next drRowEntrega

            '//Incluimos los pagos en la Factura
            If blnVtosAñadidos Then
                Dim dblImpAdelantos As Double = 0
                Dim dtPagos As DataTable = New Pago().Filter(fPagos)
                For Each drRowPago As DataRow In dtPagos.Select
                    drRowPago("IDFactura") = data.Factura.HeaderRow("IDFactura")
                    drRowPago("NFactura") = data.Factura.HeaderRow("NFactura")
                    dblImpAdelantos = dblImpAdelantos + drRowPago("ImpVencimiento")

                    data.Factura.dtPagos.ImportRow(drRowPago)
                Next
                'data.Factura.HeaderRow("ImpTotal") -= dblImpAdelantos
                Dim ValAyB As New ValoresAyB(New DataRowPropertyAccessor(data.Factura.HeaderRow), data.Factura.IDMoneda, data.Factura.CambioA, data.Factura.CambioB)
                ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)
            End If
        End If

    End Sub

#End Region

#Region " Añadir Entregas a una Factura de Venta "

    <Task()> Public Shared Sub AñadirEntregasAFacturaVenta(ByVal data As DataEntregas, ByVal services As ServiceProvider)
        If data.IDFactura > 0 Then
            '//Recuperamos la cabecera de la Factura Destino.
            Dim Doc As DocumentoFacturaVenta = New DocumentoFacturaVenta(data.IDFactura)

            '//Recuperamos las entregas seleccionadas.
            Dim IDEntregasCopy(data.IDEntregas.Length - 1) As Object
            data.IDEntregas.CopyTo(IDEntregasCopy, 0)
            Dim dtEntregas As DataTable = New EntregasACuenta().Filter(New InListFilterItem("IDEntrega", IDEntregasCopy, FilterType.Numeric))


            '/////////////////////////  ENTREGAS DE TIPO FACTURA  ///////////////////////////
            Dim datEntregas As New DataFacturaVentaEntregas(Doc, dtEntregas)
            ProcessServer.ExecuteTask(Of DataFacturaVentaEntregas)(AddressOf AddEntregasTipoFacturaVentas, datEntregas, services)

            '/////////////////////////  ENTREGAS DE TIPO COBRO/PAGO  ///////////////////////////
            ProcessServer.ExecuteTask(Of DataFacturaVentaEntregas)(AddressOf AddEntregasTipoCobro, datEntregas, services)

            ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularAnalitica, Doc, services)
            ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularVencimientos, Doc, services)

            Dim f As New Filter
            f.Add(New BooleanFilterItem("GenerarFactura", False))
            f.Add(New IsNullFilterItem("IDCliente", False))
            Dim WhereEntregasTipoCobroPago As String = AdminData.ComposeFilter(f)
            Dim adrCobros As DataRow() = dtEntregas.Select(WhereEntregasTipoCobroPago)
            If Not adrCobros Is Nothing AndAlso adrCobros.Length > 0 Then
                ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularTotales, Doc, services)
            End If

            '//Guardamos la factura y las entregas a cuenta
            ProcessServer.ExecuteTask(Of Object)(AddressOf Business.General.Comunes.BeginTransaction, Nothing, services)
            ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf Business.General.Comunes.UpdateDocument, Doc, services)
            Dim ClsEnt As New EntregasACuenta
            ClsEnt.Update(dtEntregas)
            ProcessServer.ExecuteTask(Of Boolean)(AddressOf Business.General.Comunes.CommitTransaction, True, services)
        Else
            ApplicationService.GenerateError("Debe indicar una factura.")
        End If
    End Sub
    <Serializable()> _
    Public Class DataFacturaVentaEntregas
        Public Factura As DocumentoFacturaVenta
        Public Entregas As DataTable

        Public Sub New(ByVal Factura As DocumentoFacturaVenta, ByVal Entregas As DataTable)
            Me.Factura = Factura
            Me.Entregas = Entregas
        End Sub
    End Class
    <Task()> Public Shared Sub AddEntregasTipoFacturaVentas(ByVal data As DataFacturaVentaEntregas, ByVal services As ServiceProvider)
        If data.Entregas Is Nothing OrElse data.Entregas.Rows.Count = 0 Then Exit Sub

        Dim datValMoneda As New DataValidarMonedaEntregaEnFactura(data.Factura.HeaderRow("IDMoneda") & String.Empty, data.Entregas)
        ProcessServer.ExecuteTask(Of DataValidarMonedaEntregaEnFactura)(AddressOf ValidarMonedaEntregaEnFactura, datValMoneda, services)

        ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoComunes.DescuentosCeroCabeceraFactura, data.Factura, services)

        Dim EntregasTipoFra As List(Of DataRow) = (From c In data.Entregas Where Not c.IsNull("IDCliente") AndAlso Not c.IsNull("GenerarFactura") AndAlso CBool(c("GenerarFactura")) = True Select c).ToList
        If Not EntregasTipoFra Is Nothing AndAlso EntregasTipoFra.Count > 0 Then
            Dim FL As New FacturaVentaLinea
            Dim context As New BusinessData(data.Factura.HeaderRow)
            For Each drRowEntrega As DataRow In EntregasTipoFra
                'If IsNothing(dtLineasFact) Then dtLineasFact = objNegLinFact.Filter(New NoRowsFilterItem)
                '//Generar líneas de la factura 
                Dim drNewRow As DataRow = data.Factura.dtLineas.NewRow
                ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoFacturacionVenta.AsignarValoresPredeterminadosLinea, drNewRow, services)
                drNewRow("IDOrdenLinea") = 1
                drNewRow("IDCentroGestion") = data.Factura.HeaderRow("IDCentroGestion")
                drNewRow("IDEntrega") = drRowEntrega("IDEntrega")
                drNewRow("IDFactura") = data.Factura.HeaderRow("IDFactura")
                FL.ApplyBusinessRule("IDArticulo", drRowEntrega("IDArticulo"), drNewRow, context)
                ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoComunes.DescuentosCeroLinea, drNewRow, services)
                FL.ApplyBusinessRule("CContable", drRowEntrega("CCArticulo"), drNewRow, context)
                FL.ApplyBusinessRule("Cantidad", 1, drNewRow, context)
                FL.ApplyBusinessRule("Precio", (-1) * drRowEntrega("Importe"), drNewRow, context)


                Dim DataInfo As New TipoIva.DataCalcularImporteIVA(drNewRow("IDTipoIva"), drRowEntrega("Importe"))
                Dim dblImpIVA As Double = ProcessServer.ExecuteTask(Of TipoIva.DataCalcularImporteIVA, Double)(AddressOf TipoIva.CalcularImporteIVA, DataInfo, services)
                If Length(drRowEntrega("IDObra")) > 0 Then FL.ApplyBusinessRule("IDObra", drRowEntrega("IDObra"), drNewRow, context)
                data.Factura.dtLineas.Rows.Add(drNewRow)

                drRowEntrega("IDFacturaDestino") = data.Factura.HeaderRow("IDFactura")
                drRowEntrega("Saldado") = True
            Next drRowEntrega

            ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularImporteLineasFacturas, data.Factura, services)
            ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf NegocioGeneral.CalcularImpuestos, data.Factura, services)

            ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularBasesImponibles, data.Factura, services)
            ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularTotales, data.Factura, services)
        End If
    End Sub

    <Serializable()> _
    Public Class DataValidarMonedaEntregaEnFactura
        Public IDMonedaFactura As String
        Public Entregas As DataTable

        Public Sub New(ByVal IDMonedaFactura As String, ByVal Entregas As DataTable)
            Me.IDMonedaFactura = IDMonedaFactura
            Me.Entregas = Entregas
        End Sub
    End Class
    <Task()> Public Shared Sub ValidarMonedaEntregaEnFactura(ByVal data As DataValidarMonedaEntregaEnFactura, ByVal services As ServiceProvider)
        If Length(data.IDMonedaFactura) > 0 AndAlso Not data.Entregas Is Nothing AndAlso data.Entregas.Rows.Count > 0 Then
            Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
            Dim MonInfoA As MonedaInfo = Monedas.MonedaA
            Dim IDMonedaFactura As String = data.IDMonedaFactura
            Dim IDMonedaEntrega As String

            Dim EntregasEnOtraMoneda As List(Of DataRow) = (From c In data.Entregas _
                                                            Where (c.IsNull("IDMoneda") AndAlso IDMonedaFactura <> MonInfoA.ID) OrElse _
                                                                  (Not c.IsNull("IDMoneda") AndAlso IDMonedaFactura <> c("IDMoneda")) _
                                                            Select c).ToList
            If Not EntregasEnOtraMoneda Is Nothing AndAlso EntregasEnOtraMoneda.Count > 0 Then
                IDMonedaEntrega = EntregasEnOtraMoneda(0)("IDMoneda") & String.Empty
            End If

            If Length(IDMonedaEntrega) > 0 Then
                ApplicationService.GenerateError("La Moneda de la Factura y la de la Entrega no coincide.{0}Moneda Factura: {1}{2}Moneda Entrega: {3}", vbNewLine, IDMonedaFactura, vbNewLine, IDMonedaEntrega)
            End If
        End If
    End Sub
    <Task()> Public Shared Sub AddEntregasTipoCobro(ByVal data As DataFacturaVentaEntregas, ByVal services As ServiceProvider)
        If data.Entregas Is Nothing OrElse data.Entregas.Rows.Count = 0 Then Exit Sub

        Dim datValMoneda As New DataValidarMonedaEntregaEnFactura(data.Factura.HeaderRow("IDMoneda") & String.Empty, data.Entregas)
        ProcessServer.ExecuteTask(Of DataValidarMonedaEntregaEnFactura)(AddressOf ValidarMonedaEntregaEnFactura, datValMoneda, services)

        Dim blnVtosAñadidos As Boolean = False
        Dim EntregasTipoVto As List(Of DataRow) = (From c In data.Entregas Where Not c.IsNull("IDCliente") AndAlso Not c.IsNull("GenerarFactura") AndAlso CBool(c("GenerarFactura")) = False Select c).ToList
        If Not EntregasTipoVto Is Nothing AndAlso EntregasTipoVto.Count > 0 Then
            Dim fCobros As New Filter(FilterUnionOperator.Or)
            For Each drRowEntrega As DataRow In EntregasTipoVto
                blnVtosAñadidos = True

                fCobros.Add(New NumberFilterItem("IDCobro", drRowEntrega("IDCobroPago")))
                '//Actualizamos la Entrega
                drRowEntrega("IDFacturaDestino") = data.Factura.HeaderRow("IDFactura")
                drRowEntrega("Saldado") = True
            Next drRowEntrega

            '//Incluimos los pagos en la Factura
            If blnVtosAñadidos Then
                Dim dblImpAdelantos As Double = 0
                Dim dtCobros As DataTable = New Cobro().Filter(fCobros)
                For Each drRowCobro As DataRow In dtCobros.Select
                    drRowCobro("IDFactura") = data.Factura.HeaderRow("IDFactura")
                    drRowCobro("NFactura") = data.Factura.HeaderRow("NFactura")
                    dblImpAdelantos = dblImpAdelantos + drRowCobro("ImpVencimiento")

                    data.Factura.dtCobros.ImportRow(drRowCobro)
                Next
                'data.Factura.HeaderRow("ImpTotal") -= dblImpAdelantos
                Dim ValAyB As New ValoresAyB(New DataRowPropertyAccessor(data.Factura.HeaderRow), data.Factura.IDMoneda, data.Factura.CambioA, data.Factura.CambioB)
                ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)
            End If
        End If
    End Sub

#End Region


    <Task()> Public Shared Sub QuitarEntregasAFacturaCompra(ByVal data As DataEntregas, ByVal services As ServiceProvider)

        ProcessServer.ExecuteTask(Of Object)(AddressOf General.Comunes.BeginTransaction, Nothing, services)

        Dim IDEntregas(data.IDEntregas.Length - 1) As Object
        data.IDEntregas.CopyTo(IDEntregas, 0)

        Dim dtEntregas As DataTable = New EntregasACuenta().Filter(New InListFilterItem("IDEntrega", IDEntregas, FilterType.Numeric))
        For Each drEntrega As DataRow In dtEntregas.Rows
            '//Actualizamos la Entrega a cuenta
            Dim datDelRes As New DatosElimRestricEntFn(drEntrega("IDEntrega"), data.IDFactura, Circuito.Compras)
            ProcessServer.ExecuteTask(Of DatosElimRestricEntFn)(AddressOf EliminarRestriccionesDeleteEntregaCuentaFn, datDelRes, services)
        Next

        '///////  Borramos la Entrega en la Factura (Linea o Pago/Cobro)  /////////
        Dim f As New Filter
        f.Add(New InListFilterItem("IDEntrega", IDEntregas, FilterType.Numeric))
        Dim WhereEntregas As String = AdminData.ComposeFilter(f)

        '// Desvinculamos los Pagos de la Factura
        Dim dtPagos As DataTable = New Pago().Filter(f)
        If Not dtPagos Is Nothing AndAlso dtPagos.Rows.Count > 0 Then
            For Each Vto As DataRow In dtPagos.Rows
                Vto("IDFactura") = System.DBNull.Value
                Vto("NFactura") = System.DBNull.Value
            Next
        End If
        BusinessHelper.UpdateTable(dtPagos)

        '//Eliminamos la/s línea/s de la Factura correspondientes a la/s entrega/s
        Dim FCL As New FacturaCompraLinea
        Dim dtFCL As DataTable = FCL.Filter(f)
        FCL.Delete(dtFCL)

        Dim DocFra As New DocumentoFacturaCompra(data.IDFactura)
        ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularBasesImponibles, DocFra, services)
        ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularTotales, DocFra, services)
        ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularVencimientos, DocFra, services)
        ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf Business.General.Comunes.UpdateDocument, DocFra, services)

        ProcessServer.ExecuteTask(Of Boolean)(AddressOf Business.General.Comunes.CommitTransaction, True, services)

    End Sub

    <Task()> Public Shared Sub QuitarEntregasAFacturaVenta(ByVal data As DataEntregas, ByVal services As ServiceProvider)

        ProcessServer.ExecuteTask(Of Object)(AddressOf General.Comunes.BeginTransaction, Nothing, services)

        Dim IDEntregas(data.IDEntregas.Length - 1) As Object
        data.IDEntregas.CopyTo(IDEntregas, 0)

        Dim dtEntregas As DataTable = New EntregasACuenta().Filter(New InListFilterItem("IDEntrega", IDEntregas, FilterType.Numeric))
        For Each drEntrega As DataRow In dtEntregas.Rows
            '//Actualizamos la Entrega a cuenta
            Dim datDelRes As New DatosElimRestricEntFn(drEntrega("IDEntrega"), data.IDFactura, Circuito.Ventas)
            ProcessServer.ExecuteTask(Of DatosElimRestricEntFn)(AddressOf EliminarRestriccionesDeleteEntregaCuentaFn, datDelRes, services)
        Next

        '///////  Borramos la Entrega en la Factura (Linea o Pago/Cobro)  /////////
        Dim f As New Filter
        f.Add(New InListFilterItem("IDEntrega", IDEntregas, FilterType.Numeric))
        Dim WhereEntregas As String = AdminData.ComposeFilter(f)

        '// Desvinculamos los Cobros de la Factura
        Dim dtCobros As DataTable = New Cobro().Filter(f)
        If Not dtCobros Is Nothing AndAlso dtCobros.Rows.Count > 0 Then
            For Each Vto As DataRow In dtCobros.Rows
                Vto("IDFactura") = System.DBNull.Value
                Vto("NFactura") = System.DBNull.Value
            Next
        End If
        BusinessHelper.UpdateTable(dtCobros)

        '//Eliminamos la/s línea/s de la Factura correspondientes a la/s entrega/s
        Dim FVL As New FacturaVentaLinea
        Dim dtFVL As DataTable = FVL.Filter(f)
        FVL.Delete(dtFVL)

        Dim DocFra As New DocumentoFacturaVenta(data.IDFactura)
        ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularBasesImponibles, DocFra, services)
        ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularTotales, DocFra, services)
        ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularVencimientos, DocFra, services)
        ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf Business.General.Comunes.UpdateDocument, DocFra, services)

        ProcessServer.ExecuteTask(Of Boolean)(AddressOf Business.General.Comunes.CommitTransaction, True, services)
    End Sub


#End Region

    <Task()> Public Shared Function CrearDTListaErrores(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
        Dim dt As New DataTable
        dt.RemotingFormat = SerializationFormat.Binary
        With dt
            .Columns.Add("IDEntrega", GetType(Integer))
            .Columns.Add("TipoEntrega", GetType(Integer))
            .Columns.Add("FechaEntrega", GetType(Date))
            .Columns.Add("IDCliente", GetType(String))
            .Columns.Add("IDProveedor", GetType(String))
            .Columns.Add("IDArticulo", GetType(String))
            .Columns.Add("Importe", GetType(Double))
            .Columns.Add("Error", GetType(String))
        End With

        Return dt
    End Function

End Class