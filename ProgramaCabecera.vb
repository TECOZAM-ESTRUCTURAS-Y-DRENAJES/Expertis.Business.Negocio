Public Class ProgramaCabeceraInfo
    Inherits ClassEntityInfo

    Public IDPrograma As String
    Public IDDireccionEnvio As Integer
    Public DescPrograma As String
    Public ProgramaCliente As String
    Public FechaPrograma As Date
    Public IDArticulo As String
    Public IDAlmacen As String
    Public Activo As Boolean
    Public IDCliente As String
    Public IDCentroGestion As String
    Public IDContador As String
    Public IDMoneda As String
    Public EDI As Boolean
    Public IDPedido As Integer

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal data As DataRow)
        MyBase.New(data)
    End Sub

    Public Overloads Overrides Sub Fill(ByVal ParamArray PrimaryKey() As Object)
        Dim dt As DataTable
        If Not IsNothing(PrimaryKey) AndAlso PrimaryKey.Length > 0 AndAlso Length(PrimaryKey(0)) > 0 Then
            dt = New ProgramaCabecera().SelOnPrimaryKey(PrimaryKey(0))
        End If

        If IsNothing(dt) OrElse dt.Rows.Count = 0 Then
            ApplicationService.GenerateError("El Programa de Venta | no existe.", Quoted(PrimaryKey(0)))
        Else
            Me.Fill(dt.Rows(0))
        End If
    End Sub

End Class

Public Class ProgramaCabecera

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbProgramaCabecera"

#End Region

#Region "Eventos RegisterAddNewTasks"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        ' Se obtiene cual es el centro de gestión del usuario
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarCentroGestionUsuario, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarAlmacen, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarContador, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarDireccionEnvio, data, services)

        data("FechaPrograma") = Date.Today
        data("Activo") = enumpcEstadoPrograma.pcActivo
    End Sub

    <Task()> Public Shared Sub AsignarCentroGestionUsuario(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim cgu As New UsuarioCentroGestion.UsuarioCentroGestionInfo
        cgu = ProcessServer.ExecuteTask(Of UsuarioCentroGestion.UsuarioCentroGestionInfo, UsuarioCentroGestion.UsuarioCentroGestionInfo)(AddressOf UsuarioCentroGestion.ObtenerUsuarioCentroGestion, cgu, services)
        Dim IDCentroGestion As String = cgu.IDCentroGestion
        If Length(IDCentroGestion) > 0 Then
            data("IDCentroGestion") = IDCentroGestion
        Else : data("IDCentroGestion") = New Parametro().CGestionPredet
        End If
        Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
        If AppParams.AlmacenCentroGestionActivo Then
            data("IDAlmacen") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Almacen.GetAlmacenCentroGestion, IDCentroGestion, services)
        End If
    End Sub

    <Task()> Public Shared Sub AsignarAlmacen(ByVal data As DataRow, ByVal services As ServiceProvider)
        If New Parametro().AlmacenCentroGestionActivo Then
            If Length(data("IDAlmacen")) = 0 Then data("IDAlmacen") = New Parametro().AlmacenPredeterminado
        Else : data("IDAlmacen") = New Parametro().AlmacenPredeterminado
        End If
    End Sub

    <Task()> Public Shared Sub AsignarContador(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim StDatos As New Contador.DatosDefaultCounterValue(data, GetType(ProgramaCabecera).Name, "IDPrograma")
        ProcessServer.ExecuteTask(Of Contador.DatosDefaultCounterValue)(AddressOf Contador.LoadDefaultCounterValue, StDatos, services)
    End Sub

    <Task()> Public Shared Sub AsignarDireccionEnvio(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.IsNull("IDDireccionEnvio") Then
            If Length(data("IDCliente")) > 0 Then
                Dim parametrosdireccion As New ClienteDireccion.DataDirecEnvio(data("IDCliente"), enumcdTipoDireccion.cdDireccionEnvio)
                Dim direccion As DataTable = ProcessServer.ExecuteTask(Of ClienteDireccion.DataDirecEnvio, DataTable)(AddressOf ClienteDireccion.ObtenerDireccionEnvio, parametrosdireccion, services)
                If Not direccion Is Nothing AndAlso direccion.Rows.Count > 0 Then
                    data("IDDireccionEnvio") = direccion.Rows(0)("IDDireccion")
                End If
            End If
        End If
    End Sub

#End Region

#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("IDArticulo", AddressOf CambioArticulo)
        oBrl.Add("IDCliente", AddressOf CambioCliente)
        oBrl.Add("IDAlmacen", AddressOf CambioAlmacen)
        Return oBrl
    End Function

    <Task()> Public Shared Sub CambioArticulo(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
            Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.Value)
            If Not ArtInfo.Venta Then
                ApplicationService.GenerateError("El artículo | no es de tipo venta.", Quoted(data.Value))
            ElseIf Not ArtInfo.Activo Then
                ApplicationService.GenerateError("El artículo | no está activo.", Quoted(data.Value))
            Else
                If New Parametro().AlmacenCentroGestionActivo Then
                    data.Current("IDAlmacen") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Almacen.GetAlmacenCentroGestion, data.Current("IDCentroGestion"), services)
                    If Length(data.Current("IDAlmacen")) = 0 Then data.Current("IDAlmacen") = New Parametro().AlmacenPredeterminado
                Else
                    Dim objFilter As New Filter
                    objFilter.Add(New StringFilterItem("IDArticulo", data.Value))
                    objFilter.Add(New BooleanFilterItem("Predeterminado", True))
                    Dim dtAlmacen As DataTable = New ArticuloAlmacen().Filter(objFilter)
                    If Not IsNothing(dtAlmacen) AndAlso dtAlmacen.Rows.Count > 0 Then
                        data.Current("IDAlmacen") = dtAlmacen.Rows(0)("IDAlmacen")
                    Else : data.Current("IDAlmacen") = New Parametro().AlmacenPredeterminado
                    End If
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioCliente(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
            Dim ClteInfo As ClienteInfo = Clientes.GetEntity(data.Value)
            data.Current("IdMoneda") = ClteInfo.Moneda

            Dim StDatosDirec As New ClienteDireccion.DataDirecEnvio(data.Value, enumpdTipoDireccion.pdDireccionPedido)
            Dim dtDireccion As DataTable = ProcessServer.ExecuteTask(Of ClienteDireccion.DataDirecEnvio, DataTable)(AddressOf ClienteDireccion.ObtenerDireccionEnvio, StDatosDirec, services)

            If Not dtDireccion Is Nothing AndAlso dtDireccion.Rows.Count <> 0 Then
                data.Current("IDDireccionEnvio") = dtDireccion.Rows(0)("IDDireccion")
            End If
        Else : data.Current("IdMoneda") = System.DBNull.Value
        End If
    End Sub

    <Task()> Public Shared Sub CambioAlmacen(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Current("IdCentroGestion")) = 0 Then
            Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
            data.Current("IdCentroGestion") = AppParams.CentroGestion
        End If
    End Sub

#End Region

#Region "EventosValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarArticulo)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarPrograma)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarCliente)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarAlmacen)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarMoneda)
    End Sub

    <Task()> Public Shared Sub ComprobarArticulo(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDArticulo")) = 0 And Length(data("EDI")) = 1 Then ApplicationService.GenerateError("El Artículo es obligatorio.")
    End Sub

    <Task()> Public Shared Sub ComprobarPrograma(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDPrograma")) = 0 Then ApplicationService.GenerateError("El identificador del Programa de Entrega no es válido. ")
    End Sub

    <Task()> Public Shared Sub ComprobarCliente(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IdCliente")) = 0 Then ApplicationService.GenerateError("El Cliente es obligatorio.")
    End Sub

    <Task()> Public Shared Sub ComprobarAlmacen(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDAlmacen")) = 0 Then ApplicationService.GenerateError("El Almacén es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ComprobarMoneda(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDMoneda")) = 0 Then ApplicationService.GenerateError("La moneda es una dato obligatorio.")
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    <Serializable()> _
    Public Class DatosRecalLinea
        Public IDPrograma As String
        Public MonedaOriginal As String
        Public MonedaNueva As String
    End Class

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClave)
        updateProcess.AddTask(Of DataRow)(AddressOf CambioMonedaLineas)
    End Sub

    <Task()> Public Shared Sub AsignarClave(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IdContador")) > 0 Then data("IDPrograma") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data("IDContador"), services)
            If Length(data("FechaPrograma")) = 0 Then data("FechaPrograma") = Today.Date
        End If
    End Sub

    <Task()> Public Shared Sub CambioMonedaLineas(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Modified Then
            If Length(data("IDMoneda", DataRowVersion.Original)) > 0 Then
                If data("IDMoneda", DataRowVersion.Original) <> data("IDMoneda") Then
                    '//Recalcular precios e importes de las lineas de programa por cambio de moneda
                    Dim StDatos As New DatosRecalLinea
                    StDatos.IDPrograma = data("IDPrograma")
                    StDatos.MonedaOriginal = data("IDMoneda", DataRowVersion.Original)
                    StDatos.MonedaNueva = data("IDMoneda")
                    ProcessServer.ExecuteTask(Of DatosRecalLinea)(AddressOf RecalcularLineasCambioMoneda, StDatos, services)
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub RecalcularLineasCambioMoneda(ByVal data As DatosRecalLinea, ByVal services As ServiceProvider)
        Dim dtLineas As DataTable = New ProgramaLinea().Filter(New FilterItem("IDPrograma", FilterOperator.Equal, data.IDPrograma))
        If Not IsNothing(dtLineas) AndAlso dtLineas.Rows.Count > 0 Then
            Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
            Dim MonOld As MonedaInfo = Monedas.GetMoneda(data.MonedaOriginal)
            Dim MonNew As MonedaInfo = Monedas.GetMoneda(data.MonedaNueva)

            If Not MonOld Is Nothing AndAlso Not MonNew Is Nothing Then
                Dim dblRazonCambios As Double
                If MonNew.CambioA > 0 Then
                    dblRazonCambios = (MonOld.CambioA / MonNew.CambioA)
                Else : ApplicationService.GenerateError("Error en el cálculo de los precios e importes de las líneas (división por cero). El Cambio en una Moneda Interna no puede valer cero. ")
                End If

                For Each drRowLinea As DataRow In dtLineas.Select
                    drRowLinea("Precio") = xRound(drRowLinea("Precio") * dblRazonCambios, MonNew.NDecimalesPrecio)
                    drRowLinea("Importe") = xRound(drRowLinea("Importe") * dblRazonCambios, MonNew.NDecimalesImporte)
                Next
                BusinessHelper.UpdateTable(dtLineas)
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterDeleteTasks"

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf ComprobarPedidos)
    End Sub

    <Task()> Public Shared Sub ComprobarPedidos(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDPedido")) > 0 Then ApplicationService.GenerateError("El programa de venta tiene asociada un pedido de venta. No se puede borrar el Programa")
    End Sub

#End Region

End Class