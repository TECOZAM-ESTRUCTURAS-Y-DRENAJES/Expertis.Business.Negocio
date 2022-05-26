Public Class CentroGestionInfo
    Inherits ClassEntityInfo

    Public IDCentroGestion As String
    Public DescCentroGestion As String
    Public Direccion As String
    Public CodPostal As String
    Public Poblacion As String
    Public Provincia As String
    Public Telefono As String
    Public Fax As String
    Public Email As String
    Public IDPais As String
    Public IdContadorPedidoVenta As String
    Public IdContadorAlbaranVenta As String
    Public IdContadorFacturaVenta As String
    Public IdContadorPedidoCompra As String
    Public IdContadorAlbaranCompra As String
    Public IdContadorFacturaCompra As String
    Public IDContadorObra As String
    Public IDContadorCodTrabajo As String
    Public IDContadorObraPresup As String
    Public IdContadorAlbaranVentaTPV As String
    Public IDContadorVale As String
    Public IDContadorProvisionalTPV As String
    Public IDContadorAvisos As String
    Public IDCliente As String
    Public IDTarifa As String
    Public SedeCentral As Boolean
    Public GestionStock As Boolean
    Public SolicitOtroCentro As Boolean
    Public ModificarSolicit As Boolean
    Public ValidCantidad As Boolean
    Public PrepararExpedicion As Boolean
    Public LanzarExpedicion As Boolean
    Public LanzarPedidoCompra As Boolean
    Public LanzarOfertaCompra As Boolean
    Public CambiarEstado As Boolean
    Public CambiarEstSolicit As Boolean
    Public FactorDimCentro As Integer
    Public IDBuzonEDI As String
    Public IDConsignatario As String
    Public IDBancoPropio As String
    Public IDObraCalendario As Integer
    Public FechaUltimaSincronizacion As Date
    Public IDGrafico As Integer
    Public IDCAE As String
    Public NIDPB As String
	
    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal data As DataRow)
        MyBase.New(data)
    End Sub

    Public Overloads Overrides Sub Fill(ByVal ParamArray PrimaryKey() As Object)
        Dim dt As DataTable
        If Not IsNothing(PrimaryKey) AndAlso PrimaryKey.Length > 0 AndAlso Length(PrimaryKey(0)) > 0 Then
            dt = New CentroGestion().SelOnPrimaryKey(PrimaryKey(0))
        End If

        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            ApplicationService.GenerateError("El Centro de Gestión | no existe.", Quoted(PrimaryKey(0)))
        Else
            Me.Fill(dt.Rows(0))
        End If
    End Sub

End Class

Public Class CentroGestion

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub
    Private Const cnEntidad As String = "tbMaestroCentroGestion"

#End Region

    Public Enum ContadorEntidad
        PedidoVenta = 0
        PedidoCompra = 1
        AlbaranVenta = 2
        AlbaranCompra = 3
        FacturaVenta = 4
        FacturaCompra = 5
        AlbaranVentaTPV = 6
        BdgOperacion = 7
        BdgOperacionPlan = 8
    End Enum

#Region "Funciones RegisterAddNewTasks"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf AsignarTarifaPred)
        addnewProcess.AddTask(Of DataRow)(AddressOf AsignarClientePred)
    End Sub

    <Task()> Public Shared Sub AsignarTarifaPred(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim StrTarifaPred As String = New Parametro().TarifaPredeterminada
        If Len(StrTarifaPred) > 0 Then
            data("IDTarifa") = StrTarifaPred
        Else : ApplicationService.GenerateError("No existe configurada una tarifa predeterminada.")
        End If
    End Sub

    <Task()> Public Shared Sub AsignarClientePred(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim StrCliePred As String = New Parametro().ClientePredeterminado
        If Len(StrCliePred) > 0 Then
            data("IDCliente") = StrCliePred
        Else : ApplicationService.GenerateError("No se ha configurado ningún cliente predeterminado.")
        End If
    End Sub

#End Region

#Region "Funciones Validate / Update"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarIDCentro)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDescCentro)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClavePrincipal)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarPais)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarCliente)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarTarifa)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarContadorPedidoCompra)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarContadorAlbaranCompra)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarContadorFacturaCompra)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarContadorPedidoVenta)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarContadorAlbaranVenta)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarContadorFacturaVenta)
    End Sub

    <Task()> Public Shared Sub ValidarIDCentro(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDCentroGestion")) = 0 Then ApplicationService.GenerateError("El Centro Gestión es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarDescCentro(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("DescCentroGestion")) = 0 Then ApplicationService.GenerateError("La descripción del Centro Gestión es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarClavePrincipal(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New CentroGestion().SelOnPrimaryKey(data("IDCentroGestion"))
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("Este Centro Gestión ya existe en la base de datos.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarPais(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDPais")) > 0 Then
            Dim DtPais As DataTable = New Pais().SelOnPrimaryKey(data("IDPais"))
            If DtPais Is Nothing OrElse DtPais.Rows.Count = 0 Then
                ApplicationService.GenerateError("El país elegido no existe en la base de datos")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarCliente(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDCliente")) > 0 Then
            Dim DtClie As DataTable = New Cliente().SelOnPrimaryKey(data("IDCliente"))
            If DtClie Is Nothing OrElse DtClie.Rows.Count = 0 Then
                ApplicationService.GenerateError("El cliente elegido no existe en la base de datos")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarTarifa(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDTarifa")) > 0 Then
            Dim DtTarifa As DataTable = New Tarifa().SelOnPrimaryKey(data("IDTarifa"))
            If DtTarifa Is Nothing OrElse DtTarifa.Rows.Count = 0 Then
                ApplicationService.GenerateError("La tarifa elegida no existe en la base de datos")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarContadorPedidoCompra(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDContadorPedidoCompra")) > 0 Then
            Dim DtPed As DataTable = New Contador().SelOnPrimaryKey(data("IDContadorPedidoCompra"))
            If DtPed Is Nothing OrElse DtPed.Rows.Count = 0 Then
                ApplicationService.GenerateError("El contador elegido para pedido compra no existe en la base de datos")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarContadorAlbaranCompra(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDContadorAlbaranCompra")) > 0 Then
            Dim DtAlb As DataTable = New Contador().SelOnPrimaryKey(data("IDContadorAlbaranCompra"))
            If DtAlb Is Nothing OrElse DtAlb.Rows.Count = 0 Then
                ApplicationService.GenerateError("El contador elegido para albarán compra no existe en la base de datos")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarContadorFacturaCompra(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDContadorFacturaCompra")) > 0 Then
            Dim DtFact As DataTable = New Contador().SelOnPrimaryKey(data("IDContadorFacturaCompra"))
            If DtFact Is Nothing OrElse DtFact.Rows.Count = 0 Then
                ApplicationService.GenerateError("El contador elegido para factura compra no existe en la base de datos")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarContadorPedidoVenta(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDContadorPedidoVenta")) > 0 Then
            Dim DtPed As DataTable = New Contador().SelOnPrimaryKey(data("IDContadorPedidoVenta"))
            If DtPed Is Nothing OrElse DtPed.Rows.Count = 0 Then
                ApplicationService.GenerateError("El contador elegido para pedido venta no existe en la base de datos")
            End If
        End If

    End Sub

    <Task()> Public Shared Sub ValidarContadorAlbaranVenta(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDContadorAlbaranVenta")) > 0 Then
            Dim DtAlb As DataTable = New Contador().SelOnPrimaryKey(data("IDContadorAlbaranVenta"))
            If DtAlb Is Nothing OrElse DtAlb.Rows.Count = 0 Then
                ApplicationService.GenerateError("El contador elegido para albarán venta no existe en la base de datos")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarContadorFacturaVenta(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDContadorFacturaVenta")) > 0 Then
            Dim DtFact As DataTable = New Contador().SelOnPrimaryKey(data("IDContadorFacturaVenta"))
            If DtFact Is Nothing OrElse DtFact.Rows.Count = 0 Then
                ApplicationService.GenerateError("El contador elegido para factura venta no existe en la base de datos")
            End If
        End If
    End Sub

#End Region

#Region " GetBusinessRules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("CodPostal", AddressOf CambioCodPostal)
        Return oBrl
    End Function

    <Task()> Public Shared Sub CambioCodPostal(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim infoCP As New CodPostalInfo(CStr(data.Value), data.Current("IDPais") & String.Empty)
            If Length(infoCP.DescPoblacion) > 0 Then data.Current("Poblacion") = infoCP.DescPoblacion
            If Length(infoCP.DescProvincia) > 0 Then data.Current("Provincia") = infoCP.DescProvincia
            If Length(infoCP.IDPais) > 0 Then data.Current("IDPais") = infoCP.IDPais
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function GetContadorPredeterminadoCGestionUsuario(ByVal data As CentroGestion.ContadorEntidad, ByVal services As ServiceProvider) As String
        Dim CentroEnt As New CentroEntidad
        Dim cgu As New UsuarioCentroGestion.UsuarioCentroGestionInfo
        cgu = ProcessServer.ExecuteTask(Of UsuarioCentroGestion.UsuarioCentroGestionInfo, UsuarioCentroGestion.UsuarioCentroGestionInfo)(AddressOf UsuarioCentroGestion.ObtenerUsuarioCentroGestion, cgu, services)
        CentroEnt.CentroGestion = cgu.IDCentroGestion

        CentroEnt.ContadorEntidad = data
        Return ProcessServer.ExecuteTask(Of CentroEntidad, String)(AddressOf GetContadorPredeterminado, CentroEnt, services)
    End Function

    <Task()> Public Shared Function GetContadorPredeterminado(ByVal CentroEntidad As CentroEntidad, ByVal services As ServiceProvider) As String
        'Dim TipoContador As CentroGestion.ContadorEntidad = services.GetService(GetType(CentroGestion.ContadorEntidad))
        'If Length(strCentroGestion) = 0 Then
        '    Dim ucg As New UsuarioCentroGestion
        '    strCentroGestion = ucg.CentroGestionUsuario()
        'End If
        Dim strCampoContador As String = String.Empty
        Dim strEntity As String = String.Empty
        Select Case CentroEntidad.ContadorEntidad
            Case ContadorEntidad.PedidoCompra
                strCampoContador = "IDContadorPedidoCompra" : strEntity = GetType(PedidoCompraCabecera).Name
            Case ContadorEntidad.AlbaranCompra
                strCampoContador = "IDContadorAlbaranCompra" : strEntity = GetType(AlbaranCompraCabecera).Name
            Case ContadorEntidad.FacturaCompra
                strCampoContador = "IDContadorFacturaCompra" : strEntity = GetType(FacturaCompraCabecera).Name
            Case ContadorEntidad.PedidoVenta
                strCampoContador = "IDContadorPedidoVenta" : strEntity = GetType(PedidoVentaCabecera).Name
            Case ContadorEntidad.AlbaranVenta
                strCampoContador = "IDContadorAlbaranVenta" : strEntity = GetType(AlbaranVentaCabecera).Name
            Case ContadorEntidad.FacturaVenta
                strCampoContador = "IDContadorFacturaVenta" : strEntity = GetType(FacturaVentaCabecera).Name
            Case ContadorEntidad.AlbaranVentaTPV
                strCampoContador = "IDContadorAlbaranVentaTPV" : strEntity = GetType(AlbaranVentaCabecera).Name
            Case ContadorEntidad.BdgOperacion
                strCampoContador = "IDContadorBdgOperacion" : strEntity = "BdgOperacion"
            Case ContadorEntidad.BdgOperacionplan
                strCampoContador = "IDContadorBdgOperacionPlan" : strEntity = "BdgOperacionPlan"
        End Select

        Dim strContador As String = String.Empty
        If Length(CentroEntidad.CentroGestion) > 0 Then
            Dim DrCentroGest As DataRow = New CentroGestion().GetItemRow(CentroEntidad.CentroGestion)
            If Length(DrCentroGest(strCampoContador)) > 0 Then
                strContador = DrCentroGest(strCampoContador)
            End If
        End If
        If Len(strContador) = 0 Then
            Dim dtContador As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf Contador.CounterDefault, strEntity, services)
            If Not IsNothing(dtContador) AndAlso dtContador.Rows.Count > 0 Then
                strContador = dtContador.Rows(0)("IDContador")
            End If
        End If
        Return strContador
    End Function

#End Region

End Class

<Serializable()> _
Public Class CentroEntidad
    Public CentroGestion As String
    Public ContadorEntidad As CentroGestion.ContadorEntidad
End Class