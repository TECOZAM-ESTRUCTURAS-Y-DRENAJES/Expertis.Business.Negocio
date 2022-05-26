Public Class ArticuloClienteInfo
    Inherits ClassEntityInfo

    Public IDCliente As String
    Public IDArticulo As String
    Public RefCliente As String
    Public DescRefCliente As String
    Public IDUDVenta As String
    Public IDUDExpedicion As String
    Public UdValoracion As Integer
    Public Revision As String
    Public IDLineaOfertaDetalle As Integer
    Public Precio As Double
    Public Dto1 As Double
    Public Dto2 As Double
    Public Dto3 As Double
    Public IDArticuloContenedor As String
    Public QContenedor As Double
    Public IDArticuloEmbalaje As String
    Public QEmbalaje As Double
    Public PVP As Double

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal data As DataRow)
        MyBase.New(data)
    End Sub

    Public Overloads Overrides Sub Fill(ByVal ParamArray PrimaryKey() As Object)
        Dim dt As DataTable = New ArticuloCliente().SelOnPrimaryKey(PrimaryKey)
        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            Me.Fill(dt.Rows(0))
        End If
    End Sub

End Class

Public Class ArticuloCliente
#Region "Constructor"
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbArticuloCliente"
#End Region

#Region "Eventos RegisterDeleteTasks"

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf ActualizarOfertaComercialDetalle)
    End Sub

   
    <Task()> Public Shared Sub ActualizarOfertaComercialDetalle(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDLineaOfertaDetalle")) > 0 Then
            Dim f As New Filter
            f.Add(New NumberFilterItem("IDLineaOfertaDetalle", data("IDLineaOfertaDetalle")))
            Dim dtO As DataTable = New BE.DataEngine().Filter("tbOfertaComercialDetalle", f, "IDLineaOfertaDetalle,EstadoCliente")
            If Not dtO Is Nothing AndAlso dtO.Rows.Count > 0 Then
                dtO.Rows(0)("EstadoCliente") = False
                dtO.TableName = "OfertaComercialDetalle"
                BusinessHelper.UpdateTable(dtO)
            End If
        End If
    End Sub
#End Region
#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("UDValoracion", AddressOf ComprobarValor)
        oBrl.Add("Dto1", AddressOf ComprobarValor)
        oBrl.Add("Dto2", AddressOf ComprobarValor)
        oBrl.Add("Dto3", AddressOf ComprobarValor)
        oBrl.Add("IDArticulo", AddressOf CambioArticulo)
        oBrl.Add("Precio", AddressOf ArticuloClienteLinea.CambioPrecios)
        oBrl.Add("PVP", AddressOf ArticuloClienteLinea.CambioPrecios)
        Return oBrl
    End Function

    <Task()> Public Shared Sub ComprobarValor(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Not IsNumeric(data.Value) Then ApplicationService.GenerateError("Campo no numérico.")
    End Sub

    <Task()> Public Shared Sub CambioArticulo(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
            Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.Value)
            data.Current("CodigoEDI") = ArtInfo.CodigoBarras
            data.Current("UDValoracion") = ArtInfo.UDValoracion
            data.Current("IDUDVenta") = ArtInfo.IDUDVenta
        End If

    End Sub

#End Region
#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf NegocioGeneral.ValidarClienteObligatorio, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf NegocioGeneral.ValidarArticuloObligatorio, data, services)
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dtAP As DataTable = New ArticuloCliente().SelOnPrimaryKey(data("IDCliente"), data("IDArticulo"))
            If Not dtAP Is Nothing AndAlso dtAP.Rows.Count > 0 Then
                ApplicationService.GenerateError("El Artículo | ya tiene asociado el Cliente |.", Quoted(data("IDArticulo")), Quoted(data("IDCliente")))
            End If
        End If

    End Sub

#End Region
#Region "Eventos RegisterUpdateTask"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf ActualizarUnidadValoracion)
        updateProcess.AddTask(Of DataRow)(AddressOf AplicarDecimales)
    End Sub

    <Task()> Public Shared Sub ActualizarUnidadValoracion(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Nz(data("UdValoracion"), 0) = 0 Then data("UdValoracion") = 1
    End Sub

    <Task()> Public Shared Sub AplicarDecimales(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDCliente")) > 0 Then
            Dim clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
            Dim CliInfo As ClienteInfo = clientes.GetEntity(data("IDCliente"))
            Dim IDMoneda As String
            If Length(CliInfo.Moneda) > 0 Then
                IDMoneda = CliInfo.Moneda
            Else
                Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
                IDMoneda = Monedas.MonedaA.ID
            End If
            Dim datosDec As New DataAplicarDecimalesMoneda(IDMoneda, data)
            ProcessServer.ExecuteTask(Of DataAplicarDecimalesMoneda)(AddressOf NegocioGeneral.AplicarDecimalesMoneda, datosDec, services)
        End If
    End Sub

#End Region
#Region "Funciones Públicas"
    <Serializable()> _
    Public Class DatosArtRef
        Public IDCliente As String
        Public Referencia As String
    End Class

    <Task()> Public Shared Function ObtenerArticuloRef(ByVal data As DatosArtRef, ByVal services As ServiceProvider) As String
        Dim strIDArticulo As String
        If Len(data.IDCliente) > 0 And Len(data.Referencia) > 0 Then
            Dim objFilter As New Filter
            objFilter.Add(New StringFilterItem("IDCliente", data.IDCliente))
            objFilter.Add(New StringFilterItem("RefCliente", data.Referencia))

            Dim dtReferencia As DataTable = New ArticuloCliente().Filter(objFilter)

            If Not IsNothing(dtReferencia) AndAlso dtReferencia.Rows.Count > 0 Then
                strIDArticulo = dtReferencia.Rows(0)("IDArticulo") & String.Empty
            End If
        End If

        Return strIDArticulo
    End Function
#End Region



End Class