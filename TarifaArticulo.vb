Public Class TarifaArticulo

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbTarifaArticulo"

#End Region

#Region "Eventos RegisterValidateTasks"

    ''' <summary>
    ''' Relación de tareas asociadas a la validación 
    ''' </summary>
    ''' <param name="validateProcess">Proceso en el que se registran las tareas de edición</param>
    ''' <remarks></remarks>
    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClavePrimaria)
    End Sub

    ''' <summary>
    ''' Comprobar que el código y la descripción no sean nulos
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDTarifa")) = 0 Then ApplicationService.GenerateError("El código de Tarifa es obligatorio.")
        If Length(data("IDArticulo")) = 0 Then ApplicationService.GenerateError("El campo Artículo es obligatorio.")
    End Sub

    ''' <summary>
    ''' Comprobar que no exista la clave
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New TarifaArticulo().SelOnPrimaryKey(data("IdTarifa"), data("IDArticulo"))
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("El Artículo introducido ya existe.")
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf TratarFecha)
        updateProcess.AddTask(Of DataRow)(AddressOf AplicarDecimales)
    End Sub

    <Task()> Public Shared Sub TratarFecha(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Modified Then data("FechaUltimaActualizacion") = Today.Date
    End Sub

    <Task()> Public Shared Sub AplicarDecimales(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDTarifa")) > 0 Then
            Dim Tarifas As EntityInfoCache(Of TarifaInfo) = services.GetService(Of EntityInfoCache(Of TarifaInfo))()
            Dim TarInfo As TarifaInfo = Tarifas.GetEntity(data("IDTarifa"))
            Dim IDMoneda As String
            If Length(TarInfo.IDMoneda) > 0 Then
                IDMoneda = TarInfo.IDMoneda
            Else
                Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
                IDMoneda = Monedas.MonedaA.ID
            End If
            Dim datosDec As New DataAplicarDecimalesMoneda(IDMoneda, data)
            ProcessServer.ExecuteTask(Of DataAplicarDecimalesMoneda)(AddressOf NegocioGeneral.AplicarDecimalesMoneda, datosDec, services)
        End If
    End Sub

#End Region

#Region "Eventos GetBusinessRules"

    ''' <summary>
    ''' Reglas de negocio. Lista de tareas asociadas a cambios.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>Solo se enstablece la lista en este punto no se ejecutan</remarks>
    Public Overrides Function GetBusinessRules() As Solmicro.Expertis.Engine.BE.BusinessRules
        Dim oBRL As New BusinessRules
        oBRL.Add("PVP", AddressOf CambioPVPPrecio)
        oBRL.Add("Precio", AddressOf CambioPVPPrecio)
        Return oBRL
    End Function

    ''' <summary>
    ''' Calcular PVP a partir de precio
    ''' </summary>
    ''' <param name="data">Estructura con la información necesaria</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub CambioPVPPrecio(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim dblFactor As Double
        If Not IsNumeric(data.Value) Then
            ApplicationService.GenerateError("Campo no numérico.")
        Else
            data.Current(data.ColumnName) = data.Value

            If Not IsNothing(data.Context) AndAlso data.Context.Contains("IdTarifa") AndAlso (data.Context.Contains("IdTipoIVA") Or data.Context.Contains("Factor")) Then
                If Nz(data.Context("Factor"), 0) > 0 Then
                    dblFactor = data.Context("Factor") / 100
                ElseIf Length(data.Context("IdTipoIVA")) > 0 Then
                    Dim TiposIVA As EntityInfoCache(Of TipoIvaInfo) = services.GetService(Of EntityInfoCache(Of TipoIvaInfo))()
                    Dim TIVAInfo As TipoIvaInfo = TiposIVA.GetEntity(data.Context("IDTipoIVA"))
                    dblFactor = TIVAInfo.Factor / 100
                End If

                Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
                Dim MonInfo As MonedaInfo
                If Length(data.Context("IDMoneda")) > 0 Then
                    MonInfo = Monedas.GetMoneda(data.Context("IDMoneda"))
                Else
                    MonInfo = Monedas.MonedaA
                End If

                If data.ColumnName = "PVP" And Nz(data.Context("TarifaPVP"), False) Then
                    data.Current("PVP") = xRound(data.Current("PVP"), MonInfo.NDecimalesImporte)
                    data.Current("Precio") = xRound(data.Current("PVP") / (1 + dblFactor), MonInfo.NDecimalesPrecio)
                End If
                If data.ColumnName = "Precio" And Not Nz(data.Context("TarifaPVP"), False) Then
                    ' Hallamos el precio de venta al público...
                    data.Current("Precio") = xRound(data.Current("Precio"), MonInfo.NDecimalesPrecio)
                    data.Current("PVP") = xRound(data.Current("Precio") * (1 + dblFactor), MonInfo.NDecimalesImporte)
                End If

            End If
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function PVP(ByVal strIDArticulo As String, ByVal services As ServiceProvider) As DataTable
        Dim blnDefault As Boolean
        Dim dblPVPA, dblPVPB As Double
        Dim M As New Moneda
        Dim dtMoneda As DataTable
        Dim dtPVP As New DataTable

        If Length(strIDArticulo) > 0 Then
            With dtPVP
                .Columns.Add("PVPA", GetType(Double))
                .Columns.Add("PVPB", GetType(Double))
                Dim rw As DataRow = .NewRow
                rw("PVPA") = 0
                rw("PVPB") = 0
                .Rows.Add(rw)
            End With

            Dim be As New Engine.BE.DataEngine
            Dim ofilter As New Filter
            ofilter.Add("IDArticulo", strIDArticulo)
            Dim dtTarifa As DataTable = be.Filter("vNegPVPTarifaArticulo", ofilter)

            blnDefault = True
            If Not dtTarifa Is Nothing AndAlso dtTarifa.Rows.Count > 0 Then
                For Each drTarifa As DataRow In dtTarifa.Rows
                    If drTarifa("Vigente") Then
                        If drTarifa("PVP") <> 0 Then
                            dtMoneda = M.Filter("CambioA,CambioB", "IDMoneda='" & drTarifa("IDMoneda") & "'")
                            If Not dtMoneda Is Nothing AndAlso dtMoneda.Rows.Count > 0 Then
                                dblPVPA = drTarifa("PVP") * dtMoneda.Rows(0)("CambioA")
                                dblPVPB = drTarifa("PVP") * dtMoneda.Rows(0)("CambioB")
                            End If
                        End If
                    End If
                Next
                blnDefault = False
            End If

            If blnDefault Then
                'PVPMinimo de Articulo
                Dim A As New Articulo
                Dim dtArticulo As DataTable = A.Filter("PVPMinimo", "IDArticulo='" & strIDArticulo & "'")
                If Not dtArticulo Is Nothing AndAlso dtArticulo.Rows.Count > 0 Then
                    dblPVPA = Nz(dtArticulo.Rows(0)("PVPMinimo"), 0)
                    If dblPVPA <> 0 Then
                        dtMoneda = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf Moneda.ObtenerMonedaA, Nothing, services)
                        If Not dtMoneda Is Nothing AndAlso dtMoneda.Rows.Count > 0 Then
                            dblPVPB = dblPVPA * dtMoneda.Rows(0)("CambioB")
                        End If
                    End If
                End If
            End If
            dtPVP.Rows(0)("PVPA") = dblPVPA
            dtPVP.Rows(0)("PVPB") = dblPVPB
        End If
        Return (dtPVP)
    End Function


    <Task()> Public Shared Sub ADDTarifaArticulo(ByVal data As DataTarifaArticulo, ByVal services As ServiceProvider)
        If Length(data.IDTarifa) > 0 AndAlso Length(data.IDArticulo) > 0 Then
            Dim dt As DataTable = New TarifaArticulo().SelOnPrimaryKey(data.IDTarifa, data.IDArticulo)
            If dt Is Nothing OrElse dt.Rows.Count = 0 Then
                Dim dr As DataRow = dt.NewRow
                dr("IDTarifa") = data.IDTarifa
                dr("IDArticulo") = data.IDArticulo
                dr("UDValoracion") = 1
                dr("FechaUltimaActualizacion") = Today.Date
                dt.Rows.Add(dr)
            End If
            dt.Rows(0)("Precio") = data.Precio
            Dim ClsTar As New TarifaArticulo
            ClsTar.Update(dt)
        End If
    End Sub

#End Region

End Class

<Serializable()> _
Public Class DataTarifaArticulo
    Public IDTarifa As String
    Public IDArticulo As String
    Public Precio As Double
End Class