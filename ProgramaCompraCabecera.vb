Public Class ProgramaCompraCabecera

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbProgramaCompraCabecera"

#End Region

#Region "Eventos RegisterAddNewTasks"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim StDatos As New Contador.DatosDefaultCounterValue(data, GetType(ProgramaCompraCabecera).Name, "IDPrograma")
        ProcessServer.ExecuteTask(Of Contador.DatosDefaultCounterValue)(AddressOf Contador.LoadDefaultCounterValue, StDatos, services)

        data("FechaPrograma") = Today.Date

        Dim cgu As New UsuarioCentroGestion.UsuarioCentroGestionInfo
        cgu = ProcessServer.ExecuteTask(Of UsuarioCentroGestion.UsuarioCentroGestionInfo, UsuarioCentroGestion.UsuarioCentroGestionInfo)(AddressOf UsuarioCentroGestion.ObtenerUsuarioCentroGestion, cgu, services)
        data("IDCentroGestion") = cgu.IDCentroGestion

        Dim strIDAlmacen As String = New Parametro().AlmacenPredeterminado
        If Length(data("IDCentroGestion")) > 0 Then
            strIDAlmacen = ProcessServer.ExecuteTask(Of String, String)(AddressOf Almacen.GetAlmacenCentroGestion, data("IDCentroGestion"), services)
        End If
        data("IDAlmacen") = strIDAlmacen
    End Sub


#End Region

#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim services As New ServiceProvider
        Dim oBrl As New BusinessRules
        oBrl = ProcessServer.ExecuteTask(Of BusinessRules, BusinessRules)(AddressOf ProcesoCompra.DetailBusinessRulesCab, oBrl, services)
        oBrl("IDProveedor") = AddressOf CambioProveedor
        Return oBrl
    End Function

    <Task()> Public Shared Sub CambioProveedor(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoCompra.CambioProveedor, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf AsignarDireccionPedido, data, services)
    End Sub

    <Task()> Public Shared Sub AsignarDireccionPedido(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim stDatosDirec As New ProveedorDireccion.DataDirecEnvio(data.Current("IDProveedor"), enumpdTipoDireccion.pdDireccionPedido)
        Dim dtDireccion As DataTable = ProcessServer.ExecuteTask(Of ProveedorDireccion.DataDirecEnvio, DataTable)(AddressOf ProveedorDireccion.ObtenerDireccionEnvio, stDatosDirec, services)
        If dtDireccion Is Nothing OrElse dtDireccion.Rows.Count = 0 Then
            data.Current("IDDireccionPedido") = System.DBNull.Value
        Else : data.Current("IDDireccionPedido") = dtDireccion.Rows(0)("IDDireccion")
        End If
    End Sub

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarDireccionPedido)
    End Sub

    <Task()> Public Shared Sub ComprobarDireccionPedido(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Modified Then
            If Length(data("IDDireccionPedido")) = 0 Then ApplicationService.GenerateError("No se ha podido obtener la Direccion de Envio.")
        End If
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClave)
    End Sub

    <Task()> Public Shared Sub AsignarClave(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDContador")) > 0 Then data("IDPrograma") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data("IDContador"), services)
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Serializable()> _
    Public Class DatosProgCompra
        Public Data As DataTable
        Public IDOperario As String
        Public IDCentroGestion As String
        Public IDContador As String
        Public AgruparProveedor As Boolean

        Public Sub New()
        End Sub
    End Class

    <Serializable()> _
    Public Class DatosProgCompraCab
        Public IDProveedor As String
        Public IDContador As String
        Public IDOperario As String
        Public IDCentroGestion As String
        Public IDAlmacen As String

        Public Sub New()
        End Sub
    End Class

    <Serializable()> _
    Public Class DatosProgCompraLin
        Public IDPrograma As String
        Public IDMoneda As String
        Public Dr As DataRow

        Public Sub New()
        End Sub
    End Class

    <Task()> Public Shared Function CrearProgramaCompra(ByVal data As DatosProgCompra, ByVal services As ServiceProvider) As DataTable
        If Not IsNothing(data) AndAlso data.Data.Rows.Count > 0 Then
            If Length(data.IDContador) > 0 Then
                Dim cont As New EntidadContador
                Dim f As New Filter
                f.Add(New StringFilterItem("Entidad", FilterOperator.Equal, "ProgramaCompraCabecera"))
                f.Add(New StringFilterItem("IDContador", FilterOperator.Equal, data.IDContador))
                Dim dtC As DataTable = cont.Filter(f)
                If IsNothing(dtC) OrElse dtC.Rows.Count = 0 Then
                    ApplicationService.GenerateError("El contador '|' no está definido para los Programas de Compra.", data.IDContador)
                End If
            End If

            Dim p As New Parametro
            Dim strAlmacen As String = p.AlmacenPredeterminado

            Dim dtPCC As DataTable = New ProgramaCompraCabecera().AddNew
            Dim PCL As New ProgramaCompraLinea
            Dim dtPCL As DataTable = PCL.AddNew
            Dim strIDProveedor As String

            Dim drNuevaCabecera, drNuevaLinea As DataRow

            data.Data.DefaultView.Sort = "IDProveedor"
            For Each dr As DataRow In data.Data.Select("", "IDProveedor")
                If data.AgruparProveedor Then
                    If strIDProveedor <> dr("IDProveedor") Then
                        strIDProveedor = dr("IDProveedor")
                        Dim StDatosCab As New DatosProgCompraCab
                        StDatosCab.IDProveedor = strIDProveedor
                        StDatosCab.IDContador = data.IDContador
                        StDatosCab.IDOperario = data.IDOperario
                        StDatosCab.IDCentroGestion = data.IDCentroGestion
                        StDatosCab.IDAlmacen = strAlmacen
                        drNuevaCabecera = ProcessServer.ExecuteTask(Of DatosProgCompraCab, DataRow)(AddressOf NuevaCabeceraPrograma, StDatosCab, services)
                        dtPCC.Rows.Add(drNuevaCabecera.ItemArray)
                        Dim dvlineas As New DataView(data.Data)
                        dvlineas.RowFilter = "IDProveedor=" & Quoted(dr("IDProveedor"))
                        For Each drv As DataRowView In dvlineas
                            Dim StDatosLin As New DatosProgCompraLin
                            StDatosLin.IDPrograma = drNuevaCabecera("IDPrograma")
                            StDatosLin.IDMoneda = drNuevaCabecera("IDMoneda")
                            StDatosLin.Dr = drv.Row
                            drNuevaLinea = ProcessServer.ExecuteTask(Of DatosProgCompraLin, DataRow)(AddressOf NuevaLineaPrograma, StDatosLin, services)
                            If IsNothing(dtPCL) Then dtPCL = New DataTable
                            dtPCL.Rows.Add(drNuevaLinea.ItemArray)
                        Next
                    End If
                Else
                    strIDProveedor = dr("IDProveedor")
                    Dim StDatosCab As New DatosProgCompraCab
                    StDatosCab.IDProveedor = strIDProveedor
                    StDatosCab.IDContador = data.IDContador
                    StDatosCab.IDOperario = data.IDOperario
                    StDatosCab.IDCentroGestion = data.IDCentroGestion
                    StDatosCab.IDAlmacen = strAlmacen
                    drNuevaCabecera = ProcessServer.ExecuteTask(Of DatosProgCompraCab, DataRow)(AddressOf NuevaCabeceraPrograma, StDatosCab, services)
                    dtPCC.Rows.Add(drNuevaCabecera.ItemArray)
                   
                    Dim StDatosLin As New DatosProgCompraLin
                    StDatosLin.IDPrograma = drNuevaCabecera("IDPrograma")
                    StDatosLin.IDMoneda = drNuevaCabecera("IDMoneda")
                    StDatosLin.Dr = dr
                    drNuevaLinea = ProcessServer.ExecuteTask(Of DatosProgCompraLin, DataRow)(AddressOf NuevaLineaPrograma, StDatosLin, services)
                    If IsNothing(dtPCL) Then dtPCL = New DataTable
                    dtPCL.Rows.Add(drNuevaLinea.ItemArray)
                End If
            Next
            BusinessHelper.UpdateTable(dtPCC)
            BusinessHelper.UpdateTable(dtPCL)
            Return dtPCC
        End If
    End Function

    <Task()> Public Shared Function NuevaCabeceraPrograma(ByVal data As DatosProgCompraCab, ByVal services As ServiceProvider) As DataRow
        Dim ClsProgCompraCab As New ProgramaCompraCabecera
        Dim drPrograma As DataRow = ClsProgCompraCab.AddNewForm.Rows(0)
        If Length(data.IDContador) = 0 Then data.IDContador = drPrograma("IDContador")
        drPrograma("IDPrograma") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data.IDContador, services)
        drPrograma("IDContador") = data.IDContador
        drPrograma("IDOperario") = data.IDOperario
        drPrograma("IDCentroGestion") = data.IDCentroGestion
        drPrograma("IDAlmacen") = data.IDAlmacen
        drPrograma("IDProveedor") = data.IDProveedor
        ClsProgCompraCab.ApplyBusinessRule("IDProveedor", drPrograma("IDProveedor"), drPrograma)
        Return drPrograma
    End Function

    <Task()> Public Shared Function NuevaLineaPrograma(ByVal data As DatosProgCompraLin, ByVal services As ServiceProvider) As DataRow
        Dim PCL As New ProgramaCompraLinea
        Dim drLinea As DataRow = PCL.AddNewForm.Rows(0)

        drLinea("IDLineaPrograma") = AdminData.GetAutoNumeric
        drLinea("IDPrograma") = data.IDPrograma
        drLinea("IDArticulo") = data.Dr("IDArticulo")

        Dim context As New BusinessData
        context("IDProveedor") = data.Dr("IDProveedor")
        PCL.ApplyBusinessRule("IDArticulo", drLinea("IDArticulo"), drLinea, context)

        If data.Dr.Table.Columns.Contains("IDAlmacen") Then
            If Not IsDBNull(data.Dr("IDAlmacen")) Then
                drLinea("IDAlmacen") = data.Dr("IDAlmacen")
            Else
                Dim StrIDAlmacen As String
                Dim FilArtAlm As New Filter
                Dim ClsArtAlm As New ArticuloAlmacen
                FilArtAlm.Add("IDArticulo", FilterOperator.Equal, data.Dr("IDArticulo"), FilterType.String)
                FilArtAlm.Add("Predeterminado", FilterOperator.Equal, 1, FilterType.Numeric)
                Dim DtArtAlm As DataTable = ClsArtAlm.Filter(FilArtAlm)
                If Not DtArtAlm Is Nothing AndAlso DtArtAlm.Rows.Count > 0 Then
                    StrIDAlmacen = DtArtAlm.Rows(0)("IDAlmacen")
                End If
                drLinea("IDAlmacen") = StrIDAlmacen
            End If
        End If
        drLinea("FechaEntrega") = data.Dr("FechaEntrega")

        If Length(data.Dr("IDUDCompra")) > 0 Then drLinea("IDUDMedida") = data.Dr("IDUDCompra")
        If Length(data.Dr("IDUDInterna")) > 0 Then drLinea("IDUDInterna") = data.Dr("IDUDInterna")
        If Length(data.Dr("IDTipoIva")) > 0 Then drLinea("IDTipoIva") = data.Dr("IDTipoIva")
        drLinea("QPrevista") = data.Dr("CantidadMarca1")

        Dim StDatos As New ArticuloUnidadAB.DatosFactorConversion
        StDatos.IDArticulo = drLinea("IDArticulo")
        StDatos.IDUdMedidaA = drLinea("IDUDMedida")
        StDatos.IDUdMedidaB = drLinea("IDUDInterna")
        drLinea("Factor") = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, StDatos, services)
        drLinea("QInterna") = drLinea("Factor") * drLinea("QPrevista")
        drLinea("QServida") = 0

        If data.Dr.Table.Columns.Contains("PrecioPrevMatA") Then
            drLinea("Precio") = data.Dr("PrecioPrevMatA")
        ElseIf data.Dr.Table.Columns.Contains("ImpPrevTrabajoA") Then
            drLinea("Precio") = data.Dr("ImpPrevTrabajoA")
        ElseIf data.Dr.Table.Columns.Contains("TasaPrevMatA") Then
            drLinea("Precio") = data.Dr("TasaPrevMatA")
        End If

        If data.Dr.Table.Columns.Contains("Dto1") Then drLinea("Dto1") = data.Dr("Dto1")
        If data.Dr.Table.Columns.Contains("Dto2") Then drLinea("Dto2") = data.Dr("Dto2")
        If data.Dr.Table.Columns.Contains("Dto3") Then drLinea("Dto3") = data.Dr("Dto3")

        context("Cantidad") = drLinea("QPrevista")
        context("IDMoneda") = data.IDMoneda
        PCL.ApplyBusinessRule("Precio", drLinea("Precio"), drLinea, context)

        If data.Dr.Table.Columns.Contains("UDValoracion") Then
            drLinea("UDValoracion") = data.Dr("UDValoracion")
        Else : drLinea("UDValoracion") = 1
        End If
        If data.Dr.Table.Columns.Contains("IDObra") Then drLinea("IDObra") = data.Dr("IDObra")
        If data.Dr.Table.Columns.Contains("IDTrabajo") Then drLinea("IDTrabajo") = data.Dr("IDTrabajo")
        If data.Dr.Table.Columns.Contains("IDLineaMaterial") Then drLinea("IDLineaMaterial") = data.Dr("IDLineaMaterial")
        If data.Dr.Table.Columns.Contains("IDMntoOTPrev") Then drLinea("IDMntoOTPrev") = data.Dr("IDMntoOTPrev")

        If data.Dr.Table.Columns.Contains("DescArticulo") Then
            If Length(data.Dr("DescArticulo")) > 0 Then drLinea("DescArticulo") = data.Dr("DescArticulo")
        End If
        Return drLinea
    End Function

#End Region

End Class