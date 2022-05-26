Public Class _ArticuloAlmacen
    Public Const IDArticulo As String = "IDArticulo"
    Public Const IDAlmacen As String = "IDAlmacen"
    Public Const StockFisico As String = "StockFisico"
    Public Const PuntoPedido As String = "PuntoPedido"
    Public Const LoteMinimo As String = "LoteMinimo"
    Public Const StockSeguridad As String = "StockSeguridad"
    Public Const PrecioMedioA As String = "PrecioMedioA"
    Public Const PrecioMedioB As String = "PrecioMedioB"
    Public Const StockMedio As String = "StockMedio"
    Public Const Rotacion As String = "Rotacion"
    Public Const Inventariado As String = "Inventariado"
    Public Const FechaUltimoInventario As String = "FechaUltimoInventario"
    Public Const FechaUltimoAjuste As String = "FechaUltimoAjuste"
    Public Const Predeterminado As String = "Predeterminado"
    Public Const GestionPuntoPedido As String = "GestionPuntoPedido"
    Public Const MarcaAuto As String = "MarcaAuto"
    Public Const PrecioFIFOFechaA As String = "PrecioFIFOFechaA"
    Public Const PrecioFIFOFechaB As String = "PrecioFIFOFechaB"
    Public Const PrecioFIFOMvtoA As String = "PrecioFIFOMvtoA"
    Public Const PrecioFIFOMvtoB As String = "PrecioFIFOMvtoB"
    Public Const FechaCalculo As String = "FechaCalculo"
    Public Const StockFechaCalculo As String = "StockFechaCalculo"
    Public Const IDArticuloGenerico As String = "IDArticuloGenerico"
    Public Const FechaUltimoMovimiento As String = "FechaUltimoMovimiento"
    Public Const FechaCreacionAudi As String = "FechaCreacionAudi"
    Public Const FechaModificacionAudi As String = "FechaModificacionAudi"
    Public Const UsuarioAudi As String = "UsuarioAudi"
End Class
<Serializable()> _
Public Class dataDisponibilidadArticuloAlquiler
    Public StockTotal As Double = 0        '//Stock Total del Articulo
    Public StockTotalDeposito As Double = 0  '//Stock Total del Articulo en Almacenes de Depósito
    Public StockNoDisponible As Double = 0   '//Stock No Disponible del Articulo
    Public EnMantenimiento As Double = 0     '//Artículos por almacén en Mantenimiento
    Public PendienteEnObra As Double = 0
    Public PedidoCompraPendiente As Double = 0
    Public PendienteActualizarCompra As Double = 0
    Public DisponibleTeorico As Double = 0
    Public DisponibleReal As Double = 0
End Class

<Serializable()> _
Public Class DataArtAlmAct
    Public IDArticulo As String
    Public IDAlmacen As String
    Public OperacionPP As enumPuntoPedido
    Public OperacionLote As enumLoteMinimo
    Public IDProcess As String
    Public PuntoPedido As Double
    Public LoteMinimo As Double
    Public Tipo As enumacsTipoArticulo
End Class
<Serializable()> _
Public Class DataArtAlm
    Public IDArticulo As String
    Public IDAlmacen As String
    Public IDCentroGestion As String
    Public AlmacenCentroGestion As Boolean
    Public dt As DataTable

    Public Sub New()
    End Sub
    Public Sub New(ByVal IDArticulo As String)
        Me.IDArticulo = IDArticulo
    End Sub
End Class

Public Class ArticuloAlmacen
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroArticuloAlmacen"

#Region "Eventos RegisterAddNewTasks"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarMarcaAuto, data, services)
    End Sub

#End Region

#Region "RegisterDeleteTasks"
    ''' <summary>
    ''' Relación de tareas asociadas al proceso de borrado
    ''' </summary>
    ''' <param name="deleteProcess">Proceso en el que se registran las tareas de borrado</param>
    ''' <remarks></remarks>
    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf DeleteArticuloAlmacen)
    End Sub
    ''' <summary>
    ''' Borrado de artículos
    ''' </summary>
    ''' <param name="data">Registro del artículo almacen a borrar</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub DeleteArticuloAlmacen(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Not IsNothing(data) Then
            If data("StockFisico") > 0 Then ApplicationService.GenerateError("No se pueden borrar Almacenes con Stock mayor de 0.")
            If data("Predeterminado") Then
                Dim dtPred As DataTable = New ArticuloAlmacen().Filter(New StringFilterItem("IDArticulo", data("IDArticulo")))
                If Not IsNothing(dtPred) AndAlso dtPred.Rows.Count > 1 Then
                    dtPred.Rows(0)("Predeterminado") = True

                    BusinessHelper.UpdateTable(dtPred)
                End If
            End If
        End If
    End Sub

#End Region

#Region "RegisterUpdateTasks"
    ''' <summary>
    ''' Relación de tareas asociadas a la edición 
    ''' </summary>
    ''' <param name="validateProcess">Proceso en el que se registran las tareas de edición</param>
    ''' <remarks></remarks>
    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarMarcaAuto)
        updateProcess.AddTask(Of DataRow)(AddressOf TratarArticuloPredeterminado)
        updateProcess.AddTask(Of DataRow)(AddressOf TratarArticuloGenerico)
    End Sub

    ''' <summary>
    ''' asignación de campo autonumérico MarcaAuto. Se utiliza para las consultas de marcas.
    ''' </summary>
    ''' <param name="validateProcess">Proceso en el que se registran las tareas de edición</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub AsignarMarcaAuto(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("MarcaAuto")) = 0 Then data("MarcaAuto") = AdminData.GetAutoNumeric
    End Sub

    ''' <summary>
    ''' Establecer el almacén predeterminado
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks>/remarks>
    <Task()> Public Shared Sub TratarArticuloPredeterminado(ByVal dr As DataRow, ByVal services As ServiceProvider)
        Dim f As New Filter
        f.Add("IDArticulo", FilterOperator.Equal, dr("IDArticulo"))
        f.Add("Predeterminado", FilterOperator.Equal, True)
        Dim dtAlm As DataTable = New ArticuloAlmacen().Filter(f)
        If IsNothing(dtAlm) OrElse dtAlm.Rows.Count = 0 Then
            ' No hay más almacenes para el articulo actual con lo cual será el predeterminado.
            dr("Predeterminado") = True
        Else
            ' Si el almacen ha sido marcado como predeterminado
            If dr("Predeterminado") Then
                If dr("IDAlmacen") <> dtAlm.Rows(0)("IDAlmacen") Then
                    dtAlm.Rows(0)("Predeterminado") = False
                    BusinessHelper.UpdateTable(dtAlm)
                End If
            ElseIf dr.RowState = DataRowState.Modified AndAlso dr("Predeterminado") <> dr("Predeterminado", DataRowVersion.Original) Then
                dr("Predeterminado") = True
            End If
        End If
    End Sub

    ''' <summary>
    ''' Establecer el artíclo genérico
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks>/remarks>
    <Task()> Public Shared Sub TratarArticuloGenerico(ByVal dr As DataRow, ByVal services As ServiceProvider)
        If Length(dr("IDArticuloGenerico")) > 0 Then
            Dim f As New Filter
            f.Add(New StringFilterItem("IDArticulo", dr("IDArticuloGenerico")))
            Dim dtArt As DataTable = New BE.DataEngine().Filter("vNegCaractArticulo", f)
            If Not dtArt Is Nothing AndAlso dtArt.Rows.Count > 0 AndAlso Not dtArt.Rows(0)("Generico") Then
                ApplicationService.GenerateError("El artículo genérico | no es válido.", "'" & dr("IDArticuloGenerico") & "'")
            End If
        End If
    End Sub
#End Region

#Region "RegisterValidateTasks"
    ''' <summary>
    ''' Relación de tareas asociadas a la validación 
    ''' </summary>
    ''' <param name="validateProcess">Proceso en el que se registran las tareas de edición</param>
    ''' <remarks></remarks>
    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidaDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)

    End Sub
    ''' <summary>
    ''' Comprobar que el artículo tenga tenga descripción, tipo y familia
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidaDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDArticulo")) = 0 Then ApplicationService.GenerateError("El Artículo es un dato obligatorio.")
        If Length(data("IDAlmacen")) = 0 Then ApplicationService.GenerateError("El Almacén es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New ArticuloAlmacen().SelOnPrimaryKey(data("IDArticulo"), data("IDAlmacen"))
            If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("El Almacén '|' ya está asignado en este artículo.", data("IDAlmacen"))
            End If
        End If
    End Sub
#End Region

#Region "Funciones Públicas- Punto Pedido y Lote"
    <Task()> Public Shared Sub ActualizacionDePuntoPedido(ByVal data As DataArtAlmAct, ByVal services As ServiceProvider)
        Dim _filter As New Filter
        Dim dt As DataTable
        Dim dtArticulosAlmacen As DataTable
        Dim i As Integer = 0
        Dim _artAlm As New ArticuloAlmacen
        _filter.Add("idprocess", FilterOperator.Equal, data.IDProcess)
        dt = New BE.DataEngine().Filter("vFrmCIActualizarPuntoPedido", _filter, , "idarticulo,idalmacen")
        Dim _filterAND As New Filter
        Dim _filterOR As New Filter(FilterUnionOperator.Or)
        For Each dr As DataRow In dt.Select
            _filterAND = New Filter
            _filterAND.Add("IDArticulo", dr("IDArticulo"))
            _filterAND.Add("IDAlmacen", dr("IDAlmacen"))
            _filterOR.Add(_filterAND)
            _filterAND = Nothing
        Next
        dtArticulosAlmacen = _artAlm.Filter(_filterOR, "idarticulo, idalmacen")
        i = 0
        For Each drArticuloAlm As DataRow In dtArticulosAlmacen.Select()
            Select Case data.OperacionPP
                Case enumPuntoPedido.enumPPIndividual
                    If Length(dt.Rows(i).Item("CantidadMarca1")) > 0 AndAlso dt.Rows(i)("CantidadMarca1") <> 0 Then
                        drArticuloAlm("PuntoPedido") = dt.Rows(i)("CantidadMarca1")
                    End If
                Case enumPuntoPedido.enumPPMasivo
                    drArticuloAlm("PuntoPedido") = data.PuntoPedido
                Case enumPuntoPedido.enumPPCompra
                    data.IDArticulo = dt.Rows(i)("IDArticulo")
                    data.IDAlmacen = dt.Rows(i)("IDAlmacen")
                    data.Tipo = enumacsTipoArticulo.acsCompra
                    drArticuloAlm("PuntoPedido") = ProcessServer.ExecuteTask(Of DataArtAlmAct, Double)(AddressOf CalcularPuntoPedido, data, services)
                Case enumPuntoPedido.enumPPFabrica
                    data.IDArticulo = dt.Rows(i)("IDArticulo")
                    data.IDAlmacen = dt.Rows(i)("IDAlmacen")
                    data.Tipo = enumacsTipoArticulo.acsFabrica
                    drArticuloAlm("PuntoPedido") = ProcessServer.ExecuteTask(Of DataArtAlmAct, Double)(AddressOf CalcularPuntoPedido, data, services)
            End Select
            i += 1
        Next
        BusinessHelper.UpdateTable(dtArticulosAlmacen)
    End Sub
    <Task()> Public Shared Function CalcularPuntoPedido(ByVal data As DataArtAlmAct, ByVal services As ServiceProvider) As Double
        Dim dblConsumoDiario As Double
        Dim _filter As New Filter
        If Len(data.IDArticulo) > 0 And Len(data.IDAlmacen) > 0 Then
            _filter.Add("Idarticulo", data.IDArticulo)
            _filter.Add("IdAlmacen", data.IDAlmacen)
            Dim dtPuntoPedido As DataTable = New BE.DataEngine().Filter("vCTLCICalculoPuntoPedidoDiario", _filter)
            If Not dtPuntoPedido Is Nothing AndAlso dtPuntoPedido.Rows.Count > 0 Then
                Select Case data.Tipo
                    Case enumacsTipoArticulo.acsCompra
                        dblConsumoDiario = Nz(dtPuntoPedido.Rows(0)("ConsumoDiarioCompra"), 0)
                    Case enumacsTipoArticulo.acsFabrica
                        dblConsumoDiario = Nz(dtPuntoPedido.Rows(0)("ConsumoDiarioFabrica"), 0)
                End Select
            End If
        End If
        Return dblConsumoDiario
    End Function
    <Task()> Public Shared Sub ActualizacionDeLoteMinimo(ByVal data As DataArtAlmAct, ByVal services As ServiceProvider)
        Dim _filter As New Filter
        Dim dt As DataTable
        Dim dtArticulosAlmacen As DataTable
        Dim i As Integer = 0
        Dim _artAlm As New ArticuloAlmacen
        _filter.Add("idprocess", FilterOperator.Equal, data.IDProcess)
        dt = New BE.DataEngine().Filter("vFrmCIActualizarLoteMnimo", _filter, , "idarticulo,idalmacen")
        Dim _filterAND As New Filter
        Dim _filterOR As New Filter(FilterUnionOperator.Or)
        For Each dr As DataRow In dt.Select
            _filterAND = New Filter
            _filterAND.Add("IDArticulo", dr("IDArticulo"))
            _filterAND.Add("IDAlmacen", dr("IDAlmacen"))
            _filterOR.Add(_filterAND)
            _filterAND = Nothing
        Next
        dtArticulosAlmacen = _artAlm.Filter(_filterOR, "idarticulo, idalmacen")
        i = 0
        For Each drArticuloAlm As DataRow In dtArticulosAlmacen.Select()
            Select Case data.OperacionLote
                Case enumLoteMinimo.enumLoteIndividual
                    If Length(dt.Rows(i).Item("CantidadMarca1")) > 0 AndAlso dt.Rows(i)("CantidadMarca1") <> 0 Then
                        drArticuloAlm("LoteMinimo") = dt.Rows(i)("CantidadMarca1")
                    End If
                Case enumLoteMinimo.enumLoteMasivo
                    drArticuloAlm("LoteMinimo") = data.LoteMinimo
            End Select
            i += 1
        Next
        BusinessHelper.UpdateTable(dtArticulosAlmacen)
    End Sub
    <Task()> Public Shared Function PlanificacionPorPuntoPedido(ByVal FilPlanif As Filter, ByVal services As ServiceProvider) As DataTable
        Dim DtPlanAlmacen As DataTable = New BE.DataEngine().Filter("VCtlCIPuntoPedidoArticuloAlmacen", "*", "")
        Dim DtPedido As DataTable = New BE.DataEngine().Filter("VCtlCIPuntoPedidoCompraLinea", "*", "")
        Dim DtPrograma As DataTable = New BE.DataEngine().Filter("VCtlCIPuntoProgramaCompraLinea", "*", "")
        Dim DtOF As DataTable = New BE.DataEngine().Filter("VCtlCIPuntoPedidoOF", "*", "")

        If (Not DtPlanAlmacen Is Nothing) AndAlso (Not DtPedido Is Nothing) _
        AndAlso (Not DtPrograma Is Nothing) AndAlso (Not DtOF Is Nothing) Then
            DtPlanAlmacen.Columns.Add("IDOrdenLinea", GetType(Integer))
            For Each Dr As DataRow In DtPlanAlmacen.Select
                Dim StrFiltro As String = "IDArticulo='" & Dr("IDArticulo") & "' AND IDAlmacen= '" & Dr("IDAlmacen") & "'"
                Dim DrPed() As DataRow = DtPedido.Select(StrFiltro)
                Dim DrProg() As DataRow = DtPrograma.Select(StrFiltro)
                Dim DrOF() As DataRow = DtOF.Select(StrFiltro)
                Dr("EnCurso") = 0
                Dr("Diferencia") = 0
                If DrPed.Length > 0 Then
                    For Each DrPedLin As DataRow In DrPed
                        Dr("EnCurso") += DrPedLin("EnCurso")
                        If Length(DrPedLin("IDOrdenLinea")) > 0 Then Dr("IDOrdenLinea") = DrPedLin("IDOrdenLinea")
                    Next
                End If
                If DrProg.Length > 0 Then Dr("EnCurso") += DrProg(0)("EnCurso")
                If DrOF.Length > 0 Then Dr("EnCurso") += DrOF(0)("EnCurso")
                Dr("Diferencia") = Dr("StockFisico") + Dr("EnCurso") - Dr("PuntoPedido")
                Dr("DiferenciaSS") = Dr("StockFisico") + Dr("EnCurso") - Dr("StockSeguridad")
                Dr("FechaInicio") = Today.Date
                Dr("FechaFin") = Today.Date
            Next
        End If
        If Not FilPlanif Is Nothing Then
            Dim DtNuevo As DataTable = DtPlanAlmacen.Clone
            Dim WherePlanificacion As String = FilPlanif.Compose(New AdoFilterComposer)
            Dim DrNew() As DataRow = DtPlanAlmacen.Select(WherePlanificacion)
            If DrNew.Length > 0 Then
                For Each Dr As DataRow In DrNew
                    DtNuevo.ImportRow(Dr)
                Next
            End If
            Return DtNuevo
            'Else
            'Return DtPlanAlmacen
        Else
            Return DtPlanAlmacen
        End If
    End Function
#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function AltaDeArticulo(ByVal data As DataArtAlm, ByVal services As ServiceProvider) As Integer
        Dim intAltas As Integer
        If Not data.dt Is Nothing AndAlso data.dt.Rows.Count > 0 Then
            Dim DtAux As DataTable = New Almacen().SelOnPrimaryKey(data.IDAlmacen)
            If IsNothing(DtAux) OrElse DtAux.Rows.Count = 0 Then
                ApplicationService.GenerateError("El Almacén no existe.")
            Else
                Dim DtAltas As DataTable = New ArticuloAlmacen().AddNew()
                Dim AA As New ArticuloAlmacen()
                For Each Dr As DataRow In data.dt.Select
                    DtAux = AA.SelOnPrimaryKey(Dr("IDArticulo"), data.IDAlmacen)
                    If Not DtAux Is Nothing AndAlso DtAux.Rows.Count > 0 Then
                        Dim DrNew As DataRow = DtAltas.NewRow
                        DrNew("IDArticulo") = Dr("IDArticulo").Value
                        DrNew("IDAlmacen") = data.IDAlmacen
                        DrNew("StockFisico") = 0 : DrNew("PuntoPedido") = 0 : DrNew("LoteMinimo") = 0
                        DrNew("StockSeguridad") = 0 : DrNew("PrecioMedioA") = 0 : DrNew("PrecioMedioB") = 0
                        DrNew("StockMedio") = 0 : DrNew("Rotacion") = 0 : DrNew("Inventariado") = 0
                        DrNew("FechaUltimoInventario") = System.DBNull.Value
                        DrNew("FechaUltimoAjuste") = System.DBNull.Value
                        DrNew("Predeterminado") = 0
                        DtAltas.Rows.Add(DrNew)
                    End If
                    intAltas += 1
                Next
                If DtAltas.Rows.Count > 0 Then AA.Update(DtAltas)
            End If
        Else : ApplicationService.GenerateError("No hay lineas seleccionadas.")
        End If
        Return intAltas
    End Function

    <Task()> Public Shared Sub AltaArticuloAlmacen(ByVal data As DataArticuloAlmacen, ByVal services As ServiceProvider)
        If Not data Is Nothing Then
            Dim DtAux As DataTable
            Dim DtAltas As DataTable = New ArticuloAlmacen().AddNew()
            Dim AA As New ArticuloAlmacen()
            DtAux = AA.SelOnPrimaryKey(data.IDArticulo, data.IDAlmacen)
            If Not DtAux Is Nothing AndAlso DtAux.Rows.Count = 0 Then
                Dim DrNew As DataRow = DtAltas.NewRow
                DrNew("IDArticulo") = data.IDArticulo
                DrNew("IDAlmacen") = data.IDAlmacen
                DrNew("StockFisico") = 0 : DrNew("PuntoPedido") = 0 : DrNew("LoteMinimo") = 0
                DrNew("StockSeguridad") = 0 : DrNew("PrecioMedioA") = 0 : DrNew("PrecioMedioB") = 0
                DrNew("StockMedio") = 0 : DrNew("Rotacion") = 0 : DrNew("Inventariado") = 0
                DrNew("FechaUltimoInventario") = System.DBNull.Value
                DrNew("FechaUltimoAjuste") = System.DBNull.Value
                DrNew("Predeterminado") = 0
                DtAltas.Rows.Add(DrNew)
            End If
            If DtAltas.Rows.Count > 0 Then AA.Update(DtAltas)
        End If
    End Sub



    <Task()> Public Shared Function AddAlmacenPredeterminadoArticulo(ByVal strIDArticulo As String, ByVal services As ServiceProvider) As DataTable
        Dim AE As New ArticuloAlmacen
        Dim dtNewAlmacen As DataTable = AE.AddNewForm
        dtNewAlmacen.Rows(0)("IDArticulo") = strIDArticulo
        dtNewAlmacen.Rows(0)("IDAlmacen") = New Parametro().AlmacenPredeterminado()
        dtNewAlmacen.Rows(0)("StockFisico") = 0
        dtNewAlmacen.Rows(0)("PuntoPedido") = 0
        dtNewAlmacen.Rows(0)("LoteMinimo") = 0
        dtNewAlmacen.Rows(0)("StockSeguridad") = 0
        dtNewAlmacen.Rows(0)("PrecioMedioA") = 0
        dtNewAlmacen.Rows(0)("PrecioMedioB") = 0
        dtNewAlmacen.Rows(0)("StockMedio") = 0
        dtNewAlmacen.Rows(0)("Rotacion") = 0
        dtNewAlmacen.Rows(0)("Inventariado") = False
        dtNewAlmacen.Rows(0)("Predeterminado") = False
        dtNewAlmacen.Rows(0)("GestionPuntoPedido") = False
        dtNewAlmacen.Rows(0)("MarcaAuto") = AdminData.GetAutoNumeric
        dtNewAlmacen.Rows(0)("PrecioFIFOFechaA") = 0
        dtNewAlmacen.Rows(0)("PrecioFIFOFechaB") = 0
        dtNewAlmacen.Rows(0)("PrecioFIFOMvtoA") = 0
        dtNewAlmacen.Rows(0)("PrecioFIFOMvtoB") = 0
        dtNewAlmacen.Rows(0)("StockFechaCalculo") = 0
        Return AE.Update(dtNewAlmacen)
    End Function
    <Task()> Public Shared Sub CambioAlmacenPredeterminado(ByVal data As DataArtAlm, ByVal services As ServiceProvider)
        For Each drArticulo As DataRow In data.dt.Rows
            If Length(drArticulo("IDArticulo")) > 0 Then
                Dim AE As New ArticuloAlmacen
                Dim dtAlmacen As DataTable = AE.SelOnPrimaryKey(drArticulo("IDArticulo"), data.IDAlmacen)
                If Not IsNothing(dtAlmacen) AndAlso dtAlmacen.Rows.Count > 0 Then
                    dtAlmacen.Rows(0)("Predeterminado") = True
                Else
                    dtAlmacen = AE.AddNewForm()
                    dtAlmacen.Rows(0)("IDArticulo") = drArticulo("IDArticulo")
                    dtAlmacen.Rows(0)("IDAlmacen") = data.IDAlmacen
                    dtAlmacen.Rows(0)("Predeterminado") = True
                End If

                AE.Update(dtAlmacen)
            End If
        Next
    End Sub
    <Task()> Public Shared Function AlmacenPredeterminadoArticulo(ByVal data As DataArtAlm, ByVal services As ServiceProvider) As String

        Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
        Dim StrIDAlmacen As String
        'Comprobamos el Valor del Parámetro ALMCC que valor tiene configurado.
        Select Case AppParams.AlmacenCentroGestionActivo
            Case True
                'Evualuamos si hemos pasado un centro de gestión al proceso
                If Length(data.IDCentroGestion) > 0 Then
                    'Buscamos el Almacen predeterminado del Centro de Gestión que hemos pasado
                    StrIDAlmacen = ProcessServer.ExecuteTask(Of String, String)(AddressOf Almacen.GetAlmacenCentroGestion, data.IDCentroGestion, services)
                    data.IDAlmacen = StrIDAlmacen
                    If Length(data.IDAlmacen) = 0 Then
                        Dim FilAlm As New Filter
                        FilAlm.Add("IDArticulo", FilterOperator.Equal, data.IDArticulo, FilterType.String)
                        FilAlm.Add("Predeterminado", FilterOperator.NotEqual, 0, FilterType.Numeric)
                        Dim DtArtAlm As DataTable = New ArticuloAlmacen().Filter(FilAlm)
                        'En caso de no tener almacen por Centro de Gestión, buscamos por el
                        'predeterminado del Articulo
                        If Not DtArtAlm Is Nothing AndAlso DtArtAlm.Rows.Count > 0 Then
                            If Length(DtArtAlm.Rows(0)("IDAlmacen")) > 0 Then
                                StrIDAlmacen = DtArtAlm.Rows(0)("IDAlmacen")
                                data.IDAlmacen = StrIDAlmacen
                            Else
                                'Por ultimo si no encontramos por ninguno, miramos por el
                                'Parámetro de Almacen Predeterminado.
                                data.IDAlmacen = AppParams.Almacen
                            End If
                        Else
                            data.IDAlmacen = AppParams.Almacen
                        End If
                    End If
                Else
                    'Si no hemos pasado un centro de gestión, buscamos nuestro centro
                    'de gestión de nuestro usuario
                    Dim cgu As New UsuarioCentroGestion.UsuarioCentroGestionInfo
                    cgu = ProcessServer.ExecuteTask(Of UsuarioCentroGestion.UsuarioCentroGestionInfo, UsuarioCentroGestion.UsuarioCentroGestionInfo)(AddressOf UsuarioCentroGestion.ObtenerUsuarioCentroGestion, cgu, services)
                    data.IDCentroGestion = cgu.IDCentroGestion


                    If Length(data.IDCentroGestion) > 0 Then
                        'Si tengo un centro de gestión busco el almacén predeterminado de
                        'ese centro de gestión
                        data.IDAlmacen = ProcessServer.ExecuteTask(Of String, String)(AddressOf Almacen.GetAlmacenCentroGestion, data.IDCentroGestion, services)
                        If Length(data.IDAlmacen) = 0 Then
                            Dim FilAlm As New Filter
                            FilAlm.Add("IDArticulo", FilterOperator.Equal, data.IDArticulo, FilterType.String)
                            FilAlm.Add("Predeterminado", FilterOperator.NotEqual, 0, FilterType.Numeric)
                            Dim DtArtAlm As DataTable = New ArticuloAlmacen().Filter(FilAlm)
                            'En caso de no tener almacen por Centro de Gestión, buscamos por el
                            'predeterminado del Articulo
                            If Not DtArtAlm Is Nothing AndAlso DtArtAlm.Rows.Count > 0 Then
                                If Length(DtArtAlm.Rows(0)("IDAlmacen")) > 0 Then
                                    data.IDAlmacen = DtArtAlm.Rows(0)("IDAlmacen")
                                Else
                                    'Por ultimo si no encontramos por ninguno, miramos por el
                                    'Parámetro de Almacen Predeterminado.
                                    data.IDAlmacen = AppParams.Almacen
                                End If
                            Else
                                'Por ultimo si no encontramos por ninguno, miramos por el
                                'Parámetro de Almacen Predeterminado.
                                data.IDAlmacen = AppParams.Almacen
                            End If
                        End If
                    Else
                        Dim FilAlm As New Filter
                        FilAlm.Add("IDArticulo", FilterOperator.Equal, data.IDArticulo, FilterType.String)
                        FilAlm.Add("Predeterminado", FilterOperator.NotEqual, 0, FilterType.Numeric)
                        Dim DtArtAlm As DataTable = New ArticuloAlmacen().Filter(FilAlm)
                        'En caso de no tener almacen por Centro de Gestión, buscamos por el
                        'predeterminado del Articulo
                        If Not DtArtAlm Is Nothing AndAlso DtArtAlm.Rows.Count > 0 Then
                            If Length(DtArtAlm.Rows(0)("IDAlmacen")) > 0 Then
                                StrIDAlmacen = DtArtAlm.Rows(0)("IDAlmacen")
                            Else
                                'Por ultimo si no encontramos por ninguno, miramos por el
                                'Parámetro de Almacen Predeterminado.
                                StrIDAlmacen = AppParams.Almacen
                            End If
                        Else
                            'Por ultimo si no encontramos por ninguno, miramos por el
                            'Parámetro de Almacen Predeterminado.
                            StrIDAlmacen = AppParams.Almacen
                        End If
                    End If
                End If
            Case False
                Dim BlnTieneAlm As Boolean = False
                Dim FilAlm As New Filter
                FilAlm.Add("IDArticulo", FilterOperator.Equal, data.IDArticulo, FilterType.String)
                FilAlm.Add("Predeterminado", FilterOperator.NotEqual, 0, FilterType.Numeric)
                Dim DtArtAlm As DataTable = New ArticuloAlmacen().Filter(FilAlm)
                'En caso de no tener almacen por Centro de Gestión, buscamos por el
                'predeterminado del Articulo
                If Not DtArtAlm Is Nothing AndAlso DtArtAlm.Rows.Count > 0 Then
                    If Length(DtArtAlm.Rows(0)("IDAlmacen")) > 0 Then
                        StrIDAlmacen = DtArtAlm.Rows(0)("IDAlmacen")
                        data.IDAlmacen = StrIDAlmacen
                        BlnTieneAlm = True
                    End If
                End If
                'Evualuamos si hemos pasado un centro de gestión al proceso
                If Length(data.IDCentroGestion) > 0 AndAlso Not BlnTieneAlm Then
                    'Buscamos el Almacen predeterminado del Centro de Gestión que hemos pasado
                    StrIDAlmacen = ProcessServer.ExecuteTask(Of String, String)(AddressOf Almacen.GetAlmacenCentroGestion, data.IDCentroGestion, services)
                    If Length(StrIDAlmacen) = 0 Then
                        'Por ultimo si no encontramos por ninguno, miramos por el
                        'Parámetro de Almacen Predeterminado.
                        data.IDAlmacen = AppParams.Almacen
                    Else
                        data.IDAlmacen = StrIDAlmacen
                    End If
                ElseIf Not BlnTieneAlm Then
                    'Si no hemos pasado un centro de gestión, buscamos nuestro centro
                    'de gestión de nuestro usuario
                    Dim cgu As New UsuarioCentroGestion.UsuarioCentroGestionInfo
                    cgu = ProcessServer.ExecuteTask(Of UsuarioCentroGestion.UsuarioCentroGestionInfo, UsuarioCentroGestion.UsuarioCentroGestionInfo)(AddressOf UsuarioCentroGestion.ObtenerUsuarioCentroGestion, cgu, services)
                    data.IDCentroGestion = cgu.IDCentroGestion

                    If Length(data.IDCentroGestion) > 0 Then
                        'Buscamos el Almacen predeterminado del Centro de Gestión que hemos pasado
                        StrIDAlmacen = ProcessServer.ExecuteTask(Of String, String)(AddressOf Almacen.GetAlmacenCentroGestion, data.IDCentroGestion, services)
                        If Length(StrIDAlmacen) = 0 Then
                            'Por ultimo si no encontramos por ninguno, miramos por el
                            'Parámetro de Almacen Predeterminado.
                            data.IDAlmacen = AppParams.Almacen
                        End If
                    Else
                        'Por ultimo si no encontramos por ninguno, miramos por el
                        'Parámetro de Almacen Predeterminado.
                        data.IDAlmacen = AppParams.Almacen
                    End If
                End If
        End Select
        Return data.IDAlmacen
    End Function
    <Task()> Public Shared Function ObtenerStockTotal(ByVal f As Filter, ByVal services As ServiceProvider) As DataTable
        Return New BE.DataEngine().Filter(ArticuloAlmacen.cnEntidad, f, "SELECT SUM(StockFisico) AS StockTotal")
    End Function
    <Task()> Public Shared Function CalcularResumenDisponibilidadAlquiler(ByVal IDArticulo As String, ByVal services As ServiceProvider) As dataDisponibilidadArticuloAlquiler
        Dim f As New Filter
        f.Add(New StringFilterItem("IDArticulo", IDArticulo))

        Dim dataDisponibilidadAlquiler As New dataDisponibilidadArticuloAlquiler
        Dim DE As New BE.DataEngine
        '//Stock Total del artículo
        Dim dt As DataTable = DE.Filter("VfrmCiStockTotal", f, "SUM(StockTotal) AS StockTotal")
        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            dataDisponibilidadAlquiler.StockTotal = Nz(dt.Rows(0)("StockTotal"), 0)
        End If

        '//Stock Total del Articulo en Almacenes de Depósito
        dt = DE.Filter("VfrmCiStockTotalDeposito", f, "SUM(StockFisico) AS StockTotal")
        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            dataDisponibilidadAlquiler.StockTotalDeposito = Nz(dt.Rows(0)("StockTotal"), 0)
        End If

        '//Stock No Disponible del Articulo
        dt = DE.Filter("VFrmAlquilerStockNodisponible", f, "SUM(StockNoDisponible) AS StockNTotal")
        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            dataDisponibilidadAlquiler.StockNoDisponible = Nz(dt.Rows(0)("StockNTotal"), 0)
        End If

        '//Artículos por almacén en Mantenimiento (Están en una OT En Curso o Lanzada)
        dt = DE.Filter("VAlquilerCIDispAlquilerEnMantenimiento", f, "SUM(EnMantenimiento) AS MantenimientoTotal")
        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            dataDisponibilidadAlquiler.EnMantenimiento = Nz(dt.Rows(0)("MantenimientoTotal"), 0)
        End If

        dt = DE.Filter("vAlquilerCIDisponibilidadAlquiler", f, "SUM(QPendienteObra) AS PendienteTotal")
        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            dataDisponibilidadAlquiler.PendienteEnObra = Nz(dt.Rows(0)("PendienteTotal"), 0)
        End If

        dt = DE.Filter("VCtlCIDisponibilidadPedidoCompra", f, "SUM(Pendiente) AS PendienteTotal")
        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            dataDisponibilidadAlquiler.PedidoCompraPendiente = Nz(dt.Rows(0)("PendienteTotal"), 0)
        End If

        dt = DE.Filter("vAlquilerCIDispAlquilerPteActualizarCompra", f, "SUM(PteActualizarCompra) AS ActualizarTotal")
        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            dataDisponibilidadAlquiler.PendienteActualizarCompra = Nz(dt.Rows(0)("ActualizarTotal"), 0)
        End If

        dataDisponibilidadAlquiler.DisponibleTeorico = dataDisponibilidadAlquiler.StockTotal - dataDisponibilidadAlquiler.StockTotalDeposito - dataDisponibilidadAlquiler.StockNoDisponible
        dataDisponibilidadAlquiler.DisponibleReal = dataDisponibilidadAlquiler.DisponibleTeorico + (dataDisponibilidadAlquiler.PedidoCompraPendiente + dataDisponibilidadAlquiler.PendienteActualizarCompra) - dataDisponibilidadAlquiler.PendienteEnObra

        Return dataDisponibilidadAlquiler
    End Function

#End Region

#Region "GetStock"
    Public Function getStock(ByVal almacen As String, ByVal articulo As String) As String
        Dim strSQL As String = "SELECT * FROM tbMaestroArticuloAlmacen WHERE IDArticulo = '" & articulo & "' AND IDAlmacen = '" & almacen & "'"
        Dim tb As DataTable = AdminData.GetData(strSQL)
        Dim stock As String = "0"
        For Each dr As DataRow In tb.Rows
            stock = dr("StockFisico")
            'MsgBox(stock)
        Next
        Return stock
    End Function
#End Region

    Public Sub ActualizaStock(ByVal stock As String, ByVal idArticulo As String)
        Dim sql As String
        sql = "UPDATE tbMaestroArticuloAlmacen"
        sql &= " SET StockFisico='" & stock & "' "
        sql &= " WHERE IDArticulo='" & idArticulo & "'"
        Try
            AdminData.Execute(sql)
        Catch ex As Exception
            ApplicationService.GenerateError(ex.ToString & ":error")
        End Try
    End Sub

End Class

