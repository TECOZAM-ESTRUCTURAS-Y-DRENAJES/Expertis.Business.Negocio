Public Class Activo

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroActivo"

#End Region

#Region "Eventos RegisterAddNewTasks"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    ''' <summary>
    ''' Método para rellenar valores por defecto, al crear un nuevo registro.
    ''' </summary>
    ''' <param name="data">Registro de Activo</param>
    ''' <param name="services">Objeto para compartir información</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim AppParamsActivo As ParametroActivo = services.GetService(Of ParametroActivo)()
        Dim Estado As String = AppParamsActivo.EstadoActivoPredeterminado
        'Dim Clase As String = AppParamsActivo.ClaseActivoPredeterminado
        'Dim Categoria As String = AppParamsActivo.CategoriaActivoPredeterminado
        ' Dim Zona As String = AppParamsActivo.ZonaActivoPredeterminado
        'Dim CentroCoste As String = AppParamsActivo.CentroCosteActivoPredeterminado

        If Len(Estado) > 0 Then
            data("IDEstadoActivo") = Estado
        End If
        data("IDCategoriaActivo") = "00"
        data("IDCentroCoste") = "00"
        data("IDClaseActivo") = "00"
        'If Len(Clase) > 0 Then
        '    data("IDClaseActivo") = Clase
        'End If
        'If Len(Categoria) > 0 Then
        '    data("IDCategoriaActivo") = Categoria
        'End If
        'If Len(Zona) > 0 Then
        '    data("IDZona") = Zona
        'End If
        'If Len(CentroCoste) > 0 Then
        '    data("IDCentroCoste") = "00"
        'End If
        data("FechaAlta") = Today
        data("FechaEstado") = Today

        Dim StDatos As New Contador.DatosDefaultCounterValue
        StDatos.Row = data
        StDatos.EntityName = "Activo"
        StDatos.FieldName = "IDActivo"
        ProcessServer.ExecuteTask(Of Contador.DatosDefaultCounterValue)(AddressOf Contador.LoadDefaultCounterValue, StDatos, services)
    End Sub

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    ''' <summary>
    ''' Método para validar datos obligatorios del Activo.
    ''' </summary>
    ''' <param name="data">Registro de Activo</param>
    ''' <param name="services">Objeto para compartir información</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim Operarios As EntityInfoCache(Of OperarioInfo) = services.GetService(Of EntityInfoCache(Of OperarioInfo))()
        Dim OpInfo As OperarioInfo = Operarios.GetEntity(AdminData.GetSessionInfo.UserID)
        If OpInfo Is Nothing OrElse Length(OpInfo.IDOperario) = 0 Then ApplicationService.GenerateError("El Usuario actual no tiene asociado ningún Operario. El Operario es obligatorio.")

        If Length(data("DescActivo")) = 0 Then ApplicationService.GenerateError("La Descripción del Activo es obligatoria.")
        If Length(data("IDCategoriaActivo")) = 0 Then ApplicationService.GenerateError("La Categoría del Activo es obligatoria.")
        If Length(data("IDCentroCoste")) = 0 Then ApplicationService.GenerateError("El Centro de Coste es obligatorio.")
        If Length(data("IDEstadoActivo")) = 0 Then ApplicationService.GenerateError("El Estado del Activo es obligatorio.")
        If Length(data("IDClaseActivo")) = 0 Then ApplicationService.GenerateError("La Clase del Activo es obligatorio.")

        Dim AppParamsActivo As ParametroActivo = services.GetService(Of ParametroActivo)()
        If AppParamsActivo.GestionNumeroSerieConActivos Then
            If Length(data("NSerie")) > 0 AndAlso data("IDActivo") <> data("NSerie") Then
                ApplicationService.GenerateError("El Número de Serie no coincide con el identificador del Activo.")
            End If
            If data.RowState = DataRowState.Modified Then
                If Nz(data("NSerie")) <> Nz(data("NSerie", DataRowVersion.Original)) Then
                    ApplicationService.GenerateError("No se permite modificar el Número de Serie.")
                End If
                If Nz(data("IDArticulo")) <> Nz(data("IDArticulo", DataRowVersion.Original)) AndAlso Length(data("NSerie")) > 0 Then
                    ApplicationService.GenerateError("No se permite modificar el Artículo si tiene un número de serie asociado.")
                End If
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoGuardado, AddressOf AsignarIDActivoPorContador)
        updateProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoGuardado, AddressOf AsignarOperario)
        updateProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoGuardado, AddressOf AsignarOperarioAnterior)
        updateProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoGuardado, AddressOf AsignarFechaEstadoAnterior)
        updateProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoGuardado, AddressOf AsignarEstadoActivoAnterior)
        updateProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoGuardado, AddressOf CambioPadre)
        updateProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoGuardado, AddressOf CambioEstadoActivo)

        updateProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoGuardado, AddressOf Business.General.Comunes.BeginTransaction)
        updateProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoGuardado, AddressOf Business.General.Comunes.UpdateEntityRow)
        updateProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoGuardado, AddressOf Business.General.Comunes.MarcarComoActualizado)
        updateProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoGuardado, AddressOf Activo.HistoricoEstadoActivo)
        updateProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoGuardado, AddressOf Activo.ActualizarNumeroDeSerie)
    End Sub

    <Task()> Public Shared Function NoHaSidoGuardado(ByVal data As DataRow, ByVal services As ServiceProvider) As Boolean
        Dim ListaTratados As ActivosTratados = services.GetService(Of ActivosTratados)()
        Dim HaSidoGuardado As Boolean = ListaTratados.IDActivo.Contains(data("IDActivo"))
        If HaSidoGuardado Then services.GetService(Of UpdateProcessContext).Updated = HaSidoGuardado
        Return Not HaSidoGuardado
    End Function

    ''' <summary>
    ''' Método que asigna un valor al campo IDActivo a través del contador asociado indicado en el registro. Así como la comprobación
    ''' de existencia del Activo en el sistema.
    ''' </summary>
    ''' <param name="data">Registo del Activo</param>
    ''' <param name="services">Objeto para compartir información </param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub AsignarIDActivoPorContador(ByVal data As DataRow, ByVal services As ServiceProvider)
        Select Case data.RowState
            Case DataRowState.Added
                If Length(data("IDContador")) > 0 Then data("IDActivo") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data("IDContador"), services)
                Dim dtActivo As DataTable = New Activo().SelOnPrimaryKey(data("IDActivo"))
                If Not IsNothing(dtActivo) AndAlso dtActivo.Rows.Count > 0 Then
                    ApplicationService.GenerateError("Ya existe el Activo {0}.", Quoted(data("IDActivo")))
                End If
        End Select
    End Sub

    ''' <summary>
    ''' Método que asigna un valor al campo IDEstadoActivoAnterior, si cambia el Estado del Activo y tenemos Gestión por Número de Serie.
    ''' </summary>
    ''' <param name="data">Registo del Activo</param>
    ''' <param name="services">Objeto para compartir información </param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub AsignarEstadoActivoAnterior(ByVal data As DataRow, ByVal services As ServiceProvider)
        Select Case data.RowState
            Case DataRowState.Added
                data("IDEstadoActivoAnterior") = data("IDEstadoActivo")
            Case DataRowState.Modified
                Dim AppParamsActivo As ParametroActivo = services.GetService(Of ParametroActivo)()
                If Nz(data("IDEstadoActivo")) <> Nz(data("IDEstadoActivo", DataRowVersion.Original)) AndAlso AppParamsActivo.GestionNumeroSerieConActivos Then
                    data("IDEstadoActivoAnterior") = data("IDEstadoActivo", DataRowVersion.Original)
                End If
        End Select
    End Sub

    ''' <summary>
    ''' Método que asigna un valor al campo IDOperario. Se le asigna el Operario asociado al Usuario.
    ''' </summary>
    ''' <param name="data">Registo del Activo</param>
    ''' <param name="services">Objeto para compartir información </param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub AsignarOperario(ByVal data As DataRow, ByVal services As ServiceProvider)
        Select Case data.RowState
            Case DataRowState.Added
                Dim Operarios As EntityInfoCache(Of OperarioInfo) = services.GetService(Of EntityInfoCache(Of OperarioInfo))()
                Dim OpInfo As OperarioInfo = Operarios.GetEntity(AdminData.GetSessionInfo.UserID)
                If Not OpInfo Is Nothing AndAlso Length(OpInfo.IDOperario) > 0 Then
                    data("IDOperario") = OpInfo.IDOperario
                Else
                    ApplicationService.GenerateError("El Usuario actual no tiene asociado ningún Operario. El Operario es obligatorio.")
                End If
        End Select
    End Sub

    ''' <summary>
    ''' Método que asigna un valor al campo IDOperarioAnterior, si cambia el Estado del Activo y tenemos Gestión por Número de Serie.
    ''' </summary>
    ''' <param name="data">Registo del Activo</param>
    ''' <param name="services">Objeto para compartir información </param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub AsignarOperarioAnterior(ByVal data As DataRow, ByVal services As ServiceProvider)
        Select Case data.RowState
            Case DataRowState.Added
                data("IDOperarioAnterior") = data("IDOperario")
            Case DataRowState.Modified
                Dim AppParamsActivo As ParametroActivo = services.GetService(Of ParametroActivo)()
                If Nz(data("IDEstadoActivo")) <> Nz(data("IDEstadoActivo", DataRowVersion.Original)) AndAlso AppParamsActivo.GestionNumeroSerieConActivos Then
                    data("IDOperarioAnterior") = data("IDOperario", DataRowVersion.Original)
                End If
        End Select
    End Sub

    ''' <summary>
    ''' Método que asigna un valor al campo FechaEstadoAnterior, si cambia el Estado del Activo y tenemos Gestión por Número de Serie.
    ''' </summary>
    ''' <param name="data">Registo del Activo</param>
    ''' <param name="services">Objeto para compartir información </param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub AsignarFechaEstadoAnterior(ByVal data As DataRow, ByVal services As ServiceProvider)
        Select Case data.RowState
            Case DataRowState.Added
                If IsDate(data("FechaEstado")) Then
                    data("FechaEstadoAnterior") = data("FechaEstado")
                Else
                    data("FechaEstadoAnterior") = Today
                End If
            Case DataRowState.Modified
                Dim AppParamsActivo As ParametroActivo = services.GetService(Of ParametroActivo)()
                If Nz(data("IDEstadoActivo")) <> Nz(data("IDEstadoActivo", DataRowVersion.Original)) AndAlso AppParamsActivo.GestionNumeroSerieConActivos Then
                    data("FechaEstadoAnterior") = data("FechaEstado", DataRowVersion.Original)
                End If
        End Select
    End Sub

    ''' <summary>
    ''' Método que valida si el Activo existe en el árbol de la estructura y si ya existe, no permite que sea padre.
    ''' </summary>
    ''' <param name="data">Registo del Activo</param>
    ''' <param name="services">Objeto para compartir información </param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub CambioPadre(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("Padre")) > 0 AndAlso (data.RowState = DataRowState.Added OrElse data.RowState = DataRowState.Modified AndAlso data("Padre") <> Nz(data("Padre", DataRowVersion.Original))) AndAlso data("Padre") Then
            Dim datos As New ActivoEstructura.DataExisteActivoEnExplosion(String.Empty, data("IDActivo"), True)
            If ProcessServer.ExecuteTask(Of ActivoEstructura.DataExisteActivoEnExplosion, Boolean)(AddressOf ActivoEstructura.ExisteActivoEnExplosion, datos, services) Then
                data("Padre") = False
            End If
        End If

    End Sub

    ''' <summary>
    ''' Método que realiza las operaciones que desencadena el cambio del Estado del Activo
    ''' </summary>
    ''' <param name="data">Registo del Activo</param>
    ''' <param name="services">Objeto para compartir información </param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub CambioEstadoActivo(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim IDActivosActualizados(-1) As String
        If data.RowState = DataRowState.Modified AndAlso Nz(data("IDEstadoActivo")) <> Nz(data("IDEstadoActivo", DataRowVersion.Original)) Then
            Dim DE As New BE.DataEngine
            Dim DtEstActual As DataTable = DE.Filter("tbMntoEstadoActivo", New FilterItem("IDEstadoActivo", data("IDEstadoActivo")))
            Dim DtEstOriginal As DataTable = DE.Filter("tbMntoEstadoActivo", New FilterItem("IDEstadoActivo", data("IDEstadoActivo", DataRowVersion.Original)))
            If DtEstActual.Rows(0)("Baja") AndAlso Not DtEstOriginal.Rows(0)("Baja") Then
                Dim serie As DataTable = New ArticuloNSerie().Filter(New StringFilterItem("IDActivo", data("IDActivo")))
                If Not serie Is Nothing AndAlso serie.Rows.Count > 0 Then
                    Dim n As Integer = ProcessServer.ExecuteTask(Of Object, Integer)(AddressOf ProcesoStocks.NuevoNumeroMovimiento, Nothing, services)
                    Dim baja As New StockData(data("IDArticulo"), serie.Rows(0)("IDAlmacen") & String.Empty, serie.Rows(0)("NSerie"), data("IDActivo"), data("IDEstadoActivo"), Nz(serie.Rows(0)("IDOperario"), data("IDOperario")), Today, enumTipoMovimiento.tmSalAjuste)
                    baja.Texto = "Cambio de estado a Baja en el Activo (Salida)"
                    Dim datAjte As New DataNumeroMovimientoSinc(n, baja)
                    Dim bajaUpdateData As StockUpdateData = ProcessServer.ExecuteTask(Of DataNumeroMovimientoSinc, StockUpdateData)(AddressOf ProcesoStocks.Ajuste, datAjte, services)
                    If Not bajaUpdateData Is Nothing Then
                        If Not bajaUpdateData.Estado = EstadoStock.Actualizado Then
                            Throw New Exception(bajaUpdateData.Detalle)
                        Else
                            For Each drActivo As DataRow In bajaUpdateData.Activo.Select
                                Dim htActivos As ActivosTratados = services.GetService(Of ActivosTratados)()
                                htActivos.IDActivo(drActivo("IDActivo")) = drActivo("IDActivo")
                            Next
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' Método que crea un nuevo registro en el Histórico de Estados del Activo
    ''' </summary>
    ''' <param name="activo">Registo del Activo</param>
    ''' <param name="services">Objeto para compartir información </param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub HistoricoEstadoActivo(ByVal data As DataRow, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataRow, DataTable)(AddressOf NewHistoricoEstadoActivo, data, services)
    End Sub

    <Task()> Public Shared Function NewHistoricoEstadoActivo(ByVal data As DataRow, ByVal services As ServiceProvider) As DataTable
        If Not data Is Nothing AndAlso Length(data("IDActivo")) > 0 Then
            If data.RowState = DataRowState.Added OrElse (data.RowState = DataRowState.Modified AndAlso data("IDEstadoActivo") & String.Empty <> data("IDEstadoActivo", DataRowVersion.Original) & String.Empty) Then
                Dim FilActivo As New Filter
                FilActivo.Add("IDEstadoActivo", FilterOperator.Equal, data("IDEstadoActivo"))
                FilActivo.Add("FechaEstado", FilterOperator.Equal, data("FechaEstado"))
                FilActivo.Add("IDActivo", FilterOperator.Equal, data("IDActivo"))
                Dim DtHist As DataTable = New HistoricoEstadoActivo().Filter(FilActivo)
                If DtHist Is Nothing OrElse DtHist.Rows.Count = 0 Then
                    Dim historico As DataTable = New HistoricoEstadoActivo().AddNewForm()
                    If historico.Rows.Count > 0 Then
                        historico.Rows(0)("IDActivo") = data("IDActivo")
                        historico.Rows(0)("IDEstadoActivo") = data("IDEstadoActivo")
                        If Length(data("IDOperario")) > 0 Then
                            historico.Rows(0)("IDOperario") = data("IDOperario")
                        Else
                            Dim Operarios As EntityInfoCache(Of OperarioInfo) = services.GetService(Of EntityInfoCache(Of OperarioInfo))()
                            Dim OpInfo As OperarioInfo = Operarios.GetEntity(AdminData.GetSessionInfo.UserID)
                            If Not OpInfo Is Nothing AndAlso Length(OpInfo.IDOperario) > 0 Then
                                historico.Rows(0)("IDOperario") = OpInfo.IDOperario
                            Else
                                ApplicationService.GenerateError("El Usuario actual no tiene asociado ningún Operario. El Operario es obligatorio.")
                            End If
                        End If
                        If IsDate(data("FechaEstado")) Then
                            historico.Rows(0)("FechaEstado") = data("FechaEstado")
                        Else
                            historico.Rows(0)("FechaEstado") = Today
                        End If
                        BusinessHelper.UpdateTable(historico)
                    End If
                End If
            End If
        End If
    End Function




    '' <summary>
    '' Método que actualiza el Nº Serie del Artículo
    '' </summary>
    '' <param name="activo">Registo del Activo</param>
    '' <param name="services">Objeto para compartir información </param>
    '' <remarks></remarks>
    <Task()> Public Shared Sub ActualizarNumeroDeSerie(ByVal activo As DataRow, ByVal services As ServiceProvider)
        Dim AppParamsActivo As ParametroActivo = services.GetService(Of ParametroActivo)()
        If AppParamsActivo.GestionNumeroSerieConActivos Then
            If Not activo Is Nothing AndAlso Length(activo("NSerie")) > 0 Then
                Dim f As New Filter
                f.Add(New StringFilterItem("IDArticulo", activo("idArticulo")))
                f.Add(New StringFilterItem("NSerie", activo("NSerie")))
                Dim serie As DataTable = New ArticuloNSerie().Filter(f)
                If serie.Rows.Count > 0 Then
                    serie.Rows(0)("IDActivo") = activo("IDActivo")
                    serie.Rows(0)("IDEstadoActivo") = activo("IDEstadoActivo")
                    serie.Rows(0)("IDArticulo") = activo("IDArticulo")
                    If Length(activo("IDOperario")) > 0 Then serie.Rows(0)("IDOperario") = activo("IDOperario")
                    BusinessHelper.UpdateTable(serie)
                Else
                    ApplicationService.GenerateError("El número de serie | no existe.", Quoted(activo("NSerie")))
                End If
            End If
        End If
    End Sub

#End Region

#Region " RegisterDeleteTask "

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf DeleteHistoricoActivo)
    End Sub

    <Task()> Public Shared Sub DeleteHistoricoActivo(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim HEA As New HistoricoEstadoActivo
        Dim f As New Filter
        f.Add(New StringFilterItem("IDActivo", data("IDActivo")))
        Dim Historico As DataTable = HEA.Filter(f)
        If Not Historico Is Nothing AndAlso Historico.Rows.Count = 1 AndAlso Historico.Rows(0)("IDEstadoActivo") = NegocioGeneral.ESTADOACTIVO_DISPONIBLE Then
            HEA.Delete(Historico)
        ElseIf Not Historico Is Nothing AndAlso Historico.Rows.Count > 0 Then
            ApplicationService.GenerateError("El Activo tiene un Histórico, o bien, no está en estado Disponible. No se puede eliminar el Activo.")
        End If
    End Sub


#End Region

#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBRL As New BusinessRules
        oBRL.Add("NSerie", AddressOf CambioNSerieEnActivo)
        oBRL.Add("Padre", AddressOf CambioPadreEnActivo)
        Return oBRL
    End Function

    ''' <summary>
    ''' Método que realiza las operaciones derivadas del cambio del NºSerie del Activo
    ''' </summary>
    ''' <param name="data">Objeto BusinessRuleData</param>
    ''' <param name="services">Objeto para compartir información</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub CambioNSerieEnActivo(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("NSerie")) > 0 Then
            '//Dado el Nº de serie, buscamos su artículo.
            Dim f As New Filter
            f.Add(New StringFilterItem("NSerie", data.Current("NSerie")))
            Dim dtArtNSerie As DataTable = New ArticuloNSerie().Filter(f)
            If Not IsNothing(dtArtNSerie) AndAlso dtArtNSerie.Rows.Count > 0 Then
                data.Current("IDArticulo") = dtArtNSerie.Rows(0)("IDArticulo")
                'Aqui obtengo la DescArticulo a partir de la IDArticulo
                Dim f2 As New Filter
                f2.Add(New StringFilterItem("IDArticulo", data.Current("IDArticulo")))
                Dim dtArticulo As DataTable = New Articulo().Filter(f2)

                data.Current("DescActivo") = dtArticulo.Rows(0)("DescArticulo")
                data.Current("IDActivo") = dtArtNSerie.Rows(0)("NSerie")
            Else
                data.Current("IDArticulo") = System.DBNull.Value
                ApplicationService.GenerateError("El Nº de Serie {0} no existe.", Quoted(data.Current("NSerie")))
            End If
        Else
            data.Current("IDArticulo") = System.DBNull.Value
        End If
    End Sub

    ''' <summary>
    ''' Método que realiza las operaciones derivadas del cambio de la marca Padre del Activo
    ''' </summary>
    ''' <param name="data">Objeto BusinessRuleData</param>
    ''' <param name="services">Objeto para compartir información</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub CambioPadreEnActivo(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("Padre")) > 0 AndAlso data.Current("Padre") Then
            '//Validar si el Activo puede ser padre, o forma parte de otro activo.
            Dim datos As New ActivoEstructura.DataExisteActivoEnExplosion
            datos.IDActivoHijo = String.Empty
            datos.IDActivoBase = data.Current("IDActivo")
            datos.ValidarPadre = True
            If ProcessServer.ExecuteTask(Of ActivoEstructura.DataExisteActivoEnExplosion, Boolean)(AddressOf ActivoEstructura.ExisteActivoEnExplosion, datos, services) Then
                ApplicationService.GenerateError("El Activo no puede ser Padre, es un Activo componente de otro Activo.")
            End If
        End If
    End Sub


#End Region

#Region "Funciones Públicas"

    <Serializable()> _
    Public Class DatosCambiarArtEnMaq
        Public IDArticuloOld As String
        Public IDArticuloNew As String
        Public IDActivo As String

        Public Sub New(ByVal IDArticuloOld As String, ByVal IDArticuloNew As String, ByVal IDActivo As String)
            Me.IDArticuloOld = IDArticuloOld
            Me.IDArticuloNew = IDArticuloNew
            Me.IDActivo = IDActivo
        End Sub
    End Class

    <Task()> Public Shared Sub CambiarArticuloEnMaquinaria(ByVal data As DatosCambiarArtEnMaq, ByVal services As ServiceProvider)
        '///Cambio de Articulo en maquinaria///'
        If Length(data.IDArticuloOld) > 0 AndAlso Length(data.IDArticuloNew) > 0 Then
            Dim ans As New ArticuloNSerie
            Dim f As New Filter
            f.Add(New StringFilterItem("IDActivo", data.IDActivo))
            f.Add(New StringFilterItem("IDArticulo", data.IDArticuloOld))
            Dim serie As DataTable = ans.Filter(f)
            If Not IsNothing(serie) AndAlso serie.Rows.Count > 0 Then
                If Length(serie.Rows(0)("IDAlmacen")) > 0 Then
                    '//Comprobar que el nuevo artículo existe en el mismo almacén.
                    Dim objFilter As New Filter
                    objFilter.Add(New StringFilterItem("IDArticulo", data.IDArticuloNew))
                    objFilter.Add(New StringFilterItem("IDAlmacen", serie.Rows(0)("IDAlmacen")))
                    Dim objNegArtAlm As New ArticuloAlmacen
                    Dim dtArtAlm As DataTable = objNegArtAlm.Filter(objFilter)
                    If IsNothing(dtArtAlm) OrElse dtArtAlm.Rows.Count = 0 Then
                        ApplicationService.GenerateError("El artículo | no existe en el almacén |.", Quoted(data.IDArticuloNew), Quoted(serie.Rows(0)("IDAlmacen")))
                    End If

                    '/////////////
                    Dim strIDEstadoActivo As String
                    Dim dtActivo As DataTable = New Activo().Filter(New StringFilterItem("IDActivo", data.IDActivo))
                    If Not IsNothing(dtActivo) AndAlso dtActivo.Rows.Count > 0 Then
                        strIDEstadoActivo = dtActivo.Rows(0)("IDEstadoActivo")
                    End If
                    '//////////

                    Dim stk As New ProcesoStocks
                    Dim n As Integer = ProcessServer.ExecuteTask(Of Object, Integer)(AddressOf ProcesoStocks.NuevoNumeroMovimiento, Nothing, services)
                    Dim baja As New StockData(data.IDArticuloOld, serie.Rows(0)("IDAlmacen"), serie.Rows(0)("NSerie"), data.IDActivo, NegocioGeneral.ESTADOACTIVO_BAJA, serie.Rows(0)("IDOperario"), Today, enumTipoMovimiento.tmSalAjuste)
                    baja.Texto = "Cambio de artículo en maquinaria (Salida)"
                    Dim datAjte As New DataNumeroMovimientoSinc(n, baja)
                    Dim bajaUpdateData As StockUpdateData = ProcessServer.ExecuteTask(Of DataNumeroMovimientoSinc, StockUpdateData)(AddressOf ProcesoStocks.Ajuste, datAjte, services)
                    If Not bajaUpdateData Is Nothing Then
                        If bajaUpdateData.Estado = EstadoStock.Actualizado Then
                            Dim alta As New StockData(data.IDArticuloNew, serie.Rows(0)("IDAlmacen"), serie.Rows(0)("NSerie"), data.IDActivo, strIDEstadoActivo, serie.Rows(0)("IDOperario"), Today, enumTipoMovimiento.tmEntAjuste)
                            alta.Texto = "Cambio de artículo en maquinaria (Entrada)"
                            datAjte = New DataNumeroMovimientoSinc(n, alta)
                            Dim altaUpdateData As StockUpdateData = ProcessServer.ExecuteTask(Of DataNumeroMovimientoSinc, StockUpdateData)(AddressOf ProcesoStocks.Ajuste, datAjte, services)
                            If altaUpdateData.Estado <> EstadoStock.Actualizado Then
                                'Se deja porque bajaUpdateData.Detalle ya viene traduccido
                                Throw New Exception(altaUpdateData.Detalle)
                            End If
                        Else
                            'Se deja porque bajaUpdateData.Detalle ya viene traduccido
                            Throw New Exception(bajaUpdateData.Detalle)
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Function EstadoPendiente(ByVal data As Object, ByVal services As ServiceProvider) As String
        EstadoPendiente = NegocioGeneral.ESTADOACTIVO_PENDIENTEDERETORNAR
    End Function

    <Task()> Public Shared Function EstadoTrabajando(ByVal data As Object, ByVal services As ServiceProvider) As String
        EstadoTrabajando = NegocioGeneral.ESTADOACTIVO_TRABAJANDO
    End Function

    <Task()> Public Shared Function ObtenerRepuestosActivo(ByVal data As String, ByVal services As ServiceProvider) As DataTable
        Return New BE.DataEngine().Filter("vFrmMntoActivoRepuesto", New StringFilterItem("IDActivo", data))
    End Function

    <Task()> Public Shared Function ObtenerEstructuraActivo(ByVal data As String, ByVal services As ServiceProvider) As DataTable
        Return New BE.DataEngine().Filter("vFrmMntoActivoEstructura", New StringFilterItem("IDActivo", data))
    End Function

    <Serializable()> _
    Public Class DatosObtAlqRetorno
        Public IDArticulo As String
        Public IDActivo As String
        Public TipoAlbaran As String

        Public Sub New(ByVal IDArticulo As String, ByVal IDActivo As String, ByVal TipoAlbaran As String)
            Me.IDArticulo = IDArticulo
            Me.IDActivo = IDActivo
            Me.TipoAlbaran = TipoAlbaran
        End Sub
    End Class

    <Task()> Public Shared Function ObtenerAlquilerRetorno(ByVal data As DatosObtAlqRetorno, ByVal services As ServiceProvider) As DataTable
        Dim f As New Filter
        f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
        f.Add(New StringFilterItem("Lote", data.IDActivo))
        f.Add(New StringFilterItem("IdTipoAlbaran", data.TipoAlbaran))
        f.Add(New NumberFilterItem("QPendiente", FilterOperator.GreaterThan, 0))
        Return New BE.DataEngine().Filter("VCtlCiAlquilerRetorno", f, , "IDObra DESC, FechaAlbaran DESC")
    End Function

    <Task()> Public Shared Function UsuarioPermisoMaquinaria(ByVal data As Object, ByVal services As ServiceProvider) As Boolean
        UsuarioPermisoMaquinaria = False
        Dim objFilterUsuario As New Filter
        objFilterUsuario.Add(New GuidFilterItem("IDUsuario", NegocioGeneral.UserID))
        Dim objNegOperario As New Operario
        Dim dtOperario As DataTable = objNegOperario.Filter(objFilterUsuario)

        If Not IsNothing(dtOperario) AndAlso dtOperario.Rows.Count > 0 Then
            UsuarioPermisoMaquinaria = dtOperario.Rows(0)("EstadoMaquina")
        End If
    End Function

    <Serializable()> _
    Public Class DatosCuentaExplotacion
        Public IDActivo As String
        Public IDProcess As String
        Public Detalle As Boolean

        Public Sub New(ByVal IDActivo As String, ByVal IDProcess As String, ByVal Detalle As Boolean)
            Me.IDActivo = IDActivo
            Me.IDProcess = IDProcess
            Me.Detalle = Detalle
        End Sub
    End Class

    <Task()> Public Shared Function CuentaExplotacion(ByVal data As DatosCuentaExplotacion, ByVal services As ServiceProvider) As DataTable
        Return AdminData.Execute("sp_CuentaExplotacionMaquina", False, data.IDProcess, CInt(data.Detalle))
    End Function

    <Task()> Public Shared Function DuplicarActivo(ByVal data As String, ByVal services As ServiceProvider) As String
        Dim strIDActivoNew As String
        Dim objFilterActivo As New Filter
        objFilterActivo.Add(New StringFilterItem("IDActivo", data))
        Dim ClsActivo As New Activo
        Dim dtActivo As DataTable = ClsActivo.Filter(objFilterActivo)
        If Not IsNothing(dtActivo) AndAlso dtActivo.Rows.Count > 0 Then
            Dim dtActivoNew As DataTable = ClsActivo.AddNewForm
            If Not IsNothing(dtActivoNew) AndAlso dtActivoNew.Rows.Count > 0 Then
                For Each col As DataColumn In dtActivoNew.Columns
                    If col.ColumnName <> "IDActivo" AndAlso col.ColumnName <> "IDContador" AndAlso col.ColumnName <> "NSerie" AndAlso col.ColumnName <> "FechaAlta" Then
                        dtActivoNew.Rows(0)(col.ColumnName) = dtActivo.Rows(0)(col.ColumnName)
                    End If
                Next
            End If
            dtActivoNew = ClsActivo.Update(dtActivoNew)

            If Not IsNothing(dtActivoNew) AndAlso dtActivoNew.Rows.Count > 0 Then
                strIDActivoNew = dtActivoNew.Rows(0)("IDActivo")
                Dim objNegActivoEstructura As New ActivoEstructura
                Dim dtEstructura As DataTable = objNegActivoEstructura.Filter(objFilterActivo)
                Dim dtEstructuraNew As DataTable = dtEstructura.Clone
                For Each drEstructura As DataRow In dtEstructura.Rows
                    Dim drEstructuraNew As DataRow = dtEstructuraNew.NewRow
                    drEstructuraNew.ItemArray = drEstructura.ItemArray
                    drEstructuraNew("IDActivo") = strIDActivoNew
                    dtEstructuraNew.Rows.Add(drEstructuraNew)
                Next
                BusinessHelper.UpdateTable(dtEstructuraNew)

                Dim objNegActivoRepuesto As New ActivoRepuesto
                Dim dtRepuesto As DataTable = objNegActivoRepuesto.Filter(objFilterActivo)
                Dim dtRepuestoNew As DataTable = dtRepuesto.Clone
                For Each drRepuesto As DataRow In dtRepuesto.Rows
                    Dim drRepuestoNew As DataRow = dtRepuestoNew.NewRow
                    drRepuestoNew.ItemArray = drRepuesto.ItemArray
                    drRepuestoNew("IDActivo") = strIDActivoNew
                    dtRepuestoNew.Rows.Add(drRepuestoNew)
                Next
                BusinessHelper.UpdateTable(dtRepuestoNew)
            End If
        End If
        Return strIDActivoNew
    End Function

    <Serializable()> _
    Public Class StActuaArtSerie
        Public DtActivo As DataTable
        Public IDArticuloDestino As String

        Public Sub New(ByVal DtActivo As DataTable, ByVal IDArticuloDestino As String)
            Me.DtActivo = DtActivo
            Me.IDArticuloDestino = IDArticuloDestino
        End Sub
    End Class

    <Task()> Public Shared Sub ActualizarArticuloSerie(ByVal data As StActuaArtSerie, ByVal services As ServiceProvider)
        Dim StrIDArticuloOrigen As String = data.DtActivo.Rows(0)("IDArticulo")
        data.DtActivo.Rows(0)("IDArticulo") = data.IDArticuloDestino
        data.DtActivo.TableName = "Activo"
        AdminData.SetData(data.DtActivo)

        Dim ClsArtNSerie As New ArticuloNSerie
        Dim FilArtSerie As New Filter
        FilArtSerie.Add("IDArticulo", FilterOperator.Equal, StrIDArticuloOrigen)
        FilArtSerie.Add("NSerie", FilterOperator.Equal, data.DtActivo.Rows(0)("NSerie"))
        Dim DtArtSerie As DataTable = ClsArtNSerie.Filter(FilArtSerie)
        If Not DtArtSerie Is Nothing AndAlso DtArtSerie.Rows.Count > 0 Then
            DtArtSerie.Rows(0)("IDArticulo") = data.IDArticuloDestino
            AdminData.SetData(DtArtSerie)
        End If
    End Sub
#End Region

    Public Sub ActualizaStocksAlquiler(ByVal IDArticulo As String, ByVal obra As String, ByVal stock As Double)

        Dim stock011 As Double
        '1º. Recupero el Stock de ese articulo en el 011.
        Dim strSQL As String = "SELECT StockFisico FROM tbMaestroArticuloAlmacen WHERE IDArticulo = '" & IDArticulo & "' AND IDAlmacen='011'"
        Dim tb As DataTable = AdminData.GetData(strSQL)
        For Each dr As DataRow In tb.Rows
            stock011 = dr("StockFisico")
            'MsgBox("IDARticulo= " & IDArticulo & " con stock " & stock011 & " en el 011.")
        Next
        '2º. Le sumo el stock que quiero(para luego darle salida desde alquileres y devolverselo).

        stock011 += stock
        'MsgBox(stock011)
        '3º. Actualizo Stock del 011
        Dim strSQL2 As String
        strSQL2 = " UPDATE tbMaestroArticuloAlmacen"
        strSQL2 &= " SET StockFisico = ('" & stock011 & "')"
        strSQL2 &= " WHERE IDArticulo = ('" & IDArticulo & "') AND IDAlmacen=('" & "011" & "')"

        Try
            AdminData.Execute(strSQL2)
        Catch ex As Exception
            ApplicationService.GenerateError(ex.ToString & ": ERROR")
        End Try

        '4º. Obtengo del stock de la obra.
        Dim strSQL3 As String = "SELECT StockFisico FROM tbMaestroArticuloAlmacen WHERE IDArticulo = '" & IDArticulo & "' AND IDAlmacen='" & obra & "'"
        Dim tb2 As DataTable = AdminData.GetData(strSQL3)
        Dim stockobra
        For Each dr As DataRow In tb2.Rows
            stockobra = dr("StockFisico")
            'MsgBox("IDARticulo= " & IDArticulo & " con stock " & stockobra & " en " & obra & ".")
        Next

        stockobra = stockobra - stock
        '5º. Actualizo el stock de la obra. 
        Dim strSQL4 As String
        strSQL4 = " UPDATE tbMaestroArticuloAlmacen"
        strSQL4 &= " SET StockFisico = ('" & stockobra & "')"
        strSQL4 &= " WHERE IDArticulo = ('" & IDArticulo & "') AND IDAlmacen=('" & obra & "')"

        Try
            AdminData.Execute(strSQL4)
        Catch ex As Exception
            ApplicationService.GenerateError(ex.ToString & ": ERROR")
        End Try
        
    End Sub

End Class

Public Class ActivosTratados
    Public IDActivo As Hashtable

    Public Sub New()
        IDActivo = New Hashtable
    End Sub
End Class
