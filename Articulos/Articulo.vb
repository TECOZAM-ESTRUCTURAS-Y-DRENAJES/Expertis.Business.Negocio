Public Class ArticuloInfo
    Inherits ClassEntityInfo

    Public IDArticulo As String
    Public DescArticulo As String
    Public IDTipo As String
    Public IDFamilia As String
    Public IDSubFamilia As String
    Public IDUDInterna As String
    Public IDUDInterna2 As String
    Public IDUDVenta As String
    Public IDUDCompra As String
    Public UDValoracion As Integer
    Public CodigoBarras As String
    Public IDTipoIVA As String
    Public Activo As Boolean
    Public Configurable As Boolean
    Public IDArticuloConfigurado As String
    Public ContRadical? As Integer
    Public Compra As Boolean
    Public Venta As Boolean
    Public Fabrica As Boolean
    Public Servicio As Boolean
    Public Generico As Boolean
    Public KitVenta As Boolean
    Public Utillaje As Boolean
    Public SubProducto As Boolean
    Public Fantasma As Boolean
    Public GestionStock As Boolean
    Public GestionStockPorLotes As Boolean
    Public StockNegativo As Boolean
    Public Subcontratacion As Boolean
    Public Embalaje As Boolean
    Public Especial As Boolean
    Public NSerieObligatorio As Boolean
    Public TipoFactAlquiler As Integer
    Public PesoNeto As Double
    Public PesoBruto As Double

    Public ReadOnly Property GestionPorNumeroSerie() As Boolean
        Get
            Return NSerieObligatorio
        End Get
    End Property
    Public CCExport As String
    Public CCExportGrupo As String
    Public CCVenta As String
    Public CCVentaGrupo As String
    Public CCImport As String
    Public CCImportGrupo As String
    Public CCCompra As String
    Public CCCompraGrupo As String
    Public CriterioValoracion As enumtaValoracion
    Public TipoPrecio As enumTipoPrecio
    Public PrecioBase As Double
    Public PrecioEstandarA As Double
    Public PrecioEstandarB As Double
    Public PrecioUltimaCompraA As Double
    Public PrecioUltimaCompraB As Double
    Public IDProveedorUltimaCompra As String
    Public IDArticuloContenedor As String
    Public QContenedor As Double
    Public IDArticuloEmbalaje As String
    Public QEmbalaje As Double
    Public ControlRecepcion As enumControlRecepcion '//Calidad
    Public RecalcularValoracion As Integer
    Public SinDtoEnAlquiler As Boolean
    Public IDConcepto As String
    Public IDTipoIVAReducido As String
    Public PorcenIVANoDeducible As Double
    Public SinSeguroEnAlquiler As Boolean
    Public FactTasaResiduos As Boolean
    Public PrecioBaseConfigurado As Double
    Public [Alias] As String
    Public RetencionIRPF As Boolean
    Public NivelPlano As String

    Public ReadOnly Property ArticuloDePortes() As Boolean
        Get
            Return (Especial AndAlso NSerieObligatorio)
        End Get
    End Property

    Public ReadOnly Property TieneSegundaUnidad() As Boolean
        Get
            Return (Length(IDUDInterna2) > 0)
        End Get
    End Property

    Public EnsambladoStock As String            '//Sincronización con Stocks de Bodega
    Public ClaseStock As String                 '//Sincronización con Stocks de Bodega

    Public GenerarOFArticuloFinal As Boolean    '//Producción
    Public IDArticuloFinal As String            '//Producción

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal data As DataRow)
        MyBase.New(data)
    End Sub

    Public Overloads Overrides Sub Fill(ByVal ParamArray PrimaryKey() As Object)
        Dim CARACT_ARTICULO As String = "vNegCaractArticulo"

        Dim dt As DataTable
        If Not IsNothing(PrimaryKey) AndAlso PrimaryKey.Length > 0 AndAlso Length(PrimaryKey(0)) > 0 Then
            dt = New BE.DataEngine().Filter(CARACT_ARTICULO, New StringFilterItem("IDArticulo", PrimaryKey(0)))
        End If

        If IsNothing(dt) OrElse dt.Rows.Count = 0 Then
            ApplicationService.GenerateError("El artículo | no existe.", Quoted(PrimaryKey(0)))
        Else
            Me.Fill(dt.Rows(0))
        End If
    End Sub

End Class

Public Class _Articulo
    Public Const IDArticulo As String = "IDArticulo"
    Public Const DescArticulo As String = "DescArticulo"
    Public Const IDContador As String = "IDContador"
    Public Const FechaAlta As String = "FechaAlta"
    Public Const IDEstado As String = "IDEstado"
    Public Const IDTipo As String = "IDTipo"
    Public Const IDFamilia As String = "IDFamilia"
    Public Const IDSubfamilia As String = "IDSubfamilia"
    Public Const CCVenta As String = "CCVenta"
    Public Const CCExport As String = "CCExport"
    Public Const CCCompra As String = "CCCompra"
    Public Const CCImport As String = "CCImport"
    Public Const CCVentaRegalo As String = "CCVentaRegalo"
    Public Const CCGastoRegalo As String = "CCGastoRegalo"
    Public Const CCStocks As String = "CCStocks"
    Public Const IDTipoIva As String = "IDTipoIva"
    Public Const IDPartidaEstadistica As String = "IDPartidaEstadistica"
    Public Const IDUdInterna As String = "IDUdInterna"
    Public Const IDUdVenta As String = "IDUdVenta"
    Public Const IDUdCompra As String = "IDUdCompra"
    Public Const PrecioEstandarA As String = "PrecioEstandarA"
    Public Const PrecioEstandarB As String = "PrecioEstandarB"
    Public Const ValorReposicionA As String = "ValorReposicionA"
    Public Const ValorReposicionB As String = "ValorReposicionB"
    Public Const FechaEstandar As String = "FechaEstandar"
    Public Const FechaValorReposicion As String = "FechaValorReposicion"
    Public Const UdValoracion As String = "UdValoracion"
    Public Const PesoNeto As String = "PesoNeto"
    Public Const PesoBruto As String = "PesoBruto"
    Public Const TipoEstructura As String = "TipoEstructura"
    Public Const IDTipoEstructura As String = "IDTipoEstructura"
    Public Const TipoRuta As String = "TipoRuta"
    Public Const IDTipoRuta As String = "IDTipoRuta"
    Public Const PuntoVerde As String = "PuntoVerde"
    Public Const CodigoBarras As String = "CodigoBarras"
    Public Const PVPMinimo As String = "PVPMinimo"
    Public Const PorcentajeRechazo As String = "PorcentajeRechazo"
    Public Const Plazo As String = "Plazo"
    Public Const Volumen As String = "Volumen"
    Public Const RecalcularValoracion As String = "RecalcularValoracion"
    Public Const CriterioValoracion As String = "CriterioValoracion"
    Public Const GestionStockPorLotes As String = "GestionStockPorLotes"
    Public Const PrecioUltimaCompraA As String = "PrecioUltimaCompraA"
    Public Const PrecioUltimaCompraB As String = "PrecioUltimaCompraB"
    Public Const FechaUltimaCompra As String = "FechaUltimaCompra"
    Public Const IDProveedorUltimaCompra As String = "IDProveedorUltimaCompra"
    Public Const LoteMultiplo As String = "LoteMultiplo"
    Public Const CantMinSolicitud As String = "CantMinSolicitud"
    Public Const CantMaxSolicitud As String = "CantMaxSolicitud"
    Public Const LimitarPetDia As String = "LimitarPetDia"
    Public Const NivelPlano As String = "NivelPlano"
    Public Const StockNegativo As String = "StockNegativo"
    Public Const PlazoFabricacion As String = "PlazoFabricacion"
    Public Const IdArticuloConfigurado As String = "IdArticuloConfigurado"
    Public Const ContRadical As String = "ContRadical"
    Public Const IdFamiliaConfiguracion As String = "IdFamiliaConfiguracion"
    Public Const PrecioBase As String = "PrecioBase"
    Public Const Configurable As String = "Configurable"
    Public Const ParamMaterial As String = "ParamMaterial"
    Public Const ParamTerminado As String = "ParamTerminado"
    Public Const CapacidadDiaria As String = "CapacidadDiaria"
    Public Const AplicarLoteMRP As String = "AplicarLoteMRP"
    Public Const NSerieObligatorio As String = "NSerieObligatorio"
    Public Const PuntosMarketing As String = "PuntosMarketing"
    Public Const ValorPuntosMarketing As String = "ValorPuntosMarketing"
    Public Const ControlRecepcion As String = "ControlRecepcion"
    Public Const IDEstadoHomologacion As String = "IDEstadoHomologacion"
    Public Const IDArticuloFinal As String = "IDArticuloFinal"
    Public Const GenerarOFArticuloFinal As String = "GenerarOFArticuloFinal"
    Public Const IdDocumentoEspecificacion As String = "IdDocumentoEspecificacion"
    Public Const NivelModificacionPlan As String = "NivelModificacionPlan"
    Public Const FechaModificacionNivelPlan As String = "FechaModificacionNivelPlan"
    Public Const TipoFactAlquiler As String = "TipoFactAlquiler"
    Public Const Observaciones As String = "Observaciones"
    Public Const Seguridad As String = "Seguridad"
    Public Const Reglamentacion As String = "Reglamentacion"
    Public Const SeguridadReglamentacion As String = "SeguridadReglamentacion"
    Public Const DiasMinimosFactAlquiler As String = "DiasMinimosFactAlquiler"
    Public Const SinDtoEnAlquiler As String = "SinDtoEnAlquiler"
    Public Const SinSeguroEnAlquiler As String = "SinSeguroEnAlquiler"
    Public Const NecesitaOperario As String = "NecesitaOperario"
    Public Const IDConcepto As String = "IDConcepto"
    Public Const FechaCreacionAudi As String = "FechaCreacionAudi"
    Public Const FechaModificacionAudi As String = "FechaModificacionAudi"
    Public Const UsuarioAudi As String = "UsuarioAudi"
End Class

#Region " Configurador "

Public Interface IConfiguradorArticulo
    Sub ADDCaracteristicas(ByVal dtArticulo As DataTable)
    Sub RecalcularPreciosCaracteristicas(ByVal dtArticulo As DataTable)
End Interface

#End Region

Public Class Articulo

#Region "Constructor"
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Private Const cnEntidad As String = "tbMaestroArticulo"
    Private mblnADD As Boolean

    Private mdtAddDatosConfigurador As DataTable
    Private mdtUpdateDatosConfigurador As DataTable

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub
#End Region

#Region "Eventos RegisterAddNewTasks "

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    ''' <summary>
    ''' Rellenar valores por defecto al crear un nuevo registro
    ''' </summary>
    ''' <param name="data">Registro Nuevo</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks>Se ejecutan las tareas de nuevo registro</remarks>
    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarContador, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarValoresPredeterminados, data, services)
    End Sub
    ''' <summary>
    ''' Asignación  de contador de artículos predeterminado
    ''' </summary>
    ''' <param name="data">Registro en el que se asignan el contador de artículos</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub AsignarContador(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim StDatos As New Contador.DatosDefaultCounterValue
        StDatos.row = data
        StDatos.EntityName = "Articulo"
        StDatos.FieldName = "IDArticulo"
        ProcessServer.ExecuteTask(Of Contador.DatosDefaultCounterValue)(AddressOf Contador.LoadDefaultCounterValue, StDatos, services)
    End Sub
    ''' <summary>
    ''' Asignación  de valores predeterminados
    ''' </summary>
    ''' <param name="data">Registro en el que se asignan los valores predeterminados</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub AsignarValoresPredeterminados(ByVal data As DataRow, ByVal services As ServiceProvider)

        data("FechaAlta") = Date.Today
        data("RecalcularValoracion") = enumtaValoracionSalidas.taNoRecalcular
        data("CriterioValoracion") = enumtaValoracion.taPrecioEstandar
        data("Plazo") = 0
        data("PlazoFabricacion") = 0
        data("PrecioUltimaCompraA") = 0
        data("FechaEstandar") = Date.Today

        Dim AE As New ArticuloEstado
        Dim dtEstado As DataTable = AE.Filter("IDEstado", "Activo=1")
        If Not dtEstado Is Nothing AndAlso dtEstado.Rows.Count > 0 Then
            data("IDEstado") = dtEstado.Rows(0)("IDEstado")
        End If

        Dim P As New Parametro
        Dim strTipoIVA As String = P.TipoIva
        Dim strUdMedida As String = P.UdMedidaPred()

        If strTipoIVA.Length > 0 Then data("IDTipoIVA") = strTipoIVA
        If strUdMedida.Length > 0 Then
            data("IDUdInterna") = strUdMedida
            data("IDUdCompra") = strUdMedida
            data("IDUdVenta") = strUdMedida
        End If
    End Sub

#End Region

#Region "Eventos Delete "
    ''' <summary>
    ''' Relación de tareas asociadas al proceso de borrado
    ''' </summary>
    ''' <param name="deleteProcess">Proceso en el que se registran las tareas de borrado</param>
    ''' <remarks></remarks>
    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf DeleteArticulo)
    End Sub
    ''' <summary>
    ''' Borrado de artículos
    ''' </summary>
    ''' <param name="data">Registro del artículo a borrar</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub DeleteArticulo(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDArticulo")) > 0 Then
            Dim ACStd As New ArticuloCosteEstandar
            Dim DtCoste As DataTable = ACStd.Filter(New StringFilterItem("IDArticulo", FilterOperator.Equal, data("IDArticulo")))
            If Not DtCoste Is Nothing AndAlso DtCoste.Rows.Count > 0 Then
                ACStd.Delete(DtCoste)
            End If
            Dim strSQL As String = "UPDATE tbHistoricoMovimiento SET IDTipoMovimiento=11 WHERE (IDArticulo='" & data("IDArticulo") & "')"
            AdminData.Execute(strSQL)
        End If
    End Sub

#End Region

#Region "Eventos Validate "
    ''' <summary>
    ''' Relación de tareas asociadas a la validación 
    ''' </summary>
    ''' <param name="validateProcess">Proceso en el que se registran las tareas de edición</param>
    ''' <remarks></remarks>
    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidaDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidaArticulo)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidaLoteArticulo)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidaLoteNSerieArticulo)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidaCriterioValoracion)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidaArticuloPadre)
    End Sub

    <Task()> Public Shared Sub ValidaArticuloPadre(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDArticuloPadre")) > 0 AndAlso Length(data("IDArticulo")) > 0 AndAlso data("IDArticuloPadre") = data("IDArticulo") Then
            ApplicationService.GenerateError("El Artículo y el Artículo Padre no pueden ser iguales.")
        End If
    End Sub
    ''' <summary>
    ''' Comprobar que el artículo tenga tenga descripción, tipo y familia
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidaDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("DescArticulo")) = 0 Then ApplicationService.GenerateError("La descripción es un dato obligatorio.")
        If Length(data("IDTipo")) = 0 Then ApplicationService.GenerateError("El Tipo es un dato obligatorio.")
        If Length(data("IDFamilia")) = 0 Then ApplicationService.GenerateError("La Familia es un dato obligatorio.")
        Dim FilArtFam As New Filter
        FilArtFam.Add("IDTipo", FilterOperator.Equal, data("IDTipo"), FilterType.String)
        FilArtFam.Add("IDFamilia", FilterOperator.Equal, data("IDFamilia"), FilterType.String)
        Dim DtArtFam As DataTable = New Familia().Filter(FilArtFam)
        If DtArtFam Is Nothing OrElse DtArtFam.Rows.Count = 0 Then
            ApplicationService.GenerateError("El Tipo y la Familia no coinciden. Por favor, revise los datos.")
        Else
            Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
            If AppParams.GestionInventarioPermanente Then
                'Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                'Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data("IDArticulo"))
                If Length(data("IDTipo")) > 0 Then
                    Dim dtTipo As DataTable = New TipoArticulo().SelOnPrimaryKey(data("IDTipo"))
                    If dtTipo.Rows.Count > 0 AndAlso Nz(dtTipo.Rows(0)("GestionStock"), False) AndAlso Length(data("CCStocks")) = 0 Then
                        If Length(DtArtFam.Rows(0)("CCStocks")) > 0 Then data("CCStocks") = DtArtFam.Rows(0)("CCStocks")
                        If Length(data("CCStocks")) = 0 Then
                            ApplicationService.GenerateError("Debe indicar una Cuenta Contable para la Gestión de Stocks. Por favor, revise los datos.")
                        End If
                    End If
                End If
            End If
        End If
    End Sub
    ''' <summary>
    ''' Comprobar que el artículo tenga tenga un código válido y asignarlo si fuera necesario
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidaArticulo(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim strArticulo As String = IIf(Length(data("IDArticulo")) = 0, "COD_AUTOM", data("IDArticulo"))
            If strArticulo = "COD_AUTOM" Then
                Dim p As New Parametro
                Dim intTipoCodificacionAutomatica As Integer = p.TipoCodificacionAutomatica
                Select Case intTipoCodificacionAutomatica
                    Case enumTipoCodAutomatica.Familia
                        Dim f As New Familia
                        Dim dtFamilia As DataTable = f.SelOnPrimaryKey(data("IDTipo"), data("IDFamilia"))
                        If Not IsNothing(dtFamilia) AndAlso dtFamilia.Rows.Count > 0 Then
                            If Length(dtFamilia.Rows(0)("NumCorrelativo")) > 0 Then
                                data("IDArticulo") = data("IDTipo") & data("IDFamilia") & VB6.Format(dtFamilia.Rows(0)("NumCorrelativo"), "0000")
                                data("IDContador") = System.DBNull.Value
                            Else : ApplicationService.GenerateError("No se ha establecido el valor Correlativo en la configuración de la Familia: |", data("IDFamilia"))
                            End If
                        End If
                        dtFamilia.Rows(0)("NumCorrelativo") = dtFamilia.Rows(0)("NumCorrelativo") + 1
                        f.Update(dtFamilia)
                    Case enumTipoCodAutomatica.SubFamilia
                        If Length(data("IDSubfamilia")) > 0 Then
                            Dim sf As New Subfamilia
                            Dim dtSubFamilia As DataTable = sf.SelOnPrimaryKey(data("IDTipo"), data("IDFamilia"), data("IDSubfamilia"))
                            If Not IsNothing(dtSubFamilia) AndAlso dtSubFamilia.Rows.Count > 0 Then
                                If Length(dtSubFamilia.Rows(0)("NumCorrelativo")) > 0 Then
                                    data("IDArticulo") = data("IDTipo") & data("IDFamilia") & data("IDSubfamilia") & VB6.Format(dtSubFamilia.Rows(0)("NumCorrelativo"), "0000")
                                    data("IDContador") = System.DBNull.Value
                                Else : ApplicationService.GenerateError("No se ha establecido el valor Correlativo en la configuración de la SubFamilia: |", data("IDSubFamilia"))
                                End If
                            End If
                            dtSubFamilia.Rows(0)("NumCorrelativo") = dtSubFamilia.Rows(0)("NumCorrelativo") + 1
                            sf.Update(dtSubFamilia)
                        Else
                            ApplicationService.GenerateError("Para utilizar la Codificación Automática debe rellenar obligatoriamente los campos de Tipo, Familia y Subfamilia del artículo.", ".Update")
                        End If
                End Select
            ElseIf Length(data("IdContador")) = 0 And Length(data("IdArticulo")) > 0 Then
            ElseIf Length(data("IdContador")) > 0 Then
                data("IDArticulo") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data("IDContador"), services)
            Else
                Dim StDatos As New Contador.DatosDefaultCounterValue
                StDatos.row = data
                StDatos.EntityName = "Articulo"
                StDatos.FieldName = "IDArticulo"
                ProcessServer.ExecuteTask(Of Contador.DatosDefaultCounterValue)(AddressOf Contador.LoadDefaultCounterValue, StDatos, services)
            End If
            Dim dtArt As DataTable = New Articulo().SelOnPrimaryKey(data("IDArticulo"))
            If Not dtArt Is Nothing AndAlso dtArt.Rows.Count > 0 Then
                ApplicationService.GenerateError("El Artículo introducido ya existe.")
            End If
        End If
    End Sub
    ''' <summary>
    ''' Comprobar si el artículo antes tenía gestión por lotes
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidaLoteArticulo(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Modified Then
            If data("GestionStockPorLotes") = False AndAlso data("GestionStockPorLotes", DataRowVersion.Original) Then
                Dim ClsArtAlm As New ArticuloAlmacen
                Dim DtArtAlm As DataTable = ClsArtAlm.Filter(New FilterItem("IDArticulo", FilterOperator.Equal, data("IDArticulo"), FilterType.String))
                If Not DtArtAlm Is Nothing AndAlso DtArtAlm.Rows.Count > 0 Then
                    For Each drArtAlm As DataRow In DtArtAlm.Select
                        If drArtAlm("StockFisico") <> 0 Then
                            ApplicationService.GenerateError("No se puede quitar la gestión de Stocks por Lotes si tiene todavia Stock en sus almacenes.|Por favor, revise los datos", vbNewLine)
                        End If
                    Next
                End If
            End If
            If data("GestionStockPorLotes") = True AndAlso Not data("GestionStockPorLotes", DataRowVersion.Original) Then
                Dim ClsArtAlm As New ArticuloAlmacen
                Dim DtArtAlm As DataTable = ClsArtAlm.Filter(New FilterItem("IDArticulo", FilterOperator.Equal, data("IDArticulo"), FilterType.String))
                If Not DtArtAlm Is Nothing AndAlso DtArtAlm.Rows.Count > 0 Then
                    For Each drArtAlm As DataRow In DtArtAlm.Select
                        If drArtAlm("StockFisico") <> 0 Then
                            ApplicationService.GenerateError("No se puede poner la gestión de Stocks por Lotes si tiene Stock en sus almacenes.|Por favor, revise los datos", vbNewLine)
                        End If
                    Next
                End If
            End If
        End If
    End Sub
    ''' <summary>
    ''' Validacion Gestion lotes y numeros de serie
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidaLoteNSerieArticulo(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Nz(data("GestionStockPorLotes"), False) > 0 AndAlso Nz(data("NSerieObligatorio"), False) > 0 Then
            ApplicationService.GenerateError("La gestión de stock por lotes es incompatible con la gestión de números de serie para el artículo |.", data("IDArticulo"))
        End If
    End Sub
    ''' <summary>
    ''' Comprobar que el criterio de valoración sea válido
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidaCriterioValoracion(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("CriterioValoracion")) > 0 AndAlso _
          data("CriterioValoracion") <> enumtaValoracion.taPrecioEstandar AndAlso _
          data("CriterioValoracion") <> enumtaValoracion.taPrecioFIFOFecha AndAlso _
          data("CriterioValoracion") <> enumtaValoracion.taPrecioFIFOMvto AndAlso _
          data("CriterioValoracion") <> enumtaValoracion.taPrecioMedio AndAlso _
          data("CriterioValoracion") <> enumtaValoracion.taPrecioUltCompra Then
            ApplicationService.GenerateError("El criterio de valoración no es válido.")
        End If
    End Sub

    <Task()> Public Shared Function EsPrecinta(ByVal idArticulo As String, ByVal services As ServiceProvider) As Boolean
        Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
        If AppParams.GestionBodegas Then
            If (String.IsNullOrEmpty(idArticulo)) Then
                Return False
            End If
            Dim filter As New Filter
            filter.Add(_Articulo.IDArticulo, idArticulo)
            Dim dttArticulo As DataTable = New DataEngine().Filter("vBdgArticuloPrecinta", filter)
            If (dttArticulo Is Nothing OrElse dttArticulo.Rows.Count = 0) Then
                Return False
            End If
            Return dttArticulo.Rows(0)("Precinta") Or dttArticulo.Rows(0)("PrecintaSubfamilia")
        End If
    End Function

#End Region

#Region "Eventos Update "
    ''' <summary>
    ''' Relación de tareas asociadas a la edición 
    ''' </summary>
    ''' <param name="validateProcess">Proceso en el que se registran las tareas de edición</param>
    ''' <remarks></remarks>
    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf ActualizarCriterioValoracion)
        updateProcess.AddTask(Of DataRow)(AddressOf ActualizarEstadoArticulo)
        updateProcess.AddTask(Of DataRow)(AddressOf ActualizarUnidadesMedida)
        updateProcess.AddTask(Of DataRow)(AddressOf ActualizarHomologacion)
        updateProcess.AddTask(Of DataRow)(AddressOf ActualizarParametrosProduccion)
        updateProcess.AddTask(Of DataRow)(AddressOf ValidarArticuloPorContador)
        updateProcess.AddTask(Of DataRow)(AddressOf ComprobarArticuloPadre)
        updateProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.UpdateEntityRow)
        updateProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.MarcarComoActualizado)
        'updateProcess.AddTask(Of DataRow)(AddressOf AñadirCaracteristicas)
        updateProcess.AddTask(Of DataRow)(AddressOf AñadirCosteEstandar)
        updateProcess.AddTask(Of DataRow)(AddressOf ArticuloNivelRevision)
        updateProcess.AddTask(Of DataRow)(AddressOf ArticuloFinal)
        updateProcess.AddTask(Of DataRow)(AddressOf AñadirAlmacenPredeterminado)

    End Sub
    ''' <summary>
    ''' Actualizar el criterio de valoración
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks>/remarks>SI viene vacío se obtendrá el criterio asociado al tipo al que pertenece
    <Task()> Public Shared Sub ActualizarCriterioValoracion(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("RecalcularValoracion")) = 0 OrElse Length(data("CriterioValoracion")) = 0 Then
                Dim dtTipo As DataTable = New TipoArticulo().SelOnPrimaryKey(data("IDTipo"))
                If Not IsNothing(dtTipo) AndAlso dtTipo.Rows.Count > 0 Then
                    If Length(data("RecalcularValoracion")) = 0 Then data("RecalcularValoracion") = dtTipo.Rows(0)("RecalcularValoracion")
                    If Length(data("CriterioValoracion")) = 0 Then data("CriterioValoracion") = dtTipo.Rows(0)("CriterioValoracion")
                Else
                    data("RecalcularValoracion") = enumtaValoracionSalidas.taNoRecalcular
                    data("CriterioValoracion") = enumtaValoracion.taPrecioEstandar
                End If
            End If
        End If
    End Sub
    ''' <summary>
    ''' Actualizar el estado activo por defecto
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks>/remarks>SI viene vacío se obtendrá el estado que esté activo
    <Task()> Public Shared Sub ActualizarEstadoArticulo(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDEstado")) = 0 Then
                Dim dtEstado As DataTable = New ArticuloEstado().Filter("IDEstado", "Activo=1")
                If Not dtEstado Is Nothing AndAlso dtEstado.Rows.Count > 0 Then
                    data("IDEstado") = dtEstado.Rows(0)("IDEstado")
                End If
            End If

        End If
    End Sub
    ''' <summary>
    ''' Actualizar las unidades de medida
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks>/remarks>
    <Task()> Public Shared Sub ActualizarUnidadesMedida(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDUDInterna")) = 0 OrElse Length(data("IDUDCompra")) = 0 OrElse Length(data("IDUDVenta")) = 0 Then
                Dim strUDMedida As String = New Parametro().UdMedidaPred
                If Len(strUDMedida) > 0 Then
                    If Length(data("IDUDInterna")) = 0 Then data("IDUDInterna") = strUDMedida
                    If Length(data("IDUDCompra")) = 0 Then data("IDUDCompra") = strUDMedida
                    If Length(data("IDUDVenta")) = 0 Then data("IDUDVenta") = strUDMedida
                Else
                    ApplicationService.GenerateError("La Unidad Interna no es válida. Compruebe la configuración del parámetro Unidad Interna Predeterminada.")
                End If
            End If
        End If
    End Sub
    ''' <summary>
    ''' Actualizar el estado de homologacion
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks>/remarks>
    <Task()> Public Shared Sub ActualizarHomologacion(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim strIDEstadoHomologacion As String = New Parametro().CalidadEstadoHomologacionPorDefecto
            If Length(data("IDEstadoHomologacion")) = 0 AndAlso Length(strIDEstadoHomologacion) > 0 Then
                data("IDEstadoHomologacion") = strIDEstadoHomologacion
            End If
        End If
    End Sub
    ''' <summary>
    ''' Actualizar parámetros producción
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks>/remarks>
    <Task()> Public Shared Sub ActualizarParametrosProduccion(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim strControlProduccion As String = New Parametro().ControlProduccion
            If Length(data("ParamMaterial")) = 0 Then data("ParamMaterial") = CInt(Left(strControlProduccion, 1))
            If Length(data("ParamTerminado")) = 0 Then data("ParamTerminado") = CInt(Right(strControlProduccion, 1))
        End If
    End Sub
    ''' <summary>
    ''' Añadir las características asociadas a ese tipo familia
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks>/remarks>
    <Task()> Public Shared Sub AñadirCaracteristicas(ByVal idarticulo As String, ByVal services As ServiceProvider)
        If Length(idarticulo) > 0 Then
            'Controlar que no me llega una de presentación
            'Dim upPresenta As UpdatePackage = services.GetService(Of UpdatePackage)()
            'Dim dt As DataTable = upPresenta.Item(GetType(ArticuloCaracteristica).Name).First
            'Dim adr() As DataRow
            'If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            '    adr = dt.Select("IDArticulo='" & Data("IDArticulo") & "'")
            'End If
            'If adr Is Nothing OrElse adr.Length = 0 Then
            Dim DtArt As DataTable = New Articulo().SelOnPrimaryKey(idarticulo)
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ArticuloCaracteristica.AddArticuloCaracteristica, DtArt.Rows(0), services)
            'End If
        End If
    End Sub
    ''' <summary>
    ''' Añadir la información del coste estándar
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks>/remarks>
    <Task()> Public Shared Sub AñadirCosteEstandar(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added OrElse (data.RowState = DataRowState.Modified AndAlso AreDifferents(data("PrecioEstandarA"), data("PrecioEstandarA", DataRowVersion.Original))) Then
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ArticuloCosteEstandar.AddArticuloCosteEstandar, data, services)
        End If
    End Sub
    ''' <summary>
    ''' Añadir la información del almacén predeterminado si llevamos gestión de stocks
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks>/remarks>
    <Task()> Public Shared Sub AñadirAlmacenPredeterminado(ByVal data As DataRow, ByVal services As ServiceProvider)

        If ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf New Articulo().ArticuloGestionStocks, data("IDArticulo"), services) Then
            'Controlar que no me llega una de presentación
            Dim upPresenta As UpdatePackage = services.GetService(Of UpdatePackage)()
            Dim dt As DataTable = upPresenta.Item(GetType(ArticuloAlmacen).Name).First
            Dim adr() As DataRow
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                adr = dt.Select("IDArticulo='" & data("IDArticulo") & "'")
            End If
            If adr Is Nothing OrElse adr.Length = 0 Then
                Dim dtAlmacenes As DataTable = New ArticuloAlmacen().Filter(New StringFilterItem("IDArticulo", data("IDArticulo")))
                If IsNothing(dtAlmacenes) OrElse dtAlmacenes.Rows.Count = 0 Then
                    '//Añadir el almacén predeterminado para el artículo indicado.
                    ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf ArticuloAlmacen.AddAlmacenPredeterminadoArticulo, data("IDArticulo"), services)
                End If
            End If
        End If

    End Sub

    <Task()> Public Shared Sub ValidarArticuloPorContador(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDArticulo")) = 0 Then
                If Length(data("IDContador")) = 0 Then
                    ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarContadorPredeterminado, data, services)
                End If
                If Length(data("IdContador")) > 0 Then data("IDArticulo") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data("IDContador"), services)
                If Length(data("IDArticulo")) = 0 Then ApplicationService.GenerateError("No existe un contador bien configurado para esta entidad.")
                ProcessServer.ExecuteTask(Of DataRow)(AddressOf ValidaArticuloPadre, data, services)
            Else
                Dim dtArticulo As DataTable = New Articulo().SelOnPrimaryKey(data("IDArticulo"))
                If Not IsNothing(dtArticulo) AndAlso dtArticulo.Rows.Count > 0 Then
                    ApplicationService.GenerateError("El Artículo '|' ya existe en la Base de Datos.", data("IDArticulo"))
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ComprobarArticuloPadre(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDArticuloPadre")) > 0 Then
            If CType(data("IDArticulo"), String).Equals(CType(data("IDArticuloPadre"), String), StringComparison.OrdinalIgnoreCase) Then
                ApplicationService.GenerateError("El artículo padre no puede ser el mismo que el artículo actual.")
            End If
            Dim control As DataTable = New Articulo().SelOnPrimaryKey(data("IDArticuloPadre"))
            If control.Rows.Count > 0 Then
                If Length(control.Rows(0)("IDArticuloPadre")) > 0 Then
                    ApplicationService.GenerateError("El artículo seleccionado como padre ya depende del artículo padre |.", control.Rows(0)("IDArticuloPadre"))
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' Asignar la información por defecto
    ''' </summary>
    ''' <param name="data">Registro Nuevo</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub AsignarContadorPredeterminado(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim Contador As New Contador
        Dim StDatos As New Contador.DatosDefaultCounterValue
        StDatos.row = data
        StDatos.EntityName = "Articulo"
        StDatos.FieldName = "IDArticulo"
        ProcessServer.ExecuteTask(Of Contador.DatosDefaultCounterValue)(AddressOf Contador.LoadDefaultCounterValue, StDatos, services)
    End Sub

    ''' <summary>
    ''' Añadir la información Niveles de revision - Calidad
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks>/remarks>
    <Task()> Public Shared Sub ArticuloNivelRevision(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Modified Then
            If Length(data("NivelModificacionPlan")) > 0 AndAlso AreDifferents(data("NivelModificacionPlan"), data("NivelModificacionPlan", DataRowVersion.Original)) Then
                ProcessServer.ExecuteTask(Of DataRow)(AddressOf ActualizarArticuloNivelRevision, data, services)
            End If
        End If
    End Sub
    <Task()> Public Shared Sub ActualizarArticuloNivelRevision(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim ANR As BusinessHelper
        ANR = BusinessHelper.CreateBusinessObject("ArticuloNivelRevision")

        Dim dtNivel As DataTable = ANR.SelOnPrimaryKey(data("IDArticulo"), data("NivelModificacionPlan"))
        If IsNothing(dtNivel) OrElse dtNivel.Rows.Count = 0 Then
            Dim drNivel As DataRow = dtNivel.NewRow

            drNivel("IDArticulo") = data("IDArticulo")
            drNivel("NivelPlano") = data("NivelModificacionPlan")
            drNivel("IDEstadoHomologacion") = data("IDEstadoHomologacion")
            drNivel("FechaCambioEstado") = Now()

            dtNivel.Rows.Add(drNivel)
            ANR.Update(dtNivel)
        End If
    End Sub
    ''' <summary>
    ''' Añadir la información Niveles de revision - Calidad
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks>/remarks>
    <Task()> Public Shared Sub ArticuloFinal(ByVal data As DataRow, ByVal services As ServiceProvider)

        If Length(data("IDArticuloFinal")) = 0 Then
            data("GenerarOFArticuloFinal") = False
        Else
            If data.RowState = DataRowState.Modified Then
                If data("IDArticuloFinal") <> data("IDArticuloFinal", DataRowVersion.Original) & String.Empty Then
                    Dim dtArt As DataTable = New Articulo().SelOnPrimaryKey(data("IDArticuloFinal"))
                    If dtArt Is Nothing OrElse dtArt.Rows.Count = 0 Then
                        ApplicationService.GenerateError("El artículo | no existe.", "'" & data("IDArticuloFinal") & "'")
                    End If
                    data("GenerarOFArticuloFinal") = True
                End If
            End If
        End If
    End Sub
    <Task()> Public Shared Sub Configurador(ByVal data As DataRow, ByVal services As ServiceProvider)
        '    ElseIf dr.RowState = DataRowState.Modified Then


        'strIDFamiliaConfiguracionOLD = dr("IDFamiliaConfiguracion", DataRowVersion.Original) & String.Empty

        'strTipoEstructuraOLD = dr("IDTipoEstructura", DataRowVersion.Original) & String.Empty
        'strNivelPlanoOLD = dr("NivelModificacionPlan", DataRowVersion.Original) & String.Empty

        '        Else
        'ApplicationService.GenerateError("El Tipo y la Familia no coinciden. Por favor, revise los datos.")
        '        End If
        'If Length(dr("PrecioBase")) > 0 AndAlso Length(dr("PrecioBase", DataRowVersion.Original)) = 0 OrElse dr("PrecioBase", DataRowVersion.Original) <> dr("PrecioBase") Then
        '    If IsNothing(mdtUpdateDatosConfigurador) Then mdtUpdateDatosConfigurador = Me.AddNew
        '    mdtUpdateDatosConfigurador.ImportRow(dr)
        'End If

        '        End If

        '' En caso de que se agregue un Tipo de Estructura habrá de validarse que no hayan
        '' componentes que formen parte de la estructura del artículo.
        'If Nz(dr("TipoEstructura"), 0) AndAlso Length(dr("IDTipoEstructura")) = 0 Then
        '    ApplicationService.GenerateError("Debe de especificar un código de Tipo de Estructura.")
        'End If
        'If Length(dr("IDTipoEstructura")) > 0 AndAlso dr("IDTipoEstructura") <> strTipoEstructuraOLD Then
        '    e.ExisteComoPadre(dr)
        'End If

        'If Length(dr("IDFamiliaConfiguracion")) > 0 AndAlso dr("IDFamiliaConfiguracion") <> strIDFamiliaConfiguracionOLD Then
        '    'ActualizarDatosConfigurador(dr)
        '    If IsNothing(mdtAddDatosConfigurador) Then mdtAddDatosConfigurador = Me.AddNew
        '    mdtAddDatosConfigurador.ImportRow(dr)
        'End If


        'ActualizarDatosConfigurador(mdtAddDatosConfigurador)
        'RecalcularPreciosConfigurador(mdtUpdateDatosConfigurador)
    End Sub

    <Task()> Public Shared Function ArticuloGestionStocks(ByVal strIDArticulo As String, ByVal services As ServiceProvider) As Boolean
        If Length(strIDArticulo) > 0 Then
            Dim f As New Filter
            f.Add(New StringFilterItem("IDArticulo", strIDArticulo))
            Dim dt As DataTable = New BE.DataEngine().Filter("vNegArticuloGestionStocks", f)
            If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
                ArticuloGestionStocks = Nz(dt.Rows(0)("GestionStock"), False)
            End If
        End If
    End Function

#Region " ActualizarDatosConfigurador "

    Private Sub ActualizarDatosConfigurador(ByVal dt As DataTable)
        Dim AC As BusinessHelper
        AC = BusinessHelper.CreateBusinessObject("CfgArticuloCaractConfg")
        CType(AC, IConfiguradorArticulo).ADDCaracteristicas(dt)
    End Sub

    Private Sub RecalcularPreciosConfigurador(ByVal dt As DataTable)
        Dim AC As BusinessHelper
        AC = BusinessHelper.CreateBusinessObject("CfgArticuloCaractConfg")
        CType(AC, IConfiguradorArticulo).RecalcularPreciosCaracteristicas(dt)
    End Sub

#End Region

#End Region

#Region "Eventos GetBusinessRules"
    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("IdTipo", AddressOf CambioTipo)
        oBrl.Add("PrecioEstandarA", AddressOf CambioPrecioEstandarA)
        oBrl.Add("ValorReposicionA", AddressOf CambioValorReposicionA)
        oBrl.Add("NSerieObligatorio", AddressOf CambioNSerieObligatorio)
        oBrl.Add("IDTipoIVA", AddressOf CambioTipoIVA)
        oBrl.Add("IDPartidaEstadistica", AddressOf CambioPartidaEstadistica)
        oBrl.Add("IDArticuloPadre", AddressOf CambioArticuloPadre)
        Return oBrl
    End Function
    <Task()> Public Shared Sub CambioPartidaEstadistica(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim pe As New PartidaEstadistica
            Dim dtPE As DataTable = pe.SelOnPrimaryKey(data.Value)
            If IsNothing(dtPE) OrElse dtPE.Rows.Count = 0 Then
                ApplicationService.GenerateError("La Partida Estadística  no existe en la Base de datos.")
            End If
        End If
    End Sub
    <Task()> Public Shared Sub CambioTipoIVA(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim ti As New TipoIva
            Dim dtIVA As DataTable = ti.SelOnPrimaryKey(data.Value)
            If IsNothing(dtIVA) OrElse dtIVA.Rows.Count = 0 Then
                ApplicationService.GenerateError("El tipo de IVA introducido no existe.")
            End If
        End If
    End Sub
    <Task()> Public Shared Sub CambioNSerieObligatorio(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If IsNumeric(data.Value) Then
            data.Current("NSerieObligatorio") = data.Value
            ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf TratarNSerieObligatorio, data.Current, services)
        Else
            ApplicationService.GenerateError("Campo no numérico.")
        End If
    End Sub
    <Task()> Public Shared Sub CambioValorReposicionA(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If IsNumeric(data.Value) Then
            data.Current("ValorReposicionB") = ProcessServer.ExecuteTask(Of Double, Double)(AddressOf CalcularImporteEnMonedaB, data.Value, services)
            data.Current("FechaValorReposicion") = DateTime.Today
        Else
            ApplicationService.GenerateError("Campo no numérico.")
        End If
    End Sub
    <Task()> Public Shared Sub CambioPrecioEstandarA(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If IsNumeric(data.Value) Then
            data.Current("PrecioEstandarB") = ProcessServer.ExecuteTask(Of Double, Double)(AddressOf CalcularImporteEnMonedaB, data.Value, services)
            data.Current("FechaEstandar") = DateTime.Today
        Else
            ApplicationService.GenerateError("Campo no numérico.")
        End If
    End Sub
    <Task()> Public Shared Sub CambioTipo(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim t As New TipoArticulo
            Dim dr As DataRow = t.GetItemRow(data.Value)

            data.Current("RecalcularValoracion") = dr("RecalcularValoracion")
            data.Current("CriterioValoracion") = dr("CriterioValoracion")
            data.Current("IDFamilia") = DBNull.Value
            data.Current("IDSubFamilia") = DBNull.Value
        End If
    End Sub
    <Task()> Public Shared Sub TratarNSerieObligatorio(ByVal current As IPropertyAccessor, ByVal services As ServiceProvider)

        If Length(current("IDArticulo")) >= 0 Then
            Dim f As New Filter
            f.Add(New StringFilterItem("IDArticulo", current("IDArticulo")))
            Dim dt As DataTable

            If current("NSerieObligatorio") Then
                dt = New BE.DataEngine().Filter("vNegStockTotalPorArticulo", f)
                If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                    If dt.Rows(0)("StockTotal") > 0 Then
                        ApplicationService.GenerateError("Existe Stock en los Almacenes. Ponerlos a cero e inventariarlos.")
                    End If
                End If
            Else
                Dim Act As New Activo
                dt = Act.Filter(f)
                If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                    ApplicationService.GenerateError("El artículo está asociado a más de un activo. Previamente hay que eliminar esta relación.")
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioArticuloPadre(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            If Length(data.Current("IDArticulo")) > 0 Then
                If CType(data.Current("IDArticulo"), String).Equals(CType(data.Value, String), StringComparison.OrdinalIgnoreCase) Then
                    ApplicationService.GenerateError("El artículo padre no puede ser el mismo que el artículo actual.")
                End If
            End If
            Dim control As DataTable = New Articulo().SelOnPrimaryKey(data.Value)
            If control.Rows.Count > 0 Then
                If Length(control.Rows(0)("IDArticuloPadre")) > 0 Then
                    ApplicationService.GenerateError("El artículo seleccionado como padre ya depende del artículo padre |.", control.Rows(0)("IDArticuloPadre"))
                End If
            End If
        End If
    End Sub
#End Region

#Region " Copia Articulo "
    <Serializable()> _
      Public Class DatosArtCopia
        Public IDArticulo As String
        Public DescArticulo As String
        Public IDContador As String
        Public BlnCopyCaractMaq As Boolean
        Public BlnCopyCaract As Boolean
        Public BlnCopyEsp As Boolean
        Public BlnCopyPromo As Boolean
        Public BlnCopyIdio As Boolean
        Public BlnCopyDoc As Boolean
        Public BlnCopyAna As Boolean
        Public BlnCopyCostesVar As Boolean
        Public BlnCopyRu As Boolean
        Public BlnCopyEst As Boolean
        Public BlnCopyProv As Boolean
        Public BlnCopyTar As Boolean
        Public BlnCopyClie As Boolean
        Public BlnCopyUd As Boolean
        Public BlnCopyAlm As Boolean
        Public IDArticuloNew As String
        Public IDArticuloPadre As String
        Public IDCaracteristicaArticulo1 As String
        Public IDCaracteristicaArticulo2 As String
        Public IDCaracteristicaArticulo3 As String
        Public IDCaracteristicaArticulo4 As String
        Public IDCaracteristicaArticulo5 As String

        Public dtRutaOrigen As DataTable
        Public dtArticuloNew As DataTable
        Public dtCosteEstandarNew As DataTable
        Public dtEstructurasNew, dtEstructuraNew As DataTable
        Public dtArticuloRutaNew, dtRutasNew, dtRutaParametroNew, dtRutaUtillajeNew, dtRutaOficioNew, dtRutasAlternativasNew, dtRutaAMFENew, dtRutaProveedorNew, dtRutaProveedorLineasNew As DataTable
        Public dtNserieCaracteristicasNew As DataTable
        Public dtCaracteristicasNew As DataTable
        Public dtPlantillasNew As DataTable
        Public dtPromocionesLineaNew As DataTable
        Public dtIdiomasNew As DataTable
        Public dtPlanosNew As DataTable
        Public dtAnaliticaNew As DataTable
        Public dtCostesVariosNew As DataTable
        Public dtProveedoresNew, dtProveedoresLineasNew As DataTable
        Public dtTarifasNew, dtTarifasLineasNew As DataTable
        Public dtClientesNew, dtClientesLineasNew As DataTable
        Public dtUnidadesABNew As DataTable
        Public dtAlmacenesNew As DataTable
        Public HasRutaNew As Hashtable
    End Class
    <Task()> Public Shared Function CopiaArticulo(ByVal data As DatosArtCopia, ByVal services As ServiceProvider) As String
        If Len(data.IDArticulo) > 0 Then
            ProcessServer.ExecuteTask(Of DatosArtCopia)(AddressOf GeneraNuevoArticulo, data, services)
            If Length(data.IDArticuloNew) > 0 Then
                ProcessServer.ExecuteTask(Of DatosArtCopia)(AddressOf CopiarArticuloCosteEstandar, data, services)
                If data.BlnCopyEst Then
                    ProcessServer.ExecuteTask(Of DatosArtCopia)(AddressOf CopiarArticuloEstructura, data, services)
                    ProcessServer.ExecuteTask(Of DatosArtCopia)(AddressOf CopiarEstructura, data, services)
                End If
                If data.BlnCopyRu Then
                    ProcessServer.ExecuteTask(Of DatosArtCopia)(AddressOf CopiarArticuloRuta, data, services)
                    ProcessServer.ExecuteTask(Of DatosArtCopia)(AddressOf CopiarRuta, data, services)
                    ProcessServer.ExecuteTask(Of DatosArtCopia)(AddressOf CopiarRutaParametro, data, services)
                    ProcessServer.ExecuteTask(Of DatosArtCopia)(AddressOf CopiarRutaUtillaje, data, services)
                    ProcessServer.ExecuteTask(Of DatosArtCopia)(AddressOf CopiarRutaOficio, data, services)
                    ProcessServer.ExecuteTask(Of DatosArtCopia)(AddressOf CopiarRutaAlternativa, data, services)
                    ProcessServer.ExecuteTask(Of DatosArtCopia)(AddressOf CopiarRutaAMFE, data, services)
                End If
                If data.BlnCopyCaractMaq Then
                    ProcessServer.ExecuteTask(Of DatosArtCopia)(AddressOf CopiarCaracteristicasMaq, data, services)
                End If
                If data.BlnCopyCaract Then
                    ProcessServer.ExecuteTask(Of DatosArtCopia)(AddressOf CopiarCaracteristicas, data, services)
                End If
                If data.BlnCopyEsp Then
                    ProcessServer.ExecuteTask(Of DatosArtCopia)(AddressOf CopiarEspecificaciones, data, services)
                End If
                If data.BlnCopyPromo Then
                    ProcessServer.ExecuteTask(Of DatosArtCopia)(AddressOf CopiarPromociones, data, services)
                End If
                If data.BlnCopyIdio Then
                    ProcessServer.ExecuteTask(Of DatosArtCopia)(AddressOf CopiarIdiomas, data, services)
                End If
                If data.BlnCopyDoc Then
                    ProcessServer.ExecuteTask(Of DatosArtCopia)(AddressOf CopiarDocumentos, data, services)
                End If
                If data.BlnCopyAna Then
                    ProcessServer.ExecuteTask(Of DatosArtCopia)(AddressOf CopiarAnalitica, data, services)
                End If
                If data.BlnCopyCostesVar Then
                    ProcessServer.ExecuteTask(Of DatosArtCopia)(AddressOf CopiarCostesVarios, data, services)
                End If
                If data.BlnCopyProv Then
                    ProcessServer.ExecuteTask(Of DatosArtCopia)(AddressOf CopiarProveedores, data, services)
                    ProcessServer.ExecuteTask(Of DatosArtCopia)(AddressOf CopiarProveedoresLineas, data, services)
                End If
                If data.BlnCopyTar Then
                    ProcessServer.ExecuteTask(Of DatosArtCopia)(AddressOf CopiarTarifasArt, data, services)
                    ProcessServer.ExecuteTask(Of DatosArtCopia)(AddressOf CopiarTarifasArtLineas, data, services)
                End If
                If data.BlnCopyClie Then
                    ProcessServer.ExecuteTask(Of DatosArtCopia)(AddressOf CopiarClientes, data, services)
                    ProcessServer.ExecuteTask(Of DatosArtCopia)(AddressOf CopiarClientesLineas, data, services)
                End If
                If data.BlnCopyUd Then
                    ProcessServer.ExecuteTask(Of DatosArtCopia)(AddressOf CopiarUnidades, data, services)
                End If
                If data.BlnCopyAlm Then
                    ProcessServer.ExecuteTask(Of DatosArtCopia)(AddressOf CopiarAlmacenes, data, services)
                End If
                ProcessServer.ExecuteTask(Of DatosArtCopia)(AddressOf GuardarCopiaArticulo, data, services)

                Return data.IDArticuloNew
            Else
                Return String.Empty
            End If
        Else
            Return String.Empty
        End If
    End Function

    <Task()> Public Shared Sub GeneraNuevoArticulo(ByVal data As DatosArtCopia, ByVal services As ServiceProvider)
        Dim A As New Articulo
        Dim dt As DataTable = A.SelOnPrimaryKey(data.IDArticulo)
        If dt.Rows.Count > 0 Then
            data.dtArticuloNew = A.AddNewForm
            For Each col As DataColumn In data.dtArticuloNew.Columns
                If dt.Columns.Contains(col.ColumnName) Then
                    data.dtArticuloNew.Rows(0)(col.ColumnName) = dt.Rows(0)(col.ColumnName)
                End If
            Next
            'data.dtArticuloNew.Rows.Add(dt.Rows(0).ItemArray)
            data.dtArticuloNew.Rows(0)("PrecioUltimaCompraA") = 0
            data.dtArticuloNew.Rows(0)("PrecioUltimaCompraB") = 0
            data.dtArticuloNew.Rows(0)("FechaUltimaCompra") = DBNull.Value
            data.dtArticuloNew.Rows(0)("IDProveedorUltimaCompra") = DBNull.Value
            data.dtArticuloNew.Rows(0)("FechaAlta") = Date.Today
            data.dtArticuloNew.Rows(0)("DescArticulo") = data.DescArticulo
            If Len(data.IDArticuloNew) > 0 Then
                data.dtArticuloNew.Rows(0)("IDArticulo") = data.IDArticuloNew
                If Len(data.IDContador) > 0 Then data.dtArticuloNew.Rows(0)("IDContador") = data.IDContador
            Else
                If Len(data.IDContador) > 0 Then
                    data.dtArticuloNew.Rows(0)("IDArticulo") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data.IDContador, services)
                    data.dtArticuloNew.Rows(0)("IDContador") = data.IDContador
                    data.IDArticuloNew = data.dtArticuloNew.Rows(0)("IDArticulo")
                Else
                    data.dtArticuloNew.Rows(0)("IDArticulo") = DBNull.Value
                    data.dtArticuloNew.Rows(0)("IDContador") = DBNull.Value
                    ProcessServer.ExecuteTask(Of DataRow)(AddressOf ValidaArticulo, data.dtArticuloNew.Rows(0), services)
                End If
            End If
            data.dtArticuloNew.Rows(0)("IDArticuloPadre") = data.IDArticuloPadre
            If Length(data.IDCaracteristicaArticulo1) > 0 Then data.dtArticuloNew.Rows(0)("IDCaracteristicaArticulo1") = data.IDCaracteristicaArticulo1
            If Length(data.IDCaracteristicaArticulo2) > 0 Then data.dtArticuloNew.Rows(0)("IDCaracteristicaArticulo2") = data.IDCaracteristicaArticulo2
            If Length(data.IDCaracteristicaArticulo3) > 0 Then data.dtArticuloNew.Rows(0)("IDCaracteristicaArticulo3") = data.IDCaracteristicaArticulo3
            If Length(data.IDCaracteristicaArticulo4) > 0 Then data.dtArticuloNew.Rows(0)("IDCaracteristicaArticulo4") = data.IDCaracteristicaArticulo4
            If Length(data.IDCaracteristicaArticulo5) > 0 Then data.dtArticuloNew.Rows(0)("IDCaracteristicaArticulo5") = data.IDCaracteristicaArticulo5

            data.IDArticuloNew = data.dtArticuloNew.Rows(0)("IDArticulo")
        End If
    End Sub

    <Task()> Public Shared Sub CopiarArticuloCosteEstandar(ByVal data As DatosArtCopia, ByVal services As ServiceProvider)
        Dim A As New ArticuloCosteEstandar
        Dim dt As DataTable = A.Filter(New StringFilterItem("IDArticulo", data.IDArticulo))
        If dt.Rows.Count > 0 Then
            data.dtCosteEstandarNew = A.AddNew
            Dim drNew As DataRow = data.dtCosteEstandarNew.NewRow
            'drNew.ItemArray = dt.Rows(0).ItemArray
            For Each col As DataColumn In data.dtCosteEstandarNew.Columns
                If dt.Columns.Contains(col.ColumnName) Then
                    drNew(col.ColumnName) = dt.Rows(0)(col.ColumnName)
                End If
            Next
            drNew("IDArticulo") = data.IDArticuloNew
            data.dtCosteEstandarNew.Rows.Add(drNew)
        End If
    End Sub

#Region " Copia Estructuras "

    <Task()> Public Shared Sub CopiarArticuloEstructura(ByVal data As DatosArtCopia, ByVal services As ServiceProvider)
        Dim A As New ArticuloEstructura
        Dim dt As DataTable = A.Filter(New StringFilterItem("IDArticulo", data.IDArticulo))
        If dt.Rows.Count > 0 Then
            data.dtEstructurasNew = A.AddNew()
            For Each dr As DataRow In dt.Select
                Dim drNew As DataRow = data.dtEstructurasNew.NewRow
                'drNew.ItemArray = dr.ItemArray
                For Each col As DataColumn In data.dtEstructurasNew.Columns
                    If dt.Columns.Contains(col.ColumnName) Then
                        drNew(col.ColumnName) = dr(col.ColumnName)
                    End If
                Next
                drNew("IDArticulo") = data.IDArticuloNew

                data.dtEstructurasNew.Rows.Add(drNew)
            Next
        End If
    End Sub

    <Task()> Public Shared Sub CopiarEstructura(ByVal data As DatosArtCopia, ByVal services As ServiceProvider)
        Dim A As New Estructura
        Dim dt As DataTable = A.Filter(New StringFilterItem("IDArticulo", data.IDArticulo))
        If dt.Rows.Count > 0 Then
            data.dtEstructuraNew = A.AddNew()
            For Each dr As DataRow In dt.Select
                Dim drNew As DataRow = data.dtEstructuraNew.NewRow
                ' drNew.ItemArray = dr.ItemArray
                For Each col As DataColumn In data.dtEstructuraNew.Columns
                    If dt.Columns.Contains(col.ColumnName) Then
                        drNew(col.ColumnName) = dr(col.ColumnName)
                    End If
                Next
                drNew("IDEstrComp") = AdminData.GetAutoNumeric
                drNew("IDArticulo") = data.IDArticuloNew

                data.dtEstructuraNew.Rows.Add(drNew)
            Next
        End If
    End Sub

#End Region

#Region " Copia Rutas "

    <Task()> Public Shared Sub CopiarArticuloRuta(ByVal data As DatosArtCopia, ByVal services As ServiceProvider)
        Dim A As New ArticuloRuta
        Dim dt As DataTable = A.Filter(New StringFilterItem("IDArticulo", data.IDArticulo))
        If dt.Rows.Count > 0 Then
            data.dtArticuloRutaNew = A.AddNew()
            For Each dr As DataRow In dt.Select
                Dim drNew As DataRow = data.dtArticuloRutaNew.NewRow
                ' drNew.ItemArray = dr.ItemArray
                For Each col As DataColumn In data.dtArticuloRutaNew.Columns
                    If dt.Columns.Contains(col.ColumnName) Then
                        drNew(col.ColumnName) = dr(col.ColumnName)
                    End If
                Next
                drNew("IDArticulo") = data.IDArticuloNew

                data.dtArticuloRutaNew.Rows.Add(drNew)
            Next
        End If
    End Sub

    <Task()> Public Shared Sub CopiarRuta(ByVal data As DatosArtCopia, ByVal services As ServiceProvider)
        Dim A As New Ruta
        data.dtRutaOrigen = A.Filter(New StringFilterItem("IDArticulo", data.IDArticulo))
        If data.dtRutaOrigen.Rows.Count > 0 Then
            data.dtRutasNew = A.AddNew

            Dim RP As BusinessHelper = CreateBusinessObject("RutaProveedor")
            data.dtRutaProveedorNew = RP.AddNew
            Dim RPL As BusinessHelper = CreateBusinessObject("RutaProveedorLinea")
            data.dtRutaProveedorLineasNew = RPL.AddNew

            If data.HasRutaNew Is Nothing Then data.HasRutaNew = New Hashtable
            For Each drRutaOrigen As DataRow In data.dtRutaOrigen.Select
                Dim drRutaNew As DataRow = data.dtRutasNew.NewRow
                'drRutaNew.ItemArray = drRutaOrigen.ItemArray
                For Each col As DataColumn In data.dtRutasNew.Columns
                    If drRutaOrigen.Table.Columns.Contains(col.ColumnName) Then
                        drRutaNew(col.ColumnName) = drRutaOrigen(col.ColumnName)
                    End If
                Next
                drRutaNew("IDRutaOp") = AdminData.GetAutoNumeric
                drRutaNew("IDArticulo") = data.IDArticuloNew
                data.dtRutasNew.Rows.Add(drRutaNew)

                data.HasRutaNew.Add(drRutaOrigen("IDRutaOp"), drRutaNew("IDRutaOp"))

                Dim dtRutasProv As DataTable = RP.Filter(New NumberFilterItem("IDRutaOp", drRutaOrigen("IDRutaOp")))
                If dtRutasProv.Rows.Count > 0 Then
                    For Each drRutaProv As DataRow In dtRutasProv.Select
                        Dim drRutaProvNew As DataRow = data.dtRutaProveedorNew.NewRow
                        'drRutaProvNew.ItemArray = drRutaProv.ItemArray
                        For Each col As DataColumn In data.dtRutaProveedorNew.Columns
                            If drRutaProv.Table.Columns.Contains(col.ColumnName) Then
                                drRutaProvNew(col.ColumnName) = drRutaProv(col.ColumnName)
                            End If
                        Next
                        drRutaProvNew("IDRutaOp") = drRutaNew("IDRutaOp")
                        data.dtRutaProveedorNew.Rows.Add(drRutaProvNew)

                        Dim f As New Filter
                        f.Add(New NumberFilterItem("IDRutaOp", drRutaProv("IDRutaOp")))
                        f.Add(New StringFilterItem("IDProveedor", drRutaProv("IDProveedor")))
                        f.Add(New StringFilterItem("IDCentro", drRutaProv("IDCentro")))
                        Dim dtRutasProvLineas As DataTable = RPL.Filter(f)
                        If dtRutasProvLineas.Rows.Count > 0 Then
                            For Each drRutaProvLinea As DataRow In dtRutasProvLineas.Select
                                Dim drRutaProvLineaNew As DataRow = data.dtRutaProveedorLineasNew.NewRow
                                'drRutaProvLineaNew.ItemArray = drRutaProvLinea.ItemArray
                                For Each col As DataColumn In data.dtRutaProveedorLineasNew.Columns
                                    If drRutaProvLinea.Table.Columns.Contains(col.ColumnName) Then
                                        drRutaProvLineaNew(col.ColumnName) = drRutaProvLinea(col.ColumnName)
                                    End If
                                Next
                                drRutaProvLineaNew("IDRutaOp") = drRutaNew("IDRutaOp")

                                data.dtRutaProveedorLineasNew.Rows.Add(drRutaProvLineaNew)
                            Next
                        End If
                    Next
                End If
            Next
        End If
    End Sub

    <Task()> Public Shared Sub CopiarRutaParametro(ByVal data As DatosArtCopia, ByVal services As ServiceProvider)
        If data.dtRutaOrigen.Rows.Count > 0 Then
            Dim RP As New RutaParametro
            data.dtRutaParametroNew = RP.AddNew
            For Each drRutaOrigen As DataRow In data.dtRutaOrigen.Rows
                Dim dt As DataTable = RP.Filter(New NumberFilterItem("IDRutaOp", drRutaOrigen("IDRutaOp")))
                If dt.Rows.Count > 0 Then
                    For Each dr As DataRow In dt.Rows
                        Dim drNew As DataRow = data.dtRutaParametroNew.NewRow
                        ' drNew.ItemArray = dr.ItemArray
                        For Each col As DataColumn In data.dtRutaParametroNew.Columns
                            If dt.Columns.Contains(col.ColumnName) Then
                                drNew(col.ColumnName) = dr(col.ColumnName)
                            End If
                        Next
                        drNew("ID") = AdminData.GetAutoNumeric
                        drNew("IDRutaOp") = data.HasRutaNew(dr("IDRutaOp"))

                        data.dtRutaParametroNew.Rows.Add(drNew)
                    Next
                End If
            Next
        End If
    End Sub

    <Task()> Public Shared Sub CopiarRutaUtillaje(ByVal data As DatosArtCopia, ByVal services As ServiceProvider)
        If data.dtRutaOrigen.Rows.Count > 0 Then
            Dim RU As New RutaUtillaje
            data.dtRutaUtillajeNew = RU.AddNew
            For Each drRutaOrigen As DataRow In data.dtRutaOrigen.Rows
                Dim dt As DataTable = RU.Filter(New NumberFilterItem("IDRutaOp", drRutaOrigen("IDRutaOp")))
                If dt.Rows.Count > 0 Then
                    For Each dr As DataRow In dt.Rows
                        Dim drNew As DataRow = data.dtRutaUtillajeNew.NewRow
                        'drNew.ItemArray = dr.ItemArray
                        For Each col As DataColumn In data.dtRutaUtillajeNew.Columns
                            If dt.Columns.Contains(col.ColumnName) Then
                                drNew(col.ColumnName) = dr(col.ColumnName)
                            End If
                        Next
                        drNew("IDRutaOp") = data.HasRutaNew(dr("IDRutaOp"))

                        data.dtRutaUtillajeNew.Rows.Add(drNew)
                    Next
                End If
            Next
        End If
    End Sub

    <Task()> Public Shared Sub CopiarRutaOficio(ByVal data As DatosArtCopia, ByVal services As ServiceProvider)
        If data.dtRutaOrigen.Rows.Count > 0 Then
            Dim RO As BusinessHelper = BusinessHelper.CreateBusinessObject("RutaOficio")
            data.dtRutaOficioNew = RO.AddNew
            For Each drRutaOrigen As DataRow In data.dtRutaOrigen.Rows
                Dim dt As DataTable = RO.Filter(New NumberFilterItem("IDRutaOp", drRutaOrigen("IDRutaOp")))
                If dt.Rows.Count > 0 Then
                    For Each dr As DataRow In dt.Rows
                        Dim drNew As DataRow = data.dtRutaOficioNew.NewRow
                        ' drNew.ItemArray = dr.ItemArray
                        For Each col As DataColumn In data.dtRutaOficioNew.Columns
                            If dt.Columns.Contains(col.ColumnName) Then
                                drNew(col.ColumnName) = dr(col.ColumnName)
                            End If
                        Next
                        drNew("IDRutaOp") = data.HasRutaNew(dr("IDRutaOp"))

                        data.dtRutaOficioNew.Rows.Add(drNew)
                    Next
                End If
            Next
        End If
    End Sub

    <Task()> Public Shared Sub CopiarRutaAlternativa(ByVal data As DatosArtCopia, ByVal services As ServiceProvider)
        If data.dtRutaOrigen.Rows.Count > 0 Then
            Dim RA As BusinessHelper = BusinessHelper.CreateBusinessObject("RutaAlternativo")
            data.dtRutasAlternativasNew = RA.AddNew
            For Each drRutaOrigen As DataRow In data.dtRutaOrigen.Rows
                Dim dt As DataTable = RA.Filter(New NumberFilterItem("IDRutaOp", drRutaOrigen("IDRutaOp")))
                If dt.Rows.Count > 0 Then
                    For Each dr As DataRow In dt.Rows
                        Dim drNew As DataRow = data.dtRutasAlternativasNew.NewRow
                        'drNew.ItemArray = dr.ItemArray
                        For Each col As DataColumn In data.dtRutasAlternativasNew.Columns
                            If dt.Columns.Contains(col.ColumnName) Then
                                drNew(col.ColumnName) = dr(col.ColumnName)
                            End If
                        Next
                        drNew("ID") = AdminData.GetAutoNumeric
                        drNew("IDRutaOp") = data.HasRutaNew(dr("IDRutaOp"))

                        data.dtRutasAlternativasNew.Rows.Add(drNew)
                    Next
                End If
            Next
        End If
    End Sub

    <Task()> Public Shared Sub CopiarRutaAMFE(ByVal data As DatosArtCopia, ByVal services As ServiceProvider)
        Dim A As BusinessHelper = BusinessHelper.CreateBusinessObject("ArticuloRutaAMFE")
        Dim dt As DataTable = A.Filter(New StringFilterItem("IDArticulo", data.IDArticulo))
        If dt.Rows.Count > 0 Then
            data.dtRutaAMFENew = A.AddNew
            For Each dr As DataRow In dt.Rows
                Dim drNew As DataRow = data.dtRutaAMFENew.NewRow
                'drNew.ItemArray = dr.ItemArray
                For Each col As DataColumn In data.dtRutaAMFENew.Columns
                    If dt.Columns.Contains(col.ColumnName) Then
                        drNew(col.ColumnName) = dr(col.ColumnName)
                    End If
                Next
                drNew("IDRutaAMFE") = AdminData.GetAutoNumeric
                drNew("IDArticulo") = data.IDArticuloNew

                data.dtRutaAMFENew.Rows.Add(drNew)
            Next
        End If
    End Sub

#End Region

    <Task()> Public Shared Sub CopiarCaracteristicasMaq(ByVal data As DatosArtCopia, ByVal services As ServiceProvider)
        Dim A As New ArticuloNserieCaract
        Dim dt As DataTable = A.Filter(New StringFilterItem("IDArticulo", data.IDArticulo))
        If dt.Rows.Count > 0 Then
            data.dtNserieCaracteristicasNew = A.AddNew
            Dim drNew As DataRow = data.dtNserieCaracteristicasNew.NewRow
            'drNew.ItemArray = dt.Rows(0).ItemArray
            For Each col As DataColumn In data.dtNserieCaracteristicasNew.Columns
                If dt.Columns.Contains(col.ColumnName) Then
                    drNew(col.ColumnName) = dt.Rows(0)(col.ColumnName)
                End If
            Next
            drNew("IDArticulo") = data.IDArticuloNew
            data.dtNserieCaracteristicasNew.Rows.Add(drNew)
        End If
    End Sub

    <Task()> Public Shared Sub CopiarCaracteristicas(ByVal data As DatosArtCopia, ByVal services As ServiceProvider)
        Dim A As New ArticuloCaracteristica
        Dim dt As DataTable = A.Filter(New StringFilterItem("IDArticulo", data.IDArticulo))
        If dt.Rows.Count > 0 Then
            data.dtCaracteristicasNew = A.AddNew
            For Each dr As DataRow In dt.Select
                Dim drNew As DataRow = data.dtCaracteristicasNew.NewRow
                'drNew.ItemArray = dr.ItemArray
                For Each col As DataColumn In data.dtCaracteristicasNew.Columns
                    If dt.Columns.Contains(col.ColumnName) Then
                        drNew(col.ColumnName) = dr(col.ColumnName)
                    End If
                Next
                drNew("IDArticulo") = data.IDArticuloNew

                data.dtCaracteristicasNew.Rows.Add(drNew)
            Next
        End If
    End Sub

    <Task()> Public Shared Sub CopiarEspecificaciones(ByVal data As DatosArtCopia, ByVal services As ServiceProvider)
        Dim A As New Plantilla
        Dim dt As DataTable = A.Filter(New StringFilterItem("IDArticulo", data.IDArticulo))
        If dt.Rows.Count > 0 Then
            data.dtPlantillasNew = A.AddNew
            For Each dr As DataRow In dt.Select
                Dim drNew As DataRow = data.dtPlantillasNew.NewRow
                'drNew.ItemArray = dr.ItemArray
                For Each col As DataColumn In data.dtPlantillasNew.Columns
                    If dt.Columns.Contains(col.ColumnName) Then
                        drNew(col.ColumnName) = dr(col.ColumnName)
                    End If
                Next
                drNew("IDArticulo") = data.IDArticuloNew

                data.dtPlantillasNew.Rows.Add(drNew)
            Next
        End If
    End Sub

    <Task()> Public Shared Sub CopiarPromociones(ByVal data As DatosArtCopia, ByVal services As ServiceProvider)
        Dim A As New PromocionLinea
        Dim dt As DataTable = A.Filter(New StringFilterItem("IDArticulo", data.IDArticulo))
        If dt.Rows.Count > 0 Then
            data.dtPromocionesLineaNew = A.AddNew
            For Each dr As DataRow In dt.Select
                Dim drNew As DataRow = data.dtPromocionesLineaNew.NewRow
                'drNew.ItemArray = dr.ItemArray
                For Each col As DataColumn In data.dtPromocionesLineaNew.Columns
                    If dt.Columns.Contains(col.ColumnName) Then
                        drNew(col.ColumnName) = dr(col.ColumnName)
                    End If
                Next
                drNew("IDArticulo") = data.IDArticuloNew
                drNew("IDPromocionLinea") = AdminData.GetAutoNumeric

                data.dtPromocionesLineaNew.Rows.Add(drNew)
            Next
        End If
    End Sub

    <Task()> Public Shared Sub CopiarIdiomas(ByVal data As DatosArtCopia, ByVal services As ServiceProvider)
        Dim A As New ArticuloIdioma
        Dim dt As DataTable = A.Filter(New StringFilterItem("IDArticulo", data.IDArticulo))
        If dt.Rows.Count > 0 Then
            data.dtIdiomasNew = A.AddNew
            For Each dr As DataRow In dt.Select
                Dim drNew As DataRow = data.dtIdiomasNew.NewRow
                'drNew.ItemArray = dr.ItemArray
                For Each col As DataColumn In data.dtIdiomasNew.Columns
                    If dt.Columns.Contains(col.ColumnName) Then
                        drNew(col.ColumnName) = dr(col.ColumnName)
                    End If
                Next
                drNew("IDArticulo") = data.IDArticuloNew

                data.dtIdiomasNew.Rows.Add(drNew)
            Next
        End If
    End Sub

    <Task()> Public Shared Sub CopiarDocumentos(ByVal data As DatosArtCopia, ByVal services As ServiceProvider)
        Dim A As New ArticuloPlano
        Dim dt As DataTable = A.Filter(New StringFilterItem("IDArticulo", data.IDArticulo))
        If dt.Rows.Count > 0 Then
            data.dtPlanosNew = A.AddNew
            For Each dr As DataRow In dt.Select
                Dim drNew As DataRow = data.dtPlanosNew.NewRow
                'drNew.ItemArray = dr.ItemArray
                For Each col As DataColumn In data.dtPlanosNew.Columns
                    If dt.Columns.Contains(col.ColumnName) Then
                        drNew(col.ColumnName) = dr(col.ColumnName)
                    End If
                Next
                drNew("IDArticulo") = data.IDArticuloNew
                drNew("IDArticuloPlano") = AdminData.GetAutoNumeric

                data.dtPlanosNew.Rows.Add(drNew)
            Next
        End If
    End Sub

    <Task()> Public Shared Sub CopiarAnalitica(ByVal data As DatosArtCopia, ByVal services As ServiceProvider)
        Dim A As New ArticuloAnalitica
        Dim dt As DataTable = A.Filter(New StringFilterItem("IDArticulo", data.IDArticulo))
        If dt.Rows.Count > 0 Then
            data.dtAnaliticaNew = A.AddNew
            For Each dr As DataRow In dt.Select
                Dim drNew As DataRow = data.dtAnaliticaNew.NewRow
                'drNew.ItemArray = dr.ItemArray
                For Each col As DataColumn In data.dtAnaliticaNew.Columns
                    If dt.Columns.Contains(col.ColumnName) Then
                        drNew(col.ColumnName) = dr(col.ColumnName)
                    End If
                Next
                drNew("IDArticulo") = data.IDArticuloNew

                data.dtAnaliticaNew.Rows.Add(drNew)
            Next
        End If
    End Sub

    <Task()> Public Shared Sub CopiarCostesVarios(ByVal data As DatosArtCopia, ByVal services As ServiceProvider)
        Dim A As New ArticuloVarios
        Dim dt As DataTable = A.Filter(New StringFilterItem("IDArticulo", data.IDArticulo))
        If dt.Rows.Count > 0 Then
            data.dtCostesVariosNew = A.AddNew
            For Each dr As DataRow In dt.Select
                Dim drNew As DataRow = data.dtCostesVariosNew.NewRow
                'drNew.ItemArray = dr.ItemArray
                For Each col As DataColumn In data.dtCostesVariosNew.Columns
                    If dt.Columns.Contains(col.ColumnName) Then
                        drNew(col.ColumnName) = dr(col.ColumnName)
                    End If
                Next
                drNew("IDArticulo") = data.IDArticuloNew

                data.dtCostesVariosNew.Rows.Add(drNew)
            Next
        End If
    End Sub

#Region " CopiarProveedores "

    <Task()> Public Shared Sub CopiarProveedores(ByVal data As DatosArtCopia, ByVal services As ServiceProvider)
        Dim A As New ArticuloProveedor
        Dim dt As DataTable = A.Filter(New StringFilterItem("IDArticulo", data.IDArticulo))
        If dt.Rows.Count > 0 Then
            data.dtProveedoresNew = A.AddNew
            For Each dr As DataRow In dt.Select
                Dim drNew As DataRow = data.dtProveedoresNew.NewRow
                'drNew.ItemArray = dr.ItemArray
                For Each col As DataColumn In data.dtProveedoresNew.Columns
                    If dt.Columns.Contains(col.ColumnName) Then
                        drNew(col.ColumnName) = dr(col.ColumnName)
                    End If
                Next
                drNew("IDArticulo") = data.IDArticuloNew

                data.dtProveedoresNew.Rows.Add(drNew)
            Next
        End If
    End Sub

    <Task()> Public Shared Sub CopiarProveedoresLineas(ByVal data As DatosArtCopia, ByVal services As ServiceProvider)
        Dim A As New ArticuloProveedorLinea
        Dim dt As DataTable = A.Filter(New StringFilterItem("IDArticulo", data.IDArticulo))
        If dt.Rows.Count > 0 Then
            data.dtProveedoresLineasNew = A.AddNew
            For Each dr As DataRow In dt.Select
                Dim drNew As DataRow = data.dtProveedoresLineasNew.NewRow
                'drNew.ItemArray = dr.ItemArray
                For Each col As DataColumn In data.dtProveedoresLineasNew.Columns
                    If dt.Columns.Contains(col.ColumnName) Then
                        drNew(col.ColumnName) = dr(col.ColumnName)
                    End If
                Next
                drNew("IDArticulo") = data.IDArticuloNew

                data.dtProveedoresLineasNew.Rows.Add(drNew)
            Next
        End If
    End Sub

#End Region

#Region " CopiarTarifas "

    <Task()> Public Shared Sub CopiarTarifasArt(ByVal data As DatosArtCopia, ByVal services As ServiceProvider)
        Dim A As New TarifaArticulo
        Dim dt As DataTable = A.Filter(New StringFilterItem("IDArticulo", data.IDArticulo))
        If dt.Rows.Count > 0 Then
            data.dtTarifasNew = A.AddNew
            For Each dr As DataRow In dt.Select
                Dim drNew As DataRow = data.dtTarifasNew.NewRow
                'drNew.ItemArray = dr.ItemArray
                For Each col As DataColumn In data.dtTarifasNew.Columns
                    If dt.Columns.Contains(col.ColumnName) Then
                        drNew(col.ColumnName) = dr(col.ColumnName)
                    End If
                Next
                drNew("IDArticulo") = data.IDArticuloNew

                data.dtTarifasNew.Rows.Add(drNew)
            Next
        End If
    End Sub

    <Task()> Public Shared Sub CopiarTarifasArtLineas(ByVal data As DatosArtCopia, ByVal services As ServiceProvider)
        Dim A As New TarifaArticuloLinea
        Dim dt As DataTable = A.Filter(New StringFilterItem("IDArticulo", data.IDArticulo))
        If dt.Rows.Count > 0 Then
            data.dtTarifasLineasNew = A.AddNew
            For Each dr As DataRow In dt.Select
                Dim drNew As DataRow = data.dtTarifasLineasNew.NewRow
                'drNew.ItemArray = dr.ItemArray
                For Each col As DataColumn In data.dtTarifasLineasNew.Columns
                    If dt.Columns.Contains(col.ColumnName) Then
                        drNew(col.ColumnName) = dr(col.ColumnName)
                    End If
                Next
                drNew("IDArticulo") = data.IDArticuloNew

                data.dtTarifasLineasNew.Rows.Add(drNew)
            Next
        End If
    End Sub

#End Region

#Region " CopiarClientes "

    <Task()> Public Shared Sub CopiarClientes(ByVal data As DatosArtCopia, ByVal services As ServiceProvider)
        Dim A As New ArticuloCliente
        Dim dt As DataTable = A.Filter(New StringFilterItem("IDArticulo", data.IDArticulo))
        If dt.Rows.Count > 0 Then
            data.dtClientesNew = A.AddNew
            For Each dr As DataRow In dt.Select
                Dim drNew As DataRow = data.dtClientesNew.NewRow
                'drNew.ItemArray = dr.ItemArray
                For Each col As DataColumn In data.dtClientesNew.Columns
                    If dt.Columns.Contains(col.ColumnName) Then
                        drNew(col.ColumnName) = dr(col.ColumnName)
                    End If
                Next
                drNew("IDArticulo") = data.IDArticuloNew

                data.dtClientesNew.Rows.Add(drNew)
            Next
        End If
    End Sub

    <Task()> Public Shared Sub CopiarClientesLineas(ByVal data As DatosArtCopia, ByVal services As ServiceProvider)
        Dim A As New ArticuloClienteLinea
        Dim dt As DataTable = A.Filter(New StringFilterItem("IDArticulo", data.IDArticulo))
        If dt.Rows.Count > 0 Then
            data.dtClientesLineasNew = A.AddNew
            For Each dr As DataRow In dt.Select
                Dim drNew As DataRow = data.dtClientesLineasNew.NewRow
                'drNew.ItemArray = dr.ItemArray
                For Each col As DataColumn In data.dtClientesLineasNew.Columns
                    If dt.Columns.Contains(col.ColumnName) Then
                        drNew(col.ColumnName) = dr(col.ColumnName)
                    End If
                Next
                drNew("IDArticulo") = data.IDArticuloNew

                data.dtClientesLineasNew.Rows.Add(drNew)
            Next
        End If
    End Sub

#End Region

    <Task()> Public Shared Sub CopiarUnidades(ByVal data As DatosArtCopia, ByVal services As ServiceProvider)
        Dim A As New ArticuloUnidadAB
        Dim dt As DataTable = A.Filter(New StringFilterItem("IDArticulo", data.IDArticulo))
        If dt.Rows.Count > 0 Then
            data.dtUnidadesABNew = A.AddNew
            For Each dr As DataRow In dt.Select
                Dim drNew As DataRow = data.dtUnidadesABNew.NewRow
                'drNew.ItemArray = dr.ItemArray
                For Each col As DataColumn In data.dtUnidadesABNew.Columns
                    If dt.Columns.Contains(col.ColumnName) Then
                        drNew(col.ColumnName) = dr(col.ColumnName)
                    End If
                Next
                drNew("IDArticulo") = data.IDArticuloNew

                data.dtUnidadesABNew.Rows.Add(drNew)
            Next
        End If
    End Sub

    <Task()> Public Shared Sub CopiarAlmacenes(ByVal data As DatosArtCopia, ByVal services As ServiceProvider)
        Dim A As New ArticuloAlmacen
        Dim dt As DataTable = A.Filter(New StringFilterItem("IDArticulo", data.IDArticulo))
        If dt.Rows.Count > 0 Then
            data.dtAlmacenesNew = A.AddNew
            For Each dr As DataRow In dt.Select
                Dim drNew As DataRow = data.dtAlmacenesNew.NewRow
                'drNew.ItemArray = dr.ItemArray
                For Each col As DataColumn In data.dtAlmacenesNew.Columns
                    If dr.Table.Columns.Contains(col.ColumnName) Then
                        drNew(col.ColumnName) = dr(col.ColumnName)
                    End If
                Next
                drNew("IDArticulo") = data.IDArticuloNew
                drNew("StockFisico") = 0
                drNew("PrecioMedioA") = 0
                drNew("PrecioMedioB") = 0
                drNew("PrecioFIFOFechaA") = 0
                drNew("PrecioFIFOFechaB") = 0
                drNew("PrecioFIFOMvtoA") = 0
                drNew("PrecioFIFOMvtoB") = 0
                drNew("StockFisico2") = 0
                drNew("Inventariado") = False
                drNew("FechaUltimoAjuste") = DBNull.Value
                drNew("FechaUltimoInventario") = DBNull.Value
                drNew("FechaUltimoMovimiento") = DBNull.Value
                drNew("IDArticuloGenerico") = DBNull.Value
                drNew("MarcaAuto") = AdminData.GetAutoNumeric
                drNew("StockFechaCalculo") = 0
                drNew("FechaCalculo") = DBNull.Value

                data.dtAlmacenesNew.Rows.Add(drNew)
            Next
        End If
    End Sub

    <Task()> Public Shared Sub GuardarCopiaArticulo(ByVal data As DatosArtCopia, ByVal services As ServiceProvider)
        AdminData.BeginTx()
        Dim u As New UpdatePackage
        u.Add(data.dtArticuloNew)
        If Not data.dtCosteEstandarNew Is Nothing Then u.Add(data.dtCosteEstandarNew)
        If Not data.dtEstructurasNew Is Nothing Then u.Add(data.dtEstructurasNew)
        If Not data.dtEstructuraNew Is Nothing Then u.Add(data.dtEstructuraNew)
        If Not data.dtArticuloRutaNew Is Nothing Then u.Add(data.dtArticuloRutaNew)
        If Not data.dtRutasNew Is Nothing Then u.Add(data.dtRutasNew)
        If Not data.dtRutaProveedorNew Is Nothing Then u.Add(data.dtRutaProveedorNew)
        If Not data.dtRutaProveedorLineasNew Is Nothing Then u.Add(data.dtRutaProveedorLineasNew)
        If Not data.dtRutaParametroNew Is Nothing Then u.Add(data.dtRutaParametroNew)
        If Not data.dtRutaUtillajeNew Is Nothing Then u.Add(data.dtRutaUtillajeNew)
        If Not data.dtRutaOficioNew Is Nothing Then u.Add(data.dtRutaOficioNew)
        If Not data.dtRutasAlternativasNew Is Nothing Then u.Add(data.dtRutasAlternativasNew)
        If Not data.dtRutaAMFENew Is Nothing Then u.Add(data.dtRutaAMFENew)
        If Not data.dtNserieCaracteristicasNew Is Nothing Then u.Add(data.dtNserieCaracteristicasNew)
        If Not data.dtCaracteristicasNew Is Nothing Then u.Add(data.dtCaracteristicasNew)
        If Not data.dtPlantillasNew Is Nothing Then u.Add(data.dtPlantillasNew)
        If Not data.dtPromocionesLineaNew Is Nothing Then u.Add(data.dtPromocionesLineaNew)
        If Not data.dtIdiomasNew Is Nothing Then u.Add(data.dtIdiomasNew)
        If Not data.dtPlanosNew Is Nothing Then u.Add(data.dtPlanosNew)
        If Not data.dtAnaliticaNew Is Nothing Then u.Add(data.dtAnaliticaNew)
        If Not data.dtCostesVariosNew Is Nothing Then u.Add(data.dtCostesVariosNew)
        If Not data.dtProveedoresNew Is Nothing Then u.Add(data.dtProveedoresNew)
        If Not data.dtProveedoresLineasNew Is Nothing Then u.Add(data.dtProveedoresLineasNew)
        If Not data.dtTarifasNew Is Nothing Then u.Add(data.dtTarifasNew)
        If Not data.dtTarifasLineasNew Is Nothing Then u.Add(data.dtTarifasLineasNew)
        If Not data.dtClientesNew Is Nothing Then u.Add(data.dtClientesNew)
        If Not data.dtClientesLineasNew Is Nothing Then u.Add(data.dtClientesLineasNew)
        If Not data.dtUnidadesABNew Is Nothing Then u.Add(data.dtUnidadesABNew)
        If Not data.dtAlmacenesNew Is Nothing Then u.Add(data.dtAlmacenesNew)

        BusinessHelper.UpdatePackage(u)
    End Sub

#End Region

#Region "Funciones Públicas"

    <Serializable()> _
    Public Class DatosArtCalculo
        Public IDArticulo As String
        Public Tipo As enumacsTipoArticulo
        Public Operacion As enumPlazo
        Public plazo As Double
        Public IDProcess As String
        Public Porcentaje As Double
        Public Valor As Double
        Public IDAlmacen As String
        Public QFabricar As Double
        Public LoteMRP As Boolean
        Public Lote As Double
        Public Capacidad As Double
    End Class
    <Serializable()> _
     Public Class UnidadMedidaPrecioInfo
        Public IDArticulo As String
        Public IDUdMedida As String
        Public IDUdInterna As String
        Public Cantidad As Double
        Public PrecioA As Double
        Public PrecioB As Double
        Public UdValoracion As Short
    End Class
    <Task()> Public Shared Function CalcularPlazo(ByVal data As DatosArtCalculo, ByVal services As ServiceProvider) As Double
        Dim dtPlazo As DataTable
        Dim _filter As New Filter
        _filter.Add("IDArticulo", FilterOperator.Equal, data.IDArticulo)
        Select Case data.Tipo
            Case enumacsTipoArticulo.acsCompra
                dtPlazo = New BE.DataEngine().Filter("vCTLCICalculoPlazoCompra", _filter)
                If Not dtPlazo Is Nothing AndAlso dtPlazo.Rows.Count > 0 Then
                    Return xRound(Nz(dtPlazo.Rows(0)("Plazo"), 0), 2)
                End If
            Case enumacsTipoArticulo.acsFabrica
                dtPlazo = New BE.DataEngine().Filter("vCTLCICalculoPlazoFabrica", _filter)
                If Not dtPlazo Is Nothing AndAlso dtPlazo.Rows.Count > 0 Then
                    Return xRound(Nz(dtPlazo.Rows(0)("plazofabrica"), 0), 2)
                End If
        End Select
    End Function
    <Task()> Public Shared Function CalcularPlazoFabrica(ByVal data As DatosArtCalculo, ByVal services As ServiceProvider) As Double
        If Length(data.IDArticulo) > 0 Then
            Dim dtArt As DataTable = New Articulo().Filter(New FilterItem("IdArticulo", data.IDArticulo))
            If Not dtArt Is Nothing AndAlso dtArt.Rows.Count > 0 Then
                Dim ClsArtAlm As New ArticuloAlmacen
                Dim dtArtAlm As DataTable = ClsArtAlm.SelOnPrimaryKey(data.IDArticulo, data.IDAlmacen)
                If Not dtArtAlm Is Nothing AndAlso dtArtAlm.Rows.Count > 0 Then
                    data.Lote = Nz(dtArtAlm.Rows(0)("LoteMinimo"), 0)
                    data.plazo = dtArt.Rows(0)("PlazoFabricacion")
                    data.LoteMRP = dtArt.Rows(0)("AplicarLoteMRP")
                    data.Capacidad = dtArt.Rows(0)("CapacidadDiaria")
                End If
            End If
        End If
        If data.Capacidad > 0 And data.QFabricar > 0 Then
            CalcularPlazoFabrica = data.plazo + System.Math.Round((1 / data.Capacidad) * data.QFabricar)
        Else
            If data.Lote > 0 And data.LoteMRP And data.QFabricar > 0 And data.QFabricar > data.Lote Then
                CalcularPlazoFabrica = data.plazo * data.QFabricar / data.Lote
            Else
                CalcularPlazoFabrica = data.plazo
            End If
        End If
        CalcularPlazoFabrica = Math.Ceiling(CalcularPlazoFabrica)
        Return CalcularPlazoFabrica

    End Function
    <Task()> Public Shared Sub ActualizacionDePlazo(ByVal data As DatosArtCalculo, ByVal services As ServiceProvider)
        Dim _filter As New Filter
        Dim dt, dtArticulos As DataTable
        Dim _articulo As New Articulo
        Dim i As Integer = 0
        _filter.Add("idprocess", FilterOperator.Equal, data.IDProcess)
        dt = AdminData.Filter("vFrmCIActualizarPlazo", , _filter.Compose(New AdoFilterComposer), "IDArticulo")
        Dim ListaArticulos(dt.Rows.Count - 1) As Object
        For Each dr As DataRow In dt.Select
            ListaArticulos(i) = dr("IDArticulo")
            i += 1
        Next
        _filter.Clear()
        _filter.Add(New InListFilterItem("IDArticulo", ListaArticulos, FilterType.String))
        dtArticulos = _articulo.Filter(_filter, "IDArticulo")
        i = 0
        For Each drArticulo As DataRow In dtArticulos.Select()
            Select Case data.Operacion
                Case enumPlazo.enumPMasivo
                    Select Case data.Tipo
                        Case enumacsTipoArticulo.acsCompra
                            drArticulo("Plazo") = data.plazo
                        Case enumacsTipoArticulo.acsFabrica
                            drArticulo("PlazoFabricacion") = data.plazo
                    End Select
                Case enumPlazo.enumPIndividual
                    If Length(dt.Rows(i).Item("CantidadMArca1")) > 0 AndAlso dt.Rows(i).Item("CantidadMArca1") <> 0 Then
                        Select Case data.Tipo
                            Case enumacsTipoArticulo.acsCompra
                                drArticulo("Plazo") = dt.Rows(i).Item("CantidadMArca1")
                            Case enumacsTipoArticulo.acsFabrica
                                drArticulo("PlazoFabricacion") = dt.Rows(i).Item("CantidadMarca1")
                        End Select
                    End If
                    i += 1
                Case enumPlazo.enumPCompra
                    data.IDArticulo = drArticulo("IDArticulo")
                    drArticulo("Plazo") = ProcessServer.ExecuteTask(Of DatosArtCalculo, Double)(AddressOf CalcularPlazo, data, services)
                Case enumPlazo.enumPFabrica1
                    data.IDArticulo = drArticulo("IDArticulo")
                    drArticulo("PlazoFabricacion") = ProcessServer.ExecuteTask(Of DatosArtCalculo, Double)(AddressOf CalcularPlazo, data, services)
            End Select
        Next
        BusinessHelper.UpdateTable(dtArticulos)
    End Sub
    <Task()> Public Shared Sub ActualizacionDePrecios(ByVal data As DatosArtCalculo, ByVal services As ServiceProvider)
        Dim _filter As New Filter
        Dim dt, dtArticulos, dtMoneda As DataTable
        Dim _articulo As New Articulo
        Dim i As Integer = 0
        Dim CambioB As Double
        Dim NDecimalesPrec As Integer
        Dim _moneda As New Moneda
        Dim _decimales As Integer


        _decimales = ProcessServer.ExecuteTask(Of Date, MonedaInfo)(AddressOf Moneda.MonedaA, cnMinDate, services).NDecimalesPrecio
        _filter.Add("idprocess", FilterOperator.Equal, data.IDProcess)
        dt = AdminData.Filter("vFrmCIActualizarPrecioEstandar", , _filter.Compose(New AdoFilterComposer), "IDArticulo")
        Dim ListaArticulos(dt.Rows.Count - 1) As Object
        For Each dr As DataRow In dt.Select
            ListaArticulos(i) = dr("IDArticulo")
            i += 1
        Next
        _filter.Clear()
        _filter.Add(New InListFilterItem("IDArticulo", ListaArticulos, FilterType.String))
        dtArticulos = _articulo.Filter(_filter, "IDArticulo")
        i = 0
        For Each drArticulo As DataRow In dtArticulos.Select()
            Select Case data.Operacion
                Case enumPrecioEstandar.enumPSIndividual
                    If Length(dt.Rows(i).Item("CantidadMArca1")) > 0 AndAlso dt.Rows(i).Item("CantidadMArca1") <> 0 Then
                        drArticulo("PrecioEstandarA") = dt.Rows(i).Item("CantidadMarca1")
                    End If
                    i += 1
                Case enumPrecioEstandar.enumPSPorcEstandar
                    drArticulo("PrecioEstandarA") = drArticulo("PrecioEstandarA") * (1 + (data.Porcentaje / 100))
                Case enumPrecioEstandar.enumPSPorcUltimo
                    drArticulo("PrecioEstandarA") = drArticulo("PrecioUltimaCompraA") * (1 + (data.Porcentaje / 100))
                Case enumPrecioEstandar.enumPSUltimo
                    drArticulo("PrecioEstandarA") = drArticulo("PrecioUltimaCompraA")
            End Select
            'drArticulo("PrecioEstandarB") = xRound(drArticulo("PrecioEstandarA") * CambioB, _decimales)
            drArticulo("PrecioEstandarB") = ProcessServer.ExecuteTask(Of Double, Double)(AddressOf CalcularImporteEnMonedaB, drArticulo("PrecioEstandarA"), services)
            drArticulo("FechaEstandar") = Today
        Next
        BusinessHelper.UpdateTable(dtArticulos)
    End Sub
    <Task()> Public Shared Sub ActualizacionPreciosPorProveedor(ByVal data As DatosArtCalculo, ByVal services As ServiceProvider)
        Dim _filter As New Filter
        Dim dt, dtMoneda, dtArticulosProveedorLinea, dtArticulosProveedor As DataTable
        Dim i As Integer = 0
        Dim _artProvLinea As New ArticuloProveedorLinea
        Dim _artProv As New ArticuloProveedor
        Dim _filterAND As Filter
        Dim _filterOR As New Filter(FilterUnionOperator.Or)
        Dim NDecimalesPrec As Integer
        Dim _decimales As Integer
        Dim monedaA As String
        Dim proveedor As ProveedorInfo

        _decimales = ProcessServer.ExecuteTask(Of Date, MonedaInfo)(AddressOf Moneda.MonedaA, cnMinDate, services).NDecimalesPrecio
        monedaA = ProcessServer.ExecuteTask(Of Date, MonedaInfo)(AddressOf Moneda.MonedaA, cnMinDate, services).ID
        _filter.Add("idprocess", FilterOperator.Equal, data.IDProcess)
        dt = AdminData.Filter("vFrmCIActualizarPrecioProveedor", , AdminData.ComposeFilter(_filter), "idarticulo, idproveedor,QDesde")
        For Each dr As DataRow In dt.Select
            _filterAND = New Filter
            _filterAND.Add("IDArticulo", dr("IDArticulo"))
            _filterAND.Add("IDProveedor", dr("IdProveedor"))
            _filterAND.Add("QDesde", dr("QDesde"))
            _filterOR.Add(_filterAND)
        Next
        dtArticulosProveedorLinea = _artProvLinea.Filter(_filterOR, "idarticulo, idproveedor,QDesde")
        i = 0
        _filterOR.Clear()
        For Each drArticuloProveedorLinea As DataRow In dtArticulosProveedorLinea.Select()
            _filterAND = New Filter
            _filterAND.Add("IDArticulo", drArticuloProveedorLinea("IDArticulo"))
            _filterAND.Add("IDProveedor", drArticuloProveedorLinea("IDProveedor"))
            _filterOR.Add(_filterAND)
            Select Case data.Operacion
                Case enumPrecioProv.enumPPDto1
                    drArticuloProveedorLinea("Dto1") = data.Valor
                Case enumPrecioProv.enumPPDto2
                    drArticuloProveedorLinea("Dto2") = data.Valor
                Case enumPrecioProv.enumPPDto3
                    drArticuloProveedorLinea("Dto3") = data.Valor
                Case enumPrecioProv.enumPPPorcPrecio
                    drArticuloProveedorLinea("Precio") = drArticuloProveedorLinea("Precio") * (1 + (data.Valor / 100))
                Case enumPrecioProv.enumPPPorcEstandar
                    Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
                    Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(drArticuloProveedorLinea("IDProveedor"))
                    If (ProvInfo.IDMoneda <> monedaA) Then
                        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
                        Dim m As MonedaInfo = Monedas.GetMoneda(ProvInfo.IDMoneda, Today())
                        drArticuloProveedorLinea("Precio") = xRound((dt.Rows(i)("PrecioEstandarA") * (1 + (data.Valor / 100))) / m.CambioA, m.NDecimalesPrecio)
                    Else
                        drArticuloProveedorLinea("Precio") = dt.Rows(i)("PrecioEstandarA") * (1 + (data.Valor / 100))
                    End If
                Case enumPrecioProv.enumPPPorcUltimo
                    Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
                    Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(drArticuloProveedorLinea("IDProveedor"))
                    If (ProvInfo.IDMoneda <> monedaA) Then
                        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
                        Dim m As MonedaInfo = Monedas.GetMoneda(ProvInfo.IDMoneda, Today())
                        drArticuloProveedorLinea("Precio") = xRound((dt.Rows(i)("PrecioUltimaCompraA") * (1 + (data.Valor / 100))) / m.CambioA, m.NDecimalesPrecio)
                    Else
                        drArticuloProveedorLinea("Precio") = dt.Rows(i)("PrecioUltimaCompraA") * (1 + (data.Valor / 100))
                    End If
                Case enumPrecioProv.enumPPEstandar
                    Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
                    Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(drArticuloProveedorLinea("IDProveedor"))

                    Dim DrArtProv As DataRow = New ArticuloProveedor().GetItemRow(drArticuloProveedorLinea("IDProveedor"), drArticuloProveedorLinea("IDArticulo"))
                    Dim DrArt As DataRow = New Articulo().GetItemRow(drArticuloProveedorLinea("IDArticulo"))
                    Dim StDataAB As New ArticuloUnidadAB.DatosFactorConversion(drArticuloProveedorLinea("IDArticulo"), DrArtProv("IDUDCompra"), DrArt("IDUDInterna"))
                    Dim DblFactor As Double = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, StDataAB, services)
                    If (ProvInfo.IDMoneda <> monedaA) Then
                        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
                        Dim m As MonedaInfo = Monedas.GetMoneda(ProvInfo.IDMoneda, Today())
                        drArticuloProveedorLinea("Precio") = xRound((dt.Rows(i)("PrecioEstandarA") * DblFactor) / m.CambioA, m.NDecimalesPrecio)
                    Else
                        drArticuloProveedorLinea("Precio") = (dt.Rows(i)("PrecioEstandarA") * DblFactor)
                    End If
                Case enumPrecioProv.enumPPUltimo
                    Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
                    Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(drArticuloProveedorLinea("IDProveedor"))
                    If (ProvInfo.IDMoneda <> monedaA) Then
                        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
                        Dim m As MonedaInfo = Monedas.GetMoneda(ProvInfo.IDMoneda, Today())
                        drArticuloProveedorLinea("Precio") = xRound(dt.Rows(i)("PrecioUltimaCompraA") / m.CambioA, m.NDecimalesPrecio)
                    Else
                        drArticuloProveedorLinea("Precio") = dt.Rows(i)("PrecioUltimaCompraA")
                    End If
                Case enumPrecioProv.enumPPIndividual
                    If Length(dt.Rows(i).Item("CantidadMarca1")) > 0 AndAlso dt.Rows(i)("CantidadMarca1") <> 0 Then
                        drArticuloProveedorLinea("Precio") = dt.Rows(i)("CantidadMarca1")
                    End If
            End Select
            drArticuloProveedorLinea("Precio") = xRound(drArticuloProveedorLinea("Precio"), _decimales)
            i += 1
        Next
        BusinessHelper.UpdateTable(dtArticulosProveedorLinea)


        Dim StrWhere As String = AdminData.ComposeFilter(New GuidFilterItem("IDProcess", data.IDProcess))

        dt = New BE.DataEngine().Filter("vFrmCIActualizarPrecioProveedor", "*", StrWhere)
        _filterOR.Clear()
        For Each dr As DataRow In dt.Select
            _filterAND = New Filter
            _filterAND.Add("IDArticulo", dr("IDArticulo"))
            _filterAND.Add("IDProveedor", dr("IdProveedor"))
            _filterOR.Add(_filterAND)
            _filterAND = Nothing
        Next
        dtArticulosProveedor = _artProv.Filter(_filterOR, "idarticulo, idproveedor")
        i = 0
        For Each drArticuloProveedor As DataRow In dtArticulosProveedor.Select()
            Select Case data.Operacion
                Case enumPrecioProv.enumPPDto1
                    drArticuloProveedor("Dto1") = data.Valor
                Case enumPrecioProv.enumPPDto2
                    drArticuloProveedor("Dto2") = data.Valor
                Case enumPrecioProv.enumPPDto3
                    drArticuloProveedor("Dto3") = data.Valor
                Case enumPrecioProv.enumPPPorcPrecio
                    drArticuloProveedor("Precio") = drArticuloProveedor("Precio") * (1 + (data.Valor / 100))
                Case enumPrecioProv.enumPPPorcEstandar
                    Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
                    Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(drArticuloProveedor("IDProveedor"))
                    If (ProvInfo.IDMoneda <> monedaA) Then
                        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
                        Dim m As MonedaInfo = Monedas.GetMoneda(ProvInfo.IDMoneda, Today())
                        drArticuloProveedor("Precio") = xRound((dt.Rows(i)("PrecioEstandarA") * (1 + (data.Valor / 100))) / m.CambioA, m.NDecimalesPrecio)
                    Else
                        drArticuloProveedor("Precio") = dt.Rows(i)("PrecioEstandarA") * (1 + (data.Valor / 100))
                    End If
                Case enumPrecioProv.enumPPPorcUltimo
                    Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
                    Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(drArticuloProveedor("IDProveedor"))
                    If (ProvInfo.IDMoneda <> monedaA) Then
                        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
                        Dim m As MonedaInfo = Monedas.GetMoneda(ProvInfo.IDMoneda, Today())
                        drArticuloProveedor("Precio") = xRound((dt.Rows(i)("PrecioUltimaCompraA") * (1 + (data.Valor / 100))) / m.CambioA, m.NDecimalesPrecio)
                    Else
                        drArticuloProveedor("Precio") = dt.Rows(i)("PrecioUltimaCompraA") * (1 + (data.Valor / 100))
                    End If
                Case enumPrecioProv.enumPPEstandar
                    Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
                    Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(drArticuloProveedor("IDProveedor"))
                    Dim DrArt As DataRow = New Articulo().GetItemRow(drArticuloProveedor("IDArticulo"))
                    Dim StDataAB As New ArticuloUnidadAB.DatosFactorConversion(drArticuloProveedor("IDArticulo"), drArticuloProveedor("IDUDCompra"), DrArt("IDUDInterna"))
                    Dim DblFactor As Double = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, StDataAB, services)
                    If (ProvInfo.IDMoneda <> monedaA) Then
                        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
                        Dim m As MonedaInfo = Monedas.GetMoneda(ProvInfo.IDMoneda, Today())
                        drArticuloProveedor("Precio") = xRound((dt.Rows(i)("PrecioEstandarA") * DblFactor) / m.CambioA, m.NDecimalesPrecio)
                    Else
                        drArticuloProveedor("Precio") = (dt.Rows(i)("PrecioEstandarA") * DblFactor)
                    End If
                Case enumPrecioProv.enumPPUltimo
                    Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
                    Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(drArticuloProveedor("IDProveedor"))
                    If (ProvInfo.IDMoneda <> monedaA) Then
                        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
                        Dim m As MonedaInfo = Monedas.GetMoneda(ProvInfo.IDMoneda, Today())
                        drArticuloProveedor("Precio") = xRound(dt.Rows(i)("PrecioUltimaCompraA") / m.CambioA, m.NDecimalesPrecio)
                    Else
                        drArticuloProveedor("Precio") = dt.Rows(i)("PrecioUltimaCompraA")
                    End If
                Case enumPrecioProv.enumPPIndividual
                    If Length(dt.Rows(i).Item("CantidadMarca1")) > 0 AndAlso dt.Rows(i)("CantidadMarca1") <> 0 Then
                        drArticuloProveedor("Precio") = dt.Rows(i)("CantidadMarca1")
                    End If
            End Select
            drArticuloProveedor("Precio") = xRound(drArticuloProveedor("Precio"), _decimales)
            i += 1
        Next
        BusinessHelper.UpdateTable(dtArticulosProveedor)
    End Sub
    <Task()> Public Shared Function CaracteristicasArticulo(ByVal strIDArticulo As String, ByVal services As ServiceProvider) As DataTable
        If Length(strIDArticulo) > 0 Then
            Dim objFilter As New Filter
            objFilter.Add("IDArticulo", FilterOperator.Equal, strIDArticulo, FilterType.String)
            CaracteristicasArticulo = New BE.DataEngine().Filter("vNegCaractArticulo", objFilter)
        End If
    End Function
    <Task()> Public Shared Function CaracteristicasArticuloInfo(ByVal strIDArticulo As String, ByVal services As ServiceProvider) As ArticuloInfo
        Dim ArtInfo As ArticuloInfo
        If Length(strIDArticulo) > 0 Then
            Dim stArticulo As New DataInfoArticulo(strIDArticulo)
            ArtInfo = ProcessServer.ExecuteTask(Of DataInfoArticulo, ArticuloInfo)(AddressOf InformacionArticulo, stArticulo, services)
        End If
        Return ArtInfo
    End Function
    <Task()> Public Shared Function Estructura(ByVal strIDArticulo As String, ByVal services As ServiceProvider) As DataTable
        If Length(strIDArticulo) > 0 Then
            Dim dtArticulo As DataTable = New Articulo().SelOnPrimaryKey(strIDArticulo)
            If Not IsNothing(dtArticulo) AndAlso dtArticulo.Rows.Count > 0 Then
                Dim strWhere As String
                If dtArticulo.Rows(0)("TipoEstructura") Then
                    If Length(dtArticulo.Rows(0)("IDTipoEstructura")) > 0 Then
                        strWhere = "IDTipoEstructura = '" & dtArticulo.Rows(0)("IDTipoEstructura") & "'"
                    End If
                Else
                    Dim strArtEstruct As String = ProcessServer.ExecuteTask(Of String, String)(AddressOf New ArticuloEstructura().EstructuraPpal, strIDArticulo, services)
                    strWhere = "IDArticulo = '" & strIDArticulo & "' AND IDEstructura = '" & strArtEstruct & "'"
                End If

                If Length(strWhere) Then
                    Dim e As New Estructura
                    Return e.Filter(, strWhere)
                End If
            End If
        End If
    End Function
    <Task()> Public Shared Function CambioUnidad(ByVal data As UnidadMedidaPrecioInfo, ByVal services As ServiceProvider) As DataTable
        If data.UdValoracion <> 0 Then
            Dim DtCambio As New DataTable
            DtCambio.Columns.Add("QInterna", GetType(Double))
            DtCambio.Columns.Add("PrecioAInterno", GetType(Double))
            DtCambio.Columns.Add("PrecioBInterno", GetType(Double))
            Dim Factor As Double
            Dim DtArt As DataTable = New Articulo().SelOnPrimaryKey(data.IDArticulo)
            If Not DtArt Is Nothing AndAlso DtArt.Rows.Count > 0 Then
                data.IDUdInterna = DtArt.Rows(0)("IDUDInterna") & String.Empty
                If data.IDUdInterna <> data.IDUdMedida Then
                    Dim StDatos As New ArticuloUnidadAB.DatosFactorConversion
                    StDatos.IDArticulo = data.IDArticulo
                    StDatos.IDUdMedidaA = data.IDUdMedida
                    StDatos.IDUdMedidaB = data.IDUdInterna
                    StDatos.UnoSiNoExiste = True
                    Factor = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, StDatos, services)

                    If Factor = 0 Then Factor = 1
                Else
                    Factor = 1
                End If

            End If

            Dim DrNew As DataRow = DtCambio.NewRow
            DrNew("PrecioAInterno") = (data.PrecioA / Factor) / data.UdValoracion
            DrNew("PrecioBInterno") = (data.PrecioB / Factor) / data.UdValoracion
            DrNew("QInterna") = data.Cantidad * Factor
            DtCambio.Rows.Add(DrNew)
            Return DtCambio
        Else
            ApplicationService.GenerateError("Unidad de Valoracion incorrecta.")
        End If
    End Function
    'Public Function FactorConversion(ByVal strIDArticulo As String, ByVal strMedidaA As String, ByVal strMedidaB As String) As Double

    '    '/*************************************** CRITERIO ***************************************************************/
    '    '/Primeramente el factor se obtiene de la tabla tbMaestroArticulo (tbArticuloUnidadAB antes). Si no se puede obtener, por defecto vale 1.   /
    '    '/De ahi en adelante ,en los calculos, se utiliza el campo 'Factor' de la linea, siguiendo SIEMPRE este criterio: /
    '    '/                                                                                                                /
    '    '/                              Q(UdMedida) * Factor = Q(UdInterna)                                               /
    '    '*****************************************************************************************************************

    '    Dim StDatos As New ArticuloUnidadAB.DatosFactorConversion
    '    StDatos.IDArticulo = strIDArticulo
    '    StDatos.IDUdMedidaA = strMedidaA
    '    StDatos.IDUdMedidaB = strMedidaB
    '    StDatos.UnoSiNoExiste = False
    '    Dim Factor As Double = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, StDatos, New ServiceProvider)
    '    If Factor = 0 Then
    '        Dim UDMedida As New UnidadAB.UnidadMedidaInfo
    '        UDMedida.IDUdMedidaA = strMedidaA
    '        UDMedida.IDUdMedidaB = strMedidaB
    '        Factor = ProcessServer.ExecuteTask(Of UnidadAB.UnidadMedidaInfo, Double)(AddressOf UnidadAB.FactorDeConversion, UDMedida, New ServiceProvider)
    '    End If

    '    Return Factor
    'End Function
    <Task()> Public Shared Function ComprobarPreciosNuevos(ByVal dt As DataTable, ByVal services As ServiceProvider) As DataTable
        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            Dim dtNew As DataTable = dt.Clone
            For Each dtRw As DataRow In dt.Rows
                If Length(dtRw("NuevoPrecio")) > 0 Then
                    dtNew.ImportRow(dtRw)
                End If
            Next
            Return dtNew
        End If
    End Function
    <Task()> Public Shared Function GetEstructurasMultiNivel(ByVal data As Estructura.DatosEstrucPrincipal, ByVal services As ServiceProvider) As DataTable
        Return AdminData.Execute("sp_EstructuraInformeMultiNivel", False, data.ArticuloPadre, data.Estructura)
    End Function
    <Task()> Public Shared Function GetTipoEstructurasMultiNivel(ByVal data As Estructura.DatosEstrucPrincipal, ByVal services As ServiceProvider) As DataTable
        Return AdminData.Execute("sp_EstructuraInformeMultiNivelTipoEstructura", False, data.ArticuloPadre, data.TipoEstructura)
    End Function
    <Task()> Public Shared Function ObtenerDatosArticulo(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider) As IPropertyAccessor
        Dim articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        Dim ArtAlm As New DataArtAlm
        If data.ContainsKey("IDArticulo") Then
            ArtAlm.IDArticulo = data("IDArticulo") & String.Empty
        ElseIf data.ContainsKey("IDMaterial") Then
            ArtAlm.IDArticulo = data("IDMaterial") & String.Empty
        End If
        If Length(ArtAlm.IDArticulo) > 0 Then
            Dim infoArticulo As ArticuloInfo = articulos.GetEntity(ArtAlm.IDArticulo)
            data("IDAlmacen") = ProcessServer.ExecuteTask(Of DataArtAlm, String)(AddressOf ArticuloAlmacen.AlmacenPredeterminadoArticulo, ArtAlm, services)
            data("IDProveedor") = ProcessServer.ExecuteTask(Of String, String)(AddressOf ArticuloProveedor.ProveedorPredeterminadoArticulo, ArtAlm.IDArticulo, services)
            data("DescArticulo") = infoArticulo.DescArticulo
            data("DescMaterial") = infoArticulo.DescArticulo
            data("IDUDInterna") = infoArticulo.IDUDInterna
            data("IDUDCompra") = infoArticulo.IDUDCompra
            data("IDUDVenta") = infoArticulo.IDUDVenta
            data("UDValoracion") = infoArticulo.UDValoracion
            data("CCVenta") = infoArticulo.CCVenta
            data("CCCompra") = infoArticulo.CCCompra
            data("TipoFactAlquiler") = infoArticulo.TipoFactAlquiler
            data("IDMaterialOrigen") = infoArticulo.IDArticulo
            data("Configurable") = infoArticulo.Configurable
            data("Activo") = infoArticulo.Activo
        End If

        Return data
    End Function

    <Task()> Public Shared Function ValidaExisteArticulo(ByVal strIDArticulo As String, ByVal services As ServiceProvider) As DataTable
        Dim dt As DataTable = New Articulo().SelOnPrimaryKey(strIDArticulo)
        If IsNothing(dt) OrElse dt.Rows.Count = 0 Then
            ApplicationService.GenerateError("El Artículo | no existe.", strIDArticulo)
        End If
        Return dt
    End Function

    Public Function CuentaContableArticuloPais(ByVal strIDArticulo As String, ByVal strIDPais As String, ByVal intCircuito As Circuito) As String
        Dim strCCArticulo As String

        Dim dtArticulo As DataTable = Me.SelOnPrimaryKey(strIDArticulo)
        If Not IsNothing(dtArticulo) AndAlso dtArticulo.Rows.Count > 0 Then
            Dim objNegPais As New Pais

            If Length(strIDPais) > 0 Then
                Dim dtPais As DataTable = objNegPais.Filter(New StringFilterItem("IDPais", FilterOperator.Equal, strIDPais))
                If IsNothing(dtPais) OrElse dtPais.Rows.Count = 0 Then
                    ApplicationService.GenerateError("El código de país | no existe.", Quoted(strIDPais))
                Else
                    If dtPais.Rows(0)("Extranjero") Then
                        Select Case intCircuito
                            Case Circuito.Ventas
                                strCCArticulo = dtArticulo.Rows(0)("CCExport")
                            Case Circuito.Compras
                                strCCArticulo = dtArticulo.Rows(0)("CCImport")
                        End Select

                    Else
                        Select Case intCircuito
                            Case Circuito.Ventas
                                strCCArticulo = dtArticulo.Rows(0)("CCVenta")
                            Case Circuito.Compras
                                strCCArticulo = dtArticulo.Rows(0)("CCCompra")
                        End Select
                    End If
                End If
            End If
        End If

        Return strCCArticulo
    End Function


    <Serializable()> _
Public Class DatosArticuloEstado
        Public ListArticulos As List(Of String)
        Public IDEstado As String
        'Public dtArticulos As DataTable
        'Public ListTarifas As List(Of String)
        'Public dtClientes As DataTable

        Public Sub New(ByVal ListArticulos As List(Of String), ByVal IDEstado As String)
            Me.ListArticulos = ListArticulos
            Me.IDEstado = IDEstado
        End Sub

        'Public Sub New(ByVal ListClientes As List(Of String), ByVal IDTarifa As String, ByVal dtClientes As DataTable)
        '    Me.ListClientes = ListClientes
        '    Me.IDTarifa = IDTarifa
        '    Me.dtClientes = dtClientes
        'End Sub

        'Public Sub New(ByVal ListClientes As List(Of String), ByVal IDTarifa As String, ByVal ListTarifas As List(Of String))
        '    Me.ListClientes = ListClientes
        '    Me.IDTarifa = IDTarifa
        '    Me.ListTarifas = ListTarifas
        'End Sub
    End Class

    <Task()> Public Shared Function ModificarArticuloEstado(ByVal data As DatosArticuloEstado, ByVal services As ServiceProvider) As Boolean
        Dim sust As Boolean = False
        Dim A As New Articulo
        Dim dt As DataTable = A.Filter()
        If data.ListArticulos.Count > 0 Then
            For value As Integer = 0 To data.ListArticulos.Count - 1
                Dim filtro As New Filter()
                filtro.Add("IDArticulo", data.ListArticulos(value))
                dt = A.Filter(filtro)
                For Each dr As DataRow In dt.Select
                    If data.ListArticulos(value) = dr("IDArticulo") Then
                        dr("IDEstado") = data.IDEstado
                        sust = True
                    End If
                Next
                A.Update(dt)
            Next
        End If
        Return sust
    End Function

#End Region

    <Task()> Public Shared Function CalcularImporteEnMonedaB(ByVal valorMonedaA As Double, ByVal services As ServiceProvider) As Double
        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
        Dim MonInfoA As MonedaInfo = Monedas.GetMoneda(Monedas.MonedaA.ID, Date.Today)
        Dim MonInfoB As MonedaInfo = Monedas.GetMoneda(Monedas.MonedaB.ID)

        Dim dblImporteB As Double = xRound(valorMonedaA * MonInfoA.CambioB, MonInfoB.NDecimalesPrecio)
        Return dblImporteB
    End Function

#Region " Configurador "

    Public Sub UpdtateContadorRadical(ByVal dtArticulos As DataTable)
        If Not IsNothing(dtArticulos) AndAlso dtArticulos.Rows.Count > 0 Then
            For Each drArticulo As DataRow In dtArticulos.Rows
                Dim dr As DataRow = Me.GetItemRow(drArticulo("IDArticulo"))
                dr("ContRadical") = dr("ContRadical") + 1
                MyBase.Update(dr.Table)
            Next
        End If
    End Sub

    Public Function EsConfigurable(ByVal IDArticulo As String) As Boolean
        Dim blnConfigurable As Boolean
        Dim dtArticulo As DataTable = Me.SelOnPrimaryKey(IDArticulo)
        If Not IsNothing(dtArticulo) AndAlso dtArticulo.Rows.Count > 0 Then
            blnConfigurable = Nz(dtArticulo.Rows(0)("Configurable"), False)
        End If
        Return blnConfigurable
    End Function

#End Region

    <Serializable()> _
    Public Class DataInfoArticulo
        Public IDArticulo As String
        Public CodigoBarras As String
        Public RefArticulo As String
        Public context As IPropertyAccessor

        Public Sub New(ByVal IDArticulo As String)
            Me.IDArticulo = IDArticulo
        End Sub

        Public Sub New(ByVal IDArticulo As String, ByVal CodigoBarras As String)
            If Length(IDArticulo) > 0 Then Me.IDArticulo = IDArticulo
            Me.CodigoBarras = CodigoBarras
        End Sub

        Public Sub New(ByVal IDArticulo As String, ByVal CodigoBarras As String, ByVal RefArticulo As String, ByVal context As IPropertyAccessor)
            If Length(IDArticulo) > 0 Then Me.IDArticulo = IDArticulo
            If Length(CodigoBarras) > 0 Then Me.CodigoBarras = CodigoBarras
            Me.RefArticulo = RefArticulo
            Me.context = context
        End Sub
    End Class

    <Task()> Public Shared Function InformacionArticulo(ByVal dataArticulo As DataInfoArticulo, ByVal services As ServiceProvider) As ArticuloInfo
        Dim ArtInfo As New ArticuloInfo
        Dim strMensaje As String
        Dim dt As DataTable
        Dim CARACT_ARTICULO As String = "vNegCaractArticulo"

        If Length(dataArticulo.CodigoBarras) > 0 Then
            dt = New BE.DataEngine().Filter(CARACT_ARTICULO, New StringFilterItem("CodigoBarras", dataArticulo.CodigoBarras))
            strMensaje = "El artículo con el código de barras " & Quoted(dataArticulo.CodigoBarras) & " no existe."
        End If

        If Length(dataArticulo.RefArticulo) > 0 Then
            If Not IsNothing(dataArticulo.context) Then
                If dataArticulo.context.ContainsKey("IDCliente") AndAlso Length(dataArticulo.context("IDCliente")) > 0 Then
                    Dim StDatosCli As New ArticuloCliente.DatosArtRef
                    StDatosCli.IDCliente = dataArticulo.context("IDCliente")
                    StDatosCli.Referencia = dataArticulo.RefArticulo
                    ArtInfo.IDArticulo = ProcessServer.ExecuteTask(Of ArticuloCliente.DatosArtRef, String)(AddressOf ArticuloCliente.ObtenerArticuloRef, StDatosCli, services)
                End If
                If dataArticulo.context.ContainsKey("IDProveedor") AndAlso Length(dataArticulo.context("IDProveedor")) > 0 Then
                    Dim StDatos As New ArticuloProveedor.DatosArtRef
                    StDatos.IDProveedor = dataArticulo.context("IDProveedor")
                    StDatos.Referencia = dataArticulo.RefArticulo
                    ArtInfo.IDArticulo = ProcessServer.ExecuteTask(Of ArticuloProveedor.DatosArtRef, String)(AddressOf ArticuloProveedor.ObtenerArticuloRef, StDatos, services)
                End If
            End If

            If Length(ArtInfo.IDArticulo) > 0 Then
                dt = New BE.DataEngine().Filter(CARACT_ARTICULO, New StringFilterItem("IDArticulo", ArtInfo.IDArticulo))
            Else
                strMensaje = "El artículo con la referencia " & Quoted(dataArticulo.RefArticulo) & " no existe."
            End If
        End If

        If Length(dataArticulo.IDArticulo) > 0 Then
            dt = New BE.DataEngine().Filter(CARACT_ARTICULO, New StringFilterItem("IDArticulo", dataArticulo.IDArticulo))
            strMensaje = "El artículo " & Quoted(dataArticulo.IDArticulo) & " no existe."
        End If

        If IsNothing(dt) OrElse dt.Rows.Count = 0 Then
            ApplicationService.GenerateError(strMensaje)
        Else
            ArtInfo = New ArticuloInfo
            ArtInfo.Fill(dt.Rows(0)("IDArticulo")) '.CriterioValoracion = Nz(dt.Rows(0)("CriterioValoracion"), enumtaValoracion.taPrecioEstandar)
        End If

        Return ArtInfo
    End Function

#Region " Punto Verde "

    <Task()> Public Shared Function TratarEcoPuntoVerde(ByVal fFiltros As Filter, ByVal services As ServiceProvider) As DataTable
        Dim clsEstructura As New ArticuloEstructura
        Dim CantidadVendidaPadre As Double = 0
        Dim PuntoVerdePadre As Double = 0

        Dim dtTablaRdo As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf CrearEstructuraTablaResultado, Nothing, services)

        Dim dtArticulosPadre As DataTable = New BE.DataEngine().Filter("vBdgConsultaEcoVidrio", fFiltros)
        If Not IsNothing(dtArticulosPadre) AndAlso dtArticulosPadre.Rows.Count > 0 Then 'Si hay articulos se cocinan
            Dim drFilaNueva As DataRow
            Dim QServida As Double
            Dim IDArticulo As String = "-2"
            Dim filtroArticulo As New Filter
            For Each drArticuloPadre As DataRow In dtArticulosPadre.Rows 'Por cada padre recoger su PuntoVerde si tiene y su Cantidad. Realizar Desglose               
                CantidadVendidaPadre = 0
                QServida = 0 '
                'Para q no haga el mismo proceso por cada articulo de la tabla y solo lo haga una vez                
                If IDArticulo <> drArticuloPadre("IDArticulo") Then
                    'Hacer un sumatorio de la QServida por cada articulo q haya en la tabla
                    filtroArticulo.Clear()
                    filtroArticulo.Add("IDArticulo", FilterOperator.Equal, drArticuloPadre("IDArticulo"))
                    For Each drFilaSum As DataRow In dtArticulosPadre.Select(filtroArticulo.Compose(New AdoFilterComposer))
                        QServida += drFilaSum("Vendido")
                    Next

                    If Length(drArticuloPadre("Vendido")) > 0 Then
                        CantidadVendidaPadre = CantidadVendidaPadre + QServida
                    End If
                    'Si tiene Punto Verde
                    If Length(drArticuloPadre("PuntoVerde")) > 0 AndAlso drArticuloPadre("PuntoVerde") > 0 Then
                        fFiltros.Clear()
                        fFiltros.Add("IDArticulo", FilterOperator.Equal, drArticuloPadre("IDArticulo"))
                        Dim drFila() = dtTablaRdo.Select(fFiltros.Compose(New AdoFilterComposer))
                        If IsNothing(drFila) AndAlso drFila.Length > 0 Then
                            'Si hay articulo igual introducir los datos en la fila correspondiente
                            drFila("NumEnvases") = drFila("NumEnvases") + CantidadVendidaPadre
                            drFila("KGSMaterial") = drFila("KGSMaterial") + (CantidadVendidaPadre * drArticuloPadre("PesoBruto"))
                            drFila("LitrosProducto") = drFila("LitrosProducto") + (CantidadVendidaPadre * drArticuloPadre("Volumen"))
                            drFila("ImporteEnvases") = drFila("ImporteEnvases") + (CantidadVendidaPadre * drArticuloPadre("PuntoVerde"))
                            drFila("ImporteKilos") = drFila("ImporteKilos") + ((CantidadVendidaPadre * drArticuloPadre("PesoBruto")) * drArticuloPadre("PuntoVerde"))
                        Else
                            'Si no hay articulo igual en la tabla, nueva fila
                            Dim datosCopiaPadre As New dataCopiaDatosPadre(drArticuloPadre, dtTablaRdo.NewRow(), CantidadVendidaPadre)
                            drFilaNueva = ProcessServer.ExecuteTask(Of dataCopiaDatosPadre, DataRow)(AddressOf CopiarDatosPadre, datosCopiaPadre, services)
                            dtTablaRdo.Rows.Add(drFilaNueva)
                        End If
                    End If
                    'Comenzar Explosion Articulo
                    Dim CantidadProducto As Double = 0
                    Dim Paso As Boolean = False
                    Dim dtHijoExplosion As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf ArticuloEstructura.CalcularEstructuraExplosion, drArticuloPadre("IDArticulo"), services)
                    If Not IsNothing(dtHijoExplosion) AndAlso dtHijoExplosion.Rows.Count > 0 Then
                        For Each drFilaHijo As DataRow In dtHijoExplosion.Select("", "Nivel")
                            'Recoger la cantidad q le corresponde a cada componente si es hijo o no
                            fFiltros.Clear()
                            fFiltros.Add("ID", FilterOperator.Equal, drFilaHijo("IDPadre"))
                            For Each drPadre As DataRow In dtHijoExplosion.Select(fFiltros.Compose(New AdoFilterComposer))
                                'Tiene padre
                                If drPadre("IDPadre") <> 0 Then
                                    Dim dataCantidadEnHijos As New dataRecogerCantidadEnHijos(dtHijoExplosion, CantidadProducto, drFilaHijo("IDPadre"), CInt(drFilaHijo("Cantidad")))
                                    ProcessServer.ExecuteTask(Of dataRecogerCantidadEnHijos)(AddressOf RecogerCantidadEnHijos, dataCantidadEnHijos, services)
                                    Paso = True
                                Else
                                    CantidadProducto = CInt(drFilaHijo("Cantidad")) * CInt(drPadre("Cantidad"))
                                    Paso = True
                                End If
                            Next
                            'Fin recoger
                            If Paso = True Then
                                CantidadProducto = CantidadProducto * CantidadVendidaPadre
                            Else
                                CantidadProducto = CType(drFilaHijo("Cantidad"), Double) * CantidadVendidaPadre
                            End If

                            Dim datosNivelInferior As New dataTratarNivelInferior(drFilaHijo, dtTablaRdo, CantidadProducto)
                            ProcessServer.ExecuteTask(Of dataTratarNivelInferior)(AddressOf TratarNivelInferior, datosNivelInferior, services)
                        Next
                    End If
                End If
                IDArticulo = drArticuloPadre("IDArticulo")
            Next
        End If
        Return dtTablaRdo
    End Function

    <Serializable()> _
    Public Class dataRecogerCantidadEnHijos
        Public dtTabla As DataTable
        Public CantidadProducto As Integer
        Public IDPadre As String
        Public CantidadActual As Integer

        Public Sub New(ByVal dtTabla As DataTable, ByRef CantidadProducto As Integer, ByVal IDPadre As String, ByVal CantidadActual As Integer)
            Me.dtTabla = dtTabla
            Me.CantidadProducto = CantidadProducto
            Me.IDPadre = IDPadre
            Me.CantidadActual = CantidadActual
        End Sub
    End Class
    <Task()> Public Shared Sub RecogerCantidadEnHijos(ByVal data As dataRecogerCantidadEnHijos, ByVal services As ServiceProvider)
        Dim fFiltros As New Filter
        fFiltros.Add("ID", FilterOperator.Equal, data.IDPadre)
        For Each drPadre As DataRow In data.dtTabla.Select(fFiltros.Compose(New AdoFilterComposer))
            'Tiene padre
            If drPadre("IDPadre") <> 0 Then
                data.CantidadProducto = (data.CantidadActual * CInt(drPadre("Cantidad")))
                data.IDPadre = drPadre("IDPadre")
                data.CantidadActual = CInt(drPadre("Cantidad"))
                ProcessServer.ExecuteTask(Of dataRecogerCantidadEnHijos)(AddressOf RecogerCantidadEnHijos, data, services)
            Else
                data.CantidadProducto = (data.CantidadActual * CInt(drPadre("Cantidad"))) * data.CantidadProducto
            End If
        Next
    End Sub

    <Serializable()> _
    Public Class dataTratarNivelInferior
        Public drFilaExplosion As DataRow
        Public dtTablaRdo As DataTable
        Public CantidadPadre As Double

        Public Sub New(ByVal drFilaExplosion As DataRow, ByRef dtTablaRdo As DataTable, ByVal CantidadPadre As Double)
            Me.drFilaExplosion = drFilaExplosion
            Me.dtTablaRdo = dtTablaRdo
            Me.CantidadPadre = CantidadPadre
        End Sub
    End Class
    <Task()> Public Shared Sub TratarNivelInferior(ByVal data As dataTratarNivelInferior, ByVal services As ServiceProvider)
        Dim NumEnvasesActual As Double = 0
        Dim ffilter As New Filter
        ffilter.Add("IDArticulo", FilterOperator.Equal, data.drFilaExplosion("IDComponente")) 'Nuevo hijo y nivel
        'Recoger datos de ArticuloNuevo
        Dim dtArticulo As DataTable = New Articulo().Filter(ffilter)
        If Not IsNothing(dtArticulo) AndAlso dtArticulo.Rows.Count > 0 Then
            If Length(dtArticulo.Rows(0)("PuntoVerde")) > 0 AndAlso dtArticulo.Rows(0)("PuntoVerde") > 0 Then
                Dim drFila() = data.dtTablaRdo.Select(ffilter.Compose(New AdoFilterComposer))
                If Not IsNothing(drFila) AndAlso drFila.Length > 0 Then
                    NumEnvasesActual = data.CantidadPadre
                    'Si hay articulo igual introducir los datos en la fila correspondiente
                    drFila(0)("NumEnvases") = Convert.ToInt32(drFila(0)("NumEnvases")) + NumEnvasesActual
                    drFila(0)("KGSMaterial") = drFila(0)("KGSMaterial") + (NumEnvasesActual * dtArticulo.Rows(0)("PesoBruto"))
                    drFila(0)("LitrosProducto") = drFila(0)("LitrosProducto") + (NumEnvasesActual * dtArticulo.Rows(0)("Volumen"))
                    drFila(0)("ImporteEnvases") = drFila(0)("ImporteEnvases") + (NumEnvasesActual * dtArticulo.Rows(0)("PuntoVerde"))
                    drFila(0)("ImporteKilos") = drFila(0)("ImporteKilos") + ((NumEnvasesActual * dtArticulo.Rows(0)("PesoBruto")) * dtArticulo.Rows(0)("PuntoVerde"))
                Else
                    'Si no hay articulo igual en la tabla, nueva fila
                    Dim datosCopiaPadre As New dataCopiaDatosPadre(dtArticulo.Rows(0), data.dtTablaRdo.NewRow(), data.CantidadPadre)
                    Dim drFilaNueva As DataRow = ProcessServer.ExecuteTask(Of dataCopiaDatosPadre, DataRow)(AddressOf CopiarDatosPadre, datosCopiaPadre, services)

                    data.dtTablaRdo.Rows.Add(drFilaNueva)
                End If
            End If 'No tiene Punto verde  
        End If  'No tiene datos
    End Sub

    <Serializable()> _
    Public Class dataCopiaDatosPadre
        Public drArticuloActual, drFilaNueva As DataRow
        Public CantidadPadre As Double

        Public Sub New(ByVal drArticuloActual As DataRow, ByVal drFilaNueva As DataRow, ByVal CantidadPadre As Double)
            Me.drArticuloActual = drArticuloActual
            Me.drFilaNueva = drFilaNueva
            Me.CantidadPadre = CantidadPadre
        End Sub
    End Class
    <Task()> Public Shared Function CopiarDatosPadre(ByVal data As dataCopiaDatosPadre, ByVal services As ServiceProvider) As DataRow
        Dim NumEnvasesActual As Double = data.CantidadPadre

        data.drFilaNueva("IDArticulo") = data.drArticuloActual("IDArticulo")
        If Length(data.drArticuloActual("DescArticulo")) > 0 Then
            data.drFilaNueva("DescArticulo") = data.drArticuloActual("DescArticulo")
        End If
        If Length(data.drArticuloActual("IDTipo")) > 0 Then
            data.drFilaNueva("IDTipo") = data.drArticuloActual("IDTipo")
            data.drFilaNueva("DescTipo") = New TipoArticulo().GetItemRow(data.drArticuloActual("IDTipo"))("DescTipo") ' ObtenerDescripcionTipo(data.drArticuloActual("IDTipo"))
        End If
        If Length(data.drArticuloActual("IDFamilia")) > 0 Then
            data.drFilaNueva("IDFamilia") = data.drArticuloActual("IDFamilia")
            data.drFilaNueva("DescFamilia") = New Familia().GetItemRow(data.drArticuloActual("IDTipo"), data.drArticuloActual("IDFamilia"))("DescFamilia") ' ObtenerDescripcionFamilia(data.drArticuloActual("IDFamilia"), data.drArticuloActual("IDTipo"))
        End If
        If Length(data.drArticuloActual("IDSubFamilia")) > 0 Then
            data.drFilaNueva("IDSubFamilia") = data.drArticuloActual("IDSubFamilia")
            data.drFilaNueva("DescSubFamilia") = New Subfamilia().GetItemRow(data.drArticuloActual("IDTipo"), data.drArticuloActual("IDFamilia"), data.drArticuloActual("IDSubFamilia"))("DescSubFamilia") 'ObtenerDescripcionSubFamilia(data.drArticuloActual("IDSubFamilia"), data.drArticuloActual("IDFamilia"), data.drArticuloActual("IDTipo")) 'drArticuloActual("DescSubFamilia")   
        End If
        data.drFilaNueva("NumEnvases") = NumEnvasesActual
        data.drFilaNueva("KGSMaterial") = (NumEnvasesActual * data.drArticuloActual("PesoBruto"))
        data.drFilaNueva("LitrosProducto") = (NumEnvasesActual * data.drArticuloActual("Volumen"))
        data.drFilaNueva("ImporteEnvases") = (NumEnvasesActual * data.drArticuloActual("PuntoVerde"))
        data.drFilaNueva("ImporteKilos") = (NumEnvasesActual * data.drArticuloActual("PesoBruto")) * data.drArticuloActual("PuntoVerde")

        Return data.drFilaNueva
    End Function

    <Task()> Public Shared Function CrearEstructuraTablaResultado(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
        Dim dtTablaRdo As New DataTable
        dtTablaRdo.Columns.Add("IDArticulo")
        dtTablaRdo.Columns.Add("DescArticulo")
        dtTablaRdo.Columns.Add("IDTipo")
        dtTablaRdo.Columns.Add("DescTipo")
        dtTablaRdo.Columns.Add("IDFamilia")
        dtTablaRdo.Columns.Add("DescFamilia")
        dtTablaRdo.Columns.Add("IDSubFamilia")
        dtTablaRdo.Columns.Add("DescSubFamilia")
        dtTablaRdo.Columns.Add("NumEnvases", GetType(Double))
        dtTablaRdo.Columns.Add("KGSMaterial", GetType(Double))
        dtTablaRdo.Columns.Add("LitrosProducto", GetType(Double))
        dtTablaRdo.Columns.Add("ImporteEnvases", GetType(Double))
        dtTablaRdo.Columns.Add("ImporteKilos", GetType(Double))

        Return dtTablaRdo
    End Function

    <Serializable()> _
    Public Class datCalcularDeclaAnualPV
        Public FechaFacturaDesde As DateTime
        Public FechaFacturaHasta As DateTime

        Public Sub New()
        End Sub
        Public Sub New(ByVal FechaFacturaDesde As DateTime, ByVal FechaFacturaHasta As DateTime)
            Me.FechaFacturaDesde = FechaFacturaDesde
            Me.FechaFacturaHasta = FechaFacturaHasta
        End Sub
    End Class

    <Task()> Public Shared Function CalcularDeclaracionAnualPuntoVerde(ByVal pFecha As datCalcularDeclaAnualPV, ByVal services As ServiceProvider) As DataTable
        Dim SqlCmd As Common.DbCommand = AdminData.GetCommand()
        SqlCmd.CommandText = "sp_EcoembesDeclaracionAnual"
        SqlCmd.CommandType = CommandType.StoredProcedure

        Dim SqlParam1 As Common.DbParameter = SqlCmd.CreateParameter
        SqlParam1.DbType = DbType.DateTime
        SqlParam1.Direction = ParameterDirection.Input
        SqlParam1.ParameterName = "@pFechaFacturaDesde"
        SqlParam1.Value = pFecha.FechaFacturaDesde
        SqlCmd.Parameters.Add(SqlParam1)

        Dim SqlParam2 As Common.DbParameter = SqlCmd.CreateParameter
        SqlParam2.DbType = DbType.DateTime
        SqlParam2.Direction = ParameterDirection.Input
        SqlParam2.ParameterName = "@pFechaFacturaHasta"
        SqlParam2.Value = pFecha.FechaFacturaHasta
        SqlCmd.Parameters.Add(SqlParam2)

        Dim SqlParam3 As Common.DbParameter = SqlCmd.CreateParameter
        SqlParam3.DbType = DbType.Int32
        SqlParam3.Direction = ParameterDirection.Input
        SqlParam3.ParameterName = "@pGenerarFicheros"
        SqlParam3.Value = 0
        SqlCmd.Parameters.Add(SqlParam3)

        Return AdminData.Execute(SqlCmd, ExecuteCommand.ExecuteReader)
    End Function

    <Serializable()> _
    Public Class datPuntoVerde
        Public FilArticulos As Filter
        Public PuntoVerde As Double

        Public Sub New()
        End Sub
        Public Sub New(ByVal FilArticulos As Filter, ByVal PuntoVerde As Double)
            Me.FilArticulos = FilArticulos
            Me.PuntoVerde = PuntoVerde
        End Sub
    End Class

    <Task()> Public Shared Sub ActualizacionDePuntoVerde(ByVal data As datPuntoVerde, ByVal services As ServiceProvider)
        Dim StrSQL As String = "UPDATE FROM tbMaestroArticulo SET PuntoVerde = " & data.PuntoVerde & " WHERE (" & AdminData.ComposeFilter(data.FilArticulos) & ")"
        AdminData.Execute(StrSQL)
    End Sub

#End Region

#Region " Gestión Doble Unidad "

    <Serializable()> _
    Public Class DataFactorDobleUnidad
        Public IDArticulo As String
        Public IDUdInterna As String
        Public CambioIDUDInterna As Boolean
        Public CambioIDUDInterna2 As Boolean

        Public Sub New(ByVal IDArticulo As String, ByVal CambioIDUDInterna As Boolean, ByVal CambioIDUDInterna2 As Boolean, Optional ByVal IDUdInterna As String = Nothing)
            Me.IDArticulo = IDArticulo
            If Length(IDUdInterna) > 0 Then Me.IDUdInterna = IDUdInterna
            Me.CambioIDUDInterna = CambioIDUDInterna
            Me.CambioIDUDInterna2 = CambioIDUDInterna2
        End Sub
    End Class
    <Task()> Public Shared Function FactorDobleUnidad(ByVal data As DataFactorDobleUnidad, ByVal services As ServiceProvider) As Double
        If ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.IDArticulo, services) Then
            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
            Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.IDArticulo)

            If Length(data.IDUdInterna) = 0 Then data.IDUdInterna = ArtInfo.IDUDInterna
            If data.CambioIDUDInterna OrElse data.CambioIDUDInterna2 Then
                Dim datFactor As New ArticuloUnidadAB.DatosFactorConversion
                If data.CambioIDUDInterna Then
                    datFactor = New ArticuloUnidadAB.DatosFactorConversion(data.IDArticulo, data.IDUdInterna, ArtInfo.IDUDInterna2, False)
                Else
                    datFactor = New ArticuloUnidadAB.DatosFactorConversion(data.IDArticulo, ArtInfo.IDUDInterna2, data.IDUdInterna, False)
                End If

                Dim Factor As Double = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf FactorDeConversionUnidadesInternas, datFactor, services)
                Return Factor
            End If

        Else
            Return -1
        End If
    End Function

    '//Tarea específica para el tratamiento de la segunda unidad
    <Task()> Public Shared Function FactorDeConversionUnidadesInternas(ByVal data As ArticuloUnidadAB.DatosFactorConversion, ByVal services As ServiceProvider) As Double
        Dim oFltr As Filter = New Filter(FilterUnionOperator.Or)
        Dim dblFactor As Double
        Dim blnDividir As Boolean

        Dim oFltrA As Filter = New Filter
        oFltrA.Add("IDArticulo", FilterOperator.Equal, data.IDArticulo)
        If Length(data.IDUdMedidaA) > 0 Then oFltrA.Add("IDUdMedidaA", FilterOperator.Equal, data.IDUdMedidaA)
        If Length(data.IDUdMedidaB) > 0 Then oFltrA.Add("IDUdMedidaB", FilterOperator.Equal, data.IDUdMedidaB)
        oFltr.Add(oFltrA)

        'Dim oFltrB As Filter = New Filter
        'oFltrB.Add("IDArticulo", FilterOperator.Equal, data.IDArticulo)
        'If Length(data.IDUdMedidaB) > 0 Then oFltrB.Add("IDUdMedidaA", FilterOperator.Equal, data.IDUdMedidaB)
        'If Length(data.IDUdMedidaA) > 0 Then oFltrB.Add("IDUdMedidaB", FilterOperator.Equal, data.IDUdMedidaA)
        'oFltr.Add(oFltrB)

        Dim dt As DataTable = New ArticuloUnidadAB().Filter(oFltr)

        Select Case dt.Rows.Count
            Case 0

                'Dim UDMedida As New UnidadAB.UnidadMedidaInfo
                'UDMedida.IDUdMedidaA = data.IDUdMedidaA
                'UDMedida.IDUdMedidaB = data.IDUdMedidaB
                'UDMedida.UnoSiNoExiste = data.UnoSiNoExiste
                'dblFactor = ProcessServer.ExecuteTask(Of UnidadAB.UnidadMedidaInfo, Double)(AddressOf UnidadAB.FactorDeConversion, UDMedida, services)
                'If dblFactor = 0 Then
                '    If data.UnoSiNoExiste Then
                '        dblFactor = 1
                '    Else : dblFactor = 0
                '    End If
                'End If
            Case 1
                Dim oRw As DataRow = dt.Rows(0)
                dblFactor = oRw("Factor")
                blnDividir = (CStr(oRw("IDUdMedidaA")) = data.IDUdMedidaB)
                'Case 2
                '    If CStr(dt.Rows(0)("IDUdMedidaA")) = data.IDUdMedidaA Then
                '        dblFactor = dt.Rows(0)("Factor")
                '    Else : dblFactor = dt.Rows(1)("Factor")
                '    End If
        End Select
        If dblFactor = 0 AndAlso data.UnoSiNoExiste Then dblFactor = 1
        If blnDividir Then
            If dblFactor <> 0 Then
                Return 1 / dblFactor
            Else : Return 0
            End If
        Else : Return dblFactor
        End If
    End Function


#End Region

#Region " Integración con Solid Works "

    <Task()> Public Shared Function GenerarObraArticulo(ByVal IDArticulo As String, ByVal services As ServiceProvider) As Integer
        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(IDArticulo)
        'Dim dtArticulo As DataTable = New Articulo().SelOnPrimaryKey(IDArticulo)
        'If Not ArtInfo Is Nothing Then
        Dim Obra As BusinessHelper = CreateBusinessObject("ObraCabecera")
        Dim Trabajo As BusinessHelper = CreateBusinessObject("ObraTrabajo")
        Dim MatPrev As BusinessHelper = CreateBusinessObject("ObraMaterial")

        Dim dtObra As DataTable = Obra.Filter(New FilterItem("NObra", FilterOperator.Equal, IDArticulo))
        Dim dtObraTrabajo As DataTable
        Dim dtObraMaterial As DataTable

        If dtObra Is Nothing OrElse dtObra.Rows.Count = 0 Then
            dtObra = Obra.AddNewForm
            dtObra.Rows(0)("IDContador") = System.DBNull.Value
            dtObra.Rows(0)("NObra") = IDArticulo
            dtObra.Rows(0)("DescObra") = ArtInfo.DescArticulo

            Dim p As New Parametro
            dtObra.Rows(0)("IDTipoObra") = p.TipoProyectoPredeterminado
            dtObra.Rows(0)("TipoMnto") = enumTipoObra.tpObra

            Dim context As New BusinessData
            Obra.ApplyBusinessRule("IDCliente", p.ClienteAutofactura, dtObra.Rows(0), context)
            Obra.ApplyBusinessRule("IDCentroGestion", p.CGestionPredet, dtObra.Rows(0), context)

            ' Dim ContextObra As New BusinessData(dtObra.Rows(0))
            dtObraTrabajo = Trabajo.AddNew
            dtObraMaterial = MatPrev.AddNew
        Else
            dtObraTrabajo = Trabajo.Filter(New FilterItem("IDObra", FilterOperator.Equal, dtObra.Rows(0)("IDObra")))
            dtObraMaterial = MatPrev.Filter(New FilterItem("IDObra", FilterOperator.Equal, dtObra.Rows(0)("IDObra")))
        End If
        Dim ContextObra As New BusinessData(dtObra.Rows(0))
        Dim datNuevoArticuloEnObra As New DataInsertarArticuloEnObra(dtObra.Rows(0)("IDObra"), dtObra.Rows(0)("IDTipoObra"), IDArticulo, ArtInfo.DescArticulo, 1, dtObraTrabajo, dtObraMaterial, Nothing, Nothing, Nothing)
        datNuevoArticuloEnObra.ContextObra = ContextObra
        ProcessServer.ExecuteTask(Of DataInsertarArticuloEnObra)(AddressOf InsertarArticuloEnObra, datNuevoArticuloEnObra, services)

        AdminData.BeginTx()
        Dim pck As New UpdatePackage(dtObra)
        pck.Add(dtObraTrabajo)
        pck.Add(dtObraMaterial)
        Obra.Update(pck)
        AdminData.CommitTx(True)

        Return Nz(dtObra.Rows(0)("IDObra"), -1)
    End Function

    Public Class DataInsertarArticuloEnObra
        Public IDObra As Integer
        Public IDTipoObra As String
        Public IDArticulo As String
        Public DescArticulo As String

        Public Cantidad As Double
        Public IDTrabajoPadre As Integer?
        Public CodTrabajoPadre As String
        Public QPrevTrabajo As Double?

        'Public RowDetalle As DataRow
        Public ContextObra As BusinessData

        'Public EstadoGenerado As Boolean
        'Public IDEvaluador As Integer?

        Public Trabajos As DataTable
        Public Materiales As DataTable

        Public Sub New(ByVal IDObra As Integer, ByVal IDTipoObra As String, ByVal IDArticulo As String, ByVal DescArticulo As String, ByVal Cantidad As Double, ByVal Trabajos As DataTable, ByVal Materiales As DataTable, ByVal IDTrabajoPadre As Integer?, ByVal CodTrabajoPadre As String, ByVal QPrevTrabajo As Double?)
            Me.IDObra = IDObra
            Me.IDTipoObra = IDTipoObra
            Me.IDArticulo = IDArticulo
            Me.DescArticulo = DescArticulo
            Me.Trabajos = Trabajos
            Me.Materiales = Materiales
            Me.Cantidad = Cantidad


            If Not IDTrabajoPadre Is Nothing Then Me.IDTrabajoPadre = IDTrabajoPadre
            If Not CodTrabajoPadre Is Nothing Then Me.CodTrabajoPadre = CodTrabajoPadre
            If Not QPrevTrabajo Is Nothing Then
                Me.QPrevTrabajo = QPrevTrabajo
            Else
                Me.QPrevTrabajo = 1
            End If
        End Sub
    End Class
    <Task()> Public Shared Sub InsertarArticuloEnObra(ByVal data As DataInsertarArticuloEnObra, ByVal services As ServiceProvider)
        Dim Trabajo As BusinessHelper = CreateBusinessObject("ObraTrabajo")
        ' Dim fArticuloEstructura As New Filter
        Dim fArticulo As New Filter
        fArticulo.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
        ' fArticuloEstructura.Add(fArticulo)
        'If data.EstadoGenerado Then
        '    fArticuloGenerado.Add(New NumberFilterItem("Estado", EnumEstadoEvaluador.Generado))
        '    'If Not data.IDEvaluador Is Nothing Then fArticuloGenerado.Add(New NumberFilterItem("IDEvaluador", data.IDEvaluador))
        'End If
        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.IDArticulo)
        Dim dtArticuloEstructura As DataTable = New BE.DataEngine().Filter("tbArticuloEstructura", fArticulo)
        If dtArticuloEstructura.Rows.Count > 0 Then
            '//Articulo configurado. NuevoTrabajo()
            If data.Cantidad <> 0 Then
                Dim FilTrabajo As New Filter
                FilTrabajo.Add("CodTrabajo", FilterOperator.Equal, data.IDArticulo)
                FilTrabajo.Add("IDObra", FilterOperator.Equal, data.IDObra)
                FilTrabajo.Add("IDTrabajoPadre", FilterOperator.Equal, data.IDTrabajoPadre)
                Dim dtTrabajo As DataTable = Trabajo.Filter(FilTrabajo)
                If dtTrabajo Is Nothing OrElse dtTrabajo.Rows.Count = 0 Then
                    dtTrabajo = Trabajo.AddNewForm
                    dtTrabajo.Rows(0)("IDObra") = data.IDObra
                    dtTrabajo.Rows(0)("IDTipoObra") = data.IDTipoObra
                    dtTrabajo.Rows(0)("CodTrabajo") = data.IDArticulo
                    dtTrabajo.Rows(0)("DescTrabajo") = data.DescArticulo
                    dtTrabajo.Rows(0)("QPrev") = data.Cantidad
                    dtTrabajo.Rows(0)("NoAcumular") = True
                    Trabajo.ApplyBusinessRule("ImpPrevTrabajoA", ArtInfo.PrecioEstandarA, dtTrabajo.Rows(0), data.ContextObra)

                    If Not data.IDTrabajoPadre Is Nothing Then dtTrabajo.Rows(0)("IDTrabajoPadre") = data.IDTrabajoPadre
                    data.Trabajos.ImportRow(dtTrabajo.Rows(0))

                    'Dim fCompoEvaluacion As New Filter
                    'fCompoEvaluacion.Add(fArticulo)
                End If
                Dim dtComponentes As DataTable = New BE.DataEngine().Filter("vNegComponentesArticulo", fArticulo)
                For Each drComponente As DataRow In dtComponentes.Rows
                    Dim datNuevoArticuloEnObra As New DataInsertarArticuloEnObra(data.IDObra, data.IDTipoObra, drComponente("IDComponente"), drComponente("DescComponente"), drComponente("Cantidad"), data.Trabajos, data.Materiales, dtTrabajo.Rows(0)("IDTrabajo"), dtTrabajo.Rows(0)("CodTrabajo"), CDbl(dtTrabajo.Rows(0)("QPrev")))
                    datNuevoArticuloEnObra.ContextObra = data.ContextObra
                    ProcessServer.ExecuteTask(Of DataInsertarArticuloEnObra)(AddressOf InsertarArticuloEnObra, datNuevoArticuloEnObra, services)
                Next
            End If
        Else
            '//Artículo normal. Nueva linea de ObraMaterialPrev
            If data.Cantidad <> 0 Then
                Dim MatPrev As BusinessHelper = CreateBusinessObject("ObraMaterial")

                Dim dtMaterial As DataTable
                Dim FilMat As New Filter
                If Not data.IDTrabajoPadre Is Nothing Then FilMat.Add("IDTrabajo", FilterOperator.Equal, data.IDTrabajoPadre)
                FilMat.Add("IDMaterial", FilterOperator.Equal, data.IDArticulo)
                dtMaterial = MatPrev.Filter(FilMat)
                If dtMaterial Is Nothing OrElse dtMaterial.Rows.Count = 0 Then
                    dtMaterial = MatPrev.AddNewForm
                    dtMaterial.Rows(0)("IDObra") = data.IDObra
                    dtMaterial.Rows(0)("IDMaterial") = data.IDArticulo
                    dtMaterial.Rows(0)("DescMaterial") = data.DescArticulo
                    'dtMaterial.Rows(0)("QPrev") = data.Cantidad
                    dtMaterial.Rows(0)("UdValoracion") = 1
                    MatPrev.ApplyBusinessRule("QUnidad", data.Cantidad, dtMaterial.Rows(0), data.ContextObra)
                    MatPrev.ApplyBusinessRule("QPrev", data.Cantidad * data.QPrevTrabajo, dtMaterial.Rows(0), data.ContextObra)

                    'tbObraMaterial (QPrev) = tbObraMaterial(QUnidad) * tbObraTrabajo(QPrev)
                    'If Not data.RowDetalle Is Nothing Then
                    dtMaterial.Rows(0)("DtoVenta1") = 0
                    dtMaterial.Rows(0)("DtoVenta2") = 0
                    dtMaterial.Rows(0)("DtoVenta3") = 0
                    MatPrev.ApplyBusinessRule("PrecioPrevMatA", ArtInfo.PrecioEstandarA, dtMaterial.Rows(0), data.ContextObra)
                    'Else
                    '    dtMaterial.Rows(0)("DtoVenta1") = 0
                    '    dtMaterial.Rows(0)("DtoVenta2") = 0
                    '    dtMaterial.Rows(0)("DtoVenta3") = 0
                    'End If

                    If Not data.IDTrabajoPadre Is Nothing Then
                        dtMaterial.Rows(0)("IDTrabajo") = data.IDTrabajoPadre
                    Else
                        Dim IDTrabajoPred As Integer?
                        Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
                        Dim CodTrabajoPred As String = AppParams.TrabajoPredeterminado
                        Dim adrTrabajoPred() As DataRow = data.Trabajos.Select("CodTrabajo='" & CodTrabajoPred & "'")
                        If Not adrTrabajoPred Is Nothing AndAlso adrTrabajoPred.Length > 0 Then
                            IDTrabajoPred = adrTrabajoPred(0)("IDTrabajo")
                        Else
                            Dim dtTrabajo As DataTable = Trabajo.AddNewForm
                            dtTrabajo.Rows(0)("IDObra") = data.IDObra
                            dtTrabajo.Rows(0)("IDTipoObra") = data.IDTipoObra
                            dtTrabajo.Rows(0)("CodTrabajo") = CodTrabajoPred
                            dtTrabajo.Rows(0)("DescTrabajo") = "MATERIALES"
                            dtTrabajo.Rows(0)("QPrev") = data.Cantidad
                            data.Trabajos.ImportRow(dtTrabajo.Rows(0))
                            IDTrabajoPred = dtTrabajo.Rows(0)("IDTrabajo")
                        End If
                        If Not IDTrabajoPred Is Nothing Then
                            dtMaterial.Rows(0)("IDTrabajo") = IDTrabajoPred
                        End If
                    End If
                    MatPrev.ApplyBusinessRule("PrecioVentaA", dtMaterial.Rows(0)("PrecioPrevMatA"), dtMaterial.Rows(0), data.ContextObra)
                    data.Materiales.ImportRow(dtMaterial.Rows(0))
                End If
            End If
        End If
    End Sub


    <Serializable()> _
    Public Class DataCrearArticulos
        Public Articulos As DataTable
        Public ArticuloEstructura As DataTable
        Public Caracteristicas As DataTable
        Public ArticuloCaracteristicaValor As DataTable
        Public ArticuloDocumento As DataTable

        Public Log As LogProcess

        Public ProductosTerminados As Hashtable
        Public Semielaborados As Hashtable

        Public Sub New(ByVal dtArticulos As DataTable, ByVal dtEstructuras As DataTable, ByVal dtCaracteristicas As DataTable, ByVal dtCaracteristicaValor As DataTable, ByVal dtDocumentos As DataTable)
            Me.Articulos = dtArticulos
            Me.ArticuloEstructura = dtEstructuras
            Me.ProductosTerminados = New Hashtable
            Me.Semielaborados = New Hashtable
            Me.Caracteristicas = dtCaracteristicas
            Me.ArticuloCaracteristicaValor = dtCaracteristicaValor
            Me.ArticuloDocumento = dtDocumentos
        End Sub
    End Class
    <Task()> Public Shared Function CrearArticulosEstructuras(ByVal data As DataCrearArticulos, ByVal services As ServiceProvider) As LogProcess
        Dim dt As DataTable = New BE.DataEngine().Filter("tbMaestroArticulo", New NoRowsFilterItem)
        If data.Log Is Nothing Then data.Log = New LogProcess
        If data.Log.Errors Is Nothing Then ReDim data.Log.Errors(0)
        If data.Log.CreatedElements Is Nothing Then ReDim data.Log.CreatedElements(0)
        ProcessServer.ExecuteTask(Of DataCrearArticulos)(AddressOf CrearArticulos, data, services)
        ProcessServer.ExecuteTask(Of DataCrearArticulos)(AddressOf CrearEstructuras, data, services)
        ProcessServer.ExecuteTask(Of DataCrearArticulos)(AddressOf CrearCararacteristicas, data, services)
        ProcessServer.ExecuteTask(Of DataCrearArticulos)(AddressOf CrearDocumentosArticulos, data, services)
        Return data.Log
    End Function
    <Task()> Public Shared Sub CrearArticulos(ByVal data As DataCrearArticulos, ByVal services As ServiceProvider)
        If data Is Nothing OrElse data.Articulos Is Nothing Then Exit Sub


        Dim TipoProductoTerminado As String = "PFR"
        Dim TipoSemielaborado As String = "PSL"
        Dim TipoMateriaPrima As String = "MPR"

        Dim Fam As New Familia
        Dim FamiliaProductoTerminado As String
        Dim dtFamilia As DataTable = Fam.Filter(New StringFilterItem("IDTipo", TipoProductoTerminado))
        If dtFamilia.Rows.Count > 0 Then
            FamiliaProductoTerminado = dtFamilia.Rows(0)("IDFamilia")
        End If
        Dim FamiliaSemielaborado As String
        dtFamilia = Fam.Filter(New StringFilterItem("IDTipo", TipoSemielaborado))
        If dtFamilia.Rows.Count > 0 Then
            FamiliaSemielaborado = dtFamilia.Rows(0)("IDFamilia")
        End If
        Dim FamiliaMateriaPrima As String
        dtFamilia = Fam.Filter(New StringFilterItem("IDTipo", TipoMateriaPrima))
        If dtFamilia.Rows.Count > 0 Then
            FamiliaMateriaPrima = dtFamilia.Rows(0)("IDFamilia")
        End If

        Dim log As LogProcess = data.Log
        Dim Art As New Articulo
        For Each drArticulo As DataRow In data.Articulos.Rows
            Try

                Dim dtArticulo As DataTable = Art.SelOnPrimaryKey(drArticulo("IDArticulo"))
                If dtArticulo.Rows.Count = 0 Then
                    '//comprobar qué tipo de artículo es
                    If Not data.ArticuloEstructura Is Nothing Then
                        Dim dvArticuloPadre As DataView = New DataView(data.ArticuloEstructura)
                        dvArticuloPadre.RowFilter = "IDArticulo='" & drArticulo("IDArticulo") & "'"
                        If dvArticuloPadre.Count = 0 Then
                            '//MATERIA PRIMA
                            AdminData.BeginTx()
                            Dim dtNewArticulos As DataTable = Art.AddNewForm
                            'Dim drNewArticulo As DataRow = dtNewArticulos.NewRow
                            dtNewArticulos.Rows(0)("IDContador") = System.DBNull.Value
                            dtNewArticulos.Rows(0)("IDArticulo") = drArticulo("IDArticulo")
                            dtNewArticulos.Rows(0)("DescArticulo") = drArticulo("DescArticulo")
                            dtNewArticulos.Rows(0)("IDTipo") = TipoMateriaPrima
                            dtNewArticulos.Rows(0)("IDFamilia") = FamiliaMateriaPrima
                            'dtNewArticulos.Rows.Add(drNewArticulo)
                            Art.Update(dtNewArticulos)
                            AdminData.CommitTx(True)
                            ReDim Preserve log.CreatedElements(log.CreatedElements.Length)
                            log.CreatedElements(log.CreatedElements.Length - 1) = New CreateElement()
                            log.CreatedElements(log.CreatedElements.Length - 1).NElement = drArticulo("IDArticulo")
                        Else
                            Dim dvComponente As DataView = New DataView(data.ArticuloEstructura)
                            dvComponente.RowFilter = "IDComponente='" & drArticulo("IDArticulo") & "'"
                            If dvArticuloPadre.Count = 0 Then
                                '//PRODUCTO TERMINADO
                                AdminData.BeginTx()
                                'Dim dtNewArticulos As DataTable = dtArticulo.Clone
                                Dim dtNewArticulos As DataTable = Art.AddNewForm
                                'Dim drNewArticulo As DataRow = dtNewArticulos.NewRow
                                dtNewArticulos.Rows(0)("IDContador") = System.DBNull.Value
                                dtNewArticulos.Rows(0)("IDArticulo") = drArticulo("IDArticulo")
                                dtNewArticulos.Rows(0)("DescArticulo") = drArticulo("DescArticulo")
                                dtNewArticulos.Rows(0)("IDTipo") = TipoProductoTerminado
                                dtNewArticulos.Rows(0)("IDFamilia") = FamiliaProductoTerminado
                                'dtNewArticulos.Rows.Add(drNewArticulo)
                                Art.Update(dtNewArticulos)
                                data.ProductosTerminados(drArticulo("IDArticulo")) = drArticulo("DescArticulo")
                                AdminData.CommitTx(True)
                                ReDim Preserve log.CreatedElements(log.CreatedElements.Length)
                                log.CreatedElements(log.CreatedElements.Length - 1) = New CreateElement()
                                log.CreatedElements(log.CreatedElements.Length - 1).NElement = drArticulo("IDArticulo")
                            Else
                                '//SEMIELABORADO
                                AdminData.BeginTx()
                                Dim dtNewArticulos As DataTable = Art.AddNewForm
                                'Dim drNewArticulo As DataRow = dtNewArticulos.NewRow
                                dtNewArticulos.Rows(0)("IDContador") = System.DBNull.Value
                                dtNewArticulos.Rows(0)("IDArticulo") = drArticulo("IDArticulo")
                                dtNewArticulos.Rows(0)("DescArticulo") = drArticulo("DescArticulo")
                                dtNewArticulos.Rows(0)("IDTipo") = TipoSemielaborado
                                dtNewArticulos.Rows(0)("IDFamilia") = FamiliaSemielaborado
                                'dtNewArticulos.Rows.Add(drNewArticulo)
                                Art.Update(dtNewArticulos)
                                data.Semielaborados(drArticulo("IDArticulo")) = drArticulo("DescArticulo")
                                AdminData.CommitTx(True)

                                ReDim Preserve log.CreatedElements(log.CreatedElements.Length)
                                log.CreatedElements(log.CreatedElements.Length - 1) = New CreateElement()
                                log.CreatedElements(log.CreatedElements.Length - 1).NElement = drArticulo("IDArticulo")
                            End If
                            dvComponente.RowFilter = Nothing
                        End If
                        dvArticuloPadre.RowFilter = Nothing
                    Else
                        '//MATERIA PRIMA
                        AdminData.BeginTx()
                        Dim dtNewArticulos As DataTable = Art.AddNewForm
                        'Dim drNewArticulo As DataRow = dtNewArticulos.NewRow
                        dtNewArticulos.Rows(0)("IDContador") = System.DBNull.Value
                        dtNewArticulos.Rows(0)("IDArticulo") = drArticulo("IDArticulo")
                        dtNewArticulos.Rows(0)("DescArticulo") = drArticulo("DescArticulo")
                        dtNewArticulos.Rows(0)("IDTipo") = TipoMateriaPrima
                        dtNewArticulos.Rows(0)("IDFamilia") = FamiliaMateriaPrima
                        'dtNewArticulos.Rows.Add(drNewArticulo)
                        Art.Update(dtNewArticulos)
                        AdminData.CommitTx(True)
                        ReDim Preserve log.CreatedElements(log.CreatedElements.Length)
                        log.CreatedElements(log.CreatedElements.Length - 1) = New CreateElement()
                        log.CreatedElements(log.CreatedElements.Length - 1).NElement = drArticulo("IDArticulo")
                    End If
                Else
                    Dim dvArticuloPadre As DataView = New DataView(data.ArticuloEstructura)
                    dvArticuloPadre.RowFilter = "IDArticulo='" & drArticulo("IDArticulo") & "'"
                    If dvArticuloPadre.Count <> 0 Then
                        Dim dvComponente As DataView = New DataView(data.ArticuloEstructura)
                        dvComponente.RowFilter = "IDComponente='" & drArticulo("IDArticulo") & "'"
                        If dvArticuloPadre.Count = 0 Then
                            '//PRODUCTO TERMINADO
                            data.ProductosTerminados(drArticulo("IDArticulo")) = drArticulo("DescArticulo")
                        Else
                            '//SEMIELABORADO
                            data.Semielaborados(drArticulo("IDArticulo")) = drArticulo("DescArticulo")
                        End If
                    End If
                    'ApplicationService.GenerateError("El Artículo {0} ya existe en el sistema.", drArticulo("IDArticulo"))
                End If
            Catch ex As Exception
                Dim Err As New ClassErrors(drArticulo("IDArticulo"), ex.Message)
                ReDim Preserve log.Errors(log.Errors.Length)
                log.Errors(log.Errors.Length - 1) = Err
            End Try
        Next
    End Sub
    <Task()> Public Shared Sub CrearEstructuras(ByVal data As DataCrearArticulos, ByVal services As ServiceProvider)
        If data.ArticuloEstructura Is Nothing Then Exit Sub
        Dim log As LogProcess = data.Log
        If data.Semielaborados.Count > 0 Then
            Dim ArtEstr As New ArticuloEstructura
            Dim Estr As New Estructura
            For Each key As String In data.Semielaborados.Keys
                Try
                    AdminData.BeginTx()
                    '//Creamos registro en  tbArticuloEstructura

                    Dim FilArtEst As New Filter
                    FilArtEst.Add("IDArticulo", FilterOperator.Equal, key)
                    FilArtEst.Add("IDEstructura", FilterOperator.Equal, "PR")
                    Dim dtArticuloEstructura As DataTable = ArtEstr.Filter(FilArtEst)
                    If dtArticuloEstructura Is Nothing OrElse dtArticuloEstructura.Rows.Count = 0 Then
                        dtArticuloEstructura = ArtEstr.AddNewForm
                        dtArticuloEstructura.Rows(0)("IDArticulo") = key
                        dtArticuloEstructura.Rows(0)("IDEstructura") = "PR"
                        dtArticuloEstructura.Rows(0)("DescEstructura") = "PRINCIPAL"
                        dtArticuloEstructura.Rows(0)("Principal") = True
                        ArtEstr.Update(dtArticuloEstructura)
                    End If

                    '//Creamos tbEstructura
                    Dim dtComponentes As DataTable = Estr.AddNew
                    Dim dvComponente As DataView = New DataView(data.ArticuloEstructura)
                    dvComponente.RowFilter = "IDArticulo='" & key & "'"
                    For Each componente As DataRowView In dvComponente
                        Dim FilEstr As New Filter
                        FilEstr.Add("IDArticulo", FilterOperator.Equal, key)
                        FilEstr.Add("IDComponente", FilterOperator.Equal, componente("IDComponente"))
                        FilEstr.Add("IDEstructura", FilterOperator.Equal, dtArticuloEstructura.Rows(0)("IDEstructura"))
                        Dim dtComponente As DataTable = Estr.Filter(FilEstr)
                        If dtComponente Is Nothing OrElse dtComponente.Rows.Count = 0 Then
                            dtComponente = Estr.AddNewForm
                            dtComponente.Rows(0)("IDArticulo") = key
                            dtComponente.Rows(0)("IDEstructura") = dtArticuloEstructura.Rows(0)("IDEstructura")
                            dtComponente.Rows(0)("IDComponente") = componente("IDComponente")
                            dtComponente.Rows(0)("Cantidad") = componente("Cantidad")
                            dtComponente.Rows(0)("Factor") = 1
                            dtComponente.Rows(0)("CantidadProduccion") = 1
                            dtComponentes.ImportRow(dtComponente.Rows(0))
                        End If
                    Next
                    Estr.Update(dtComponentes)
                    dvComponente.RowFilter = Nothing
                    AdminData.CommitTx(True)

                    ReDim Preserve log.CreatedElements(log.CreatedElements.Length)
                    log.CreatedElements(log.CreatedElements.Length - 1) = New CreateElement()
                    log.CreatedElements(log.CreatedElements.Length - 1).ExtraInfo = "Estructura del artículo " & Quoted(key)

                Catch ex As Exception
                    Dim Err As New ClassErrors(key, "Error creando la estructura del artículo " & Quoted(key) & vbNewLine & ex.Message)
                    ReDim Preserve log.Errors(log.Errors.Length)
                    log.Errors(log.Errors.Length - 1) = Err
                End Try

            Next
        End If


        If data.ProductosTerminados.Count > 0 Then
            Dim Estr As New Estructura
            Dim ArtEstr As New ArticuloEstructura
            For Each key As String In data.ProductosTerminados.Keys
                Try

                    AdminData.BeginTx()
                    '//Creamos registro en  tbArticuloEstructura
                    Dim FilArtEst As New Filter
                    FilArtEst.Add("IDArticulo", FilterOperator.Equal, key)
                    FilArtEst.Add("IDEstructura", FilterOperator.Equal, "PR")
                    Dim dtArticuloEstructura As DataTable = ArtEstr.Filter(FilArtEst)
                    If dtArticuloEstructura Is Nothing OrElse dtArticuloEstructura.Rows.Count = 0 Then
                        dtArticuloEstructura = ArtEstr.AddNewForm
                        dtArticuloEstructura.Rows(0)("IDArticulo") = key
                        dtArticuloEstructura.Rows(0)("IDEstructura") = "PR"
                        dtArticuloEstructura.Rows(0)("DescEstructura") = "PRINCIPAL"
                        dtArticuloEstructura.Rows(0)("Principal") = True
                        ArtEstr.Update(dtArticuloEstructura)
                    End If

                    '//Creamos tbEstructura
                    Dim dtComponentes As DataTable = Estr.AddNew
                    Dim dvComponente As DataView = New DataView(data.ArticuloEstructura)
                    dvComponente.RowFilter = "IDArticulo='" & key & "'"
                    For Each componente As DataRowView In dvComponente
                        Dim FilEstr As New Filter
                        FilEstr.Add("IDArticulo", FilterOperator.Equal, key)
                        FilEstr.Add("IDComponente", FilterOperator.Equal, componente("IDComponente"))
                        FilEstr.Add("IDEstructura", FilterOperator.Equal, dtArticuloEstructura.Rows(0)("IDEstructura"))
                        Dim dtComponente As DataTable = Estr.Filter(FilEstr)
                        If dtComponente Is Nothing OrElse dtComponente.Rows.Count = 0 Then
                            dtComponente = Estr.AddNewForm
                            dtComponente.Rows(0)("IDArticulo") = key
                            dtComponente.Rows(0)("IDEstructura") = dtArticuloEstructura.Rows(0)("IDEstructura")
                            dtComponente.Rows(0)("IDComponente") = componente("IDComponente")
                            dtComponente.Rows(0)("Cantidad") = componente("Cantidad")
                            dtComponente.Rows(0)("Factor") = 1
                            dtComponente.Rows(0)("CantidadProduccion") = 1
                            dtComponentes.ImportRow(dtComponente.Rows(0))
                        End If
                    Next
                    Estr.Update(dtComponentes)
                    dvComponente.RowFilter = Nothing
                    AdminData.CommitTx(True)

                    ReDim Preserve log.CreatedElements(log.CreatedElements.Length)
                    log.CreatedElements(log.CreatedElements.Length - 1) = New CreateElement()
                    log.CreatedElements(log.CreatedElements.Length - 1).ExtraInfo = "Estructura del artículo " & Quoted(key)
                Catch ex As Exception
                    Dim Err As New ClassErrors(key, "Error creando la estructura del artículo " & Quoted(key) & vbNewLine & ex.Message)
                    ReDim Preserve log.Errors(log.Errors.Length)
                    log.Errors(log.Errors.Length - 1) = Err
                End Try
            Next
        End If
    End Sub
    <Task()> Public Shared Sub CrearCararacteristicas(ByVal data As DataCrearArticulos, ByVal services As ServiceProvider)
        Dim log As LogProcess = data.Log
        If data.Caracteristicas Is Nothing OrElse data.ArticuloCaracteristicaValor Is Nothing Then Exit Sub
        Dim CaracteristicasExistentes As New Hashtable
        Dim Car As New Caracteristica()
        For Each drCaracteristica As DataRow In data.Caracteristicas.Rows
            Dim IDCaracteritica As String = UCase(Trim(drCaracteristica("IDCaracteristica") & String.Empty))
            Try
                AdminData.BeginTx()
                '//Creamos registro en  tbMaestroCaracteristica
                Dim dtCaracteristica As DataTable = Car.SelOnPrimaryKey(IDCaracteritica)
                If dtCaracteristica.Rows.Count = 0 Then
                    Dim dtNewCaracteristica As DataTable = Car.AddNewForm
                    dtNewCaracteristica.Rows(0)("IDCaracteristica") = IDCaracteritica
                    dtNewCaracteristica.Rows(0)("DescCaracteristica") = IDCaracteritica
                    dtNewCaracteristica.Rows(0)("TipoCaracteristica") = enumTipoCaracteristica.Valor
                    dtNewCaracteristica.Rows(0)("TipoValor") = enumTipoValor.Continuo
                    dtNewCaracteristica.Rows(0)("TipoDato") = enumTipoDato.Alfanumerico

                    Car.Update(dtNewCaracteristica)
                    CaracteristicasExistentes(IDCaracteritica) = IDCaracteritica
                Else
                    CaracteristicasExistentes(IDCaracteritica) = IDCaracteritica
                    ApplicationService.GenerateError("La Característica {0} ya esiste en el sistema.", Quoted(IDCaracteritica))
                End If
                AdminData.CommitTx(True)

                ReDim Preserve log.CreatedElements(log.CreatedElements.Length)
                log.CreatedElements(log.CreatedElements.Length - 1) = New CreateElement()
                log.CreatedElements(log.CreatedElements.Length - 1).NElement = IDCaracteritica
            Catch ex As Exception
                Dim Err As New ClassErrors(IDCaracteritica, "Error creando la Característica " & Quoted(IDCaracteritica) & vbNewLine & ex.Message)
                ReDim Preserve log.Errors(log.Errors.Length)
                log.Errors(log.Errors.Length - 1) = Err
            End Try
        Next

        Dim IDArticuloAnt As String
        For Each drCaracteristica As DataRow In data.ArticuloCaracteristicaValor.Select(Nothing, "IDArticulo")
            Dim IDCaracteristica As String = Trim(drCaracteristica("IDCaracteristica") & String.Empty)
            If IDArticuloAnt <> drCaracteristica("IDArticulo") Then
                If CaracteristicasExistentes.ContainsKey(IDCaracteristica) Then
                    Dim ArtCaract As New ArticuloCaracteristica
                    '//Creamos tbArticuloCaracteristica
                    Dim dvCaracteristica As DataView = New DataView(data.ArticuloCaracteristicaValor)
                    dvCaracteristica.RowFilter = "IDArticulo='" & drCaracteristica("IDArticulo") & "'"
                    For Each caracteristica As DataRowView In dvCaracteristica
                        Dim IDCaracteristicaValor As String = UCase(Trim(caracteristica("IDCaracteristica") & String.Empty))
                        Try
                            AdminData.BeginTx()

                            Dim dtExiste As DataTable = ArtCaract.SelOnPrimaryKey(caracteristica("IDArticulo"), IDCaracteristicaValor)
                            If dtExiste.Rows.Count = 0 Then
                                Dim dtCaracteristica As DataTable = ArtCaract.AddNewForm
                                dtCaracteristica.Rows(0)("IDArticulo") = caracteristica("IDArticulo")
                                dtCaracteristica.Rows(0)("IDCaracteristica") = IDCaracteristicaValor
                                dtCaracteristica.Rows(0)("Valor") = caracteristica("Valor")
                                ArtCaract.Update(dtCaracteristica)
                            Else
                                dtExiste.Rows(0)("Valor") = caracteristica("Valor")
                                BusinessHelper.UpdateTable(dtExiste)
                            End If

                            AdminData.CommitTx(True)
                            ReDim Preserve log.CreatedElements(log.CreatedElements.Length)
                            log.CreatedElements(log.CreatedElements.Length - 1) = New CreateElement()
                            log.CreatedElements(log.CreatedElements.Length - 1).ExtraInfo = "Característica del artículo " & Quoted(drCaracteristica("IDArticulo"))

                        Catch ex As Exception
                            Dim Err As New ClassErrors(IDCaracteristicaValor, "Error creando las características del artículo " & Quoted(drCaracteristica("IDArticulo")) & vbNewLine & ex.Message)
                            ReDim Preserve log.Errors(log.Errors.Length)
                            log.Errors(log.Errors.Length - 1) = Err
                        End Try
                    Next
                    dvCaracteristica.RowFilter = Nothing
                Else
                    'ApplicationService.GenerateError("No se incluye el valor de la característica {0}, ya que no se encuentra en el sistema.", Quoted(IDCaracteristica))
                    Dim Err As New ClassErrors(IDCaracteristica, "No se incluye el valor de la característica " & Quoted(IDCaracteristica) & ", ya que no se encuentra en el sistema.")
                    ReDim Preserve log.Errors(log.Errors.Length)
                    log.Errors(log.Errors.Length - 1) = Err
                End If
                IDArticuloAnt = drCaracteristica("IDArticulo")
            End If
        Next
    End Sub

    <Task()> Public Shared Sub CrearDocumentosArticulos(ByVal data As DataCrearArticulos, ByVal services As ServiceProvider)
        Dim log As LogProcess = data.Log
        If data.ArticuloDocumento Is Nothing Then Exit Sub

        For Each drDoc As DataRow In data.ArticuloDocumento.Rows
            'tbDcmMaestroDocumento
            If Length(drDoc("PathDocumento")) > 0 Then
                Dim IDDocumento As Integer = -1
                Try
                    AdminData.BeginTx()
                    Dim GDDoc As BusinessHelper = BusinessHelper.CreateBusinessObject("DcmMaestroDocumento")
                    Dim dtGDDoc As DataTable = GDDoc.AddNewForm
                    dtGDDoc.Rows(0)("IDDocumento") = AdminData.GetAutoNumeric

                    Dim NombreFichero As String = IO.Path.GetFileName(drDoc("PathDocumento"))
                    dtGDDoc.Rows(0)("DescDocumento") = NombreFichero
                    dtGDDoc.Rows(0)("URL") = drDoc("PathDocumento")
                    dtGDDoc.Rows(0)("URLDestino") = drDoc("PathDocumento")
                    dtGDDoc.Rows(0)("FechaDocumento") = Today
                    dtGDDoc.Rows(0)("Estado") = enumEstadoDoc.Traspasado
                    GDDoc.Update(dtGDDoc)
                    AdminData.CommitTx(True)
                    IDDocumento = dtGDDoc.Rows(0)("IDDocumento")

                    ReDim Preserve log.CreatedElements(log.CreatedElements.Length)
                    log.CreatedElements(log.CreatedElements.Length - 1) = New CreateElement()
                    log.CreatedElements(log.CreatedElements.Length - 1).ExtraInfo = "Documento en GD: " & Quoted(drDoc("PathDocumento"))

                Catch ex As Exception
                    Dim Err As New ClassErrors(drDoc("PathDocumento"), "Error creando Documento en la Gestión Documental " & vbNewLine & ex.Message)
                    ReDim Preserve log.Errors(log.Errors.Length)
                    log.Errors(log.Errors.Length - 1) = Err
                End Try


                If Length(drDoc("IDArticulo")) > 0 AndAlso IDDocumento <> -1 Then
                    'tbDcmDocumentoEntidad
                    Try
                        AdminData.BeginTx()
                        Dim GDDocEnt As BusinessHelper = BusinessHelper.CreateBusinessObject("DcmDocumentoEntidad")
                        Dim dtGDDocEnt As DataTable = GDDocEnt.AddNewForm
                        dtGDDocEnt.Rows(0)("Entidad") = "Articulo"
                        dtGDDocEnt.Rows(0)("IDDocumento") = IDDocumento
                        dtGDDocEnt.Rows(0)("Campo1") = "IDArticulo"
                        dtGDDocEnt.Rows(0)("Valor1") = drDoc("IDArticulo")
                        GDDocEnt.Update(dtGDDocEnt)
                        AdminData.CommitTx(True)

                        ReDim Preserve log.CreatedElements(log.CreatedElements.Length)
                        log.CreatedElements(log.CreatedElements.Length - 1) = New CreateElement()
                        log.CreatedElements(log.CreatedElements.Length - 1).ExtraInfo = "Documento en GD (Relación con Entidad): " & Quoted(drDoc("PathDocumento"))

                    Catch ex As Exception
                        Dim Err As New ClassErrors(drDoc("PathDocumento"), "Error asociando Documento en la Gestión Documental " & vbNewLine & ex.Message)
                        ReDim Preserve log.Errors(log.Errors.Length)
                        log.Errors(log.Errors.Length - 1) = Err
                    End Try

                End If
            End If
        Next

    End Sub

    <Serializable()> _
    Public Class DataEstructuraConfigurado
        Public IDArticuloModelo As String
        Public IDArticuloRaiz As String
        Public Estructura As DataTable
        Public CaracteristicasArticulo As DataTable
        Public DocumentosArticulos As DataTable
    End Class
    <Task()> Public Shared Function GetEstructuraConfigurado(ByVal IDArticulo As String, ByVal services As ServiceProvider) As DataEstructuraConfigurado
        Dim datEstr As New DataEstructuraConfigurado
        datEstr.IDArticuloRaiz = IDArticulo

        ''//comprobamos si tiene documento
        'Dim fDoc As New Filter
        'fDoc.Add(New StringFilterItem("Entidad", "Articulo"))
        'fDoc.Add(New StringFilterItem("Campo1", "IDArticulo"))
        'fDoc.Add(New StringFilterItem("Valor1", IDArticulo))
        'Dim dtDocumentos As DataTable = AdminData.GetData("vCIDocumentos", fDoc)
        'If dtDocumentos.Rows.Count > 0 Then
        '    Dim drNew As DataRow = datEstr.DocumentosArticulos.NewRow
        '    drNew("IDArticulo") = dtDocumentos.Rows(0)("IDArticulo")
        '    drNew("PathDocumento") = dtDocumentos.Rows(0)("URL")
        '    datEstr.DocumentosArticulos.Rows.Add(drNew)
        'End If

        datEstr.Estructura = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf ArticuloEstructura.CalcularEstructuraExplosion, IDArticulo, services)
        datEstr.CaracteristicasArticulo = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf CrearDtCaracteriticasArticulo, Nothing, services)
        If Not datEstr.Estructura Is Nothing Then
            Dim CV As New ArticuloCaracteristica
            For Each drComponente As DataRow In datEstr.Estructura.Rows
                Dim dtCaractValor As DataTable = CV.Filter(New StringFilterItem("IDArticulo", drComponente("IDComponente")))
                If dtCaractValor.Rows.Count > 0 Then
                    For Each drCaracteristica As DataRow In dtCaractValor.Rows
                        Dim adr() As DataRow = datEstr.CaracteristicasArticulo.Select("IDArticulo ='" & drCaracteristica("IDArticulo") & "' AND IDCaracteristica='" & drCaracteristica("IDCaracteristica") & "'")
                        If adr Is Nothing OrElse adr.Length = 0 Then
                            Dim drNew As DataRow = datEstr.CaracteristicasArticulo.NewRow
                            drNew("IDArticulo") = drCaracteristica("IDArticulo")
                            drNew("IDCaracteristica") = drCaracteristica("IDCaracteristica")
                            drNew("Valor") = drCaracteristica("Valor")

                            datEstr.CaracteristicasArticulo.Rows.Add(drNew)
                        End If
                    Next
                End If

                ''//comprobamos si tiene documento
                'fDoc.Clear()
                'fDoc.Add(New StringFilterItem("Entidad", "Articulo"))
                'fDoc.Add(New StringFilterItem("Campo1", "IDArticulo"))
                'fDoc.Add(New StringFilterItem("Valor1", drComponente("IDComponente")))
                'dtDocumentos = AdminData.GetData("vCIDocumentos", fDoc)
                'If dtDocumentos.Rows.Count > 0 Then
                '    Dim drNew As DataRow = datEstr.DocumentosArticulos.NewRow
                '    drNew("IDArticulo") = dtDocumentos.Rows(0)("IDArticulo")
                '    drNew("PathDocumento") = dtDocumentos.Rows(0)("URL")
                '    datEstr.DocumentosArticulos.Rows.Add(drNew)
                'End If
            Next
        End If
        Return datEstr
    End Function
    <Task()> Public Shared Function CrearDtCaracteriticasArticulo(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
        Dim dt As New DataTable
        dt.Columns.Add("IDArticulo", GetType(String))
        dt.Columns.Add("IDCaracteristica", GetType(String))
        dt.Columns.Add("Valor", GetType(String))
        Return dt
    End Function

#End Region

#Region "Codificación Artículo"

#Region "Objetos de Clase"

    <Serializable()> _
    Public Class DataCodifArt
        Public IDTipo As String
        Public IDFamilia As String
        Public IDSubFamilia As String
        Public DtCaract As DataTable
        Public DtArt As DataTable
        Public IDAnada As String
        Public IDCodigo As Integer?
        Public DtDetalle As DataTable
        Public UpdateContador As Boolean = False

        Public Sub New()
        End Sub
        Public Sub New(ByVal IDTipo As String, ByVal IDFamilia As String, ByVal IDSubFamilia As String, ByVal DtCaract As DataTable, _
                       Optional ByVal DtArt As DataTable = Nothing, Optional ByVal IDAnada As String = "")
            Me.IDTipo = IDTipo
            Me.IDFamilia = IDFamilia
            Me.IDSubFamilia = IDSubFamilia
            Me.DtCaract = DtCaract
            Me.IDAnada = IDAnada
            Me.IDCodigo = IDCodigo
            Me.DtArt = DtArt
        End Sub
    End Class

    <Serializable()> _
    Public Class DataCodifData
        Public IDOrigen() As String
        Public Origen As String
        Public IDCodigo As String
        Public DtCodifCab As DataTable
        Public DtCodifDetalle As DataTable

        Public Sub New()
        End Sub
        Public Sub New(ByVal IDOrigen() As String, ByVal Origen As String, ByVal IDCodigo As String)
            Me.IDOrigen = IDOrigen
            Me.Origen = Origen
            Me.IDCodigo = IDCodigo
        End Sub
    End Class

    <Serializable()> _
    Public Class DataCodifReturn
        Public IDArticulo As String
        Public DescArticulo As String

        Public Sub New()
        End Sub
        Public Sub New(ByVal IDArticulo As String, ByVal DescArticulo As String)
            Me.IDArticulo = IDArticulo
            Me.DescArticulo = DescArticulo
        End Sub
    End Class

#End Region

#Region "Tareas"

    <Task()> Public Shared Function CodificarArticulo(ByVal data As DataCodifArt, ByVal services As ServiceProvider) As DataCodifReturn
        If Length(data.IDTipo) > 0 AndAlso Length(data.IDFamilia) > 0 OrElse Length(data.IDCodigo) > 0 Then
            Dim StDataCodif As New DataCodifData
            Dim StDataReturn As New DataCodifReturn
            If Length(data.IDCodigo) > 0 AndAlso (Not data.DtDetalle Is Nothing AndAlso data.DtDetalle.Rows.Count > 0) Then
                StDataCodif.IDCodigo = data.IDCodigo
                StDataCodif.DtCodifDetalle = data.DtDetalle
            End If
            '1º Comprobamos si el artículo tiene SubFamilia y de ser así comprobar si tiene IDCódigo Asociado.
            If Length(StDataCodif.IDCodigo) = 0 AndAlso Length(data.IDSubFamilia) > 0 Then
                Dim StrOr(2) As String
                StrOr(0) = data.IDTipo
                StrOr(1) = data.IDFamilia
                StrOr(2) = data.IDSubFamilia
                StDataCodif.IDOrigen = StrOr
                StDataCodif.Origen = "SubFamilia"
                StDataCodif = ProcessServer.ExecuteTask(Of DataCodifData, DataCodifData)(AddressOf CodifGetCodigo, StDataCodif, services)
            End If
            '2º En caso de no haber encontrado IDCodigo en SubFamilia buscamos en el nivel de la Familia
            If Length(StDataCodif.IDCodigo) = 0 AndAlso Length(data.IDFamilia) > 0 Then
                Dim StrOr(1) As String
                StrOr(0) = data.IDTipo
                StrOr(1) = data.IDFamilia
                StDataCodif.IDOrigen = StrOr
                StDataCodif.Origen = "Familia"
                StDataCodif = ProcessServer.ExecuteTask(Of DataCodifData, DataCodifData)(AddressOf CodifGetCodigo, StDataCodif, services)
            End If
            '3º En caso de no haber encontrado IDCodigo en Familia buscamos en el nivel del Tipo y último
            If Length(StDataCodif.IDCodigo) = 0 AndAlso Length(data.IDTipo) > 0 Then
                Dim StrOr(0) As String
                StrOr(0) = data.IDTipo
                StDataCodif.Origen = "TipoArticulo"
                StDataCodif.IDOrigen = StrOr
                StDataCodif = ProcessServer.ExecuteTask(Of DataCodifData, DataCodifData)(AddressOf CodifGetCodigo, StDataCodif, services)
            End If
            'Comprobamos el Código si se ha encontrado uno y comprobamos si se ha cargado detalle de dicho código
            If Length(StDataCodif.IDCodigo) > 0 Then
                If Not StDataCodif.DtCodifDetalle Is Nothing AndAlso StDataCodif.DtCodifDetalle.Rows.Count > 0 Then
                    'Recorremos la tabla de Detalle por su campo Orden para ir viendo la generación del código
                    For Each DrCodif As DataRow In StDataCodif.DtCodifDetalle.Select("", "Orden ASC")
                        'Comprobamos si se ha rellenado configuración por Artículo-Característica o por Campos y Tablas
                        If Length(DrCodif("IDCampo")) > 0 Then
                            If Length(DrCodif("TablaRelacionada")) > 0 Then
                                Dim ClsEnt As BusinessHelper = BusinessHelper.CreateBusinessObject(DrCodif("TablaRelacionada"))
                                Dim FilTabla As New Filter
                                If DrCodif("TablaRelacionada") = "BdgAnada" AndAlso DrCodif("IDCampo") = "IDAnada" AndAlso Length(data.IDAnada) > 0 Then
                                    FilTabla.Add(DrCodif("IDCampo"), FilterOperator.Equal, data.IDAnada)
                                Else
                                    If Not data.DtArt Is Nothing AndAlso data.DtArt.Rows.Count > 0 Then
                                        Select Case DrCodif("TablaRelacionada")
                                            Case "TipoArticulo"
                                                FilTabla.Add(DrCodif("IDCampo"), FilterOperator.Equal, data.DtArt.Rows(0)(DrCodif("IDCampo")))
                                            Case "Familia"
                                                FilTabla.Add("IDTipo", FilterOperator.Equal, data.IDTipo)
                                                FilTabla.Add(DrCodif("IDCampo"), FilterOperator.Equal, data.DtArt.Rows(0)(DrCodif("IDCampo")))
                                            Case "SubFamilia"
                                                FilTabla.Add("IDTipo", FilterOperator.Equal, data.IDTipo)
                                                FilTabla.Add("IDFamilia", FilterOperator.Equal, data.IDFamilia)
                                                FilTabla.Add(DrCodif("IDCampo"), FilterOperator.Equal, data.DtArt.Rows(0)(DrCodif("IDCampo")))
                                            Case Else
                                                If (DrCodif("TipoCodigo") AndAlso Length(data.DtArt.Rows(0)(DrCodif("IDCampo"))) > 0 AndAlso Length(DrCodif("CampoCodigoRel")) > 0) OrElse _
                                                (Not DrCodif("TipoCodigo") AndAlso Length(data.DtArt.Rows(0)(DrCodif("IDCampo"))) > 0 AndAlso Length(DrCodif("CampoDescRel")) > 0) Then
                                                    FilTabla.Add(DrCodif("IDCampo"), FilterOperator.Equal, data.DtArt.Rows(0)(DrCodif("IDCampo")))
                                                End If
                                        End Select
                                    End If
                                End If
                                Dim DtEnt As New DataTable
                                If FilTabla.Count > 0 Then DtEnt = ClsEnt.Filter(FilTabla)
                                If Not DtEnt Is Nothing AndAlso DtEnt.Rows.Count > 0 Then
                                    If DrCodif("TipoCodigo") Then
                                        If Length(DrCodif("CampoCodigoRel")) > 0 Then
                                            StDataReturn.IDArticulo &= DtEnt.Rows(0)(DrCodif("CampoCodigoRel"))
                                        Else : StDataReturn.IDArticulo &= "??"
                                        End If
                                    End If
                                    If Not DrCodif("TipoCodigo") AndAlso Length(DrCodif("CampoDescRel")) > 0 Then
                                        StDataReturn.DescArticulo &= Strings.Space(1) & DtEnt.Rows(0)(DrCodif("CampoDescRel"))
                                    End If
                                ElseIf DrCodif("TipoCodigo") Then
                                    StDataReturn.IDArticulo &= "??"
                                End If
                            Else
                                If DrCodif("TipoCodigo") Then
                                    If Length(data.DtArt.Rows(0)(DrCodif("IDCampo"))) > 0 Then
                                        StDataReturn.IDArticulo &= data.DtArt.Rows(0)(DrCodif("IDCampo"))
                                    Else : StDataReturn.IDArticulo &= "??"
                                    End If
                                ElseIf Not DrCodif("TipoCodigo") Then
                                    If Length(data.DtArt.Rows(0)(DrCodif("IDCampo"))) > 0 Then
                                        StDataReturn.DescArticulo &= Strings.Space(1) & data.DtArt.Rows(0)(DrCodif("IDCampo"))
                                    End If
                                End If
                            End If
                        ElseIf Length(DrCodif("IDCaracteristica")) > 0 Then
                            Dim DtArtCaract As New DataTable
                            If data.DtCaract Is Nothing OrElse data.DtCaract.Rows.Count = 0 Then
                                Dim FilCaract As New Filter
                                FilCaract.Add("IDArticulo", FilterOperator.Equal, data.DtArt.Rows(0)("IDArticulo"))
                                DtArtCaract = New ArticuloCaracteristica().Filter(FilCaract)
                            Else : DtArtCaract = data.DtCaract
                            End If
                            Dim DrFindCaract() As DataRow = DtArtCaract.Select("IDCaracteristica = '" & DrCodif("IDCaracteristica") & "'")
                            If DrFindCaract.Length > 0 Then
                                Dim DtCaract As DataTable = New Caracteristica().SelOnPrimaryKey(DrCodif("IDCaracteristica"))
                                If Not DtCaract Is Nothing AndAlso DtCaract.Rows.Count > 0 Then
                                    Select Case DtCaract.Rows(0)("TipoValor")
                                        Case enumTipoValor.Continuo
                                            If DrCodif("TipoCodigo") Then
                                                StDataReturn.IDArticulo &= DrFindCaract(0)("Valor")
                                            ElseIf Not DrCodif("TipoCodigo") Then
                                                StDataReturn.DescArticulo &= Strings.Space(1) & DrFindCaract(0)("Valor")
                                            End If
                                        Case enumTipoValor.Discreto
                                            Dim DtCaractValor As DataTable = New CaracteristicaValor().SelOnPrimaryKey(DrCodif("IDCaracteristica"), DrFindCaract(0)("Valor"))
                                            If Not DtCaractValor Is Nothing AndAlso DtCaractValor.Rows.Count > 0 Then
                                                If DrCodif("TipoCodigo") Then
                                                    If Length(DrCodif("IDCaracteristicaCodigo")) > 0 Then
                                                        StDataReturn.IDArticulo &= DtCaractValor.Rows(0)(DrCodif("IDCaracteristicaCodigo"))
                                                    Else : StDataReturn.IDArticulo &= DtCaractValor.Rows(0)("IDValor")
                                                    End If
                                                ElseIf Not DrCodif("TipoCodigo") Then
                                                    If Length(DrCodif("IDCaracteristicaDescCodigo")) > 0 Then
                                                        StDataReturn.DescArticulo &= Strings.Space(1) & DtCaractValor.Rows(0)(DrCodif("IDCaracteristicaDescCodigo"))
                                                    Else : StDataReturn.DescArticulo &= Strings.Space(1) & DtCaractValor.Rows(0)("DescValor")
                                                    End If
                                                End If
                                            ElseIf DrCodif("TipoCodigo") Then
                                                StDataReturn.IDArticulo &= "??"
                                            End If
                                    End Select
                                End If
                            ElseIf DrCodif("TipoCodigo") Then
                                StDataReturn.IDArticulo &= "??"
                            End If
                        End If
                    Next
                    'Por último comprobamos que en la cabecera si se ha configurado un contador para añadir a la parte final.
                    If Not StDataCodif.DtCodifCab Is Nothing AndAlso StDataCodif.DtCodifCab.Rows.Count > 0 Then
                        If Length(StDataCodif.DtCodifCab.Rows(0)("IDContador")) > 0 Then
                            If data.UpdateContador AndAlso Not StDataReturn.IDArticulo.Contains("?") Then
                                StDataReturn.IDArticulo &= ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, StDataCodif.DtCodifCab.Rows(0)("IDContador"), services)
                            Else
                                Dim StCont As Contador.CounterTx = ProcessServer.ExecuteTask(Of String, Contador.CounterTx)(AddressOf Contador.CounterValueTx, StDataCodif.DtCodifCab.Rows(0)("IDContador"), services)
                                StDataReturn.IDArticulo &= StCont.strCounterValue
                            End If
                        End If
                    End If
                    'Comprobamos si la longitud final del artículo supera por 25 la longitud máxima del campo IDArticulo
                    If Length(StDataReturn.IDArticulo) > 25 Then ApplicationService.GenerateError("La longitud del Artículo codificado (|) supera la longitud máxima permitida(25): |.", StDataReturn.IDArticulo.Length, StDataReturn.IDArticulo)
                    Return StDataReturn
                End If
            End If
        Else
            If Length(data.DtArt.Rows(0)("IDTipo")) = 0 Then
                ApplicationService.GenerateError("El Tipo de Artículo es necesario para la codificación.")
            ElseIf Length(data.DtArt.Rows(0)("IDFamilia")) = 0 Then
                ApplicationService.GenerateError("La Familia de Artículo es necesaria para la codificación.")
            End If
        End If
    End Function

    <Task()> Public Shared Function CodifGetCodigo(ByVal data As DataCodifData, ByVal services As ServiceProvider) As DataCodifData
        If Not data.IDOrigen Is Nothing AndAlso data.IDOrigen.Length > 0 Then
            Dim ClsOrigen As BusinessHelper = BusinessHelper.CreateBusinessObject(data.Origen)
            Dim DtOrigen As DataTable = ClsOrigen.SelOnPrimaryKey(data.IDOrigen)
            If Not DtOrigen Is Nothing AndAlso DtOrigen.Rows.Count > 0 Then
                If Length(DtOrigen.Rows(0)("IDCodigo")) > 0 Then
                    data.IDCodigo = DtOrigen.Rows(0)("IDCodigo")
                    If Length(data.IDCodigo) > 0 Then
                        data.DtCodifCab = New CodificacionCabecera().SelOnPrimaryKey(data.IDCodigo)
                        data.DtCodifDetalle = New CodificacionDetalle().Filter(New FilterItem("IDCodigo", FilterOperator.Equal, data.IDCodigo))
                    End If
                End If
            End If
        End If

        Return data
    End Function

    <Serializable()> _
      Public Class DataObtenerArticulosCompatibles
        Public IDArticuloOriginal As String
        Public IDUDMedidadArtOriginal As String
        Public IDLineaPedido As Integer

        Public dtArticulosPrimerNivel As DataTable
        Public dtArticulosSegundoNivel As DataTable

        Public Sub New(ByVal IDArticuloOriginal As String, ByVal IDUDMedidadArtOriginal As String, Optional ByVal IDLineaPedido As Integer = 0)
            Me.IDArticuloOriginal = IDArticuloOriginal
            Me.IDUDMedidadArtOriginal = IDUDMedidadArtOriginal
            Me.IDLineaPedido = IDLineaPedido
        End Sub
    End Class

    <Task()> Public Shared Function ObtenerArticulosCompatibles(ByVal data As DataObtenerArticulosCompatibles, ByVal services As ServiceProvider) As DataObtenerArticulosCompatibles
        Dim lstArticulosPrimerNivel As New List(Of Object)
        Dim lstArticulosSegundoNivel As New List(Of Object)

        Dim ClsArticulo As New Articulo
        Dim DtArt As DataTable = ClsArticulo.SelOnPrimaryKey(data.IDArticuloOriginal)
        If Not DtArt Is Nothing AndAlso DtArt.Rows.Count > 0 Then

            Dim StDataCodif As New Articulo.DataCodifData
            '1º Comprobamos si el artículo tiene SubFamilia y de ser así comprobar si tiene IDCódigo Asociado.
            If Length(DtArt.Rows(0)("IDSubFamilia")) > 0 Then
                Dim StrOr(2) As String
                StrOr(0) = DtArt.Rows(0)("IDTipo")
                StrOr(1) = DtArt.Rows(0)("IDFamilia")
                StrOr(2) = DtArt.Rows(0)("IDSubFamilia")
                StDataCodif.IDOrigen = StrOr
               
                StDataCodif.Origen = "SubFamilia"
                StDataCodif = ProcessServer.ExecuteTask(Of Articulo.DataCodifData, Articulo.DataCodifData)(AddressOf Articulo.CodifGetCodigo, StDataCodif, services)
            End If
            '2º En caso de no haber encontrado IDCodigo en SubFamilia buscamos en el nivel de la Familia
            If Length(StDataCodif.IDCodigo) = 0 AndAlso Length(DtArt.Rows(0)("IDFamilia")) > 0 Then
                Dim StrOr(1) As String
                StrOr(0) = DtArt.Rows(0)("IDTipo")
                StrOr(1) = DtArt.Rows(0)("IDFamilia")

                StDataCodif.IDOrigen = StrOr
                StDataCodif.Origen = "Familia"
                StDataCodif = ProcessServer.ExecuteTask(Of Articulo.DataCodifData, Articulo.DataCodifData)(AddressOf Articulo.CodifGetCodigo, StDataCodif, services)
            End If
            '3º En caso de no haber encontrado IDCodigo en Familia buscamos en el nivel del Tipo y último
            If Length(StDataCodif.IDCodigo) = 0 AndAlso Length(DtArt.Rows(0)("IDTipo")) > 0 Then
                Dim StrOr(0) As String
                StrOr(0) = DtArt.Rows(0)("IDTipo")
                StDataCodif.IDOrigen = StrOr
                StDataCodif.Origen = "TipoArticulo"
                StDataCodif = ProcessServer.ExecuteTask(Of Articulo.DataCodifData, Articulo.DataCodifData)(AddressOf Articulo.CodifGetCodigo, StDataCodif, services)
            End If
            If Length(StDataCodif.IDCodigo) > 0 Then
                If Not StDataCodif.DtCodifDetalle Is Nothing AndAlso StDataCodif.DtCodifDetalle.Rows.Count > 0 Then
                    Dim FilPrincipal As New Filter
                    Dim FilSecundario As New Filter
                    Dim FilCaractPrincipal As New Filter
                    Dim FilCaractSecundario As New Filter
                    For Each DrCodif As DataRow In StDataCodif.DtCodifDetalle.Select("", "Orden ASC")
                        If Length(DrCodif("IDCampo")) > 0 AndAlso DrCodif("TipoCodigo") Then
                            If DrCodif("FiltroPrincipal") Then
                                FilPrincipal.Add(DrCodif("IDCampo"), FilterOperator.Equal, DtArt.Rows(0)(DrCodif("IDCampo")))
                            End If
                            If DrCodif("FiltroSecundario") Then
                                FilSecundario.Add(DrCodif("IDCampo"), FilterOperator.Equal, DtArt.Rows(0)(DrCodif("IDCampo")))
                            End If
                        ElseIf Length(DrCodif("IDCaracteristica")) > 0 AndAlso DrCodif("TipoCodigo") Then
                            Dim FilArtCaract As New Filter
                            FilArtCaract.Add("IDArticulo", FilterOperator.Equal, data.IDArticuloOriginal)
                            FilArtCaract.Add("IDCaracteristica", FilterOperator.Equal, DrCodif("IDCaracteristica"))
                            Dim DtArtCaract As DataTable = New ArticuloCaracteristica().Filter(FilArtCaract)
                            If Not DtArtCaract Is Nothing AndAlso DtArtCaract.Rows.Count > 0 Then
                                If DrCodif("FiltroPrincipal") Then
                                    FilCaractPrincipal.Add("IDCaracteristica", FilterOperator.Equal, DrCodif("IDCaracteristica"))
                                    FilCaractPrincipal.Add("Valor", FilterOperator.Equal, DtArtCaract.Rows(0)("Valor"))
                                End If
                                If DrCodif("FiltroSecundario") Then
                                    FilCaractSecundario.Add("IDCaracteristica", FilterOperator.Equal, DrCodif("IDCaracteristica"))
                                    FilCaractSecundario.Add("Valor", FilterOperator.Equal, DtArtCaract.Rows(0)("Valor"))
                                End If
                            End If
                        End If
                    Next
                    If FilPrincipal.Count > 0 Then
                        FilPrincipal.Add("IDArticulo", FilterOperator.NotEqual, data.IDArticuloOriginal)
                        Dim DtArtPrincipal As DataTable = ClsArticulo.Filter(FilPrincipal, "IDArticulo", "IDArticulo")
                        If Not DtArtPrincipal Is Nothing AndAlso DtArtPrincipal.Rows.Count > 0 Then
                            For Each DrPrin As DataRow In DtArtPrincipal.Select
                                If FilCaractPrincipal.Count > 0 Then
                                    Dim FilCaract As New Filter
                                    FilCaract.Add(FilCaractPrincipal)
                                    FilCaract.Add("IDArticulo", FilterOperator.Equal, DrPrin("IDArticulo"))
                                    Dim DtCaract As DataTable = New ArticuloCaracteristica().Filter(FilCaract)
                                    If Not DtCaract Is Nothing AndAlso DtCaract.Rows.Count > 0 Then
                                        lstArticulosPrimerNivel.Add(DrPrin("IDArticulo"))
                                    End If
                                Else : lstArticulosPrimerNivel.Add(DrPrin("IDArticulo"))
                                End If
                            Next
                        End If
                    End If
                    If FilSecundario.Count > 0 Then
                        FilSecundario.Add("IDArticulo", FilterOperator.NotEqual, data.IDArticuloOriginal)
                        For Each ArtPrin As String In lstArticulosPrimerNivel
                            FilSecundario.Add("IDArticulo", FilterOperator.NotEqual, ArtPrin)
                        Next
                        Dim DtArtSecundario As DataTable = ClsArticulo.Filter(FilSecundario, "IDArticulo", "IDArticulo")
                        If Not DtArtSecundario Is Nothing AndAlso DtArtSecundario.Rows.Count > 0 Then
                            For Each DrSecon As DataRow In DtArtSecundario.Select
                                If FilCaractSecundario.Count > 0 Then
                                    Dim FilCaract As New Filter
                                    FilCaract.Add(FilCaractSecundario)
                                    FilCaract.Add("IDArticulo", FilterOperator.Equal, DrSecon("IDArticulo"))
                                    Dim DtCaract As DataTable = New ArticuloCaracteristica().Filter(FilCaract)
                                    If Not DtCaract Is Nothing AndAlso DtCaract.Rows.Count > 0 Then
                                        lstArticulosSegundoNivel.Add(DrSecon("IDArticulo"))
                                    End If
                                Else : lstArticulosSegundoNivel.Add(DrSecon("IDArticulo"))
                                End If
                            Next
                        End If
                    End If
                Else : ApplicationService.GenerateError("No se ha encontrado detalle de la codificación | para el Artículo.", StDataCodif.IDCodigo)
                End If
            Else : ApplicationService.GenerateError("No se ha encontrado configuración de codificación para el Artículo.")
            End If
        End If

        If Not lstArticulosPrimerNivel Is Nothing AndAlso lstArticulosPrimerNivel.Count > 0 Then
            Dim f As New Filter
            f.Add(New InListFilterItem("IDArticulo", lstArticulosPrimerNivel.ToArray, FilterType.String))
            data.dtArticulosPrimerNivel = New BE.DataEngine().Filter("vfrmBdgArticulosCompatibles", f, , "IDArticulo")
            If data.IDLineaPedido <> 0 AndAlso Not data.dtArticulosPrimerNivel.Columns.Contains("IDLineaPedido") Then
                data.dtArticulosPrimerNivel.Columns.Add("IDLineaPedido", GetType(Integer))
            End If
            If Not data.dtArticulosPrimerNivel Is Nothing AndAlso data.dtArticulosPrimerNivel.Rows.Count > 0 Then
                For Each dr As DataRow In data.dtArticulosPrimerNivel.Rows
                    If data.IDLineaPedido <> 0 Then dr("IDLineaPedido") = data.IDLineaPedido
                    Dim datFactor As New ArticuloUnidadAB.DatosFactorConversion(dr("IDArticulo"), dr("IDUDInterna"), data.IDUDMedidadArtOriginal, True)
                    dr("Factor") = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, datFactor, services)
                Next
            End If
        End If

        If Not lstArticulosSegundoNivel Is Nothing AndAlso lstArticulosSegundoNivel.Count > 0 Then
            Dim f As New Filter
            f.Add(New InListFilterItem("IDArticulo", lstArticulosSegundoNivel.ToArray, FilterType.String))
            data.dtArticulosSegundoNivel = New BE.DataEngine().Filter("vfrmBdgArticulosCompatibles", f, , "IDArticulo")
            If data.IDLineaPedido <> 0 AndAlso Not data.dtArticulosSegundoNivel.Columns.Contains("IDLineaPedido") Then
                data.dtArticulosSegundoNivel.Columns.Add("IDLineaPedido", GetType(Integer))
            End If

            If Not data.dtArticulosSegundoNivel Is Nothing AndAlso data.dtArticulosSegundoNivel.Rows.Count > 0 Then
                For Each dr As DataRow In data.dtArticulosSegundoNivel.Rows
                    If data.IDLineaPedido <> 0 Then dr("IDLineaPedido") = data.IDLineaPedido

                    Dim datFactor As New ArticuloUnidadAB.DatosFactorConversion(dr("IDArticulo"), dr("IDUDInterna"), data.IDUDMedidadArtOriginal, True)
                    dr("Factor") = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, datFactor, services)
                Next
            End If
        End If


        Return data
    End Function

#End Region

#End Region

#Region " Actualizar Precio Estandar "

    <Serializable()> _
    Public Class DataActualizarPrecioEstandar
        Public Articulos As DataTable
        Public Fecha As Date
        Public Sub New(ByVal Articulos As DataTable, ByVal Fecha As Date)
            Me.Articulos = Articulos
            Me.Fecha = Fecha
        End Sub
    End Class
    <Task()> Public Shared Sub ActualizarPrecioEstandar(ByVal data As DataActualizarPrecioEstandar, ByVal services As ServiceProvider)
        If data.Fecha = cnMinDate Then
            ApplicationService.GenerateError("Debe indicar una Fecha de Actualización. Se cancelará el proceso.")
            Exit Sub
        End If
        If Not data.Articulos Is Nothing Then
            For Each drArticulo As DataRow In data.Articulos.Rows
                Dim datActPrecio As New DataActualizarPrecioEstandarArticulo(drArticulo("IDArticulo"), data.Fecha)
                ProcessServer.ExecuteTask(Of DataActualizarPrecioEstandarArticulo)(AddressOf ActualizarPrecioEstandarArticulo, datActPrecio, services)
            Next
        End If
    End Sub

    <Serializable()> _
    Public Class DataActualizarPrecioEstandarArticulo
        Public IDArticulo As String
        Public Fecha As Date
        Public Sub New(ByVal IDArticulo As String, ByVal Fecha As Date)
            Me.IDArticulo = IDArticulo
            Me.Fecha = Fecha
        End Sub
    End Class

    <Task()> Public Shared Sub ActualizarPrecioEstandarArticulo(ByVal data As DataActualizarPrecioEstandarArticulo, ByVal services As ServiceProvider)
        If data.Fecha = cnMinDate Then
            ApplicationService.GenerateError("Debe indicar una Fecha de Actualización. Se cancelará el proceso.")
            Exit Sub
        End If
        If Length(data.IDArticulo) > 0 Then
            '//Llamar a P.A.
            AdminData.Execute("sp_ActualizarPrecioEstandar", False, data.IDArticulo, data.Fecha)
        End If
    End Sub

#End Region

    '21/01/2022 David Velasc
    'Actualiza tipo, fam y subfami por idarticulo
    Public Sub ActualizaTipFamSub(ByVal IDArticulo As String)

        Dim strSQL As String
        strSQL = " UPDATE tbMaestroArticulo"
        strSQL &= " SET IDtipo = ('" & 20 & "'),IDFamilia = ('" & 2001 & "'),IDSubFamilia = ('" & "TORNILLOS" & "')"
        strSQL &= " WHERE IDArticulo = ('" & IDArticulo & "')"

        Try
            AdminData.Execute(strSQL)
        Catch ex As Exception
            ApplicationService.GenerateError(ex.ToString & ": ERROR")
        End Try
        'MsgBox("Modificación realizada con exito en la tabla de ArticuloNSerie")
    End Sub
    Public Sub actualizaPrecioYReposicion(ByVal IDArticulo As String, ByVal ValorA As String, ByVal PrecioA As String)

        'Dim strSQL As String
        'strSQL = " UPDATE tbMaestroArticulo"
        'strSQL &= " SET ValorReposicionA = ('" & ValorA & "'),ValorReposicionB = ('" & ValorA & "'),FechaEstandar = ('" & Today & "'),FechaValorReposicion = ('" & Today & "'), PrecioEstandarA = ('" & PrecioA & "'),PrecioEstandarB = ('" & PrecioA & "')"
        'strSQL &= " WHERE IDArticulo = ('" & IDArticulo & "')"

        Dim strSQL As String
        strSQL = " UPDATE tbMaestroArticulo "
        strSQL &= "SET ValorReposicionA ='" & ValorA & "',ValorReposicionB ='" & ValorA & "',FechaEstandar ='" & Today & "',FechaValorReposicion ='" & Today & "', PrecioEstandarA ='" & PrecioA & "', PrecioEstandarB ='" & PrecioA & "'"
        strSQL &= " WHERE IDArticulo='" & IDArticulo & "'"

        Try
            AdminData.Execute(strSQL)
        Catch ex As Exception
            ApplicationService.GenerateError(ex.ToString & ": ERROR")
        End Try
        'MsgBox("Modificación realizada con exito en la tabla de ArticuloNSerie")
    End Sub

    'Informa del tipo del articulo pasado para ver si es del 30
    Public Function DevuelveTipo(ByVal IDArticulo As String)
        Dim strSQL As String = "SELECT * FROM tbMaestroArticulo WHERE IDArticulo = '" & IDArticulo & "'"
        Dim tb As DataTable = AdminData.GetData(strSQL)
        Dim tipo As String = "0"
        For Each dr As DataRow In tb.Rows
            tipo = dr("IDTipo")
            'MsgBox(stock)
        Next
        Return tipo
    End Function

    'David Velasco-Consulta Movimientos-02/2022
    Public Function DevuelveTabla(ByVal idFamilia As String) As DataTable
        Dim dtFami As DataTable = AdminData.GetData("select DescFamilia from tbMaestroFamilia where IDFamilia='" & idFamilia & "'", False)
        Return dtFami
    End Function
    '-Consulta Movimientos.2ª Parte.
    Public Function DevuelveTabla2(ByVal strSelect As String) As DataTable
        Dim dt As DataTable = AdminData.GetData(strSelect, False)
        Return dt
    End Function

    Public Function DevuelveID() As Integer
        Return AdminData.GetAutoNumeric
    End Function

End Class