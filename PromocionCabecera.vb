Public Class PromocionCabecera

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbPromocionCabecera"

#End Region

#Region "Eventos RegisterAddNewTasks"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim StDatos As New Contador.DatosDefaultCounterValue(data, GetType(PromocionCabecera).Name, "IDPromocion")
        ProcessServer.ExecuteTask(Of Contador.DatosDefaultCounterValue)(AddressOf Contador.LoadDefaultCounterValue, StDatos, services)
    End Sub

#End Region

#Region "Eventos RegisterDeleteTasks"

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf Reordenar)
    End Sub

#End Region

#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("FechaDesde", AddressOf CambioFecha)
        oBrl.Add("FechaHasta", AddressOf CambioFecha)
        oBrl.Add("IDTarifa", AddressOf CambioTarifa)
        oBrl.Add("PromocionGeneral", AddressOf CambioPromoGeneral)
        oBrl.Add("Orden", AddressOf CambioOrden)
        Return oBrl
    End Function

    <Task()> Public Shared Sub CambioFecha(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            data.Current(data.ColumnName) = data.Value
            If Length(data.Current("FechaDesde")) > 0 And Length(data.Current("FechaHasta")) > 0 Then
                If data.Current("FechaDesde") > data.Current("FechaHasta") Then
                    ApplicationService.GenerateError("La Fecha Hasta no puede ser más pequeña que la Fecha Desde.")
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioTarifa(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim ClsTar As New Tarifa
            ClsTar.GetItemRow(data.Value)
        End If
    End Sub

    <Task()> Public Shared Sub CambioPromoGeneral(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then data.Current("Orden") = IIf(data.Value, 0, DBNull.Value)
    End Sub

    <Task()> Public Shared Sub CambioOrden(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            If Not IsNumeric(data.Value) Then
                ApplicationService.GenerateError("Campo no numérico.")
            ElseIf data.Value <= 0 Then
                ApplicationService.GenerateError("El Orden ha de ser un valor numérico mayor de 0.")
            End If
            Dim f As New Filter
            f.Add(New StringFilterItem("IDPromocion", FilterOperator.NotEqual, data.Current("IDPromocion")))
            f.Add(New BooleanFilterItem("PromocionGeneral", True))
            f.Add(New NumberFilterItem("Orden", data.Value))

            Dim dtPromosEnCurso As DataTable = New PromocionCabecera().Filter(f)
            If Not dtPromosEnCurso Is Nothing AndAlso dtPromosEnCurso.Rows.Count > 0 Then
                data.Current("MsgPromocionExiste") = "El Orden indicado ya existe.  Al grabar se reordenarán las Promociones."
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        'validateProcess.AddTask(Of DataRow)(AddressOf ComprobarTarifa)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarDescPromo)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarFechasVigencia)
    End Sub

    '<Task()> Public Shared Sub ComprobarTarifa(ByVal data As DataRow, ByVal services As ServiceProvider)
    '    If Length(data("IDTarifa")) = 0 Then ApplicationService.GenerateError("La Tarifa es un dato obligatorio.")
    'End Sub

    <Task()> Public Shared Sub ComprobarDescPromo(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("DescPromocion")) = 0 Then ApplicationService.GenerateError("La descripción es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ComprobarFechasVigencia(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("FechaDesde")) = 0 Then ApplicationService.GenerateError("La Fecha Desde es un dato obligatorio.")
        If Length(data("FechaHasta")) = 0 Then ApplicationService.GenerateError("La Fecha Hasta es un dato obligatorio.")
        If data("FechaDesde") > data("FechaHasta") Then
            ApplicationService.GenerateError("La Fecha Desde debe ser anterior a la Fecha Hasta.")
        End If
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClave)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarOrden)
        updateProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.UpdateEntityRow)
        updateProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.MarcarComoActualizado)
        updateProcess.AddTask(Of DataRow)(AddressOf ComprobarReorden)
    End Sub

    <Task()> Public Shared Sub AsignarClave(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDContador")) > 0 Then data("IDPromocion") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data("IDContador"), services)
            'Comprobación de la existencia de la Promoción
            Dim dt As DataTable = New PromocionCabecera().SelOnPrimaryKey(data("IDPromocion"))
            If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then ApplicationService.GenerateError("La Promoción ya existe.")
        End If
    End Sub

    <Task()> Public Shared Sub AsignarOrden(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("Orden")) = 0 Then data("Orden") = 0
        End If
    End Sub

    <Task()> Public Shared Sub ComprobarReOrden(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Modified Then
            If data("Orden", DataRowVersion.Original) & String.Empty <> data("Orden") & String.Empty Then
                ProcessServer.ExecuteTask(Of DataRow)(AddressOf Reordenar, data, services)
            End If
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Serializable()> _
    Public Class DatosCompPromo
        Public IDArticulo As String
        Public Cantidad As Double
        Public CantidadAnterior As Double
        Public IDCliente As String
        Public Fecha As Date

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDArticulo As String, ByVal Cantidad As Double, ByVal CantidadAnterior As Double, ByVal IDCliente As String, ByVal Fecha As Date)
            Me.IDArticulo = IDArticulo
            Me.Cantidad = Cantidad
            Me.CantidadAnterior = CantidadAnterior
            Me.IDCliente = IDCliente
            Me.Fecha = Fecha
        End Sub
    End Class

    <Serializable()> _
    Public Class DatosGetPromocion
        Public Dv As DataView
        Public IDArticulo As String
        Public Cantidad As Double
        Public CantidadAnterior As Double
        Public Fecha As Date

        Public Sub New()
        End Sub

        Public Sub New(ByVal Dv As DataView, ByVal IDArticulo As String, ByVal Cantidad As Double, ByVal CantidadAnterior As Double, ByVal Fecha As Date)
            Me.Dv = Dv
            Me.IDArticulo = IDArticulo
            Me.Cantidad = Cantidad
            Me.CantidadAnterior = CantidadAnterior
            Me.Fecha = Fecha
        End Sub
    End Class

    <Serializable()> _
    Public Class DatosVigenciaPromo
        Public IDPromocion As String
        Public Fecha As Date
    End Class

    'Función que devuelve la promocion vigente asignable a un determinado artículo
    <Task()> Public Shared Function ComprobarPromociones(ByVal data As DatosCompPromo, ByVal services As ServiceProvider) As DataTable
        Dim dtResultado As DataTable
        'Primero se comprueba si hay o no gestión de Promociones
        Dim p As New Parametro
        If p.GPromociones() Then
            'Algoritmo para buscar la promoción asignable a un artículo
            '1º - Buscar en las promociones asignadas al CLIENTE siguiendo el ORDEN de Prioridad
            '2º - Buscar alguna "promoción temporal" que afecte a ese artículo (ORDEN)
            '3º - Buscar en la promoción predeterminada

            '1º.-
            Dim f As New Filter
            If Length(data.IDCliente) > 0 Then
                f.Add(New StringFilterItem("IDCliente", data.IDCliente))
                Dim dtClienteProm As DataTable = New BE.DataEngine().Filter("vNegTarifaClientePromocion", f, , "Orden")
                If Not dtClienteProm Is Nothing AndAlso dtClienteProm.Rows.Count > 0 Then
                    'En el caso de encontrar varias promociones para ese cliente buscaremos cual de ellas tiene el artículo en cuestión.
                    'Tienen prioridad los registros que tienen "Orden".
                    Dim dvClienteProm As New DataView(dtClienteProm)
                    dvClienteProm.RowFilter = "Orden <> 0"
                    If dvClienteProm.Count > 0 Then
                        Dim StDatos As New DatosGetPromocion(dvClienteProm, data.IDArticulo, data.Cantidad, data.CantidadAnterior, data.Fecha)
                        dtResultado = ProcessServer.ExecuteTask(Of DatosGetPromocion, DataTable)(AddressOf GetPromocion, StDatos, services)
                    End If
                    'En caso de no encontrar una promoción con prioridad buscamos en las que el campo Orden no tiene valor.
                    If dtResultado Is Nothing OrElse dtResultado.Rows.Count = 0 Then
                        dvClienteProm.RowFilter = "Orden=0"
                        If dvClienteProm.Count > 0 Then
                            Dim StDatos As New DatosGetPromocion(dvClienteProm, data.IDArticulo, data.Cantidad, data.CantidadAnterior, data.Fecha)
                            dtResultado = ProcessServer.ExecuteTask(Of DatosGetPromocion, DataTable)(AddressOf GetPromocion, StDatos, services)
                        End If
                    End If
                End If
            End If
            '2º.-
            Dim dtPromCab As DataTable
            If dtResultado Is Nothing OrElse dtResultado.Rows.Count = 0 Then
                f.Clear()
                f.Add(New BooleanFilterItem("PromocionGeneral", True))
                dtPromCab = New PromocionCabecera().Filter(f, "Orden")
                If Not dtPromCab Is Nothing AndAlso dtPromCab.Rows.Count > 0 Then
                    'En el caso de encontrar varias promociones generales buscaremos cual de ellas tiene el artículo en cuestión.
                    'Tienen prioridad los registros que tienen "Orden".
                    Dim dvPromCab As New DataView(dtPromCab)
                    dvPromCab.RowFilter = "Orden <> 0"
                    If dvPromCab.Count > 0 Then
                        Dim StDatos As New DatosGetPromocion(dvPromCab, data.IDArticulo, data.Cantidad, data.CantidadAnterior, data.Fecha)
                        dtResultado = ProcessServer.ExecuteTask(Of DatosGetPromocion, DataTable)(AddressOf GetPromocion, StDatos, services)
                    End If
                    If dtResultado Is Nothing OrElse dtResultado.Rows.Count = 0 Then
                        dvPromCab.RowFilter = "Orden=0"
                        If dvPromCab.Count > 0 Then
                            Dim StDatos As New DatosGetPromocion(dvPromCab, data.IDArticulo, data.Cantidad, data.CantidadAnterior, data.Fecha)
                            dtResultado = ProcessServer.ExecuteTask(Of DatosGetPromocion, DataTable)(AddressOf GetPromocion, StDatos, services)
                        End If
                    End If
                End If
            End If

            '3º.-
            If dtResultado Is Nothing OrElse dtResultado.Rows.Count = 0 Then
                Dim strIDPromocion As String = p.PromocionPredeterminada()
                If Len(strIDPromocion) <> 0 Then
                    dtPromCab = New PromocionCabecera().SelOnPrimaryKey(strIDPromocion)
                    If Not dtPromCab Is Nothing AndAlso dtPromCab.Rows.Count > 0 Then
                        Dim StDatos As New DatosGetPromocion(dtPromCab.DefaultView, data.IDArticulo, data.Cantidad, data.CantidadAnterior, data.Fecha)
                        dtResultado = ProcessServer.ExecuteTask(Of DatosGetPromocion, DataTable)(AddressOf GetPromocion, StDatos, services)
                    End If
                End If
            End If
        End If
        Return dtResultado
    End Function

    <Task()> Public Shared Function GetPromocion(ByVal data As DatosGetPromocion, ByVal services As ServiceProvider) As DataTable
        Dim dtPromocion As New DataTable
        With dtPromocion
            .Columns.Add("IDPromocionLinea", GetType(Integer))
            .Columns.Add("QMaxPromocionable", GetType(Double))
            .Columns.Add("IDPromocion", GetType(String))
            .Columns.Add("IDTarifa", GetType(String))
        End With

        Dim f As New Filter
        f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))

        Dim strIDPromocion As String
        Dim pl As New PromocionLinea
        For Each drv As DataRowView In data.Dv
            strIDPromocion = drv("IDPromocion")
            'Antes de buscar si el artículo elegido está promocionado en esta promocion, se comprobará la vigencia de la misma.
            Dim StDatosV As New DatosVigenciaPromo
            StDatosV.IDPromocion = strIDPromocion
            StDatosV.Fecha = data.Fecha
            If ProcessServer.ExecuteTask(Of DatosVigenciaPromo, Boolean)(AddressOf VigenciaPromocion, StDatosV, services) Then
                'Se comenta ya que queremos permitir promociones sin tarifa asociada y aplique la de cliente.
                'If Length(drv("IDTarifa")) > 0 Then
                'Llamamos a la función que buscará en la tabla PromocionLinea
                Dim fPromo As New Filter
                fPromo.Add(New StringFilterItem("IDPromocion", strIDPromocion))
                fPromo.Add(f)

                Dim StDatos As New PromocionLinea.DatosFindArtPromo(fPromo, data.Cantidad, data.CantidadAnterior)
                Dim dt As DataTable = ProcessServer.ExecuteTask(Of PromocionLinea.DatosFindArtPromo, DataTable)(AddressOf PromocionLinea.FindArticuloPromocion, StDatos, services)
                If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                    'Promocion encontrada
                    Dim drPromocion As DataRow = dtPromocion.NewRow
                    drPromocion("IDPromocionLinea") = dt.Rows(0)("IDPromocionLinea")
                    drPromocion("QMaxPromocionable") = dt.Rows(0)("QMaxPromocionable")
                    drPromocion("IDPromocion") = strIDPromocion
                    If Length(drv("IDTarifa")) > 0 Then drPromocion("IDTarifa") = drv("IDTarifa")

                    dtPromocion.Rows.Add(drPromocion)
                    Exit For
                End If
                'End If
            End If
        Next
        Return dtPromocion
    End Function

    <Task()> Public Shared Function VigenciaPromocion(ByVal data As DatosVigenciaPromo, ByVal services As ServiceProvider) As Boolean
        Dim blnVigente As Boolean = False
        If Length(data.IDPromocion) Then
            Dim dtPromocion As DataTable = New PromocionCabecera().SelOnPrimaryKey(data.IDPromocion)
            If Not IsNothing(dtPromocion) AndAlso dtPromocion.Rows.Count > 0 Then
                If IsDate(dtPromocion.Rows(0)("FechaDesde")) AndAlso dtPromocion.Rows(0)("FechaDesde") <= data.Fecha Then
                    blnVigente = True
                End If
                If blnVigente AndAlso IsDate(dtPromocion.Rows(0)("FechaHasta")) Then
                    If dtPromocion.Rows(0)("FechaHasta") < data.Fecha Then
                        blnVigente = False
                    End If
                End If
            End If
        End If

        Return blnVigente
    End Function

    <Task()> Public Shared Sub Reordenar(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Not IsNothing(data) Then
            'Funcion que establece el orden de una determinada columna, haciendo que sus valores sean correlativos y que esten ordenados.
            'Cuando se le llama desde el DELETE: Se le pasa el rs con un 0 en la columna que hay que reordenar.
            'La llamada a esta funcion se hace despues de hacer una modificación o un borrado.

            'Hay que seleccionar solo las promociones generales.
            Dim f As New Filter
            f.Add(New StringFilterItem("IDPromocion", FilterOperator.NotEqual, data("IDPromocion")))
            f.Add(New BooleanFilterItem("PromocionGeneral", True))
            f.Add(New NumberFilterItem("Orden", FilterOperator.NotEqual, 0))

            Dim dtPromosEnCurso As DataTable = New PromocionCabecera().Filter(f, "Orden")
            If Not dtPromosEnCurso Is Nothing AndAlso dtPromosEnCurso.Rows.Count > 0 Then
                If Nz(data("Orden"), 0) > dtPromosEnCurso.Rows.Count + 1 Then
                    Dim dr As DataRow = New PromocionCabecera().GetItemRow(data("IDPromocion"))
                    dr("Orden") = dtPromosEnCurso.Rows.Count + 1
                    BusinessHelper.UpdateTable(dr.Table)
                End If

                Dim intOrden As Integer = 1
                For Each drPromosEnCurso As DataRow In dtPromosEnCurso.Rows
                    Dim dr As DataRow = New PromocionCabecera().GetItemRow(drPromosEnCurso("IDPromocion"))
                    If intOrden = data("Orden") Then intOrden += 1
                    dr("Orden") = intOrden
                    BusinessHelper.UpdateTable(dr.Table)
                    intOrden += 1
                Next
            End If
        End If
    End Sub

#End Region

End Class