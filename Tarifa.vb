Public Class TarifaInfo
    Inherits ClassEntityInfo

    Public IDTarifa As String
    Public DescTarifa As String
    Public IDMoneda As String
    Public IDEstado As String
    Public FechaDesde As Date
    Public FechaHasta As Date
    Public IDTarifaOrigen As String
    Public MaxPrioridad As Boolean
    Public IdContador As String
    Public TarifaPVP As Boolean

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal data As DataRow)
        MyBase.New(data)
    End Sub

    Public Overloads Overrides Sub Fill(ByVal ParamArray PrimaryKey() As Object)
        Dim dt As DataTable
        If Not IsNothing(PrimaryKey) AndAlso PrimaryKey.Length > 0 AndAlso Length(PrimaryKey(0)) > 0 Then
            dt = New Tarifa().SelOnPrimaryKey(PrimaryKey(0))
        End If

        If IsNothing(dt) OrElse dt.Rows.Count = 0 Then
            ApplicationService.GenerateError("La tarifa | no existe.", Quoted(PrimaryKey(0)))
        Else
            Me.Fill(dt.Rows(0))
        End If
    End Sub
End Class

<Serializable()> _
Public Class DatosActualizacionTarifa
    Public IDTarifaOrigen As String
    Public IDTarifaNew As String
    Public DescTarifaNew As String
    Public IDTipo As String
    Public IDFamilia As String
    Public IDSubFamilia As String
    Public IDMoneda As String
    Public IDEstado As String
    Public Dto1 As Double
    Public Dto2 As Double
    Public Dto3 As Double
    Public IncPrecio As Double
    Public AñadirRegATarifa As Boolean
    Public DtNewTarifa As DataTable
    Public TarifaPVP As Boolean
    Public NumDecPrecio As Integer
    Public NumDecImp As Integer
    Public Activo As Boolean

End Class


Public Class Tarifa

#Region "Constructor"
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroTarifa"

#End Region

#Region " RegisterAddNewTasks "

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim StDatos As New Contador.DatosDefaultCounterValue
        StDatos.Row = data
        StDatos.EntityName = GetType(Tarifa).Name
        StDatos.FieldName = "IDTarifa"
        ProcessServer.ExecuteTask(Of Contador.DatosDefaultCounterValue)(AddressOf Contador.LoadDefaultCounterValue, StDatos, services)
        data("IDEstado") = New Parametro().EstadoTarifa
        data("MaxPrioridad") = False
        data("TarifaPVP") = False
    End Sub

#End Region

#Region "Eventos RegisterDeleteTasks"

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf ValidarDelTarifaPredeterminada)
    End Sub

    <Task()> Public Shared Sub ValidarDelTarifaPredeterminada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data("IDTarifa") = New Parametro().TarifaPredeterminada Then
            ApplicationService.GenerateError("El registro no se puede borrar. Es el predeterminado en parámetros.")
        End If
    End Sub

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarMonedaObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDescripcionObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarEstadoObligatorio)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarFechasVigenciaObligatorias)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarTarifaExistente)
    End Sub

    <Task()> Public Shared Sub ValidarDescripcionObligatoria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("DescTarifa")) = 0 Then ApplicationService.GenerateError("La descripción es obligatoria.")
    End Sub

    <Task()> Public Shared Sub ValidarEstadoObligatorio(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDEstado")) = 0 Then ApplicationService.GenerateError("El Estado es obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarFechasVigenciaObligatorias(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("FechaDesde")) = 0 OrElse Length(data("FechaHasta")) = 0 Then
            ApplicationService.GenerateError("Las fechas de Vigencia son Obligatorias.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarTarifaExistente(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IdTarifa")) = 0 Then Exit Sub
            Dim dtTarifa As DataTable = New Tarifa().SelOnPrimaryKey(data("IdTarifa"))
            If Not dtTarifa Is Nothing AndAlso dtTarifa.Rows.Count > 0 Then
                ApplicationService.GenerateError("La Tarifa {0} ya existe en la base de datos.", Quoted(data("IdTarifa")))
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarIdentificadorTarifa)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarPrioridadMaxima)
    End Sub

    <Task()> Public Shared Sub AsignarIdentificadorTarifa(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Not IsDBNull(data("IdContador")) Then
                data("IdTarifa") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data("IDContador"), services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarPrioridadMaxima(ByVal data As DataRow, ByVal services As ServiceProvider)
        'Sólo puede haber una tarifa con PrioridadMaxima para el mismo periodo.
        If Length(data("MaxPrioridad")) > 0 AndAlso data("MaxPrioridad") Then
            Dim F As New Filter
            F.Add(New StringFilterItem("IDTarifa", FilterOperator.NotEqual, data("IdTarifa")))
            If Length(data("FechaDesde")) > 0 Then F.Add(New DateFilterItem("FechaDesde", FilterOperator.GreaterThanOrEqual, data("FechaDesde")))
            If Length(data("FechaHasta")) > 0 Then F.Add(New DateFilterItem("FechaHasta", FilterOperator.LessThanOrEqual, data("FechaHasta")))
            F.Add(New BooleanFilterItem("MaxPrioridad", FilterOperator.Equal, True))

            Dim dtTarifa As DataTable = New Tarifa().Filter(F)
            If Not dtTarifa Is Nothing AndAlso dtTarifa.Rows.Count > 0 Then
                For Each drTarifa As DataRow In dtTarifa.Rows
                    drTarifa("MaxPrioridad") = False
                Next
                BusinessHelper.UpdateTable(dtTarifa)
            End If
        End If
    End Sub

#End Region

#Region "Funciones Publicas"

    <Task()> Public Shared Function NuevaTarifa(ByVal TarInfo As DatosActualizacionTarifa, ByVal services As ServiceProvider) As String
        Dim StrIDTarifaNew As String
        Dim DtNewTarifa As DataTable = TarInfo.DtNewTarifa
        Dim dtTarifa As DataTable = New Tarifa().SelOnPrimaryKey(TarInfo.IDTarifaOrigen)
        If dtTarifa.Rows.Count > 0 Then
            TarInfo.TarifaPVP = dtTarifa.Rows(0)("TarifaPVP")
            TarInfo.IDMoneda = dtTarifa.Rows(0)("IDMoneda")
            TarInfo.IDEstado = dtTarifa.Rows(0)("IDEstado")
            DtNewTarifa.Rows(0)("TarifaPVP") = TarInfo.TarifaPVP
            DtNewTarifa.Rows(0)("IDMoneda") = TarInfo.IDMoneda
            DtNewTarifa.Rows(0)("IDEstado") = TarInfo.IDEstado
            DtNewTarifa.Rows(0)("FechaDesde") = dtTarifa.Rows(0)("FechaDesde")
            DtNewTarifa.Rows(0)("FechaHasta") = dtTarifa.Rows(0)("FechaHasta")
        End If

        TarInfo = ProcessServer.ExecuteTask(Of DatosActualizacionTarifa, DatosActualizacionTarifa)(AddressOf DecimalesMoneda, TarInfo, services)


        If TarInfo.AñadirRegATarifa Then
            StrIDTarifaNew = DtNewTarifa.Rows(0)("IdTarifa")
        Else
            If Length(DtNewTarifa.Rows(0)("IdContador") & String.Empty) > 0 Then
                If Length(DtNewTarifa.Rows(0)("IdTarifa")) = 0 Then
                    DtNewTarifa.Rows(0)("IdTarifa") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, DtNewTarifa.Rows(0)("IDContador"), services)
                Else
                    Dim DefCont As Contador.DefaultCounter = ProcessServer.ExecuteTask(Of String, Contador.DefaultCounter)(AddressOf Contador.GetDefaultCounterValue, "Tarifa", services)
                    If DtNewTarifa.Rows(0)("IdTarifa") = DefCont.CounterValue Then
                        DtNewTarifa.Rows(0)("IdTarifa") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, DtNewTarifa.Rows(0)("IDContador"), services)
                    End If
                End If
            End If
            StrIDTarifaNew = DtNewTarifa.Rows(0)("IdTarifa")
        End If
        BusinessHelper.UpdateTable(DtNewTarifa)

        Dim FilTar As New Filter
        FilTar.Add("IDTarifa", FilterOperator.Equal, TarInfo.IDTarifaOrigen)
        If Len(TarInfo.IDTipo) > 0 Then FilTar.Add("IDTipo", FilterOperator.Equal, TarInfo.IDTipo)
        If Len(TarInfo.IDFamilia) > 0 Then FilTar.Add("IDFamilia", FilterOperator.Equal, TarInfo.IDFamilia)
        If Len(TarInfo.IDSubFamilia) > 0 Then FilTar.Add("IDSubFamilia", FilterOperator.Equal, TarInfo.IDSubFamilia)
        If TarInfo.Activo Then FilTar.Add("Activo", FilterOperator.Equal, True)

        'NUEVA CABECERA DE TARIFA DE ARTICULOS
        Dim StrInsert As String = String.Empty
        StrInsert = "INSERT INTO tbTarifaArticulo "
        StrInsert &= "(IDTarifa, IDArticulo, UdValoracion, Dto1, Dto2, Dto3, FechaUltimaActualizacion, Precio, PVP, Referencia, DescReferencia, FechaCreacionAudi, FechaModificacionAudi)  "
        StrInsert &= "SELECT '" & StrIDTarifaNew & "' AS IDTarifa, IDArticulo, UdValoracion, " & IIf(TarInfo.Dto1 <> -1, CStr(TarInfo.Dto1).Replace(",", "."), "Dto1") & " AS Dto1, " & _
        IIf(TarInfo.Dto2 <> -1, CStr(TarInfo.Dto2).Replace(",", "."), "Dto2") & " AS Dto2, " & IIf(TarInfo.Dto3 <> -1, CStr(TarInfo.Dto3).Replace(",", "."), "Dto3") & " AS Dto3, GETDATE() AS FechaUltimaActualizacion, Precio, PVP, Referencia, DescReferencia, GETDATE() AS FechaCreacionAudi, GETDATE() AS FechaModificacionAudi "
        StrInsert &= "FROM VFrmMntoTarifaTarifaArticulos "
        StrInsert &= "WHERE " & AdminData.ComposeFilter(FilTar)
        AdminData.Execute(StrInsert)

        ProcessServer.ExecuteTask(Of DatosActualizacionTarifa)(AddressOf ModificarArticuloPrecioPVP, TarInfo, services)


        'NUEVAS LINEAS DE TARIFAS ARTICULOS
        StrInsert = "INSERT INTO tbTarifaArticuloLinea "
        StrInsert &= "(IDTarifa, IDArticulo, QDesde, Precio, Dto1, Dto2, Dto3, PVP, Incremento, NuevoPrecio, NuevoPVP, FechaCreacionAudi, FechaModificacionAudi)  "
        StrInsert &= "SELECT '" & StrIDTarifaNew & "' AS IDTarifa, IDArticulo, QDesde, Precio, " & IIf(TarInfo.Dto1 <> -1, CStr(TarInfo.Dto1).Replace(",", "."), "Dto1") & " AS Dto1, " & _
        IIf(TarInfo.Dto2 <> -1, CStr(TarInfo.Dto2).Replace(",", "."), "Dto2") & " AS Dto2, " & IIf(TarInfo.Dto3 <> -1, CStr(TarInfo.Dto3).Replace(",", "."), "Dto3") & " AS Dto3, PVP, Incremento, NuevoPrecio, NuevoPVP, GETDATE() AS FechaCreacionAudi, GETDATE() AS FechaModificacionAudi "
        StrInsert &= "FROM vFrmMntoTarifaArticuloLinea "
        StrInsert &= "WHERE (IDTarifa = N'" & TarInfo.IDTarifaOrigen & "' and IDArticulo in (select IDArticulo from VFrmMntoTarifaTarifaArticulos WHERE " & AdminData.ComposeFilter(FilTar) & "))"
        AdminData.Execute(StrInsert)

        ProcessServer.ExecuteTask(Of DatosActualizacionTarifa)(AddressOf ModificarArticuloLineaPrecioPVP, TarInfo, services)
        Return TarInfo.IDTarifaNew
    End Function
    <Task()> Public Shared Sub ModificarArticuloPrecioPVP(ByVal Tarinfo As DatosActualizacionTarifa, ByVal services As ServiceProvider)
        Dim StrUpdate As String
        Dim FilTar As New Filter

        FilTar.Add("IDTarifa", FilterOperator.Equal, Tarinfo.IDTarifaNew)
        If Len(Tarinfo.IDTipo) > 0 Then FilTar.Add("IDTipo", FilterOperator.Equal, Tarinfo.IDTipo)
        If Len(Tarinfo.IDFamilia) > 0 Then FilTar.Add("IDFamilia", FilterOperator.Equal, Tarinfo.IDFamilia)
        If Len(Tarinfo.IDSubFamilia) > 0 Then FilTar.Add("IDSubFamilia", FilterOperator.Equal, Tarinfo.IDSubFamilia)

        If Tarinfo.IncPrecio <> 0 Then
            If Tarinfo.TarifaPVP Then
                StrUpdate = "UPDATE tbTarifaArticulo "
                StrUpdate &= "SET PVP = ROUND(PVP * (1 + " & CStr(Tarinfo.IncPrecio / 100).Replace(",", ".") & "), " & Tarinfo.NumDecImp & ") "
                StrUpdate &= ", FechaModificacionAudi = GETDATE()"
                StrUpdate &= " FROM tbTarifaArticulo INNER JOIN"
                StrUpdate &= " tbMaestroArticulo ON tbTarifaArticulo.IDArticulo = tbMaestroArticulo.IDArticulo INNER JOIN"
                StrUpdate &= " tbMaestroTipoIva ON tbMaestroArticulo.IDTipoIva = tbMaestroTipoIva.IDTipoIva"
                StrUpdate &= " WHERE " & AdminData.ComposeFilter(FilTar)
                AdminData.Execute(StrUpdate)
                StrUpdate = "Update tbTarifaArticulo "
                StrUpdate &= "SET Precio = round(tbTarifaArticulo.PVP / (1 + tbMaestroTipoIva.Factor / 100)," & Tarinfo.NumDecPrecio & ") "
                StrUpdate &= ", FechaModificacionAudi = GETDATE()"
                StrUpdate &= " FROM tbTarifaArticulo INNER JOIN"
                StrUpdate &= " tbMaestroArticulo ON tbTarifaArticulo.IDArticulo = tbMaestroArticulo.IDArticulo INNER JOIN"
                StrUpdate &= " tbMaestroTipoIva ON tbMaestroArticulo.IDTipoIva = tbMaestroTipoIva.IDTipoIva"
                StrUpdate &= " WHERE " & AdminData.ComposeFilter(FilTar)
                AdminData.Execute(StrUpdate)

            Else
                StrUpdate = "UPDATE tbTarifaArticulo "
                StrUpdate &= "SET Precio = ROUND(Precio * (1 + " & CStr(Tarinfo.IncPrecio / 100).Replace(",", ".") & "), " & Tarinfo.NumDecPrecio & ") "
                StrUpdate &= ", FechaModificacionAudi = GETDATE()"
                StrUpdate &= " FROM   tbTarifaArticulo INNER JOIN"
                StrUpdate &= " tbMaestroArticulo ON tbTarifaArticulo.IDArticulo = tbMaestroArticulo.IDArticulo INNER JOIN"
                StrUpdate &= " tbMaestroTipoIva ON tbMaestroArticulo.IDTipoIva = tbMaestroTipoIva.IDTipoIva"
                StrUpdate &= " WHERE " & AdminData.ComposeFilter(FilTar)
                AdminData.Execute(StrUpdate)
                StrUpdate = "Update tbTarifaArticulo "
                StrUpdate &= "SET PVP = round(tbTarifaArticulo.Precio + tbTarifaArticulo.Precio * tbMaestroTipoIva.Factor / 100 ," & Tarinfo.NumDecImp & ")  "
                StrUpdate &= ", FechaModificacionAudi = GETDATE()"
                StrUpdate &= " FROM   tbTarifaArticulo INNER JOIN"
                StrUpdate &= " tbMaestroArticulo ON tbTarifaArticulo.IDArticulo = tbMaestroArticulo.IDArticulo INNER JOIN"
                StrUpdate &= " tbMaestroTipoIva ON tbMaestroArticulo.IDTipoIva = tbMaestroTipoIva.IDTipoIva"
                StrUpdate &= " WHERE " & AdminData.ComposeFilter(FilTar)
                AdminData.Execute(StrUpdate)
            End If

        End If
    End Sub
    <Task()> Public Shared Sub ModificarArticuloLineaPrecioPVP(ByVal Tarinfo As DatosActualizacionTarifa, ByVal services As ServiceProvider)
        Dim StrUpdate As String
        Dim FilTar As New Filter

        FilTar.Add("IDTarifa", FilterOperator.Equal, Tarinfo.IDTarifaNew)
        If Len(Tarinfo.IDTipo) > 0 Then FilTar.Add("IDTipo", FilterOperator.Equal, Tarinfo.IDTipo)
        If Len(Tarinfo.IDFamilia) > 0 Then FilTar.Add("IDFamilia", FilterOperator.Equal, Tarinfo.IDFamilia)
        If Len(Tarinfo.IDSubFamilia) > 0 Then FilTar.Add("IDSubFamilia", FilterOperator.Equal, Tarinfo.IDSubFamilia)

        If Tarinfo.IncPrecio <> 0 Then
            If Tarinfo.TarifaPVP Then
                StrUpdate = "UPDATE tbTarifaArticuloLinea "
                StrUpdate &= "SET PVP = ROUND(PVP * (1 + " & CStr(Tarinfo.IncPrecio / 100).Replace(",", ".") & "), " & Tarinfo.NumDecImp & ")"
                StrUpdate &= ", FechaModificacionAudi = GETDATE()"
                StrUpdate &= " FROM   tbTarifaArticuloLinea INNER JOIN"
                StrUpdate &= " tbMaestroArticulo ON tbTarifaArticuloLinea.IDArticulo = tbMaestroArticulo.IDArticulo INNER JOIN"
                StrUpdate &= " tbMaestroTipoIva ON tbMaestroArticulo.IDTipoIva = tbMaestroTipoIva.IDTipoIva"
                StrUpdate &= " WHERE " & AdminData.ComposeFilter(FilTar)
                AdminData.Execute(StrUpdate)
                StrUpdate = "Update tbTarifaArticuloLinea "
                StrUpdate &= "SET Precio = round(tbTarifaArticuloLinea.PVP / (1 + tbMaestroTipoIva.Factor / 100)," & Tarinfo.NumDecPrecio & ") "
                StrUpdate &= ", FechaModificacionAudi = GETDATE()"
                StrUpdate &= " FROM   tbTarifaArticuloLinea INNER JOIN"
                StrUpdate &= " tbMaestroArticulo ON tbTarifaArticuloLinea.IDArticulo = tbMaestroArticulo.IDArticulo INNER JOIN"
                StrUpdate &= " tbMaestroTipoIva ON tbMaestroArticulo.IDTipoIva = tbMaestroTipoIva.IDTipoIva"
                StrUpdate &= " WHERE " & AdminData.ComposeFilter(FilTar)
                AdminData.Execute(StrUpdate)

            Else
                StrUpdate = "UPDATE tbTarifaArticuloLinea "
                StrUpdate &= "SET Precio = ROUND(tbTarifaArticuloLinea.Precio * (1 + " & CStr(Tarinfo.IncPrecio / 100).Replace(",", ".") & "), " & Tarinfo.NumDecPrecio & ") "
                StrUpdate &= ", FechaModificacionAudi = GETDATE()"
                StrUpdate &= " FROM   tbTarifaArticuloLinea INNER JOIN"
                StrUpdate &= " tbMaestroArticulo ON tbTarifaArticuloLinea.IDArticulo = tbMaestroArticulo.IDArticulo INNER JOIN"
                StrUpdate &= " tbMaestroTipoIva ON tbMaestroArticulo.IDTipoIva = tbMaestroTipoIva.IDTipoIva"
                StrUpdate &= " WHERE " & AdminData.ComposeFilter(FilTar)
                AdminData.Execute(StrUpdate)
                StrUpdate = "Update tbTarifaArticuloLinea "
                StrUpdate &= "SET PVP = round(tbTarifaArticuloLinea.Precio + tbTarifaArticuloLinea.Precio * tbMaestroTipoIva.Factor / 100 ," & Tarinfo.NumDecImp & ")  "
                StrUpdate &= ", FechaModificacionAudi = GETDATE()"
                StrUpdate &= " FROM   tbTarifaArticuloLinea INNER JOIN"
                StrUpdate &= " tbMaestroArticulo ON tbTarifaArticuloLinea.IDArticulo = tbMaestroArticulo.IDArticulo INNER JOIN"
                StrUpdate &= " tbMaestroTipoIva ON tbMaestroArticulo.IDTipoIva = tbMaestroTipoIva.IDTipoIva"
                StrUpdate &= " WHERE " & AdminData.ComposeFilter(FilTar)
                AdminData.Execute(StrUpdate)
            End If

        End If
    End Sub
    <Task()> Public Shared Function SobreescribirTarifasArticulo(ByVal TarInfo As DatosActualizacionTarifa, ByVal services As ServiceProvider) As String
        Dim StrIDTarifaNew As String
        Dim DtNewTarifa As DataTable = TarInfo.DtNewTarifa

        If Length(DtNewTarifa.Rows(0)("IdContador") & String.Empty) > 0 Then
            If Length(DtNewTarifa.Rows(0)("IdTarifa")) = 0 Then
                DtNewTarifa.Rows(0)("IdTarifa") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, DtNewTarifa.Rows(0)("IDContador"), services)
            Else
                Dim DefCont As Contador.DefaultCounter = ProcessServer.ExecuteTask(Of String, Contador.DefaultCounter)(AddressOf Contador.GetDefaultCounterValue, "Tarifa", services)
                If DtNewTarifa.Rows(0)("IdTarifa") = DefCont.CounterValue Then
                    DtNewTarifa.Rows(0)("IdTarifa") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, DtNewTarifa.Rows(0)("IDContador"), services)
                End If
            End If
        End If
        StrIDTarifaNew = DtNewTarifa.Rows(0)("IdTarifa")

        Dim FilTar As New Filter
        If Len(TarInfo.IDTipo) > 0 Then FilTar.Add("IDTipo", FilterOperator.Equal, TarInfo.IDTipo)
        If Len(TarInfo.IDFamilia) > 0 Then FilTar.Add("IDFamilia", FilterOperator.Equal, TarInfo.IDFamilia)
        If Len(TarInfo.IDSubFamilia) > 0 Then FilTar.Add("IDSubFamilia", FilterOperator.Equal, TarInfo.IDSubFamilia)
        If TarInfo.Activo Then FilTar.Add("Activo", FilterOperator.Equal, True)

        BusinessHelper.UpdateTable(DtNewTarifa)

        Dim StrUpdate As String = String.Empty
        StrUpdate = "INSERT INTO tbTarifaArticulo "
        StrUpdate &= "(IDTarifa, IDArticulo, UdValoracion, Dto1, Dto2, Dto3, Precio, PVP, FechaUltimaActualizacion, FechaCreacionAudi, FechaModificacionAudi) "
        StrUpdate &= "SELECT '" & StrIDTarifaNew & "' AS IDTarifa, IDArticulo, UdValoracion, " & IIf(TarInfo.Dto1 <> -1, Strings.Replace(CStr(TarInfo.Dto1), ",", "."), 0) & " AS Dto1, "
        StrUpdate &= IIf(TarInfo.Dto2 <> -1, Strings.Replace(CStr(TarInfo.Dto2), ",", "."), 0) & " AS Dto2, " & IIf(TarInfo.Dto3 <> -1, Strings.Replace(CStr(TarInfo.Dto3), ",", "."), 0) & " AS Dto3, 0 AS Precio, 0 AS PVP, GETDATE() AS FechaUltimaActualizacion, GETDATE() AS FechaCreacionAudi , GETDATE() AS FechaModificacionAudi "
        StrUpdate &= "FROM tbMaestroArticulo INNER JOIN tbMaestroArticuloEstado ON tbMaestroArticulo.IDEstado = tbMaestroArticuloEstado.IDEstado "
        If FilTar.Count > 0 Then
            StrUpdate &= "WHERE " & AdminData.ComposeFilter(FilTar)
        End If

        AdminData.Execute(StrUpdate, ExecuteCommand.ExecuteScalar)

        Return StrIDTarifaNew
    End Function
    <Task()> Public Shared Sub ModificaTarifa(ByVal TarInfo As DatosActualizacionTarifa, ByVal services As ServiceProvider)
        Dim DtTarModif As DataTable = TarInfo.DtNewTarifa
        Dim dtTarifa As DataTable = New Tarifa().SelOnPrimaryKey(TarInfo.IDTarifaOrigen)
        If dtTarifa.Rows.Count > 0 Then
            TarInfo.TarifaPVP = dtTarifa.Rows(0)("TarifaPVP")
            TarInfo.IDMoneda = dtTarifa.Rows(0)("IDMoneda")
            TarInfo.IDEstado = dtTarifa.Rows(0)("IDEstado")
        End If
        TarInfo.IDTarifaNew = TarInfo.IDTarifaOrigen
        TarInfo = ProcessServer.ExecuteTask(Of DatosActualizacionTarifa, DatosActualizacionTarifa)(AddressOf DecimalesMoneda, TarInfo, services)
        '   BusinessHelper.UpdateTable(DtTarModif)

        Dim FilTar As New Filter
        FilTar.Add("IDTarifa", FilterOperator.Equal, TarInfo.IDTarifaOrigen)
        If Len(TarInfo.IDTipo) > 0 Then FilTar.Add("IDTipo", FilterOperator.Equal, TarInfo.IDTipo)
        If Len(TarInfo.IDFamilia) > 0 Then FilTar.Add("IDFamilia", FilterOperator.Equal, TarInfo.IDFamilia)
        If Len(TarInfo.IDSubFamilia) > 0 Then FilTar.Add("IDSubFamilia", FilterOperator.Equal, TarInfo.IDSubFamilia)
        If TarInfo.Activo Then FilTar.Add("Activo", FilterOperator.Equal, True)

        Dim StrUpdate As String = String.Empty

        StrUpdate = "UPDATE tbTarifaArticulo "
        StrUpdate &= "SET "
        If TarInfo.Dto1 <> -1 Then
            StrUpdate &= "Dto1 = " & CStr(TarInfo.Dto1).Replace(",", ".") & ", "
        End If
        If TarInfo.Dto2 <> -1 Then
            StrUpdate &= "Dto2 = " & CStr(TarInfo.Dto2).Replace(",", ".") & ", "
        End If
        If TarInfo.Dto3 <> -1 Then
            StrUpdate &= "Dto3 = " & CStr(TarInfo.Dto3).Replace(",", ".") & ", "
        End If
        StrUpdate &= "FechaUltimaActualizacion = GETDATE(), "
        StrUpdate &= "FechaModificacionAudi = GETDATE() "
        StrUpdate &= "WHERE IDTarifa = '" & TarInfo.IDTarifaOrigen & "' AND IDArticulo IN (SELECT IDArticulo FROM vFrmMntoTarifaTarifaArticulos WHERE " & AdminData.ComposeFilter(FilTar) & ")"
        AdminData.Execute(StrUpdate)

        ProcessServer.ExecuteTask(Of DatosActualizacionTarifa)(AddressOf ModificarArticuloPrecioPVP, TarInfo, services)

        'Modificacion LINEAS DE TARIFAS ARTICULOS
        If TarInfo.Dto1 <> -1 OrElse TarInfo.Dto2 <> -1 OrElse TarInfo.Dto3 <> -1 Then
            StrUpdate = "UPDATE tbTarifaArticuloLinea "
            StrUpdate &= "SET "
            Dim StrDtos As String = String.Empty
            If TarInfo.Dto1 <> -1 Then
                StrDtos &= "Dto1 = " & CStr(TarInfo.Dto1).Replace(",", ".")
            End If
            If TarInfo.Dto2 <> -1 Then
                If Len(StrDtos) > 0 Then StrDtos &= ", "
                StrDtos &= "Dto2 = " & CStr(TarInfo.Dto2).Replace(",", ".")
            End If
            If TarInfo.Dto3 <> -1 Then
                If Len(StrDtos) > 0 Then StrDtos &= ", "
                StrDtos &= "Dto3 = " & CStr(TarInfo.Dto3).Replace(",", ".")
            End If
            If Len(StrDtos) > 0 Then StrUpdate &= StrDtos & ", "
            StrUpdate &= " FechaModificacionAudi = GETDATE() "
            StrUpdate &= " WHERE IDTarifa = '" & TarInfo.IDTarifaOrigen & "'"
            AdminData.Execute(StrUpdate)
        End If

        ProcessServer.ExecuteTask(Of DatosActualizacionTarifa)(AddressOf ModificarArticuloLineaPrecioPVP, TarInfo, services)



    End Sub
    <Task()> Public Shared Sub ModificaTarifaArticulo(ByVal TarInfo As DatosActualizacionTarifa, ByVal services As ServiceProvider)
        Dim DtTarModif As DataTable = TarInfo.DtNewTarifa
        Dim dtTarifa As DataTable = New Tarifa().SelOnPrimaryKey(TarInfo.IDTarifaOrigen)
        If dtTarifa.Rows.Count > 0 Then
            TarInfo.TarifaPVP = dtTarifa.Rows(0)("TarifaPVP")
            TarInfo.IDMoneda = dtTarifa.Rows(0)("IDMoneda")
            TarInfo.IDEstado = dtTarifa.Rows(0)("IDEstado")
        End If
        TarInfo = ProcessServer.ExecuteTask(Of DatosActualizacionTarifa, DatosActualizacionTarifa)(AddressOf DecimalesMoneda, TarInfo, services)
        '  BusinessHelper.UpdateTable(DtTarModif)

        Dim FilTar As New Filter
        FilTar.Add("IDTarifa", FilterOperator.Equal, TarInfo.IDTarifaOrigen)
        If Len(TarInfo.IDTipo) > 0 Then FilTar.Add("IDTipo", FilterOperator.Equal, TarInfo.IDTipo)
        If Len(TarInfo.IDFamilia) > 0 Then FilTar.Add("IDFamilia", FilterOperator.Equal, TarInfo.IDFamilia)
        If Len(TarInfo.IDSubFamilia) > 0 Then FilTar.Add("IDSubFamilia", FilterOperator.Equal, TarInfo.IDSubFamilia)
        If TarInfo.Activo Then FilTar.Add("Activo", FilterOperator.Equal, True)

        Dim StrDelete As String = String.Empty
        StrDelete = "DELETE FROM tbTarifaArticuloLinea "
        StrDelete &= "WHERE IDTarifa = '" & TarInfo.IDTarifaNew & "' AND IDArticulo IN (SELECT IDArticulo FROM vFrmMntoTarifaTarifaArticulos WHERE " & AdminData.ComposeFilter(FilTar) & ")"
        AdminData.Execute(StrDelete)

        StrDelete = "DELETE FROM tbTarifaArticulo "
        StrDelete &= "WHERE IDTarifa = '" & TarInfo.IDTarifaNew & "' AND IDArticulo IN (SELECT IDArticulo FROM vFrmMntoTarifaTarifaArticulos WHERE " & AdminData.ComposeFilter(FilTar) & ")"
        AdminData.Execute(StrDelete)

        Dim StrInsert As String = String.Empty
        StrInsert = "INSERT INTO tbTarifaArticulo "
        StrInsert &= "(IDTarifa, IDArticulo, UdValoracion, Dto1, Dto2, Dto3, FechaUltimaActualizacion, Precio, PVP, Referencia, DescReferencia, FechaCreacionAudi, FechaModificacionAudi)  "
        StrInsert &= "SELECT '" & TarInfo.IDTarifaNew & "' AS IDTarifa, IDArticulo, UdValoracion, " & IIf(TarInfo.Dto1 <> -1, CStr(TarInfo.Dto1).Replace(",", "."), "Dto1") & " AS Dto1, " & _
        IIf(TarInfo.Dto2 <> -1, CStr(TarInfo.Dto2).Replace(",", "."), "Dto2") & " AS Dto2, " & IIf(TarInfo.Dto3 <> -1, CStr(TarInfo.Dto3).Replace(",", "."), "Dto3") & " AS Dto3, GETDATE() AS FechaUltimaActualizacion, Precio, PVP, Referencia, DescReferencia, GETDATE() AS FechaCreacionAudi , GETDATE() AS FechaModificacionAudi "
        StrInsert &= "FROM VFrmMntoTarifaTarifaArticulos "
        StrInsert &= "WHERE " & AdminData.ComposeFilter(FilTar)
        AdminData.Execute(StrInsert)

        ProcessServer.ExecuteTask(Of DatosActualizacionTarifa)(AddressOf ModificarArticuloPrecioPVP, TarInfo, services)



        'Modificacion LINEAS DE TARIFAS ARTICULOS
        StrInsert = "INSERT INTO tbTarifaArticuloLinea "
        StrInsert &= "(IDTarifa, IDArticulo, QDesde, Precio, Dto1, Dto2, Dto3, PVP, Incremento, NuevoPrecio, NuevoPVP, FechaCreacionAudi, FechaModificacionAudi)  "
        StrInsert &= "SELECT '" & TarInfo.IDTarifaNew & "' AS IDTarifa, IDArticulo, QDesde, Precio, " & IIf(TarInfo.Dto1 <> -1, CStr(TarInfo.Dto1).Replace(",", "."), "Dto1") & " AS Dto1, " & _
        IIf(TarInfo.Dto2 <> -1, CStr(TarInfo.Dto2).Replace(",", "."), "Dto2") & " AS Dto2, " & IIf(TarInfo.Dto3 <> -1, CStr(TarInfo.Dto3).Replace(",", "."), "Dto3") & " AS Dto3, PVP, Incremento, NuevoPrecio, NuevoPVP, GETDATE() AS FechaCreacionAudi, GETDATE() AS FechaModificacionAudi "
        StrInsert &= "FROM tbTarifaArticuloLinea "
        StrInsert &= "WHERE IDTarifa = '" & TarInfo.IDTarifaOrigen & "'"
        AdminData.Execute(StrInsert)
        ProcessServer.ExecuteTask(Of DatosActualizacionTarifa)(AddressOf ModificarArticuloLineaPrecioPVP, TarInfo, services)

    End Sub
    <Task()> Public Shared Sub SobreescribirTarifas(ByVal TarInfo As DatosActualizacionTarifa, ByVal services As ServiceProvider)
        Dim StrDelete As String = String.Empty
        Dim StrInsert As String = String.Empty

        Dim FilTar As New Filter
        FilTar.Add("IDTarifa", FilterOperator.Equal, TarInfo.IDTarifaOrigen)
        If Len(TarInfo.IDTipo) > 0 Then FilTar.Add("IDTipo", FilterOperator.Equal, TarInfo.IDTipo)
        If Len(TarInfo.IDFamilia) > 0 Then FilTar.Add("IDFamilia", FilterOperator.Equal, TarInfo.IDFamilia)
        If Len(TarInfo.IDSubFamilia) > 0 Then FilTar.Add("IDSubFamilia", FilterOperator.Equal, TarInfo.IDSubFamilia)

        StrDelete = "DELETE FROM tbTarifaArticuloLinea "
        StrDelete &= "WHERE IDTarifa = '" & TarInfo.IDTarifaNew & "' AND IDArticulo IN (SELECT IDArticulo FROM vFrmMntoTarifaTarifaArticulos WHERE " & AdminData.ComposeFilter(FilTar) & ")"

        AdminData.Execute(StrDelete, ExecuteCommand.ExecuteNonQuery)

        StrDelete = "DELETE FROM tbTarifaArticulo "
        StrDelete &= "WHERE IDTarifa = '" & TarInfo.IDTarifaNew & "' AND IDArticulo IN (SELECT IDArticulo FROM vFrmMntoTarifaTarifaArticulos WHERE " & AdminData.ComposeFilter(FilTar) & ")"

        AdminData.Execute(StrDelete, ExecuteCommand.ExecuteNonQuery)

        StrInsert = "INSERT INTO tbTarifaArticulo "
        StrInsert &= "(IDTarifa, IDArticulo, UdValoracion, Dto1, Dto2, Dto3, FechaUltimaActualizacion, Precio, PVP, Referencia, DescReferencia, FechaCreacionAudi, FechaModificacionAudi) "
        StrInsert &= "SELECT '" & TarInfo.IDTarifaNew & "' AS IDTarifa, IDArticulo, UdValoracion, " & IIf(TarInfo.Dto1 <> -1, TarInfo.Dto1, "Dto1") & " AS Dto1, " & IIf(TarInfo.Dto2 <> -1, TarInfo.Dto2, "Dto2") & " AS Dto2, " & _
        IIf(TarInfo.Dto3 <> -1, TarInfo.Dto3, "Dto3") & " AS Dto3, FechaUltimaActualizacion, Precio, PVP, Referencia, DescReferencia, GETDATE() AS FechaCreacionAudi , GETDATE() AS FechaModificacionAudi  "
        StrInsert &= "FROM tbTarifaArticulo "
        StrInsert &= "WHERE IDTarifa = '" & TarInfo.IDTarifaOrigen & "' AND IDArticulo IN (SELECT IDArticulo FROM vFrmMntoTarifaTarifaArticulos WHERE " & AdminData.ComposeFilter(FilTar) & ")"

        AdminData.Execute(StrInsert, ExecuteCommand.ExecuteNonQuery)

        StrInsert = "INSERT INTO tbTarifaArticuloLinea "
        StrInsert &= "(IDTarifa, IDArticulo, QDesde, Precio, Dto1, Dto2, Dto3, PVP, Incremento, NuevoPrecio, NuevoPVP, FechaCreacionAudi, FechaModificacionAudi) "
        StrInsert &= "SELECT '" & TarInfo.IDTarifaNew & "' AS IDTarifa, IDArticulo, QDesde, Precio, " & IIf(TarInfo.Dto1 <> -1, TarInfo.Dto1, "Dto1") & " AS Dto1, " & _
        IIf(TarInfo.Dto2 <> -1, TarInfo.Dto2, "Dto2") & " AS Dto2, " & IIf(TarInfo.Dto3 <> -1, TarInfo.Dto3, "Dto3") & " AS Dto3, PVP, Incremento, NuevoPrecio, NuevoPVP, GETDATE() AS FechaCreacionAudi, GETDATE() AS FechaModificacionAudi  "
        StrInsert &= "FROM tbTarifaArticuloLinea "
        StrInsert &= "WHERE IDTarifa = '" & TarInfo.IDTarifaOrigen & "' AND IDArticulo IN (SELECT IDArticulo FROM vFrmMntoTarifaTarifaArticulos WHERE " & AdminData.ComposeFilter(FilTar) & ")"

        AdminData.Execute(StrInsert, ExecuteCommand.ExecuteNonQuery)
    End Sub
    <Task()> Public Shared Sub GenerarTarifaArticulo(ByVal data As DataTarifaArticulo, ByVal services As ServiceProvider)

        ProcessServer.ExecuteTask(Of DataTarifaArticulo)(AddressOf TarifaArticulo.ADDTarifaArticulo, data, services)
    End Sub
    <Task()> Public Shared Function FactorPrecioPVP(ByVal strIDArticulo As String) As Double
        Dim dblFactor As Double
        Dim A As New Articulo

        Dim dtArt As DataTable = A.SelOnPrimaryKey(strIDArticulo)

        If Not IsNothing(dtArt) AndAlso dtArt.Rows.Count > 0 Then
            If Length(dtArt.Rows(0)("IDTipoIva")) > 0 Then
                Dim TI As New TipoIva
                Dim dtIVA As DataTable = TI.SelOnPrimaryKey(dtArt.Rows(0)("IDTipoIva"))
                If Not IsNothing(dtIVA) AndAlso dtIVA.Rows.Count > 0 Then
                    dblFactor = dtIVA.Rows(0)("Factor") / 100
                End If
            End If
        End If

        Return dblFactor
    End Function
    <Task()> Public Shared Function DecimalesMoneda(ByVal Tarifa As DatosActualizacionTarifa, ByVal services As ServiceProvider) As DatosActualizacionTarifa

        'Obtenemos los decimales de la moneda de la tarifa para poder aplicarlos al campo PVP.
        Dim M As New Moneda
        Dim dtMoneda As DataTable = M.SelOnPrimaryKey(Tarifa.IDMoneda)
        If Not IsNothing(dtMoneda) AndAlso dtMoneda.Rows.Count > 0 Then
            Tarifa.NumDecPrecio = dtMoneda.Rows(0)("NDecimalesPrec")
            Tarifa.NumDecImp = dtMoneda.Rows(0)("NDecimalesImp")
        End If

        Return Tarifa
    End Function

    <Serializable()> _
    Public Class DataInsertArtTar
        Public IDTarifaDestino As String
        Public IDTipo As String
        Public IDFamilia As String
        Public IDSubFamilia As String
        Public Activo As Boolean

        Public Sub New(ByVal IDTarifaDestino As String, ByVal IDTipo As String, ByVal IDFamilia As String, ByVal IDSubFamilia As String, ByVal Activo As Boolean)
            Me.IDTarifaDestino = IDTarifaDestino
            Me.IDTipo = IDTipo
            Me.IDFamilia = IDFamilia
            Me.IDSubFamilia = IDSubFamilia
            Me.Activo = Activo
        End Sub
    End Class

    <Task()> Public Shared Sub InsertarArticulosenTarifa(ByVal data As DataInsertArtTar, ByVal services As ServiceProvider)
        Dim FilArt As New Filter
        If Length(data.IDTipo) > 0 Then FilArt.Add("IDTipo", FilterOperator.Equal, data.IDTipo)
        If Length(data.IDFamilia) > 0 Then FilArt.Add("IDFamilia", FilterOperator.Equal, data.IDFamilia)
        If Length(data.IDSubFamilia) > 0 Then FilArt.Add("IDSubFamilia", FilterOperator.Equal, data.IDSubFamilia)
        If data.Activo Then FilArt.Add("Activo", FilterOperator.Equal, True)
        Dim DtArt As DataTable = New BE.DataEngine().Filter("vFrmMntoTarifaTarifaArticulos", FilArt, "IDArticulo", "IDArticulo")
        If Not DtArt Is Nothing AndAlso DtArt.Rows.Count > 0 Then
            Dim ClsArtTar As New TarifaArticulo
            Dim FilArtTar As New Filter
            FilArtTar.Add("IDTarifa", FilterOperator.Equal, data.IDTarifaDestino)
            Dim DtArtTar As DataTable = ClsArtTar.Filter(FilArtTar)
            For Each Dr As DataRow In DtArt.Select
                Dim DrFind() As DataRow = DtArtTar.Select("IDArticulo = '" & Dr("IDArticulo") & "'")
                If DrFind.Length <= 0 Then
                    Dim DrNew As DataRow = DtArtTar.NewRow
                    DrNew("IDTarifa") = data.IDTarifaDestino
                    DrNew("IDArticulo") = Dr("IDArticulo")
                    DrNew("UDValoracion") = 1
                    DrNew("Precio") = 0
                    DrNew("PVP") = 0
                    DrNew("Dto1") = 0
                    DrNew("Dto2") = 0
                    DrNew("Dto3") = 0
                    DtArtTar.Rows.Add(DrNew)
                End If
            Next
            ClsArtTar.Update(DtArtTar)
        End If
    End Sub

    <Serializable()> _
Public Class DatosClienteTarifa
        Public ListClientes As List(Of String)
        Public IDTarifa As String
        Public Order As Integer
        Public dtClienteTarifa As DataTable
        Public ListTarifas As List(Of String)
        Public dtClientes As DataTable

        Public Sub New(ByVal ListClientes As List(Of String), ByVal IDTarifa As String)
            Me.ListClientes = ListClientes
            Me.IDTarifa = IDTarifa
        End Sub

        Public Sub New(ByVal ListClientes As List(Of String), ByVal IDTarifa As String, ByVal dtClientes As DataTable)
            Me.ListClientes = ListClientes
            Me.IDTarifa = IDTarifa
            Me.dtClientes = dtClientes
        End Sub

        Public Sub New(ByVal ListClientes As List(Of String), ByVal IDTarifa As String, ByVal ListTarifas As List(Of String))
            Me.ListClientes = ListClientes
            Me.IDTarifa = IDTarifa
            Me.ListTarifas = ListTarifas
        End Sub
    End Class


    <Task()> Public Shared Function AsignarTarifaCliente(ByVal data As DatosClienteTarifa, ByVal services As ServiceProvider) As Boolean
        Dim asig As Boolean = False
        Dim A As New ClienteTarifa
        Dim existe As Boolean
        If data.ListClientes.Count > 0 Then
            data.dtClienteTarifa = A.AddNew()
            For value As Integer = 0 To data.ListClientes.Count - 1
                Dim drNew As DataRow = data.dtClienteTarifa.NewRow
                drNew("IDCliente") = data.ListClientes(value)
                drNew("IDTarifa") = data.IDTarifa

                Dim StrUpdate As String = String.Empty
                StrUpdate &= "SELECT COALESCE(MAX(Orden), 0) FROM tbClienteTarifa WHERE IDCliente = '" & data.ListClientes(value) & "'"
                data.Order = AdminData.Execute(StrUpdate, ExecuteCommand.ExecuteScalar)
                existe = False
                For Each dr As DataRow In data.dtClientes.Select
                    If dr("IDCliente") = data.ListClientes(value) Then
                        existe = True
                    End If
                Next
                If existe Then
                    data.Order += 1
                End If
                drNew("Orden") = data.Order
                drNew("Predeterminado") = 0
                data.dtClienteTarifa.Rows.Add(drNew)
            Next
            asig = True
        End If
        A.Update(data.dtClienteTarifa)
        Return asig
    End Function


    <Task()> Public Shared Function EliminarTarifaCliente(ByVal data As DatosClienteTarifa, ByVal services As ServiceProvider) As Boolean
        Dim elim As Boolean = False
        Dim A As New ClienteTarifa
        If data.ListClientes.Count > 0 Then
            data.dtClienteTarifa = A.AddNew()
            For value As Integer = 0 To data.ListClientes.Count - 1
                Dim drNew As DataRow = data.dtClienteTarifa.NewRow
                drNew("IDCliente") = data.ListClientes(value)
                drNew("IDTarifa") = data.IDTarifa
                Dim StrUpdate As String = String.Empty
                StrUpdate &= "SELECT Orden FROM tbClienteTarifa WHERE IDCliente = '" & data.ListClientes(value) & "'"
                data.Order = AdminData.Execute(StrUpdate, ExecuteCommand.ExecuteScalar)

                drNew("Orden") = data.Order
                data.dtClienteTarifa.Rows.Add(drNew)
            Next
            elim = True
        End If
        A.Delete(data.dtClienteTarifa)
        Return elim
    End Function

    <Task()> Public Shared Function SustituirTarifaCliente(ByVal data As DatosClienteTarifa, ByVal services As ServiceProvider) As Boolean
        Dim sust As Boolean = False
        Dim A As New ClienteTarifa
        Dim filtroOr As New Filter(FilterUnionOperator.Or)
        Dim dt As DataTable 
        If data.ListClientes.Count > 0 Then
            data.dtClienteTarifa = A.AddNew()
            For value As Integer = 0 To data.ListClientes.Count - 1
                Dim drNew As DataRow = data.dtClienteTarifa.NewRow
                drNew("IDCliente") = data.ListClientes(value)
                drNew("IDTarifa") = data.IDTarifa

                Dim StrUpdate As String = String.Empty
                StrUpdate &= "SELECT Orden FROM tbClienteTarifa WHERE IDCliente = '" & data.ListClientes(value) & "'"
                data.Order = AdminData.Execute(StrUpdate, ExecuteCommand.ExecuteScalar)
                drNew("Orden") = data.Order
                drNew("Predeterminado") = 0
                data.dtClienteTarifa.Rows.Add(drNew)
                Dim filtro As New Filter()
                filtro.Add("IDCliente", data.ListClientes(value))
                filtro.Add("IDTarifa", data.ListTarifas(value))
                filtroOr.Add(filtro)
            Next
            dt = A.Filter(filtroOr)
            sust = True
        End If
        A.Delete(dt)
        A.Update(data.dtClienteTarifa)
        Return sust
    End Function

#End Region

End Class