Public Class PromocionLinea

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbPromocionLinea"

#End Region

#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("IDArticulo", AddressOf CambioArticulo)
        oBrl.Add("QPedida", AddressOf CambioCantidades)
        oBrl.Add("QMinPedido", AddressOf CambioCantidades)
        oBrl.Add("QMaxPedido", AddressOf CambioCantidades)
        oBrl.Add("QMaxPromocionable", AddressOf CambioCantidades)
        oBrl.Add("QPromocionada", AddressOf CambioCantidades)
        Return oBrl
    End Function

    <Task()> Public Shared Sub CambioArticulo(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim dr As DataRow = New Articulo().GetItemRow(data.Value)
            data.Current("DescArticulo") = dr("DescArticulo")
        End If
    End Sub

    <Task()> Public Shared Sub CambioCantidades(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            data.Current(data.ColumnName) = data.Value
            If Not IsNumeric(data.Value) Then
                ApplicationService.GenerateError("Campo no numérico.")
            Else
                If Length(data.Current("QMinPedido")) > 0 AndAlso Length(data.Current("QMaxPedido")) > 0 Then
                    If data.Current("QMinPedido") > data.Current("QMaxPedido") Then
                        ApplicationService.GenerateError("La Cantida Mínima de Pedido no puede ser superior a la Cantidad Máxima.")
                    End If
                End If
                If Length(data.Current("QPromocionada")) > 0 AndAlso Length(data.Current("QMaxPromocionable")) > 0 Then
                    If data.Current("QMaxPromocionable") > 0 AndAlso data.Current("QPromocionada") > data.Current("QMaxPromocionable") Then
                        ApplicationService.GenerateError("La Cantidad Promocionada no puede ser superior a la Cantidad Máxima Promocionable.")
                    End If
                End If
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarArticulo)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarPromocion)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarQMaxPromocionableObligatoria)
    End Sub

    <Task()> Public Shared Sub ComprobarArticulo(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDArticulo")) = 0 Then ApplicationService.GenerateError("El Artículo es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ComprobarPromocion(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim f As New Filter
        f.Add(New StringFilterItem("IDArticulo", data("IDArticulo")))
        f.Add(New StringFilterItem("IDPromocion", data("IDPromocion")))
        If data.RowState = DataRowState.Modified Then f.Add(New NumberFilterItem("IDPromocionLinea", FilterOperator.NotEqual, data("IDPromocionLinea")))
        Dim dt As DataTable = New PromocionLinea().Filter(f)
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then ApplicationService.GenerateError("El Artículo ya está incluido en esta promoción.")
    End Sub

    <Task()> Public Shared Sub ComprobarQMaxPromocionableObligatoria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Nz(data("QMaxPromocionable"), 0) <= 0 Then ApplicationService.GenerateError("Debe indicar una Cantidad Máxima Promocionable.")
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClave)
    End Sub

    <Task()> Public Shared Sub AsignarClave(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDPromocionLinea")) = 0 Then data("IDPromocionLinea") = AdminData.GetAutoNumeric
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Serializable()> _
    Public Class DatosFindArtPromo
        Public FilFind As Filter
        Public Cantidad As Double
        Public CantidadAnterior As Double

        Public Sub New()
        End Sub

        Public Sub New(ByVal FilFind As Filter, ByVal Cantidad As Double, ByVal CantidadAnterior As Double)
            Me.FilFind = FilFind
            Me.Cantidad = Cantidad
            Me.CantidadAnterior = CantidadAnterior
        End Sub
    End Class

    <Serializable()> _
    Public Class DatosActuaLinPromoDr
        Public Doc As DocumentCabLin
        Public Dt As DataTable
        Public Delete As Boolean

        Public Sub New(ByVal Dt As DataTable, Optional ByVal Delete As Boolean = False, Optional ByVal Doc As DocumentCabLin = Nothing)
            Me.Doc = Doc
            Me.Dt = Dt
            Me.Delete = Delete
            Me.Doc = Doc
        End Sub
    End Class

    <Task()> Public Shared Function FindArticuloPromocion(ByVal data As DatosFindArtPromo, ByVal services As ServiceProvider) As DataTable
        Dim dtPromocion As New DataTable
        With dtPromocion
            .Columns.Add("IDPromocionLinea", GetType(Integer))
            .Columns.Add("QMaxPromocionable", GetType(Double))
            .Columns.Add("IDPromocion", GetType(String))
        End With

        Dim blnCantidadOk As Boolean = False

        Dim dtPromoLinea As DataTable = New PromocionLinea().Filter(data.FilFind)
        If Not IsNothing(dtPromoLinea) AndAlso dtPromoLinea.Rows.Count = 1 Then
            'Comprobamos si la cantidad está dentro de los límites (si estos existen).
            If dtPromoLinea.Rows(0)("QMinPedido") = 0 Or dtPromoLinea.Rows(0)("QMinPedido") <= data.Cantidad Then
                blnCantidadOk = True
            End If
            'Existe una promocion que cumple los requisitos
            If blnCantidadOk Then
                Dim drPromocion As DataRow = dtPromocion.NewRow
                drPromocion("IDPromocionLinea") = dtPromoLinea.Rows(0)("IDPromocionLinea")
                'Comprobamos si la QPromocionada no supera la QMaxPromocionable. Solo si QMaxPromocionable <> 0
                If dtPromoLinea.Rows(0)("QMaxPromocionable") <> 0 Then
                    Dim dblQ As Double = data.Cantidad
                    Dim dblQAnterior As Double = data.CantidadAnterior
                    If dtPromoLinea.Rows(0)("QMaxPedido") <> 0 Then
                        If data.Cantidad > dtPromoLinea.Rows(0)("QMaxPedido") Then
                            dblQ = dtPromoLinea.Rows(0)("QMaxPedido")
                        End If
                        If data.CantidadAnterior > dtPromoLinea.Rows(0)("QMaxPedido") Then
                            dblQAnterior = dtPromoLinea.Rows(0)("QMaxPedido")
                        End If
                    End If

                    If dtPromoLinea.Rows(0)("QPromocionada") - dblQAnterior + dblQ <= dtPromoLinea.Rows(0)("QMaxPromocionable") Then
                        drPromocion("QMaxPromocionable") = dtPromoLinea.Rows(0)("QMaxPromocionable")
                    Else
                        drPromocion("QMaxPromocionable") = 0
                    End If
                    drPromocion("IDPromocion") = dtPromoLinea.Rows(0)("IDPromocion")
                End If

                dtPromocion.Rows.Add(drPromocion)
            End If
        End If
        Return dtPromocion
    End Function

    <Task()> Public Shared Sub ActualizarLineaPromocion(ByVal data As DatosActuaLinPromoDr, ByVal services As ServiceProvider)
        'Función que actualiza la Cantidad Promocionada en una Promoción cuando se elimina una linea
        'Tb. elimina aquellas lineas de REGALO generadas por la linea borrada.
        If data.Dt.Rows(0)("Regalo") = 0 AndAlso _
          ((data.Dt.Rows(0).RowState = DataRowState.Added AndAlso Nz(data.Dt.Rows(0)("IDPromocionLinea"), 0) <> 0) OrElse _
            (Nz(data.Dt.Rows(0)("IDPromocionLinea"), 0) <> 0 Or Nz(data.Dt.Rows(0)("IDPromocionLinea", DataRowVersion.Original), 0) <> 0)) Then

            Dim strCampoQ As String = String.Empty
            Dim strCampoIDLinea As String = String.Empty
            Dim strCampoID As String = String.Empty
            Dim StrEntidad As String = String.Empty
            If data.Dt.TableName = GetType(PedidoVentaLinea).Name Then
                strCampoQ = "QPedida"
                strCampoID = "IDPedido" : strCampoIDLinea = "IDLineaPedido"
                StrEntidad = GetType(PedidoVentaLinea).Name
            ElseIf data.Dt.TableName = GetType(AlbaranVentaLinea).Name Then
                strCampoQ = "QServida"
                strCampoID = "IDAlbaran" : strCampoIDLinea = "IDLineaAlbaran"
                StrEntidad = GetType(AlbaranVentaLinea).Name
            ElseIf data.Dt.TableName = GetType(FacturaVentaLinea).Name Then
                strCampoQ = "Cantidad"
                strCampoID = "IDFactura" : strCampoIDLinea = "IDLineaFactura"
                StrEntidad = GetType(FacturaVentaLinea).Name
            End If
            Dim clsEntidad As BusinessHelper = BusinessHelper.CreateBusinessObject(StrEntidad)

            If data.Delete Then
                ProcessServer.ExecuteTask(Of Integer)(AddressOf ActualizarLineaPromocionQ, data.Dt.Rows(0)("IDPromocionLinea", DataRowVersion.Original), services)
            Else
                ProcessServer.ExecuteTask(Of Integer)(AddressOf ActualizarLineaPromocionQ, data.Dt.Rows(0)("IDPromocionLinea"), services)
            End If

            'Borrar las líneas de regalo.
            If data.Delete Then
                Dim lineasEliminadas As Hashtable
                Dim PR As New PromocionRegalo
                Dim F1 As New Filter

                F1.Add(New NumberFilterItem(strCampoID, data.Dt.Rows(0)(strCampoID)))
                F1.Add(New StringFilterItem("IDPromocionLinea", data.Dt.Rows(0)("IDPromocionLinea", DataRowVersion.Original)))
                F1.Add(New BooleanFilterItem("Regalo", True))

                If data.Dt.TableName = GetType(PedidoVentaLinea).Name Then
                    lineasEliminadas = services.GetService(Of LineasPedidoEliminadas).IDLineas
                End If
                If data.Dt.TableName = GetType(AlbaranVentaLinea).Name Then
                    Dim f2 As New Filter(FilterUnionOperator.Or)
                    f2.Add(New IsNullFilterItem("IDLineaPedido"))
                    f2.Add(New NumberFilterItem("IDLineaPedido", 0))
                    F1.Add(f2)
                    lineasEliminadas = services.GetService(Of LineasAlbaranEliminadas).IDLineas
                End If
                If data.Dt.TableName = GetType(FacturaVentaLinea).Name Then
                    Dim f2 As New Filter(FilterUnionOperator.Or)
                    f2.Add(New IsNullFilterItem("IDLineaAlbaran"))
                    f2.Add(New NumberFilterItem("IDLineaAlbaran", 0))
                    F1.Add(f2)
                    lineasEliminadas = services.GetService(Of LineasFacturaEliminadas).IDLineas
                End If

                Dim dtLinea As DataTable
                If data.Doc Is Nothing Then
                    '//Si borramos la línea de promoción, tendremos que borrar las líneas de los regalos.
                    dtLinea = clsEntidad.Filter(F1)
                    'Dim AVL As New AlbaranVentaLinea
                    If Not IsNothing(dtLinea) AndAlso dtLinea.Rows.Count > 0 Then
                        clsEntidad.Delete(dtLinea, services)

                        For Each DrDel As DataRow In dtLinea.Select
                            lineasEliminadas(DrDel(strCampoIDLinea)) = DrDel(strCampoIDLinea)
                        Next
                    End If
                Else
                    '//Si modificamos la cantidad del artículo en promoción, eliminaremos 
                    '//las líneas de regalo para volverlas a crear. Esta es la manera de actualizarlas.
                    Dim dtLineas As DataTable = data.Doc.dtLineas.Clone
                    Dim WhereLineasRegalo As String = F1.Compose(New AdoFilterComposer)
                    For Each dr As DataRow In data.Doc.dtLineas.Select(WhereLineasRegalo)
                        lineasEliminadas(dr(strCampoIDLinea)) = dr(strCampoIDLinea)

                        For Each drAnalitica As DataRow In data.Doc.dtAnalitica.Select(strCampoIDLinea & "=" & dr(strCampoIDLinea))
                            drAnalitica.AcceptChanges() '//Que no haga nada con ella
                        Next

                        Dim Doc As DocumentoComercial = data.Doc
                        If TypeOf data.Doc Is DocumentoAlbaranVenta Then
                            For Each drLote As DataRow In CType(data.Doc, DocumentoAlbaranVenta).dtLote.Select(strCampoIDLinea & "=" & dr(strCampoIDLinea))
                                drLote.AcceptChanges() '//Que no haga nada con ella
                            Next
                        End If


                        For Each drRepresentante As DataRow In Doc.dtVentaRepresentante.Select(strCampoIDLinea & "=" & dr(strCampoIDLinea))
                            drRepresentante.AcceptChanges() '//Que no haga nada con ella
                        Next

                        dr.AcceptChanges() '//Que no haga nada con ella
                        ProcessServer.ExecuteTask(Of DataRow)(AddressOf Comunes.DeleteEntityRow, dr, services)
                        dr.Delete()
                        dr.AcceptChanges()
                    Next
                End If

            End If
        End If
    End Sub
    <Task()> Public Shared Sub ActualizarLineaPromocionQ(ByVal IDPromocionLinea As Integer, ByVal services As ServiceProvider)
        'Función que actualiza la Cantidad Promocionada en una Promoción
        Dim PL As New PromocionLinea
        Dim dtPL As DataTable = PL.SelOnPrimaryKey(IDPromocionLinea)
        If Not dtPL Is Nothing AndAlso dtPL.Rows.Count > 0 Then
            Dim dblQPromocionada, dbQ As Double

            'Líneas de Pedidos
            Dim f1 As New Filter
            f1.Add(New NumberFilterItem("IDPromocionLinea", IDPromocionLinea))
            f1.Add(New NumberFilterItem("Regalo", 0))
            Dim dtPvl As DataTable = New PedidoVentaLinea().Filter(f1)
            For Each dr As DataRow In dtPvl.Rows
                If dr("Estado") = enumpvlEstado.pvlCerrado Then
                    dbQ = dr("QServida")
                Else
                    dbQ = dr("QPedida")
                End If
                If dbQ > dtPL.Rows(0)("QMaxPedido") Then
                    dblQPromocionada = dblQPromocionada + dtPL.Rows(0)("QMaxPedido")
                Else
                    dblQPromocionada = dblQPromocionada + dbQ
                End If
            Next

            'Líneas de Albaranes sin pedido
            Dim f2 As New Filter
            f2.Add(New NumberFilterItem("IDPromocionLinea", IDPromocionLinea))
            f2.Add(New NumberFilterItem("Regalo", 0))
            Dim f3 As New Filter(FilterUnionOperator.Or)
            f3.Add(New IsNullFilterItem("IDLineaPedido"))
            f3.Add(New NumberFilterItem("IDLineaPedido", 0))
            f2.Add(f3)
            Dim dtAvl As DataTable = New AlbaranVentaLinea().Filter(f2)
            For Each dr As DataRow In dtAvl.Rows
                If dr("QServida") > dtPL.Rows(0)("QMaxPedido") Then
                    dblQPromocionada = dblQPromocionada + dtPL.Rows(0)("QMaxPedido")
                Else
                    dblQPromocionada = dblQPromocionada + dr("QServida")
                End If
            Next

            'Líneas de Facturas sin albarán
            Dim f4 As New Filter
            f4.Add(New NumberFilterItem("IDPromocionLinea", IDPromocionLinea))
            f4.Add(New NumberFilterItem("Regalo", 0))
            Dim f5 As New Filter(FilterUnionOperator.Or)
            f5.Add(New IsNullFilterItem("IDLineaAlbaran"))
            f5.Add(New NumberFilterItem("IDLineaAlbaran", 0))
            f4.Add(f5)
            Dim dtFvl As DataTable = New FacturaVentaLinea().Filter(f4)
            For Each dr As DataRow In dtFvl.Rows
                If dr("Cantidad") > dtPL.Rows(0)("QMaxPedido") Then
                    dblQPromocionada = dblQPromocionada + dtPL.Rows(0)("QMaxPedido")
                Else
                    dblQPromocionada = dblQPromocionada + dr("Cantidad")
                End If
            Next

            dtPL.Rows(0)("QPromocionada") = dblQPromocionada

            PL.Update(dtPL)
        End If

    End Sub

    '<Task()> Public Shared Sub TratarPromocion(ByVal dr As DataRow, ByVal services As ServiceProvider)
    '    If dr("Regalo") = 0 Then
    '        Dim strCampoQ, strCampoIDLinea, strCampoID, strCampoArticulo As String
    '        Dim clsEntidad As BusinessHelper
    '        If dr.Table.TableName = GetType(PedidoVentaLinea).Name Then
    '            strCampoQ = "QPedida"
    '            strCampoID = "IDPedido" : strCampoIDLinea = "IDLineaPedido"
    '            strCampoArticulo = "IDArticulo"
    '            clsEntidad = BusinessHelper.CreateBusinessObject(GetType(PedidoVentaLinea).Name)
    '        ElseIf dr.Table.TableName = GetType(AlbaranVentaLinea).Name Then
    '            strCampoQ = "QServida"
    '            strCampoID = "IDAlbaran" : strCampoIDLinea = "IDLineaAlbaran"
    '            strCampoArticulo = "IDArticulo"
    '            clsEntidad = BusinessHelper.CreateBusinessObject(GetType(AlbaranVentaLinea).Name)
    '        ElseIf dr.Table.TableName = GetType(FacturaVentaLinea).Name Then
    '            strCampoQ = "Cantidad"
    '            strCampoID = "IDFactura" : strCampoIDLinea = "IDLineaFactura"
    '            strCampoArticulo = "IDArticulo"
    '            clsEntidad = BusinessHelper.CreateBusinessObject(GetType(FacturaVentaLinea).Name)
    '        End If

    '        If dr.Table.TableName = GetType(AlbaranVentaLinea).Name AndAlso Nz(dr("IDLineaPedido"), 0) <> 0 Then
    '            If dr.RowState = DataRowState.Modified _
    '                                    AndAlso (dr(strCampoQ) <> dr(strCampoQ, DataRowVersion.Original) _
    '                             Or dr(strCampoArticulo) <> dr(strCampoArticulo, DataRowVersion.Original)) Then
    '                ProcessServer.ExecuteTask(Of Integer)(AddressOf ActualizarLineaPromocionQ, dr("IDPromocionLinea"), services)
    '            End If
    '            Exit Sub
    '        End If

    '        If dr.Table.TableName = "FacturaVentaLinea" AndAlso Nz(dr("IDLineaAlbaran"), 0) <> 0 Then
    '            If dr.RowState = DataRowState.Modified _
    '                                    AndAlso (dr(strCampoQ) <> dr(strCampoQ, DataRowVersion.Original) _
    '                             Or dr(strCampoArticulo) <> dr(strCampoArticulo, DataRowVersion.Original)) Then
    '                ProcessServer.ExecuteTask(Of Integer)(AddressOf ActualizarLineaPromocionQ, dr("IDPromocionLinea"), services)
    '            End If
    '            Exit Sub
    '        End If

    '        '10. Quitamos la información anterior
    '        If dr.RowState = DataRowState.Modified Then
    '            If Nz(dr("IDPromocionLinea", DataRowVersion.Original), 0) <> 0 _
    '                    AndAlso (dr(strCampoQ) <> dr(strCampoQ, DataRowVersion.Original) _
    '                             OrElse dr(strCampoArticulo) <> dr(strCampoArticulo, DataRowVersion.Original)) Then
    '                Dim datActLinPromo As New DatosActuaLinPromoDr(dr, True)
    '                ProcessServer.ExecuteTask(Of DatosActuaLinPromoDr)(AddressOf ActualizarLineaPromocion, datActLinPromo, services)
    '            End If
    '        End If

    '        '20. Insertamos la nueva información
    '        If Nz(dr("IDPromocionLinea"), 0) <> 0 Then
    '            If dr.RowState = DataRowState.Added _
    '                    OrElse (dr(strCampoQ) <> dr(strCampoQ, DataRowVersion.Original) _
    '                            Or dr(strCampoArticulo) <> dr(strCampoArticulo, DataRowVersion.Original)) Then

    '                Dim dtPromLinea As DataTable = New PromocionLinea().SelOnPrimaryKey(dr("IDPromocionLinea"))
    '                If Not IsNothing(dtPromLinea) AndAlso dtPromLinea.Rows.Count > 0 Then
    '                    If dr(strCampoQ) >= dtPromLinea.Rows(0)("QMinPedido") Then
    '                        Dim datActLinPromo As New DatosActuaLinPromoDr(dr, False)
    '                        ProcessServer.ExecuteTask(Of DatosActuaLinPromoDr)(AddressOf ActualizarLineaPromocion, datActLinPromo, services)

    '                        If dr.Table.TableName = GetType(PedidoVentaLinea).Name Then
    '                            Dim pvl As New PedidoVentaLinea
    '                            pvl.ADDLineaRegaloPedido(dr, dtPromLinea.Rows(0))
    '                            'ProcesoPedidoVenta.NuevaLineaRegalo()
    '                        ElseIf dr.Table.TableName = GetType(AlbaranVentaLinea).Name Then
    '                            Dim avl As New AlbaranVentaLinea
    '                            avl.ADDLineaRegaloAlbaran(dr, dtPromLinea.Rows(0))
    '                        ElseIf dr.Table.TableName = GetType(FacturaVentaLinea).Name Then
    '                            Dim fvl As New FacturaVentaLinea
    '                            fvl.ADDLineaRegaloFactura(dr, dtPromLinea.Rows(0))
    '                        End If

    '                    End If
    '                End If
    '            End If
    '        End If

    '        '30. En pedidos de venta actualizamos en función de la QServida al cerrar la línea.
    '        If dr.RowState = DataRowState.Modified And dr.Table.TableName = GetType(PedidoVentaLinea).Name Then
    '            If Nz(dr("IDPromocionLinea"), 0) > 0 _
    '                        AndAlso (dr("Estado") = enumpvlEstado.pvlCerrado Or dr("Estado", DataRowVersion.Original) = enumpvlEstado.pvlCerrado) Then
    '                'Si ha cambiado el estado a cerrado hay que comprobar si
    '                'existe alguna diferencia entre la cantidad pedida y la cantidad servida.
    '                If dr("Estado", DataRowVersion.Original) & String.Empty <> dr("Estado") Then
    '                    If dr("QPedida") <> dr("QServida") Then
    '                        'Hay que actualizar la cantidad promocionada
    '                        ProcessServer.ExecuteTask(Of Integer)(AddressOf ActualizarLineaPromocionQ, dr("IDPromocionLinea"), services)
    '                    End If
    '                End If
    '            End If
    '        End If
    '    End If
    'End Sub

#End Region

End Class