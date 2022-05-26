Public Class PrevisionLinea

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbPrevisionLinea"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarTipoPrevision)
    End Sub

    <Task()> Public Shared Sub ValidarObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("FechaPrevision")) = 0 Then ApplicationService.GenerateError("La fecha de previsi�n es obligatoria")
        If Length(data("IDArticulo")) > 0 Then
            Dim dtArticulo As DataTable = New Articulo().SelOnPrimaryKey(data("IDArticulo"))
            If dtArticulo Is Nothing OrElse dtArticulo.Rows.Count = 0 Then
                ApplicationService.GenerateError("El Art�culo no existe.")
            End If
        End If

        If Length(data("IDCliente")) > 0 Then
            Dim dtCliente As DataTable = New Cliente().SelOnPrimaryKey(data("IDCliente"))
            If dtCliente Is Nothing OrElse dtCliente.Rows.Count = 0 Then
                ApplicationService.GenerateError("El cliente no existe.")
            End If
        End If

        If Length(data("IDZona")) > 0 Then
            Dim dtZona As DataTable = New Zona().SelOnPrimaryKey(data("IDZona"))
            If dtZona Is Nothing OrElse dtZona.Rows.Count = 0 Then
                ApplicationService.GenerateError("La Zona no existe.")
            End If
        End If

        If Length(data("IDTipo")) > 0 Then
            Dim dtTipoArticulo As DataTable = New TipoArticulo().SelOnPrimaryKey(data("IDTipo"))
            If dtTipoArticulo Is Nothing OrElse dtTipoArticulo.Rows.Count = 0 Then
                ApplicationService.GenerateError("El tipo de art�culo introducido no existe.")
            End If
        End If

        If Length(data("IDFamilia")) > 0 Then
            Dim dtFamilia As DataTable = New Familia().SelOnPrimaryKey(data("IDTipo").ToString(), data("IDFamilia"))
            If dtFamilia Is Nothing OrElse dtFamilia.Rows.Count = 0 Then
                ApplicationService.GenerateError("La familia introducida no existe para el tipo de art�culo actual.")
            End If
        End If

        If Length(data("IDSubFamilia")) > 0 Then
            Dim dtSubfamilia As DataTable = New Subfamilia().SelOnPrimaryKey(data("IDTipo").ToString(), _
                data("IDFamilia").ToString(), data("IDSubFamilia"))
            If dtSubfamilia Is Nothing OrElse dtSubfamilia.Rows.Count = 0 Then
                ApplicationService.GenerateError("La subfamilia introducida no existe para la familia y tipo de art�culo actuales.")
            End If
        End If
    End Sub
    <Task()> Public Shared Sub ValidarTipoPrevision(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim dtDatos As DataTable = New PrevisionCabecera().SelOnPrimaryKey(data("IDPrevision"))
        If dtDatos Is Nothing OrElse dtDatos.Rows.Count = 0 Then ApplicationService.GenerateError("La previsi�n no existe")

        Select Case CInt(dtDatos.Rows(0)("TipoPrevision"))
            Case enumtpTipoPrevision.tpPorArticulo
                If Length(data("IdArticulo")) = 0 Then
                    ApplicationService.GenerateError("El art�culo es obligatorio para el tipo de previsi�n actual.")
                Else
                    data("IDCliente") = System.DBNull.Value
                    data("IDZona") = System.DBNull.Value
                    data("IDTipo") = System.DBNull.Value
                    data("IDFamilia") = System.DBNull.Value
                    data("IDSubFamilia") = System.DBNull.Value
                    data("ImporteA") = 0
                    data("PrecioA") = 0
                End If
            Case enumtpTipoPrevision.tpPorCliente
                If Length(data("IdCliente")) = 0 Then
                    ApplicationService.GenerateError("El cliente es obligatorio para el tipo de previsi�n actual.")
                Else
                    data("IDArticulo") = System.DBNull.Value
                    data("IDZona") = System.DBNull.Value
                    data("IDTipo") = System.DBNull.Value
                    data("IDFamilia") = System.DBNull.Value
                    data("IDSubFamilia") = System.DBNull.Value
                    data("QPrevista") = 0
                    data("PrecioA") = 0
                End If
            Case enumtpTipoPrevision.tpArticuloCliente
                If Length(data("IdArticulo")) = 0 OrElse Length(data("IdCliente")) = 0 Then
                    ApplicationService.GenerateError("El art�culo y el cliente son obligatorios para el tipo de previsi�n actual.")
                Else
                    data("IDZona") = System.DBNull.Value
                    data("IDTipo") = System.DBNull.Value
                    data("IDFamilia") = System.DBNull.Value
                    data("IDSubFamilia") = System.DBNull.Value
                    data("PrecioA") = 0
                End If
            Case enumtpTipoPrevision.tpZona
                If Length(data("IdZona")) = 0 Then
                    ApplicationService.GenerateError("La zona es obligatoria para el tipo de previsi�n actual.")
                Else
                    data("IDArticulo") = System.DBNull.Value
                    data("IDCliente") = System.DBNull.Value
                    data("IDTipo") = System.DBNull.Value
                    data("IDFamilia") = System.DBNull.Value
                    data("IDSubFamilia") = System.DBNull.Value
                    data("QPrevista") = 0
                    data("PrecioA") = 0
                End If
            Case enumtpTipoPrevision.tpClienteTipoFamiliaSubFamilia
                If Length(data("IDCliente")) = 0 OrElse Length(data("IDTipo")) = 0 Then
                    ApplicationService.GenerateError("El cliente y el tipo de art�culo son obligatorios para el tipo de previsi�n actual.")
                Else
                    data("IDArticulo") = System.DBNull.Value
                    data("IDZona") = System.DBNull.Value
                End If
        End Select
    End Sub
    '<Task()> Public Shared Sub ValidarTipoPrevision(ByVal data As DataRow, ByVal services As ServiceProvider)
    '    Dim Fil As New Filter
    '    Fil.Add("IDPrevision", FilterOperator.Equal, data("IDPrevision"))
    '    Dim IDTipoPrevision As Integer


    '    Dim DtDatos As DataTable
    '    Select Case data("IDTipoPrevision")
    '        Case enumtpTipoPrevision.tpPorArticulo
    '            Fil.Add("IDArticulo", FilterOperator.Equal, data("IDArticulo"))
    '            DtDatos = New PrevisionLinea().Filter(Fil)
    '            If Not dtDatos Is Nothing AndAlso dtDatos.Rows.Count > 0 Then
    '                ApplicationService.GenerateError("Ya existe una l�nea de previsi�n para ese art�culo y esa fecha")
    '            End If
    '        Case enumtpTipoPrevision.tpPorCliente
    '            Fil.Add("IDCliente", FilterOperator.Equal, data("IDCliente"))
    '            DtDatos = New PrevisionLinea().Filter(Fil)
    '            If Not dtDatos Is Nothing AndAlso dtDatos.Rows.Count > 0 Then
    '                ApplicationService.GenerateError("Ya existe una l�nea de previsi�n para ese cliente y esa fecha")
    '            End If
    '        Case enumtpTipoPrevision.tpArticuloCliente
    '            Fil.Add("IDArticulo", FilterOperator.Equal, data("IDArticulo"))
    '            Fil.Add("IDCliente", FilterOperator.Equal, data("IDCliente"))
    '            DtDatos = New PrevisionLinea().Filter(Fil)
    '            If Not dtDatos Is Nothing AndAlso dtDatos.Rows.Count > 0 Then
    '                ApplicationService.GenerateError("Ya existe una l�nea de previsi�n para ese cliente, art�culo y fecha")
    '            End If
    '        Case enumtpTipoPrevision.tpZona
    '            Fil.Add("IDZona", FilterOperator.Equal, data("IDZona"))
    '            DtDatos = New PrevisionLinea().Filter(Fil)
    '            If Not dtDatos Is Nothing AndAlso dtDatos.Rows.Count > 0 Then
    '                ApplicationService.GenerateError("Ya existe una l�nea de previsi�n para esa zona y esa fecha")
    '            End If
    '        Case enumtpTipoPrevision.tpClienteTipoFamiliaSubFamilia
    '            ' Recordar que cliente y tipo son obligatorios, pero family y sub no
    '            Fil.Add("IDCliente", FilterOperator.Equal, data("IDCliente"))
    '            Fil.Add("IDTipo", FilterOperator.Equal, data("IDTipo"))
    '            If Length(data("IDFamilia")) = 0 Then
    '                Fil.Add(New IsNullFilterItem("IDFamilia", True))
    '            Else
    '                Fil.Add(New StringFilterItem("IDFamilia", FilterOperator.Equal, data("IDFamilia")))
    '                If Length(data("IDSubfamilia")) = 0 Then
    '                    Fil.Add(New IsNullFilterItem("IDSubfamilia", True))
    '                Else
    '                    Fil.Add("IDSubfamilia", FilterOperator.Equal, data("IDSubFamilia"))
    '                End If
    '            End If
    '            DtDatos = New PrevisionLinea().Filter(Fil)
    '            If Not dtDatos Is Nothing AndAlso dtDatos.Rows.Count > 0 Then
    '                ApplicationService.GenerateError("Ya existe una l�nea de previsi�n para ese cliente, tipo de art�culo/familia/subfamilia y fecha")
    '            End If
    '    End Select
    'End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClave)
    End Sub

    <Task()> Public Shared Sub AsignarClave(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDLineaPrevision")) = 0 Then data("IDLineaPrevision") = AdminData.GetAutoNumeric
        End If
    End Sub

#End Region

#Region "Funciones P�blicas"

    <Serializable()> _
    Public Class DatosGenLinPrev
        Public OrigenFacturas As Boolean
        Public BorrarLineasActuales As Boolean
        Public IDPrevisionDestino As String
        Public Fil As Filter
        Public TipoPrevisionOrigen As enumtpTipoPrevision
        Public TipoPrevisionDestino As enumtpTipoPrevision
        Public IncrementoA�os As Integer
        Public IncrementoCantidad As Double
        Public IncrementoPrecio As Double
        Public IncrementoImporte As Double
    End Class

    <Serializable()> _
    Public Class ResGeneracionLineasPrevision
        Public Exito As Boolean 'True si se han llevado a cabo acciones, False si no se ha pasado ninguna
        Public NLineasNoPasadas As Integer
        Public ValoresNulos As Boolean
        Public HayRepeticion As Boolean
    End Class

    <Task()> Public Shared Function GenerarLineasPrevistas(ByVal data As DatosGenLinPrev, ByVal services As ServiceProvider) As ResGeneracionLineasPrevision
        ' 4 pasos:
        ' 1� Obtenemos dtFinal, que almacenar� los datos a guardar finalmente en la BD.
        '       + Si borramos l�neas actuales, se carga con las l�neas a borrar de la previsi�n destino.
        '       + Si agregamos, se carga con las l�neas actuales de la previsi�n destino.
        '       + Si la previsi�n destino no tiene ninguna fila, obtenemos la estructura de la
        '           tabla previsi�n l�nea
        '
        ' 2� Obtenemos los datos que corresponden con el filtro. Lo obtenemos de:
        '       + Vista de Facturas de Venta, � 
        '       + Vista de las l�neas de previsi�n, una de ellas en funci�n del tipo de previsi�n a generar.
        '
        ' 3� Procesamos los datos y los pasamos a dtFinal, que ser� la tabla a pasar a la base de datos.
        '       + Preparamos columnas, descartando aquellos que no necesitamos seg�n el tipo de previs�n a generar.
        '       + Si se agrega una fila repetida cuando el origen es facturas, sumamos cantidades e importes
        '           siempre que no hagamos la suma a una de las filas preexistentes.
        '       + Transformamos fecha, importe y cantidad con los criterios de c�lculo.
        '
        ' 4� Guardamos dtFinal en la base de datos

        Dim res As New ResGeneracionLineasPrevision

        If data.BorrarLineasActuales Then
            Dim fLineasActuales As New Filter
            Dim ClsPrev As New PrevisionLinea
            fLineasActuales.Add(New StringFilterItem("IDPrevision", data.IDPrevisionDestino))
            Dim dt As DataTable = ClsPrev.Filter(fLineasActuales)
            ClsPrev.Delete(dt)
        End If

        ' 1 - Obtenemos dtFinal, que almacenar� los datos a guardar finalmente en la BD.
        Dim dtFinal As DataTable = New PrevisionLinea().AddNew
        '-----------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------
        ' 2 - Obtenemos los datos que corresponden con el filtro
        Dim dtDatos As New DataTable
        If data.OrigenFacturas Then
            dtDatos = New BE.DataEngine().Filter("vLineasPrevistasFacturas", data.Fil)
        Else
            Select Case data.TipoPrevisionOrigen
                Case enumtpTipoPrevision.tpArticuloCliente
                    dtDatos = New BE.DataEngine().Filter("vLineasPrevistasPrevisionesArticulosClientes", data.Fil)
                Case enumtpTipoPrevision.tpClienteTipoFamiliaSubFamilia
                    dtDatos = New BE.DataEngine().Filter("vLineasPrevistasPrevisionesClientesTiposFamilias", data.Fil)
                Case enumtpTipoPrevision.tpPorArticulo
                    dtDatos = New BE.DataEngine().Filter("vLineasPrevistasPrevisionesArticulos", data.Fil)
                Case enumtpTipoPrevision.tpPorCliente
                    dtDatos = New BE.DataEngine().Filter("vLineasPrevistasPrevisionesClientes", data.Fil)
                Case enumtpTipoPrevision.tpZona
                    dtDatos = New PrevisionLinea().Filter(data.Fil)
            End Select
        End If
        '-----------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------
        ' 3 - Trabajamos con los datos que cumplen el filtro, que los pasaremos a dtFinal
        '       si corresponde.
        If Not dtDatos Is Nothing AndAlso dtDatos.Rows.Count > 0 Then
            If data.OrigenFacturas Then
                'PREPARAMOS COLUMNAS DE dtDatos PQ SU ESTRUCTURA ES DIFERENTE AL VENIR DE FACTURAS
                dtDatos.Columns.Remove(dtDatos.Columns("IDFactura"))
                dtDatos.Columns.Add("PrecioA", GetType(Double))
                'Metemos �sta porque la que nos viene, ImporteFactura, es de s�lo lectura
                dtDatos.Columns.Add("ImporteA", GetType(Double))
            End If

            Dim agregarFila As Boolean
            Dim claves() As Object
            Dim anularPrecio As Boolean
            Dim anularCantidad As Boolean
            Dim anularImporte As Boolean
            Dim drFind As DataRow

            dtFinal.BeginLoadData()

            'Definimos claves principales de dtFinal para evitar repeticiones:
            Dim dc(1) As DataColumn
            dc(0) = dtFinal.Columns("IDPrevision")
            dc(1) = dtFinal.Columns("FechaPrevision")
            Select Case data.TipoPrevisionDestino
                Case enumtpTipoPrevision.tpArticuloCliente
                    ReDim Preserve dc(3)
                    dc(2) = dtFinal.Columns("IDArticulo")
                    dc(3) = dtFinal.Columns("IDCliente")
                    dtFinal.PrimaryKey = dc
                    anularPrecio = True
                Case enumtpTipoPrevision.tpClienteTipoFamiliaSubFamilia
                    ReDim Preserve dc(5)
                    dc(2) = dtFinal.Columns("IDCliente")
                    dc(3) = dtFinal.Columns("IDTipo")
                    dc(4) = dtFinal.Columns("IDFamilia")
                    dc(5) = dtFinal.Columns("IDSubfamilia")
                    dtFinal.PrimaryKey = dc
                Case enumtpTipoPrevision.tpPorArticulo
                    ReDim Preserve dc(2)
                    dc(2) = dtFinal.Columns("IDArticulo")
                    dtFinal.PrimaryKey = dc
                    anularPrecio = True
                    anularImporte = True
                Case enumtpTipoPrevision.tpPorCliente
                    ReDim Preserve dc(2)
                    dc(2) = dtFinal.Columns("IDCliente")
                    dtFinal.PrimaryKey = dc
                    anularPrecio = True
                    anularCantidad = True
                Case enumtpTipoPrevision.tpZona
                    ReDim Preserve dc(2)
                    dc(2) = dtFinal.Columns("IDZona")
                    dtFinal.PrimaryKey = dc
                    anularPrecio = True
                    anularCantidad = True
            End Select

            If data.OrigenFacturas Then
                anularPrecio = True     ' Si venimos de facturas no pasamos el precio.
            End If

            For Each dr As DataRow In dtDatos.Rows
                agregarFila = True

                If data.OrigenFacturas Then
                    dr("ImporteA") = dr("ImporteFactura")
                End If
                If dr.IsNull("FechaPrevision") Then
                    agregarFila = False
                Else
                    If data.IncrementoA�os <> 0 Then
                        dr("FechaPrevision") = CType(dr("FechaPrevision"), System.DateTime).AddYears(data.IncrementoA�os)
                    End If
                End If

                ' COGEMOS LOS DATOS QUE NOS INTERESAN SEG�N EL TIPO DE PREVISI�N DE DESTINO
                Select Case data.TipoPrevisionDestino
                    Case enumtpTipoPrevision.tpArticuloCliente
                        If dr.IsNull("IDArticulo") Or dr.IsNull("IDCliente") Then
                            agregarFila = False
                        Else
                            claves = New Object() {data.IDPrevisionDestino, dr("FechaPrevision"), dr("IDArticulo"), dr("IDCliente")}
                            dr("IDTipo") = System.DBNull.Value
                            dr("IDFamilia") = System.DBNull.Value
                            dr("IDSubfamilia") = System.DBNull.Value
                            dr("IDZona") = System.DBNull.Value
                        End If
                    Case enumtpTipoPrevision.tpClienteTipoFamiliaSubFamilia
                        If dr.IsNull("IDCliente") Or dr.IsNull("IDTipo") Then
                            agregarFila = False
                        Else
                            If dr.IsNull("IDFamilia") And Not dr.IsNull("IDSubfamilia") Then
                                dr("IDSubfamilia") = System.DBNull.Value
                            End If
                            claves = New Object() {data.IDPrevisionDestino, dr("FechaPrevision"), dr("IDCliente"), dr("IDTipo"), dr("IDFamilia"), dr("IDSubfamilia")}
                            dr("IDArticulo") = System.DBNull.Value
                            dr("IDZona") = System.DBNull.Value
                        End If
                    Case enumtpTipoPrevision.tpPorArticulo
                        If dr.IsNull("IDArticulo") Then
                            agregarFila = False
                        Else
                            claves = New Object() {data.IDPrevisionDestino, dr("FechaPrevision"), dr("IDArticulo")}
                            dr("IDTipo") = System.DBNull.Value
                            dr("IDFamilia") = System.DBNull.Value
                            dr("IDSubfamilia") = System.DBNull.Value
                            dr("IDCliente") = System.DBNull.Value
                            dr("IDZona") = System.DBNull.Value
                        End If
                    Case enumtpTipoPrevision.tpPorCliente
                        If dr.IsNull("IDCliente") Then
                            agregarFila = False
                        Else
                            claves = New Object() {data.IDPrevisionDestino, dr("FechaPrevision"), dr("IDCliente")}
                            dr("IDTipo") = System.DBNull.Value
                            dr("IDFamilia") = System.DBNull.Value
                            dr("IDSubfamilia") = System.DBNull.Value
                            dr("IDArticulo") = System.DBNull.Value
                            dr("IDZona") = System.DBNull.Value
                        End If
                    Case enumtpTipoPrevision.tpZona
                        If dr.IsNull("IDZona") Then
                            agregarFila = False
                        Else
                            claves = New Object() {data.IDPrevisionDestino, dr("FechaPrevision"), dr("IDZona")}
                            dr("IDTipo") = System.DBNull.Value
                            dr("IDFamilia") = System.DBNull.Value
                            dr("IDSubfamilia") = System.DBNull.Value
                            dr("IDArticulo") = System.DBNull.Value
                            dr("IDCliente") = System.DBNull.Value
                        End If
                End Select

                If agregarFila Then
                    ' APLICAMOS C�LCULOS SOBRE CANTIDAD, PRECIO E IMPORTE
                    If anularCantidad Then
                        dr("QPrevista") = System.DBNull.Value
                    Else
                        If data.IncrementoCantidad <> 0 And Not dr.IsNull("QPrevista") Then
                            dr("QPrevista") = dr("QPrevista") * (1 + (data.IncrementoCantidad / 100))
                        End If
                    End If

                    If anularPrecio Then
                        dr("PrecioA") = System.DBNull.Value
                    Else
                        If data.IncrementoPrecio <> 0 And Not dr.IsNull("PrecioA") Then
                            dr("PrecioA") = dr("PrecioA") * (1 + (data.IncrementoPrecio / 100))
                        End If
                    End If

                    If anularImporte Then
                        dr("ImporteA") = System.DBNull.Value
                    Else
                        If data.IncrementoImporte <> 0 And Not dr.IsNull("ImporteA") Then
                            dr("ImporteA") = dr("ImporteA") * (1 + (data.IncrementoImporte / 100))
                        End If
                    End If

                    'DETERMINAR SI LA FILA DE dtDatos PASA O NO A dtFinal
                    drFind = dtFinal.Rows.Find(claves)
                    If drFind Is Nothing Then
                        Dim drNewFinal As DataRow = dtFinal.NewRow
                        drNewFinal("IDLineaPrevision") = AdminData.GetAutoNumeric
                        drNewFinal("IDPrevision") = data.IDPrevisionDestino
                        drNewFinal("IDCliente") = dr("IDCliente")
                        drNewFinal("IDTipo") = dr("IDTipo")
                        drNewFinal("IDFamilia") = dr("IDFamilia")
                        drNewFinal("IDSubfamilia") = dr("IDSubfamilia")
                        drNewFinal("IDArticulo") = dr("IDArticulo")
                        drNewFinal("QPrevista") = dr("QPrevista")
                        drNewFinal("FechaPrevision") = dr("FechaPrevision")
                        drNewFinal("IDZona") = dr("IDZona")
                        drNewFinal("PrecioA") = dr("PrecioA")
                        drNewFinal("ImporteA") = dr("ImporteA")
                        dtFinal.Rows.Add(drNewFinal)
                        'dtFinal.LoadDataRow(New Object() {dr("IDLineaPrevision"), dr("IDPrevision"), _
                        '    dr("IDCliente"), dr("IDTipo"), dr("IDFamilia"), dr("IDSubfamilia"), _
                        '    dr("IDArticulo"), dr("QPrevista"), dr("FechaPrevision"), _
                        '    dr("IDZona"), dr("PrecioA"), dr("ImporteA")}, False)
                    Else
                        Dim sumarCTFS As Boolean
                        ' Si la previsi�n no va a ser tpClienteTipoFamiliaSubFamilia
                        If data.OrigenFacturas And drFind.RowState = DataRowState.Added Then
                            sumarCTFS = True
                        End If

                        If Not sumarCTFS Then
                            res.NLineasNoPasadas += 1
                            res.HayRepeticion = True
                        Else
                            ' Si no podemos sumar por nulos la cantidad no pasa nada: 
                            ' no agregamos la l�nea pero tampoco nos lo tomamos como un error.
                            If Not dr.IsNull("QPrevista") Then
                                If Not drFind.IsNull("QPrevista") Then
                                    drFind("QPrevista") = drFind("QPrevista") + dr("QPrevista")
                                Else
                                    drFind("QPrevista") = dr("QPrevista")
                                End If
                            End If
                            If Not dr.IsNull("ImporteA") Then
                                If Not drFind.IsNull("ImporteA") Then
                                    drFind("ImporteA") = drFind("ImporteA") + dr("ImporteA")
                                Else
                                    drFind("ImporteA") = dr("ImporteA")
                                End If
                            End If
                        End If
                    End If
                Else
                    ' agregarFila es False debido a que hay valores nulos
                    res.ValoresNulos = True
                    res.NLineasNoPasadas += 1
                End If
            Next

            If data.TipoPrevisionDestino = enumtpTipoPrevision.tpClienteTipoFamiliaSubFamilia Then
                'En este tipo de previsi�n han variado las claves. Cada vez que se establece 
                'unaclave, aunque sea temporalmente, queda allowdbnull = true y luego 
                'puede dar problemas al cerrar la carga de datos con EndLoadData
                dtFinal.PrimaryKey = Nothing
                dtFinal.Columns("IDFamilia").AllowDBNull = True
                dtFinal.Columns("IDSubfamilia").AllowDBNull = True
            End If
            dtFinal.EndLoadData()
            '-----------------------------------------------------------------------------------

            '-----------------------------------------------------------------------------------
            ' 4 - Guardamos dtFinal en la base de datos
            Dim dtNuevas As DataTable = dtFinal.GetChanges(DataRowState.Added)
            If dtNuevas Is Nothing Then
                res.Exito = False
            Else
                Dim ClsPrevLin As New PrevisionLinea
                If data.BorrarLineasActuales Then
                    Dim dtBorradas As DataTable = dtFinal.GetChanges(DataRowState.Deleted)
                    If Not dtBorradas Is Nothing Then
                        ClsPrevLin.Update(dtBorradas)
                    End If
                End If
                ClsPrevLin.Update(dtNuevas)
                res.Exito = True
            End If
        Else : res.Exito = False
        End If
        Return res
    End Function

#End Region

End Class