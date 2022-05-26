Public MustInherit Class DocumentCabLin
    Inherits Document

    Public Moneda As MonedaInfo
    Public MonedaA As MonedaInfo
    Public MonedaB As MonedaInfo

    Public Cabecera As Cabecera

    Public mAIva As Boolean
    Public Property AIva() As Boolean
        Get
            Return mAIva
        End Get
        Set(ByVal value As Boolean)
            mAIva = value
        End Set
    End Property

    Public MustOverride Function EntidadCabecera() As String
    Public MustOverride Function EntidadLineas() As String
    Public MustOverride Function EntidadAnalitica() As String
    Public MustOverride Property Fecha(Optional ByVal OriginalValue As DataRowVersion = DataRowVersion.Current) As Date
    Public MustOverride Function PrimaryKeyCab() As String()
    Public MustOverride Function PrimaryKeyLin() As String()

#Region " Creación de instancias "

    '//New a utilizar en los procesos de creación de elementos de tipo cabecera/lineas
    Public Sub New(ByVal Cab As Cabecera, ByVal services As ServiceProvider)
        LoadEntityHeader(True, Cab)

        Inicializar(services)
        LoadEntitiesChilds(True)
    End Sub

    '//New a utilizar desde presentación (utilizado por el motor para realizar las actualizaciones de los elementos cabecera/lineas)
    Public Sub New(ByVal UpdtCtx As UpdatePackage)
        '// al crearse la bola desde el UpdatePackage, se debieran eliminar los conjuntos de datos que se van a tratar
        '// y dejar que el motor trate el resto
        LoadEntityHeader(False, Nothing, UpdtCtx)

        Dim services As New ServiceProvider
        Inicializar(services)

        LoadEntitiesChilds(False, UpdtCtx)
    End Sub

    '//New utilizado para obtener un Documento alamacenado en la BBDD.
    Public Sub New(ByVal ParamArray PrimaryKey() As Object)
        LoadEntityHeader(False, Nothing, Nothing, PrimaryKey)

        Dim services As New ServiceProvider
        Inicializar(services)

        LoadEntitiesChilds(False)
    End Sub


    Protected Overridable Sub LoadEntityHeader(ByVal AddNew As Boolean, Optional ByVal Cab As Cabecera = Nothing, Optional ByVal UpdtCtx As UpdatePackage = Nothing, Optional ByVal PrimaryKey() As Object = Nothing)
        If AddNew Then '//New de Procesos
            Dim oBusinessEntity As BusinessHelper = BusinessHelper.CreateBusinessObject(EntidadCabecera)
            Dim dtCabeceras As DataTable = oBusinessEntity.AddNew

            dtCabeceras.Rows.Add(dtCabeceras.NewRow)
            Me.AddHeader(EntidadCabecera, dtCabeceras)     '//Creamos el HeaderRow

            Me.Cabecera = Cab

            If HeaderRow.Table.Columns.Contains("IDMoneda") Then
                If Length(Me.Cabecera.IDMoneda) > 0 Then Me.HeaderRow("IDMoneda") = Me.Cabecera.IDMoneda
            End If

            Me.Fecha = Me.Cabecera.Fecha

            'Me.HeaderRow("IDBancoPropio") = Me.Cabecera.IDBancoPropio
            If Length(Me.Cabecera.IDCentroGestion) > 0 Then Me.HeaderRow("IDCentroGestion") = Me.Cabecera.IDCentroGestion
            If Length(Me.Cabecera.IDFormaPago) > 0 Then Me.HeaderRow("IDFormaPago") = Me.Cabecera.IDFormaPago
            If Length(Me.Cabecera.IDCondicionPago) > 0 Then Me.HeaderRow("IDCondicionPago") = Me.Cabecera.IDCondicionPago

        ElseIf UpdtCtx Is Nothing Then  '//New de PrimaryKey
            Dim oBusinessEntity As BusinessHelper = BusinessHelper.CreateBusinessObject(EntidadCabecera)
            Dim dtCabeceras As DataTable = oBusinessEntity.SelOnPrimaryKey(PrimaryKey)
            AddHeader(EntidadCabecera, dtCabeceras)
        Else '//New del formulario
            Dim dtCabecera As DataTable = UpdtCtx(EntidadCabecera).First
            AddHeader(EntidadCabecera, UpdtCtx(EntidadCabecera).First)
            If HeaderRow.Table.Columns.Contains("CambioA") AndAlso HeaderRow.Table.Columns.Contains("CambioB") Then
                If Not dtCabecera Is Nothing AndAlso dtCabecera.Rows.Count > 0 Then
                    If (dtCabecera.Rows(0).RowState = DataRowState.Modified AndAlso _
                       (dtCabecera.Rows(0)("CambioA") <> dtCabecera.Rows(0)("CambioA", DataRowVersion.Original) OrElse _
                        dtCabecera.Rows(0)("CambioB") <> dtCabecera.Rows(0)("CambioB", DataRowVersion.Original))) OrElse _
                        dtCabecera.Rows(0).RowState = DataRowState.Added Then
                        Me.HeaderRow("CambioA") = dtCabecera.Rows(0)("CambioA")
                        Me.HeaderRow("CambioB") = dtCabecera.Rows(0)("CambioB")

                        Me.CambioA = dtCabecera.Rows(0)("CambioA")
                        Me.CambioB = dtCabecera.Rows(0)("CambioB")
                    End If
                End If

                Dim PKCabecera() As String = PrimaryKeyCab()
                MergeData(UpdtCtx, EntidadCabecera, PKCabecera, PKCabecera, True)
            End If
        End If
    End Sub
    Protected Overridable Sub LoadEntitiesChilds(ByVal AddNew As Boolean, Optional ByVal UpdtCtx As UpdatePackage = Nothing)
        Dim oEntidadLineas As BusinessHelper = BusinessHelper.CreateBusinessObject(EntidadLineas)
        Dim dtLineas As DataTable
        If AddNew Then  '//New de Procesos
            dtLineas = oEntidadLineas.AddNew
        ElseIf UpdtCtx Is Nothing Then  '//New de PrimaryKey
            Dim PKCabecera() As String = PrimaryKeyCab()
            dtLineas = oEntidadLineas.Filter(New FilterItem(PKCabecera(0), HeaderRow(PKCabecera(0))))
        Else  '//New del formulario
            '// al crearse la bola desde el UpdatePackage, se debieran eliminar los conjuntos de datos que se van a tratar
            '// y dejar que el motor trate el resto  (MergeData)
            Dim PKCabecera() As String = PrimaryKeyCab()
            dtLineas = MergeData(UpdtCtx, EntidadLineas, PKCabecera, PKCabecera, True)
        End If
        Me.Add(EntidadLineas, dtLineas)

        LoadEntitiesGrandChilds(AddNew)
    End Sub
    Protected Overridable Sub LoadEntitiesGrandChilds(ByVal AddNew As Boolean)
        Dim PKLineas() As String = PrimaryKeyLin()
        Dim ids(dtLineas.Rows.Count - 1) As Object
        For i As Integer = 0 To dtLineas.Rows.Count - 1
            ids(i) = dtLineas.Rows(i)(PKLineas(0))
        Next
        If ids.Length = 0 Then ids = New Object() {0}

        Dim f As New Filter
        f.Add(New InListFilterItem(PKLineas(0), ids, FilterType.Numeric))

        'TODO falta cargar la analitica de las lineas nuevas
        LoadEntityGrandChild(EntidadAnalitica, f, AddNew)
    End Sub

    Protected Overridable Sub LoadEntityGrandChild(ByVal EntityName As String, ByVal f As Filter, ByVal AddNew As Boolean)
        If Length(EntityName) > 0 Then
            Dim oEntidad As BusinessHelper = BusinessHelper.CreateBusinessObject(EntityName)
            Dim dt As DataTable
            If AddNew Then
                dt = oEntidad.AddNew
            Else
                dt = oEntidad.Filter(f)
            End If
            Me.Add(EntityName, dt)
        End If
    End Sub

    Protected Overridable Sub Inicializar(ByVal services As ServiceProvider)
        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()

        If HeaderRow.Table.Columns.Contains("IDMoneda") Then
            Me.IDMoneda = HeaderRow("IDMoneda")
            If Moneda Is Nothing Then
                Me.MonedaA = Monedas.MonedaA
                Me.MonedaB = Monedas.MonedaB
                Me.Moneda = Monedas.GetMoneda(IDMoneda, Fecha)
            End If

            If HeaderRow.Table.Columns.Contains("CambioA") AndAlso HeaderRow.Table.Columns.Contains("CambioB") Then
                If Me.HeaderRow.IsNull("CambioA") OrElse Me.HeaderRow("CambioA") = 0 Then
                    Me.HeaderRow("CambioA") = Me.Moneda.CambioA
                    Me.HeaderRow("CambioB") = Me.Moneda.CambioB

                    Me.CambioA = Me.Moneda.CambioA
                    Me.CambioB = Me.Moneda.CambioB
                End If
            End If

        End If

        If Length(Me.HeaderRow("IDContador")) > 0 Then
            Dim dtContador As DataTable = New Contador().SelOnPrimaryKey(Me.HeaderRow("IDContador"))
            If dtContador.Rows.Count > 0 Then
                Me.AIva = dtContador.Rows(0)("AIVA")
            End If
        Else
            Me.AIva = True
        End If

        ''TODO esto se hace así porque despues se usan los cambios de Fra.Moneda
        'If DocFra.HeaderRow.IsNull("CambioA") OrElse DocFra.HeaderRow("CambioA") = 0 Then DocFra.HeaderRow("CambioA") = DocFra.Moneda.CambioA '= DocFra.HeaderRow("CambioA")
        'If DocFra.HeaderRow.IsNull("CambioB") Then DocFra.Moneda.CambioB = DocFra.HeaderRow("CambioB")
    End Sub
#End Region


    Public Property IDMoneda() As String
        Get
            Return Me.HeaderRow("IDMoneda") & String.Empty
        End Get
        Set(ByVal value As String)
            Me.HeaderRow("IDMoneda") = value
        End Set
    End Property

    Public Property CambioA() As Double
        Get
            Return Nz(Me.HeaderRow("CambioA"), 0)
        End Get
        Set(ByVal value As Double)
            HeaderRow("CambioA") = value
        End Set
    End Property

    Public Property CambioB() As Double
        Get
            Return Nz(Me.HeaderRow("CambioB"), 0)
        End Get
        Set(ByVal value As Double)
            HeaderRow("CambioB") = value
        End Set
    End Property

    Public ReadOnly Property dtLineas() As DataTable
        Get
            Return MyBase.Item(EntidadLineas)
        End Get
    End Property

    Public ReadOnly Property dtAnalitica() As DataTable
        Get
            Return MyBase.Item(EntidadAnalitica)
        End Get
    End Property


#Region " Métodos auxiliares "

    Protected Function GetPrimaryKey(ByVal oBusinessEntity As BusinessHelper) As String()
        Dim dtKeys As DataTable = oBusinessEntity.PrimaryKeyTable
        Dim PrimaryKey() As DataColumn
        dtKeys.Columns.CopyTo(PrimaryKey, 0)
        Return GetPrimaryKey(PrimaryKey)
    End Function

    Protected Function GetPrimaryKey(ByVal PrimaryKey() As DataColumn) As String()
        Dim Keys(PrimaryKey.Length) As String
        For i As Integer = 0 To PrimaryKey.Length - 1
            Keys(i) = PrimaryKey(i).ColumnName
        Next
        Return Keys
        'Return Join(Keys, ", ")
    End Function

#End Region

End Class
