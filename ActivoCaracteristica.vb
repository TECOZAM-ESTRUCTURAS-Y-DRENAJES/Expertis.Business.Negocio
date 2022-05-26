Public Class ActivoCaracteristica

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbActivoCaracteristica"

#End Region

#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("IDCaracteristica", AddressOf ValidarCaracteristica)
        oBrl.Add("Valor", AddressOf ValidarValor)
        Return oBrl
    End Function

    <Task()> Public Shared Sub ValidarCaracteristica(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDCaracteristica")) > 0 Then
            Dim dr As DataRow = New Caracteristica().GetItemRow(data.Current("IDCaracteristica"))
            If Not IsNothing(dr) Then
                If data.Current.ContainsKey("DescCaracteristica") Then data.Current("DescCaracteristica") = dr("DescCaracteristica")
                If data.Current.ContainsKey("Orden") Then data.Current("Orden") = dr("Orden")
                If data.Current.ContainsKey("TipoValor") Then data.Current("TipoValor") = dr("TipoValor")

                Dim dtAG As DataTable = New CaracteristicaAgrupacion().SelOnPrimaryKey(dr("IDAgrupacion"))
                If Not IsNothing(dtAG) AndAlso dtAG.Rows.Count > 0 Then
                    If data.Current.ContainsKey("IDAgrupacion") Then data.Current("IDAgrupacion") = dtAG.Rows(0)("IDAgrupacion") & String.Empty
                    If data.Current.ContainsKey("DescAgrupacion") Then data.Current("DescAgrupacion") = dtAG.Rows(0)("DescAgrupacion") & String.Empty
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarValor(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDCaracteristica")) > 0 AndAlso Length(data.Current("Valor")) > 0 Then
            Dim objNegCaract As New Caracteristica
            Dim dt As DataTable = objNegCaract.SelOnPrimaryKey(data.Current("IDCaracteristica"))
            If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
                If dt.Rows(0)("TipoDato") = enumTipoDato.Numerico AndAlso Not IsNumeric(data.Current("Valor")) Then
                    ApplicationService.GenerateError("Tipo de valor incorrecto debe ser numérico")
                End If
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarCaracteristicaObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarExisteCaracteristica)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarTipoDato)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarCaracteristicaValor)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarCaracteristicaActivo)
    End Sub

    <Task()> Public Shared Sub ValidarCaracteristicaObligatoria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDCaracteristica")) = 0 Then ApplicationService.GenerateError("El código de la Característica es obligatorio")
    End Sub

    <Task()> Public Shared Sub ValidarExisteCaracteristica(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim dtCrt As DataTable = New Caracteristica().SelOnPrimaryKey(data("IDCaracteristica"))
        If IsNothing(dtCrt) OrElse dtCrt.Rows.Count = 0 Then ApplicationService.GenerateError("El código de Característica no existe")
    End Sub

    <Task()> Public Shared Sub ValidarTipoDato(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim dtCrt As DataTable = New Caracteristica().SelOnPrimaryKey(data("IDCaracteristica"))
        If dtCrt.Rows(0)("TipoDato") = enumTipoDato.Numerico Then
            If Not IsNumeric(data("Valor")) Then
                ApplicationService.GenerateError("Tipo de valor incorrecto debe ser numérico")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarCaracteristicaValor(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim dtCrt As DataTable = New Caracteristica().SelOnPrimaryKey(data("IDCaracteristica"))
        If Length(data("Valor")) > 0 AndAlso dtCrt.Rows(0)("TipoValor") = enumTipoValor.Discreto Then
            'Valor discreto. Ciertos valores sólo
            '//PENDIENTE- El Configurador no está todavía.
            Dim dt As DataTable = New CaracteristicaValor().SelOnPrimaryKey(data("IDCaracteristica"), data("Valor"))
            If IsNothing(dt) OrElse dt.Rows.Count = 0 Then
                ApplicationService.GenerateError("El valor introducido no se encuentra entre los valores posibles de la característica")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarCaracteristicaActivo(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dtCrt As DataTable = New ActivoCaracteristica().SelOnPrimaryKey(data("IDActivo"), data("IDCaracteristica"))
            If Not IsNothing(dtCrt) AndAlso dtCrt.Rows.Count > 0 Then
                ApplicationService.GenerateError("Ya está definida ésta característica para el activo actual")
            End If
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Serializable()> _
    Public Class DatosRecupCaract
        Public IDArticulo As String
        Public IDActivo As String
        Public DtCaractActualesActivo As DataTable

        Public Sub New(ByVal IDArticulo As String, ByVal IDActivo As String, ByVal DtCaractActualesActivo As DataTable)
            Me.IDArticulo = IDArticulo
            Me.IDActivo = IDActivo
            Me.DtCaractActualesActivo = DtCaractActualesActivo
        End Sub
    End Class

    <Task()> Public Shared Function RecuperaCaracteristicas(ByVal data As DatosRecupCaract, ByVal services As ServiceProvider) As DataTable
        '//Se obtienen las características del artículo
        Dim dtArticuloCaract As DataTable = New BE.DataEngine().Filter("vFrmMntoArticuloNSerieCaracteristica", New StringFilterItem("IDArticulo", data.IDArticulo))
        Dim dtResultado As DataTable = data.DtCaractActualesActivo.Clone
        If Not IsNothing(dtArticuloCaract) AndAlso dtArticuloCaract.Rows.Count > 0 Then
            For Each dr As DataRow In dtArticuloCaract.Rows
                Dim agregar As Boolean = False

                'Miramos si la característica del artículo ya está con el activo actual
                If Not data.DtCaractActualesActivo Is Nothing AndAlso data.DtCaractActualesActivo.Rows.Count > 0 Then
                    Dim f As New Filter
                    f.Add(New StringFilterItem("IDCaracteristica", dr("IDCaracteristica")))
                    Dim WhereCaracteristica As String = f.Compose(New AdoFilterComposer)
                    Dim drFiltro() As DataRow = data.DtCaractActualesActivo.Select(WhereCaracteristica)
                    If drFiltro.Length = 0 Then
                        agregar = True
                    End If
                Else
                    agregar = True
                End If

                'Si no lo está, la damos de alta.
                If agregar Then
                    Dim drNew As DataRow = dtResultado.NewRow
                    drNew("IDActivo") = data.IDActivo
                    drNew("IDCaracteristica") = dr("IDCaracteristica")
                    drNew("DescCaracteristica") = dr("DescCaracteristica")
                    drNew("IDAgrupacion") = dr("IDAgrupacion")
                    drNew("Valor") = System.DBNull.Value
                    dtResultado.Rows.Add(drNew)
                End If
            Next
        End If

        Return dtResultado
    End Function

#End Region

End Class