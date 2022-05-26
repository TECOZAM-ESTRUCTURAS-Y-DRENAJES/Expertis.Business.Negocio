Public Class CaracteristicaInfo
    Inherits ClassEntityInfo

    Public IDCaracteristica As String
    Public DescCaracteristica As String
    Public IDAgrupacion As String
    Public TipoDato As enumTipoDato
    Public TipoValor As enumTipoValor
    Public IDUdMedida As String
    Public IDCaracteristicaPadre As String
    Public TipoCaracteristica As enumTipoCaracteristica
    Public IDFormula As String

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal data As DataRow)
        MyBase.New(data)
    End Sub

    Public Sub New(ByVal IDCaracteristica As String)
        MyBase.New()
        Me.Fill(IDCaracteristica)
    End Sub

    Public Overloads Overrides Sub Fill(ByVal ParamArray PrimaryKey() As Object)
        If Length(PrimaryKey(0)) = 0 Then Exit Sub
        'Dim dtCaractInfo As DataTable = New Caracteristica().SelOnPrimaryKey(PrimaryKey(0))
        Dim dtCaractInfo As DataTable = New BE.DataEngine().Filter("tbMaestroCaracteristica", New StringFilterItem("IDCaracteristica", PrimaryKey(0)), "IDCaracteristica,DescCaracteristica,IDAgrupacion,TipoDato,TipoValor,IDUdMedida,IDCaracteristicaPadre,TipoCaracteristica,IDFormula")
        If dtCaractInfo.Rows.Count > 0 Then
            Me.Fill(dtCaractInfo.Rows(0))
        End If
    End Sub

End Class

Public Class Caracteristica

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroCaracteristica"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarIdentificadorCorrecto)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim fwnFormula As Object = BusinessHelper.CreateBusinessObject(AdminData.GetEntityInfo("cfgMaestroFormula"))
        If Length(data("IDCaracteristica")) = 0 Then
            ApplicationService.GenerateError("El IDCaracteristica es un dato obligatorio.")
        Else
            Dim dtFormula As DataTable = fwnFormula.Filter(New FilterItem("IDFormula", FilterOperator.Equal, data("IDCaracteristica")))
            If dtFormula.Rows.Count <> 0 Then ApplicationService.GenerateError("El código de la característica no puede existir como código de fórmula")
        End If
        If Length(data("TipoValor")) = 0 Then ApplicationService.GenerateError("Tipo Valor es un dato obligatorio.")
        If Length(data("QObligatoria")) = 0 Then data("QObligatoria") = 0
        If Length(data("TipoCaracteristica")) = 0 Then
            ApplicationService.GenerateError("Tipo Caracteristica es un dato obligatorio.")
        ElseIf data("TipoCaracteristica") = enumTipoCaracteristica.Formula Then
            If Length(data("IDFormula")) = 0 Then ApplicationService.GenerateError("El tipo Característica es Fórmula, debe introducir la Fórmula")
        End If
        If data.RowState = DataRowState.Modified Then
            If data("TipoDato") <> data("TipoDato", DataRowVersion.Original) AndAlso data("TipoDato") = enumTipoDato.Numerico Then
                Dim dtCaracteristicaValor As DataTable = New CaracteristicaValor().Filter(New StringFilterItem("IDCaracteristica", data("IDCaracteristica")))
                If Not IsNothing(dtCaracteristicaValor) AndAlso dtCaracteristicaValor.Rows.Count > 0 Then
                    For Each drValor As DataRow In dtCaracteristicaValor.Rows
                        If Not IsNumeric(drValor("IDValor")) Then
                            ApplicationService.GenerateError("Los valores de la característica deben ser numéricos.")
                        End If
                    Next
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarIdentificadorCorrecto(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim PalabrasReservadas() As String = {"IF", "THEN", "TABLA", "TRUE", "FALSE", "AND", "OR", "ANDALSO", "ORELSE", "NOT", "DIM", "AS", "ELSE", "ELSEIF", "END", "MOD", "INT", "ABS", _
                                            "SQRT", "STRING", "BOOLEAN", "INTEGER", "XROUND", "ROUND", "DOUBLE", "ME", _
                                            "=", "<>", "<", ">", "(", ")", "<=", ">=", "=>", "=<", _
                                            "+", "-", "*", "/", " ", ",", Chr(13), Chr(10)}
        Dim CaracteresEspeciales() As Char = {"+", "-", "*", "/", ",", " ", "=", "(", ")", "<", ">", Chr(13), Chr(10), "."}

        If data.RowState = DataRowState.Added AndAlso Length(data("IDCaracteristica")) > 0 Then
            For Each car As Char In CaracteresEspeciales
                If InStr(data("IDCaracteristica"), car, CompareMethod.Text) > 0 Then
                    ApplicationService.GenerateError("El identificador de la Característica no puede contener el carácter {0}.", Quoted(car))
                End If
            Next

            For Each palabra As String In PalabrasReservadas
                If UCase(data("IDCaracteristica")) = UCase(palabra) Then
                    ApplicationService.GenerateError("El identificador de la Característica no puede ser la palabra reservada {0}.", Quoted(palabra))
                End If
            Next

            If IsNumeric(Left(data("IDCaracteristica"), 1)) Then
                ApplicationService.GenerateError("El identificador de la Característica no puede ser numérico ni empezar por un número.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New Caracteristica().Filter(New FilterItem("IDCaracteristica", data("IDCaracteristica")))
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("El IDCaracteristica introducido ya existe")
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf ActualizarInformacionConfigurador)
    End Sub

    <Task()> Public Shared Sub ActualizarInformacionConfigurador(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Modified Then
            Dim dtCrt As DataTable = New Caracteristica().SelOnPrimaryKey(data("IDCaracteristica"))
            If Not IsNothing(dtCrt) AndAlso dtCrt.Rows.Count > 0 Then
                'Si el campo TipoValor cambia de Discreto a Continuo
                If data("TipoValor") <> dtCrt.Rows(0)("TipoValor") AndAlso dtCrt.Rows(0)("TipoValor") = enumTipoValor.Discreto AndAlso data("TipoValor") = enumTipoValor.Continuo Then
                    'Eliminamos
                    Dim objFilter As New Filter
                    objFilter.Add(New StringFilterItem("IDCaracteristica", data("IDCaracteristica")))

                    Dim objNegFamiliaCaracteristicaValor As Object = BusinessHelper.CreateBusinessObject("CfgFamiliaCaracteristicaValor")
                    Dim objNegArticuloCaracteristicaDiscreta As Object = BusinessHelper.CreateBusinessObject("CfgArticuloCaractDiscreta")
                    Dim dtFamiliaCaracteristicaValor As DataTable = objNegFamiliaCaracteristicaValor.Filter(objFilter)
                    Dim dtArticuloCaracteristicaDiscreta As DataTable = objNegArticuloCaracteristicaDiscreta.Filter(objFilter)

                    If dtFamiliaCaracteristicaValor.Rows.Count Then
                        ApplicationService.GenerateError("La caracteristica esta siendo utilizada en una familia de configuración, como Discreta.")
                    ElseIf dtArticuloCaracteristicaDiscreta.Rows.Count Then
                        ApplicationService.GenerateError("La caracteristica esta siendo utilizada en un artículo de configuración, como Discreta.")
                    Else
                        Dim objNegCaractValor As New CaracteristicaValor
                        Dim dtCrtValor As DataTable = objNegCaractValor.Filter(objFilter)
                        If Not IsNothing(dtCrtValor) AndAlso dtCrtValor.Rows.Count > 0 Then
                            objNegCaractValor.Delete(dtCrtValor)
                        End If
                    End If
                End If
            End If
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Serializable()> _
    Public Class DatosDuplicarCarac
        Public IDCaracteristicaOrigen As String
        Public IDCaracteristicaDestino As String
        Public DescCaracteristicaDestino As String
    End Class

    <Task()> Public Shared Sub DuplicarCaracteristica(ByVal data As DatosDuplicarCarac, ByVal services As ServiceProvider)
        Dim blnDiscreta As Boolean
        '//Duplicamos la cabecera
        Dim ClsCaract As New Caracteristica
        Dim dtCaractOrigen As DataTable = ClsCaract.SelOnPrimaryKey(data.IDCaracteristicaOrigen)
        If Not IsNothing(dtCaractOrigen) AndAlso dtCaractOrigen.Rows.Count > 0 Then
            Dim dtCaractDestino As DataTable = ClsCaract.AddNewForm
            dtCaractDestino.Rows(0).ItemArray = dtCaractOrigen.Rows(0).ItemArray
            dtCaractDestino.Rows(0)("IDCaracteristica") = data.IDCaracteristicaDestino
            dtCaractDestino.Rows(0)("DescCaracteristica") = data.DescCaracteristicaDestino
            blnDiscreta = (dtCaractOrigen.Rows(0)("TipoValor") = enumTipoValor.Discreto)
            ClsCaract.Update(dtCaractDestino)
        End If

        '//Duplicamos los valores si es una característica Discreta.
        If blnDiscreta Then
            Dim objNegCfgCaractValor As New CaracteristicaValor
            Dim objFilterCaract As New Filter
            objFilterCaract.Add(New StringFilterItem("IDCaracteristica", data.IDCaracteristicaOrigen))
            Dim dtCaractValorOrigen As DataTable = objNegCfgCaractValor.Filter(objFilterCaract)
            Dim dtCaractValorDestino As DataTable = dtCaractValorOrigen.Clone
            Dim drNewCaractValor As DataRow
            If Not IsNothing(dtCaractValorOrigen) AndAlso dtCaractValorOrigen.Rows.Count > 0 Then
                For Each drCaractValor As DataRow In dtCaractValorOrigen.Rows
                    drNewCaractValor = dtCaractValorDestino.NewRow
                    drNewCaractValor.ItemArray = drCaractValor.ItemArray
                    drNewCaractValor("IDCaracteristica") = data.IDCaracteristicaDestino
                    dtCaractValorDestino.Rows.Add(drNewCaractValor)
                Next
                objNegCfgCaractValor.Update(dtCaractValorDestino)
            End If
        End If
    End Sub

#End Region

End Class