Public Class ArticuloUnidadAB

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Private Const cnEntidad As String = "tbArticuloUnidadAB"

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

#End Region

#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("IDUDMedidaA", AddressOf CambioMedida)
        oBrl.Add("Factor", AddressOf CambioFactor)
        Return oBrl
    End Function

    <Task()> Public Shared Sub CambioMedida(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim ud As New UdMedida
        ud.GetItemRow(data.Value)
    End Sub

    <Task()> Public Shared Sub CambioFactor(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Not IsNumeric(data.Value) Then ApplicationService.GenerateError("Campo no numérico.")
    End Sub

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDArticulo")) = 0 Then ApplicationService.GenerateError("El Artículo es obligatorio.")
        If Length(data("IDUdMedidaA")) = 0 Then ApplicationService.GenerateError("La Ud. Medida A es obligatoria.")
        If Length(data("IDUdMedidaB")) = 0 Then ApplicationService.GenerateError("La Ud. Medida B es obligatoria.")
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New ArticuloUnidadAB().SelOnPrimaryKey(data("IDArticulo"), data("IDUdMedidaA"), data("IDUdMedidaB"))
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("El factor ya esta configurado para este Artículo respecto a las unidades de medida A y B.")
            End If
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Serializable()> _
    Public Class DatosFactorConversion
        Public IDArticulo As String
        Public IDUdMedidaA As String
        Public IDUdMedidaB As String
        Public UnoSiNoExiste As Boolean

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDArticulo As String, ByVal IDUdMedidaA As String, ByVal IDUdMedidaB As String, Optional ByVal UnoSiNoExiste As Boolean = True)
            Me.IDArticulo = IDArticulo
            Me.IDUdMedidaA = IDUdMedidaA
            Me.IDUdMedidaB = IDUdMedidaB
            Me.UnoSiNoExiste = UnoSiNoExiste
        End Sub
    End Class

    <Task()> Public Shared Function FactorDeConversion(ByVal data As DatosFactorConversion, ByVal services As ServiceProvider) As Double
        Dim oFltr As Filter = New Filter(FilterUnionOperator.Or)
        Dim dblFactor As Double
        Dim blnDividir As Boolean

        Dim oFltrA As Filter = New Filter
        oFltrA.Add("IDArticulo", FilterOperator.Equal, data.IDArticulo)
        If Length(data.IDUdMedidaA) > 0 Then oFltrA.Add("IDUdMedidaA", FilterOperator.Equal, data.IDUdMedidaA)
        If Length(data.IDUdMedidaB) > 0 Then oFltrA.Add("IDUdMedidaB", FilterOperator.Equal, data.IDUdMedidaB)
        oFltr.Add(oFltrA)

        Dim oFltrB As Filter = New Filter
        oFltrB.Add("IDArticulo", FilterOperator.Equal, data.IDArticulo)
        If Length(data.IDUdMedidaB) > 0 Then oFltrB.Add("IDUdMedidaA", FilterOperator.Equal, data.IDUdMedidaB)
        If Length(data.IDUdMedidaA) > 0 Then oFltrB.Add("IDUdMedidaB", FilterOperator.Equal, data.IDUdMedidaA)
        oFltr.Add(oFltrB)

        Dim dt As DataTable = New ArticuloUnidadAB().Filter(oFltr)

        Select Case dt.Rows.Count
            Case 0
                Dim UDMedida As New UnidadAB.UnidadMedidaInfo
                UDMedida.IDUdMedidaA = data.IDUdMedidaA
                UDMedida.IDUdMedidaB = data.IDUdMedidaB
                UDMedida.UnoSiNoExiste = data.UnoSiNoExiste
                dblFactor = ProcessServer.ExecuteTask(Of UnidadAB.UnidadMedidaInfo, Double)(AddressOf UnidadAB.FactorDeConversion, UDMedida, services)
                If dblFactor = 0 Then
                    If data.UnoSiNoExiste OrElse data.IDUdMedidaA & String.Empty = data.IDUdMedidaB & String.Empty Then
                        dblFactor = 1
                    Else : dblFactor = 0
                    End If
                End If
            Case 1
                Dim oRw As DataRow = dt.Rows(0)
                dblFactor = oRw("Factor")
                blnDividir = (CStr(oRw("IDUdMedidaA")) = data.IDUdMedidaB)
            Case 2
                If CStr(dt.Rows(0)("IDUdMedidaA")) = data.IDUdMedidaA Then
                    dblFactor = dt.Rows(0)("Factor")
                Else : dblFactor = dt.Rows(1)("Factor")
                End If
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

    <Serializable()> _
    Public Class DatosFactorConverInterna
        Public IDArticulo As String
        Public IDUdMedida As String
        Public IDUdInterna As String
    End Class

    <Task()> Public Shared Function FactorDeConversionConUdInterna(ByVal data As DatosFactorConverInterna, ByVal services As ServiceProvider) As Double
        Dim dblFactor As Double = 1
        If Len(data.IDArticulo) > 0 Then
            If Len(data.IDUdInterna) = 0 Then
                Dim Art As New Articulo
                Dim dtArticulo As DataTable = Art.SelOnPrimaryKey(data.IDArticulo)
                If Not dtArticulo Is Nothing AndAlso dtArticulo.Rows.Count > 0 Then
                    data.IDUdInterna = dtArticulo.Rows(0)("IDUDInterna")
                End If
            End If
            Dim StDatos As New DatosFactorConversion
            StDatos.IDArticulo = data.IDArticulo
            StDatos.IDUdMedidaA = data.IDUdMedida
            StDatos.IDUdMedidaB = data.IDUdInterna
            dblFactor = ProcessServer.ExecuteTask(Of DatosFactorConversion, Double)(AddressOf FactorDeConversion, StDatos, services)
        End If
        Return dblFactor
    End Function

#End Region

End Class