Public Class PromocionRegalo

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbPromocionRegalo"

#End Region

#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim OBRL As New BusinessRules
        OBRL.Add("IDArticuloRegalo", AddressOf CambioArticulo)
        OBRL.Add("QRegalo", AddressOf CambioCantidad)
        Return OBRL
    End Function

    <Task()> Public Shared Sub CambioArticulo(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim a As New Articulo
            Dim dr As DataRow = a.GetItemRow(data.Value)
            data.Current("DescArticulo") = dr("DescArticulo")
        End If
    End Sub

    <Task()> Public Shared Sub CambioCantidad(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            If Not IsNumeric(data.Value) Then
                ApplicationService.GenerateError("Campo no numérico.")
            ElseIf data.Value <= 0 Then
                ApplicationService.GenerateError("La cantidad a regalar debe ser superior a cero.")
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarIDArticuloRegalo)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClavePrimaria)
    End Sub

    <Task()> Public Shared Sub ValidarIDArticuloRegalo(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDArticuloRegalo")) = 0 Then ApplicationService.GenerateError("El Artículo Regalo es un dato obligatorio.")
    End Sub

    ''' <summary>
    ''' Comprobar que no exista la clave
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarClavePrimaria(ByVal dr As DataRow, ByVal services As ServiceProvider)
        If dr.RowState = DataRowState.Added Then
            If Length(dr("IDArticuloRegalo")) = 0 Then ApplicationService.GenerateError("El Artículo Regalo es un dato obligatorio.")
            If Length(dr("IDPromocionLinea")) > 0 Then
                Dim dt As DataTable = New PromocionRegalo().SelOnPrimaryKey(dr("IDPromocionLinea"), dr("IDArticuloRegalo"))
                If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                    ApplicationService.GenerateError("El Artículo Regalo introducido ya existe para la línea de promoción.")
                End If
            End If
        End If
    End Sub

#End Region

End Class