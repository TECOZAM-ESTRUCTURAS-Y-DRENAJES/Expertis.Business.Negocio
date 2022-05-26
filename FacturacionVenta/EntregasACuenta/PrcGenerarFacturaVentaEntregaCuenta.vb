Public Class PrcGenerarFacturaVentaEntregaCuenta
    Inherits Process(Of DataPrcFacturacionEntregas, ResultFacturacion)

    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of DataPrcFacturacionEntregas)(AddressOf DatosIniciales)
        Me.AddTask(Of DataPrcFacturacionEntregas, FraCabEntregaCta())(AddressOf AgruparEntregas)
        Me.AddForEachTask(Of PrcCrearFacturaVentaEntregaCta)(BusinessProcesses.OnExceptionBehaviour.NextLoop)
        Me.AddTask(Of Object, ResultFacturacion)(AddressOf ProcesoComunes.GetResultadoFacturacion)
    End Sub

    <Task()> Public Shared Sub DatosIniciales(ByVal data As DataPrcFacturacionEntregas, ByVal services As ServiceProvider)
        '//Prepara en el service información del proceso.
        If Length(data.IDContador) > 0 Then
            services.RegisterService(New ProcessInfo(data.IDContador))
        End If
    End Sub

    <Task()> Public Shared Function AgruparEntregas(ByVal data As DataPrcFacturacionEntregas, ByVal services As ServiceProvider) As FraCabEntregaCta()
        Dim IDEntregasCopy(data.IDEntregas.Length - 1) As Object
        data.IDEntregas.CopyTo(IDEntregasCopy, 0)
        Dim dtEntregas As DataTable = New EntregasACuenta().Filter(New InListFilterItem("IDEntrega", IDEntregasCopy, FilterType.Numeric))
        If dtEntregas.Rows.Count > 0 Then
            Dim oGrprUser As New GroupUserEntregasCtaVenta

            Dim grpEntrega(0) As DataColumn
            grpEntrega(0) = dtEntregas.Columns("IDEntrega")
            Dim groupers(0) As GroupHelper
            groupers(0) = New GroupHelper(grpEntrega, oGrprUser)
            For Each drEntrega As DataRow In dtEntregas.Rows
                groupers(0).Group(drEntrega)
            Next

            Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
            Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
            For Each fra As FraCabEntregaCta In oGrprUser.Fras
                If Length(fra.IDFormaPago) = 0 Then fra.IDFormaPago = AppParams.FormaPagoEfectivo
                If Length(fra.IDMoneda) = 0 Then
                    Dim MonInfoA As MonedaInfo = Monedas.MonedaA
                    fra.IDMoneda = MonInfoA.ID
                End If
                fra.IDCondicionPago = AppParams.CondicionPagoEfectivo
            Next

            Return oGrprUser.Fras
        End If
    End Function

End Class


<Serializable()> _
Public Class DataPrcFacturacionEntregas

    Public IDEntregas() As Integer
    Public IDContador As String

    Public Sub New(ByVal IDEntregas() As Integer, Optional ByVal IDContador As String = Nothing)
        Me.IDEntregas = IDEntregas
        If Length(IDContador) > 0 Then Me.IDContador = IDContador
    End Sub

End Class