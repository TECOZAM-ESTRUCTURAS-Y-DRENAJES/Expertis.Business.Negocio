Public Class ClienteDireccion

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub
    Private Const cnEntidad As String = "tbClienteDireccion"

#End Region

#Region "Eventos RegisterAddNewTasks"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        data("IDDireccion") = AdminData.GetAutoNumeric
    End Sub

#End Region

#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim Obrl As New BusinessRules
        Obrl.Add("ENVIO", AddressOf CambioEnvio)
        Obrl.Add("FACTURA", AddressOf CambioFactura)
        Obrl.Add("GIRO", AddressOf CambioGiro)
        Obrl.Add("CodPostal", AddressOf CambioCodPostal)
        Return Obrl
    End Function

    <Task()> Public Shared Sub CambioEnvio(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If data.Value = False Then data.Current("PredeterminadaEnvio") = False
    End Sub

    <Task()> Public Shared Sub CambioFactura(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If data.Value = False Then data.Current("PredeterminadaFactura") = False
    End Sub

    <Task()> Public Shared Sub CambioGiro(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If data.Value = False Then data.Current("PredeterminadaGiro") = False
    End Sub

    <Task()> Public Shared Sub CambioCodPostal(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim infoCP As New CodPostalInfo(CStr(data.Value), data.Current("IDPais") & String.Empty)
            If Length(infoCP.DescPoblacion) > 0 Then
                data.Current("Poblacion") = infoCP.DescPoblacion
                'Else : data.Current("Poblacion") = String.Empty
            End If
            If Length(infoCP.DescProvincia) > 0 Then
                data.Current("Provincia") = infoCP.DescProvincia
                'Else : data.Current("Provincia") = String.Empty
            End If
            If Length(infoCP.IDPais) > 0 Then
                data.Current("IDPais") = infoCP.IDPais
                'Else : data.Current("IDPais") = String.Empty
            End If
            If Length(infoCP.DescPais) > 0 Then
                data.Current("DescPais") = infoCP.DescPais
                ' Else : data.Current("DescPais") = String.Empty
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarIDCliente)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDireccion)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarCIF)
    End Sub

    <Task()> Public Shared Sub ValidarIDCliente(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDCliente")) = 0 Then ApplicationService.GenerateError("El Cliente es un dato obligatorio.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarDireccion(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If (data.IsNull("Envio") OrElse data("Envio") = False) And _
                (data.IsNull("Factura") OrElse data("Factura") = False) And _
                (data.IsNull("Giro") OrElse data("Giro") = False) Then
                ApplicationService.GenerateError("La dirección del cliente debe ser de envío, de factura o de giro")
            End If
        ElseIf data.RowState = DataRowState.Modified Then
            If (data.IsNull("Envio") OrElse data("Envio") = False) And _
                                            (data.IsNull("Factura") OrElse data("Factura") = False) And _
                                            (data.IsNull("Giro") OrElse data("Giro") = False) Then
                ApplicationService.GenerateError("La dirección del cliente debe ser de envío, de factura o de giro")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarCIF(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Nz(data("DomicilioFiscal"), 0) Then
            If Length(data("CifCliente")) = 0 Then ApplicationService.GenerateError("El Cif de Cliente es un dato oligatorio para una dirección como Domicilio Fiscal.")
        End If
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarPrimaryKey)
        updateProcess.AddTask(Of DataRow)(AddressOf TratarPredeterminado)
    End Sub

    <Task()> Public Shared Sub AsignarPrimaryKey(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If IsDBNull(data("IDDireccion")) Then data("IDDireccion") = AdminData.GetAutoNumeric
        End If
    End Sub

    <Task()> Public Shared Sub TratarPredeterminado(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim f As New Filter(FilterUnionOperator.And)
        f.Add(New StringFilterItem("IDCliente", FilterOperator.Equal, data("IDCliente")))

        If Not data.IsNull("Envio") AndAlso data("Envio") = True Then
            Dim fEnvio As New Filter
            fEnvio.Add(f)
            fEnvio.Add(New BooleanFilterItem("Envio", FilterOperator.Equal, True))
            fEnvio.Add(New BooleanFilterItem("PredeterminadaEnvio", FilterOperator.Equal, True))
            Dim StPred As New ClienteDireccion.DataPredeterminada
            StPred.FilPred = fEnvio
            StPred.DtPred = data.Table
            StPred.DireccionPred = "PredeterminadaEnvio"
            ProcessServer.ExecuteTask(Of DataPredeterminada)(AddressOf ClienteDireccion.ActualizarPredeterminada, StPred, services)
        End If

        If Not data.IsNull("Factura") AndAlso data("Factura") = True Then
            Dim fFactura As New Filter
            fFactura.Add(f)
            fFactura.Add(New BooleanFilterItem("Factura", FilterOperator.Equal, True))
            fFactura.Add(New BooleanFilterItem("PredeterminadaFactura", FilterOperator.Equal, True))
            Dim StPred As New ClienteDireccion.DataPredeterminada
            StPred.FilPred = fFactura
            StPred.DtPred = data.Table
            StPred.DireccionPred = "PredeterminadaFactura"
            ProcessServer.ExecuteTask(Of DataPredeterminada)(AddressOf ClienteDireccion.ActualizarPredeterminada, StPred, services)
        End If

        If Not data.IsNull("Giro") AndAlso data("Giro") = True Then
            Dim fGiro As New Filter
            fGiro.Add(f)
            fGiro.Add(New BooleanFilterItem("Giro", FilterOperator.Equal, True))
            fGiro.Add(New BooleanFilterItem("PredeterminadaGiro", FilterOperator.Equal, True))
            Dim StPred As New ClienteDireccion.DataPredeterminada
            StPred.FilPred = fGiro
            StPred.DtPred = data.Table
            StPred.DireccionPred = "PredeterminadaGiro"
            ProcessServer.ExecuteTask(Of DataPredeterminada)(AddressOf ClienteDireccion.ActualizarPredeterminada, StPred, services)
        End If
    End Sub

    <Serializable()> _
    Public Class DataPredeterminada
        Public FilPred As Filter
        Public DtPred As DataTable
        Public DireccionPred As String
    End Class

    <Task()> Public Shared Sub ActualizarPredeterminada(ByVal data As DataPredeterminada, ByVal services As ServiceProvider)
        Dim dtCD As DataTable = New ClienteDireccion().Filter(data.FilPred)
        If IsNothing(dtCD) OrElse dtCD.Rows.Count = 0 Then
            ' No hay más IDDireccion de ese tipo dentro del cliente actual con lo cual será el predeterminado.
            data.DtPred.Rows(0)(data.DireccionPred) = True
        Else
            If IsDBNull(data.DtPred.Rows(0)(data.DireccionPred)) Then data.DtPred.Rows(0)(data.DireccionPred) = False
            ' Si IDDireccion ha sido marcado como predeterminado
            If data.DtPred.Rows(0)(data.DireccionPred) Then
                If data.DtPred.Rows(0)("IDDireccion") <> dtCD.Rows(0)("IDDireccion") Then
                    dtCD.Rows(0)(data.DireccionPred) = False
                    BusinessHelper.UpdateTable(dtCD)
                End If
            ElseIf data.DtPred.Rows(0).RowState = DataRowState.Modified AndAlso _
                data.DtPred.Rows(0)(data.DireccionPred) <> data.DtPred.Rows(0)(data.DireccionPred, DataRowVersion.Original) AndAlso _
                dtCD.Rows.Count = 1 Then
                data.DtPred.Rows(0)(data.DireccionPred) = True
            End If
        End If
    End Sub

#End Region

    <Serializable()> _
    Public Class dataDireccionAlquiler
        Public IDDireccion As Integer
        Public IDObra As Integer
        Public IDAlmacen As String

        Public Sub New(ByVal IDDireccion As Integer, ByVal IDObra As Integer, ByVal IDAlmacen As String)
            Me.IDDireccion = IDDireccion
            Me.IDObra = IDObra
            Me.IDAlmacen = IDAlmacen
        End Sub
    End Class

    <Task()> Public Shared Sub ActualizarEnDireccion(ByVal data As dataDireccionAlquiler, ByVal services As ServiceProvider)
        If data.IDDireccion > 0 AndAlso data.IDObra > 0 AndAlso Length(data.IDAlmacen) > 0 Then
            Dim dtCliDir As DataTable = New ClienteDireccion().SelOnPrimaryKey(data.IDDireccion)
            If Not dtCliDir Is Nothing AndAlso dtCliDir.Rows.Count > 0 Then
                If Length(dtCliDir.Rows(0)("IDObra")) = 0 Then dtCliDir.Rows(0)("IDObra") = data.IDObra
                If dtCliDir.Rows(0)("IDAlmacen") & String.Empty <> data.IDAlmacen Then
                    dtCliDir.Rows(0)("IDAlmacen") = data.IDAlmacen
                End If
                BusinessHelper.UpdateTable(dtCliDir)
            End If
        End If
    End Sub

    <Serializable()> _
    Public Class DataNuevaDireccion
        Public DtDirec As DataTable
        Public IDCliente As String
    End Class
    <Task()> Public Shared Sub NuevaDireccion(ByVal data As DataNuevaDireccion, ByVal services As ServiceProvider)
        Dim dtNewDireccion As DataTable = New ClienteDireccion().AddNewForm()
        If Not dtNewDireccion Is Nothing Then
            dtNewDireccion.Rows(0)("IDCliente") = data.IDCliente
            dtNewDireccion.Rows(0)("CodPostal") = data.DtDirec.Rows(0)("CodPostal")
            dtNewDireccion.Rows(0)("Direccion") = data.DtDirec.Rows(0)("Direccion")
            dtNewDireccion.Rows(0)("IDPais") = data.DtDirec.Rows(0)("IDPais")
            dtNewDireccion.Rows(0)("Poblacion") = data.DtDirec.Rows(0)("Poblacion")
            dtNewDireccion.Rows(0)("Provincia") = data.DtDirec.Rows(0)("Provincia")
            dtNewDireccion.Rows(0)("RazonSocial") = data.DtDirec.Rows(0)("RazonSocial")
            dtNewDireccion.Rows(0)("Envio") = 1
            dtNewDireccion.Rows(0)("Factura") = 1
            dtNewDireccion.Rows(0)("Giro") = 1
            BusinessHelper.UpdateTable(dtNewDireccion)
        End If
    End Sub

    <Serializable()> _
    Public Class DataDirecEnvio
        Public IDCliente As String
        Public TipoDireccion As enumcdTipoDireccion

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDCliente As String, ByVal TipoDireccion As enumcdTipoDireccion)
            Me.IDCliente = IDCliente
            Me.TipoDireccion = TipoDireccion
        End Sub
    End Class

    <Task()> Public Shared Function ObtenerDireccionEnvio(ByVal data As DataDirecEnvio, ByVal services As ServiceProvider) As DataTable
        Dim f As New Filter
        f.Add(New StringFilterItem("IDCliente", data.IDCliente))
        Select Case data.TipoDireccion
            Case enumcdTipoDireccion.cdDireccionEnvio
                f.Add(New BooleanFilterItem("Envio", True))
                f.Add(New BooleanFilterItem("PredeterminadaEnvio", True))
            Case enumcdTipoDireccion.cdDireccionFactura
                f.Add(New BooleanFilterItem("Factura", True))
                f.Add(New BooleanFilterItem("PredeterminadaFactura", True))
            Case enumcdTipoDireccion.cdDireccionGiro
                f.Add(New BooleanFilterItem("Giro", True))
                f.Add(New BooleanFilterItem("PredeterminadaGiro", True))
        End Select
        Dim dtDireccion As DataTable = New ClienteDireccion().Filter(f)
        If dtDireccion Is Nothing OrElse dtDireccion.Rows.Count = 0 Then
            If data.TipoDireccion <> enumcdTipoDireccion.cdDireccionGiro Then
                ApplicationService.GenerateError("Este Cliente no tiene una direccion predeterminada. Debe de crear una en el mantenimiento de Clientes.")
            End If
        End If
        Return dtDireccion
    End Function

    <Serializable()> _
    Public Class DataDirecDe
        Public IDDireccion As Integer
        Public TipoDireccion As enumcdTipoDireccion
        Public Sub New(ByVal IDDireccion As Integer, ByVal TipoDireccion As enumcdTipoDireccion)
            Me.IDDireccion = IDDireccion
            Me.TipoDireccion = TipoDireccion
        End Sub
    End Class
    <Task()> Public Shared Function EsDireccionDe(ByVal data As DataDirecDe, ByVal services As ServiceProvider) As Boolean
        Dim f As New Filter
        f.Add(New NumberFilterItem("IDDireccion", data.IDDireccion))
        Select Case data.TipoDireccion
            Case enumcdTipoDireccion.cdDireccionEnvio
                f.Add(New BooleanFilterItem("Envio", True))
            Case enumcdTipoDireccion.cdDireccionFactura
                f.Add(New BooleanFilterItem("Factura", True))
            Case enumcdTipoDireccion.cdDireccionGiro
                f.Add(New BooleanFilterItem("Giro", True))
        End Select
        Dim dtDireccion As DataTable = New ClienteDireccion().Filter(f)
        Return (Not dtDireccion Is Nothing AndAlso dtDireccion.Rows.Count > 0)
    End Function

End Class