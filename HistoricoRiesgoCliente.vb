Public Class HistoricoRiesgoCliente
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub
    Private Const cnEntidad As String = "tbHistoricoRiesgoCliente"

#Region " RegisterAddnewTasks "

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
    End Sub

#End Region

#Region " RegisterValidateTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarRegistroExistente)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        ' If Length(data("IDCliente")) = 0 Then ApplicationService.GenerateError("El Cliente es un dato obligatorio.")
        If Length(data("CIFCliente")) = 0 Then ApplicationService.GenerateError("El CIF Cliente es un dato obligatorio.")
        'If Nz(data("FechaSolicitud"), cnMinDate) = cnMinDate Then ApplicationService.GenerateError("La Fecha Solicitud es un dato obligatorio.")
        If Length(data("IDAseguradora")) = 0 Then ApplicationService.GenerateError("La aseguradora es un dato obligatorio.")
        If Length(data("NumPoliza")) = 0 Then ApplicationService.GenerateError("El Nº de Póliza es un dato obligatorio.")
        If Nz(data("FechaConcedido"), cnMinDate) <> cnMinDate Then
            If Nz(data("ImporteConcedido")) = 0 AndAlso Length(data("IDMotivoNoAsegurado")) = 0 Then
                ApplicationService.GenerateError("Debe indicar el Importe Concedido o el Motivo No Asegurado.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarRegistroExistente(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added AndAlso Nz(data("ID")) <> 0 Then
            Dim dt As DataTable = New HistoricoRiesgoCliente().SelOnPrimaryKey(data("ID"))
            If dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("Ya existe el Resgistro indicado en el sistema.")
            End If
        End If
    End Sub

#End Region

#Region " RegisterUpdateTasks "

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
        updateProcess.AddTask(Of DataRow)(AddressOf Comunes.UpdateEntityRow)
        updateProcess.AddTask(Of DataRow)(AddressOf Comunes.MarcarComoActualizado)
        updateProcess.AddTask(Of DataRow)(AddressOf ActualizarDatosClientes)
    End Sub

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("ID")) = 0 Then data("ID") = AdminData.GetAutoNumeric
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarDatosClientes(ByVal data As DataRow, ByVal services As ServiceProvider)

        Dim c As New Cliente
        Dim dtClientesCIF As DataTable
        If Length(data("CIFCliente")) > 0 Then
            dtClientesCIF = c.Filter(New StringFilterItem("CIFCliente", data("CIFCliente")))
        End If
        If dtClientesCIF.Rows.Count > 0 Then

            For Each drCliente As DataRow In dtClientesCIF.Rows
                Dim IDAseguradora As String = String.Empty

                Dim ImporteConcedido As Double = 0
                Dim NumPoliza As String = String.Empty
                Dim IDMotivoNoAsegurado As String '= cboMotivoNoAsegurado.Value & String.Empty
                Dim fCliente As New Filter
                fCliente.Add(New StringFilterItem("CIFCliente", data("CIFCliente")))
                Dim dtHistorico As DataTable = New HistoricoRiesgoCliente().Filter(fCliente, "FechaConcedido DESC", "TOP 2 *")
                If dtHistorico.Rows.Count > 0 Then

                    ImporteConcedido = dtHistorico.Rows(0)("ImporteConcedido")
                    IDMotivoNoAsegurado = dtHistorico.Rows(0)("IDMotivoNoAsegurado") & String.Empty
                    IDAseguradora = dtHistorico.Rows(0)("IDAseguradora") & String.Empty
                    NumPoliza = dtHistorico.Rows(0)("NumPoliza") & String.Empty
                    If ImporteConcedido = 0 AndAlso Nz(dtHistorico.Rows(0)("FechaConcedido"), cnMinDate) = cnMinDate AndAlso dtHistorico.Rows.Count > 1 Then
                        ImporteConcedido = dtHistorico.Rows(1)("ImporteConcedido")
                        IDMotivoNoAsegurado = dtHistorico.Rows(1)("IDMotivoNoAsegurado") & String.Empty
                        IDAseguradora = dtHistorico.Rows(1)("IDAseguradora") & String.Empty
                        NumPoliza = dtHistorico.Rows(1)("NumPoliza") & String.Empty
                    End If
                End If
                If Length(IDAseguradora) > 0 Then
                    drCliente("IDAseguradora") = IDAseguradora
                Else
                    drCliente("IDAseguradora") = System.DBNull.Value
                End If
                If Length(NumPoliza) > 0 Then
                    drCliente("NPolizaAseg") = NumPoliza
                Else
                    drCliente("NPolizaAseg") = System.DBNull.Value
                End If
                drCliente = c.ApplyBusinessRule("LimiteCapitalAsegurado", ImporteConcedido, drCliente, Nothing)
                If Length(IDMotivoNoAsegurado) > 0 Then
                    drCliente("IDMotivoNoAsegurado") = IDMotivoNoAsegurado
                Else
                    drCliente("IDMotivoNoAsegurado") = System.DBNull.Value
                End If
            Next

        End If

        c.Update(New UpdatePackage(dtClientesCIF))

    End Sub

#End Region

#Region " RegisterDeleteTasks "

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf Comunes.DeleteEntityRow)
        deleteProcess.AddTask(Of DataRow)(AddressOf Comunes.MarcarComoEliminado)
        deleteProcess.AddTask(Of DataRow)(AddressOf ActualizarDatosClientes)
    End Sub

#End Region


End Class
