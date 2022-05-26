Public Class TarifaOrden

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbTarifaOrden"

#End Region

    '#Region "Eventos RegisterDeleteTasks"

    '    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
    '        MyBase.RegisterDeleteTasks(deleteProcess)
    '        deleteProcess.AddTask(Of DataRow)(AddressOf Reordenar)
    '    End Sub



    '#End Region

#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("IDTarifa", AddressOf ComprobarTarifa)
        Return oBrl
    End Function

    <Task()> Public Shared Sub ComprobarTarifa(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) = 0 Then
            ApplicationService.GenerateError("Hay que introducir un identificativo de Tarifa.")
        Else
            'Se comprueba que la tarifa exista en el maestro de tarifas.
            Dim ClsTarifa As New Tarifa
            Dim DtTarifa As DataTable = ClsTarifa.SelOnPrimaryKey(data.Value)
            If DtTarifa.Rows.Count = 0 Then
                ApplicationService.GenerateError("La tarifa introducida no existe.")
            Else : data.Current("DescTarifa") = DtTarifa.Rows(0)("DescTarifa")
            End If
            'Se comprueba que la tarifa no ha sido introducida ya en las tarifas en curso.
            Dim DtTarifaOrden As DataTable = New TarifaOrden().SelOnPrimaryKey(data.Value)
            If Not DtTarifaOrden Is Nothing AndAlso DtTarifaOrden.Rows.Count > 0 Then
                ApplicationService.GenerateError("La Nueva Tarifa ya existe. Introduzca otra.")
                data.Current("DescTarifa") = String.Empty
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf RecuperarUltimoOrden)
        updateProcess.AddTask(Of DataRow)(AddressOf ValidarOrdenExistente)
    End Sub

    <Task()> Public Shared Sub RecuperarUltimoOrden(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added OrElse data.RowState = DataRowState.Modified Then
            If Length(data("Orden")) <= 0 Then
                Dim DtOrden As DataTable = New TarifaOrden().Filter(, , "Orden DESC")
                If Not DtOrden Is Nothing AndAlso DtOrden.Rows.Count > 0 Then
                    data("Orden") = DtOrden.Rows(0)("Orden") + 1
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarOrdenExistente(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added OrElse (data.RowState = DataRowState.Modified AndAlso (data("Orden") & String.Empty <> data("Orden", DataRowVersion.Original) & String.Empty)) Then
            If Length(data("Orden")) > 0 Then
                Dim f As New Filter
                f.Add(New NumberFilterItem("Orden", data("Orden")))
                f.Add(New StringFilterItem("IDTarifa", FilterOperator.NotEqual, data("IDTarifa")))
                Dim dtOrden As DataTable = New TarifaOrden().Filter(f)
                If dtOrden.Rows.Count > 0 Then
                    ApplicationService.GenerateError("El Orden indicado para la Tarifa {0} ya existe en el sistema.", Quoted(data("IDTarifa")))
                End If
            End If
        End If
    End Sub


#End Region

End Class