Public Class AmortizacionRegistro

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbAmortizacionRegistro"

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
    End Sub

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDAmortizacionRegistro")) = 0 Then data("IDAmortizacionRegistro") = AdminData.GetAutoNumeric
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function InsertAmortizacionRegistro(ByVal data As DataRow, ByVal services As ServiceProvider) As DataTable
        Dim dtResultado As DataTable
        Dim lngDecimalesA As Integer
        Dim lngDecimalesB As Integer
        Dim dblCambioBdeA As Double

        Dim dtAmortReg As DataTable


        'Comienzo del Cuerpo de la Función
        If Not data Is Nothing Then
            'DATOS MONEDA...
            Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
            Dim MonInfoA As MonedaInfo = Monedas.MonedaA
            lngDecimalesA = MonInfoA.NDecimalesImporte
            dblCambioBdeA = MonInfoA.CambioB

            Dim MonInfoB As MonedaInfo = Monedas.MonedaB
            lngDecimalesB = MonInfoB.NDecimalesImporte

            If data.IsNull("AmortizacionAutomatica") OrElse data("AmortizacionAutomatica") = False Then
                'Elimina los registros que ya existen para ese elemento
                dtResultado = New AmortizacionRegistro().Filter(New StringFilterItem("IdElemento", data("IdElemento").ToString()))
                If Not dtResultado Is Nothing AndAlso dtResultado.Rows.Count > 0 Then
                    Dim ClsAmort As New AmortizacionRegistro
                    ClsAmort.Delete(dtResultado)
                End If
            End If

            Dim valorAmortizado As Double
            Dim valorAmortizadoPlusvalia As Double
            If data.IsNull("ValorAmortizadoElementoA") Then
                valorAmortizado = 0
            Else
                If data.RowState <> DataRowState.Added AndAlso Nz(data("AmortizacionAutomatica"), False) Then
                    valorAmortizado = data("ValorAmortizadoElementoA") - data("ValorAmortizadoElementoA", DataRowVersion.Original)
                Else
                    valorAmortizado = data("ValorAmortizadoElementoA")
                End If
            End If
            If data.IsNull("ValorAmortizadoPlusvaliaA") Then
                valorAmortizadoPlusvalia = 0
            Else
                valorAmortizadoPlusvalia = data("ValorAmortizadoPlusvaliaA")
            End If

            If Nz(data("ValorAmortizadoElementoA"), 0) > 0 Or valorAmortizadoPlusvalia > 0 Then
                'Inserta el nuevo registro AmortizacionRegistro
                dtAmortReg = New AmortizacionRegistro().AddNew
                Dim drNew As DataRow = dtAmortReg.NewRow
                drNew("IDAmortizacionRegistro") = AdminData.GetAutoNumeric
                drNew("IdElemento") = data("IdElemento")
                drNew("IDGrupoAmortizacion") = data("IDGrupoAmortizacion")
                drNew("MesContabilizacion") = Month(data("FechaUltimaContabilizacion"))
                drNew("AñoContabilizacion") = Year(data("FechaUltimaContabilizacion"))
                drNew("FechaContabilizacion") = data("FechaUltimaContabilizacion")
                drNew("ValorAmortizadoA") = xRound(valorAmortizado, lngDecimalesA)
                drNew("ValorAmortizadoB") = xRound(valorAmortizado * dblCambioBdeA, lngDecimalesB)
                drNew("ValorAmortizadoPlusvaliaA") = xRound(data("ValorAmortizadoPlusvaliaA"), lngDecimalesA)
                drNew("ValorAmortizadoPlusvaliaB") = xRound(data("ValorAmortizadoPlusvaliaA") * dblCambioBdeA, lngDecimalesB)
                dtAmortReg.Rows.Add(drNew)
                'If dtResultado Is Nothing Then
                '    Return dtAmortReg
                'Else
                '    dtResultado.ImportRow(dtAmortReg.Rows(0))
                '    Return dtResultado
                'End If
                Return dtAmortReg
            Else : Return dtResultado
            End If
        End If
    End Function

#End Region

End Class