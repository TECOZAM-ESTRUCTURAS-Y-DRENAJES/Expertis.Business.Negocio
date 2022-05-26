Public Class CentroGestAlmSuminist

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroCentroGestAlmSuminist"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarAlmacen)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As Object, ByVal services As ServiceProvider)
        If data("DescAlmacen").ToString.Trim.Length = 0 Then
            ApplicationService.GenerateError("Introduzca la descripci�n del almac�n")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarAlmacen(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            ' Que no sea ya un almac�n del centro (de los de la primera solapa)
            ' S�lo en filas nuevas porque en las actualizadas no se modifica el IDAlmacen.
            'Comprobar que el almac�n no est� gestionado por el Centro de Gesti�n
            Dim f As New Filter
            f.Add(New StringFilterItem("IDAlmacen", data("IDAlmacen")))
            f.Add(New StringFilterItem("IDCentroGestion", data("IDCentroGestion")))
            Dim DtAlmacen As DataTable = New Almacen().Filter(f)
            If Not DtAlmacen Is Nothing AndAlso DtAlmacen.Rows.Count > 0 Then
                ApplicationService.GenerateError("El Almac�n no puede sar dado de alta como Almac�n Suministrador porque es un Almac�n gestionado por el Centro de Gesti�n.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDAlmacen")) > 0 Then
                Dim DtTemp As DataTable = New CentroGestAlmSuminist().SelOnPrimaryKey(data("IDCentroGestion"), data("IDAlmacen"))
                If Not DtTemp Is Nothing AndAlso DtTemp.Rows.Count > 0 Then
                    ApplicationService.GenerateError("Ya ha introducido este almac�n suministrador. -")
                End If
            Else : ApplicationService.GenerateError("Introduzca el c�digo del almac�")
            End If
        End If
    End Sub

#End Region

End Class