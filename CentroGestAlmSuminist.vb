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
            ApplicationService.GenerateError("Introduzca la descripción del almacén")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarAlmacen(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            ' Que no sea ya un almacén del centro (de los de la primera solapa)
            ' Sólo en filas nuevas porque en las actualizadas no se modifica el IDAlmacen.
            'Comprobar que el almacén no está gestionado por el Centro de Gestión
            Dim f As New Filter
            f.Add(New StringFilterItem("IDAlmacen", data("IDAlmacen")))
            f.Add(New StringFilterItem("IDCentroGestion", data("IDCentroGestion")))
            Dim DtAlmacen As DataTable = New Almacen().Filter(f)
            If Not DtAlmacen Is Nothing AndAlso DtAlmacen.Rows.Count > 0 Then
                ApplicationService.GenerateError("El Almacén no puede sar dado de alta como Almacén Suministrador porque es un Almacén gestionado por el Centro de Gestión.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDAlmacen")) > 0 Then
                Dim DtTemp As DataTable = New CentroGestAlmSuminist().SelOnPrimaryKey(data("IDCentroGestion"), data("IDAlmacen"))
                If Not DtTemp Is Nothing AndAlso DtTemp.Rows.Count > 0 Then
                    ApplicationService.GenerateError("Ya ha introducido este almacén suministrador. -")
                End If
            Else : ApplicationService.GenerateError("Introduzca el código del almacé")
            End If
        End If
    End Sub

#End Region

End Class