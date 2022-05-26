Public Class AlmacenUbicacion
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroAlmacenUbicacion"


#Region " Delete "

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf DeletePredeterminada)
    End Sub

    <Task()> Public Shared Sub DeletePredeterminada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data("Predeterminada") Then
            Dim dt As DataTable = New AlmacenUbicacion().Filter(New StringFilterItem("IDAlmacen", data("IDAlmacen")))
            If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
                dt.Rows(0)("Predeterminada") = True

                BusinessHelper.UpdateTable(dt)
            End If
        End If
    End Sub

#End Region

#Region " Validate "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarAlmacenObligatorio)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarUbicacionObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarAlmacenUbicacionExistente)
    End Sub

    <Task()> Public Shared Sub ValidarUbicacionObligatoria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDUbicacion")) = 0 Then ApplicationService.GenerateError("Introduzca el código de la ubicación.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarAlmacenUbicacionExistente(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New AlmacenUbicacion().SelOnPrimaryKey(data("IDAlmacen"), data("IDUbicacion"))
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("Ya existe una Ubicación con esa clave para el Almacén actual.")
            End If
        End If
    End Sub
#End Region

#Region " Update "

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.BeginTransaction)
        updateProcess.AddTask(Of DataRow)(AddressOf TratarPredeterminada)
    End Sub

    <Task()> Public Shared Sub TratarPredeterminada(ByVal dr As DataRow, ByVal services As ServiceProvider)
        Dim f As New Filter
        f.Add(New StringFilterItem("IDAlmacen", dr("IDAlmacen")))
        f.Add(New BooleanFilterItem("Predeterminada", True))
        Dim dtAU As DataTable = New AlmacenUbicacion().Filter(f)

        If IsNothing(dtAU) OrElse dtAU.Rows.Count = 0 Then
            ' No hay más IDUbicacion dentro del almacen actual con lo cual será el predeterminado.
            dr("Predeterminada") = True
        Else
            If IsDBNull(dr("Predeterminada")) Then dr("Predeterminada") = False
            ' Si IDUbicacion ha sido marcado como predeterminado
            If dr("Predeterminada") Then
                If dr("IDUbicacion") <> dtAU.Rows(0)("IDUbicacion") Then
                    dtAU.Rows(0)("Predeterminada") = False
                    BusinessHelper.UpdateTable(dtAU)
                End If
            ElseIf dr.RowState = DataRowState.Modified AndAlso dr("Predeterminada") <> dr("Predeterminada", DataRowVersion.Original) AndAlso dtAU.Rows.Count = 1 Then
                dr("Predeterminada") = True
            End If
        End If
    End Sub

#End Region

    <Task()> Public Shared Sub NuevaUbicacion(ByVal strIDAlmacen As String, ByVal services As ServiceProvider)
        Dim info As Parametro.infoUbicacion = New Parametro().UbicacionNoDefinida
        If Length(info.IDUbicacion) Then
            Dim AU As New AlmacenUbicacion
            Dim dt As DataTable = AU.AddNewForm
            dt.Rows(0)("IDAlmacen") = strIDAlmacen
            dt.Rows(0)("IDUbicacion") = info.IDUbicacion
            dt.Rows(0)("DescUbicacion") = info.DescUbicacion
            dt.Rows(0)("Predeterminada") = True

            AU.Update(dt)
        End If
    End Sub

    <Task()> Public Shared Function UbicacionPredeterminada(ByVal IDAlmacen As String, ByVal services As ServiceProvider) As DataTable
        If Length(IDAlmacen) > 0 Then
            Dim f As New Filter
            f.Add(New StringFilterItem("IDAlmacen", IDAlmacen))
            f.Add(New BooleanFilterItem("Predeterminada", True))
            Dim AU As New AlmacenUbicacion
            Dim ubicacion As DataTable = AU.Filter(f)
            If ubicacion.Rows.Count > 0 Then
                Return ubicacion
            Else
                Dim info As Parametro.infoUbicacion = New Parametro().UbicacionNoDefinida
                ubicacion = AU.AddNewForm
                ubicacion.Rows(0)("IDUbicacion") = info.IDUbicacion
                ubicacion.Rows(0)("DescUbicacion") = info.DescUbicacion
            End If
            Return ubicacion
        End If
    End Function

End Class