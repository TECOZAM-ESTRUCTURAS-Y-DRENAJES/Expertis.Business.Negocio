Option Strict Off
Option Explicit On
Option Compare Text

Imports Solmicro.Expertis.Business.Financiero

Public Class OrdenTrabajoCabecera
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbOrdenTrabajoCabecera"

    Public Overloads Sub Delete(ByVal strIDTrabajo As Integer)
        ' Error al borrar, tener encuenta que el id valor pasado sea del mismo tipo
        ' Luego desde el admin Actualizar entidades.

    End Sub

    Public Overloads Function Update(ByVal dttSource As System.Data.DataTable) As System.Data.DataTable

        Dim services As ServiceProvider

        If Not dttSource Is Nothing AndAlso dttSource.Rows.Count > 0 Then
            Me.BeginTx()
            For Each dr As DataRow In dttSource.Rows
                If Length(dr("Fentrada")) = 0 Then ApplicationService.GenerateError("Debe de indicar una fecha de creación de la Orden.")
                If Length(dr("IDCliente")) = 0 Then ApplicationService.GenerateError("Debe de indicar un cliente.")
                ' Actualizar el contador por el IDContador seleccionado
                If Not IsDBNull(dr("IDContador")) And Not IsDBNull(dr("tipopago")) Then
                    dr("NTrabajo") = Contador.CounterValue(dr("IDContador"), services)
                End If
                ' Ojo utilizo tipopago para controlar si actualizo o no el contador
                dr("tipopago") = DBNull.Value
                If dr.RowState = DataRowState.Added Then


                    ''Comprobación de la existencia de la Prevision
                    'dtTarifa = SelOnPrimaryKey(dr("IDPrevision"))
                    'If dtTarifa.Rows.Count <> 0 Then GenerateMessage("La Previsión ya existe", Me.GetType.Name & ".Update")

                End If
            Next
            AdminData.SetData(dttSource)
        End If
        Return dttSource
    End Function

    Public Overrides Function AddNewForm() As DataTable
        Dim dt As DataTable = MyBase.AddNewForm
        Dim services As ServiceProvider
        'Contadores ------------------------------------------------------------------------------------------------------

        Dim c As Contador
        Dim dtContadores As DataTable
        Dim DtContadorPred As DataTable = Contador.CounterDefault("OrdenTrabajoCabecera", services)
        If Not DtContadorPred Is Nothing AndAlso DtContadorPred.Rows.Count > 0 Then
            dt.Rows(0)("IDContador") = DtContadorPred.Rows(0)("IDContador")
            dt.Rows(0)("NTrabajo") = DtContadorPred.Rows(0)("Contador")
            ' Ojo utilizo esta columna para controlar si act. contador o no.
            dt.Rows(0)("tipopago") = 1
        End If

        dt.Rows(0)("idOrdenTrabajo") = AdminData.GetAutoNumeric
        'Saco el Contador que Corresponda
        '----------------------------------------
        dt.Rows(0)("Fentrada") = Now



        Return dt

    End Function

#Region "Creación de ordenes desde mediciones"
    Public Function CrearOrdenMediciones(ByVal IDObra As Integer, ByVal IDCliente As String, ByVal consulta As String) As Integer
        Dim blnExiste As Boolean = False
        Dim obj As New OrdenTrabajoCabecera
        Dim EjercPredet As New EjercicioContable
        'Cliente ----------------------------------------------------------------------------------------------------------
        Dim dtObra As New DataTable
        Dim objObra As New Obra.ObraCabecera
        Dim fObra As New Filter
        Dim services As ServiceProvider

        Try
            If IDCliente = "" Then
                MessageBox.Show("Debe indicar un Cliente." & Chr(13), "No puede continuar", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return 0
            End If

            '--------------------------------------------------------------------------------------------------------------

            Dim dt As DataTable = AddNewForm()

            'dt.Rows(0)("idOrdenTrabajo") = AdminData.GetAutoNumeric
            'dt.Rows(0)("IDContador") = adr(0)("IDContador")
            'dt.Rows(0)("NTrabajo") = Contador.CounterValue(dt.Rows(0)("IDContador"), Me, "NTrabajo", "Fentrada", Nz(dt.Rows(0)("Fentrada"), Date.Today))
            MessageBox.Show("AQUI LLEGA")
            dt.Rows(0)("IDObra") = IDObra
            ' Datos de obra select
            MessageBox.Show("AQUI LLEGA 2")
            fObra.Add("IDObra", IDObra)
            dtObra = objObra.Filter(fObra)
            If Not IsNothing(dtObra) AndAlso dtObra.Rows.Count > 0 Then
                dt.Rows(0)("DescObra") = dtObra.Rows(0)("DescObra")
            End If
            ' Coger el contador predeterminado y no actualzar en el update
            If Not IsDBNull(dt.Rows(0)("IDContador")) Then
                dt.Rows(0)("NTrabajo") = Contador.CounterValue(dt.Rows(0)("IDContador"), services)
                dt.Rows(0)("tipopago") = DBNull.Value
            End If
            ' Borrar datos de obra
            dtObra = Nothing
            objObra = Nothing
            fObra = Nothing
            dt.Rows(0)("IDCliente") = IDCliente
            dt.Rows(0)("IDEjercicio") = EjercicioContable.Predeterminado(Date.Today, services)
            dt.Rows(0)("IDCliente") = IDCliente
            dt.Rows(0)("Fentrada") = Today
            dt.Rows(0)("comentario") = "LOS PAQUETES NO SUPERAN LOS 3.000 KGS"
            MyBase.Update(dt)

            'Creamos las Lineas 
            If consulta.Trim <> "" Then
                CrearLineas(dt.Rows(0)("idOrdenTrabajo"), consulta)
            End If
            ' Retornar el Id de orden
            Return dt.Rows(0)("idOrdenTrabajo")

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            EjercPredet = Nothing
        End Try

    End Function

    Public Sub CrearLineas(ByVal idOrden As Integer, ByVal consulta As String)
        Dim dt As New DataTable
        Dim dtMedicion As New DataTable
        Dim obj As New OrdenTrabajodet
        Try
            dtMedicion = AdminData.Filter("tbObraMedicionAcero", , "IDLineaMedicionA IN " & consulta)

            For Each dr As DataRow In dtMedicion.Rows
                dt = obj.AddNewForm
                dt.Rows(0)("idOrdenTrabajo") = idOrden
                dt.Rows(0)("idLineaMedicionA") = dr("IDLineaMedicionA")
                dt.Rows(0)("estructura") = dr("Estructura")
                dt.Rows(0)("localizacion1") = dr("Localizacion1")
                dt.Rows(0)("localizacion2") = dr("Localizacion2")
                dt.Rows(0)("FechaProduccion") = dr("fproduccion")
                dt.Rows(0)("Numpedido") = dr("numPedido")
                dt.Rows(0)("kg") = dr("PesoPlanilla")
                '-----------------------------------------------------------------------------------------------------------------
                obj.Update(dt)
            Next


        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            obj = Nothing
            dt = Nothing
        End Try
    End Sub

#End Region

End Class
