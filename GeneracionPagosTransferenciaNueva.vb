Imports Microsoft.VisualBasic

Imports Solmicro.Expertis.Engine.UI
Imports Solmicro.Expertis.Engine.BE.BusinessHelper
Imports Solmicro.Expertis.Business.SEPA
Imports Solmicro.Expertis.Application.ERP.CommonClasses

Public Class GeneracionPagosTransferenciaNueva

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public mIDProcess As System.Guid
    Private mblnDesmarcar As Boolean

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "vSEPA_frmPagosGenerarFichero_34_14_Pago"
    'Pasa la tabla de las filas a realizar la transferencia
    Public Function filtroSEPA(ByVal dt As DataTable) As DataTable
        Dim f As New Filter
        Dim f2 As New Filter
        Dim frm As New GeneracionPagosTransferencia
        Dim IDPago As String
        f2.Add("IDPago", FilterOperator.Equal, "0")
        Dim dtSEPA As DataTable = AdminData.GetData("vSEPA_frmPagosGenerarFichero_34_14_Pago", f2) 'PAra crear la tabla vacia
        Dim dtSEPA2 As DataTable

        'Dim frm As New GeneracionPagosTransferencia()
        'ExpertisApp.GenerateMessage("Número de lineas Origen: " & dt.Rows.Count, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        For Each dr As DataRow In dt.Rows
            f.Add("IDPago", FilterOperator.Equal, dr("IDPago"))
            dtSEPA2 = AdminData.GetData("vSEPA_frmPagosGenerarFichero_34_14_Pago", f)
            dtSEPA.Merge(dtSEPA2)
            'ExpertisApp.GenerateMessage("IDPago" & dr("IDPago"))
            f.Clear()
        Next

        'ExpertisApp.GenerateMessage("Número de lineas FIN: " & dtSEPA.Rows.Count, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        Return dtSEPA
    End Function

    Public Function GenerarFicheroTransferencia(ByVal strBanco As String, ByVal ruta As String, ByVal dtSEPA As DataTable)

        Dim rcsFichero As DataTable
        Dim strRuta As String

        If Len(strBanco) Then
            'Ibis. Guille 2016/01/26 SEPA xml
            'SaveFileDialog1.DefaultExt = ".N34"
            strRuta = ruta

            If Len(strRuta & vbNullString) Then
                'cambiado

                Dim dtFichero As New DataTable
                dtFichero.Columns.Add("Linea", GetType(String))

                Dim htPagos As New Hashtable
                htPagos.Clear()
                '//Marcamos en el servidor los pagos
                Dim i As Integer = 0
                For Each drRowMarcados As DataRow In dtSEPA.Rows
                    htPagos("IDEnlace" & i) = drRowMarcados("IdPago")
                    i = i + 1
                Next drRowMarcados

                mIDProcess = MarcarRegistro(htPagos, FilterType.Numeric)
                mblnDesmarcar = True
                htPagos = Nothing
                'rcsFichero = GenerarFicheroTransferencia3SEPA(Me.ProgramInfo.Alias, ExpertisApp.UserName, strBanco)
                Dim datFich As New DataGenerarFichero
                datFich.IDProcess = mIDProcess
                datFich.IDBancoPropio = strBanco
                datFich.FechaEmision = Today

                Dim lstRegsFich As Byte() 'entre 0 y 255
                Dim ClsFichCSB As New Fichero_PAIN_001_001_03
                lstRegsFich = ClsFichCSB.GenerarFichero(datFich, New ServiceProvider())
                If Not lstRegsFich Is Nothing AndAlso lstRegsFich.Length > 0 AndAlso Expertis.Application.ERP.SEPA.General.GuardarFicheroXML(lstRegsFich, strRuta, -1, datFich.FechaCargo, True) Then
                    ExpertisApp.GenerateMessage("Fichero generado.", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Information)
                    DesmarcarRegistro(mIDProcess, FilterType.Numeric)
                    Exit Function
                End If

            End If
        End If

    End Function

    Public Function GenerarFicheroTransferencia3SEPA(ByVal strIDPrograma As String, ByVal pIdMaquinaUsuario As String) As DataTable
        Dim datFich As New DataGenerarFichero
        Dim DteFechaEmision As Date
        Dim DtFichero As DataTable
        Dim service As ServiceProvider
        Dim datFichero As DataAddRegistroFichero
        DteFechaEmision = Today
        datFich.IDProcess = mIDProcess
        datFich.FechaEmision = DteFechaEmision
        ProcessServer.ExecuteTask(Of DataGenerarFichero, Byte())(AddressOf Fichero_PAIN_001_001_03.GenerarFichero, datFich, service)

    End Function

    Public Function GuardarFichero(ByVal StrRuta As String, ByVal DtFichero As DataTable) As Boolean
        If Len(StrRuta) > 0 Then
            Dim IntPos As Integer = 1
            Dim IntPosPar As Integer = Strings.InStr(IntPos, StrRuta, "\")
            While IntPosPar <> 0
                IntPos = IntPosPar + 1
                IntPosPar = InStr(IntPos, StrRuta, "\")
            End While
            Dim StrRutaFinal As String = Strings.Left(StrRuta, IntPos - 2)
            If IO.Directory.Exists(StrRutaFinal) Then
                'If IO.File.Exists(StrRuta) Then IO.File.Delete(StrRuta)
                'Dim FichDest As New IO.StreamWriter(StrRuta, False)
                'For Each Dr As DataRow In DtFichero.Select
                '    FichDest.WriteLine(Dr("Linea"))
                'Next
                'FichDest.Close()
                'Return True

                'Abrimos el fichero
                Dim intFile As Integer
                intFile = FreeFile()

                FileOpen(intFile, StrRuta, OpenMode.Output, OpenAccess.Write, OpenShare.LockReadWrite)

                With DtFichero
                    For Each Dr As DataRow In DtFichero.Select
                        Print(intFile, Dr("Linea"))
                        Print(intFile, vbNewLine)
                    Next
                End With

                If Not IsNothing(DtFichero) Then DtFichero.Dispose()
                FileClose(intFile)
                Return True
            Else
                ExpertisApp.GenerateMessage("La ruta | no existe. Debe crear la ruta para generar el fichero.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, StrRuta)
            End If
        Else
            ExpertisApp.GenerateMessage("La ruta | no existe. Debe crear la ruta para generar el fichero.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, StrRuta)
        End If

        'ExpertisApp.GenerateMessage("Ya estaría", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    End Function

End Class
