Public Class ArticuloNumeroSerie
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub
    Private Const cnEntidad As String = "tbArticuloNumeroSerie"

    Public Overloads Sub Delete(ByVal strIDArticulo As String, ByVal strNSerie As String)

        If Not DeleteSide(strIDArticulo, strNSerie) Then
            ApplicationService.GenerateError(DELETECONSTRAINTMESSAGE)
        Else

        End If
        Exit Sub
    End Sub

    Public Function DeleteSide(ByVal strIDArticulo As String, ByVal strNSerie As String)

        Delete(strIDArticulo, strNSerie)

    End Function
    Public Function UltimoDocCompra(ByVal strIDArticulo As String, ByVal strNSerie As String) As Long

        Dim rscReferencia As New DataTable
        Try
            rscReferencia = AdminData.Filter("tbArticuloNumeroSerie", "UltimoDocCompra", "IDArticulo = '" & strIDArticulo & "' and NSerie = '" & strNSerie & "'")
            UltimoDocCompra = 0
            If Not rscReferencia Is Nothing Then
                If rscReferencia.Rows.Count > 0 Then
                    UltimoDocCompra = Nz(rscReferencia("UltimoDocCompra"), 0)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            rscReferencia = Nothing
        End Try
    End Function


    Public Function UltimoDocVenta(ByVal strIDArticulo As String, ByVal strNSerie As String) As Long

        Dim rscReferencia As New DataTable

        Try
            rscReferencia = AdminData.Filter("tbArticuloNumeroSerie", "UltimoDocVenta", "IDArticulo = '" & strIDArticulo & "' and NSerie = '" & strNSerie & "'")

            UltimoDocVenta = 0
            If Not rscReferencia Is Nothing Then
                If rscReferencia.Rows.Count > 0 Then
                    UltimoDocVenta = Nz(rscReferencia("UltimoDocVenta"), 0)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            rscReferencia = Nothing
        End Try

    End Function

    Public Function StockNSerie(ByVal strIDArticulo As String, ByVal strNSerie As String) As Double
        Dim rscReferencia As New DataTable

        Try
            rscReferencia = AdminData.Filter("tbArticuloNumeroSerie", "Stock", "IDArticulo = '" & strIDArticulo & "' and NSerie = '" & strNSerie & "'")

            StockNSerie = 0
            If Not rscReferencia Is Nothing Then
                If rscReferencia.Rows.Count > 0 Then
                    StockNSerie = Nz(rscReferencia("Stock"), 0)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            rscReferencia = Nothing
        End Try
    End Function

    Public Function ModificarNSerie(ByVal strIDArticulo As String, ByVal strNSerie As String, ByVal strNuevoNumero As String) As Boolean
        Dim rscReferencia As DataTable
        Dim vSql As String

        Try
            rscReferencia = AdminData.Filter("tbArticuloNumeroSerie", "NSerie", "IDArticulo = '" & strIDArticulo & "' and NSerie = '" & strNSerie & "'")
            ModificarNSerie = False
            ModificarNSerie = 0
            If Not rscReferencia Is Nothing Then
                If rscReferencia.Rows.Count > 0 Then
                    If Nz(rscReferencia("UltimoDocVenta"), 0) = 0 Then
                        rscReferencia(0)("NSerie") = strNuevoNumero
                        vSql = "UPDATE tbArticuloNumeroSerie SET NSerie = '" & strNuevoNumero & "' WHERE IDArticulo = '" & strIDArticulo & "' and NSerie = '" & strNSerie & "'"
                        AdminData.Execute(vSql)
                        ModificarNSerie = True
                    End If
                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            rscReferencia = Nothing
        End Try

    End Function

    Public Function SePuedeComprarNSerie(ByVal strIDArticulo As String, ByVal strNSerie As String, ByVal Abono As Boolean) As Boolean
        Dim rscReferencia As New DataTable

        Try
            rscReferencia = AdminData.Filter("tbArticuloNumeroSerie", "Stock", "IDArticulo = '" & strIDArticulo & "' and NSerie = '" & strNSerie & "'")
            SePuedeComprarNSerie = True

            If Not rscReferencia Is Nothing Then
                If rscReferencia.Rows.Count > 0 Then
                    If Abono Then
                        If Nz(rscReferencia("Stock"), 0) <= 0 Then
                            SePuedeComprarNSerie = False
                        End If
                    Else
                        If Nz(rscReferencia("Stock"), 0) > 0 Then
                            SePuedeComprarNSerie = False
                        End If
                    End If
                Else
                    'David, pq si le pasamos cantidades negativas, se cree q no existen referencias
                    'If Abono Then SePuedeComprarNSerie = False
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            rscReferencia = Nothing
        End Try

    End Function

    Public Function SePuedeVenderNSerie(ByVal strIDArticulo As String, ByVal strNSerie As String, ByVal Abono As Boolean, ByRef Mensaje As String) As Boolean
        Dim rscReferencia As DataTable

        Try
            rscReferencia = AdminData.Filter("tbArticuloNumeroSerie", "Stock,UltimoDocVenta", "IDArticulo = '" & strIDArticulo & "' and NSerie = '" & strNSerie & "'")
            SePuedeVenderNSerie = True

            If Not rscReferencia Is Nothing Then
                If rscReferencia.Rows.Count > 0 Then
                    If Abono Then
                        'If Nz(rscReferencia("Stock"), 0) > 0 Then
                        '    SePuedeVenderNSerie = False
                        '    Mensaje = "1 No se puede generar un abono de este número de Ref: " & strNSerie
                        'End If
                    Else
                        If Nz(rscReferencia("Stock"), 0) <= 0 Then
                            SePuedeVenderNSerie = False
                            Mensaje = "2 El Número de Ref: " & strNSerie & " NO tiene stock o ya se ha realizado una Venta"
                        End If
                    End If
                Else
                    If Abono Then
                        SePuedeVenderNSerie = False
                        Mensaje = "3 No se puede generar un abono de este número de Ref: " & strNSerie & Chr(13) & Chr(10) & " No tiene registrado ninguna venta"
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            rscReferencia = Nothing
        End Try
    End Function

    Public Function HayStock(ByVal strIDArticulo As String, ByVal strNSerie As String, ByVal Cant As Double) As Boolean
        Dim rscReferencia As DataTable
        Try
            rscReferencia = AdminData.Filter("tbArticuloNumeroSerie", "Stock", "IDArticulo = '" & strIDArticulo & "' and NSerie = '" & strNSerie & "'")
            HayStock = False

            If Not rscReferencia Is Nothing Then

                If rscReferencia.Rows.Count > 0 Then

                    If Cant > 0 Then

                        If Nz(rscReferencia("Stock"), 0) >= Cant Then
                            HayStock = True
                        End If
                    Else
                        HayStock = True
                        'Si es un abono no  se hace nada
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            rscReferencia = Nothing
        End Try

    End Function

    Public Function EliminarNSerie(ByVal strIDArticulo As String, ByVal strNSerie As String) As Boolean

        Dim rscReferencia As DataTable

        Try
            AdminData.Execute("DELETE FROM tbArticuloNumeroSerie WHERE IDArticulo = '" & strIDArticulo & "' and NSerie = '" & strNSerie & "' AND STOCK = 0")
            rscReferencia = AdminData.Filter("tbArticuloNumeroSerie", "Stock", "IDArticulo = '" & strIDArticulo & "' and NSerie = '" & strNSerie & "'")
            If rscReferencia.Rows.Count = 0 Then EliminarNSerie = True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            rscReferencia = Nothing
        End Try

    End Function

    Public Function ActualizarArticuloNSerie(ByVal strIDArticulo As String, ByVal strNSerie As String, ByVal Cant As Double, ByVal IDLineaAlbaran As Long) As Boolean
        Dim rsArticulo As DataTable
        Dim vstock As Double
        Dim vSql As String

        Try
            rsArticulo = AdminData.Filter("tbArticuloNumeroSerie", "Stock,UltimoDocCompra", "IDArticulo = '" & strIDArticulo & "' and NSerie = '" & strNSerie & "'")
            If Not rsArticulo Is Nothing Then
                If rsArticulo.Rows.Count > 0 Then
                    'acutualizar la linea anterior en albaranes compra
                    'esto es para seguir la traza
                    If IDLineaAlbaran <> Nz(rsArticulo("UltimoDocCompra"), 0) Then
                        vSql = "UPDATE tbAlbaranCompraNSerie SET IDLineaAnterior = " & Nz(rsArticulo("UltimoDocCompra"), 0) & " where AlbaranCompraNSerie = " & IDLineaAlbaran
                        AdminData.Execute(vSql)
                    End If
                    vstock = Cant
                    'Obtenemos stock
                    rsArticulo = AdminData.Filter("vFrmArticuloNSerieCompra", "TotalC", "IDArticulo = '" & strIDArticulo & "' and NSerie = '" & strNSerie & "'")
                    If rsArticulo.Rows.Count > 0 Then
                        vstock = Nz(rsArticulo("TotalC"), 0)
                    End If

                    rsArticulo = AdminData.Filter("vFrmArticuloNSerieVenta", "TotalC", "IDArticulo = '" & strIDArticulo & "' and NSerie = '" & strNSerie & "'")
                    If rsArticulo.Rows.Count > 0 Then
                        vstock = vstock - Nz(rsArticulo("TotalC"), 0)
                    End If
                    'Fin Obtener stock
                    vSql = "UPDATE tbArticuloNumeroSerie  " _
                    & " SET UltimoDocCompra = " & IDLineaAlbaran & "," _
                    & " TipoDocumento = 'AC',Stock = " & Replace(vstock, ",", ".") _
                    & " where IDArticulo = '" & strIDArticulo & "' and NSerie = '" & strNSerie & "'"
                    AdminData.Execute(vSql)
                ElseIf strIDArticulo.Trim <> "" AndAlso strNSerie.Trim <> "" Then
                    vSql = "INSERT INTO tbArticuloNumeroSerie  " _
                    & "( IDArticulo , NSerie  ,  UltimoDocCompra ,  TipoDocumento  ,   Stock  )" _
                    & " VALUES ( " _
                    & "'" & strIDArticulo & "','" & strNSerie & "'," & IDLineaAlbaran & ", 'AC', " & Replace(Cant, ",", ".") & ")"
                    AdminData.Execute(vSql)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            rsArticulo = Nothing
        End Try
    End Function

    Public Function ActualizarArticuloNSerieV(ByVal strIDArticulo As String, ByVal strNSerie As String, ByVal Cant As Double, ByVal IDLineaAlbaran As Long) As Boolean
        Dim rsArticulo As New DataTable
        Dim vstock As Double
        Dim vSql As String

        Try
            rsArticulo = AdminData.Filter("tbArticuloNumeroSerie", "Stock,UltimoDocVenta", "IDArticulo = '" & strIDArticulo & "' and NSerie = '" & strNSerie & "'")
            If Not rsArticulo Is Nothing Then
                If rsArticulo.Rows.Count > 0 Then
                    'acutualizar la linea anterior en albaranes compra
                    'esto es para seguir la traza
                    'If IDLineaAlbaran <> Nz(rsArticulo("UltimoDocVenta"), 0) Then
                    ' vSql = "UPDATE tbAlbaranVentaNSerie SET IDLineaAnterior = " & Nz(rsArticulo("UltimoDocVenta"), 0) & " where AlbaranVentaNSerie = " & idAlb
                    ' fwArticulo.Ejecutar (vSql)
                    ' End If
                    vstock = Cant
                    'Obtenemos stock
                    rsArticulo = AdminData.Filter("vFrmArticuloNSerieCompra", "TotalC", "IDArticulo = '" & strIDArticulo & "' and NSerie = '" & strNSerie & "'")
                    If rsArticulo.Rows.Count > 0 Then
                        vstock = Nz(rsArticulo("TotalC"), 0)
                    End If

                    rsArticulo = AdminData.Filter("vFrmArticuloNSerieVenta", "TotalC", "IDArticulo = '" & strIDArticulo & "' and NSerie = '" & strNSerie & "'")
                    If rsArticulo.Rows.Count > 0 Then
                        vstock = vstock - Nz(rsArticulo("TotalC"), 0)
                    End If
                    'Fin Obtener stock
                    vSql = "UPDATE tbArticuloNumeroSerie  " _
                    & " SET UltimoDocVenta = " & IDLineaAlbaran & "," _
                    & " TipoDocumento = 'AV',Stock = " & Replace(vstock, ",", ".") _
                    & " where IDArticulo = '" & strIDArticulo & "' and NSerie = '" & strNSerie & "'"
                    AdminData.Execute(vSql)
                Else
                    vSql = "INSERT INTO tbArticuloNumeroSerie  " _
                    & "( IDArticulo , NSerie  ,  UltimoDocVenta ,  TipoDocumento  ,   Stock  )" _
                    & " VALUES ( " _
                    & "'" & strIDArticulo & "','" & strNSerie & "'," & IDLineaAlbaran & ", 'AV', " & Replace(Cant, ",", ".") & ")"
                    AdminData.Execute(vSql)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            rsArticulo = Nothing
        End Try

    End Function

    'Actualiza el stock, después de eliminar lineas de albaran
    Public Function ActualizarStockNSerie(ByVal strIDArticulo As String, ByVal strNSerie As String) As Boolean
        Dim rsArticulo As New DataTable
        Dim vstock As Double
        Dim vSql As String

        Try
            rsArticulo = AdminData.Filter("tbArticuloNumeroSerie", "Stock , UltimoDocCompra", "IDArticulo = '" & strIDArticulo & "' and NSerie = '" & strNSerie & "'")
            If Not rsArticulo Is Nothing Then
                If rsArticulo.Rows.Count > 0 Then

                    rsArticulo = AdminData.Filter("vFrmArticuloNSerieCompra", "TotalC", "IDArticulo = '" & strIDArticulo & "' and NSerie = '" & strNSerie & "'")
                    If rsArticulo.Rows.Count > 0 Then
                        vstock = Nz(rsArticulo("TotalC"), 0)
                    End If

                    rsArticulo = AdminData.Filter("vFrmArticuloNSerieVenta", "TotalC", "IDArticulo = '" & strIDArticulo & "' and NSerie = '" & strNSerie & "'")
                    If rsArticulo.Rows.Count > 0 Then
                        vstock = vstock - Nz(rsArticulo("TotalC"), 0)
                    End If
                    'Fin Obtener stock
                    vSql = "UPDATE tbArticuloNumeroSerie SET Stock = " & Replace(vstock, ",", ".") _
                    & " where IDArticulo = '" & strIDArticulo & "' and NSerie = '" & strNSerie & "'"
                    AdminData.Execute(vSql)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            rsArticulo = Nothing
        End Try

    End Function

    'Actualiza el id del último albaran de compra
    Public Function ActualizarUltimaCompra(ByVal strIDArticulo As String, ByVal strNSerie As String, ByVal IDAlbaran As Long) As Boolean
        Dim vSql As String

        Try
            vSql = "UPDATE tbArticuloNumeroSerie  " _
                      & " SET UltimoDocCompra = " & IDAlbaran _
                      & " where IDArticulo = '" & strIDArticulo & "' and NSerie = '" & strNSerie & "'"
            AdminData.Execute(vSql)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Function

    'Actualiza el id del último albaran de Venta
    Public Function ActualizarUltimaVenta(ByVal strIDArticulo As String, ByVal strNSerie As String, ByVal IDAlbaran As Long) As Boolean
        Dim vSql As String

        Try
            vSql = "UPDATE tbArticuloNumeroSerie  " _
                             & " SET UltimoDocVenta = " & IDAlbaran _
                             & " where IDArticulo = '" & strIDArticulo & "' and NSerie = '" & strNSerie & "'"
            AdminData.Execute(vSql)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

End Class
