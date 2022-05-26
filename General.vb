
'Module General

'    Public Const cnLENGTH_NIVELES_ANALITICA As Integer = 3

'    'Estados de activo predeterminados de la aplicacion (con marca Sistema=1)
'    Public Const ESTADOACTIVO_DISPONIBLE As String = "0"
'    Public Const ESTADOACTIVO_ENMANTENIMIENTO As String = "1"
'    Public Const ESTADOACTIVO_RESERVADA As String = "2"
'    Public Const ESTADOACTIVO_TRABAJANDO As String = "3"
'    Public Const ESTADOACTIVO_VENDIDO As String = "4"
'    Public Const ESTADOACTIVO_BAJA As String = "5"
'    Public Const ESTADOACTIVO_AVERIADO As String = "6"
'    Public Const ESTADOACTIVO_ENTRANSITO As String = "7"
'    Public Const ESTADOACTIVO_OCUPADOENPORTE As String = "8"
'    Public Const ESTADOACTIVO_PENDIENTEDERETORNAR As String = "14"

'End Module

Public Class DataAplicarDecimalesMoneda
    Public IDMoneda As String
    Public Fecha As Date
    Public Row As DataRow

    Public Sub New(ByVal IDMoneda As String, ByVal Row As DataRow, Optional ByVal Fecha As Date = cnMinDate)
        Me.IDMoneda = IDMoneda
        Me.Fecha = Fecha
        Me.Row = Row
    End Sub
End Class


Public Class ArrayManager

#Region "Object"
    Public Shared Sub Copy(ByVal sourceArray As Object(), ByRef destinationArray As Object())
        If sourceArray.Length > 0 Then
            ReDim Preserve destinationArray(UBound(destinationArray) + sourceArray.Length)
            Array.Copy(sourceArray, 0, destinationArray, destinationArray.Length - sourceArray.Length, sourceArray.Length)
        End If
    End Sub

    Public Shared Sub Copy(ByVal data As Object, ByRef destinationArray As Object())
        If Not data Is Nothing Then
            ReDim Preserve destinationArray(UBound(destinationArray) + 1)
            destinationArray(UBound(destinationArray)) = data
        End If
    End Sub
#End Region

#Region "Datatable"
    Public Shared Sub Copy(ByVal sourceArray As DataTable(), ByRef destinationArray As DataTable())
        If sourceArray.Length > 0 Then
            ReDim Preserve destinationArray(UBound(destinationArray) + sourceArray.Length)
            Array.Copy(sourceArray, 0, destinationArray, destinationArray.Length - sourceArray.Length, sourceArray.Length)
        End If
    End Sub

    Public Shared Sub Copy(ByVal data As DataTable, ByRef destinationArray As DataTable())
        If Not data Is Nothing Then
            ReDim Preserve destinationArray(UBound(destinationArray) + 1)
            destinationArray(UBound(destinationArray)) = data
        End If
    End Sub
#End Region

#Region "Datarow"
    Public Shared Sub Copy(ByVal sourceArray As DataRow(), ByRef destinationArray As DataRow())
        If sourceArray.Length > 0 Then
            ReDim Preserve destinationArray(UBound(destinationArray) + sourceArray.Length)
            Array.Copy(sourceArray, 0, destinationArray, destinationArray.Length - sourceArray.Length, sourceArray.Length)
        End If
    End Sub

    Public Shared Sub Copy(ByVal data As DataRow, ByRef destinationArray As DataRow())
        If Not data Is Nothing Then
            ReDim Preserve destinationArray(UBound(destinationArray) + 1)
            destinationArray(UBound(destinationArray)) = data
        End If
    End Sub

#End Region

#Region "Integer"
    Public Shared Sub Copy(ByVal sourceArray As Integer(), ByRef destinationArray As Integer())
        If sourceArray.Length > 0 Then
            ReDim Preserve destinationArray(UBound(destinationArray) + sourceArray.Length)
            Array.Copy(sourceArray, 0, destinationArray, destinationArray.Length - sourceArray.Length, sourceArray.Length)
        End If
    End Sub

    Public Shared Sub Copy(ByVal data As Integer, ByRef destinationArray As Integer())
        ReDim Preserve destinationArray(UBound(destinationArray) + 1)
        destinationArray(UBound(destinationArray)) = data
    End Sub
#End Region

#Region "Date"
    Public Shared Sub Copy(ByVal sourceArray As Date(), ByRef destinationArray As Date())
        If sourceArray.Length > 0 Then
            ReDim Preserve destinationArray(UBound(destinationArray) + sourceArray.Length)
            Array.Copy(sourceArray, 0, destinationArray, destinationArray.Length - sourceArray.Length, sourceArray.Length)
        End If
    End Sub

    Public Shared Sub Copy(ByVal data As Date, ByRef destinationArray As Date())
        ReDim Preserve destinationArray(UBound(destinationArray) + 1)
        destinationArray(UBound(destinationArray)) = data
    End Sub
#End Region

#Region "Double"
    Public Shared Sub Copy(ByVal sourceArray As Double(), ByRef destinationArray As Double())
        If sourceArray.Length > 0 Then
            ReDim Preserve destinationArray(UBound(destinationArray) + sourceArray.Length)
            Array.Copy(sourceArray, 0, destinationArray, destinationArray.Length - sourceArray.Length, sourceArray.Length)
        End If
    End Sub

    Public Shared Sub Copy(ByVal data As Double, ByRef destinationArray As Double())
        ReDim Preserve destinationArray(UBound(destinationArray) + 1)
        destinationArray(UBound(destinationArray)) = data
    End Sub
#End Region

#Region "String"
    Public Shared Sub Copy(ByVal sourceArray As String(), ByRef destinationArray As String())
        If sourceArray.Length > 0 Then
            ReDim Preserve destinationArray(UBound(destinationArray) + sourceArray.Length)
            Array.Copy(sourceArray, 0, destinationArray, destinationArray.Length - sourceArray.Length, sourceArray.Length)
        End If
    End Sub

    Public Shared Sub Copy(ByVal data As String, ByRef destinationArray As String())
        ReDim Preserve destinationArray(UBound(destinationArray) + 1)
        destinationArray(UBound(destinationArray)) = data
    End Sub
#End Region

#Region "StockData"
    Public Shared Sub Copy(ByVal sourceArray As StockData(), ByRef destinationArray As StockData())
        If sourceArray.Length > 0 Then
            ReDim Preserve destinationArray(UBound(destinationArray) + sourceArray.Length)
            Array.Copy(sourceArray, 0, destinationArray, destinationArray.Length - sourceArray.Length, sourceArray.Length)
        End If
    End Sub

    Public Shared Sub Copy(ByVal data As StockData, ByRef destinationArray As StockData())
        If Not data Is Nothing Then
            ReDim Preserve destinationArray(UBound(destinationArray) + 1)
            destinationArray(UBound(destinationArray)) = data
        End If
    End Sub
#End Region

#Region "StockUpdateData"
    Public Shared Sub Copy(ByVal sourceArray As StockUpdateData(), ByRef destinationArray As StockUpdateData())
        If sourceArray.Length > 0 Then
            ReDim Preserve destinationArray(UBound(destinationArray) + sourceArray.Length)
            Array.Copy(sourceArray, 0, destinationArray, destinationArray.Length - sourceArray.Length, sourceArray.Length)
        End If
    End Sub

    Public Shared Sub Copy(ByVal data As StockUpdateData, ByRef destinationArray As StockUpdateData())
        If Not data Is Nothing Then
            ReDim Preserve destinationArray(UBound(destinationArray) + 1)
            destinationArray(UBound(destinationArray)) = data
        End If
    End Sub
#End Region

End Class

<Serializable()> _
Public Class AlbaranLogProcess
    Implements System.Runtime.Remoting.Messaging.ILogicalThreadAffinative

    Public CreateData As LogProcess
    Public StockUpdateData(-1) As StockUpdateData
End Class



<Serializable()> _
Public Class DataDeclaraciones
    Public NDeclaracion As Integer
    Public AnioDeclaracion As Integer
    Public Filtro As Filter

    Public Sub New(ByVal NDeclaracion As Integer, ByVal AnioDeclaracion As Integer, ByVal Filtro As Filter)
        Me.NDeclaracion = NDeclaracion
        Me.AnioDeclaracion = AnioDeclaracion
        Me.Filtro = filtro
    End Sub

End Class