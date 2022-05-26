Public Class DataPistolas

    <Serializable()> _
    Public Class PedidoCompraPistolas_Info
        Public IDPedido As Integer
        Public NPedido As String
        Public IDLineaPedido As Integer
        Public IDArticulo As String
        Public DescArticulo As String
        Public FechaEntrega As Date
        Public QPendiente As Double
        Public QServida As Double
        Public IDProveedor As String
        Public IDMoneda As String
        Public Sub New()
        End Sub
        Public Sub New(ByVal IDPedido As Integer, ByVal NPedido As String, ByVal IDLineaPedido As Integer, ByVal IDArticulo As String, ByVal QPendiente As Double, Optional ByVal FechaEntrega As Date = cnMinDate, Optional ByVal IDProveedor As String = "", Optional ByVal DescArticulo As String = "", Optional ByVal IDMoneda As String = "")
            Me.IDPedido = IDPedido
            Me.NPedido = NPedido
            Me.IDLineaPedido = IDLineaPedido
            Me.IDArticulo = IDArticulo
            Me.QPendiente = QPendiente
            Me.FechaEntrega = FechaEntrega
            If DescArticulo.Length > 0 Then Me.DescArticulo = DescArticulo
            Me.QServida = 0
            If IDProveedor.Length > 0 Then Me.IDProveedor = IDProveedor
            If IDMoneda.Length > 0 Then Me.IDMoneda = IDMoneda
        End Sub

    End Class

    <Serializable()> _
    Public Class ActionResultPistolas
        Public OK As Boolean
        Public Message As String
        Public Sub New()
            Me.OK = False
            Me.Message = String.Empty
        End Sub
        Public Sub New(ByVal OK As Boolean, ByVal Message As String)
            Me.OK = OK
            Me.Message = Message
        End Sub
    End Class

    <Serializable()> _
    Public Class Almacen_Info
        Public B As String
        Public Almacen As String
        Public Descripcion As String
        Public Stock As Double
        Public UltInvent As DateTime
        Public Inventariado As String
        Public Cantidad As Double
        Public IDArticulo As String

        Public Sub New()
        End Sub
        Public Sub New(ByVal Almacen As String, ByVal Descripcion As String, ByVal Stock As Double, ByVal UltInvent As DateTime, ByVal Inventariado As String, ByVal IDArticulo As String)
            Me.B = "N"
            Me.Almacen = Almacen
            Me.Descripcion = Descripcion
            Me.Stock = xRound(CDbl(Stock), 2)
            Me.UltInvent = UltInvent
            Me.Inventariado = Inventariado
            Me.Cantidad = 0
            Me.IDArticulo = IDArticulo
        End Sub

    End Class

    <Serializable()> _
    Public Class Proveedor_Info
        Public IDProveedor As String
        Public DescProveedor As String
        Public Sub New()
        End Sub
        Public Sub New(ByVal IDProveedor As String, ByVal DescProveedor As String)
            Me.IDProveedor = IDProveedor
            If DescProveedor.Length > 0 Then Me.DescProveedor = DescProveedor
        End Sub
    End Class

    <Serializable()> _
    Public Class Preparacion_Info

        Public Preparacion As String
        Public Estado As String
        Public FechaPrevista As DateTime
        Public Cliente As String
        Public Destino As String
        Public IDLineaPedido As String
        Public IDLineaPreparacion As String
        Public NPedido As Integer
        Public CodigoArticulo As String
        Public DescArticulo As String
        Public Cantidad As Double
        Public QExpedir As Double
        Public CodigoBarras As String
        Public QEmbalaje As Double
        Public DescCliente As String
        Public ProvinciaEnvio As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal Preparacion As String, ByVal Estado As String, ByVal FechaPrevista As DateTime, ByVal Cliente As String, ByVal Destino As String, ByVal IDLineaPedido As String, ByVal IDLineaPreparacion As String, ByVal NPedido As Integer, ByVal CodigoArticulo As String, ByVal DescArticulo As String, ByVal Cantidad As Double, ByVal QExpedir As Double, ByVal CodigoBarras As String, ByVal QEmbalaje As Double, ByVal DescCliente As String, ByVal ProvinciaEnvio As String)
            Me.Preparacion = Preparacion
            Me.Estado = Estado
            Me.FechaPrevista = FechaPrevista
            Me.Cliente = Cliente
            Me.Destino = Destino
            Me.IDLineaPedido = IDLineaPedido
            Me.IDLineaPreparacion = IDLineaPreparacion
            Me.NPedido = NPedido
            Me.CodigoArticulo = CodigoArticulo
            Me.DescArticulo = DescArticulo
            Me.Cantidad = Cantidad
            Me.QExpedir = QExpedir
            Me.CodigoBarras = CodigoBarras
            Me.QEmbalaje = QEmbalaje
            Me.DescCliente = DescCliente
            Me.ProvinciaEnvio = ProvinciaEnvio
        End Sub

    End Class

    <Serializable()> _
    Public Class ExisteArticulo
        Public Existe As Boolean
        Public IDArticulo As String
        Public DescArticulo As String
        Public CodigoBarras As String
        Public QEmbalaje As Double
        Public Sub New()
            Me.Existe = False
            Me.IDArticulo = String.Empty
            Me.DescArticulo = String.Empty
        End Sub
        Public Sub New(ByVal Existe As Boolean, ByVal IDArticulo As String, ByVal DescArticulo As String, ByVal CodigoBarras As String, ByVal QEmbalaje As Double)
            Me.Existe = Existe
            Me.IDArticulo = IDArticulo
            Me.DescArticulo = DescArticulo & String.Empty
            Me.CodigoBarras = CodigoBarras & String.Empty
            Me.QEmbalaje = QEmbalaje
        End Sub
    End Class

    <Serializable()> _
    Public Class ExisteProveedor
        Public Existe As Boolean
        Public DescProveedor As String
        Public Sub New()
            Me.Existe = False
            Me.DescProveedor = String.Empty
        End Sub
        Public Sub New(ByVal Existe As Boolean, ByVal DescProveedor As String)
            Me.Existe = Existe
            Me.DescProveedor = DescProveedor
        End Sub
    End Class


End Class
