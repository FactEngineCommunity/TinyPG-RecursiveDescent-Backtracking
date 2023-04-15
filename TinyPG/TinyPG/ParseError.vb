Imports System

Namespace TinyPG
    <Serializable> _
    Public Class ParseError

        Private codeField As Integer
        Private colField As Integer
        Private lengthField As Integer
        Private lineField As Integer
        Private messageField As String
        Private posField As Integer
        Private m_expected_token As String = ""

        ' Methods
        Public Sub New()
        End Sub

        Public Sub New(ByVal message As String,
                       ByVal code As Integer)

        End Sub

        Public Sub New(ByVal message As String, ByVal code As Integer, ByVal node As ParseNode)
            Me.New(message, code, 0, node.Token.StartPos, node.Token.StartPos, node.Token.Length)
        End Sub

        Public Sub New(ByVal message As String,
                       ByVal code As Integer,
                       ByVal line As Integer,
                       ByVal col As Integer,
                       ByVal pos As Integer,
                       ByVal length As Integer,
                       Optional ByVal asExpectedToken As String = "")

            Me.messageField = message
            Me.codeField = code
            Me.lineField = line
            Me.colField = col
            Me.posField = pos
            Me.lengthField = length
            m_expected_token = asExpectedToken
        End Sub


        ' Properties
        Public ReadOnly Property Code As Integer
            Get
                Return Me.codeField
            End Get
        End Property

        Public ReadOnly Property Column As Integer
            Get
                Return Me.colField
            End Get
        End Property

        Public ReadOnly Property Length As Integer
            Get
                Return Me.lengthField
            End Get
        End Property

        Public ReadOnly Property Line As Integer
            Get
                Return Me.lineField
            End Get
        End Property

        Public ReadOnly Property Message As String
            Get
                Return Me.messageField
            End Get
        End Property

        Public ReadOnly Property Position As Integer
            Get
                Return Me.posField
            End Get
        End Property
        
    End Class
End Namespace

