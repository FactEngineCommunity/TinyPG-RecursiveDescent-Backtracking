Imports System
Imports System.Collections.Generic
Imports System.Xml.Serialization

Namespace TinyPG
    Public Class Token

        Private endposField As Integer
        Private skippedField As List(Of Token)
        Private startposField As Integer
        Private textField As String
        <XmlAttribute()> _
        Public Type As TokenType
        Private valueField As Object

        ' Methods
        Public Sub New()
            Me.New(0, 0)
        End Sub

        Public Sub New(ByVal start As Integer, ByVal [end] As Integer)
            Me.Type = TokenType._UNDETERMINED_
            Me.StartPos = start
            Me.EndPos = [end]
            Me.Text = ""
            Me.Value = Nothing
        End Sub

        Public Overrides Function ToString() As String
            If (Not Me.Text Is Nothing) Then
                Return (Me.Type.ToString & " '" & Me.Text & "'")
            End If
            Return Me.Type.ToString
        End Function

        Public Sub UpdateRange(ByVal token As Token)
            If (token.StartPos < Me.StartPos) Then
                Me.StartPos = token.StartPos
            End If
            If (token.EndPos > Me.EndPos) Then
                Me.EndPos = token.EndPos
            End If
        End Sub


        ' Properties
        Public Property EndPos As Integer
            Get
                Return Me.endposField

            End Get
            Set(ByVal value As Integer)
                Me.endposField = value
            End Set
        End Property

        Public ReadOnly Property Length As Integer
            Get
                Return (Me.EndPos - Me.StartPos)
            End Get
        End Property

        Public Property Skipped As List(Of Token)
            Get
                Return Me.skippedField
            End Get
            Set(ByVal value As List(Of Token))
                Me.skippedField = value
            End Set
        End Property

        Public Property StartPos As Integer
            Get
                Return Me.startposField
            End Get
            Set(ByVal value As Integer)
                Me.startposField = value
            End Set
        End Property

        Public Property [Text] As String
            Get
                Return Me.textField
            End Get
            Set(ByVal value As String)
                Me.textField = value
            End Set
        End Property

        Public Property Value As Object
            Get
                Return Me.valueField
            End Get
            Set(ByVal value As Object)
                Me.valueField = value
            End Set
        End Property

    End Class
End Namespace

