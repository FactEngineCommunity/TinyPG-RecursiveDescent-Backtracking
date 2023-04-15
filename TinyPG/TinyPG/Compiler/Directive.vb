Imports System
Imports System.Collections.Generic
Imports System.Runtime.CompilerServices

Namespace TinyPG.Compiler
    Public Class Directive
        Inherits Dictionary(Of String, String)

        Private nameField As String

        ' Methods
        Public Sub New(ByVal name As String)
            Me.Name = name
        End Sub


        ' Properties
        Public Property Name As String
            Get
                Return Me.nameField
            End Get
            Set(ByVal value As String)
                Me.nameField = value
            End Set
        End Property
    End Class
End Namespace

