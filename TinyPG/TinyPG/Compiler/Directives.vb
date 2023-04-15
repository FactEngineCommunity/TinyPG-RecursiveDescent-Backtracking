Imports System
Imports System.Collections.Generic
Imports System.Reflection

Namespace TinyPG.Compiler
    <DefaultMember("Item")> _
    Public Class Directives
        Inherits List(Of Directive)
        ' Methods
        Public Overloads Function Exists(ByVal directive As Directive) As Boolean
            Return MyBase.Exists(Function(d) (d.Name = directive.Name))
        End Function

        Public Overloads Function Find(ByVal name As String) As Directive
            Return MyBase.Find(Function(d) (d.Name = name))
        End Function


        ' Properties
        Default Public Overloads ReadOnly Property Item(ByVal name As String) As Directive
            Get
                Return Me.Find(name)
            End Get
        End Property

    End Class
End Namespace

