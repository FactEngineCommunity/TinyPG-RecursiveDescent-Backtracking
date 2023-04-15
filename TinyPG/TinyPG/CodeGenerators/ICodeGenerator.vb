Imports System
Imports TinyPG.Compiler

Namespace TinyPG.CodeGenerators
    Public Interface ICodeGenerator
        ' Methods
        Function Generate(ByVal grammar As Grammar, ByVal debug As Boolean) As String

        ' Properties
        ReadOnly Property FileName As String

    End Interface
End Namespace

