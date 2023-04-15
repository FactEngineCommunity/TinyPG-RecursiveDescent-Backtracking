Imports System

Namespace TinyPG.Debug
    Public Interface IParseTree
        Inherits IParseNode
        ' Methods
        Function Eval(ByVal ParamArray paramlist As Object()) As Object
        Function PrintTree() As String
    End Interface
End Namespace

