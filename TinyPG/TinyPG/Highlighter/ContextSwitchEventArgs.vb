Imports System
Imports System.Runtime.InteropServices

Namespace TinyPG.Highlighter
    <Serializable, ComVisible(True)> _
    Public Class ContextSwitchEventArgs
        Inherits EventArgs
        ' Methods
        Public Sub New(ByVal prevContext As ParseNode, ByVal nextContext As ParseNode)
            Me.PreviousContext = prevContext
            Me.NewContext = nextContext
        End Sub


        ' Fields
        Public ReadOnly NewContext As ParseNode
        Public ReadOnly PreviousContext As ParseNode
    End Class
End Namespace

