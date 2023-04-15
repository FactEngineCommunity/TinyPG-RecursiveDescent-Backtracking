Imports System.Collections.Generic
Imports System.Windows.Forms

Namespace TinyPG.Controls
    Public Class Floaties
        Inherits List(Of IFloaty)
        ' Methods
        Public Overloads Function Find(ByVal container As Control) As IFloaty
            Dim f As Floaty
            For Each f In Me
                If f.DockState.Container.Equals(container) Then
                    Return f
                End If
            Next
            Return Nothing
        End Function

    End Class
End Namespace

