Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports TinyPG.Debug

Namespace TinyPG
    Public NotInheritable Class ParseTreeViewer
        ' Methods
        Private Sub New()
        End Sub

        Public Shared Sub Populate(ByVal treeview As TreeView, ByVal parsetree As IParseTree)
            treeview.Visible = False
            treeview.SuspendLayout
            treeview.Nodes.Clear
            treeview.Tag = parsetree
            Dim start As IParseNode = parsetree.INodes.Item(0)
            Dim node As New TreeNode(start.Text) With { _
                .Tag = start, _
                .ForeColor = Color.SteelBlue _
            }
            treeview.Nodes.Add(node)
            ParseTreeViewer.PopulateNode(node, start)
            treeview.ExpandAll
            treeview.ResumeLayout
            treeview.Visible = True
        End Sub

        Private Shared Sub PopulateNode(ByVal node As TreeNode, ByVal start As IParseNode)
            Dim ipn As IParseNode
            For Each ipn In start.INodes
                Dim tn As New TreeNode(ipn.Text) With { _
                    .Tag = ipn _
                }
                node.Nodes.Add(tn)
                ParseTreeViewer.PopulateNode(tn, ipn)
            Next
        End Sub

    End Class
End Namespace

