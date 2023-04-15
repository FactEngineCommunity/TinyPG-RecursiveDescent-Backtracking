Imports System
Imports System.Threading
Imports System.Windows.Forms

Namespace TinyPG
    Friend Module Program
        ' Methods
        Private Sub Application_ThreadException(ByVal sender As Object, ByVal e As ThreadExceptionEventArgs)
            MessageBox.Show(("An unhandled exception occured: " & e.Exception.Message))
        End Sub

        <STAThread()> _
        Public Sub Main()
            AddHandler Application.ThreadException, New ThreadExceptionEventHandler(AddressOf Program.Application_ThreadException)
            Application.EnableVisualStyles()
            Application.SetCompatibleTextRenderingDefault(False)
            Application.Run(New MainForm)
        End Sub

    End Module
End Namespace

