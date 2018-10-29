Public Class Form1
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim AutoCadObj As Object
        Dim AutoCadPath As String

        AutoCadObj = GetObject(, "AutoCAD.Application")
        AutoCadPath = AutoCadObj.Path
        AutoCadObj.LoadDVB(My.Application.Info.DirectoryPath & "\Test.dvb")
        AutoCadObj.RunMacro("PlotTest")
    End Sub
End Class
