Module Module1

    Sub Main()
        Dim AutoCadObj As Object
        Dim AutoCadPath As String

        AutoCadObj = GetObject(, "AutoCAD.Application")
        AutoCadPath = AutoCadObj.Path
        AutoCadObj.LoadDVB(My.Application.Info.DirectoryPath & "\Test.dvb")
        AutoCadObj.RunMacro("PlotTest")
    End Sub

End Module
