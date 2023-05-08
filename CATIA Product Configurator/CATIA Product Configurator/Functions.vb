Module Functions
    Friend Function getFileLocation(directoryType) As String
        'Dim projectPath As String = CurDir() + ".\..\.."
        'Dim projectPath As String = AppDomain.CurrentDomain.BaseDirectory + ".\..\..\"
        Dim projectPath As String = My.Application.Info.DirectoryPath + ".\..\..\"
        Dim fileLocationFile As String = "FileLocation.XML"
        Dim doc As XDocument = XDocument.Load(projectPath + fileLocationFile)
        If directoryType = "CAD" Then
            Return doc.Root.<CAD>.Value
        ElseIf directoryType = "STL" Then
            Return doc.Root.<STL>.Value
        ElseIf directoryType = "temp" Then
            Return doc.Root.<temp>.Value
        End If
        Return vbNull
    End Function

    ' The following two functions provide a functional infrastructure to calculate and set all of the different combinations of configurations as defined
    ' in the ExportSTL.vb and in a Geometry Affecting Parameter Associative Structure Matrix. See the readme for more information.
    Public Sub GenerateNestedForLoops(ByVal loops As ArrayList, ByVal currentLoop As Integer, ByVal values As ArrayList,
                                      ByRef result As ArrayList, ByRef defaultSingleListResult As ArrayList, changingSettings As ArrayList)
        If currentLoop = loops.Count Then
            ' base case: we have generated all the loops, so process the values
            ProcessValues(values, result, defaultSingleListResult, changingSettings)
            Return
        End If

        ' recursive case: generate the current loop
        For Each value In loops(currentLoop)
            values(currentLoop) = value
            GenerateNestedForLoops(loops, currentLoop + 1, values, result, defaultSingleListResult, changingSettings)
        Next
    End Sub
    Public Sub ProcessValues(ByVal values As ArrayList, ByRef result As ArrayList, ByRef defaultSingleListResult As ArrayList, ByVal changingSettings As ArrayList)
        Dim currentConfiguration = New ArrayList(defaultSingleListResult)
        For i = 0 To changingSettings.Count - 1
            currentConfiguration(changingSettings(i)) = values(i)
        Next
        defaultSingleListResult = currentConfiguration
        result.Add(currentConfiguration)
    End Sub
End Module
