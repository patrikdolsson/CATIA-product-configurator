Imports System.IO
Imports System.Text
Imports Newtonsoft.Json
Module JSONWriter
    Public Sub writeJsonAndExportSTL(CATIA, documents, products, objectsToSelect, cadFilesLocation, stlFilesLocation,
                                     tempFilesLocation, referenceName, stlExportingInformation, stlExportingSettingsInformation)
        ' This function writes the stl info json and exports the desired stl models.
        ' It uses Newtonsoft.Json.StringBuilder and StringWriter to accomplish this
        Dim sb As StringBuilder = New StringBuilder()
        Dim sw As StringWriter = New StringWriter(sb)

        Using writer As Newtonsoft.Json.JsonWriter = New JsonTextWriter(sw)
            writer.WriteStartObject()
            Dim basePartName = ""
            Dim firstLoop = True
            For Each stlExportingSettingDataRow As DataRow In stlExportingInformation.Rows
                If Not basePartName = stlExportingSettingDataRow("partName") Then
                    basePartName = stlExportingSettingDataRow("partName")
                    ' If it's the first loop it should add a "settings" object in the first level for the json that will contain the readable names.
                    If Not firstLoop Then
                        writer.WriteEndObject()
                    Else
                        firstLoop = False
                        writer.WritePropertyName("settings")
                        writer.WriteStartObject()
                        writer.WritePropertyName("readableNames")
                        writer.WriteStartObject()
                        Dim readableNamesHashtable As Hashtable = stlExportingSettingDataRow("readableNames")
                        Dim readableNames As New ArrayList
                        readableNames.AddRange(readableNamesHashtable.Keys)
                        For Each name In readableNames
                            writer.WritePropertyName(name)
                            writer.WriteValue(readableNamesHashtable(name))
                        Next
                        writer.WriteEndObject()
                        writer.WriteEndObject()
                        ' It then creates a "parts" object in the first level where all the stl info will be stored.
                        writer.WritePropertyName("parts")
                        writer.WriteStartObject()
                    End If
                    writer.WritePropertyName(stlExportingSettingDataRow("partName"))
                    writer.WriteStartObject()
                End If
                writer.WritePropertyName(stlExportingSettingDataRow("type"))
                writer.WriteStartArray()

                ' Inside each type, every single exported STL model gets its own array element to store parameter data, point coordinates and stl names.

                Dim firstInnerLoop = True
                For Each config In stlExportingSettingDataRow.Field(Of ArrayList)("settingsList")
                    Dim typeArray = stlExportingSettingDataRow.Field(Of Array)("typeArray")
                    ConfigureFunctions.Configure(CATIA, documents, products, objectsToSelect, referenceName, cadFilesLocation, typeArray, config, tempFilesLocation)

                    writer.WriteStartObject()

                    ' "stl" : "stlName"
                    Dim currentParameterValues As List(Of Double) = New List(Of Double)
                    For Each parameterIndex In stlExportingSettingDataRow("settingsIndexList")
                        currentParameterValues.Add(config(parameterIndex))
                    Next
                    Dim currParameterValueStrings = String.Join("_", currentParameterValues)
                    Dim partName = stlExportingSettingDataRow("partName") & If(stlExportingSettingDataRow("type").Length > 0, "_" & stlExportingSettingDataRow("type"), "")
                    Dim name = partName & "_" & currParameterValueStrings
                    Dim stlName = name & ".stl"

                    writer.WritePropertyName("name")
                    writer.WriteValue(name)

                    If stlExportingSettingDataRow("exportBoolean") Then
                        writer.WritePropertyName("path")
                        writer.WriteValue("STL" & "/" & partName & "/" & stlName)
                    End If

                    ' The first inner loop for each type of base part will have information used for the gui creation.
                    ' The parameters inside the GUIControlParameters and GUIControlSlidingParameters will be available
                    ' for change in the gui configuration in the WebGL implementation
                    If firstInnerLoop Then
                        writer.WritePropertyName("GUIControlParameters")
                        writer.WriteStartArray()
                        For Each parameterIndex In stlExportingSettingDataRow("GUIControlParametersIndex")
                            writer.WriteValue(stlExportingSettingDataRow("parameterArray")(parameterIndex))
                        Next
                        writer.WriteEndArray()

                        writer.WritePropertyName("GUIControlSlidingParameters")
                        writer.WriteStartObject()
                        For Each parameterIndex In stlExportingSettingDataRow("GUIControlSlidingParametersIndex")
                            writer.WritePropertyName(stlExportingSettingDataRow("parameterArray")(parameterIndex))
                            writer.WriteStartObject()
                            writer.WritePropertyName("default")
                            writer.WriteValue(stlExportingSettingsInformation(parameterIndex)(0))
                            writer.WritePropertyName("min")
                            writer.WriteValue(stlExportingSettingsInformation(parameterIndex)(1))
                            writer.WritePropertyName("max")
                            writer.WriteValue(stlExportingSettingsInformation(parameterIndex)(2))
                            writer.WritePropertyName("step")
                            writer.WriteValue(stlExportingSettingsInformation(parameterIndex)(3))
                            writer.WriteEndObject()
                        Next
                        writer.WriteEndObject()
                        firstInnerLoop = False
                    End If

                    ' "parameters" : {"parameter1" : parameter1Value, "parameter2" : parameter2Value, ...}
                    writer.WritePropertyName("parameters")
                    writer.WriteStartObject()
                    For Each parameterIndex In stlExportingSettingDataRow("settingsIndexList")
                        writer.WritePropertyName(stlExportingSettingDataRow("parameterArray")(parameterIndex))
                        writer.WriteValue(config(parameterIndex))
                    Next
                    writer.WriteEndObject()

                    ' "placement" : {"PlacementPoint1" : {"x" : xValue, "y" : yValue, "z" : zValue}, "PlacementPoint2" : {}, ...}
                    writer.WritePropertyName("placement")
                    writer.WriteStartObject()
                    For Each placementPoint In stlExportingSettingDataRow("placementPoints")
                        writer.WritePropertyName(placementPoint)
                        writer.WriteStartObject()
                        Dim coordinates = ExportSTL.GetPointCoordinatesFromPart(CATIA, documents, products, partName, "Placement", placementPoint)
                        writer.WritePropertyName("x")
                        writer.WriteValue(coordinates(0))
                        writer.WritePropertyName("y")
                        writer.WriteValue(coordinates(1))
                        writer.WritePropertyName("z")
                        writer.WriteValue(coordinates(2))
                        writer.WriteEndObject()
                    Next
                    writer.WriteEndObject()

                    ' "output" : {"OutputPoint1" : {"x" : xValue, "y" : yValue, "z" : zValue}, "OutputPoint2" : {}, ...}
                    writer.WritePropertyName("output")
                    writer.WriteStartObject()
                    For Each outputPoint In stlExportingSettingDataRow("outputPoints")
                        writer.WritePropertyName(outputPoint)
                        writer.WriteStartObject()
                        Dim coordinates = ExportSTL.GetPointCoordinatesFromPart(CATIA, documents, products, partName, "Output", outputPoint)
                        writer.WritePropertyName("x")
                        writer.WriteValue(coordinates(0))
                        writer.WritePropertyName("y")
                        writer.WriteValue(coordinates(1))
                        writer.WritePropertyName("z")
                        writer.WriteValue(coordinates(2))
                        writer.WriteEndObject()
                    Next
                    writer.WriteEndObject()

                    writer.WriteEndObject()

                    If stlExportingSettingDataRow("exportBoolean") Then
                        ExportSTL.ExportPartCATPartSTL(CATIA, documents, products, stlFilesLocation, tempFilesLocation, partName, stlName)
                    End If
                Next
                writer.WriteEndArray()
            Next
            writer.WriteEndObject()
            writer.WriteEndObject()
        End Using

        Dim json As String = sw.ToString()
        'Console.WriteLine(json)
        ' Once json is complete, it should be written to stlInfo.json in the desired location.
        IO.File.WriteAllText(stlFilesLocation & "stlInfo.json", json)
    End Sub
End Module
