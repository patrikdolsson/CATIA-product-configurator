'Imports System.Linq
'Imports System.Xml.Linq
Imports System.IO
Imports System.Text
Imports INFITF
Imports KnowledgewareTypeLib
Imports MECMOD
Imports Newtonsoft.Json
Imports ProductStructureTypeLib

Module ExportSTL

    Public Function GetPartsToExport() As DataTable
        ' This is the main function that declares how the product should be exported.

        ' Declare a default type for each part as allowed by the ConfigureFunctions.GetAcceptableTypes() function
        ' and used in ConfigureFunctions.Configure() sub
        Dim typeArray(7) As String
        typeArray(0) = "Box"
        typeArray(1) = "Round"
        typeArray(2) = ""
        typeArray(3) = ""
        typeArray(4) = "Round"
        typeArray(5) = "Round"
        typeArray(6) = "Round"
        typeArray(7) = "Type1"

        ' Declare the list of parameter names as used in configurationParameterArray in the ConfigureFunctions.Configure() sub
        ' Note that the name needs to have the exact identifier for each parameter that they are inside of the HLCts. The order is also essential.
        Dim parameterList(21) As String
        parameterList(0) = "baseWidth"
        parameterList(1) = "baseDepth"
        parameterList(2) = "baseHeight"
        parameterList(3) = "baseConnectionHeight"
        parameterList(4) = "baseConnectionDiameter"
        parameterList(5) = "lampSwivelAngle"
        parameterList(6) = "lowerConnectionWidth"
        parameterList(7) = "upperConnectionWidth"
        parameterList(8) = "lowerAxleDiameter"
        parameterList(9) = "lowerAxleAngle"
        parameterList(10) = "upperAxleDiameter"
        parameterList(11) = "upperAxleAngle"
        parameterList(12) = "lowerSupportHeight"
        parameterList(13) = "lowerSupportDiameter"
        parameterList(14) = "upperSupportHeight"
        parameterList(15) = "upperSupportDiameter"
        parameterList(16) = "headSupportWidth"
        parameterList(17) = "headSupportHeight"
        parameterList(18) = "headSupportAngle"
        parameterList(19) = "lampHeadDiameter"
        parameterList(20) = "lampHeadLength"
        parameterList(21) = "lampHeadAngle"

        ' Set readable names for each instantiated part, types, and parameters.
        ' The purpose of the readable names is to make the names for things in
        ' the gui configurator inside the WebGL implementation look a bit nicer
        Dim readableNames As New Hashtable From {
            {"LampBase", "Lamp Base"},
            {"BaseConnection", "Base Connection"},
            {"LowerAxle", "Lower Axle"},
            {"UpperAxle", "Upper Axle"},
            {"LowerSupportRod1", "Lower Support"},
            {"UpperSupportRod", "Upper Support"},
            {"HeadSupport", "Head Support"},
            {"LampHead", "Lamp Head"},
            {"Box", "Box"},
            {"Oval", "Oval"},
            {"Round", "Round"},
            {"Type1", "Type 1"},
            {parameterList(0), "Base Width"},
            {parameterList(1), "Base Depth"},
            {parameterList(2), "Base Height"},
            {parameterList(3), "Base Connection Height"},
            {parameterList(4), "Base Connection Diameter"},
            {parameterList(5), "Lamp Swivel Angle"},
            {parameterList(6), "Lower Connection Width"},
            {parameterList(7), "Upper Connection Width"},
            {parameterList(8), "Lower Axle Diameter"},
            {parameterList(9), "Lower Axle Angle"},
            {parameterList(10), "Upper Axle Diameter"},
            {parameterList(11), "Upper Axle Angle"},
            {parameterList(12), "Lower Support Height"},
            {parameterList(13), "Lower Support Diameter"},
            {parameterList(14), "Upper Support Height"},
            {parameterList(15), "Upper Support Diameter"},
            {parameterList(16), "Head Support Width"},
            {parameterList(17), "Head Support Height"},
            {parameterList(18), "Head Support Angle"},
            {parameterList(19), "Lamp Head Diameter"},
            {parameterList(20), "Lamp Head Length"},
            {parameterList(21), "Lamp Head Angle"}
        }

        ' Set the settings to configure the stl export for each part. 
        ' First argument ("LampBase") represents the base name for the part
        ' Second argument ("Box") represents the type of base part
        ' Third argument (GetTypeArray(typeArray)) represents the default type array.
        ' Fourth argument (parameterList) represents the list of parameter names
        ' Fifth argument represents the array of placement points as they are called inside of the HLCts
        ' Sixth argument represents the array of output points as they are called inside of the HLCts
        ' Seventh argument represents the geometry affecting parameters as the index of the parameterList.
        '       As defined in the Geometry Affecting Parameter Associative Structure Matrix.
        ' Eighth argument represents the arraylist that will become the full list of configurations
        ' Ninth argument represents a boolean that decides if this part should be exported as stl or not
        '       If set to false. Only the placement points and output points will be recorded into the stl info json
        ' 10th argument represents the readable names.
        ' 11th argument represents the parameters that should be available for change in the gui for that part.
        '       As a general rule, one parameter should only be able to be changed in one place at one time,
        '       meaning that multiple parts should in general not have the same parameters in this arraylist.
        ' 12th argument represents continuous positional or rotational parameters that should be available
        '       for change in the gui for that part
        Dim productDataTable = CreateExportingSettingsDataTable()
        productDataTable.Rows.Add("LampBase", "Box", GetTypeArray(typeArray), parameterList,
                               {"PlacementPoint"},
                               {"OutputPoint1"},
                               New ArrayList From {0, 1, 2}, New ArrayList, True, readableNames,
                               New ArrayList From {0, 1, 2}, New ArrayList From {})

        typeArray(0) = "Oval"
        productDataTable.Rows.Add("LampBase", "Oval", GetTypeArray(typeArray), parameterList,
                               {"PlacementPoint"},
                               {"OutputPoint1"},
                               New ArrayList From {0, 1, 2}, New ArrayList, True, readableNames,
                               New ArrayList From {0, 1, 2}, New ArrayList From {})

        productDataTable.Rows.Add("BaseConnection", "Round", GetTypeArray(typeArray), parameterList,
                                     {"PlacementPoint"},
                                     {"OutputPoint1"},
                                     New ArrayList From {3, 4}, New ArrayList, True, readableNames,
                                     New ArrayList From {3, 4}, New ArrayList From {5})

        productDataTable.Rows.Add("LowerAxle", "", GetTypeArray(typeArray), parameterList,
                                {"PlacementPoint1", "PlacementPoint2"},
                                {"OutputPoint1", "OutputPoint2", "OutputPoint3"},
                                New ArrayList From {6, 8}, New ArrayList, True, readableNames,
                                New ArrayList From {6, 8}, New ArrayList From {9})

        productDataTable.Rows.Add("UpperAxle", "", GetTypeArray(typeArray), parameterList,
                                {"PlacementPoint1", "PlacementPoint2"},
                                {"OutputPoint1", "OutputPoint2", "OutputPoint3"},
                                New ArrayList From {7, 10}, New ArrayList, True, readableNames,
                                New ArrayList From {7, 10}, New ArrayList From {11})

        productDataTable.Rows.Add("LowerSupportRod1", "Round", GetTypeArray(typeArray), parameterList,
                                      {"PlacementPoint"},
                                      {"OutputPoint1"},
                                      New ArrayList From {12, 13}, New ArrayList, True, readableNames,
                                      New ArrayList From {12, 13}, New ArrayList From {})

        productDataTable.Rows.Add("LowerSupportRod2", "Round", GetTypeArray(typeArray), parameterList,
                                      {"PlacementPoint"},
                                      {"OutputPoint1"},
                                      New ArrayList From {12, 13}, New ArrayList, False, readableNames,
                                      New ArrayList From {}, New ArrayList From {})

        productDataTable.Rows.Add("UpperSupportRod", "Round", GetTypeArray(typeArray), parameterList,
                                      {"PlacementPoint"},
                                      {"OutputPoint1"},
                                      New ArrayList From {14, 15}, New ArrayList, True, readableNames,
                                      New ArrayList From {14, 15}, New ArrayList From {})

        productDataTable.Rows.Add("HeadSupport", "Round", GetTypeArray(typeArray), parameterList,
                                  {"PlacementPoint"},
                                  {"OutputPoint1"},
                                  New ArrayList From {15, 16, 17}, New ArrayList, True, readableNames,
                                  New ArrayList From {16, 17}, New ArrayList From {18})

        productDataTable.Rows.Add("LampHead", "Type1", GetTypeArray(typeArray), parameterList,
                               {"PlacementPoint"},
                               {},
                               New ArrayList From {16, 17, 19, 20}, New ArrayList, True, readableNames,
                               New ArrayList From {19, 20}, New ArrayList From {21})

        Return productDataTable
    End Function
    Public Function GetTypeArray(typeArray)
        Dim newTypeArray(typeArray.Length) As String
        Array.Copy(typeArray, newTypeArray, typeArray.Length)
        Return newTypeArray
    End Function

    Public Function GetExportSTLConfigurationsArray(stlExportingSettingsArray)
        ' This function sets the uniform steps for each of the parameters that should be used to determine all of the combinations.
        ' All of the different combinations for each of the parameter values are calculated and added to the configurationList.

        Dim partsToExport = GetPartsToExport()

        For Each partType As DataRow In partsToExport.Rows
            Dim defaultSTLExportConfigurationSetting = GetDefaultSTLExportConfigurationSetting(stlExportingSettingsArray)

            Dim changingSettings = partType.Field(Of ArrayList)("settingsIndexList")
            Dim changingSettingsList As ArrayList = New ArrayList

            For Each setting As Integer In changingSettings
                Dim innerListLength = stlExportingSettingsArray(setting)(2)
                Dim innerListValues As ArrayList = New ArrayList
                For i = 0 To innerListLength
                    If innerListLength = 0 Then
                        innerListValues.Add(stlExportingSettingsArray(setting)(0))
                    Else
                        innerListValues.Add((stlExportingSettingsArray(setting)(1) - stlExportingSettingsArray(setting)(0)) * i / innerListLength + stlExportingSettingsArray(setting)(0))
                    End If
                Next

                changingSettingsList.Add(innerListValues)
            Next

            Dim values As New ArrayList
            For i As Integer = 1 To changingSettingsList.Count
                values.Add(0)
            Next
            Dim configurationList As New ArrayList

            Functions.GenerateNestedForLoops(changingSettingsList, 0, values, configurationList, defaultSTLExportConfigurationSetting, changingSettings)
            partType("settingsList") = configurationList

            Console.WriteLine(partType("partName") & "_" & partType("type") & ": " & partType("settingsList").Count)
        Next

        Dim exportSTLConfigurationsArray = partsToExport
        Return exportSTLConfigurationsArray
    End Function
    Public Function GetDefaultSTLExportConfigurationSetting(stlExportingSettingsArray)
        Dim defaultSTLExportConfigurationSetting As ArrayList = New ArrayList

        For i = 0 To stlExportingSettingsArray.Length - 1
            defaultSTLExportConfigurationSetting.Add(stlExportingSettingsArray(i)(0))
        Next
        Return defaultSTLExportConfigurationSetting
    End Function
    Public Function CreateExportingSettingsDataTable()
        Dim table As New DataTable
        table.Columns.Add("partName", GetType(String))
        table.Columns.Add("type", GetType(String))
        table.Columns.Add("typeArray", GetType(Array))
        table.Columns.Add("parameterArray", GetType(Array))
        table.Columns.Add("placementPoints", GetType(Array))
        table.Columns.Add("outputPoints", GetType(Array))
        table.Columns.Add("settingsIndexList", GetType(ArrayList))
        table.Columns.Add("settingsList", GetType(ArrayList))
        table.Columns.Add("exportBoolean", GetType(Boolean))
        table.Columns.Add("readableNames", GetType(Hashtable))
        table.Columns.Add("GUIControlParametersIndex", GetType(ArrayList))
        table.Columns.Add("GUIControlSlidingParametersIndex", GetType(ArrayList))
        Return table
    End Function

    Public Function GetPointCoordinatesFromPart(CATIA As Object, documents As Documents, products As Products, partName As String, geoSetName As String, pointName As String)
        ' Extract the coordinates from a point inside the part
        Dim part = documents.Item(products.Item(partName).PartNumber & ".CATPart").Part

        Dim geoSet = part.FindObjectByName(geoSetName).HybridShapes
        Dim point = geoSet.Item(pointName)

        Dim coordinates = GetPartCoordinatesOfPoint(CATIA, point)
        Return coordinates
    End Function
    Public Function GetPartCoordinatesOfPoint(CATIA As Object, point As Object)
        Dim myWorkbench = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench")
        Dim myMeasure = myWorkbench.GetMeasurable(point)
        Dim coordinates(2)
        myMeasure.GetPoint(coordinates)
        Return coordinates
    End Function
    Public Sub ExportPartCATPartSTL(CATIA, documents, products, stlFilesLocation, tempFilesLocation, partName, stlName)
        ' Export a certain part in stl format to the specified directory.
        If Not Directory.Exists(stlFilesLocation & partName & "\") Then
            Directory.CreateDirectory(stlFilesLocation & partName & "\")
        End If
        SaveProduct(CATIA, tempFilesLocation)
        CATIA.DisplayFileAlerts = False
        documents.Open(tempFilesLocation & products.Item(partName).PartNumber & ".CATPart")
        CATIA.ActiveDocument.ExportData(stlFilesLocation & partName & "\" & stlName, "stl")
        CATIA.ActiveDocument.Close
        CATIA.DisplayFileAlerts = True
    End Sub
End Module
