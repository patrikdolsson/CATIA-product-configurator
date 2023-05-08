Imports AnnotationTypeLib
Imports INFITF
Imports MECMOD
Imports ProductStructureTypeLib

Module ConfigureFunctions
    Public Function GetAcceptableTypes() As Array()
        ' Get an array of arrays that contain the supported types in the product configurator.
        ' If you only have one type, you can keep an empty string or not include it at all.
        Dim acceptableBaseTypes As String() = {"Box", "Oval"}
        Dim acceptableBaseConnectionTypes As String() = {"Round"}
        Dim acceptableLowerAxleTypes As String() = {""}
        Dim acceptableUpperAxleTypes As String() = {""}
        Dim acceptableLowerSupportRodTypes As String() = {"Round"}
        Dim acceptableUpperSupportRodTypes As String() = {"Round"}
        Dim acceptableHeadSupportTypes As String() = {"Round"}
        Dim acceptableLampHeadTypes As String() = {"Type1"}

        Dim acceptableTypes = New Array() {acceptableBaseTypes, acceptableBaseConnectionTypes, acceptableLowerAxleTypes, acceptableUpperAxleTypes,
            acceptableLowerSupportRodTypes, acceptableUpperSupportRodTypes, acceptableHeadSupportTypes, acceptableLampHeadTypes}

        Return acceptableTypes
    End Function
    Public Sub Configure(CATIA, documents, products, objectsToSelect, referenceName, cadFilesLocation, typeArray, configurationParameterArray, tempFilesLocation)
        ' Set parameter identifiers to their values according to configurationParameterArray
        Dim baseWidth = configurationParameterArray(0)
        Dim baseDepth = configurationParameterArray(1)
        Dim baseHeight = configurationParameterArray(2)
        Dim baseConnectionHeight = configurationParameterArray(3)
        Dim baseConnectionDiameter = configurationParameterArray(4)
        Dim lampSwivelAngle = configurationParameterArray(5)
        Dim lowerConnectionWidth = configurationParameterArray(6)
        Dim upperConnectionWidth = configurationParameterArray(7)
        Dim lowerAxleDiameter = configurationParameterArray(8)
        Dim lowerAxleAngle = configurationParameterArray(9)
        Dim upperAxleDiameter = configurationParameterArray(10)
        Dim upperAxleAngle = configurationParameterArray(11)
        Dim lowerSupportHeight = configurationParameterArray(12)
        Dim lowerSupportDiameter = configurationParameterArray(13)
        Dim upperSupportHeight = configurationParameterArray(14)
        Dim upperSupportDiameter = configurationParameterArray(15)
        Dim headSupportWidth = configurationParameterArray(16)
        Dim headSupportHeight = configurationParameterArray(17)
        Dim headSupportAngle = configurationParameterArray(18)
        Dim lampHeadDiameter = configurationParameterArray(19)
        Dim lampHeadLength = configurationParameterArray(20)
        Dim lampHeadAngle = configurationParameterArray(21)

        ' Separate the parameters that affect the reference part
        Dim referenceParameterArray(9) As Double
        referenceParameterArray(0) = baseWidth
        referenceParameterArray(1) = baseDepth
        referenceParameterArray(2) = baseHeight
        referenceParameterArray(3) = baseConnectionHeight
        referenceParameterArray(4) = lowerConnectionWidth
        referenceParameterArray(5) = upperConnectionWidth
        referenceParameterArray(6) = lowerSupportHeight
        referenceParameterArray(7) = upperSupportHeight
        referenceParameterArray(8) = headSupportWidth
        referenceParameterArray(9) = headSupportHeight

        ' Separate parameters that need to be adjusted in instantiated powercopies
        Dim partParameterArray(7) As Double
        partParameterArray(0) = baseHeight
        partParameterArray(1) = baseConnectionDiameter
        partParameterArray(2) = lowerAxleDiameter
        partParameterArray(3) = upperAxleDiameter
        partParameterArray(4) = lowerSupportDiameter
        partParameterArray(5) = upperSupportDiameter
        partParameterArray(6) = lampHeadDiameter
        partParameterArray(7) = lampHeadLength

        ' Modify the reference part and then configure all of the parts according the parameters
        ConfigureFunctions.ModifyReferenceParameters(documents, products, referenceName, referenceParameterArray)
        ConfigureFunctions.ConfigureParts(CATIA, documents, products, objectsToSelect, referenceName, cadFilesLocation, typeArray, partParameterArray, tempFilesLocation)
    End Sub
    Public Sub ModifyAngles(documents, products, reference, angleArray)
        Dim lampSwivelAngle = angleArray(0)
        Dim lowerAxleAngle = angleArray(1)
        Dim upperAxleAngle = angleArray(2)
        Dim headSupportAngle = angleArray(3)
        Dim lampHeadAngle = angleArray(4)

        CATIAFunctions.ModifyParameters(documents, products, reference, "BaseParameters\lampSwivelAngle", lampSwivelAngle)
        CATIAFunctions.ModifyParameters(documents, products, reference, "LowerSupportParameters\lowerAxleAngle", lowerAxleAngle)
        CATIAFunctions.ModifyParameters(documents, products, reference, "UpperSupportParameters\upperAxleAngle", upperAxleAngle)
        CATIAFunctions.ModifyParameters(documents, products, reference, "HeadSupportParameters\headSupportAngle", headSupportAngle)
        CATIAFunctions.ModifyParameters(documents, products, reference, "LampHeadParameters\lampHeadAngle", lampHeadAngle)
        CATIAFunctions.UpdatePart(products, reference)
    End Sub
    Public Sub ModifyReferenceParameters(documents, products, reference, referenceParameterArray)
        Dim baseWidth = referenceParameterArray(0)
        Dim baseDepth = referenceParameterArray(1)
        Dim baseHeight = referenceParameterArray(2)
        Dim baseConnectionHeight = referenceParameterArray(3)
        Dim lowerConnectionWidth = referenceParameterArray(4)
        Dim upperConnectionWidth = referenceParameterArray(5)
        Dim lowerSupportHeight = referenceParameterArray(6)
        Dim upperSupportHeight = referenceParameterArray(7)
        Dim headSupportWidth = referenceParameterArray(8)
        Dim headSupportHeight = referenceParameterArray(9)

        CATIAFunctions.ModifyParameters(documents, products, reference, "BaseParameters\baseWidth", baseWidth)
        CATIAFunctions.ModifyParameters(documents, products, reference, "BaseParameters\baseDepth", baseDepth)
        CATIAFunctions.ModifyParameters(documents, products, reference, "BaseParameters\baseHeight", baseHeight)
        CATIAFunctions.ModifyParameters(documents, products, reference, "BaseParameters\baseConnectionHeight", baseConnectionHeight)
        CATIAFunctions.ModifyParameters(documents, products, reference, "LowerSupportParameters\lowerConnectionWidth", lowerConnectionWidth)
        CATIAFunctions.ModifyParameters(documents, products, reference, "LowerSupportParameters\upperConnectionWidth", upperConnectionWidth)
        CATIAFunctions.ModifyParameters(documents, products, reference, "LowerSupportParameters\lowerSupportHeight", lowerSupportHeight)
        CATIAFunctions.ModifyParameters(documents, products, reference, "UpperSupportParameters\upperSupportHeight", upperSupportHeight)
        CATIAFunctions.ModifyParameters(documents, products, reference, "HeadSupportParameters\headSupportWidth", headSupportWidth)
        CATIAFunctions.ModifyParameters(documents, products, reference, "HeadSupportParameters\headSupportHeight", headSupportHeight)
        CATIAFunctions.UpdatePart(products, reference)
    End Sub
    Public Sub ConfigureParts(CATIA, documents, products, objectsToSelect, reference, powerCopyFilesLocation, typeArray, parameterArray, tempFilesLocation)
        ' Define variables based on input arrays
        Dim baseType = typeArray(0)
        Dim baseConnectionType = typeArray(1)
        'Dim lowerAxleType = typeArray(2)
        'Dim upperAxleType = typeArray(3)
        Dim lowerSupportRodType = typeArray(4)
        Dim upperSupportRodType = typeArray(5)
        Dim headSupportType = typeArray(6)
        Dim lampHeadType = typeArray(7)

        Dim baseHeight = parameterArray(0)
        Dim baseConnectionDiameter = parameterArray(1)
        Dim lowerAxleDiameter = parameterArray(2)
        Dim upperAxleDiameter = parameterArray(3)
        Dim lowerSupportDiameter = parameterArray(4)
        Dim upperSupportDiameter = parameterArray(5)
        Dim lampHeadDiameter = parameterArray(6)
        Dim lampHeadLength = parameterArray(7)

        ' Create an array list of product parts
        Dim productParts As ArrayList = New ArrayList()
        For Each part As Product In products
            productParts.Add(part.Name)
        Next

        ' Check if LampBase exists and if it does, check if it's the right version
        ' Delete and instantiate a new one and modify it accordingly if necessary.
        ' The same methodology is used for all parts.
        If Not productParts.Contains("LampBase_" & baseType) Then
            For Each part In products
                If part.Name.StartsWith("LampBase_") Then
                    DeletePart(documents, products, objectsToSelect, part)
                    SaveProduct(CATIA, tempFilesLocation)
                End If
            Next
            InstantiateLampBase(documents, products, objectsToSelect, reference, powerCopyFilesLocation, baseType, baseHeight)
        Else
            ModifyLampBase(documents, products, baseType, baseHeight)
        End If

        If Not productParts.Contains("BaseConnection_" & baseConnectionType) Then
            For Each part In products
                If part.Name.StartsWith("BaseConnection_") Then
                    DeletePart(documents, products, objectsToSelect, part)
                    SaveProduct(CATIA, tempFilesLocation)
                End If
            Next
            InstantiateBaseConnection(documents, products, objectsToSelect, reference, powerCopyFilesLocation, baseConnectionType, baseConnectionDiameter)
        Else
            ModifyBaseConnection(documents, products, baseConnectionType, baseConnectionDiameter)
        End If

        If Not productParts.Contains("LowerAxle") Then
            For Each part In products
                If part.Name.StartsWith("LowerAxle") Then
                    DeletePart(documents, products, objectsToSelect, products.Item("LowerAxle"))
                    SaveProduct(CATIA, tempFilesLocation)
                End If
            Next
            InstantiateLowerAxle(documents, products, objectsToSelect, reference, powerCopyFilesLocation, lowerAxleDiameter)
        Else
            ModifyLowerAxle(documents, products, lowerAxleDiameter)
        End If

        If Not productParts.Contains("UpperAxle") Then
            For Each part In products
                If part.Name.StartsWith("UpperAxle") Then
                    DeletePart(documents, products, objectsToSelect, products.Item("UpperAxle"))
                    SaveProduct(CATIA, tempFilesLocation)
                End If
            Next
            InstantiateUpperAxle(documents, products, objectsToSelect, reference, powerCopyFilesLocation, upperAxleDiameter)
        Else
            ModifyUpperAxle(documents, products, upperAxleDiameter)
        End If

        If Not productParts.Contains("LowerSupportRod1_" & lowerSupportRodType) Then
            For Each part In products
                If part.Name.StartsWith("LowerSupportRod") Then
                    DeletePart(documents, products, objectsToSelect, part)
                    SaveProduct(CATIA, tempFilesLocation)
                End If
            Next
            InstantiateLowerSupportRods(documents, products, objectsToSelect, reference, powerCopyFilesLocation, lowerSupportRodType, lowerSupportDiameter)
        Else
            ModifyLowerSupportRods(documents, products, lowerSupportRodType, lowerSupportDiameter)
        End If

        If Not productParts.Contains("UpperSupportRod_" & upperSupportRodType) Then
            For Each part In products
                If part.Name.StartsWith("UpperSupportRod_") Then
                    DeletePart(documents, products, objectsToSelect, part)
                    SaveProduct(CATIA, tempFilesLocation)
                End If
            Next
            InstantiateUpperSupportRod(documents, products, objectsToSelect, reference, powerCopyFilesLocation, upperSupportRodType, upperSupportDiameter)
        Else
            ModifyUpperSupportRod(documents, products, upperSupportRodType, upperSupportDiameter)
        End If

        If Not productParts.Contains("HeadSupport_" & headSupportType) Then
            For Each part In products
                If part.Name.StartsWith("HeadSupport_") Then
                    DeletePart(documents, products, objectsToSelect, part)
                    SaveProduct(CATIA, tempFilesLocation)
                End If
            Next
            InstantiateHeadSupport(documents, products, objectsToSelect, reference, powerCopyFilesLocation, headSupportType, upperSupportDiameter)
        Else
            ModifyHeadSupport(documents, products, headSupportType, upperSupportDiameter)
        End If

        If Not productParts.Contains("LampHead_" & lampHeadType) Then
            For Each part In products
                If part.Name.StartsWith("LampHead_") Then
                    DeletePart(documents, products, objectsToSelect, part)
                    SaveProduct(CATIA, tempFilesLocation)
                End If
            Next
            InstantiateLampHead(documents, products, objectsToSelect, reference, powerCopyFilesLocation, lampHeadType, lampHeadDiameter, lampHeadLength)
        Else
            ModifyLampHead(documents, products, lampHeadType, lampHeadDiameter, lampHeadLength)
        End If

    End Sub
    Public Sub InstantiateLampBase(documents, products, objectsToSelect, reference, fileLocation, type, baseHeight)
        ' Define how LampBase should be instantiated. If using different parts different powercopies need to be referenced.
        ' Different types probably need different references to be copied as well.
        Dim partName = "LampBase_" & type
        If type = "Oval" Then
            ' Specify powercopy part and powercopy
            Dim powerCopyPart As String = "PowerCopy_LampBase_Oval"
            Dim powerCopy As String = "LampBase"
            ' Create a reference DataTable that contains all the references that needs to be copied over to the new part
            ' First argument (reference) represents the partName of the part that has the reference
            ' Second argument ("Base") represents the name of the geoset where the reference that needs to be copied is included
            ' Third argument ("BaseEllipse") represents the name of the reference that needs to be copied over.
            Dim referencesDataTable As DataTable = CATIAFunctions.CreateReferenceDataTable()
            referencesDataTable.Rows.Add(reference, "Base", "BaseEllipse")
            referencesDataTable.Rows.Add(reference, "Base", "BaseHeightPlane")
            ' Call the function to add and instantiate the part
            CATIAFunctions.AddAndInstantiate(documents, products, objectsToSelect, partName, powerCopyPart, powerCopy, referencesDataTable, fileLocation)
        ElseIf type = "Box" Then
            ' Specify the powercopy part and powercopy name
            Dim powerCopyPart As String = "PowerCopy_LampBase_Box"
            Dim powerCopy As String = "LampBase"
            ' Create a DataTable containing all the references that need to be copied
            Dim referencesDataTable As DataTable = CATIAFunctions.CreateReferenceDataTable()
            referencesDataTable.Rows.Add(reference, "Base", "P1_Box")
            referencesDataTable.Rows.Add(reference, "Base", "P2_Box")
            referencesDataTable.Rows.Add(reference, "Base", "P3_Box")
            referencesDataTable.Rows.Add(reference, "Base", "P4_Box")
            referencesDataTable.Rows.Add(reference, "Base", "BaseHeightPlane")
            CATIAFunctions.AddAndInstantiate(documents, products, objectsToSelect, partName, powerCopyPart, powerCopy, referencesDataTable, fileLocation)
        Else
            ' If type is not Oval or Box, throw an exception
            Throw New Exception(type & " not found while trying to configure " & partName)
        End If
        ' After instantiation, parameters produced by the powercopy might need to be adjusted to fit the configuration.
        ModifyLampBase(documents, products, type, baseHeight)
    End Sub
    Public Sub ModifyLampBase(documents, products, type, baseHeight)
        ' Define the part name based on type
        Dim partName = "LampBase_" & type

        ' Check if the type is Oval or Box
        If type = "Oval" Then
            ' Modify the "radius" parameter based on the base height
            CATIAFunctions.ModifyParameters(documents, products, partName, "radius", baseHeight / 2)
            ' Update the part
            CATIAFunctions.UpdatePart(products, partName)
        ElseIf type = "Box" Then
            ' Modify the "radius" parameter based on the base height
            CATIAFunctions.ModifyParameters(documents, products, partName, "radius", baseHeight / 2)
            ' Update the part
            CATIAFunctions.UpdatePart(products, partName)
        Else
            ' If type is not Oval or Box, throw an exception
            Throw New Exception(type & " not found while trying to modify " & partName)
        End If
    End Sub
    ' Similar implementation for the rest of the products.
    Public Sub InstantiateBaseConnection(documents, products, objectsToSelect, reference, fileLocation, type, baseConnectionDiameter)
        Dim partName = "BaseConnection_" & type
        If type = "Round" Then
            Dim powerCopyPart As String = "PowerCopy_BaseConnection_Round"
            Dim powerCopy As String = "BaseConnection"
            Dim referencesDataTable As DataTable = CATIAFunctions.CreateReferenceDataTable()
            referencesDataTable.Rows.Add(reference, "Base", "BaseConnectionPoint_1")
            referencesDataTable.Rows.Add(reference, "LowerSupport", "BaseConnectionPoint_2")
            CATIAFunctions.AddAndInstantiate(documents, products, objectsToSelect, partName, powerCopyPart, powerCopy, referencesDataTable, fileLocation)
        Else
            Throw New Exception(type & " not found while trying to configure " & partName)
        End If
        ModifyBaseConnection(documents, products, type, baseConnectionDiameter)
    End Sub
    Public Sub ModifyBaseConnection(documents, products, type, baseConnectionDiameter)
        Dim partName = "BaseConnection_" & type
        If type = "Round" Then
            CATIAFunctions.ModifyParameters(documents, products, partName, "diameter", baseConnectionDiameter)
            CATIAFunctions.UpdatePart(products, partName)
        Else
            Throw New Exception(type & " not found while trying to modify " & partName)
        End If
    End Sub
    Public Sub InstantiateLowerAxle(documents, products, objectsToSelect, reference, fileLocation, lowerAxleDiameter)
        Dim partName = "LowerAxle"
        Dim powerCopyPart As String = "PowerCopy_HorizontalAxle"
        Dim powerCopy As String = "HorizontalAxle"
        Dim referencesDataTable As DataTable = CATIAFunctions.CreateReferenceDataTable()
        referencesDataTable.Rows.Add(reference, "LowerSupport", "LowerAxlePointLeft")
        referencesDataTable.Rows.Add(reference, "LowerSupport", "LowerAxlePointRight")
        CATIAFunctions.AddAndInstantiate(documents, products, objectsToSelect, partName, powerCopyPart, powerCopy, referencesDataTable, fileLocation)
        ModifyLowerAxle(documents, products, lowerAxleDiameter)
    End Sub
    Public Sub ModifyLowerAxle(documents, products, lowerAxleDiameter)
        Dim partName = "LowerAxle"
        CATIAFunctions.ModifyParameters(documents, products, partName, "diameter", lowerAxleDiameter)
        CATIAFunctions.UpdatePart(products, partName)
    End Sub

    Public Sub InstantiateUpperAxle(documents, products, objectsToSelect, reference, fileLocation, upperAxleDiameter)
        Dim partName = "UpperAxle"
        Dim powerCopyPart As String = "PowerCopy_HorizontalAxle"
        Dim powerCopy As String = "HorizontalAxle"
        Dim referencesDataTable As DataTable = CATIAFunctions.CreateReferenceDataTable()
        referencesDataTable.Rows.Add(reference, "UpperSupport", "UpperAxlePointLeft")
        referencesDataTable.Rows.Add(reference, "UpperSupport", "UpperAxlePointRight")
        CATIAFunctions.AddAndInstantiate(documents, products, objectsToSelect, partName, powerCopyPart, powerCopy, referencesDataTable, fileLocation)
        ModifyUpperAxle(documents, products, upperAxleDiameter)
    End Sub
    Public Sub ModifyUpperAxle(documents, products, upperAxleDiameter)
        Dim partName = "UpperAxle"
        CATIAFunctions.ModifyParameters(documents, products, partName, "diameter", upperAxleDiameter)
        CATIAFunctions.UpdatePart(products, partName)
    End Sub
    Public Sub InstantiateLowerSupportRods(documents, products, objectsToSelect, reference, fileLocation, type, lowerSupportDiameter)
        Dim parts = "LowerSupportRods"
        If type = "Round" Then
            For i = 1 To 2
                Dim partName = "LowerSupportRod" & i & "_" & type
                Dim powerCopyPart As String = "PowerCopy_LowerSupportRod_Round"
                Dim powerCopy As String = "LowerSupportRod"
                Dim referencesDataTable As DataTable = CATIAFunctions.CreateReferenceDataTable()
                If i = 1 Then
                    referencesDataTable.Rows.Add(reference, "LowerSupport", "LowerAxlePointLeft")
                    referencesDataTable.Rows.Add(reference, "UpperSupport", "UpperAxlePointLeft")
                Else
                    referencesDataTable.Rows.Add(reference, "LowerSupport", "LowerAxlePointRight")
                    referencesDataTable.Rows.Add(reference, "UpperSupport", "UpperAxlePointRight")
                End If
                CATIAFunctions.AddAndInstantiate(documents, products, objectsToSelect, partName, powerCopyPart, powerCopy, referencesDataTable, fileLocation)
            Next
        Else
            Throw New Exception(type & " not found while trying to configure " & parts)
        End If
        ModifyLowerSupportRods(documents, products, type, lowerSupportDiameter)
    End Sub
    Public Sub ModifyLowerSupportRods(documents, products, type, lowerSupportDiameter)
        Dim parts = "LowerSupportRods"
        If type = "Round" Then
            For i = 1 To 2
                Dim partName = "LowerSupportRod" & i & "_" & type
                CATIAFunctions.ModifyParameters(documents, products, partName, "diameter", lowerSupportDiameter)
                CATIAFunctions.ModifyParameters(documents, products, partName, "extraLength_OneSide", lowerSupportDiameter)
                CATIAFunctions.UpdatePart(products, partName)
            Next
        Else
            Throw New Exception(type & " not found while trying to modify " & parts)
        End If
    End Sub
    Public Sub InstantiateUpperSupportRod(documents, products, objectsToSelect, reference, fileLocation, type, upperSupportDiameter)
        Dim partName = "UpperSupportRod_" & type
        If type = "Round" Then
            Dim powerCopyPart As String = "PowerCopy_LowerSupportRod_Round"
            Dim powerCopy As String = "LowerSupportRod"
            Dim referencesDataTable As DataTable = CATIAFunctions.CreateReferenceDataTable()
            referencesDataTable.Rows.Add(reference, "UpperSupport", "UpperSupportRodPoint_1")
            referencesDataTable.Rows.Add(reference, "UpperSupport", "UpperSupportRodPoint_2")
            CATIAFunctions.AddAndInstantiate(documents, products, objectsToSelect, partName, powerCopyPart, powerCopy, referencesDataTable, fileLocation)
        Else
            Throw New Exception(type & " not found while trying to configure " & partName)
        End If
        ModifyUpperSupportRod(documents, products, type, upperSupportDiameter)
    End Sub
    Public Sub ModifyUpperSupportRod(documents, products, type, upperSupportDiameter)
        Dim partName = "UpperSupportRod_" & type
        If type = "Round" Then
            CATIAFunctions.ModifyParameters(documents, products, partName, "diameter", upperSupportDiameter)
            CATIAFunctions.ModifyParameters(documents, products, partName, "extraLength_OneSide", 0)
            CATIAFunctions.UpdatePart(products, partName)
        Else
            Throw New Exception(type & " not found while trying to modify " & partName)
        End If
    End Sub
    Public Sub InstantiateHeadSupport(documents, products, objectsToSelect, reference, fileLocation, type, upperSupportDiameter)
        Dim partName = "HeadSupport_" & type
        If type = "Round" Then
            Dim powerCopyPart As String = "PowerCopy_LampHolder"
            Dim powerCopy As String = "LampHolder"
            Dim referencesDataTable As DataTable = CATIAFunctions.CreateReferenceDataTable()
            referencesDataTable.Rows.Add(reference, "HeadSupport", "xy")
            referencesDataTable.Rows.Add(reference, "UpperSupport", "UpperSupportRodPoint_2")
            referencesDataTable.Rows.Add(reference, "HeadSupport", "HeadSupportPointLeft")
            referencesDataTable.Rows.Add(reference, "HeadSupport", "HeadSupportPointRight")
            CATIAFunctions.AddAndInstantiate(documents, products, objectsToSelect, partName, powerCopyPart, powerCopy, referencesDataTable, fileLocation)
        Else
            Throw New Exception(type & " not found while trying to configure " & partName)
        End If
        ModifyHeadSupport(documents, products, type, upperSupportDiameter)
    End Sub
    Public Sub ModifyHeadSupport(documents, products, type, upperSupportDiameter)
        Dim partName = "HeadSupport_" & type
        If type = "Round" Then
            CATIAFunctions.ModifyParameters(documents, products, partName, "diameter", upperSupportDiameter)
            CATIAFunctions.UpdatePart(products, partName)
        Else
            Throw New Exception(type & " not found while trying to modify " & partName)
        End If
    End Sub
    Public Sub InstantiateLampHead(documents, products, objectsToSelect, reference, fileLocation, type, lampHeadDiameter, lampHeadLength)
        Dim partName = "LampHead_" & type
        If type = "Type1" Then
            Dim powerCopyPart As String = "PowerCopy_LampHead"
            Dim powerCopy As String = "LampHead"
            Dim referencesDataTable As DataTable = CATIAFunctions.CreateReferenceDataTable()
            referencesDataTable.Rows.Add(reference, "LampHead", "xy")
            referencesDataTable.Rows.Add(reference, "LampHead", "yz")
            referencesDataTable.Rows.Add(reference, "LampHead", "zx")
            referencesDataTable.Rows.Add(reference, "LampHead", "P1")
            referencesDataTable.Rows.Add(reference, "HeadSupport", "HeadSupportPointRight")
            referencesDataTable.Rows.Add(reference, "HeadSupport", "HeadSupportPointMiddle")
            CATIAFunctions.AddAndInstantiate(documents, products, objectsToSelect, partName, powerCopyPart, powerCopy, referencesDataTable, fileLocation)
        Else
            Throw New Exception(type & " not found while trying to configure " & partName)
        End If
        ModifyLampHead(documents, products, type, lampHeadDiameter, lampHeadLength)
    End Sub
    Public Sub ModifyLampHead(documents, products, type, lampHeadDiameter, lampHeadLength)
        Dim partName = "LampHead_" & type
        If type = "Type1" Then
            CATIAFunctions.ModifyParameters(documents, products, partName, "diameter", 10)
            CATIAFunctions.ModifyParameters(documents, products, partName, "headDiameter", lampHeadDiameter)
            CATIAFunctions.ModifyParameters(documents, products, partName, "headLength", lampHeadLength)
            CATIAFunctions.UpdatePart(products, partName)
        Else
            Throw New Exception(type & " not found while trying to modify " & partName)
        End If
    End Sub
End Module
