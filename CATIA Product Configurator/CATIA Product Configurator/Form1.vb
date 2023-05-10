'Imports System.IO
'Imports System.Text
'Imports AnnotationTypeLib
'Imports DNBIgpTagPath
'Imports DRAFTINGITF
'Imports INFITF
'Imports KnowledgewareTypeLib
'Imports MECMOD
'Imports PROCESSITF
Imports ProductStructureTypeLib
Public Class Form1


    Public Shared CATIA As INFITF.Application = GetObject(vbNullString, "Catia.Application")
    'Create references to the active document and selection
    Public Shared objectsToSelect = CATIA.ActiveDocument.Selection
    Public Shared documents = CATIA.Documents

    'File locations for the CAD, STL and temporary files
    Public Shared cadFilesLocation As String = Functions.getFileLocation("CAD")
    Public Shared stlFilesLocation As String = Functions.getFileLocation("STL")
    Public Shared tempFilesLocation As String = Functions.getFileLocation("temp")

    'Load the product document
    Dim productDocument As ProductDocument = documents.Item("TableLamp.CATProduct")

    'Default reference name for the lamp
    Public Shared referenceName As String = "LampReference"

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Load the current product and its products
        Dim product As Product = productDocument.Product
        Dim products As Products = product.Products

        ' Get Supported Types
        Dim acceptableTypes = ConfigureFunctions.GetAcceptableTypes()
        Dim acceptableBaseTypes As String() = acceptableTypes(0)
        Dim acceptableBaseConnectionTypes As String() = acceptableTypes(1)
        Dim acceptableLowerSupportRodTypes As String() = acceptableTypes(2)
        Dim acceptableUpperSupportRodTypes As String() = acceptableTypes(3)
        Dim acceptableHeadSupportTypes As String() = acceptableTypes(4)
        Dim acceptableLampHeadTypes As String() = acceptableTypes(5)

        ' Load types from the configurator
        Dim baseType As String = ComboBox1.SelectedItem
        Dim baseConnectionType As String = ComboBox6.SelectedItem
        Dim lowerSupportRodType As String = ComboBox2.SelectedItem
        Dim upperSupportRodType As String = ComboBox3.SelectedItem
        Dim headSupportType As String = ComboBox4.SelectedItem
        Dim lampHeadType As String = ComboBox5.SelectedItem

        ' Set default type if type is not specified correctly
        If Not acceptableBaseTypes.Contains(baseType) Then
            baseType = "Oval"
        End If

        If Not acceptableBaseConnectionTypes.Contains(baseConnectionType) Then
            baseConnectionType = "Round"
        End If

        If Not acceptableLowerSupportRodTypes.Contains(lowerSupportRodType) Then
            lowerSupportRodType = "Round"
        End If

        If Not acceptableUpperSupportRodTypes.Contains(upperSupportRodType) Then
            upperSupportRodType = "Round"
        End If

        If Not acceptableHeadSupportTypes.Contains(headSupportType) Then
            headSupportType = "Round"
        End If

        If Not acceptableLampHeadTypes.Contains(lampHeadType) Then
            lampHeadType = "Type1"
        End If

        ' Load parameters from the configurator
        Dim baseWidth As Double = NumericUpDown1.Value
        Dim baseDepth As Double = NumericUpDown2.Value
        Dim baseHeight As Double = NumericUpDown3.Value
        Dim baseConnectionHeight As Double = NumericUpDown7.Value
        Dim baseConnectionDiameter As Double = NumericUpDown4.Value
        Dim lampSwivelAngle As Double = NumericUpDown12.Value

        Dim lowerConnectionWidth As Double = NumericUpDown18.Value
        Dim upperConnectionWidth As Double = NumericUpDown17.Value
        Dim lowerAxleDiameter As Double = NumericUpDown45.Value
        Dim upperAxleDiameter As Double = NumericUpDown46.Value
        Dim lowerSupportHeight As Double = NumericUpDown14.Value
        Dim lowerSupportDiameter As Double = NumericUpDown47.Value
        Dim lowerAxleAngle As Double = NumericUpDown19.Value

        Dim upperSupportHeight As Double = NumericUpDown25.Value
        Dim upperSupportDiameter As Double = NumericUpDown48.Value
        Dim upperAxleAngle As Double = NumericUpDown20.Value

        Dim headSupportWidth As Double = NumericUpDown22.Value
        Dim headSupportHeight As Double = NumericUpDown21.Value
        Dim headSupportAngle As Double = NumericUpDown10.Value

        Dim lampHeadDiameter As Double = NumericUpDown24.Value
        Dim lampHeadLength As Double = NumericUpDown23.Value
        Dim lampHeadAngle As Double = NumericUpDown13.Value

        'Explicitly set the values for the different arrays used
        Dim typeArray(7) As String
        typeArray(0) = baseType
        typeArray(1) = baseConnectionType
        typeArray(2) = ""
        typeArray(3) = ""
        typeArray(4) = lowerSupportRodType
        typeArray(5) = upperSupportRodType
        typeArray(6) = headSupportType
        typeArray(7) = lampHeadType

        Dim angleArray(4) As Double
        angleArray(0) = lampSwivelAngle
        angleArray(1) = lowerAxleAngle
        angleArray(2) = upperAxleAngle
        angleArray(3) = headSupportAngle
        angleArray(4) = lampHeadAngle

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

        Dim partParameterArray(7) As Double
        partParameterArray(0) = baseHeight
        partParameterArray(1) = baseConnectionDiameter
        partParameterArray(2) = lowerAxleDiameter
        partParameterArray(3) = upperAxleDiameter
        partParameterArray(4) = lowerSupportDiameter
        partParameterArray(5) = upperSupportDiameter
        partParameterArray(6) = lampHeadDiameter
        partParameterArray(7) = lampHeadLength

        ' Modify angles, reference parameters and then instantiate or update the parts as specified
        ConfigureFunctions.ModifyAngles(documents, products, referenceName, angleArray)
        ConfigureFunctions.ModifyReferenceParameters(documents, products, referenceName, referenceParameterArray)
        ConfigureFunctions.ConfigureParts(CATIA, documents, products, objectsToSelect, referenceName, cadFilesLocation, typeArray, partParameterArray, tempFilesLocation)
        CATIAFunctions.UpdateAllParts(products)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' Clear previous instantiations
        Dim productDocument As ProductDocument = documents.Item("TableLamp.CATProduct")
        Dim product = productDocument.Product
        Dim products = product.Products

        CATIAFunctions.ClearPreviousInstantiations(CATIA, documents, products, objectsToSelect, tempFilesLocation)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim productDocument As ProductDocument = documents.Item("TableLamp.CATProduct")
        Dim product = productDocument.Product
        Dim products = product.Products

        ' Define variables to hold the minimum, maximum, and step values for each of the numeric input fields
        ' Note: you don't need to chose this through the configurator
        Dim baseWidth_Min = NumericUpDown26.Value
        Dim baseWidth_Max = NumericUpDown27.Value
        Dim baseWidth_Steps = NumericUpDown49.Value

        Dim baseDepth_Min = NumericUpDown11.Value
        Dim baseDepth_Max = NumericUpDown28.Value
        Dim baseDepth_Steps = NumericUpDown51.Value

        Dim baseHeight_Min = NumericUpDown8.Value
        Dim baseHeight_Max = NumericUpDown30.Value
        Dim baseHeight_Steps = NumericUpDown55.Value

        Dim baseConnectionHeight_Min = NumericUpDown9.Value
        Dim baseConnectionHeight_Max = NumericUpDown29.Value
        Dim baseConnectionHeight_Steps = NumericUpDown53.Value

        Dim baseConnectionDiameter_Min = NumericUpDown6.Value
        Dim baseConnectionDiameter_Max = NumericUpDown31.Value
        Dim baseConnectionDiameter_Steps = NumericUpDown57.Value

        Dim lowerConnectionWidth_Min = NumericUpDown33.Value
        Dim lowerConnectionWidth_Max = NumericUpDown34.Value
        Dim lowerConnectionWidth_Steps = NumericUpDown50.Value

        Dim upperConnectionWidth_Min = NumericUpDown35.Value
        Dim upperConnectionWidth_Max = NumericUpDown36.Value
        Dim upperConnectionWidth_Steps = NumericUpDown52.Value

        Dim lowerAxleDiameter_Min = NumericUpDown5.Value
        Dim lowerAxleDiameter_Max = NumericUpDown32.Value
        Dim lowerAxleDiameter_Steps = NumericUpDown43.Value

        Dim upperAxleDiameter_Min = NumericUpDown44.Value
        Dim upperAxleDiameter_Max = NumericUpDown59.Value
        Dim upperAxleDiameter_Steps = NumericUpDown60.Value

        Dim lowerSupportHeight_Min = NumericUpDown41.Value
        Dim lowerSupportHeight_Max = NumericUpDown42.Value
        Dim lowerSupportHeight_Steps = NumericUpDown58.Value

        Dim lowerSupportDiameter_Min = NumericUpDown61.Value
        Dim lowerSupportDiameter_Max = NumericUpDown62.Value
        Dim lowerSupportDiameter_Steps = NumericUpDown63.Value

        Dim upperSupportHeight_Min = NumericUpDown64.Value
        Dim upperSupportHeight_Max = NumericUpDown66.Value
        Dim upperSupportHeight_Steps = NumericUpDown67.Value

        Dim upperSupportDiameter_Min = NumericUpDown65.Value
        Dim upperSupportDiameter_Max = NumericUpDown68.Value
        Dim upperSupportDiameter_Steps = NumericUpDown69.Value

        Dim headSupportWidth_Min = NumericUpDown70.Value
        Dim headSupportWidth_Max = NumericUpDown73.Value
        Dim headSupportWidth_Steps = NumericUpDown75.Value

        Dim headSupportHeight_Min = NumericUpDown72.Value
        Dim headSupportHeight_Max = NumericUpDown78.Value
        Dim headSupportHeight_Steps = NumericUpDown80.Value

        Dim lampHeadDiameter_Min = NumericUpDown71.Value
        Dim lampHeadDiameter_Max = NumericUpDown76.Value
        Dim lampHeadDiameter_Steps = NumericUpDown77.Value

        Dim lampHeadLength_Min = NumericUpDown74.Value
        Dim lampHeadLength_Max = NumericUpDown79.Value
        Dim lampHeadLength_Steps = NumericUpDown81.Value

        ' Define arrays for each configuration setting
        Dim baseWidthArray = {baseWidth_Min, baseWidth_Max, baseWidth_Steps}
        Dim baseDepthArray = {baseDepth_Min, baseDepth_Max, baseDepth_Steps}
        Dim baseHeightArray = {baseHeight_Min, baseHeight_Max, baseHeight_Steps}
        Dim baseConnectionHeightArray = {baseConnectionHeight_Min, baseConnectionHeight_Max, baseConnectionHeight_Steps}
        Dim baseConnectionDiameterArray = {baseConnectionDiameter_Min, baseConnectionDiameter_Max, baseConnectionDiameter_Steps}
        Dim lampSwivelAngleArray = {70, -180, 180, 1}
        Dim lowerConnectionWidthArray = {lowerConnectionWidth_Min, lowerConnectionWidth_Max, lowerConnectionWidth_Steps}
        Dim upperConnectionWidthArray = {upperConnectionWidth_Min, upperConnectionWidth_Max, upperConnectionWidth_Steps}
        Dim lowerAxleDiameterArray = {lowerAxleDiameter_Min, lowerAxleDiameter_Max, lowerAxleDiameter_Steps}
        Dim lowerAxleAngleArray = {-30, -90, 90, 1}
        Dim upperAxleDiameterArray = {upperAxleDiameter_Min, upperAxleDiameter_Max, upperAxleDiameter_Steps}
        Dim upperAxleAngleArray = {90, -150, 150, 1}
        Dim lowerSupportHeightArray = {lowerSupportHeight_Min, lowerSupportHeight_Max, lowerSupportHeight_Steps}
        Dim lowerSupportDiameterArray = {lowerSupportDiameter_Min, lowerSupportDiameter_Max, lowerSupportDiameter_Steps}
        Dim upperSupportHeightArray = {upperSupportHeight_Min, upperSupportHeight_Max, upperSupportHeight_Steps}
        Dim upperSupportDiameterArray = {upperSupportDiameter_Min, upperSupportDiameter_Max, upperSupportDiameter_Steps}
        Dim headSupportWidthArray = {headSupportWidth_Min, headSupportWidth_Max, headSupportWidth_Steps}
        Dim headSupportHeightArray = {headSupportHeight_Min, headSupportHeight_Max, headSupportHeight_Steps}
        Dim headSupportAngleArray = {-50, -180, 180, 1}
        Dim lampHeadDiameterArray = {lampHeadDiameter_Min, lampHeadDiameter_Max, lampHeadDiameter_Steps}
        Dim lampHeadLengthArray = {lampHeadLength_Min, lampHeadLength_Max, lampHeadLength_Steps}
        Dim lampHeadAngleArray = {70, -120, 120, 1}

        ' Save the configuration setting to the stlExportingSettingsArray
        Dim stlExportingSettingsArray(21) As Array
        stlExportingSettingsArray(0) = baseWidthArray
        stlExportingSettingsArray(1) = baseDepthArray
        stlExportingSettingsArray(2) = baseHeightArray
        stlExportingSettingsArray(3) = baseConnectionHeightArray
        stlExportingSettingsArray(4) = baseConnectionDiameterArray
        stlExportingSettingsArray(5) = lampSwivelAngleArray
        stlExportingSettingsArray(6) = lowerConnectionWidthArray
        stlExportingSettingsArray(7) = upperConnectionWidthArray
        stlExportingSettingsArray(8) = lowerAxleDiameterArray
        stlExportingSettingsArray(9) = lowerAxleAngleArray
        stlExportingSettingsArray(10) = upperAxleDiameterArray
        stlExportingSettingsArray(11) = upperAxleAngleArray
        stlExportingSettingsArray(12) = lowerSupportHeightArray
        stlExportingSettingsArray(13) = lowerSupportDiameterArray
        stlExportingSettingsArray(14) = upperSupportHeightArray
        stlExportingSettingsArray(15) = upperSupportDiameterArray
        stlExportingSettingsArray(16) = headSupportWidthArray
        stlExportingSettingsArray(17) = headSupportHeightArray
        stlExportingSettingsArray(18) = headSupportAngleArray
        stlExportingSettingsArray(19) = lampHeadDiameterArray
        stlExportingSettingsArray(20) = lampHeadLengthArray
        stlExportingSettingsArray(21) = lampHeadAngleArray



        ' Keep all angles at 0
        Dim zeroAngleArray = {0, 0, 0, 0, 0}
        ConfigureFunctions.ModifyAngles(documents, products, referenceName, zeroAngleArray)

        ' Get ExportSTLConfigurationArray and run the exporting function writeJsonAndExportSTL() 
        Dim exportSTLConfigurationArray = ExportSTL.GetExportSTLConfigurationsArray(stlExportingSettingsArray)
        JSONWriter.writeJsonAndExportSTL(CATIA, documents, products, objectsToSelect, cadFilesLocation, stlFilesLocation, tempFilesLocation, referenceName, exportSTLConfigurationArray, stlExportingSettingsArray)

        ' Run the default configuration. This changes the angles back according to the configurator.
        Button1_Click(sender, e)
    End Sub
End Class

