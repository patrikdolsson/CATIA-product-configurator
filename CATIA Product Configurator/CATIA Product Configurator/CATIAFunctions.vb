'Imports System.Linq
'Imports System.Xml.Linq
Imports INFITF
Imports KnowledgewareTypeLib
Imports MECMOD
Imports ProductStructureTypeLib

Module CATIAFunctions
    Public Sub ModifyParameters(documents As Documents, products As Products, partName As String, parameterName As String, parameterValue As Object)
        ' Modify the parameter called parameterName inside the part named partName
        Dim part As Part = documents.Item(products.Item(partName).PartNumber & ".CATPart").Part

        Dim partParameters As Parameters = part.Parameters
        Dim partParameter As Parameter = partParameters.Item(parameterName)
        partParameter.value = parameterValue
    End Sub
    Public Sub ModifyParameters(part As Part, parameterName As String, parameterValue As Object)
        ' Modify the parameter called parameterName inside the part
        Dim partParameters = part.Parameters
        Dim partParameter = partParameters.Item(parameterName)
        partParameter.value = parameterValue
    End Sub

    Public Sub CopyReferencesToPaste(documents1 As Object, products1 As Products, reference As String, geoSet As String, itemToCopyString As String, objectsToSelect As Object)
        ' This function only serves to copy the desired reference to the objectsToSelect object.
        Dim referencePart = documents1.item(products1.Item(reference).PartNumber & ".CATPart")

        objectsToSelect.Clear()

        ' Set the list of the hybridshapes inside the desired geometrical set
        Dim geoSetList = referencePart.Part.FindObjectByName(geoSet).HybridShapes

        If itemToCopyString.EndsWith("AxisSystem") Then
            ' Copies any axis system inside the reference part named according to itemToCopyString, but only if it ends with "AxisSystem"
            objectsToSelect.Search("'Part Design'.'Axis System'.Name=" & itemToCopyString & ";all")
            Console.WriteLine(objectsToSelect.Count2)
            Dim itemToCopy = objectsToSelect.Item2(1).Value
            objectsToSelect.Clear()
            objectsToSelect.Add(itemToCopy)
            objectsToSelect.Copy()
            objectsToSelect.Clear()
        Else
            ' Copies the hybridShape called itemToCopyString inside the geoSet of the reference part
            Dim itemToCopy = geoSetList.Item(itemToCopyString)
            objectsToSelect.Add(itemToCopy)
            objectsToSelect.Copy()
            objectsToSelect.Clear()
        End If

    End Sub

    Public Sub CopyReferences(documents1 As Object, products1 As Products, partName As String, references As DataTable, objectsToSelect As Object)
        Dim pasteInPart As PartDocument = documents1.Item(products1.Item(partName).PartNumber & ".CATPart")
        objectsToSelect.Clear()
        ' Define place that references should be pasted to
        Dim pasteIn = pasteInPart.Part.HybridBodies

        ' Attempt to copy and paste all references as defined by the each DataRow in the DataTable called references
        For Each row As DataRow In references.Rows
            ' Define the elements from the DataRow
            Dim referencePart = row.Field(Of String)(0)
            Dim geoSet = row.Field(Of String)(1)
            Dim itemToCopyString = row.Field(Of String)(2)

            ' Set the label statement that can be used further down with the GoTo Statement to retry copying and
            ' pasting in case it doesn't wotk the first time
H1:
            CopyReferencesToPaste(documents1, products1, referencePart, geoSet, itemToCopyString, objectsToSelect)
            objectsToSelect.Add(pasteIn)
            objectsToSelect.PasteSpecial("CATPrtResult")
            objectsToSelect.Clear()

            ' Attempt to access the pasted item. If they're not accessible, it means that the copy and paste operation
            ' failed and it needs to be redone 
            Try
                Dim externalReferences = pasteIn.Item("External References")
                Dim pastedReference = externalReferences.HybridShapes.Item(itemToCopyString)
            Catch ex1 As Exception
                Try
                    If itemToCopyString.EndsWith("AxisSystem") Then
                        pasteInPart.Part.AxisSystems.Item(itemToCopyString)
                    Else
                        Console.WriteLine("SAVEEED")
                        GoTo H1
                    End If
                Catch ex2 As Exception
                    GoTo H1
                End Try
            End Try

        Next
    End Sub

    Public Function CreateReferenceDataTable()
        Dim table As New DataTable
        table.Columns.Add("referencePart", GetType(String))
        table.Columns.Add("geoSet", GetType(String))
        table.Columns.Add("reference", GetType(String))
        Return table
    End Function

    'Public Sub CopyParameters(documents1 As Object, Reference As String, GeoSet As String, Objects2select As Object)
    '    Dim Reference_Part = documents1.Item(Reference & ".CATPart")
    '    Dim parameters_to_copy = Reference_Part.Part.Parameters.RootParameterSet.ParameterSets.item(GeoSet).DirectParameters 'FindObjectByName(GeoSet).ParameterSets
    '    Objects2select.Clear()

    '    For i As Int32 = 1 To parameters_to_copy.Count
    '        Dim item_to_Copy = parameters_to_copy.Item(i)
    '        Objects2select.Add(item_to_Copy)
    '    Next i

    '    Objects2select.Copy()
    '    Objects2select.Clear()
    'End Sub
    Public Sub AddAndInstantiate(documents As Documents, products As Products, objectsToSelect As Object, partName As String,
                                 powerCopyPart As String, powerCopy As String, referencesDataTable As DataTable, fileLocation As String)
        ' This function serves to add a new part, copy all of the references needed, and instantiate a single powercopy to
        ' the new part utilizing the copied references.
        AddPart(products, partName)
        CopyReferences(documents, products, partName, referencesDataTable, objectsToSelect)
        Dim powerCopyLocation = fileLocation & powerCopyPart & ".CATPart"
        Dim powerCopyToUse As String = powerCopy & "_PC"
        InstantiatePC(documents, products, partName, powerCopyLocation, powerCopyPart, powerCopyToUse)

        objectsToSelect.Clear()
        Dim part As Part = documents.Item(products.Item(partName).PartNumber & ".CATPart").Part

        ' Hide the standard xy, yz, and zx plane
        Dim xyplane = part.FindObjectByName("xy plane")
        Dim zxplane = part.FindObjectByName("zx plane")
        Dim yzplane = part.FindObjectByName("yz plane")
        objectsToSelect.Add(xyplane)
        objectsToSelect.VisProperties.SetShow(1)
        objectsToSelect.Clear()
        objectsToSelect.Add(zxplane)
        objectsToSelect.VisProperties.SetShow(1)
        objectsToSelect.Clear()
        objectsToSelect.Add(yzplane)
        objectsToSelect.VisProperties.SetShow(1)
        objectsToSelect.Clear()
    End Sub

    Public Sub AddPart(products As Products, partName As String)
        Dim newPart = products.AddNewComponent("Part", "")
        newPart.Name = partName
    End Sub

    Public Sub InstantiatePC(documents, products, partName, powerCopyLocation, powerCopyPart, powerCopy)
        ' This function serves to instantiate the powerCopy into the part called partName utilizing copied
        ' references inside the geometrical set "External References".

        ' Define the part where the powercopy should be instantiated to
        Dim iPart As PartDocument = documents.Item(products.Item(partName).PartNumber & ".CATPart")
        Dim iPartPart = iPart.Part

        ' Start the powercopy instantiation
        Dim factory = iPartPart.GetCustomerFactory("InstanceFactory")
        factory.BeginInstanceFactory(powerCopy, powerCopyLocation)
        factory.BeginInstantiate

        ' Set the list of shapes that will be matched with powercopy inputs
        Dim shapesToSet = iPartPart.HybridBodies.Item("External References").HybridShapes

        Dim itemsToSet As ArrayList = New ArrayList()

        For Each shape In shapesToSet
            itemsToSet.Add(shape)
        Next

        ' If any axis systems have been pasted to this part they will be added to the list of items
        ' that will be matched with powercopy inputs.
        For Each axisSystem In iPartPart.AxisSystems
            itemsToSet.Add(axisSystem)
        Next

        ' Get the names of the powercopy inputs. Note that axis systems that are to be used to instantiate
        ' powercopies must be placed in the reference geometrical sets in the powercopy part
        Dim powerCopyDoc As Document = documents.Item(powerCopyPart & ".CATPart")
        Dim referenceGeoSetPowerCopy = powerCopyDoc.Part.HybridBodies.Item("References").HybridShapes

        ' Match the power copy inputs to the contextual items to set in the part.
        For i = 1 To itemsToSet.Count
            factory.PutInputData(referenceGeoSetPowerCopy.Item(i).Name, itemsToSet.Item(i - 1))
        Next

        Dim Instance = factory.Instantiate
        factory.EndInstantiate
        factory.EndInstanceFactory

        ' Attempt an update
        Try
            iPartPart.Update()
        Catch ex As Exception

        End Try
    End Sub

    Public Sub UpdatePart(products, partName)
        Dim part = products.Item(partName)
        'Dim part = partToUpdate.Part
        part.Update()
    End Sub
    Public Sub UpdateAllParts(products)
        For Each part As Product In products
            part.Update()
        Next
    End Sub
    Public Sub RemoveDelWarning(documents, products, objectsToSelect, partToDelete)
        Dim externalref = documents.Item(products.Item(partToDelete).PartNumber & ".CATPart").Part.FindObjectByName("External References")
        objectsToSelect.clear()
        objectsToSelect.add(externalref)
        objectsToSelect.delete()
    End Sub

    Public Sub RemoveAxisSystems(documents, products, objectsToSelect, partToDelete)
        For i = 1 To documents.Item(products.Item(partToDelete).PartNumber & ".CATPart").Part.AxisSystems.Count
            Dim axisSystem = documents.Item(products.Item(partToDelete).PartNumber & ".CATPart").Part.AxisSystems.Item(1)
            objectsToSelect.clear()
            objectsToSelect.add(axisSystem)
            objectsToSelect.delete()
        Next
    End Sub
    Public Sub DeletePart(documents1, products1, objectsToSelect, partToDelete)
        Try
            RemoveDelWarning(documents1, products1, objectsToSelect, partToDelete.Name)
        Catch ex As Exception
        End Try
        Try
            RemoveAxisSystems(documents1, products1, objectsToSelect, partToDelete.Name)
        Catch ex As Exception

        End Try
        objectsToSelect.Clear()
        objectsToSelect.add(partToDelete)
        objectsToSelect.delete()
        objectsToSelect.Clear()
    End Sub

    Public Sub ClearPreviousInstantiations(CATIA, documents, products, objectsToSelect, tempFilesLocation)
        Try
            For i = 2 To products.Count
                Dim partToDelete = products.Item(2)
                DeletePart(documents, products, objectsToSelect, partToDelete)
            Next
        Catch ex As Exception
        End Try
        SaveProduct(CATIA, tempFilesLocation)
    End Sub

    Public Sub SaveProduct(CATIA, tempFilesLocation)
        If Not Directory.Exists(tempFilesLocation) Then
            Directory.CreateDirectory(tempFilesLocation)
        End If
        CATIA.DisplayFileAlerts = False
        CATIA.ActiveDocument.SaveAs(tempFilesLocation & "TableLamp.CATProduct")
        CATIA.DisplayFileAlerts = True
    End Sub
End Module
