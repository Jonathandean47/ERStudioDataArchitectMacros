Sub Main

    Dim theDiagram As Diagram
    Dim logicalModel As Model
    Dim physicalModel As Model
    Dim theDDLGenerator As DDLGeneration
    Dim physicalModelGenerator As PhysicalGeneration
    Dim model As Model

    Set fs = CreateObject("Scripting.FileSystemObject")


    Set theDiagram = DiagramManager.ActiveDiagram
    theDiagram.SaveFile("C:\Users\jonathan.dean\OneDrive - Neudesic\Documents\Neudesic\CoorsTek\CoorsTek Data Warehouse_Detached.DM1")
    fs.CopyFile("C:\Users\jonathan.dean\OneDrive - Neudesic\Documents\Neudesic\CoorsTek\CoorsTek Data Warehouse_Detached.DM1", "C:\Users\jonathan.dean\Git\Data Warehouse\CoorsTek Data Warehouse.DM1")
    Set logicalModel = theDiagram.ActiveModel

    For Each model In theDiagram.Models
    	If model.Name = "CoorsTekDW" Then
    		theDiagram.Models.Remove(model.ID)
    		Exit For
    	End If
    Next

	Set physicalModelGenerator = logicalModel.PhysicalGenerationObject
	physicalModelGenerator.GenPhysicalUsingFileBasedQuickLaunch("C:\Users\jonathan.dean\AppData\Roaming\Idera\ERStudio\XML\GeneratePhysicalModel.gpo")

	For Each model In theDiagram.Models
    	If model.Name = "CoorsTekDW" Then
    		Set physicalModel = model
    		Exit For
    	End If
    Next

	Set theDDLGenerator = physicalModel.DDLGenerationObject
	theDDLGenerator.GenDDLUsingFileBasedQuickLaunch("C:\Users\jonathan.dean\AppData\Roaming\Idera\ERStudio\XML\GenerateLocalDeploymentScript.ddo")
	theDDLGenerator.GenDDLUsingFileBasedQuickLaunch("C:\Users\jonathan.dean\AppData\Roaming\Idera\ERStudio\XML\GenerateDDL.ddo")

    For Each model In theDiagram.Models
    	If model.Name = "CoorsTekDW" Then
    		theDiagram.Models.Remove(model.ID)
    		Exit For
    	End If
    Next

    'theDDLGenerator.GenerateDatabaseMethod = MULTIPLE_FILES
    'theDDLGenerator.GenerateAllEntities
    'theDDLGenerator.CreationScriptFilePath = "C:\demo"
    'theDDLGenerator.GenerateDDL

    'MsgBox "Done"

End Sub
