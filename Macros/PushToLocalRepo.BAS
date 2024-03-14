Sub Main

    Dim theDiagram As Diagram
    Dim logicalModel As Model
    Dim physicalModel As Model
    Dim theDDLGenerator As DDLGeneration
    Dim physicalModelGenerator As PhysicalGeneration
    Dim model As Model

    Set fs = CreateObject("Scripting.FileSystemObject")


    Set theDiagram = DiagramManager.ActiveDiagram
    theDiagram.SaveFile("MyLocalCopy.DM1")
    fs.CopyFile("MyLocalCopy.DM1", "RepoCopy.DM1")
    Set logicalModel = theDiagram.ActiveModel

    For Each model In theDiagram.Models
    	If model.Name = "CoorsTekDW" Then
    		theDiagram.Models.Remove(model.ID)
    		Exit For
    	End If
    Next

	Set physicalModelGenerator = logicalModel.PhysicalGenerationObject
	physicalModelGenerator.GenPhysicalUsingFileBasedQuickLaunch("{home}\AppData\Roaming\Idera\ERStudio\XML\GeneratePhysicalModel.gpo")

	For Each model In theDiagram.Models
    	If model.Name = "CoorsTekDW" Then
    		Set physicalModel = model
    		Exit For
    	End If
    Next

	Set theDDLGenerator = physicalModel.DDLGenerationObject
	theDDLGenerator.GenDDLUsingFileBasedQuickLaunch("{home}\AppData\Roaming\Idera\ERStudio\XML\GenerateLocalDeploymentScript.ddo")
	theDDLGenerator.GenDDLUsingFileBasedQuickLaunch("{home}\AppData\Roaming\Idera\ERStudio\XML\GenerateDDL.ddo")

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
