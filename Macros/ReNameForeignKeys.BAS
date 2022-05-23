'#Language "WWB-COM"

Option Explicit

Sub Main
	Dim MyDiagram As Diagram
	Dim MyModel As Model
	Dim MySubModel As SubModel
	Dim MyDictionary As Dictionary
	Dim MyRelationship As Relationship

	Dim MySelectedObject As SelectedObject

	' Set base variables
	Set MyDiagram = DiagramManager.ActiveDiagram
	Set MyModel = MyDiagram.ActiveModel
	Set MySubModel = MyModel.ActiveSubModel
	Set MyDictionary = MyDiagram.Dictionary

	If MySubModel.SelectedObjects.Count > 0 Then
		For Each MySelectedObject In MySubModel.SelectedObjects
			
			If MySelectedObject.Type = 2 Then '2=Relationships
				Set MyRelationship = MyModel.Relationships.Item(MySelectedObject.ID)
				ReNameRelationship(MyRelationship)
			End If

		Next
	Else
		For Each MyRelationship In MyModel.Relationships
			ReNameRelationship(MyRelationship)
		Next
		
	End If
End Sub
Function ReNameRelationship(ByVal MyRelationship As Relationship)
	Dim st As String
	Dim MyFKColumnPair As FKColumnPair
	st = "FK_"
	st = st & MyRelationship.ChildEntity.EntityName
	For Each MyFKColumnPair In MyRelationship.FKColumnPairs
		st = st & "_" & MyRelationship.FKColumnPairs.Item(MyFKColumnPair.ID).ChildAttribute.RoleName
	Next
	'Debug.Print st
	MyRelationship.Name = st '"FK_" + MyRelationship.ChildEntity.EntityName + "_" + MyRelationship.FKColumnPairs.Item(1).ChildAttribute.RoleName
	MyRelationship.BusinessName = MyRelationship.Name
End Function
