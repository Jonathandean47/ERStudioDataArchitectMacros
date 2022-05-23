' MACRO TITLE: ADD BASE ATTRIBUTES TO PERSON ENTITY
' This macro adds base attributes to selected entities, which 
' represent people. It will also introduce a primary key based
' upon the entity's name.
'----------------------------------------------------------------

Sub Main
    Dim MyDiagram As Diagram
	Dim MyModel As Model
	Dim MySubModel As SubModel
	Dim MyDictionary As Dictionary
	Dim MyDomain As Domains
	Dim MyEntity As Entity
	Dim MyAttribute As AttributeObj
	'Dim MyPrimaryKey As Index
	'Dim MyAlternateKey As Index
	Dim AttributeName As String
	Dim PrimaryKeyName As String
	Dim AlternateKeyName As String
	Dim EntityName As String
	Dim Logical As Boolean
	Dim MySelObject As SelectedObject
	Dim SmallKeyID As Integer
	Dim ShortNameID As Integer
	Dim MetaIsActiveID As Integer
	Dim MetaIsDeletedID As Integer
	Dim MetaHashbyteValueID As Integer
	Dim MetaCreatedDateID As Integer
	Dim MetaModifiedDateID As Integer
	Dim MetaInsertAuditKeyID As Integer
	Dim MetaUpdateAuditKeyID As Integer
	Dim MyRelationship As Relationship
	Dim MyFKColumnPair As FKColumnPair
	Dim SystemKeyID As Integer
	Const RelationshipTypeNonIdentifyingMandatory As Integer = 1

    Dim ID As Integer
	Dim ObjType As Integer
	Dim MyIndex As Index

    ' Set base variables
	Set MyDiagram = DiagramManager.ActiveDiagram
	Set MyModel = MyDiagram.ActiveModel
	Set MySubModel = MyModel.ActiveSubModel
	Set MyDictionary = MyDiagram.Dictionary

    ' Determine Domains.
    Set MyDomain = MyDictionary.Domains
    SmallKeyID = MyDomain.Item("SmallKey").ID
    ShortNameID = MyDomain.Item("ShortName").ID
    MetaIsActiveID = MyDomain.Item("MetaIsActive").ID
    MetaIsDeletedID = MyDomain.Item("MetaIsDeleted").ID
	MetaHashbyteValueID = MyDomain.Item("MetaHashbyteValue").ID
	MetaCreatedDateID = MyDomain.Item("MetaCreatedDate").ID
	MetaModifiedDateID = MyDomain.Item("MetaModifiedDate").ID
	MetaInsertAuditKeyID = MyDomain.Item("MetaInsertAuditKey").ID
	MetaUpdateAuditKeyID = MyDomain.Item("MetaUpdateAuditKey").ID

	' Determine if the current model is logical or physical.
	Logical = MyModel.Logical

	' Iterate through all the selected objects in the current
	' model.
	For Each MySelObject In MySubModel.SelectedObjects

		'Get the object type - we are only concerned
		'with entities.
		ObjType = MySelObject.Type
		
		If ObjType = 1 Then
			' Get the ID for the selected object.
			ID = MySelObject.ID
		
			' Get the actual entity object with this ID. 
			' The model contains the collection of all the
			' entities.
			Set MyEntity = MyModel.Entities.Item(ID)
	
			If Logical = True Then
				' Get the entity name.
				EntityName = MyEntity.EntityName
				
				' Set name of PK and PK attribute.
				AttributeName = Replace(Replace(EntityName,"Dim","",1),"Fact","",1)
				PrimaryKeyName = "PK_" + EntityName
				AlternateKeyName = "AK_" + EntityName
				
				' Add the PK attribute								
                Set MyAttribute = MyEntity.Attributes.Add(AttributeName + "Key", True)
				MyAttribute.DomainId = SmallKeyID

				'Set the alternate key
				Set MyAttribute = MyEntity.Attributes.Add(AttributeName + "Code", False)
				MyAttribute.DomainId = ShortNameID
				MyEntity.Indexes.Add(MyAttribute.AttributeName,AlternateKeyName)

				' Go through the collection of indexes in the entity and name the PK index
				For Each MyIndex In MyEntity.Indexes
					If (MyIndex.IsPK = True) Then
						MyIndex.Name = PrimaryKeyName
						Exit For
					End If
				Next 'MyIndex

				Set MyRelationship = MyModel.Relationships.Add(MyModel.Entities.Item("DimSystem").EntityName,MyEntity.EntityName,RelationshipTypeNonIdentifyingMandatory)
				For Each MyFKColumnPair In MyRelationship.FKColumnPairs
					'Debug.Print MyFKColumnPair.SequenceNo
					If MyFKColumnPair.ChildAttribute.AttributeName = "SystemKey" Then
						SystemKeyID = MyFKColumnPair.ID
						Exit For
					End If
				Next

				MyRelationship.Name = "FK_" + MyRelationship.ChildEntity.EntityName + "_" + MyRelationship.FKColumnPairs.Item(SystemKeyID).ChildAttribute.RoleName
				MyRelationship.BusinessName = MyRelationship.Name

				' Add Meta Attributes
	            Set MyAttribute = MyEntity.Attributes.Add("MetaIsActive", False)
				MyAttribute.DomainId = MetaIsActiveID
				Set MyAttribute = MyEntity.Attributes.Add("MetaIsDeleted", False)
				MyAttribute.DomainId = MetaIsDeletedID
				Set MyAttribute = MyEntity.Attributes.Add("MetaHashbyteValue", False)
				MyAttribute.DomainId = MetaHashbyteValueID
				Set MyAttribute = MyEntity.Attributes.Add("MetaCreatedDate", False)
				MyAttribute.DomainId = MetaCreatedDateID
				Set MyAttribute = MyEntity.Attributes.Add("MetaModifiedDate", False)
				MyAttribute.DomainId = MetaModifiedDateID
				Set MyAttribute = MyEntity.Attributes.Add("MetaInsertAuditKey", False)
				MyAttribute.DomainId = MetaInsertAuditKeyID
				Set MyAttribute = MyEntity.Attributes.Add("MetaUpdateAuditKey", False)
				MyAttribute.DomainId = MetaUpdateAuditKeyID
			End If
			
		End If
		
	Next MySelObject
	
End Sub
