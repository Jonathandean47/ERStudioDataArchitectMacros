'#Language "WWB-COM"

Option Explicit

Sub Main
    Dim MyDiagram As Diagram
	Dim MyModel As Model
	Dim MySubModel As SubModel
	Dim MyDictionary As Dictionary
	Dim MyDomain As Domains
	Dim MyEntity As Entity
	Dim MyAttribute As AttributeObj
	Dim LargeKeyDomainId As Integer
	Dim SmallKeyDomainId As Integer
	Dim TargetTables As String

	TargetTables = "Fact"

	' Set base variables
	Set MyDiagram = DiagramManager.ActiveDiagram
	Set MyModel = MyDiagram.ActiveModel
	Set MySubModel = MyModel.ActiveSubModel
	Set MyDictionary = MyDiagram.Dictionary
	Set MyDomain = MyDictionary.Domains

	LargeKeyDomainId = MyDomain.Item("LargeKey").ID
	SmallKeyDomainId = MyDomain.Item("SmallKey").ID

	For Each MyEntity In MyModel.Entities
		If Left(MyEntity.EntityName,Len(TargetTables)) = TargetTables Then
			For Each MyAttribute In MyEntity.Attributes
				If MyAttribute.PrimaryKey Then
					If MyAttribute.DomainId = SmallKeyDomainId Then
						MyAttribute.DomainId = LargeKeyDomainId
						Debug.Print MyEntity.EntityName & "." & MyAttribute.AttributeName
					End If
					Exit For
				End If
			Next
		End If
	Next
End Sub
