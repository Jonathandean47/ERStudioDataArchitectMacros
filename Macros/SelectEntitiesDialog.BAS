'#Language "WWB-COM"

Option Explicit

Public AvailableEntitieNameArr() As String
Public AvailableEntitieIdArr() As Integer
Public SelectedEntitieNameArr() As String
Public SelectedEntitieIdArr() As Integer

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
	Dim OldDomainStr As String
	Dim i As Integer
	Dim j As Integer
	Dim buttonPressed As Integer

	' Set base variables
	Set MyDiagram = DiagramManager.ActiveDiagram
	Set MyModel = MyDiagram.ActiveModel
	Set MySubModel = MyModel.ActiveSubModel
	Set MyDictionary = MyDiagram.Dictionary
	Set MyDomain = MyDictionary.Domains

	'Dialog box for selecting entities
	Begin Dialog UserDialog 630,231,.DlgFunc ' %GRID:10,7,1,1
		Text 20,0,240,14,"Available Entities",.Text1,2
		Text 370,0,240,14,"Selected Entities",.Text2,2
		MultiListBox 20,21,240,203,AvailableEntitieNameArr(),.AvailableEntities,3
		MultiListBox 370,21,240,203,SelectedEntitieNameArr(),.SelectedEntities,3
		PushButton 270,126,90,21,"Add All",.AddAll
		PushButton 270,70,90,21,"<",.RemoveSelection
		PushButton 270,98,90,21,">",.AddSelection
		PushButton 270,42,90,21,"Remove All",.RemoveAll
		OKButton 270,154,90,21
		CancelButton 270,182,90,21
	End Dialog
	Dim dlg As UserDialog

	'Determine which button (ok, cancel, etc.) was pressed
	buttonPressed = Dialog(dlg)
	If buttonPressed = -1 Then 'Ok 
		For i = LBound(SelectedEntitieIdArr) To UBound(SelectedEntitieIdArr)
			Debug.Print SelectedEntitieIdArr(i); "-"; SelectedEntitieNameArr(i); ":"; MyModel.Entities.Item(SelectedEntitieIdArr(i)).EntityName
		Next
	End If

End Sub
Private Function DlgFunc(DlgItem$, Action%, SuppValue?) As Boolean
    Dim MyDiagram As Diagram
	Dim MyModel As Model
	Dim MyEntity As Entity
	ReDim removeArr(-1)
	Dim i As Integer
	Dim availableEntitesInt As Integer

	'set base variables
	Set MyDiagram = DiagramManager.ActiveDiagram
	Set MyModel = MyDiagram.ActiveModel

    Select Case Action%
    Case 1 ' Dialog box initialization
		Beep
		ReDim AvailableEntitieNameArr(-1)
		ReDim AvailableEntitieIdArr(-1)
		ReDim SelectedEntitieNameArr(-1)
		ReDim SelectedEntitieIdArr(-1)

		'initialize all entities into Available MultiSelectBox
		For Each MyEntity In MyModel.Entities
			ReDim Preserve AvailableEntitieNameArr(UBound(AvailableEntitieNameArr)+1)
			ReDim Preserve AvailableEntitieIdArr(UBound(AvailableEntitieIdArr)+1)
			AvailableEntitieIdArr(UBound(AvailableEntitieIdArr)) =  MyEntity.ID
			AvailableEntitieNameArr(UBound(AvailableEntitieNameArr)) =  MyEntity.EntityName
		Next
		DlgListBoxArray "AvailableEntities", AvailableEntitieNameArr
    Case 2 ' Value changing or button pressed
        If DlgItem = "AddSelection" Then
			'Debug.Print "Add:";
			For Each availableEntitesInt In DlgValue "AvailableEntities"
				AddItemId SelectedEntitieIdArr, MyModel.Entities.Item(AvailableEntitieIdArr(availableEntitesInt)).ID
				AddItemName "SelectedEntities", SelectedEntitieNameArr, MyModel.Entities.Item(AvailableEntitieIdArr(availableEntitesInt)).EntityName
				ReDim Preserve removeArr(UBound(removeArr)+1)
				removeArr(UBound(removeArr)) = availableEntitesInt
				'Debug.Print removeArr(UBound(removeArr)); ",";
			Next
			Array_BubbleSort(removeArr)
			'Debug.Print "Remove:";
			For i = UBound(removeArr) To 0 Step -1
				RemoveItemId AvailableEntitieIdArr, removeArr(i)
				RemoveItemName "AvailableEntities", AvailableEntitieNameArr,  removeArr(i)
				'Debug.Print removeArr(i); ",";
			Next i
            DlgFunc = True ' Prevent button press from closing the dialog box
		ElseIf DlgItem = "RemoveSelection" Then
			'Debug.Print "Add:";
			For Each availableEntitesInt In DlgValue "SelectedEntities"
				AddItemId AvailableEntitieIdArr, MyModel.Entities.Item(SelectedEntitieIdArr(availableEntitesInt)).ID
				AddItemName "AvailableEntities", AvailableEntitieNameArr, MyModel.Entities.Item(SelectedEntitieIdArr(availableEntitesInt)).EntityName
				ReDim Preserve removeArr(UBound(removeArr)+1)
				removeArr(UBound(removeArr)) = availableEntitesInt
				Debug.Print removeArr(UBound(removeArr)); ",";
			Next
			Array_BubbleSort(removeArr)
			'Debug.Print "Remove:";
			For i = UBound(removeArr) To 0 Step -1
				RemoveItemId SelectedEntitieIdArr, removeArr(i)
				RemoveItemName "SelectedEntities", SelectedEntitieNameArr, removeArr(i)
				'Debug.Print removeArr(i); ",";
			Next i
		    DlgFunc = True ' Prevent button press from closing the dialog box
		ElseIf DlgItem = "AddAll" Then
			ReDim AvailableEntitieNameArr(-1)
			ReDim AvailableEntitieIdArr(-1)
			ReDim SelectedEntitieNameArr(-1)
			ReDim SelectedEntitieIdArr(-1)

			For Each MyEntity In MyModel.Entities
				ReDim Preserve SelectedEntitieNameArr(UBound(SelectedEntitieNameArr)+1)
				ReDim Preserve SelectedEntitieIdArr(UBound(SelectedEntitieIdArr)+1)
				SelectedEntitieIdArr(UBound(SelectedEntitieIdArr)) =  MyEntity.ID
				SelectedEntitieNameArr(UBound(SelectedEntitieNameArr)) =  MyEntity.EntityName
			Next
			DlgListBoxArray "SelectedEntities", SelectedEntitieNameArr
			DlgListBoxArray "AvailableEntities", AvailableEntitieNameArr

			DlgFunc = True ' Prevent button press from closing the dialog box
		ElseIf DlgItem = "RemoveAll" Then
			ReDim AvailableEntitieNameArr(-1)
			ReDim AvailableEntitieIdArr(-1)
			ReDim SelectedEntitieNameArr(-1)
			ReDim SelectedEntitieIdArr(-1)
			For Each MyEntity In MyModel.Entities
				ReDim Preserve AvailableEntitieNameArr(UBound(AvailableEntitieNameArr)+1)
				ReDim Preserve AvailableEntitieIdArr(UBound(AvailableEntitieIdArr)+1)
				AvailableEntitieIdArr(UBound(AvailableEntitieIdArr)) =  MyEntity.ID
				AvailableEntitieNameArr(UBound(AvailableEntitieNameArr)) =  MyEntity.EntityName
			Next
			DlgListBoxArray "SelectedEntities", SelectedEntitieNameArr
			DlgListBoxArray "AvailableEntities", AvailableEntitieNameArr
			DlgFunc = True ' Prevent button press from closing the dialog box
        End If
		'Debug.Print
    Case 3 ' TextBox or ComboBox text changed
    Case 4 ' Focus changed
		'Debug.Print "DlgFocus("; SuppValue; " > "; DlgFocus(); ")"
    Case 5 ' Idle
        Rem DlgFunc = True ' Continue getting idle actions
    Case 6 ' Function key
    End Select
End Function
Sub AddItemName(List As String, ListArray() As String, Value As String)
    Dim N As Integer
    N = UBound(ListArray)
    ReDim Preserve ListArray(0 To N+1)
    ListArray(N+1) = Value
    DlgListBoxArray List, ListArray()
End Sub
Sub AddItemId(ListArray() As Integer, Value As Integer)
    Dim N As Integer
    N = UBound(ListArray)
    ReDim Preserve ListArray(0 To N+1)
    ListArray(N+1) = Value
End Sub
 
Sub RemoveItemName(List As String, ListArray() As String, Index As Integer)
    Dim N As Integer
    N = UBound(ListArray)
    Dim I As Integer
    For i = Index To N-1
        ListArray(i) = ListArray(i+1)
    Next i
    ReDim Preserve ListArray(0 To N-1)
    DlgListBoxArray List, ListArray()
End Sub
Sub RemoveItemId(ListArray() As Integer, Index As Integer)
    Dim N As Integer
    N = UBound(ListArray)
    Dim I As Integer
    For i = Index To N-1
        ListArray(i) = ListArray(i+1)
    Next i
    ReDim Preserve ListArray(0 To N-1)
End Sub
Sub Array_BubbleSort(ByRef arr() As Variant)

   Dim n As Long, newn As Variant, i As Long, tmp As Variant
   n = UBound(arr)

   Do
      newn = 1
      For i = 0 To n - 1
         If arr(i) > arr(i+1) Then
            'swap
            tmp = arr(i+1)
            arr(i+1) = arr(i)
            arr(i) = tmp
            newn = i+1
         End If
      Next
      n = newn
   Loop While (n > 1)

End Sub
