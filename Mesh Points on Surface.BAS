Sub Main
    Dim App As femap.model
    Set App = feFemap()

    '==================================
    'This macro defines mesh points on a surface similar to the
    'method prior to Femap 12.
    '
    'Copyright (c) 2019, Aaron Jasso
    '
    'This program is free software: you can redistribute it and/or modify
    'it under the terms of the GNU General Public License as published by
    'the Free Software Foundation, either version 3 of the License, or
    '(at your option) any later version.
    '
    'This program is distributed in the hope that it will be useful,
    'but WITHOUT ANY WARRANTY; without even the implied warranty of
    'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    'GNU General Public License for more details.
    '
    'You should have received a copy of the GNU General Public License
    'along with this program.  If not, see <https://www.gnu.org/licenses/>.
    '
    'Version 1.1
    '
    'Created by AMJ, 4.17.2019
    'Please comment or contribute at https://github.com/aaronjasso/Femap-Old-Style-Mesh-Points-on-Surface
    '==================================

    Dim rc As Variant

    '==================================
    'Prompt user to select a surface and check for existing mesh points
    '==================================
    Dim userSurface As femap.Surface
    Set userSurface = App.feSurface
    Dim meshSet As femap.Set
    Set meshSet = App.feSet
    Dim hasMeshPoints As Boolean
    Dim meshPoint As femap.MeshHardPoint
    Set meshPoint = App.feMeshHardPoint

    hasMeshPoints = False

    rc = userSurface.SelectID("Select Surface to add Mesh Point(s)")
    If rc = 2 Then	'exit if the user presses cancel
    	Exit Sub
    End If

    'Check selected surface for existing mesh points
    rc = meshSet.AddRule(userSurface.ID, FGD_MESHPOINT_BYSURFACE)

	meshSet.Reset
	While meshSet.Next
		meshPoint.Get(meshSet.CurrentID)
		If meshPoint.OnGeometryType = FT_SURFACE _
			And meshPoint.OnGeometryID = userSurface.ID Then
			hasMeshPoints = True
			Exit While
		End If
	Wend

    '==================================
    'Prompt user for action to be taken (Add points, remove points, etc)
    '==================================
	Dim userPoints As femap.Set
	Set userPoints = App.feSet
	Dim userNodes As femap.Set
	Set userNodes = App.feSet
	Dim position As Integer

	Begin Dialog UserDialog 230,203,"Mesh Points",.DialogFunc ' %GRID:10,7,1,1
		GroupBox 10,0,210,161,"Add/Remove Mesh Points",.GroupBox1
		text 20,105,170,14,"No Existing Mesh Points",.Text1
		OptionGroup .Group1
			OptionButton 20,28,180,14,"Use Existing Points",.OptionButton1
			OptionButton 20,49,180,14,"Use Existing Nodes",.OptionButton2
			OptionButton 20,70,180,14,"Create Points",.OptionButton3
			OptionButton 20,105,180,14,"Remove Existing Points",.OptionButton4
		CheckBox 20,126,190,14,"Also remove underlying",.RemGPoints,1
		text 40,140,130,14,"geometric points",.Text2
		OKButton 10,168,100,28
		PushButton 120,168,100,28,"Cancel",.PushButton1
		'CheckBox 500,500,500,500,"HasPoints",.hiddenHasPoints
		CheckBox 0,0,0,0,"HasPoints",.hiddenHasPoints
	End Dialog
	Dim dlg As UserDialog

	'A dlgItem is necessary to pass information into DialogFunc (in this case, whether the surface already has mesh points or not)
	If hasMeshPoints Then
		dlg.hiddenHasPoints = 1
	End If

	'handle cancel button
	If Dialog(dlg) = 1 Then
		Exit Sub
	End If
	
    '==================================
    'Get entities to add MPs, or remove existing MPs as selected by user
    '==================================

	'Get user defined points or nodes
	Dim newPoint As femap.Point
	Set newPoint = App.fePoint
	Dim iterNode As femap.Node
	Set iterNode = App.feNode
	Dim iterPoint As femap.Point
	Set iterPoint = App.fePoint

	Select Case dlg.Group1
	Case 0	'Choose existing points
		userPoints.Select(FT_POINT,True,"Select points to designate mesh points")

	Case 1	'Choose existing nodes
		userNodes.Select(FT_NODE,True,"Select nodes to designate mesh points")
		userNodes.Reset
		While userNodes.Next
			iterNode.Get(userNodes.CurrentID)
			With newPoint
				newPoint.layer = userSurface.layer
				newPoint.xyz =iterNode.xyz
				newPoint.ID = App.Info_NextID(FT_POINT)
			End With
			newPoint.Put(0)
			userPoints.Add(newPoint.ID)
		Wend

	Case 2	'Define points
		Dim coords As Variant

		userPoints.Reset
		While App.feCoordPick("Define coodinates to designate mesh points. Cancel when done.",coords) <> 2
			With newPoint
				.x = coords(0)
				.y = coords(1)
				.z = coords(2)
				.ID = App.Info_NextID(FT_POINT)
				.layer = userSurface.layer
			End With
			rc = newPoint.Put(newPoint.ID)
			userPoints.Add(newPoint.ID)
			App.feViewRegenerate(0)
		Wend

	Case 3	'Remove mesh points from surface
		Dim iterMP As femap.MeshHardPoint
		Set iterMP = App.feMeshHardPoint
		Dim i As Integer

		i = 0

		meshSet.Reset
		While meshSet.Next
			iterMP.Get(meshSet.CurrentID)
			If iterMP.OnGeometryType = FT_SURFACE _
				And iterMP.OnGeometryID = userSurface.ID Then
				iterPoint.Get(iterMP.pointID)
				rc = iterMP.Delete(meshSet.CurrentID)
				If dlg.RemGPoints = 1 Then
					rc = iterPoint.Delete(iterPoint.ID)
				End If
				i = i + 1
			End If
		Wend
		App.feViewRegenerate(0)
		App.feAppUpdatePanes(True)
		App.feAppMessage(FCM_NORMAL,i & " Mesh Points deleted")
		App.feAppMessage(FCM_WARNING,"Note: the Mesh Point Editor dockable pane might still show mesh points; close and reopen it to refresh")
		Exit Sub
	End Select

	'==================================
	'Add mesh points to surface
	'==================================
	Dim newMP As femap.MeshHardPoint
	Set newMP = App.feMeshHardPoint

	userPoints.Reset
	While userPoints.Next
		iterPoint.Get(userPoints.CurrentID)
		With newMP
			.pointID = iterPoint.ID
			.layer = userSurface.layer
			.OnGeometryType = FT_SURFACE
			.OnGeometryID = userSurface.ID
			.NodeOnGeometry = False
			.DistanceToGeometry = 0
			.vLocationOnGeometry = iterPoint.xyz
		End With
		newMP.Put(App.Info_NextID(FT_MESH_POINT))
	Wend

	App.feViewRegenerate(0)
	App.feAppUpdatePanes(True)

End Sub

Function dialogfunc(DlgItem As String, Action As Integer, SuppValue As Integer) As Boolean
    Select Case Action
    Case 1 ' Dialog box initialization
        DlgVisible "hiddenHasPoints",False
        If DlgValue("hiddenHasPoints") = 0 Then
			DlgEnable "OptionButton4", False
			DlgEnable "RemGPoints", False
			DlgEnable "Text2",False
        End If
    Case 5 ' Idle
        Wait .1
    End Select
End Function

