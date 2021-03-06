Option Explicit

Sub includeFile(fSpec)
    executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(fSpec).readAll()
End Sub

includeFile "MFiles-VBScript.vbs"

VaultName = "test"
AuthType=MFAuthTypeSpecificMFilesUser
Username="MFAdmin"
Password="K@ml0op$"

ConnectToVault

Dim BaseObject, BaseObjectProperty, BaseObjectPropertyType, BaseObjectPropertyValue
BaseObject="vClass.Employee"
Dim ObjVers : set ObjVers = SearchForObjects(BaseObject, Array( -1 ) )
if (ObjVers is Nothing) Then
    msgbox "Found Nothing"
    wscript.quit
end if
'wscript.quit
Dim ObjVer
Dim MFU
For Each ObjVer In ObjVers
    Dim Vals : Set Vals = Vault.ObjectPropertyOperations.GetProperties(ObjVer.ObjVer)
    Dim CheckedOutObj : set CheckedOutObj = Vault.ObjectOperations.Checkout(ObjVer.ObjVer.ObjID)
    Set MFU = Vals.SearchForPropertyByAlias(Vault, "vProperty.MFilesusers", True)
    if not (MFU is nothing) Then
        MFU.PropertyDef = Vault.PropertyDefOperations.GetPropertyDefIDByAlias("vProperty.MFilesUser")
        MFU.TypedValue.SetValue MFDatatypeLookup, MFU.TypedValue.GetValueAsLookup()
        Vault.ObjectPropertyOperations.SetProperty CheckedOutObj.ObjVer, MFU
        Vault.ObjectPropertyOperations.RemoveProperty CheckedOutObj.ObjVer, GetPropertyID("vProperty.MFilesusers")
    End If
    Vault.ObjectOperations.Checkin CheckedOutObj.ObjVer
Next
msgbox "Done!"

' Gets the ObjectID by alias
Function GetObjectID( alias )
	Dim iObjectID : iObjectID = Vault.ObjectTypeOperations.GetObjectTypeIDByAlias(alias)
	
	If (iObjectID = -1) Then
	        Err.Raise MFScriptCancel, "Object definition with alias '" & alias & _
        "' was not found, or there are multiple objects with the same alias."
	End IF
	GetObjectID = iObjectID
End Function

' Gets the ClassID by alias
Function GetClassID( alias )
	Dim iClassID : iClassID = Vault.ClassOperations.GetObjectClassIDByAlias(alias)
	If (iClassID = -1) Then
		Err.Raise MFScriptCancel, "Class definition with alias '" & alias & _
        "' was not found, or there are multiple classes with the same alias."
	End IF
	GetClassID = iClassID
End Function

' Gets the property definition ID by alias
Function GetPropertyID(alias)
	Dim iID
	iID = Vault.PropertyDefOperations.GetPropertyDefIDByAlias(alias)
	If iID = -1 Then
		Err.Raise MFScriptCancel, "Property definition not found by the specified alias of: " & alias
	End If
	GetPropertyID = iID
End Function

' Gets objects for a given ClassAlias, and an array of multiple Property Alias(es), Property Type(s), Property Condition Type(s), and Property Value(s)
Function SearchForObjects(ClassAlias, Criteria)
' Each element of the Criteria Array
    Dim CriteriaData
' Criteria Data Type
    Dim CriteriaDataType
' Criteria Condition
    Dim CriteriaCondition
' Property ID
    Dim PropertyID
' Search Conditions
    Dim oSearchConditions: Set oSearchConditions = CreateObject("MFilesAPI.SearchConditions")
' Search Condition Object used for Searches
    Dim oSearchCondition: Set oSearchCondition = CreateObject("MFilesAPI.SearchCondition")
' Object Search Results
    Dim oObjSeaResult
' Create "Not Deleted" Condition
    oSearchCondition.ConditionType = MFConditionTypeEqual
    oSearchCondition.Expression.DataStatusValueType = MFStatusTypeDeleted
    oSearchCondition.TypedValue.SetValue MFDatatypeBoolean, False
    oSearchConditions.Add -1, oSearchCondition
' Create Class Condition
    oSearchCondition.ConditionType = MFConditionTypeEqual
    oSearchCondition.Expression.DataPropertyValuePropertyDef = MFBuiltInPropertyDefClass
    oSearchCondition.TypedValue.SetValue MFDatatypeLookup, GetClassID(ClassAlias)
    oSearchConditions.Add -1, oSearchCondition
' Set up Condition for each grouping of Criteria
'
' Criteria are:
'   1: A Property Alias
'   2: A Property Type
'   3: A Property Condition Type
'   4: A Property Value
'
    Dim Counter: Counter = 1
    For Each CriteriaData In Criteria
        If Counter = 1 Then
            If IsNumeric(CriteriaData) Then
                PropertyID = CInt(CriteriaData)
            Else
                PropertyID = GetPropertyID(CriteriaData)
            End If
            If PropertyID = -1 Then
                Exit For
            End If
            oSearchCondition.Expression.DataPropertyValuePropertyDef = PropertyID
            Counter = 2
        ElseIf Counter = 2 Then
            CriteriaDataType = CriteriaData
            Counter = 3
        ElseIf Counter = 3 Then
            CriteriaCondition = CriteriaData
            oSearchCondition.ConditionType = CriteriaCondition
            Counter = 4
        ElseIf Counter = 4 Then
            oSearchCondition.TypedValue.SetValue CriteriaDataType, CriteriaData
            oSearchConditions.Add -1, oSearchCondition
            Counter = 1
        End If
    Next
' Search for Object
    Set oObjSeaResult = Vault.ObjectSearchOperations.SearchForObjectsByConditionsEX(oSearchConditions, MFSearchFlagNone, False,0,0)
    If oObjSeaResult.Count <> 0 Then
' Return the Objects found
        Set SearchForObjects = oObjSeaResult
    Else
' Return the Nothing Object
        Set SearchForObjects = Nothing
    End If
End Function

