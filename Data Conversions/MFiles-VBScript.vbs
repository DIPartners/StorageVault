Const MFConditionTypeContains = 7 ' Contains the given text string (similar to LIKE 'xyz'). 

Const MFBuiltInObjectTypeAssignment = 10 ' The assignment object type. 
Const MFBuiltInObjectTypeDocument = 0 ' The document object type. 
Const MFBuiltInObjectTypeDocumentCollection = 9 ' The document collection type. 
Const MFBuiltInObjectTypeDocumentOrDocumentCollection = -102 ' A special value that refers to both the document and the document collection types. Can be used in e.g. search conditions. 

Const MFConditionTypeContainsAnyBitwise= 16 ' Contains any of the bits of the given number. Treats its operands as a vector of bits rather than a single number. 
Const MFConditionTypeDoesNotContain = 8 ' Does not contain the given text string (similar to NOT LIKE 'xyz'). 
Const MFConditionTypeDoesNotContainAnyBitwise = 17 ' Does not contain any of the bits of the given number. Treats its operands as a vector of bits rather than a single number. 
Const MFConditionTypeDoesNotMatchWildcardPattern = 12 ' Does not match the given DOS-style wildcard pattern (such as '*.txt'). 
Const MFConditionTypeDoesNotStartWith = 10 ' Does not start with the given text string (similar to NOT LIKE 'xyz%'). 
Const MFConditionTypeEqual = 1 ' Must be equal to. 
Const MFConditionTypeGreaterThan = 3 ' Must be greater than. 
Const MFConditionTypeGreaterThanOrEqual = 5 'Must be greater than or equal to. 
Const MFConditionTypeLessThan = 4 ' Must be less than. 
Const MFConditionTypeLessThanOrEqual = 6 ' Must be less than or equal to. 
Const MFConditionTypeMatchesWildcardPattern = 11 ' Matches the given DOS-style wildcard pattern (such as '*.txt'). 
Const MFConditionTypeNotEqual = 2 ' Must not be equal to. 
Const MFConditionTypeStartsWith = 9 ' Starts with the given text string (similar to LIKE 'xyz%'). 
Const MFConditionTypeStartsWithAtWordBoundary = 15 ' Starts with the given text string at a word boundary. Similar to LIKE 'xyz%', but evaluated at the beginning of each word). 

Const MFStatusTypeCheckedOut = 0 ' The object's checked-out state. 
Const MFStatusTypeCheckedOutAt = 2 ' The date when the object was checked out. 
Const MFStatusTypeCheckedOutTo = 1 ' The user who has checked out the object. 
Const MFStatusTypeDeleted = 5 ' The deletion state of the object. 
Const MFStatusTypeExtID = 8 ' External ID. 
Const MFStatusTypeIsLatestCheckedInVersion = 7 ' The version must be the latest checked-in version. 
Const MFStatusTypeLatestOrSpecific = 9 ' The latest version / specific version indicator. For the latest version, the indicator yields 'Latest' and 'Specific'. For others, the indicator yields only 'Specific'. 
Const MFStatusTypeObjectFlags = 11 ' The object flags of the object. Does not include the deleted flag. 
Const MFStatusTypeObjectID = 3 ' The ID of the object. 
Const MFStatusTypeObjectTypeAndID = 10 ' The type and ID of the object. 
Const MFStatusTypeObjectTypeID = 6 ' The ID of the object type. 
Const MFStatusTypeObjectVersionID = 4 ' The ID of the object version. 
Const MFStatusTypeOriginalObjectID = 14 ' The ID of the object in the vault in which the object was originally created. 
Const MFStatusTypeOriginalObjectIDSegment = 15 ' The ID segment of the object in the vault in which the object was originally created. Segment size is always 1000. 
Const MFStatusTypeOriginalObjectType = 13 ' The type of the object in the vault in which the object was originally created. 
Const MFStatusTypeOriginalVaultGUID = 12 ' The GUID of the vault in which the object was originally created. 

Const MFDatatypeACL = 14 ' The access control list (ACL). 
Const MFDatatypeBoolean = 8 ' Boolean. 
Const MFDatatypeDate = 5 ' Date. 
Const MFDatatypeFILETIME = 12 ' FILETIME (a 64-bit integer). Not used in the properties. 
Const MFDatatypeFloating = 3 ' A double-precision floating point. 
Const MFDatatypeInteger = 2 ' A 32-bit integer. 
Const MFDatatypeInteger64 = 11 ' A 64-bit integer. Not used in the properties. 
Const MFDatatypeLookup = 9 ' Lookup (from a value list). 
Const MFDatatypeMultiLineText = 13 ' Multi-line text. 
Const MFDatatypeMultiSelectLookup = 10 ' Multiple selection from a value list. 
Const MFDatatypeText = 1 ' Text. 
Const MFDatatypeTime = 6 ' Time. 
Const MFDatatypeTimestamp = 7 ' Timestamp. 
Const MFDatatypeUninitialized = 0 ' Unknown type (not yet set to any type). 

Const MFBuiltInPropertyDefAccessedByMe = 81 ' The last time the object was accessed by the current user. 
Const MFBuiltInPropertyDefACLChanged = 90 ' The date and time of the last change to the ACL of the object version. 
Const MFBuiltInPropertyDefAdditionalClasses = 36 ' A list of additional classes for the object. 
Const MFBuiltInPropertyDefAssignedTo = 44 ' A list of users to whom the object version is assigned. 
Const MFBuiltInPropertyDefAssignmentDescription = 41 ' The assignment description for the current object version assignment. 
Const MFBuiltInPropertyDefClass = 100 ' The class of the object. 
Const MFBuiltInPropertyDefClassGroups = 101 ' The class group of the object. 
Const MFBuiltInPropertyDefCollectionMemberCollections = 47 ' A list of document collections belonging to the document collection. 
Const MFBuiltInPropertyDefCollectionMemberDocuments = 46 ' A list of documents belonging to the document collection. 
Const MFBuiltInPropertyDefComment = 33 ' Comment text for the object version. Same as MFBuiltInPropertyDefVersionComment. 
Const MFBuiltInPropertyDefCompleted = 98 ' Specifies whether or not the assignment has been completed. 
Const MFBuiltInPropertyDefCompletedBy = 45 ' A list of users who have completed the current assignment. 
Const MFBuiltInPropertyDefConflictResolved = 96 ' The date and time a conflict was last resolved in favor of the object. 
Const MFBuiltInPropertyDefConstituent = 48 ' Constituent documents for the current object. 
Const MFBuiltInPropertyDefCreated = 20 ' The creation date and time of an object. 
Const MFBuiltInPropertyDefCreatedBy = 25 ' Identifies the user who created the object. 
Const MFBuiltInPropertyDefCreatedFromExternalLocation = 35 ' The external source from which the object was imported. 
Const MFBuiltInPropertyDefDeadline = 42 ' The deadline date for the current object version assignment. 
Const MFBuiltInPropertyDefDeleted = 27 ' The deletion date and time of the object. 
Const MFBuiltInPropertyDefDeletedBy = 28 ' Identifies the user who deleted the object. 
Const MFBuiltInPropertyDefDeletionStatusChanged = 93 ' The date and time of the last change to the deletion status of the object. 
Const MFBuiltInPropertyDefFavoriteView = 82 ' The 'favorite views' where the object should be shown. 
Const MFBuiltInPropertyDefInReplyTo = 84 ' E-mail in-reply-to internet header value. 
Const MFBuiltInPropertyDefInReplyToReference = 85 ' E-mail in-reply-to references between documents belonging to the same conversation. 
Const MFBuiltInPropertyDefIsTemplate = 37 ' A Boolean property identifying whether the object is a template. 
Const MFBuiltInPropertyDefKeywords = 26 ' Keywords for the object. 
Const MFBuiltInPropertyDefLastModified = 21 ' The last modification date and time of an object. 
Const MFBuiltInPropertyDefLastModifiedBy = 23 ' Identifies the user who performed the last modification for the object. 
Const MFBuiltInPropertyDefLocation = 103 ' The location in a repository. 
Const MFBuiltInPropertyDefMarkedForArchiving = 32 ' A Boolean property identifying whether the object is marked for archiving. 
Const MFBuiltInPropertyDefMessageID = 83 ' E-mail message-id from internet headers. 
Const MFBuiltInPropertyDefMonitoredBy = 43 ' A list of users who are monitoring the current assignment. 
Const MFBuiltInPropertyDefNameOrTitle = 0 ' The name or title property definition. 
Const MFBuiltInPropertyDefObjectChanged = 89 ' The date and time of the last change to the object. 
Const MFBuiltInPropertyDefObjectID = -102 ' This special value is used for referring to Object ID in ObjectTypeColumnMapping.TargetPropertyDef. 
Const MFBuiltInPropertyDefOriginalPath = 75 ' The location from which the object was imported to M-Files. 
Const MFBuiltInPropertyDefOriginalPath2 = 77 ' The location from which the object was imported to M-Files (continued). 
Const MFBuiltInPropertyDefOriginalPath3 = 78 ' The location from which the object was imported to M-Files (continued). 
Const MFBuiltInPropertyDefReference = 76 ' A list of referenced documents. 
Const MFBuiltInPropertyDefRejectedBy = 97 ' A list of users who have rejected the current assignment. 
Const MFBuiltInPropertyDefReportPlacement = 88 ' Report placement. 
Const MFBuiltInPropertyDefReportURL = 87 ' Report URL. 
Const MFBuiltInPropertyDefRepository = 102 ' The repository of an object. 
Const MFBuiltInPropertyDefSharedFiles = 95 ' The shared location paths of the object's shared files. 
Const MFBuiltInPropertyDefSignatureManifestation = 86 ' Signature manifestation. 
Const MFBuiltInPropertyDefSingleFileObject = 22 ' A Boolean property identifying whether the object is a single-file object. 
Const MFBuiltInPropertyDefSizeOnServerAllVersions = 31 ' The total size of all object versions. 
Const MFBuiltInPropertyDefSizeOnServerThisVersion = 30 ' The size of this object version. 
Const MFBuiltInPropertyDefState = 39 ' The workflow state of the object. 
Const MFBuiltInPropertyDefStateEntered = 40 ' The date when the object state was changed to the current state. 
Const MFBuiltInPropertyDefStateTransition = 99 ' The workflow state transition of the object. 
Const MFBuiltInPropertyDefStatusChanged = 24 ' The date and time of the last status change of the object. 
Const MFBuiltInPropertyDefTraditionalFolder = 34 ' A traditional folder containing this object version. 
Const MFBuiltInPropertyDefVaultGUID = 94 ' The GUID of the vault that contains or contained the object to which the shortcut object refers to. 
Const MFBuiltInPropertyDefVersionComment = 33 ' Comment text for the object version. Same as MFBuiltInPropertyDefComment. 
Const MFBuiltInPropertyDefVersionCommentChanged = 92 ' The date and time of the last change to the comment of the object version. 
Const MFBuiltInPropertyDefVersionLabel = 29 ' The version label for the object. 
Const MFBuiltInPropertyDefVersionLabelChanged = 91 ' The date and time of the last change to the version label of the object version. 
Const MFBuiltInPropertyDefWorkflow = 38 ' The workflow for the object. 
Const MFBuiltInPropertyDefWorkflowAssignment = 79 ' A property which indicates the assignment related to the workflow of the object. 

Const MFSearchFlagDisableRelevancyRanking = 16 ' Search results are not sorted according to their relevancy. 
Const MFSearchFlagIncludeUnmanagedObjects = 64 ' Include unmanaged objects in the search results. 
Const MFSearchFlagLookAllObjectTypes = 4 ' Perform a separate search for each object type and merges the results. 
Const MFSearchFlagLookInAllVersions = 1 ' Orders to look for matches in all object versions. By default, only the latest object versions are included in the search. 
Const MFSearchFlagNone = 0 ' No flags. 
Const MFSearchFlagResultsFromEachRepository = 128 ' Limits the search results per search location instead of limiting the total number of results from all vaults and repositories. This ensures that the search results include items from all search locations with search hits. 
Const MFSearchFlagReturnLatestVisibleVersion = 2 ' Orders to return the latest version for each object, regardless of whether it is an actual matching version. This flag is effective only if MFSearchFlagLookInAllVersions is on. 

Const MFAuthTypeLoggedOnWindowsUser = 1 ' Windows authentication with current user credentials. 
Const MFAuthTypeSpecificMFilesUser = 3 ' M-Files authentication, user needs to specify the credentials. 
Const MFAuthTypeSpecificWindowsUser = 2 ' Windows authentication, user needs to specify the credentials. 
Const MFAuthTypeUnknown = 0 ' Unspecified authentication method. 
                                                    
Const MFScriptCancel = -1

Dim VaultName, Vault
Dim AuthType, Username, Password, Domain
AuthType=MFAuthTypeLoggedOnWindowsUser

Sub ConnectToVault()
    Dim Vaults
    Dim VaultGUID
    Dim MFAPIApp
    Set MFAPIApp = CreateObject("MFilesAPI.MFilesServerApplication")
    If AuthType = MFAuthTypeLoggedOnWindowsUser Then
        MFAPIApp.Connect AuthType
    ElseIf AuthType = MFAuthTypeSpecificMFilesUser Then
        MFAPIApp.Connect AuthType, UserName, Password
    ElseIf AuthType = MFAuthTypeSpecificWindowsUser Then
        MFAPIApp.Connect AuthType, UserName, Password, Domain
    End If
    Set Vaults = MFAPIApp.GetVaults
    VaultGUID = FindVault(VaultName, Vaults)
    If VaultGUID <> "" Then
        Set Vault = MFAPIApp.LogInToVault(VaultGUID)
    Else
        Set Vault = Nothing
    End If
End Sub

Function FindVault(VaultName, Vaults)
    Dim Vault
    FindVault = ""
    For Each Vault In Vaults
        If Vault.Name = VaultName Then
            FindVault = Vault.GUID
            Exit For
        End If
    Next
End Function

Sub ErrRaise(errtype, msg)
    msgbox msg
end sub

' Gets object properties for a given lookup
'
' EXAMPLE:
'-------------------------------------------------------------------------------------
'		Dim oaLookup: Set oaLookup = CreateObject("MFilesAPI.Lookup")
'		oaLookup.Item = default_val
'		oaLookup.ObjectType = oValListObjType.ID
'		set oRelationVals=GetPropsByLookup(oaLookup)
'-------------------------------------------------------------------------------------
Function GetPropsByLookup(lookup)
    Dim oObjID : Set oObjID = CreateObject( "MFilesAPI.ObjID" )
    Call oObjID.SetIDs( lookup.ObjectType, lookup.Item )
    Dim oObjVer: Set oObjVer = Vault.ObjectOperations.GetLatestObjVer( oObjID, False )
    Set GetPropsByLookup = Vault.ObjectPropertyOperations.GetProperties(oObjVer)
End Function  

' Gets a property definition value
Function GetPropertyValue(propVals, propID)
    GetPropertyValue = propVals.SearchForProperty(propID).TypedValue.DisplayValue
End Function


' Gets the ObjectID by alias
Function GetObjectID( alias )
	Dim iObjectID : iObjectID = Vault.ObjectTypeOperations.GetObjectTypeIDByAlias(alias)
	
	If (iObjectID = -1) Then
	        ErrRaise MFScriptCancel, "Object definition with alias '" & alias & _
        "' was not found, or there are multiple objects with the same alias."
	End IF
	GetObjectID = iObjectID
End Function


' Gets the ValueListID by alias
Function GetVLID( alias )
	Dim iValID : iValID = Vault.ValueListOperations.GetValueListIDByAlias(alias)
	If (iValID = -1) Then
	        ErrRaise MFScriptCancel, "Value List with alias '" & alias & _
        "' was not found, or there are multiple objects with the same alias."
	End IF
	GetVLID = iValID
End Function

' Dumps Data to an Output File
'
' EXAMPLE:
'-------------------------------------------------------------------------------------
'	Dump_to_File "C:\Temp\InvoiceXML.txt", Vault.ObjectPropertyOperations.GetPropertiesAsXML(ObjVer), 0
'
'   This will overwrite the file on each call. 
'   To append to the file, replace the trailing 0 with a 1.
'
'-------------------------------------------------------------------------------------
Sub Dump_to_File( OutputFileName, OutputData, Append)
    Const Write_to_File = 2
    Const Append_to_file = 8
    Dim objFileToWrite
    Dim Write_Type
    if Append <> 0 then
        Write_Type = Append_to_file
    else
        Write_Type = Write_to_File
    end if
	Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(OutputFileName,Write_Type,true)
	objFileToWrite.WriteLine(OutputData)
	objFileToWrite.Close
	Set objFileToWrite = Nothing
end sub

' Gets the ClassID by alias
Function GetClassID( alias )
	Dim iClassID : iClassID = Vault.ClassOperations.GetObjectClassIDByAlias(alias)
	If (iClassID = -1) Then
	        ErrRaise MFScriptCancel, "Class definition with alias '" & alias & _
        "' was not found, or there are multiple classes with the same alias."
	End IF
	GetClassID = iClassID
End Function

' Gets the property definition ID by alias
Function GetPropertyID(alias)
    Dim iID
    iID = Vault.PropertyDefOperations.GetPropertyDefIDByAlias(alias)
    If iID = -1 Then
        ErrRaise MFScriptCancel, "Property definition not found by the specified alias of: " & alias
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

Function FindOT(OTtoFind)
    Dim ObjectTypes, ObjectType
    Set ObjectTypes = Vault.ObjectTypeOperations.GetObjectTypesAdmin()
    Set FindOT = Nothing
    For Each ObjectType In ObjectTypes
        If ObjectType.ObjectType.NameSingular = OTtoFind Then
            Set FindOT = ObjectType
            Exit For
        End If
    Next
End Function

Function FindClass(ClasstoFind)
    Dim ObjectTypes, ObjectType, ObjClass
    Set ObjectTypes = Vault.ClassOperations.GetAllObjectClassesAdmin()
    For Each ObjectType In ObjectTypes
        If ObjectType.Name = ClasstoFind Then
            Set ObjClass = Vault.ClassOperations.GetObjectClass(ObjectType.ID)
            ObjectType.AssociatedPropertyDefs = CreateObject("MFilesAPI.AssociatedPropertyDefs")
            Dim PropDef
            For Each PropDef In ObjClass.AssociatedPropertyDefs
                If PropDef.PropertyDef > 1000 Then
                    ObjectType.AssociatedPropertyDefs.Add -1, PropDef
                End If
            Next
            Set FindClass = ObjectType
            Exit For
        End If
    Next
End Function

Function FindProp(ProptoFind)
    Dim PropertyDefs, PropertyDef
    Set PropertyDefs = Vault.PropertyDefOperations.GetPropertyDefsAdmin()
    Set FindProp = Nothing
    For Each PropertyDef In PropertyDefs
        If PropertyDef.PropertyDef.Name = ProptoFind Then
            Set FindProp = PropertyDef
            Exit For
        End If
    Next
End Function

Function FindVL(VLtoFind)
    Dim ObjectTypes, ObjectType
    Set ObjectTypes = Vault.ValueListOperations.GetValueListsAdmin()
    Set FindVL = Nothing
    For Each ObjectType In ObjectTypes
        If ObjectType.ObjectType.NameSingular = VLtoFind Then
            Set FindVL = ObjectType
            Exit For
        End If
    Next
End Function

Sub CreateProperty(PropertyName, PropertyType, PropertyLookup, PropertyFilter, SimpleConcatenation)
    Dim AdminProp
    Dim FoundAdmin, VL, StatFilt
    Set FoundAdmin = FindProp(PropertyName)
    Dim Alias: Alias = "vProperty." & Replace(Replace(Replace(Replace(PropertyName, "/", ""), " ", ""), ".", ""), "-", "")
    If Not FoundAdmin Is Nothing Then
        Set AdminProp = FoundAdmin
        AdminProp.SemanticAliases.Value = Alias
        AdminProp.PropertyDef.AllowedAsGroupingLevel = False
        AdminProp.PropertyDef.BasedOnValueList = False
        If PropertyLookup <> "" Then
            AdminProp.PropertyDef.BasedOnValueList = True
            Set VL = FindVL(PropertyLookup)
            If VL Is Nothing Then
                AdminProp.PropertyDef.StaticFilter = CreateStaticFilter("Class<=>" & PropertyLookup)
            Else
                AdminProp.PropertyDef.ValueList = VL.ObjectType.ID
            End If
        End If
        If PropertyFilter <> "" Then
            AdminProp.PropertyDef.StaticFilter = CreateStaticFilter(PropertyFilter)
        End If
        AdminProp.PropertyDef.DataType = PropertyType
        AdminProp.PropertyDef.ObjectsSearchableByThisProperty = True
        If SimpleConcatenation <> "" Then
            AdminProp.AutomaticValue.CVSExpression = SimpleConcatenation
            AdminProp.PropertyDef.AutomaticValueType = MFAutomaticValueTypeCalculatedWithPlaceholders
        End If
        Vault.PropertyDefOperations.UpdatePropertyDefAdmin AdminProp, Nothing
    Else
        Set AdminProp = CreateObject("MFilesAPI.PropertyDefAdmin")
        AdminProp.SemanticAliases.Value = Alias
        AdminProp.PropertyDef.Name = PropertyName
        AdminProp.PropertyDef.AllowedAsGroupingLevel = False
        AdminProp.PropertyDef.BasedOnValueList = False
        If PropertyLookup <> "" Then
            AdminProp.PropertyDef.BasedOnValueList = True
            Set VL = FindVL(PropertyLookup)
            If VL Is Nothing Then
                AdminProp.PropertyDef.StaticFilter = CreateStaticFilter("Class<=>" & PropertyLookup)
            Else
                AdminProp.PropertyDef.ValueList = VL.ObjectType.ID
            End If
        End If
        If PropertyFilter <> "" Then
            AdminProp.PropertyDef.StaticFilter = CreateStaticFilter(PropertyFilter)
        End If
        AdminProp.PropertyDef.DataType = PropertyType
        AdminProp.PropertyDef.ObjectsSearchableByThisProperty = True
        If SimpleConcatenation <> "" Then
            AdminProp.AutomaticValue.CVSExpression = SimpleConcatenation
            AdminProp.PropertyDef.AutomaticValueType = MFAutomaticValueTypeCalculatedWithPlaceholders
        End If
        Vault.PropertyDefOperations.AddPropertyDefAdmin AdminProp, Nothing
    End If
End Sub

Function CreateStaticFilter(FilterString)
    Dim SearchCondition: Set SearchCondition = CreateObject("MFilesAPI.SearchCondition")
    Dim SearchConditions: Set SearchConditions = CreateObject("MFilesAPI.SearchConditions")
    Dim Field, Operator, Value
    Dim Conditions: Conditions = Split(FilterString, "~")
    Dim Condition
    For Each Condition In Conditions
        Dim OpStart: OpStart = InStr(Condition, "<")
        Dim OpEnd: OpEnd = InStr(Condition, ">")
        Field = Left(Condition, OpStart - 1)
        Value = Right(Condition, Len(Condition) - OpEnd)
        Operator = Mid(Condition, OpStart + 1, OpEnd - OpStart - 1)
        Set SearchCondition = BuildSearchConditions(Field, Operator, Value)
        If Not (SearchCondition Is Nothing) Then
            SearchConditions.Add -1, SearchCondition
        Else
            MsgBox "Cannot Create Static Filter base on Field [" & Field & "], Operator [" & Operator & "], Value [" & Value & "]"
        End If
    Next
    Set CreateStaticFilter = SearchConditions
End Function

Function BuildSearchConditions(Field, Operator, Value)
    Dim MFConditionType
    Dim SearchCondition: Set SearchCondition = CreateObject("MFilesAPI.SearchCondition")
    Dim Lookup: Set Lookup = CreateObject("MFilesAPI.Lookup")
    Dim Lookups: Set Lookups = CreateObject("MFilesAPI.Lookups")
    MFConditionType = ""
    Set BuildSearchConditions = Nothing
    Dim FieldPropVal: Set FieldPropVal = FindProp(Field)
    If Not (FieldPropVal Is Nothing) Then
        Select Case UCase(Operator)
            Case "="
                MFConditionType = MFConditionTypeEqual
            Case "!="
                MFConditionType = MFConditionTypeNotEqual
            Case "ONE OF"
                MFConditionType = MFConditionTypeEqual
            Case "NOT ONE OF"
                MFConditionType = MFConditionTypeNotEqual
            Case "IS NOT EMPTY"
                MFConditionType = MFConditionTypeNotEqual 'not equal
                'Value should be ==0 @ t
            Case "IS EMPTY"
                MFConditionType = MFConditionTypeEqual
                
        End Select
        Dim Result
        Dim Results(): Results() = Split(Value, ";")
        Dim Count: Count = 0
        'if length(results) = 0, this doesnt execute,  (IS not empty)
            'count stays at 0
        For Each Result In Results
            Dim ResultStr: ResultStr = Result
            If FieldPropVal.PropertyDef.BasedOnValueList Then
                Lookup.Item = SearchValueList(FieldPropVal.PropertyDef.ValueList, ResultStr)
                Lookup.DisplayValue = ResultStr
                Lookups.Add -1, Lookup
            End If
            Count = Count + 1
        Next
        If MFConditionType <> "" Then
            SearchCondition.ConditionType = MFConditionType
            SearchCondition.Expression.DataPropertyValuePropertyDef = FieldPropVal.PropertyDef.ID
            If Count = 1 Then
                SearchCondition.TypedValue.SetValueToLookup Lookup
            ElseIf Count > 1 Then
                SearchCondition.TypedValue.SetValueToMultiSelectLookup Lookups 'same as .SetValue Lookups MFDatatypeLookup
            Else '0
                SearchCondition.TypedValue.SetValueToNULL MFDatatypeLookup
            End If
            Set BuildSearchConditions = SearchCondition
        End If
    End If
End Function
