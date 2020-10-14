Option Explicit

Sub includeFile(fSpec)
    executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(fSpec).readAll()
End Sub

includeFile "MFiles-VBScript.vbs"

VaultName = "SV - Test"
AuthType=MFAuthTypeSpecificMFilesUser
Username="MFAdmin"
Password="K@ml0op$"

ConnectToVault
Msgbox "Start!"

' FixObject "Company", "Companies"
' FixObject "Employee", "Employees"
' FixObject "GL Account", "GL Accounts"
' FixObject "Location", "Locations"
' FixObject "Vendor", "Vendors"
' FixClass "Company"
' FixClass "Employee"
' FixClass "GL Account"
' FixClass "Location"
' FixClass "Vendor"
' FixProp "Company Name",""
' FixProp "Full Name",""
' FixProp "First Name",""
' FixProp "Last Name",""
' FixProp "M-Files user", "M-Files Users"
' FixProp "E-mail","Email Address"
' FixProp "Role","Roles"
' FixProp "Company","Companies"
' FixProp "Account Name",""
' FixProp "Account Number",""
' FixProp "Description","Account Description"
' FixProp "Location Name",""
' FixProp "Address",""
' FixProp "City",""
' FixProp "Province/State",""
' FixProp "Phone","Phone Number"
' FixProp "Regional Manager","Regional Managers"
' FixProp "Sr. Regional Manager","Sr. Regional Managers"
' FixProp "Asset Manager","Asset Managers"
' FixProp "Director","Directors"
' FixProp "VP","VPs"
' FixProp "Office Supplies Approver","Office Supplies Approvers"
' FixProp "Vendor Name",""
' FixProp "Postal Code",""
' FixProp "FAX","Fax Number"
' FixProp "Notes",""
' FixProp "Recurring",""
' FixProp "Default account","Default Account"
' CreateProperty "Regional Manager", MFDatatypeLookup, "Employee", "", ""
' CreateProperty "Sr. Regional Manager", MFDatatypeLookup, "Employee", "", ""
' CreateProperty "Director", MFDatatypeLookup, "Employee", "", ""
' CreateProperty "VP", MFDatatypeLookup, "Employee", "", ""
' CreateProperty "Asset Manager", MFDatatypeLookup, "Employee", "", ""
' CreateProperty "M-Files User", MFDatatypeLookup, "User", "", ""
' CreateProperty "Role", MFDatatypeLookup, "Role", "", ""
' CreateProperty "Company", MFDatatypeLookup, "Company", "", ""

' ReplaceField "Location", Array( _
'     "Regional Managers","Regional Manager", _
'     "Sr. Regional Managers","Sr. Regional Manager", _
'     "Asset Managers","Asset Manager", _
'     "Directors","Director", _
'     "VPs","VP" _
' )

' Dim arr: arr = Array()
' ReDim arr(4, 1)
' arr(0,0) = "Regional Managers"
' arr(0,1) = "Regional Manager"
' arr(1,0) = "Sr. Regional Managers"
' arr(1,1) = "Sr. Regional Manager"
' arr(2,0) = "Asset Managers"
' arr(2,1) = "Asset Manager"
' arr(3,0) = "Directors"
' arr(3,1) = "Director"
' arr(4,0) = "VPs"
' arr(4,1) = "VP"

' ReplaceFieldHKo "Location", arr

' ReplaceField "Employee", Array( _
'     "M-Files Users","M-Files User" _
' )
' CreateProperty "Legal Workflows", MFDatatypeMultiSelectLookup, "Workflow", "", ""
' CreateProperty "Book Keepers", MFDatatypeLookup, "User group", "", ""
' AssociateProperty "Company","Legal Workflows"
' AssociateProperty "Company","Book Keepers"

' CreateProperty "Only Final Approval", MFDatatypeBoolean, "", "", ""
' CreateProperty "Supplies Vendor", MFDatatypeBoolean, "", "", ""
' CreateProperty "Is Vendor", MFDatatypeBoolean, "", "", ""
' AssociateProperty "Location","Only Final Approval"
' AssociateProperty "Vendor","Supplies Vendor"
' AssociateProperty "Vendor","Is Vendor"

Msgbox "Done!"

'1) Add Aliases to Objects
Sub FixObject(Singular, Plural)
    Dim FoundAdmin
    Set FoundAdmin = FindOT(Singular)
    Dim Alias
    Alias = "vObject." & Replace(Replace(Replace(Replace(Singular, "/", ""), " ", ""),".",""),"-","") 
    If Not FoundAdmin Is Nothing Then
        FoundAdmin.ObjectType.NameSingular = Singular
        FoundAdmin.ObjectType.NamePlural = Plural
        FoundAdmin.SemanticAliases.Value = Alias
        Vault.ObjectTypeOperations.UpdateObjectTypeAdmin FoundAdmin
    Else
        Msgbox "Can't find Object " & Singular
        wscript.quit
    End If
End Sub

'1) Add Aliases to Classes
Sub FixClass(ClassName)
    Dim FoundAdmin
    Set FoundAdmin = FindClass(ClassName)
    Dim Alias
    Alias = "vClass." & Replace(Replace(Replace(Replace(ClassName, "/", ""), " ", ""),".",""),"-","") 
    If Not FoundAdmin Is Nothing Then
        FoundAdmin.SemanticAliases.Value = Alias
        Vault.ClassOperations.UpdateObjectClassAdmin FoundAdmin
    Else
        Msgbox "Can't find Class " & ClassName
        wscript.quit
    End If
End Sub

'3) Add Aliases to Properties
Sub FixProp(PropertyName, NewName)
    Dim FoundAdmin
    if NewName="" Then
        NewName=PropertyName
    End IF
    Set FoundAdmin = FindProp(PropertyName)
    Dim Alias
    Alias = "vProperty." & Replace(Replace(Replace(Replace(NewName, "/", ""), " ", ""),".",""),"-","") 
    If Not FoundAdmin Is Nothing Then
        FoundAdmin.SemanticAliases.Value = Alias
        FoundAdmin.PropertyDef.Name = NewName
        Vault.PropertyDefOperations.UpdatePropertyDefAdmin FoundAdmin, Nothing
    Else
        Msgbox "Can't find Property " & PropertyName
        wscript.quit
    End If
End Sub

'6) Add singular fields and remove plural to Locations
'7) Add singular fields and remove plural to Employee
'8) Run VbScript Module to Fix Locations, moving plural property Values to singular Values
Sub ReplaceField(ClassName, FieldList)
    Dim FoundAdmin: Set FoundAdmin = FindClass(ClassName)
    If Not FoundAdmin Is Nothing Then
        Dim FieldName, Alias, OrigID, RepID, PropDef
        Dim Counter: Counter = 1
        For Each FieldName In FieldList
            If Counter = 1 Then
                Alias = "vProperty." & Replace(Replace(Replace(Replace(FieldName, "/", ""), " ", ""), ".", ""), "-", "")
                OrigID = Vault.PropertyDefOperations.GetPropertyDefIDByAlias(Alias)
                Counter = Counter + 1
            ElseIf Counter = 2 Then
                Alias = "vProperty." & Replace(Replace(Replace(Replace(FieldName, "/", ""), " ", ""), ".", ""), "-", "")
                RepID = Vault.PropertyDefOperations.GetPropertyDefIDByAlias(Alias)
                For Each PropDef In FoundAdmin.AssociatedPropertyDefs
                    If PropDef.PropertyDef = OrigID Then
                        PropDef.PropertyDef = RepID
                    End If
                Next
                Counter = 1
            End If
        Next
        Vault.ClassOperations.UpdateObjectClassAdmin FoundAdmin
        Alias = "vClass." & Replace(Replace(Replace(Replace(ClassName, "/", ""), " ", ""), ".", ""), "-", "")
        Dim ObjVers, ObjVer, Vals, CheckedOutObj, FieldVal, SrcAlias
        Set ObjVers = SearchForObjects(Alias, Array( -1 ) )
        if Not (ObjVers is Nothing) Then
            For Each ObjVer In ObjVers
                Set Vals = Vault.ObjectPropertyOperations.GetProperties(ObjVer.ObjVer)
                Set CheckedOutObj = Vault.ObjectOperations.Checkout(ObjVer.ObjVer.ObjID)
                For Each FieldName In FieldList
                    If Counter = 1 Then
                        SrcAlias = "vProperty." & Replace(Replace(Replace(Replace(FieldName, "/", ""), " ", ""), ".", ""), "-", "")
                        Set FieldVal = Vals.SearchForPropertyByAlias(Vault, SrcAlias, True)
                        Counter = Counter + 1
                    ElseIf Counter = 2 Then
                        if Not (FieldVal is Nothing) Then
                            Alias = "vProperty." & Replace(Replace(Replace(Replace(FieldName, "/", ""), " ", ""), ".", ""), "-", "")
                            FieldVal.PropertyDef = Vault.PropertyDefOperations.GetPropertyDefIDByAlias(Alias)
                            FieldVal.TypedValue.SetValue MFDatatypeLookup, FieldVal.TypedValue.GetValueAsLookup()
                            Vault.ObjectPropertyOperations.SetProperty CheckedOutObj.ObjVer, FieldVal
                            Vault.ObjectPropertyOperations.RemoveProperty CheckedOutObj.ObjVer, GetPropertyID(SrcAlias)
                        End If
                        Counter = 1
                    End If
                Next
                Vault.ObjectOperations.Checkin CheckedOutObj.ObjVer
            Next
        End If
    Else
        MsgBox "Can't find Class " & ClassName
        wscript.quit
    End If

End Sub

'Needs create property first then
'8) Add fields to Company Class
'9) Add fields to Location Class
'10) Add fields to Vendor Class
Sub AssociateProperty(ClassName, PropertyName)
    Dim FoundAdmin, FoundProp, PropDef
    Set FoundAdmin = FindClass(ClassName)
    Dim Alias, PropertyID
    Alias = "vProperty." & Replace(Replace(Replace(Replace(PropertyName, "/", ""), " ", ""), ".", ""), "-", "")
    If Not FoundAdmin Is Nothing Then
        PropertyID = Vault.PropertyDefOperations.GetPropertyDefIDByAlias(Alias)
        FoundProp = False
        For Each PropDef In FoundAdmin.AssociatedPropertyDefs
            If PropDef.PropertyDef = PropertyID Then
                FoundProp = True
            End If
        Next
        If Not FoundProp Then
            Set PropDef = CreateObject("MFilesAPI.AssociatedPropertyDef")
            PropDef.Required = False
            PropDef.PropertyDef = PropertyID
            FoundAdmin.AssociatedPropertyDefs.Add -1, PropDef
            Vault.ClassOperations.UpdateObjectClassAdmin FoundAdmin
        End If
    Else
        MsgBox "Can't find Class " & ClassName
    End If
End Sub

Sub ReplaceFieldHKo(ClassName, FieldList)
    Dim FoundAdmin: Set FoundAdmin = FindClass(ClassName)
    If Not FoundAdmin Is Nothing Then
        Dim OrigID, RepID, PropDef, i
        For i = 0 to ubound(FieldList)
            OrigID = Vault.PropertyDefOperations.GetPropertyDefIDByAlias(getAlias("vProperty", FieldList(i, 0)))
            RepID = Vault.PropertyDefOperations.GetPropertyDefIDByAlias(getAlias("vProperty", FieldList(i, 1)))
            For Each PropDef In FoundAdmin.AssociatedPropertyDefs
                If PropDef.PropertyDef = OrigID Then
                    PropDef.PropertyDef = RepID
                End If
            Next            
        Next

        Vault.ClassOperations.UpdateObjectClassAdmin FoundAdmin

        Dim ObjVers, ObjVer, Vals, CheckedOutObj, FieldVal
        Set ObjVers = SearchForObjects(getAlias("vClass", ClassName), Array( -1 ) )
        if Not (ObjVers is Nothing) Then
            For Each ObjVer In ObjVers
                Set Vals = Vault.ObjectPropertyOperations.GetProperties(ObjVer.ObjVer)
                Set CheckedOutObj = Vault.ObjectOperations.Checkout(ObjVer.ObjVer.ObjID)
                For i = 0 to ubound(FieldList)
                    Set FieldVal = Vals.SearchForPropertyByAlias(Vault, getAlias("vProperty", FieldList(i, 0)), True)
                    if Not (FieldVal is Nothing) Then
                        FieldVal.PropertyDef = Vault.PropertyDefOperations.GetPropertyDefIDByAlias(getAlias("vProperty", FieldList(i, 1)))
                        FieldVal.TypedValue.SetValue MFDatatypeLookup, FieldVal.TypedValue.GetValueAsLookup()
                        Vault.ObjectPropertyOperations.SetProperty CheckedOutObj.ObjVer, FieldVal
                        Vault.ObjectPropertyOperations.RemoveProperty CheckedOutObj.ObjVer, GetPropertyID(SrcAlias)                    
                    End if
                Next

                Vault.ObjectOperations.Checkin CheckedOutObj.ObjVer
            Next
        End If
    Else
        MsgBox "Can't find Class " & ClassName
        wscript.quit
    End If

End Sub

Function getAlias(Selector, findName)
    getAlias = Selector & "." & Replace(Replace(Replace(Replace(findName, "/", ""), " ", ""), ".", ""), "-", "")
end Function