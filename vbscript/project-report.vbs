' Description: Opens an MPP file in Microsoft Project, prints project properties, document properties, and custom field names, then closes without saving.
' Run: cscript //nologo project-report.vbs "C:\path\to\project.mpp"
Option Explicit

Const pjDoNotSave = 0
Const pjTask = 0
Const pjResource = 1
Const pjProject = 2

Dim mppPath
Dim projectApp
Dim projectFile

If LCase(Right(WScript.FullName, 11)) <> "cscript.exe" Then
    WScript.Echo "Usage: cscript //nologo project-report.vbs ""C:\path\to\project.mpp"""
    WScript.Quit 1
End If

Info "Running script: " & WScript.ScriptFullName

If WScript.Arguments.Count < 1 Then
    WScript.Echo "Usage: cscript //nologo project-report.vbs ""C:\path\to\project.mpp"""
    WScript.Quit 1
End If

mppPath = WScript.Arguments(0)

If Not FileExists(mppPath) Then
    WScript.Echo "MPP file was not found: " & mppPath
    WScript.Quit 1
End If

On Error Resume Next
Set projectApp = GetObject(, "MSProject.Application")
If Err.Number <> 0 Then
    Err.Clear
    Set projectApp = CreateObject("MSProject.Application")
End If

If Err.Number <> 0 Or projectApp Is Nothing Then
    WScript.Echo "Could not start Microsoft Project. Error " & Err.Number & ": " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0

projectApp.Visible = True

Info "Opening MPP file: " & mppPath
Info "MPP file size: " & GetFileSize(mppPath) & " bytes"

On Error Resume Next
projectApp.FileOpen mppPath
If Err.Number <> 0 Then
    WScript.Echo "Could not open MPP file. Error " & Err.Number & ": " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0

Info "MPP file opened:  " & mppPath

Set projectFile = projectApp.ActiveProject
EnumerateProjectProperties projectFile
EnumerateCustomFields projectApp
WScript.Echo "------------------"
CloseActiveProjectWithoutSaving projectApp

Sub Info(ByVal message)
    WScript.Echo "info: " & message
End Sub

Function FileExists(ByVal filePath)
    Dim fileSystem
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    FileExists = fileSystem.FileExists(filePath)
End Function

Function GetFileSize(ByVal filePath)
    Dim fileSystem
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    GetFileSize = fileSystem.GetFile(filePath).Size
End Function

Sub CloseActiveProjectWithoutSaving(ByVal projectApp)
    On Error Resume Next
    projectApp.FileClose pjDoNotSave
    If Err.Number <> 0 Then
        WScript.Echo "Could not close MPP file without saving. Error " & Err.Number & ": " & Err.Description
        Err.Clear
    Else
        Info "Closed MPP file without saving changes."
    End If
    On Error GoTo 0
End Sub

Sub EnumerateProjectProperties(ByVal projectFile)
    WScript.Echo ""
    WScript.Echo "Project properties"
    WScript.Echo "------------------"

    EchoNamedProjectProperty projectFile, "Name"
    EchoNamedProjectProperty projectFile, "FullName"
    EchoNamedProjectProperty projectFile, "Path"
    EchoNamedProjectProperty projectFile, "Title"
    EchoNamedProjectProperty projectFile, "Subject"
    EchoNamedProjectProperty projectFile, "Author"
    EchoNamedProjectProperty projectFile, "Manager"
    EchoNamedProjectProperty projectFile, "Company"
    EchoNamedProjectProperty projectFile, "Comments"
    EchoNamedProjectProperty projectFile, "CreationDate"
    EchoNamedProjectProperty projectFile, "LastSaveDate"
    EchoNamedProjectProperty projectFile, "Start"
    EchoNamedProjectProperty projectFile, "Finish"
    EchoNamedProjectProperty projectFile, "StatusDate"
    EchoNamedProjectProperty projectFile, "CurrentDate"
    EchoNamedProjectProperty projectFile, "Calendar"
    EchoNamedProjectProperty projectFile, "CurrencySymbol"
    EchoNamedProjectProperty projectFile, "CurrencyCode"

    EnumerateDocumentProperties projectFile, "Built-in document properties", "BuiltinDocumentProperties"
    EnumerateDocumentProperties projectFile, "Custom document properties", "CustomDocumentProperties"
End Sub

Sub EnumerateCustomFields(ByVal projectApp)
    WScript.Echo ""
    WScript.Echo "Custom fields"
    WScript.Echo "-------------"

    EnumerateCustomFieldsForScope projectApp, "Task", pjTask
    EnumerateCustomFieldsForScope projectApp, "Resource", pjResource
    EnumerateCustomFieldsForScope projectApp, "Project", pjProject
End Sub

Sub EnumerateCustomFieldsForScope(ByVal projectApp, ByVal scopeName, ByVal fieldType)
    Dim customFieldTypes
    Dim customFieldType
    Dim index
    Dim baseFieldName
    Dim fieldId
    Dim customFieldName
    Dim valueCount
    Dim foundAny

    foundAny = False
    customFieldTypes = Array("Text", "Number", "Date", "Duration", "Cost", "Flag", "Outline Code")

    For Each customFieldType In customFieldTypes
        For index = 1 To CustomFieldTypeCount(customFieldType)
            baseFieldName = customFieldType & CStr(index)
            fieldId = FieldConstantFor(projectApp, baseFieldName, fieldType)

            If Not IsEmpty(fieldId) Then
                customFieldName = CustomFieldNameFor(projectApp, fieldId)
                If Len(customFieldName) > 0 Then
                    valueCount = CountCustomFieldValuesForScope(projectApp, fieldId, fieldType)

                    If Not foundAny Then
                        WScript.Echo ""
                        WScript.Echo scopeName & " custom fields"
                        WScript.Echo String(Len(scopeName & " custom fields"), "-")
                        foundAny = True
                    End If

                    WScript.Echo baseFieldName & ": " & customFieldName & " (" & ValueCountLabel(valueCount) & ")"
                End If
            End If
        Next
    Next

    If Not foundAny Then
        WScript.Echo ""
        WScript.Echo scopeName & " custom fields"
        WScript.Echo String(Len(scopeName & " custom fields"), "-")
        WScript.Echo "<none>"
    End If
End Sub

Function CountCustomFieldValuesForScope(ByVal projectApp, ByVal fieldId, ByVal fieldType)
    On Error Resume Next
    Select Case fieldType
        Case pjTask
            CountCustomFieldValuesForScope = CountCustomFieldValuesInItems(projectApp.ActiveProject.Tasks, fieldId)
        Case pjResource
            CountCustomFieldValuesForScope = CountCustomFieldValuesInItems(projectApp.ActiveProject.Resources, fieldId)
        Case pjProject
            CountCustomFieldValuesForScope = CountCustomFieldValueOnItem(projectApp.ActiveProject.ProjectSummaryTask, fieldId)
        Case Else
            CountCustomFieldValuesForScope = 0
    End Select

    If Err.Number <> 0 Then
        CountCustomFieldValuesForScope = 0
        Err.Clear
    End If
    On Error GoTo 0
End Function

Function CountCustomFieldValuesInItems(ByVal items, ByVal fieldId)
    Dim item

    CountCustomFieldValuesInItems = 0
    On Error Resume Next
    For Each item In items
        CountCustomFieldValuesInItems = CountCustomFieldValuesInItems + CountCustomFieldValueOnItem(item, fieldId)
    Next

    If Err.Number <> 0 Then
        CountCustomFieldValuesInItems = 0
        Err.Clear
    End If
    On Error GoTo 0
End Function

Function CountCustomFieldValueOnItem(ByVal item, ByVal fieldId)
    Dim customFieldValue

    CountCustomFieldValueOnItem = 0
    On Error Resume Next
    If item Is Nothing Then
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    customFieldValue = CustomFieldValueFor(item, fieldId)
    If Len(customFieldValue) > 0 Then
        CountCustomFieldValueOnItem = 1
    End If
End Function

Function CustomFieldValueFor(ByVal item, ByVal fieldId)
    On Error Resume Next
    CustomFieldValueFor = item.GetField(fieldId)
    If Err.Number <> 0 Or IsNull(CustomFieldValueFor) Then
        CustomFieldValueFor = ""
        Err.Clear
    Else
        CustomFieldValueFor = Trim(CStr(CustomFieldValueFor))
    End If
    On Error GoTo 0
End Function

Function ValueCountLabel(ByVal valueCount)
    If valueCount = 1 Then
        ValueCountLabel = "1 value"
    Else
        ValueCountLabel = CStr(valueCount) & " values"
    End If
End Function

Function CustomFieldTypeCount(ByVal customFieldType)
    Select Case customFieldType
        Case "Text"
            CustomFieldTypeCount = 30
        Case "Number", "Flag"
            CustomFieldTypeCount = 20
        Case "Date", "Duration", "Cost", "Outline Code"
            CustomFieldTypeCount = 10
        Case Else
            CustomFieldTypeCount = 0
    End Select
End Function

Function FieldConstantFor(ByVal projectApp, ByVal baseFieldName, ByVal fieldType)
    On Error Resume Next
    FieldConstantFor = projectApp.FieldNameToFieldConstant(baseFieldName, fieldType)
    If Err.Number <> 0 Or FieldConstantFor = 0 Then
        FieldConstantFor = Empty
        Err.Clear
    End If
    On Error GoTo 0
End Function

Function CustomFieldNameFor(ByVal projectApp, ByVal fieldId)
    On Error Resume Next
    CustomFieldNameFor = projectApp.CustomFieldGetName(fieldId)
    If Err.Number <> 0 Or IsNull(CustomFieldNameFor) Then
        CustomFieldNameFor = ""
        Err.Clear
    Else
        CustomFieldNameFor = Trim(CStr(CustomFieldNameFor))
    End If
    On Error GoTo 0
End Function

Sub EchoNamedProjectProperty(ByVal projectFile, ByVal propertyName)
    Dim value
    value = GetProjectPropertyValue(projectFile, propertyName)
    WScript.Echo propertyName & ": " & value
End Sub

Function GetProjectPropertyValue(ByVal projectFile, ByVal propertyName)
    On Error Resume Next
    Select Case propertyName
        Case "Name"
            GetProjectPropertyValue = projectFile.Name
        Case "FullName"
            GetProjectPropertyValue = projectFile.FullName
        Case "Path"
            GetProjectPropertyValue = projectFile.Path
        Case "Title"
            GetProjectPropertyValue = projectFile.Title
        Case "Subject"
            GetProjectPropertyValue = projectFile.Subject
        Case "Author"
            GetProjectPropertyValue = projectFile.Author
        Case "Manager"
            GetProjectPropertyValue = projectFile.Manager
        Case "Company"
            GetProjectPropertyValue = projectFile.Company
        Case "Comments"
            GetProjectPropertyValue = projectFile.Comments
        Case "CreationDate"
            GetProjectPropertyValue = projectFile.CreationDate
        Case "LastSaveDate"
            GetProjectPropertyValue = projectFile.LastSaveDate
        Case "Start"
            GetProjectPropertyValue = projectFile.Start
        Case "Finish"
            GetProjectPropertyValue = projectFile.Finish
        Case "StatusDate"
            GetProjectPropertyValue = projectFile.StatusDate
        Case "CurrentDate"
            GetProjectPropertyValue = projectFile.CurrentDate
        Case "Calendar"
            GetProjectPropertyValue = projectFile.Calendar
        Case "CurrencySymbol"
            GetProjectPropertyValue = projectFile.CurrencySymbol
        Case "CurrencyCode"
            GetProjectPropertyValue = projectFile.CurrencyCode
        Case Else
            GetProjectPropertyValue = ""
    End Select

    If Err.Number <> 0 Then
        GetProjectPropertyValue = "<unavailable: " & Err.Description & ">"
        Err.Clear
    ElseIf IsEmpty(GetProjectPropertyValue) Or IsNull(GetProjectPropertyValue) Then
        GetProjectPropertyValue = ""
    End If
    On Error GoTo 0
End Function

Sub EnumerateDocumentProperties(ByVal projectFile, ByVal heading, ByVal collectionName)
    Dim properties
    Dim propertyItem
    Dim propertyValue

    WScript.Echo ""
    WScript.Echo heading
    WScript.Echo String(Len(heading), "-")

    On Error Resume Next
    Select Case collectionName
        Case "BuiltinDocumentProperties"
            Set properties = projectFile.BuiltinDocumentProperties
        Case "CustomDocumentProperties"
            Set properties = projectFile.CustomDocumentProperties
    End Select

    If Err.Number <> 0 Or properties Is Nothing Then
        WScript.Echo "<unavailable: " & Err.Description & ">"
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    For Each propertyItem In properties
        propertyValue = GetDocumentPropertyValue(propertyItem)
        WScript.Echo propertyItem.Name & ": " & propertyValue
    Next
End Sub

Function GetDocumentPropertyValue(ByVal propertyItem)
    On Error Resume Next
    GetDocumentPropertyValue = propertyItem.Value
    If Err.Number <> 0 Then
        GetDocumentPropertyValue = "<unavailable: " & Err.Description & ">"
        Err.Clear
    ElseIf IsEmpty(GetDocumentPropertyValue) Or IsNull(GetDocumentPropertyValue) Then
        GetDocumentPropertyValue = ""
    End If
    On Error GoTo 0
End Function