' Description: Opens an MPP file read-only in Microsoft Project, renames tasks (task-001...) and resources (resource-001...), deletes resource email addresses, and saves a copy as <source-filename>-sanitised-yyyyMMdd-HHmmss.mpp. The source file is never modified.
' Run: cscript //nologo mpp-sanitise.vbs "C:\path\to\project.mpp"
Option Explicit

Const pjDoNotSave = 0
Const pjViews = 0
Const pjTask = 0
Const pjResource = 1
Const pjProject = 2
Const pjTables = 1
Const pjMPP = 0

Dim mppPath
Dim outputPath
Dim projectApp
Dim projectFile
Dim fileSystem
Dim builtIns

If LCase(Right(WScript.FullName, 11)) <> "cscript.exe" Then
    WScript.Echo "Usage: cscript //nologo mpp-sanitise.vbs ""C:\path\to\project.mpp"""
    WScript.Quit 1
End If

If WScript.Arguments.Count < 1 Then
    WScript.Echo "Usage: cscript //nologo mpp-sanitise.vbs ""C:\path\to\project.mpp"""
    WScript.Quit 1
End If

Set fileSystem = CreateObject("Scripting.FileSystemObject")
mppPath = fileSystem.GetAbsolutePathName(WScript.Arguments(0))

If Not fileSystem.FileExists(mppPath) Then
    WScript.Echo "MPP file was not found: " & mppPath
    WScript.Quit 1
End If

outputPath = fileSystem.BuildPath(fileSystem.GetParentFolderName(mppPath), _
    fileSystem.GetBaseName(mppPath) & "-sanitised-" & Timestamp() & ".mpp")

If LCase(outputPath) = LCase(mppPath) Then
    WScript.Echo "Output path would overwrite the source file: " & mppPath
    WScript.Quit 1
End If

On Error Resume Next
Set projectApp = CreateObject("MSProject.Application")
If Err.Number <> 0 Or projectApp Is Nothing Then
    WScript.Echo "Could not start Microsoft Project. Error " & Err.Number & ": " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0

projectApp.Visible = False
projectApp.DisplayAlerts = False

Info "Reading built-in views and tables"
Set builtIns = GetBuiltIns(projectApp)

Info "Opening read-only: " & mppPath
On Error Resume Next
projectApp.FileOpenEx mppPath, True
If Err.Number <> 0 Then
    WScript.Echo "Could not open MPP file. Error " & Err.Number & ": " & Err.Description
    projectApp.Quit
    Set projectApp = Nothing
    WScript.Quit 1
End If
On Error GoTo 0

Set projectFile = projectApp.ActiveProject

On Error Resume Next
SanitiseTasks projectFile
SanitiseResources projectFile
SanitiseAssignmentNotes projectFile
SanitiseProperties projectFile
SanitiseProjectSummary projectFile
SanitiseViews projectApp, projectFile, builtIns
SanitiseTables projectApp, projectFile, builtIns
SanitiseCustomFields projectApp, projectFile
If Err.Number <> 0 Then
    WScript.Echo "Sanitising failed. Error " & Err.Number & ": " & Err.Description
    CloseProject projectApp
    WScript.Quit 1
End If
On Error GoTo 0

Info "Saving: " & outputPath
On Error Resume Next
If fileSystem.FileExists(outputPath) Then fileSystem.DeleteFile outputPath, True
projectApp.FileSaveAs outputPath, pjMPP
If Err.Number <> 0 Then
    WScript.Echo "Could not save sanitised file. Error " & Err.Number & ": " & Err.Description
    CloseProject projectApp
    WScript.Quit 1
End If
On Error GoTo 0

CloseProject projectApp
Info "Done."

Sub CloseProject(ByRef app)
    On Error Resume Next
    app.FileClose pjDoNotSave
    app.DisplayAlerts = True
    app.Quit
    Set app = Nothing
    Set projectFile = Nothing
End Sub

Sub SanitiseTasks(ByVal project)
    Dim t
    Dim n
    n = 0
    For Each t In project.Tasks
        If Not t Is Nothing Then
            n = n + 1
            t.Name = "task-" & Pad3(n)
            t.Notes = ""
        End If
    Next
    Info "Tasks renamed: " & n
End Sub

Sub SanitiseResources(ByVal project)
    Dim r
    Dim n
    Dim emails
    n = 0
    emails = 0
    For Each r In project.Resources
        If Not r Is Nothing Then
            n = n + 1
            r.Name = "resource-" & Pad3(n)
            r.Notes = ""
            If Len(r.EMailAddress) > 0 Then
                r.EMailAddress = ""
                emails = emails + 1
            End If
        End If
    Next
    Info "Resources renamed: " & n & ", emails deleted: " & emails
End Sub

Sub SanitiseAssignmentNotes(ByVal project)
    Dim a
    Dim n
    n = 0
    For Each a In project.Assignments
        If Not a Is Nothing Then
            a.Notes = ""
            n = n + 1
        End If
    Next
    Info "Assignment notes cleared: " & n
End Sub

Sub SanitiseProperties(ByVal project)
    Dim names
    Dim name
    Dim prop
    Dim cleared
    cleared = 0
    names = Array("Title", "Subject", "Author", "Manager", "Company", "Category", _
        "Keywords", "Comments", "Hyperlink base", "Last author")

    On Error Resume Next
    For Each name In names
        Err.Clear
        project.BuiltinDocumentProperties(name).Value = ""
        If Err.Number = 0 Then
            cleared = cleared + 1
        Else
            Info "Could not clear property '" & name & "': " & Err.Description
        End If
    Next

    Dim i
    For i = project.CustomDocumentProperties.Count To 1 Step -1
        Err.Clear
        project.CustomDocumentProperties(i).Delete
        If Err.Number = 0 Then cleared = cleared + 1
    Next
    On Error GoTo 0
    Info "Project properties cleared: " & cleared
End Sub

' Built-in items are those present in a new blank project (Global template).
Function GetBuiltIns(ByVal app)
    Dim result
    Dim item
    Set result = CreateObject("Scripting.Dictionary")
    result.Add "views", NewNameSet()
    result.Add "taskTables", NewNameSet()
    result.Add "resourceTables", NewNameSet()
    app.FileNew
    For Each item In app.ActiveProject.Views
        result("views")(item.Name) = True
    Next
    For Each item In app.ActiveProject.TaskTables
        result("taskTables")(item.Name) = True
    Next
    For Each item In app.ActiveProject.ResourceTables
        result("resourceTables")(item.Name) = True
    Next
    app.FileClose pjDoNotSave
    Set GetBuiltIns = result
End Function

Function NewNameSet()
    Dim d
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set NewNameSet = d
End Function

Sub SanitiseViews(ByVal app, ByVal project, ByVal builtIn)
    Dim i
    Dim viewName
    Dim deleted
    deleted = 0

    On Error Resume Next
    app.ViewApply "Task Usage"
    Err.Clear
    For i = project.Views.Count To 1 Step -1
        viewName = project.Views(i).Name
        If Not builtIn("views").Exists(viewName) Then
            Err.Clear
            app.OrganizerDeleteItem pjViews, project.FullName, viewName
            If Err.Number = 0 Then
                deleted = deleted + 1
            Else
                Info "Could not delete view '" & viewName & "': " & Err.Description
            End If
        End If
    Next
    On Error GoTo 0
    Info "Custom views deleted: " & deleted
End Sub

Sub SanitiseTables(ByVal app, ByVal project, ByVal builtIn)
    Dim deleted
    deleted = 0
    deleted = deleted + DeleteCustomTables(app, project, project.TaskTables, builtIn("taskTables"), True)
    deleted = deleted + DeleteCustomTables(app, project, project.ResourceTables, builtIn("resourceTables"), False)
    Info "Custom tables deleted: " & deleted
End Sub

Function DeleteCustomTables(ByVal app, ByVal project, ByVal tables, ByVal builtInNames, ByVal isTask)
    Dim i
    Dim tableName
    Dim deleted
    deleted = 0
    On Error Resume Next
    For i = tables.Count To 1 Step -1
        tableName = tables(i).Name
        If Not builtInNames.Exists(tableName) Then
            Err.Clear
            app.OrganizerDeleteItem pjTables, project.FullName, tableName, isTask
            If Err.Number = 0 Then
                deleted = deleted + 1
            Else
                Info "Could not delete table '" & tableName & "': " & Err.Description
            End If
        End If
    Next
    On Error GoTo 0
    DeleteCustomTables = deleted
End Function

Sub SanitiseProjectSummary(ByVal project)
    project.ProjectSummaryTask.Name = "project-001"
    project.ProjectSummaryTask.Notes = ""
    Info "Project summary task renamed: project-001"
End Sub

Function Timestamp()
    Dim d
    d = Now
    Timestamp = Year(d) & Right("0" & Month(d), 2) & Right("0" & Day(d), 2) & "-" & _
        Right("0" & Hour(d), 2) & Right("0" & Minute(d), 2) & Right("0" & Second(d), 2)
End Function

Function FieldConstantFor(ByVal app, ByVal baseFieldName, ByVal fieldType)
    On Error Resume Next
    FieldConstantFor = app.FieldNameToFieldConstant(baseFieldName, fieldType)
    If Err.Number <> 0 Or FieldConstantFor = 0 Then
        FieldConstantFor = Empty
        Err.Clear
    End If
    On Error GoTo 0
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

' Clears values of named custom fields, then deletes the field definitions.
Sub SanitiseCustomFields(ByVal app, ByVal project)
    Dim deleted
    deleted = 0
    deleted = deleted + DeleteCustomFieldsForScope(app, project, pjTask)
    deleted = deleted + DeleteCustomFieldsForScope(app, project, pjResource)
    deleted = deleted + DeleteCustomFieldsForScope(app, project, pjProject)
    Info "Custom fields deleted: " & deleted
End Sub

Function DeleteCustomFieldsForScope(ByVal app, ByVal project, ByVal fieldType)
    Dim customFieldTypes
    Dim customFieldType
    Dim index
    Dim baseFieldName
    Dim fieldId
    Dim fieldName
    Dim item
    Dim deleted
    deleted = 0
    customFieldTypes = Array("Text", "Number", "Date", "Duration", "Cost", "Flag", "Outline Code")

    On Error Resume Next
    For Each customFieldType In customFieldTypes
        For index = 1 To CustomFieldTypeCount(customFieldType)
            baseFieldName = customFieldType & CStr(index)
            fieldId = FieldConstantFor(app, baseFieldName, fieldType)
            If Not IsEmpty(fieldId) Then
                Err.Clear
                fieldName = app.CustomFieldGetName(fieldId)
                If Err.Number = 0 And Len(Trim(CStr(fieldName))) > 0 Then
                    If fieldType = pjTask Then
                        For Each item In project.Tasks
                            If Not item Is Nothing Then item.SetField fieldId, ""
                        Next
                    ElseIf fieldType = pjResource Then
                        For Each item In project.Resources
                            If Not item Is Nothing Then item.SetField fieldId, ""
                        Next
                    End If
                    Err.Clear
                    app.CustomFieldDelete fieldId
                    If Err.Number = 0 Then
                        deleted = deleted + 1
                    Else
                        Info "Could not delete custom field '" & baseFieldName & "': " & Err.Description
                    End If
                End If
            End If
        Next
    Next
    On Error GoTo 0
    DeleteCustomFieldsForScope = deleted
End Function

Function Pad3(ByVal value)
    Pad3 = Right("000" & CStr(value), 3)
    If value > 999 Then Pad3 = CStr(value)
End Function

Sub Info(ByVal message)
    WScript.Echo "info: " & message
End Sub
