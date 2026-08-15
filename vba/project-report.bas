Attribute VB_Name = "project_report"
' Description: Prints project properties, document properties, and custom field names for the active Microsoft Project file.
' Run: Import this module into Microsoft Project, open the Immediate Window, then run PrintActiveProjectReportToImmediate.
Option Explicit

Private Const FIELD_TYPE_TASK As Long = 0
Private Const FIELD_TYPE_RESOURCE As Long = 1
Private Const FIELD_TYPE_PROJECT As Long = 2

Public Sub PrintActiveProjectReportToImmediate()
    Dim projectFile As Project

    Set projectFile = ActiveProject

    Info "Running VBA module: project_report"
    Info "MPP file opened:  " & projectFile.FullName
    Info "MPP file size: " & GetFileSize(projectFile.FullName) & " bytes"

    EnumerateProjectProperties projectFile
    EnumerateCustomFields
    Debug.Print "------------------"
End Sub

Private Sub Info(ByVal message As String)
    Debug.Print "info: " & message
End Sub

Private Function GetFileSize(ByVal filePath As String) As String
    On Error Resume Next
    GetFileSize = CStr(FileLen(filePath))
    If Err.Number <> 0 Then
        GetFileSize = "<unavailable: " & Err.Description & ">"
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Sub EnumerateProjectProperties(ByVal projectFile As Project)
    Debug.Print ""
    Debug.Print "Project properties"
    Debug.Print "------------------"

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

Private Sub EnumerateCustomFields()
    Debug.Print ""
    Debug.Print "Custom fields"
    Debug.Print "-------------"

    EnumerateCustomFieldsForScope "Task", FIELD_TYPE_TASK
    EnumerateCustomFieldsForScope "Resource", FIELD_TYPE_RESOURCE
    EnumerateCustomFieldsForScope "Project", FIELD_TYPE_PROJECT
End Sub

Private Sub EnumerateCustomFieldsForScope(ByVal scopeName As String, ByVal fieldType As Long)
    Dim customFieldTypes As Variant
    Dim customFieldType As Variant
    Dim index As Long
    Dim baseFieldName As String
    Dim fieldId As Variant
    Dim customFieldName As String
    Dim valueCount As Long
    Dim foundAny As Boolean

    foundAny = False
    customFieldTypes = Array("Text", "Number", "Date", "Duration", "Cost", "Flag", "Outline Code")

    For Each customFieldType In customFieldTypes
        For index = 1 To CustomFieldTypeCount(CStr(customFieldType))
            baseFieldName = CStr(customFieldType) & CStr(index)
            fieldId = FieldConstantFor(baseFieldName, fieldType)

            If Not IsEmpty(fieldId) Then
                customFieldName = CustomFieldNameFor(CLng(fieldId))
                If Len(customFieldName) > 0 Then
                    valueCount = CountCustomFieldValuesForScope(CLng(fieldId), fieldType, baseFieldName)

                    If Not foundAny Then
                        Debug.Print ""
                        Debug.Print scopeName & " custom fields"
                        Debug.Print String(Len(scopeName & " custom fields"), "-")
                        foundAny = True
                    End If

                    Debug.Print baseFieldName & ": " & customFieldName & " (" & ValueCountLabel(valueCount) & ")"
                End If
            End If
        Next index
    Next customFieldType

    If Not foundAny Then
        Debug.Print ""
        Debug.Print scopeName & " custom fields"
        Debug.Print String(Len(scopeName & " custom fields"), "-")
        Debug.Print "<none>"
    End If
End Sub

Private Function CountCustomFieldValuesForScope(ByVal fieldId As Long, ByVal fieldType As Long, ByVal baseFieldName As String) As Long
    On Error Resume Next
    Select Case fieldType
        Case FIELD_TYPE_TASK
            CountCustomFieldValuesForScope = CountCustomFieldValuesInItems(ActiveProject.Tasks, fieldId, baseFieldName)
        Case FIELD_TYPE_RESOURCE
            CountCustomFieldValuesForScope = CountCustomFieldValuesInItems(ActiveProject.Resources, fieldId, baseFieldName)
        Case FIELD_TYPE_PROJECT
            CountCustomFieldValuesForScope = CountCustomFieldValueOnItem(ActiveProject.ProjectSummaryTask, fieldId, baseFieldName)
        Case Else
            CountCustomFieldValuesForScope = 0
    End Select

    If Err.Number <> 0 Then
        CountCustomFieldValuesForScope = 0
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Function CountCustomFieldValuesInItems(ByVal items As Object, ByVal fieldId As Long, ByVal baseFieldName As String) As Long
    Dim item As Object

    CountCustomFieldValuesInItems = 0
    On Error Resume Next
    For Each item In items
        CountCustomFieldValuesInItems = CountCustomFieldValuesInItems + CountCustomFieldValueOnItem(item, fieldId, baseFieldName)
    Next item

    If Err.Number <> 0 Then
        CountCustomFieldValuesInItems = 0
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Function CountCustomFieldValueOnItem(ByVal item As Object, ByVal fieldId As Long, ByVal baseFieldName As String) As Long
    Dim customFieldValue As String

    CountCustomFieldValueOnItem = 0
    On Error Resume Next
    If item Is Nothing Then
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    customFieldValue = CustomFieldValueFor(item, fieldId, baseFieldName)
    If Len(customFieldValue) > 0 Then
        CountCustomFieldValueOnItem = 1
    End If
End Function

Private Function CustomFieldValueFor(ByVal item As Object, ByVal fieldId As Long, ByVal baseFieldName As String) As String
    Dim rawValue As Variant

    On Error Resume Next
    CustomFieldValueFor = item.GetField(fieldId)
    If Err.Number <> 0 Or IsNull(CustomFieldValueFor) Then
        CustomFieldValueFor = ""
        Err.Clear
    Else
        CustomFieldValueFor = Trim(CStr(CustomFieldValueFor))
    End If
    On Error GoTo 0

    If UCase(CustomFieldValueFor) <> "#ERROR" Then
        Exit Function
    End If

    rawValue = ""
    On Error Resume Next
    rawValue = CallByName(item, Replace(baseFieldName, " ", ""), VbGet)
    If Err.Number <> 0 Or IsNull(rawValue) Then
        CustomFieldValueFor = ""
        Err.Clear
    Else
        CustomFieldValueFor = Trim(CStr(rawValue))
    End If
    On Error GoTo 0
End Function

Private Function ValueCountLabel(ByVal valueCount As Long) As String
    If valueCount = 1 Then
        ValueCountLabel = "1 value"
    Else
        ValueCountLabel = CStr(valueCount) & " values"
    End If
End Function

Private Function CustomFieldTypeCount(ByVal customFieldType As String) As Long
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

Private Function FieldConstantFor(ByVal baseFieldName As String, ByVal fieldType As Long) As Variant
    On Error Resume Next
    FieldConstantFor = Application.FieldNameToFieldConstant(baseFieldName, fieldType)
    If Err.Number <> 0 Or FieldConstantFor = 0 Then
        FieldConstantFor = Empty
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Function CustomFieldNameFor(ByVal fieldId As Long) As String
    On Error Resume Next
    CustomFieldNameFor = Application.CustomFieldGetName(fieldId)
    If Err.Number <> 0 Or IsNull(CustomFieldNameFor) Then
        CustomFieldNameFor = ""
        Err.Clear
    Else
        CustomFieldNameFor = Trim(CStr(CustomFieldNameFor))
    End If
    On Error GoTo 0
End Function

Private Sub EchoNamedProjectProperty(ByVal projectFile As Project, ByVal propertyName As String)
    Dim value As String
    value = GetProjectPropertyValue(projectFile, propertyName)
    Debug.Print propertyName & ": " & value
End Sub

Private Function GetProjectPropertyValue(ByVal projectFile As Project, ByVal propertyName As String) As String
    On Error Resume Next
    Select Case propertyName
        Case "Name"
            GetProjectPropertyValue = CStr(projectFile.Name)
        Case "FullName"
            GetProjectPropertyValue = CStr(projectFile.FullName)
        Case "Path"
            GetProjectPropertyValue = CStr(projectFile.Path)
        Case "Title"
            GetProjectPropertyValue = CStr(projectFile.Title)
        Case "Subject"
            GetProjectPropertyValue = CStr(projectFile.Subject)
        Case "Author"
            GetProjectPropertyValue = CStr(projectFile.Author)
        Case "Manager"
            GetProjectPropertyValue = CStr(projectFile.Manager)
        Case "Company"
            GetProjectPropertyValue = CStr(projectFile.Company)
        Case "Comments"
            GetProjectPropertyValue = CStr(projectFile.Comments)
        Case "CreationDate"
            GetProjectPropertyValue = CStr(projectFile.CreationDate)
        Case "LastSaveDate"
            GetProjectPropertyValue = CStr(projectFile.LastSaveDate)
        Case "Start"
            GetProjectPropertyValue = CStr(projectFile.Start)
        Case "Finish"
            GetProjectPropertyValue = CStr(projectFile.Finish)
        Case "StatusDate"
            GetProjectPropertyValue = CStr(projectFile.StatusDate)
        Case "CurrentDate"
            GetProjectPropertyValue = CStr(projectFile.CurrentDate)
        Case "Calendar"
            GetProjectPropertyValue = CStr(projectFile.Calendar)
        Case "CurrencySymbol"
            GetProjectPropertyValue = CStr(projectFile.CurrencySymbol)
        Case "CurrencyCode"
            GetProjectPropertyValue = CStr(projectFile.CurrencyCode)
        Case Else
            GetProjectPropertyValue = ""
    End Select

    If Err.Number <> 0 Then
        GetProjectPropertyValue = "<unavailable: " & Err.Description & ">"
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Sub EnumerateDocumentProperties(ByVal projectFile As Project, ByVal heading As String, ByVal collectionName As String)
    Dim properties As Object
    Dim propertyItem As Object
    Dim propertyValue As String

    Debug.Print ""
    Debug.Print heading
    Debug.Print String(Len(heading), "-")

    On Error Resume Next
    Select Case collectionName
        Case "BuiltinDocumentProperties"
            Set properties = projectFile.BuiltinDocumentProperties
        Case "CustomDocumentProperties"
            Set properties = projectFile.CustomDocumentProperties
    End Select

    If Err.Number <> 0 Or properties Is Nothing Then
        Debug.Print "<unavailable: " & Err.Description & ">"
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    For Each propertyItem In properties
        propertyValue = GetDocumentPropertyValue(propertyItem)
        Debug.Print propertyItem.Name & ": " & propertyValue
    Next propertyItem
End Sub

Private Function GetDocumentPropertyValue(ByVal propertyItem As Object) As String
    Dim value As Variant

    On Error Resume Next
    value = propertyItem.Value
    If Err.Number <> 0 Then
        GetDocumentPropertyValue = "<unavailable: " & Err.Description & ">"
        Err.Clear
    ElseIf IsEmpty(value) Or IsNull(value) Then
        GetDocumentPropertyValue = ""
    Else
        GetDocumentPropertyValue = CStr(value)
    End If
    On Error GoTo 0
End Function
