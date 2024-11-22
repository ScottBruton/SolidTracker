Dim swApp As SldWorks.SldWorks
Dim activeDoc As SldWorks.ModelDoc2
Dim jsonFilePath As String
Dim rebuildStartTime As Double
Dim activeStartTime As Double
Dim rebuildActive As Boolean

Sub Main()
    ' JSON file to store the data
    jsonFilePath = "U:\\SolidTracker\\SolidWorksActivityLog.json"
    
    Set swApp = Application.SldWorks
    Debug.Print "SolidWorks Application Set: " & Not swApp Is Nothing

    ' Initialize variables
    rebuildActive = False
    activeStartTime = Timer

    ' Initialize JSON file
    If Dir(jsonFilePath) = "" Then
        CreateEmptyJsonFile
    End If

    ' Start event handlers with error checking
    On Error Resume Next
    Dim success As Boolean
    
    ' Make sure to use the full path to the callback functions
    success = swApp.SetAddInCallbackInfo(0, Me, 1)
    Debug.Print "Callback Info Set: " & success
    
    success = swApp.RegisterNotify(SwConst.swDocumentActiveNotify, "OnActiveDocChange")
    Debug.Print "ActiveDoc Handler Added: " & success
    
    success = swApp.RegisterNotify(SwConst.swRebuildNotify, "OnRebuildNotify")
    Debug.Print "Rebuild Handler Added: " & success
    
    success = swApp.RegisterNotify(SwConst.swIdleNotify, "OnIdleNotify")
    Debug.Print "Idle Handler Added: " & success
    
    On Error GoTo 0
End Sub

' Event: Active Document Changed
Function OnActiveDocChange() As Long
    ' Log the time spent on the current document
    If Not activeDoc Is Nothing Then
        LogTimeSpent(activeDoc, Timer - activeStartTime)
    End If

    ' Track the new active document
    Set activeDoc = swApp.ActiveDoc
    activeStartTime = Timer
    OnActiveDocChange = 0
End Function

' Event: Rebuild Notify (Triggered on both rebuild start and end)
Function OnRebuildNotify(rebuildType As Long) As Long
    If rebuildActive Then
        ' Rebuild has finished; calculate rebuild duration
        Dim rebuildDuration As Double
        rebuildDuration = Timer - rebuildStartTime
        LogRebuildData activeDoc, rebuildDuration
        rebuildActive = False
    Else
        ' Rebuild is starting
        rebuildStartTime = Timer
        rebuildActive = True
    End If

    OnRebuildNotify = 0
End Function

' Event: Idle Notify (Called periodically by SolidWorks)
Function OnIdleNotify() As Long
    ' Update time spent on the active document
    If Not activeDoc Is Nothing Then
        LogTimeSpent(activeDoc, Timer - activeStartTime)
        activeStartTime = Timer
    End If

    OnIdleNotify = 0
End Function

' Function: Log Rebuild Data to JSON File
Sub LogRebuildData(doc As SldWorks.ModelDoc2, rebuildDuration As Double)
    LogDocumentData doc, rebuildDuration, False
End Sub

' Function: Log Time Spent to JSON File
Sub LogTimeSpent(doc As SldWorks.ModelDoc2, duration As Double)
    LogDocumentData doc, 0, True, duration
End Sub

' Function: Log Document Data to JSON File
Sub LogDocumentData(doc As SldWorks.ModelDoc2, rebuildDuration As Double, updateTime As Boolean, Optional timeSpent As Double = 0)
    If doc Is Nothing Then Exit Sub

    Dim fileName As String
    Dim filePath As String
    Dim fileType As String
    Dim dateCreated As String
    Dim featureCount As Long
    Dim planeCount As Long
    Dim sketchCount As Long
    Dim solidBodyCount As Long
    Dim surfaceBodyCount As Long
    Dim componentCount As Long
    Dim fso As Object, fileObj As Object
    Dim jsonData As Object, document As Object, fileContent As String
    Dim found As Boolean
    Dim currentTime As String
    Dim swFeat As SldWorks.Feature
    Dim swBodies As Variant
    Dim swAsm As SldWorks.AssemblyDoc

    ' Extract filename, full path, and file type
    fileName = GetFileName(doc.GetPathName)
    filePath = doc.GetPathName
    fileType = GetFileType(doc.GetType)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fileObj = fso.GetFile(filePath)
    dateCreated = Format(fileObj.DateCreated, "yyyy-mm-ddTHH:MM:SS") ' File creation date
    currentTime = Format(Now, "yyyy-mm-ddTHH:MM:SS")

    ' Count features, planes, and sketches
    featureCount = 0
    planeCount = 0
    sketchCount = 0
    solidBodyCount = 0
    surfaceBodyCount = 0
    componentCount = 0
    Set swFeat = doc.FirstFeature
    While Not swFeat Is Nothing
        featureCount = featureCount + 1
        If swFeat.GetTypeName2 = "ReferencePlane" Then
            planeCount = planeCount + 1
        ElseIf swFeat.GetTypeName2 = "Sketch" Then
            sketchCount = sketchCount + 1
        End If
        Set swFeat = swFeat.GetNextFeature
    Wend

    ' Count solid bodies and surface bodies
    If doc.GetType = swDocPART Then
        swBodies = doc.GetBodies2(swSolidBody, False) ' Get solid bodies
        If Not IsEmpty(swBodies) Then
            solidBodyCount = UBound(swBodies) + 1
        End If
        swBodies = doc.GetBodies2(swSurfaceBody, False) ' Get surface bodies
        If Not IsEmpty(swBodies) Then
            surfaceBodyCount = UBound(swBodies) + 1
        End If
    ElseIf doc.GetType = swDocASSEMBLY Then
        ' For assemblies, count components
        Set swAsm = doc
        componentCount = swAsm.GetComponentCount(True) ' Count all components
    End If

    ' Read the JSON file
    If fso.FileExists(jsonFilePath) Then
        fileContent = fso.OpenTextFile(jsonFilePath, 1).ReadAll
        Set jsonData = JsonConverter.ParseJson(fileContent)
    Else
        Set jsonData = CreateObject("Scripting.Dictionary")
        jsonData("documents") = CreateObject("Scripting.Dictionary")
    End If

    ' Update the data for the document
    found = False
    For Each document In jsonData("documents")
        If document("name") = fileName Then
            document("path") = filePath
            document("fileType") = fileType
            document("lastRebuildTimeSeconds") = rebuildDuration
            document("timeSpentSeconds") = document("timeSpentSeconds") + timeSpent
            document("featureCount") = featureCount
            document("planeCount") = planeCount
            document("sketchCount") = sketchCount
            document("solidBodyCount") = solidBodyCount
            document("surfaceBodyCount") = surfaceBodyCount
            document("componentCount") = componentCount
            If Not updateTime Then document("lastUpdated") = currentTime
            found = True
            Exit For
        End If
    Next

    ' If the document is not found, add a new entry
    If Not found Then
        Set document = CreateObject("Scripting.Dictionary")
        document("name") = fileName
        document("path") = filePath
        document("fileType") = fileType
        document("dateCreated") = dateCreated
        document("lastRebuildTimeSeconds") = rebuildDuration
        document("timeSpentSeconds") = timeSpent
        document("featureCount") = featureCount
        document("planeCount") = planeCount
        document("sketchCount") = sketchCount
        document("solidBodyCount") = solidBodyCount
        document("surfaceBodyCount") = surfaceBodyCount
        document("componentCount") = componentCount
        document("lastUpdated") = currentTime
        jsonData("documents").Add document
    End If

    ' Save the JSON file
    Dim jsonString As String
    jsonString = JsonConverter.ConvertToJson(jsonData, False)
    Dim logFile As Object
    Set logFile = fso.CreateTextFile(jsonFilePath, True)
    logFile.WriteLine jsonString
    logFile.Close
End Sub

' Function: Create an Empty JSON File
Sub CreateEmptyJsonFile()
    Dim fso As Object, jsonFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set jsonFile = fso.CreateTextFile(jsonFilePath, True)
    jsonFile.WriteLine "{""documents"":[]}"
    jsonFile.Close
End Sub

' Helper Function: Extract Filename and Extension from Path
Function GetFileName(fullPath As String) As String
    Dim fileParts() As String
    fileParts = Split(fullPath, "\")
    GetFileName = fileParts(UBound(fileParts)) ' Return the last part (filename with extension)
End Function

' Helper Function: Map File Type Constants to Strings
Function GetFileType(fileTypeConst As Long) As String
    Select Case fileTypeConst
        Case swDocPART
            GetFileType = "Part"
        Case swDocASSEMBLY
            GetFileType = "Assembly"
        Case swDocDRAWING
            GetFileType = "Drawing"
        Case Else
            GetFileType = "Unknown"
    End Select
End Function
