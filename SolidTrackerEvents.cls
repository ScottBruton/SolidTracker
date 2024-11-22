Option Explicit

Private WithEvents swApp As SldWorks.SldWorks
Private activeDoc As SldWorks.ModelDoc2
Private jsonFilePath As String
Private rebuildStartTime As Double
Private activeStartTime As Double
Private rebuildActive As Boolean

Private Sub Class_Initialize()
    Main
End Sub

Private Sub Main()
    jsonFilePath = "U:\\SolidTracker\\SolidWorksActivityLog.json"
    
    Set swApp = Application.SldWorks
    Debug.Print "SolidWorks Application Set: " & Not swApp Is Nothing

    rebuildActive = False
    activeStartTime = Timer

    If Dir(jsonFilePath) = "" Then
        CreateEmptyJsonFile
    End If

    On Error Resume Next
    Dim success As Boolean
    
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

' [Rest of your existing functions go here - copy all functions from PartTracker.vba]
' Include: OnActiveDocChange, OnRebuildNotify, OnIdleNotify, LogRebuildData, 
' LogTimeSpent, LogDocumentData, CreateEmptyJsonFile, GetFileName, GetFileType
