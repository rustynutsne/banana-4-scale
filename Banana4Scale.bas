' identical to Banana4Scale.swp, just in a github readable format

Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swPart As SldWorks.PartDoc
Dim swAssy As SldWorks.AssemblyDoc
Dim swComp As SldWorks.Component2

Dim longstatus As Long
Dim longwarnings As Long

Sub main()
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    'Set swPart = swModel

    If swModel Is Nothing Then
        MsgBox "No active document open.", vbCritical
        Exit Sub
    End If

    If swModel.GetType <> swDocASSEMBLY Then
        MsgBox "This macro only works in an Assembly document." & vbCrLf & "Please open or create an assembly first.", vbExclamation
        Exit Sub
    End If

    Set swAssy = swModel

    ' *** PLACE THE BANANA FILE IN THE MACRO FOLDER OR CHANGE THIS PATH TO YOUR BANANA PART FILE ***
    Dim bananaPath As String: bananaPath = getBananaPath() ' <--- Edit this!
    ' I will not check if the path exists because the normal file missing dialog is enough

    ' Load the model invisibly in the background
    Dim bananaModel As SldWorks.ModelDoc2
    swApp.DocumentVisible False, swDocumentTypes_e.swDocPART
    Set bananaModel = swApp.OpenDoc6(bananaPath, swDocumentTypes_e.swDocPART, swOpenDocOptions_e.swOpenDocOptions_Silent, "", longstatus, longwarnings)
    swApp.DocumentVisible True, swDocumentTypes_e.swDocPART

    ' Insert the banana component at the assembly origin (0,0,0)
    Set swComp = swAssy.AddComponent5(bananaPath, swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig, "", False, "", 0, 0, 0)

    If swComp Is Nothing Then
        MsgBox "Failed to insert the banana. Check the file path and ensure the part exists.", vbCritical
    Else
        MsgBox "Banana for scale inserted successfully!", vbInformation
    End If

    ' Rebuild and zoom to fit
    swModel.ForceRebuild3 False
    swModel.ViewZoomtofit2

End Sub

Function getBananaPath() As String
    macroPath = swApp.GetCurrentMacroPathName()
    lastslashpos = InStrRev(macroPath, "\")
    If lastslashpos = 0 Then lastslashpos = InStrRev(macroPath, "/")
    folderPath = ""
    If lastslashpos > 0 Then folderPath = Left(macroPath, lastslashpos)
    getBananaPath = folderPath & "Banana.SLDPRT"
End Function
