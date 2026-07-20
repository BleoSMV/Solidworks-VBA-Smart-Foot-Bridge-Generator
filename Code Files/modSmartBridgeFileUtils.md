Option Explicit

Public Sub PrepareDesignPaths(ByRef d As BridgeDesign)

    Dim runFolder As String

    If Len(Trim$(d.SaveFolder)) = 0 Then d.SaveFolder = Environ$("USERPROFILE") & "\Documents\SmartBridgeAPI\Output"
    If Len(Trim$(d.ProjectName)) = 0 Then d.ProjectName = "SmartBridge"

    d.ProjectName = CleanFileName(d.ProjectName)
    d.SaveFolder = NormalizeFolderPath(d.SaveFolder)
    d.CreatedOn = Format$(Now, "yyyy-mm-dd hh:nn:ss")

    EnsureFolderExists d.SaveFolder

    runFolder = PathCombine(d.SaveFolder, d.ProjectName & "_" & TimeStampForFolder())
    EnsureFolderExists runFolder

    d.OutputFolder = runFolder
    d.MainBeamPath = PathCombine(runFolder, "SB_001_MainBeam.SLDPRT")
    d.CrossMemberPath = PathCombine(runFolder, "SB_002_CrossMember.SLDPRT")
    d.DeckPanelPath = PathCombine(runFolder, "SB_003_DeckPanel.SLDPRT")
    d.GuardPostPath = PathCombine(runFolder, "SB_004_GuardPost.SLDPRT")
    d.HandrailPath = PathCombine(runFolder, "SB_005_Handrail.SLDPRT")
    d.KickPlatePath = PathCombine(runFolder, "SB_006_KickPlate.SLDPRT")
    d.EndPlatePath = PathCombine(runFolder, "SB_007_EndPlate.SLDPRT")
    d.AssemblyPath = PathCombine(runFolder, d.ProjectName & "_Bridge_Assembly.SLDASM")
    d.drawingPath = PathCombine(runFolder, d.ProjectName & "_Bridge_Drawing.SLDDRW")
    d.DrawingPdfPath = PathCombine(runFolder, d.ProjectName & "_Bridge_Drawing.pdf")
    d.DrawingDwgPath = PathCombine(runFolder, d.ProjectName & "_Bridge_Drawing.dwg")
    d.BomCsvPath = PathCombine(runFolder, d.ProjectName & "_Bridge_BOM.csv")
    d.ReportPath = PathCombine(runFolder, d.ProjectName & "_Bridge_Report.txt")

End Sub

Public Function CleanInputPath(ByVal rawPath As String) As String

    Dim s As String

    s = Trim$(rawPath)
    If Left$(s, 1) = "=" Then s = Trim$(Mid$(s, 2))
    s = Replace(s, Chr$(34), "")
    s = Replace(s, ChrW$(8220), "")
    s = Replace(s, ChrW$(8221), "")
    s = Replace(s, "'", "")
    s = Replace(s, "/", "\")

    CleanInputPath = Trim$(s)

End Function

Public Function NormalizeFolderPath(ByVal rawFolder As String) As String

    Dim s As String
    Dim docsFolder As String

    s = CleanInputPath(rawFolder)
    docsFolder = Environ$("USERPROFILE") & "\Documents"

    If Len(s) = 0 Then
        NormalizeFolderPath = docsFolder & "\SmartBridgeAPI\Output"
        Exit Function
    End If

    If Len(s) >= 2 Then
        If Mid$(s, 2, 1) = ":" Then
            If Len(s) = 2 Then
                s = s & "\"
            ElseIf Mid$(s, 3, 1) <> "\" Then
                s = Left$(s, 2) & "\" & Mid$(s, 3)
            End If
        End If
    End If

    If Not IsAbsoluteFolderPath(s) Then s = docsFolder & "\" & s

    Do While Len(s) > 3 And Right$(s, 1) = "\"
        s = Left$(s, Len(s) - 1)
    Loop

    ValidateNoBadPathCharacters s
    NormalizeFolderPath = s

End Function

Public Function NormalizeFilePath(ByVal rawFilePath As String) As String

    Dim s As String
    Dim docsFolder As String

    s = CleanInputPath(rawFilePath)
    docsFolder = Environ$("USERPROFILE") & "\Documents"

    If Len(s) = 0 Then
        NormalizeFilePath = ""
        Exit Function
    End If

    If Len(s) >= 2 Then
        If Mid$(s, 2, 1) = ":" Then
            If Len(s) = 2 Then
                s = s & "\"
            ElseIf Mid$(s, 3, 1) <> "\" Then
                s = Left$(s, 2) & "\" & Mid$(s, 3)
            End If
        End If
    End If

    If Not IsAbsoluteFolderPath(s) Then s = docsFolder & "\" & s

    ValidateNoBadPathCharacters s
    NormalizeFilePath = s

End Function

Public Function IsAbsoluteFolderPath(ByVal folderPath As String) As Boolean

    Dim s As String

    s = CleanInputPath(folderPath)
    IsAbsoluteFolderPath = False

    If Len(s) >= 3 Then
        If Mid$(s, 2, 1) = ":" And Mid$(s, 3, 1) = "\" Then
            IsAbsoluteFolderPath = True
            Exit Function
        End If
    End If

    If Len(s) >= 2 Then
        If Left$(s, 2) = "\\" Then IsAbsoluteFolderPath = True
    End If

End Function

Public Sub EnsureFolderExists(ByVal folderPath As String)

    Dim fso As Object
    Dim cleanPath As String
    Dim parentPath As String

    cleanPath = NormalizeFolderPath(folderPath)
    Set fso = CreateObject("Scripting.FileSystemObject")

    CheckDriveExists cleanPath

    If fso.FolderExists(cleanPath) Then Exit Sub

    parentPath = fso.GetParentFolderName(cleanPath)

    If Len(parentPath) > 0 Then
        If parentPath <> cleanPath Then
            If Not fso.FolderExists(parentPath) Then EnsureFolderExists parentPath
        End If
    End If

    fso.CreateFolder cleanPath

End Sub

Public Function PathCombine(ByVal folderPath As String, ByVal fileName As String) As String

    Dim cleanFolder As String
    Dim cleanFile As String

    cleanFolder = NormalizeFolderPath(folderPath)
    cleanFile = CleanInputPath(fileName)

    Do While Left$(cleanFile, 1) = "\"
        cleanFile = Mid$(cleanFile, 2)
    Loop

    If Right$(cleanFolder, 1) = "\" Then
        PathCombine = cleanFolder & cleanFile
    Else
        PathCombine = cleanFolder & "\" & cleanFile
    End If

End Function

Public Function CleanFileName(ByVal rawName As String) As String

    Dim badChars As Variant
    Dim i As Long
    Dim s As String

    s = Trim$(rawName)
    badChars = Array("<", ">", ":", """", "/", "\", "|", "?", "*")

    For i = LBound(badChars) To UBound(badChars)
        s = Replace(s, CStr(badChars(i)), "_")
    Next i

    CleanFileName = s

End Function

Public Function FileExistsSB(ByVal filePath As String) As Boolean

    Dim fso As Object
    Dim p As String

    On Error GoTo MissingFile
    p = NormalizeFilePath(filePath)
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExistsSB = fso.FileExists(p)
    Exit Function

MissingFile:
    FileExistsSB = False

End Function

Private Sub CheckDriveExists(ByVal pathText As String)

    Dim fso As Object
    Dim driveLetter As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Len(pathText) >= 2 Then
        If Mid$(pathText, 2, 1) = ":" Then
            driveLetter = Left$(pathText, 1)
            If Not fso.DriveExists(driveLetter) Then
                Err.Raise vbObjectError + 1500, "CheckDriveExists", "The drive '" & driveLetter & ":' is not available."
            End If
        End If
    End If

End Sub

Private Sub ValidateNoBadPathCharacters(ByVal pathText As String)

    Dim s As String

    s = pathText

    If Len(s) >= 2 Then
        If InStr(1, Mid$(s, 3), ":", vbBinaryCompare) > 0 Then
            Err.Raise vbObjectError + 1501, "ValidateNoBadPathCharacters", "Invalid colon found in path:" & vbCrLf & s
        End If
    End If

    If InStr(1, s, "<", vbBinaryCompare) > 0 Then GoTo BadPath
    If InStr(1, s, ">", vbBinaryCompare) > 0 Then GoTo BadPath
    If InStr(1, s, "|", vbBinaryCompare) > 0 Then GoTo BadPath
    If InStr(1, s, "?", vbBinaryCompare) > 0 Then GoTo BadPath
    If InStr(1, s, "*", vbBinaryCompare) > 0 Then GoTo BadPath
    If InStr(1, s, Chr$(34), vbBinaryCompare) > 0 Then GoTo BadPath

    Exit Sub

BadPath:
    Err.Raise vbObjectError + 1502, "ValidateNoBadPathCharacters", "The path contains invalid filename characters:" & vbCrLf & s

End Sub

Public Function TimeStampForFolder() As String
    TimeStampForFolder = Format$(Now, "yyyymmdd_hhnnss")
End Function

Public Function FormatM(ByVal valueM As Double) As String
    FormatM = Format$(valueM, "0.000") & " m"
End Function

Public Function FormatMM(ByVal valueM As Double) As String
    FormatMM = Format$(valueM * 1000#, "0.0") & " mm"
End Function
