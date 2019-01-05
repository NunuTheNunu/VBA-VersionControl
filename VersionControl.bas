Attribute VB_Name = "VersionControl"
Public SourceCodeFolder As String

Public Sub ImportCodeModules()
SourceCodeFolder = "M:\SP Projects\Administration\SW Source Code\Estimate Template\Current\"
        on error resume next
Dim TotalModule, BasFiles
If Dir(SourceCodeFolder, vbDirectory) <> vbNullString Then
        TotalModule = ThisWorkbook.VBProject.VBComponents.Count
        Dim ModuleName() As Variant
        ReDim ModuleName(TotalModule)
        With ThisWorkbook.VBProject
            For i% = .VBComponents.Count To 0 Step -1
        
                ModuleName(i%) = .VBComponents(i%).CodeModule.Name
            Next i
            For Each item In ModuleName
                If item <> "VersionControl" Then
                    If Left(item, 2) = "VC" Then
                        .VBComponents.Remove .VBComponents(item)
                        
                    End If
                End If
            Next
            BasFiles = Dir(SourceCodeFolder & "\*.bas")
            Do While BasFiles <> ""
                        .VBComponents.Import SourceCodeFolder & BasFiles
                        BasFiles = Dir
            Loop
            End With

    Else
       Exit Sub
    End If

End Sub
