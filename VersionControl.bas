Attribute VB_Name = "VersionControl"
Public SourceCodeFolder As String



Sub SaveCodeModules()

Dim ModuleName
'This code Exports VersionControled VBA modules
Dim i%, sName$

With ThisWorkbook.VBProject
    For i% = 1 To .VBComponents.Count
    ModuleName = .VBComponents(i%).CodeModule.Name
        If .VBComponents(i%).CodeModule.CountOfLines > 0 Then
        If Left(ModuleName, 2) = "VC" Then
            sName$ = .VBComponents(i%).CodeModule.Name
            .VBComponents(i%).Export "C:\Users\swei\OneDrive\Documents\Work\Synergy Stuff\Kevin's Estimate\Version Control\Current\" & sName$ & ".bas"
            End If
        End If
    Next i
End With

End Sub

Public Sub ImportCodeModules()
SourceCodeFolder = "M:\SP Projects\Administration\SW Source Code\Estimate Template\Current\"
'On Error EXIT SUB
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
