Attribute VB_Name = "SourceCode"
Public Sub PrepareSourcesForGithub()
    Dim project As InstallerProject
    Set project = StartInstaller()
    
    project.ExportModules CreateModules
    
    project.CloseProject
End Sub


Public Sub UpdateSourcesFromGithub()
    
    Dim project As InstallerProject
    Set project = StartInstaller()
    
    project.InstallModules CreateModules
    
    project.CloseProject
End Sub


Private Function StartInstaller()
    Dim project As New InstallerProject
    project.Path = ThisWorkbook.Path & Application.PathSeparator & ThisWorkbook.name
    
    Set StartInstaller = project
End Function

Private Function CreateModules()
    Dim modules As New Collection
    For Each component In ThisWorkbook.VBProject.VBComponents
        If Not (component.name = "SourceCode" Or Left(component.name, 9) = "Installer") Then
            If component.Type = 1 Then
                modules.Add CreateModule(component.name)
            ElseIf component.Type = 2 Then
                modules.Add CreateClassModule(component.name)
            End If
        End If
    Next
    Set CreateModules = modules
End Function

Private Function CreateModule(name As String)
    Dim module As New InstallerModule
    module.name = name
    module.Path = "Modules/" & name & ".bas"

    Set CreateModule = module
End Function

Private Function CreateClassModule(name As String)
    Dim module As New InstallerModule
    module.name = name
    module.Path = "Class Modules/" & name & ".cls"

    Set CreateClassModule = module
End Function


