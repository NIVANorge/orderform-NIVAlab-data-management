VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InstallerProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' @class InstallerProject
''
Option Explicit

Private pPath As String
Private pApplication As Object
Private pProject As Object
Private pProjectWasOpen As Boolean
Private pShouldRestore As Boolean

Public ProjectType As ApplicationType
Public HideProgress As Boolean
Public ProgressCallback As String

Public Property Get Path() As String
    Path = pPath
End Property
Public Property Let Path(Value As String)
    pPath = Value
    
    ' Set ProjectType by extension
    If Value <> "" Then
        If VBA.Left$(GetExtension(Value), 3) = "acc" Then
            Me.ProjectType = ApplicationType.AccessApplication
        Else
            Me.ProjectType = ApplicationType.ExcelApplication
        End If
    End If
End Property

''
' Install modules from module path to project
' - Safe: Backup existing module (if found) before install and restore on error
' - Smart: Keep project open if it's already open
'
' @method InstallModules
' @param {Collection of InstallerModule} Modules
''
Public Sub InstallModules(modules As Collection)
    On Error GoTo ErrorHandling

    Dim i As Long
    Dim InstallComplete As Boolean
    Dim module As InstallerModule
    
    If Me.ProjectType = ApplicationType.AccessApplication Then
        If MsgBox("WARNING: The VBA-Web installer for Access is under development and caution should be used for production projects. " & _
            "Be sure to close the database before installation and creating a backup copy before installation is recommended. " & vbNewLine & vbNewLine & _
            "Are you sure you want to install VBA-Web into this Access project?", vbOKCancel, "WARNING") <> vbOK Then
            Exit Sub
        End If
    End If
    
    OpenApplication
    OpenProject
    
    If VbaIsTrusted Then
        If modules.Count > 0 Then
            Progress modules.Count, 0
            i = 0
            
            ' Backup and import modules
            Debug.Print "Installing: " & Me.Path
            For Each module In modules
                Debug.Print "Install: " & module.name
            
                'BackupModule module, i
                ImportModule module
                
                i = i + 1
                Progress modules.Count, i
            Next module
            
            InstallComplete = True
            
            ' Remove backups
            'i = 0
            'For Each module In modules
            '    RemoveBackupModule module, i
            '    i = i + 1
            'Next module
            Debug.Print "Done: " & Me.Path
        End If
    Else
        Err.Raise 10022, "InstallerProject.InstallModules", _
            "In order to install modules, access to the VBA project object model must be trusted." & _
            vbNewLine & vbNewLine & _
            "To enable:" & vbNewLine & _
            "Options > Trust Center > TrustCenterSettings > Macro Settings > " & vbNewLine & _
            "Trust access to the VBA project object model"
    End If

    ClearProgress
    CloseProject
    CloseApplication

    Exit Sub
    
ErrorHandling:
    
    Dim ErrorNumber As Long
    Dim ErrorSource As String
    Dim ErrorDescription As String
    
    ' Store original error
    ErrorNumber = Err.Number
    ErrorSource = Err.Source
    ErrorDescription = Err.Description

    ClearProgress
    Err.Clear
    
    If InstallComplete Then
        ErrorDescription = "Failed to remove backups after install." & _
            vbNewLine & vbNewLine & _
            "Modules were installed successfully, but cleanup after install failed." & vbNewLine & _
            "It is safe to remove the backup files manually (""backup__..."")." & _
            vbNewLine & vbNewLine & _
            "Remove Error:" & vbNewLine & _
            ErrorNumber & ": " & ErrorSource & " - " & ErrorDescription
        ErrorNumber = 10025
        ErrorSource = "InstallerProject.InstallModules"
    ElseIf pShouldRestore Then
        ' Attempt to restore backups if project was open
        ' (Otherwise, they'll be removed by closing without saving changes)
        i = 0
        For Each module In modules
            RestoreModule module, i
            i = i + 1
        Next module
        
        If Err.Number <> 0 Then
            ErrorDescription = "Failed to restore backups after error." & _
                vbNewLine & vbNewLine & _
                "Original Error:" & vbNewLine & _
                ErrorNumber & ": " & ErrorSource & " - " & ErrorDescription & _
                vbNewLine & vbNewLine & _
                "Restore Error:" & vbNewLine & _
                Err.Number & ": " & Err.Source & " - " & Err.Description
            ErrorNumber = 10026
            ErrorSource = "Installer.InstallModules"
        End If
    End If
    
    CloseProject SaveChanges:=False
    CloseApplication
        
    ' Rethrow error
    Err.Raise ErrorNumber, Source:=ErrorSource, Description:=ErrorDescription
End Sub

''
' Export modules from project to module path
'
' @method ExportModules
' @param {Collection of InstallerModule} Modules
''
Public Sub ExportModules(modules As Collection)
    On Error GoTo Cleanup

    Dim i As Long
    Dim module As InstallerModule
    
    OpenApplication
    OpenProject
    
    If VbaIsTrusted Then
        If modules.Count > 0 Then
            Progress modules.Count, 0
            i = 0
            
            For Each module In modules
                ExportModule module
                
                i = i + 1
                Progress modules.Count, i
            Next module
        End If
    Else
        Err.Raise 10022, "InstallerProject.ExportModules", _
            "In order to export modules, access to the VBA project object model must be trusted." & _
            vbNewLine & vbNewLine & _
            "To enable:" & vbNewLine & _
            "Options > Trust Center > TrustCenterSettings > Macro Settings > " & vbNewLine & _
            "Trust access to the VBA project object model"
    End If
    
Cleanup:

    ClearProgress
    CloseProject SaveChanges:=False
    CloseApplication
        
    If Err.Number <> 0 Then
        ' Rethrow error
        Err.Raise Err.Number, Source:=Err.Source, Description:=Err.Description
    End If
End Sub

''
' Open application
'
' @method OpenApplication
''
Public Sub OpenApplication()
    If Me.ProjectType = ApplicationType.AccessApplication Then
        OpenAccess
    Else
        OpenExcel
    End If
End Sub

''
' Close application
'
' @method CloseApplication
''
Public Sub CloseApplication()
    If Me.ProjectType = ApplicationType.AccessApplication Then
        CloseAccess
    Else
        CloseExcel
    End If
End Sub

''
' Open project
'
' @method OpenProject
''
Public Sub OpenProject()
    If Me.ProjectType = ApplicationType.AccessApplication Then
        OpenDatabase
    Else
        OpenWorkbook
    End If
End Sub

''
' Close project (saving changes if specified)
'
' @method CloseProject
' @param {Boolean} [SaveChanges=True]
''
Public Sub CloseProject(Optional SaveChanges As Boolean = True)
    If Me.ProjectType = ApplicationType.AccessApplication Then
        CloseDatabase
    Else
        CloseWorkbook SaveChanges
    End If
End Sub

''
' Import module into project
'
' @method ImportModule
' @param {InstallerModule} Module
''
Public Sub ImportModule(module As InstallerModule)
    If Me.ProjectType = ApplicationType.AccessApplication Then
        ImportDatabaseModule module
    Else
        ImportWorkbookModule module
    End If
End Sub

''
' Export module from project
'
' @method ExportModule
' @param {InstallerModule} Module
''
Public Sub ExportModule(module As InstallerModule)
    If Me.ProjectType = ApplicationType.AccessApplication Then
        ExportDatabaseModule module
    Else
        ExportWorkbookModule module
    End If
End Sub

''
' Backup module (if found) by renaming with backup prefix
'
' Example:
' Module.Name: "Testing", ModuleIndex: 2, BackupPrefix: "backup__"
' Testing -> backup__02Testing
'
' @method BackupModule
' @param {InstallerModule} Module
' @param {Long} ModuleIndex
' @param {String} [BackupPrefix="backup__"]
''
Public Sub BackupModule(module As InstallerModule, ModuleIndex As Long, Optional BackupPrefix As String = "backup__")
    Dim BackupName As String
    BackupName = GetBackupName(module.name, ModuleIndex, BackupPrefix)

    If Me.ProjectType = ApplicationType.AccessApplication Then
        BackupDatabaseModule module, BackupName
    Else
        BackupWorkbookModule module, BackupName
    End If
End Sub

''
' Restore module from backup (if found)
' Finds file with BackupPrefix + ModuleIndex + Module.Name and restore to Module.Name
'
' Example:
' Module.Name: "Testing", ModuleIndex: 2, BackupPrefix: "backup__"
' backup__02Testing -> Testing
'
' @method RestoreModule
' @param {InstallerModule} Module
' @param {Long} ModuleIndex
' @param {String} [BackupPrefix="backup__"]
''
Public Sub RestoreModule(module As InstallerModule, ModuleIndex As Long, Optional BackupPrefix As String = "backup__")
    Dim BackupName As String
    BackupName = GetBackupName(module.name, ModuleIndex, BackupPrefix)

    If Me.ProjectType = ApplicationType.AccessApplication Then
        RestoreDatabaseModule module, BackupName
    Else
        RestoreWorkbookModule module, BackupName
    End If
End Sub

''
' Remove backup module
'
' @method RemoveBackupModule
' @param {InstallerModule} Module
' @param {Long} ModuleIndex
' @param {String} [BackupPrefix="backup__"]
''
Public Sub RemoveBackupModule(module As InstallerModule, ModuleIndex As Long, Optional BackupPrefix As String = "backup__")
    Dim BackupModule As New InstallerModule
    BackupModule.name = GetBackupName(module.name, ModuleIndex, BackupPrefix)

    If Me.ProjectType = ApplicationType.AccessApplication Then
        RemoveDatabaseModule BackupModule
    Else
        RemoveWorkbookModule BackupModule
    End If
End Sub

''
' Check if VBA is trusted
'
' @param {Workbook} Project
' @return {Boolean}
''
Public Function VbaIsTrusted() As Boolean
    On Error Resume Next

    Dim ComponentCount As Long
    If Me.ProjectType = ApplicationType.AccessApplication Then
        ComponentCount = pApplication.VBE.ActiveVBProject.VBComponents.Count
    Else
        If pProject.VBProject.Protection = 1 Then ' 1 = Locked
            On Error GoTo 0
            Err.Raise 10026, "InstallerProject.VbaIsTrusted", _
                "VBA Project is locked, to install in this Workbook, open Workbook, unlock VBA, and leave the Workbook open while running the installer again"
        End If
        
        ComponentCount = pProject.VBProject.VBComponents.Count
    End If
    
    If Err.Number <> 0 Then
        Err.Clear
        VbaIsTrusted = False
    Else
        VbaIsTrusted = True
    End If
End Function

' ============================================= '
' Private Methods
' ============================================= '

Private Function GetBackupName(ModuleName As String, ModuleIndex As Long, BackupPrefix As String) As String
    ' Truncate backup name to keep below 31 characters and include index to avoid name collisions
    GetBackupName = BackupPrefix & VBA.Right$("00" & ModuleIndex, 3)
    GetBackupName = GetBackupName & VBA.Left$(ModuleName, 31 - VBA.Len(GetBackupName))
End Function

Private Sub Progress(Optional ModuleCount As Long, Optional CompleteCount As Long)
    On Error Resume Next

    If HideProgress Then
        Exit Sub
    ElseIf ProgressCallback <> "" Then
        ' Callback
        Application.Run ProgressCallback, ModuleCount, CompleteCount
    Else
        Dim ProgressMessage As String
        ProgressMessage = "Installing Modules [" & _
            VBA.Left$(VBA.String$(CompleteCount, "=") & VBA.Space$(ModuleCount), ModuleCount) & _
            "]"
        
        If ModuleCount = CompleteCount Then
            ProgressMessage = ProgressMessage & " Done!"
        End If
        
        Application.StatusBar = ProgressMessage
    End If
End Sub

Private Sub ClearProgress()
    Application.StatusBar = ""
End Sub

' ============================================= '
' Excel
' ============================================= '

Private Sub OpenExcel()
    Set pApplication = Application
End Sub

Private Sub CloseExcel()
    ' Do nothing, Excel was already running
End Sub

Private Sub OpenWorkbook()
    Dim ScreenWasUpdating As Boolean
    
    On Error Resume Next
    Set pProject = pApplication.Workbooks(GetFilename(Me.Path))
    On Error GoTo 0
    
    If pProject Is Nothing Or Err.Number <> 0 Then
        Err.Clear
        
        If FileExists(Me.Path) Then
            ScreenWasUpdating = pApplication.ScreenUpdating
            pApplication.ScreenUpdating = False
            
            Set pProject = pApplication.Workbooks.Open(Me.Path)
            
            pApplication.ScreenUpdating = ScreenWasUpdating
        Else
            Err.Raise 10021, "InstallerProject.OpenWorkbook", "Project not found at path"
        End If
    Else
        pProjectWasOpen = True
        pShouldRestore = True
    End If
End Sub

Private Sub CloseWorkbook(Optional SaveChanges As Boolean = True)
    If Not pProject Is Nothing And Not pProjectWasOpen Then
        pProject.Close SaveChanges
    End If
    
    Set pProject = Nothing
End Sub

Private Function GetWorkbookModule(module As InstallerModule, Optional FromWorkbook As Object) As Object
    If FromWorkbook Is Nothing Then
        Set FromWorkbook = pProject
    End If

    On Error Resume Next
    Set GetWorkbookModule = FromWorkbook.VBProject.VBComponents(module.name)
    On Error GoTo 0
End Function

Private Sub ImportWorkbookModule(module As InstallerModule)
    Dim Path As String
    Dim ExistingModule As Object
    Dim TempPath As String
    
    On Error GoTo Cleanup
    
    ' For local file, export to temporary path
    If module.FromLocal Then
        TempPath = FullPath(module.Path)
        ExportWorkbookModule module, FromWorkbook:=ThisWorkbook
    End If
    
    Path = FullPath(module.Path)
    If Not FileExists(Path) Then
        Err.Raise 10025, "InstallerProject.ImportWorkbookModule", "No module found at " & Path
    End If
    
    ' Check for existing module and remove
    Set ExistingModule = GetWorkbookModule(module)
    If Not ExistingModule Is Nothing Then
        RemoveWorkbookModule module
    End If
    
    ' Import module
    pProject.VBProject.VBComponents.Import Path
    
Cleanup:
    
    Dim ErrorNumber As Long
    Dim ErrorSource As String
    Dim ErrorDescription As String
    
    ' Store original error
    If Err.Number <> 0 Then
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorDescription = Err.Description
    End If
    
    If TempPath <> "" Then
        ' Remove temporary file
        DeleteFile TempPath
    End If
    
    If ErrorNumber <> 0 Then
        Err.Raise ErrorNumber, ErrorSource, ErrorDescription
    End If
End Sub

Private Sub ExportWorkbookModule(module As InstallerModule, Optional FromWorkbook As Object)
    Dim Path As String
    Dim ExistingModule As Object
    
    If FromWorkbook Is Nothing Then
        Set FromWorkbook = pProject
    End If
    Path = FullPath(module.Path)
    Set ExistingModule = GetWorkbookModule(module, FromWorkbook)
    
    If Not ExistingModule Is Nothing Then
        ' Export fails for existing file, so remove if exists
        If FileExists(Path) Then
            DeleteFile Path
        End If
    
        ExistingModule.Export Path
    End If
End Sub

Private Sub RemoveWorkbookModule(module As InstallerModule)
    Dim ExistingModule As Object
    Set ExistingModule = GetWorkbookModule(module)
    If Not ExistingModule Is Nothing Then
        pProject.VBProject.VBComponents.Remove ExistingModule
    End If
End Sub

Private Sub BackupWorkbookModule(module As InstallerModule, BackupName As String)
    Dim Backup As Object
    Dim Existing As Object
    
    Set Backup = GetWorkbookModule(module)
    If Not Backup Is Nothing Then
        ' Remove previous backup (if found)
        Dim BackupModule As New InstallerModule
        BackupModule.name = BackupName
        Set Existing = GetWorkbookModule(BackupModule)
        If Not Existing Is Nothing Then
            RemoveWorkbookModule BackupModule
        End If
            
        ' Update name of module to backup
        Backup.name = BackupName
    End If
End Sub

Private Sub RestoreWorkbookModule(module As InstallerModule, BackupName As String)
    Dim BackupModule As New InstallerModule
    Dim Backup As Object
    Dim Existing As Object
    
    BackupModule.name = BackupName
    Set Backup = GetWorkbookModule(BackupModule)
    If Not Backup Is Nothing Then
        ' Remove restored module (if found)
        Set Existing = GetWorkbookModule(module)
        If Not Existing Is Nothing Then
            RemoveWorkbookModule module
        End If
        
        ' Restore backup
        Backup.name = module.name
    End If
End Sub

' ============================================= '
' Access
' ============================================= '

Public Sub OpenAccess()
    Set pApplication = CreateObject("Access.Application")
    pApplication.Visible = True
End Sub

Public Sub CloseAccess()
    If Not pApplication Is Nothing Then
        ' Access should have been quit with/without saving before here
        ' So if still around, something bad happened, quit without saving changes
        pApplication.Quit 2 ' 2 = acQuitSaveNone
    End If
    Set pApplication = Nothing
End Sub

Public Sub OpenDatabase()
    ' First, check for already open database
    ' TODO Find alternative approach that store pProjectWasOpen
    '      or don't close Access projects
'    On Error Resume Next
'
'    Dim OpenDb As Object
'    Set OpenDb = GetObject(Me.Path, "Access.Application")
'
'    On Error GoTo 0
'
'    If Not OpenDb Is Nothing Then
'        pApplication.Quit
'        Set pApplication = OpenDb.Application
'    Else
        pApplication.OpenCurrentDatabase Me.Path, Exclusive:=True
'    End If
End Sub

Public Sub CloseDatabase(Optional SaveChanges As Boolean = True)
    If SaveChanges Then
        pApplication.Quit 1 ' 1 = acQuitSaveAll
    Else
        pApplication.Quit 2 ' 2 = acQuitSaveNone
    End If
    Set pApplication = Nothing
End Sub

Public Function GetDatabaseModule(module As InstallerModule) As Object
    On Error Resume Next
    Set GetDatabaseModule = pApplication.VBE.ActiveVBProject.VBComponents(module.name)
    On Error GoTo 0
End Function

Public Sub ImportDatabaseModule(module As InstallerModule)
    Dim Path As String
    Dim ExistingModule As Object
    Dim TempPath As String

    On Error GoTo Cleanup

    ' For local file, export to temporary path
    If module.FromLocal Then
        TempPath = FullPath(module.Path)
        ExportWorkbookModule module, FromWorkbook:=ThisWorkbook
    End If

    Path = FullPath(module.Path)
    If Not FileExists(Path) Then
        Err.Raise 10025, "InstallerProject.ImportDatabaseModule", "No module found at " & Path
    End If

    ' Check for existing module and remove
    Set ExistingModule = GetDatabaseModule(module)
    If Not ExistingModule Is Nothing Then
        RemoveDatabaseModule module
    End If

    ' Import module
    pApplication.VBE.ActiveVBProject.VBComponents.Import Path
    pApplication.DoCmd.Save 5, module.name ' 5 = acModule

Cleanup:

    Dim ErrorNumber As Long
    Dim ErrorSource As String
    Dim ErrorDescription As String

    ' Store original error
    If Err.Number <> 0 Then
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorDescription = Err.Description
    End If

    If TempPath <> "" Then
        ' Remove temporary file
        DeleteFile TempPath
    End If

    If ErrorNumber <> 0 Then
        Err.Raise ErrorNumber, ErrorSource, ErrorDescription
    End If
End Sub

Public Sub ExportDatabaseModule(module As InstallerModule)
    Dim Path As String
    Dim ExistingModule As Object

    Path = FullPath(module.Path)
    Set ExistingModule = GetDatabaseModule(module)

    If Not ExistingModule Is Nothing Then
        ' Export fails for existing file, so remove if exists
        If FileExists(Path) Then
            DeleteFile Path
        End If

        ExistingModule.Export Path
    End If
End Sub

Public Sub RemoveDatabaseModule(module As InstallerModule)
    Dim ExistingModule As Object

    Set ExistingModule = GetDatabaseModule(module)
    If Not ExistingModule Is Nothing Then
        pApplication.VBE.ActiveVBProject.VBComponents.Remove ExistingModule
    End If
End Sub

Public Sub BackupDatabaseModule(module As InstallerModule, BackupName As String)
    Dim Backup As Object
    Dim Existing As Object

    Set Backup = GetDatabaseModule(module)
    If Not Backup Is Nothing Then
        ' Remove previous backup (if found)
        Dim BackupModule As New InstallerModule
        BackupModule.name = BackupName
        Set Existing = GetDatabaseModule(BackupModule)
        If Not Existing Is Nothing Then
            RemoveDatabaseModule BackupModule
        End If

        ' Update name of module to backup
        Backup.name = BackupName
    End If
End Sub

Public Sub RestoreDatabaseModule(module As InstallerModule, BackupName As String)
    Dim BackupModule As New InstallerModule
    Dim Backup As Object
    Dim Existing As Object

    BackupModule.name = BackupName
    Set Backup = GetDatabaseModule(BackupModule)
    If Not Backup Is Nothing Then
        ' Remove restored module (if found)
        Set Existing = GetDatabaseModule(module)
        If Not Existing Is Nothing Then
            RemoveDatabaseModule module
        End If

        ' Restore backup
        Backup.name = module.name
    End If
End Sub
