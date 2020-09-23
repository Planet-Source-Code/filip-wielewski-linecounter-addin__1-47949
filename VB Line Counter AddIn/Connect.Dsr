VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   9975
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   10650
   _ExtentX        =   18785
   _ExtentY        =   17595
   _Version        =   393216
   Description     =   $"Connect.dsx":0000
   DisplayName     =   "LineCounter AddIn"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Command Line / Startup"
   LoadBehavior    =   5
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSafe     =   -1  'True
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public FormDisplayed          As Boolean
Public VBInstance             As VBIDE.VBE
Dim mcbMenuCommandBar         As Office.CommandBarControl
Dim mfrmMain                  As New frmMain
Public WithEvents MenuHandler As CommandBarEvents          'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1


Sub Hide()
    
    On Error Resume Next
    
    FormDisplayed = False
    mfrmMain.Hide
    
End Sub

Sub Show()
  
    On Error GoTo Blad
    
    booShow = True
    
    Dim intProjects As Integer      'will store number of opened projects
    Dim intComponents As Integer    'will store number of components...
                                    '...in chosen project
    Dim intI As Integer             'variable for "For...Next"
    
    If mfrmMain Is Nothing Then
        Set mfrmMain = New frmMain
    End If
    
    Set mfrmMain.VBInstance = VBInstance
    Set mfrmMain.Connect = Me
    FormDisplayed = True
    
    'clear both lists
    mfrmMain.lstProject.Clear
    mfrmMain.lstComponent.Clear
    
    'get number of opened projects
    intProjects = VBInstance.VBProjects.Count
    'if there is more than one opened project then:
    If intProjects <> 0 Then
        'add name of each project to lstProject
        For intI = 1 To intProjects
            mfrmMain.lstProject.AddItem VBInstance.VBProjects.Item(intI)
        Next intI
        'set active item of lstProject
        mfrmMain.lstProject.Text = VBInstance.VBProjects.Item(1)
        'get number of components of selected project
        intComponents = VBInstance.VBProjects.Item(1).VBComponents.Count
        'if there is more than one component then:
        If intComponents <> 0 Then
            'add name of each component to lstComponent
            For intI = 1 To intComponents
                If VBInstance.VBProjects.Item(1).VBComponents.Item(intI).Name <> "" Then mfrmMain.lstComponent.AddItem VBInstance.VBProjects.Item(1).VBComponents.Item(intI).Name
            Next intI
            'set active item of lstComponent
            mfrmMain.lstComponent.Text = VBInstance.VBProjects.Item(1).VBComponents.Item(1).Name
        End If
    End If
    
    booShow = False
    mfrmMain.Show 1
    
    Exit Sub
    
Blad:
    MsgBox Err.Description
    booShow = False
    mfrmMain.Show 1
    
End Sub

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    
    'save the vb instance
    Set VBInstance = Application
    
    'this is a good place to set a breakpoint and
    'test various addin objects, properties and methods
    Debug.Print VBInstance.FullName

    If ConnectMode = ext_cm_External Then
        'Used by the wizard toolbar to start this wizard
        Me.Show
    Else
        Set mcbMenuCommandBar = AddToAddInCommandBar("LineCounter AddIn")
        'sink the event
        Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
    End If
  
    If ConnectMode = ext_cm_AfterStartup Then
        If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
            'set this to display the form on connect
            Me.Show
        End If
    End If
    
    Exit Sub
    
error_handler:
    
    MsgBox Err.Description
    
End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    'delete the command bar entry
    mcbMenuCommandBar.Delete
    
    'shut down the Add-In
    If FormDisplayed Then
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"
        FormDisplayed = False
    Else
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "0"
    End If
    
    Unload mfrmMain
    Set mfrmMain = Nothing
    
End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
        'set this to display the form on connect
        Me.Show
    End If
    
End Sub

'this event fires when the menu is clicked in the IDE
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Me.Show
End Sub

Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarControl  'command bar object
    Dim cbMenu As Object
  
    On Error GoTo AddToAddInCommandBarErr
    
    'see if we can find the Add-Ins menu
    Set cbMenu = VBInstance.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then
        'not available so we fail
        Exit Function
    End If
    
    'add it to the command bar
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    'set the caption
    cbMenuCommandBar.Caption = sCaption
    
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
    
AddToAddInCommandBarErr:

End Function

