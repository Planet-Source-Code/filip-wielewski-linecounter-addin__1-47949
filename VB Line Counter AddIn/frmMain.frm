VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LineCounter AddIn"
   ClientHeight    =   5730
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   5190
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   382
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   346
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmOpt 
      Caption         =   "Count:"
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   4935
      Begin VB.OptionButton optAll 
         Caption         =   "lines in all components in selected project"
         Height          =   350
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   4695
      End
      Begin VB.OptionButton optOne 
         Caption         =   "lines in 1 selected component"
         Height          =   350
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   4575
      End
   End
   Begin VB.ListBox lstProject 
      Height          =   645
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   600
      Width           =   4935
   End
   Begin VB.ListBox lstComponent 
      Height          =   1815
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   1680
      Width           =   4935
   End
   Begin VB.TextBox txtResult 
      Alignment       =   2  'Center
      Height          =   615
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   4800
      Width           =   3495
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton cmdCount 
      Caption         =   "Count lines"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label lblAutor 
      Caption         =   "By Filip Wielewski"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
   Begin VB.Image imgICO 
      Height          =   495
      Left            =   2280
      Top             =   45
      Width           =   495
   End
   Begin VB.Label lblComponent 
      Caption         =   "Component:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lblProject 
      Caption         =   "Project:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'=============================================================='
'=                                                            ='
'= ======AUTHOR======                                         ='
'= THIS IS A FREE CODE                                        ='
'= BY FILIP WIELEWSKI                                         ='
'= E-MAIL: WIELFILIST@WP.PL                                   ='
'=                                                            ='
'= ======SORRY FOR:======                                     ='
'= my bad english which I use in descriptions :]              ='
'=                                                            ='
'=============================================================='

'--To install this AddIn you have to click "File" and select...
'----..."Make LineCounterAddIn.dll" (compile it to VB's folder - it is important)
'--find file "vbaddin.ini" in windows' directory (ex. c:\WINNT\vbaddin.ini)...
'----...and add this line: "LineCounterAddIn.Connect=3"
'--Now restart your Visual Basic and you can use "LineCounter AddIn" from...
'----..."Add-Ins" menu


Option Explicit

Public VBInstance As VBIDE.VBE      'VB Instance
Attribute VBInstance.VB_VarHelpID = -1
Public Connect As Connect           'Connact as connect
Attribute Connect.VB_VarHelpID = -1
Dim strVBProject As String          'this string will store name of current...
                                    '...project
Dim strVBComponent As String        'this string will store name of current...
                                    '...component

Private Sub cmdCount_Click()
    
    On Error GoTo Blad

    Dim lonLines As Long            'will store number of lines of component
    Dim intI As Integer             'variable for "For...Next"
    Dim intListCount As Integer     'will store number of components
    
    'get name of selected project
    strVBProject = lstProject.Text
    'get name of selected component
    strVBComponent = lstComponent.Text
    
    'if optOne.Value = True then count lines in one selected component
    If optOne.Value = True Then
        
        'get number of lines in selected component
        lonLines = Str(VBInstance.VBProjects.Item(strVBProject).VBComponents.Item(strVBComponent).CodeModule.CountOfLines)
        
        'display result in txtResult
        Select Case lonLines
            Case 0
                txtResult.Text = BigLetter(strVBComponent) & " hasn't any line of code."
            Case 1
                txtResult.Text = BigLetter(strVBComponent) & " has 1 line of code."
            Case Else
                txtResult.Text = BigLetter(strVBComponent) & " has " & lonLines & " lines of code."
        End Select
        
    'if optAll.Value = True then count all lines in selected project
    Else
        intListCount = lstComponent.ListCount
        lonLines = 0
        For intI = 1 To intListCount
            lonLines = lonLines + Str(VBInstance.VBProjects.Item(strVBProject).VBComponents.Item(intI).CodeModule.CountOfLines)
        Next intI
        Select Case lonLines
            Case 0
                txtResult.Text = BigLetter(strVBProject) & " hasn't any line of code."
            Case 1
                txtResult.Text = BigLetter(strVBProject) & " has 1 line of code."
            Case Else
                txtResult.Text = BigLetter(strVBProject) & " has " & lonLines & " lines of code."
        End Select
    End If
    
    Exit Sub
    
Blad:
    MsgBox Err.Description
    
End Sub

Private Sub cmdOK_Click()
    
    On Error GoTo Blad
    
    'hide frmMain
    Connect.Hide
    
    Exit Sub
    
Blad:
    MsgBox Err.Description
    Connect.Hide
        
End Sub



Private Sub Form_Load()
    
    On Error GoTo Blad
    
    imgICO.Picture = frmMain.Icon
    
    Exit Sub
    
Blad:
    MsgBox Err.Description
    
End Sub


Private Sub lstProject_Click()
    
        On Error GoTo Blad
    
        Dim intI As Integer             'variable for "For...Next"
        Dim intComponents As Integer    'will store number of components in...
                                        '...chosen project
        
        If booShow = False Then
            'clear lstComponents
            lstComponent.Clear
            'get number of components in chosen project
            intComponents = VBInstance.VBProjects.Item(lstProject.Text).VBComponents.Count
            'if there is at least one component and its name isn't "" ...
            '...(component's name is "" when component is a related document,...
            '...ex. RES file) then get name of every component and add it ...
            '...to lstComponent
            If intComponents <> 0 Then
                For intI = 1 To intComponents
                    If VBInstance.VBProjects.Item(lstProject.Text).VBComponents.Item(intI).Name <> "" Then lstComponent.AddItem VBInstance.VBProjects.Item(lstProject.Text).VBComponents.Item(intI).Name
                Next intI
                lstComponent.Text = VBInstance.VBProjects.Item(lstProject.Text).VBComponents.Item(1).Name
            End If
        End If
        
        Exit Sub
        
Blad:
        MsgBox Err.Description
    
End Sub

Private Sub optAll_Click()
    
    lstComponent.Enabled = False
    
End Sub

Private Sub optOne_Click()
    
    lstComponent.Enabled = True
    
End Sub

Private Function BigLetter(blWord As String) As String
    
    'change first letter of blWord argument into large letter
    BigLetter = UCase(Left(blWord, 1)) & Mid(blWord, 2)
    
End Function
