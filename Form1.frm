VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Visualicious - Visual Styles Wizard"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   5745
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   5745
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4650
      Left            =   5880
      ScaleHeight     =   4650
      ScaleWidth      =   5355
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   5360
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove From List"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   120
         TabIndex        =   9
         Top             =   3960
         Width           =   1935
      End
      Begin ComctlLib.ListView ListView1 
         Height          =   3375
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   5953
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Application's Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Application Path"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   ""
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton cmdRestore 
         Caption         =   "Restore Application"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3240
         TabIndex        =   8
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recovery List:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   1200
      End
   End
   Begin VB.TextBox tbx2 
      Height          =   2895
      Left            =   2160
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "Form1.frx":1CFA
      Top             =   8040
      Width           =   5775
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4650
      Left            =   240
      ScaleHeight     =   4650
      ScaleWidth      =   5355
      TabIndex        =   2
      Top             =   480
      Width           =   5360
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse For Application"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   120
         TabIndex        =   4
         Top             =   3960
         Width           =   1935
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update Application"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3240
         TabIndex        =   3
         Top             =   3960
         UseMaskColor    =   -1  'True
         Width           =   1935
      End
      Begin VB.Label tbxPath 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1320
         TabIndex        =   22
         Top             =   1440
         Width           =   3975
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File Path:"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   660
      End
      Begin VB.Label lblADescription 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1320
         TabIndex        =   20
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label lblACompany 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1320
         TabIndex        =   19
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label lblAProdName 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1320
         TabIndex        =   18
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company:"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name:"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File Name:"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   750
      End
      Begin VB.Label lblAFileName 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1320
         TabIndex        =   13
         Top             =   1200
         Width           =   3255
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Application Information:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"Form1.frx":1EA6
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1035
         Left            =   240
         TabIndex        =   11
         Top             =   2400
         Width           =   4725
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Notice!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Width           =   585
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9600
      Top             =   8040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox tbx1 
      Height          =   1605
      Left            =   2160
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":1FEE
      Top             =   7560
      Width           =   6015
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   5055
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   8916
      TabWidthStyle   =   2
      TabFixedWidth   =   3351
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Update Application"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Application Recovery"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuBA 
         Caption         =   "Browse For Application"
      End
      Begin VB.Menu mnuUA 
         Caption         =   "Update Application"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu MnuSettings 
      Caption         =   "Options"
      Begin VB.Menu mnuRA 
         Caption         =   "Always Run Application After Updating"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuReadme 
         Caption         =   "ReadMe.txt"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private FSO As New FileSystemObject

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function BeginUpdateResource Lib "kernel32" Alias "BeginUpdateResourceA" _
        (ByVal pFileName As String, _
        ByVal bDeleteExistingResources As Long) As Long

Private Declare Function UpdateResource Lib "kernel32" Alias "UpdateResourceA" _
        (ByVal hUpdate As Long, _
        ByVal lpType As Integer, _
        ByVal lpName As Integer, _
        ByVal wLanguage As Long, _
        lpData As Any, _
        ByVal cbData As Long) As Long

Private Declare Function EndUpdateResource Lib "kernel32" Alias "EndUpdateResourceA" _
        (ByVal hUpdate As Long, _
        ByVal fDiscard As Long) As Long
Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long


Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const CS_DROPSHADOW = &H20000
Private Const GCL_STYLE = (-26)
Private Declare Function InitCommonControls Lib "COMCTL32" () As Long

Private sManifest As String, strFile As String



Private Sub cmdBrowse_Click()






    With CommonDialog1
    .DialogTitle = "Portable Executable"
    .CancelError = False
    .FileName = ""
    
    .Filter = "Portable Executible|*.exe"
    .MaxFileSize = 32000
    .ShowOpen
    
    
    End With
    If CommonDialog1.FileName = "" Then Exit Sub
    lblAFileName.Caption = CommonDialog1.FileTitle
    tbxPath.Caption = CommonDialog1.FileName
    
    If lblAFileName.Caption = "" Then
    cmdUpdate.Enabled = False
    mnuUA.Enabled = False
    Else
    cmdUpdate.Enabled = True
    mnuUA.Enabled = True
    End If
    
    
    'load application's company name
    lblACompany = GetCompanyName(CommonDialog1.FileName)
    
    'load application's description
    lblADescription = GetFileDescription(CommonDialog1.FileName)
    
    'load application's product name
    lblAProdName = GetProductName(CommonDialog1.FileName)
End Sub

Private Sub cmdUpdate_Click()


strFile = tbxPath.Caption
CreateManifest
ApplyManifest sManifest, Len(sManifest)

'save to manifest settings
SaveRegString HKEY_LOCAL_MACHINE, "Software\OSI\Visualicious\Recovery\" & lblAFileName.Caption, "FileName", lblAFileName.Caption
SaveRegString HKEY_LOCAL_MACHINE, "Software\OSI\Visualicious\Recovery\" & lblAFileName.Caption, "FilePath", tbxPath.Caption

'save to win reg
SaveRegString HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", tbxPath.Caption, "WIN2000"


With ListView1.ListItems.Add(, , lblAProdName)
    .SubItems(1) = tbxPath.Caption
End With

 
 If mnuRA.Checked = True Then Shell tbxPath.Caption, vbNormalFocus
 
lblACompany = ""
lblADescription = ""
lblAProdName = ""

lblAFileName.Caption = ""
tbxPath.Caption = ""
cmdUpdate.Enabled = False

cmdRestore.Enabled = True

End Sub
 

Private Sub tbxDescription_Change()
 cmdUpdate.Enabled = True
End Sub

Private Sub cmdRestore_Click()
On Error Resume Next

If ListView1.SelectedItem.Selected = False Then
MsgBox "Please Select Application To Restore"
Exit Sub
End If


'Delete  from registry
DeleteKey HKEY_LOCAL_MACHINE, "Software\OSI\Visualicious\Recovery\" & ListView1.SelectedItem.Text & ".exe"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", ListView1.SelectedItem.SubItems(1)

'delete file
FSO.DeleteFile (ListView1.SelectedItem.SubItems(1))
FSO.CopyFile ListView1.SelectedItem.SubItems(1) & ".original", ListView1.SelectedItem.SubItems(1)
Pause (2)
FSO.DeleteFile (ListView1.SelectedItem.SubItems(1) & ".original")

'Delete from listview
ListView1.ListItems.Remove (ListView1.SelectedItem.Index)

If ListView1.ListItems.Count = 0 Then cmdRestore.Enabled = False
MsgBox "Application Restored"
End Sub

Private Sub cmdRemove_Click()
'On Error Resume Next

If ListView1.SelectedItem.Selected = False Then
MsgBox "Please Select Application To Restore"
Exit Sub
End If


'Delete  from registry
DeleteKey HKEY_LOCAL_MACHINE, "Software\OSI\Visualicious\Recovery\" & ListView1.SelectedItem.Text & ".exe"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", ListView1.SelectedItem.SubItems(1)


'Delete from listview
ListView1.ListItems.Remove (ListView1.SelectedItem.Index)

If ListView1.ListItems.Count = 0 Then cmdRestore.Enabled = False

End Sub

Private Sub Form_Initialize()
Call InitCommonControls
End Sub

Private Sub Form_Load()
On Error Resume Next

Dim rApp As String, fName As String, fPath As String, prodname As String, resi As Long

SetClassLong Me.hWnd, GCL_STYLE, GetClassLong(Me.hWnd, GCL_STYLE) Or CS_DROPSHADOW






Picture1.Left = 240
Picture2.Left = 240

resi = GetRegLong(HKEY_LOCAL_MACHINE, "Software\OSI\Visualicious", "AutoRun")

If resi = 1 Then mnuRA.Checked = True


'load saved info to listview

countit = CountRegKeys(HKEY_LOCAL_MACHINE, "Software\OSI\Visualicious\Recovery")

For i = 0 To countit - 1
fApp = GetRegKey(HKEY_LOCAL_MACHINE, "Software\OSI\Visualicious\Recovery", i)
fName = GetRegString(HKEY_LOCAL_MACHINE, "Software\OSI\Visualicious\Recovery\" & fApp, "FileName")
fPath = GetRegString(HKEY_LOCAL_MACHINE, "Software\OSI\Visualicious\Recovery\" & fApp, "FilePath")
prodname = GetProductName(GetRegString(HKEY_LOCAL_MACHINE, "Software\OSI\Visualicious\Recovery\" & fApp, "FilePath"))
    
    With ListView1.ListItems.Add(, , prodname)
        .SubItems(1) = fPath
    End With


Next i

If ListView1.ListItems.Count = 0 Then cmdRestore.Enabled = False

End Sub



Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuBA_Click()
cmdBrowse_Click
End Sub

Private Sub mnuExit_Click()
Unload Me
End
End Sub

Private Sub mnuRA_Click()
If mnuRA.Checked = False Then
mnuRA.Checked = True
SaveRegLong HKEY_LOCAL_MACHINE, "Software\OSI\Visualicious", "AutoRun", 1

Else
mnuRA.Checked = False
SaveRegLong HKEY_LOCAL_MACHINE, "Software\OSI\Visualicious", "AutoRun", 0

End If

End Sub

Private Sub mnuUA_Click()
If cmdUpdate.Enabled = True Then cmdUpdate_Click
End Sub



Private Sub TabStrip1_Click()
Select Case TabStrip1.SelectedItem.Index
Case 1
Picture1.Visible = True
Picture2.Visible = False
Case 2
Picture2.Visible = True
Picture1.Visible = False


End Select
End Sub



Public Function CreateManifest()
        sManifest = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & " standalone=" & Chr(34) & "yes" & Chr(34) & "?>"
        sManifest = sManifest & Chr(13) & "<assembly xmlns=" & Chr(34) & "urn:schemas-microsoft-com:asm.v1" & Chr(34) & " manifestVersion=" & Chr(34) & "1.0" & Chr(34) & ">"
        sManifest = sManifest & Chr(13) & "<assemblyIdentity"
        sManifest = sManifest & Chr(13) & "  name=" & Chr(34) & lblAProdName & Chr(34)
        sManifest = sManifest & Chr(13) & "  processorArchitecture=" & Chr(34) & "x86" & Chr(34)
        sManifest = sManifest & Chr(13) & "  version=" & Chr(34) & "1.0.0.0" & Chr(34)
        sManifest = sManifest & Chr(13) & "  type=" & Chr(34) & "win32" & Chr(34) & "/>"
        sManifest = sManifest & Chr(13) & "<description>" & Me.lblADescription & "</description>"
        sManifest = sManifest & Chr(13) & "<dependency>"
        sManifest = sManifest & Chr(13) & "  <dependentAssembly>"
        sManifest = sManifest & Chr(13) & "    <assemblyIdentity"
        sManifest = sManifest & Chr(13) & "      type=" & Chr(34) & "win32" & Chr(34)
        sManifest = sManifest & Chr(13) & "      name=" & Chr(34) & "Microsoft.Windows.Common-Controls" & Chr(34)
        sManifest = sManifest & Chr(13) & "      version=" & Chr(34) & "6.0.0.0" & Chr(34)
        sManifest = sManifest & Chr(13) & "      processorArchitecture=" & Chr(34) & "x86" & Chr(34)
        sManifest = sManifest & Chr(13) & "      publicKeyToken=" & Chr(34) & "6595b64144ccf1df" & Chr(34)
        sManifest = sManifest & Chr(13) & "      language=" & Chr(34) & "*" & Chr(34)
        sManifest = sManifest & Chr(13) & "    />"
        sManifest = sManifest & Chr(13) & "  </dependentAssembly>"
        sManifest = sManifest & Chr(13) & "</dependency>"
        sManifest = sManifest & Chr(13) & "</assembly>"
End Function

Private Sub ApplyManifest(lpData As String, ByVal cbData As Long)

    Dim hUpdateRes As Long, lRet As Long
    FSO.CopyFile strFile, strFile & ".original", True
    Pause (3)
    
        
    'get handle for UpdateResource. strFile must be an executable.
    hUpdateRes = BeginUpdateResource(strFile, False)

    If hUpdateRes = 0 Then GoTo FileError
    
    'modify the resource.
    lRet = UpdateResource(hUpdateRes, 24, 1, 1033, ByVal lpData, cbData)
    
    'commit the changes to the executable file.
    lRet = EndUpdateResource(hUpdateRes, False)

    MsgBox lblAProdName & " Updated"
    
Exit Sub

FileError:
        MsgBox "Could not " & lblAFileName.Caption & " for writing!"
        
End Sub
Public Function Pause(Seconds)
    Dim CurrentTime As Long
    CurrentTime = Timer
    
    Do
        DoEvents
        Loop Until CurrentTime + Seconds <= Timer
    End Function
