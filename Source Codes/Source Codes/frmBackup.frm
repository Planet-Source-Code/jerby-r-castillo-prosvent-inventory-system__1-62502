VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBackup 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BackUp"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00F8D9CB&
      Caption         =   "Last BackUp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   0
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   6375
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   2760
         Width           =   6135
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         Width           =   3495
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Label7"
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   2520
         Width           =   6135
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Label8"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   3480
         Width           =   3615
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Label9"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   3240
         Width           =   3135
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5400
         TabIndex        =   15
         Top             =   3960
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00F8D9CB&
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   1440
      TabIndex        =   11
      Top             =   960
      Visible         =   0   'False
      Width           =   3495
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00DC705C&
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   13
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label12"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   2970
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00000000&
      Caption         =   "View Database Size"
      Height          =   495
      Left            =   480
      TabIndex        =   9
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000000&
      Caption         =   "View System Size"
      Height          =   495
      Left            =   1800
      TabIndex        =   8
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Left            =   5400
      Top             =   3120
   End
   Begin MSComctlLib.ProgressBar p1 
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   3840
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   101
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdsave2 
      BackColor       =   &H00000000&
      Caption         =   "Back Up System"
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00000000&
      Caption         =   "Back Up Database"
      Height          =   495
      Left            =   4680
      TabIndex        =   3
      Top             =   2520
      Width           =   1455
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   1665
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   2775
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   1560
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.FileListBox File2 
      Appearance      =   0  'Flat
      Height          =   1590
      Left            =   3360
      TabIndex        =   7
      Top             =   780
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "View Last BackUp Information"
      Height          =   495
      Left            =   1800
      TabIndex        =   10
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   4080
      Width           =   1215
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit








Private Sub Command2_Click()
On Error GoTo CMD2ERR

 Dim Fsys2, File2, msg2
    Set Fsys2 = CreateObject("Scripting.FileSystemObject")
    Set File2 = Fsys2.GetFile(App.Path & "\Database\Thesis.mdb")
    msg2 = File2.Name & " uses " & Left(File2.size / 1000000, 5) & " MB."
    Frame2.Visible = True
    Frame2.Caption = "Database Info"
    Label12.Caption = msg2
    
Exit Sub
CMD2ERR:
    MsgBox Err.Description, vbInformation
    
End Sub

Private Sub Command3_Click()
On Error GoTo CMD3ERR

frmBackup.Height = 5085
Frame1.Visible = True
    Dim LastPath, SysLastPath As String
    Dim LastDate, SysLastDate As String
    Dim LastTime, SysLastTime As String
    Dim fsys, Fldr, Nme
    
    Set fsys = CreateObject("Scripting.FileSystemObject")
    Set Fldr = fsys.GetFolder(App.Path)
    Nme = Fldr.Name
    
    LastPath = GetSetting(App.title, "Settings", "BackupPath")
    LastDate = GetSetting(App.title, "Settings", "BackupDate")
    LastTime = GetSetting(App.title, "Settings", "BackupTime")
    
    SysLastPath = GetSetting(App.title, "Settings", "SysBackupPath")
    SysLastDate = GetSetting(App.title, "Settings", "SysBackupDate")
    SysLastTime = GetSetting(App.title, "Settings", "SysBackupTime")
    
    
    If LastPath = "" Then
        Label1.Caption = "No Backup made previously"
        Label2.Caption = " "
        Label3.Caption = " "
    Else
        Label1.Caption = "Path: " & LastPath
        Label2.Caption = "Date: " & LastDate
        Label3.Caption = "Time: " & LastTime
    End If
    
    If SysLastPath = "" Then
        Label7.Caption = "No Backup made previously"
        Label8.Caption = " "
        Label9.Caption = " "
      
    Else
        
        Label7.Caption = "Path: " & SysLastPath & Nme
        Label8.Caption = "Date: " & SysLastDate
        Label9.Caption = "Time: " & SysLastTime
    End If

Exit Sub
CMD3ERR:
    MsgBox Err.Description, vbInformation
    
End Sub

Private Sub cmdsave_click()
On Error GoTo BackUPerr
    cmdSave.Enabled = False
    cmdsave2.Enabled = False

    
    CURRDATE = Format$(Now, "dddd mmmm dd, yyyy")
    CURRTIME = Format$(Now, "hh:mm:ss AM/PM")
    DESTINATION = File1.Path & "Thesis.mdb"
    Source = App.Path & "\Database\Thesis.mdb"
    
 
    FSYSTEM.CopyFile Source, DESTINATION, True
    

    SaveSetting App.title, "Settings", "BackupPath", DESTINATION
    SaveSetting App.title, "Settings", "BackupDate", CURRDATE
    SaveSetting App.title, "Settings", "BackupTime", CURRTIME
    
    cmdsave2.Enabled = True
    cmdSave.Enabled = True
       Frame2.Visible = True
    Frame2.Caption = "Backup Info"
    Label12.Caption = "Backup Process Success"
    File2.Refresh
Exit Sub
BackUPerr:
    MsgBox Err.Description, vbInformation & Space(5), "Error"
cmdSave.Enabled = True
    cmdsave2.Enabled = True
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Drive1_Change()
On Error GoTo DrvErr
    Dir1.Path = Drive1.Drive
Exit Sub
DrvErr:
    MsgBox Err.Description & Space(5), vbInformation
    Drive1.ListIndex = 1
    
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    File2.Path = Dir1.Path
    
End Sub



Private Sub cmdsave2_Click()
p1.Visible = True
Timer1.Interval = 1

  cmdsave2.Enabled = False
    cmdSave.Enabled = False
frmBackup.Height = 4860


End Sub
Private Sub Command1_Click()
On Error GoTo CMD1ERR
    Dim fsys, Folder2, msg
    Set fsys = CreateObject("Scripting.FileSystemObject")
    Set Folder2 = fsys.GetFolder(App.Path)
    msg = Folder2.Name & " uses " & Left(Folder2.size / 1000000, 5) & " MB."
     Frame2.Visible = True
    Frame2.Caption = "System Info"
    Label12.Caption = msg
Exit Sub
CMD1ERR:
    MsgBox Err.Description, vbInformation
    
End Sub

Private Sub Form_Activate()
frmBackup.Height = 4320
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.ForeColor = vbBlack

End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.ForeColor = vbBlack
End Sub

Private Sub Label10_Click()
Frame1.Visible = False
frmBackup.Height = 4320
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label10.ForeColor = vbBlue

End Sub

Private Sub Label13_Click()
Frame2.Visible = False
End Sub

Private Sub Timer1_Timer()
On Error GoTo SysCopyErr
Dim BProg As Long
 Dim Fsys2, File2, msg2
    Set Fsys2 = CreateObject("Scripting.FileSystemObject")
 
    Set File2 = Fsys2.GetFile(App.Path & "\Database\Thesis.mdb")
    msg2 = File2.size

   
   
    'p1.Max = msg2
    BProg = msg2 / 2
    Label4.AutoSize = True
   Label4.Caption = "Backup Progress " & p1.Value & "%"
  
   
If p1.Value < 101 Then
    p1.Value = p1.Value + 1
 
      
       If p1.Value = 25 Then
  
    DESTINATION = File1.Path & "\"
    Source = App.Path
     DoEvents
     ElseIf p1.Value = 83 Then
     
    FSYSTEM.CopyFolder Source, DESTINATION, True
    
    

        DoEvents
    
    ElseIf p1.Value = 101 Then
      Label4.AutoSize = True
     Label4.Caption = "Backup Progress Complete"
       Frame2.Visible = True
    Frame2.Caption = "Backup Info"
    Label12.Caption = "Backup Process Success"
    frmBackup.Height = 4770
       CURRDATE = Format$(Now, "dddd mmmm dd, yyyy")
        CURRTIME = Format$(Now, "hh:mm:ss AM/PM")
    
    SaveSetting App.title, "Settings", "SysBackupPath", DESTINATION
    SaveSetting App.title, "Settings", "SysBackupDate", CURRDATE
    SaveSetting App.title, "Settings", "SysBackupTime", CURRTIME
    
     Label4.Caption = ""
    Timer1.Interval = 0
    p1.Value = 0
    p1.Visible = False
   
    cmdSave.Enabled = True
    cmdsave2.Enabled = True
    Dir1.Refresh
   End If

End If

Exit Sub

SysCopyErr:
    MsgBox Err.Description & Space(5), vbInformation
  Timer1.Interval = 0
    p1.Value = 0
    p1.Visible = False
    Label4.Caption = ""
    cmdSave.Enabled = True
    cmdsave2.Enabled = True
End Sub
