VERSION 5.00
Begin VB.Form C_frmLocation 
   BackColor       =   &H80000014&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Location Information"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3570
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00404040&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   3570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtremarks 
      Height          =   1095
      Left            =   120
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   1680
      Width           =   3375
   End
   Begin VB.TextBox txtname 
      Height          =   375
      Left            =   120
      MaxLength       =   40
      TabIndex        =   5
      Top             =   1080
      Width           =   3375
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2565
      TabIndex        =   4
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1650
      TabIndex        =   3
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label lbl_cust 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   675
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Location Name"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1050
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fill The Location Information Sheet"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3000
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      Picture         =   "C_frmLocation.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3570
   End
End
Attribute VB_Name = "C_frmLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
If ADDING = True Then
If Trim(txtname.text) <> "" Then
    
    Call LocSave
Else
    MsgBox "Please Input atleast Location name", vbInformation
End If

Else
    Call LocUpdate
    
End If
    Call GridRefresh
    



End Sub
Private Sub LocUpdate()
 Call GetNewConnection2
    Set Rs1 = New Recordset
    Set Rs1 = DCON.Execute("Select * from location where location='" & txtname.text & "'")
        
        If Rs1.RecordCount = 0 Then
          
            DCON.Execute "Update location set location='" & Trim(txtname.text) & "',note1='" & txtremarks & "' where location='" & MODIFYID & "'"
            MsgBox "Record has been updated", vbInformation
            Unload Me
          Else
          
            If txtname.text <> MODIFYID Then
            
               MsgBox "Record is already exist", vbInformation
             
             Else
            MsgBox "Record has been updated", vbInformation
             Unload Me
            End If
        
        End If
    Set Rs1 = Nothing
    Set DCON = Nothing
    
        
End Sub
Private Sub LocSave()
Dim CDb As CDbase
Dim CIns As New CInsert
Dim CustID As String


Call GetNewConnection(CIns)
Set CDb = CIns



CDb.TableName = "Location"


CIns.FieldVal txtname, CText
CIns.FieldVal txtremarks, CText

CIns.Insert
    MsgBox "Record has been saved", vbInformation
    txtname.text = ""
    txtremarks.text = ""
'For Each Control In C_frmCustomer
'    If TypeOf Control Is TextBox Then
'        Control.text = ""
'    End If
'Next

Set CIns = Nothing
End Sub

Private Sub Form_Activate()
On Error Resume Next

If ADDING = False Then
    Call GetNewConnection2
    Set Rs1 = New Recordset
    Set Rs1 = DCON.Execute("Select * from location where location='" & MODIFYID & "'")

    If Rs1.RecordCount <> 0 Then
        txtname.text = Rs1!Location & " "
        txtremarks.text = Rs1!note1 & " "
        
    End If
    Set Rs1 = Nothing
    Set DCON = Nothing
    
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub txtname_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub
