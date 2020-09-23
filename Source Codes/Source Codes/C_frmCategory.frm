VERSION 5.00
Begin VB.Form C_frmCategory 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Customers Information"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3720
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   3720
      TabIndex        =   6
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox txtname 
      Height          =   375
      Left            =   120
      MaxLength       =   45
      TabIndex        =   5
      Top             =   960
      Width           =   3375
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1725
      TabIndex        =   4
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2625
      TabIndex        =   3
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label lblError 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   405
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   1485
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Category Name"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1125
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fill The Category Information"
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
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   2475
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   0
      Picture         =   "C_frmCategory.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3810
   End
End
Attribute VB_Name = "C_frmCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ShowHide As Boolean

Private Sub cmdOk_Click()
If ADDING = True Then
    
    If Trim(txtname.text) <> "" Then
    
    Call CatSave
Else
    MsgBox "Please Input Category name", vbInformation
End If
Else
    Call CatUPdate
    
End If
    Call GridRefresh
    
End Sub
Private Sub CatSave()
Dim CDb As CDbase
Dim CIns As New CInsert
Dim CustID As String

Call GetNewConnection2

Set Rs1 = New Recordset
Set Rs1 = DCON.Execute("Select * from Category where category='" & txtname.text & "'")

If Rs1.RecordCount = 0 Then

Call GetNewConnection(CIns)
Set CDb = CIns


CDb.TableName = "category"


CIns.FieldVal txtname, CText



CIns.Insert

txtname.text = ""
MsgBox "Record has been saved", vbInformation

Set CIns = Nothing

Else

    MsgBox "The Category Name was exist", vbInformation, "Category"
End If

End Sub
Private Sub CatUPdate()


    Call GetNewConnection2
    Set Rs1 = New Recordset
    Set Rs1 = DCON.Execute("Select * from category where category='" & txtname.text & "'")
        
        If Rs1.RecordCount = 0 Then
          
            DCON.Execute "Update product set category='" & Trim(txtname.text) & "' where category='" & MODIFYID & "'"
            DCON.Execute "Update category set category='" & Trim(txtname.text) & "' where category='" & MODIFYID & "'"
          
          
            MsgBox "Category has been updated", vbInformation
            Unload Me
          
          Else
          
            If txtname.text <> MODIFYID Then
            
                MsgBox "The Category was already exist", vbInformation
            Else
                 MsgBox "Category has been updated", vbInformation
            Unload Me
            End If
        
        End If
    Set Rs1 = Nothing
    Set DCON = Nothing
    
        


End Sub
Private Sub Form_Activate()
On Error Resume Next

If ADDING = False Then
    Call GetNewConnection2
    Set Rs1 = New Recordset
    Set Rs1 = DCON.Execute("Select * from category where category='" & MODIFYID & "'")

    If Rs1.RecordCount <> 0 Then
        txtname.text = Rs1!category & " "
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

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    If ShowHide = False Then
        Me.Width = Me.Width + (List1.Width + 600)
        Me.Left = Me.Left - ((List1.Width + 600) \ 2)
        ShowHide = True
        Command1.Caption = "<<"
    Else
        Me.Width = Me.Width - (List1.Width + 600)
        Me.Left = Me.Left + ((List1.Width + 600) \ 2)
        Command1.Caption = ">>"
        ShowHide = False
    End If
List1.Enabled = ShowHide
End Sub

Private Sub Form_Load()


    

    
    lblError.BackColor = vbWhite
End Sub

Private Sub txtname_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))


End Sub
