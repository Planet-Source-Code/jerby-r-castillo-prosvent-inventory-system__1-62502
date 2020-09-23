VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4590
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Login"
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "l"
      TabIndex        =   1
      Text            =   "jerby"
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Text            =   "jerby"
      Top             =   840
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Users Name and Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   435
      Left            =   1800
      TabIndex        =   7
      Top             =   120
      Width           =   1515
   End
   Begin VB.Image Image1 
      Height          =   630
      Left            =   0
      Picture         =   "frmLogin.frx":0000
      Top             =   0
      Width           =   4500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "errLabel"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image Image4 
      Height          =   15
      Left            =   0
      Picture         =   "frmLogin.frx":0F0F
      Stretch         =   -1  'True
      Top             =   705
      Width           =   4635
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00696969&
      Height          =   210
      Index           =   1
      Left            =   360
      TabIndex        =   6
      Top             =   1440
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00696969&
      Height          =   210
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   840
      Width           =   885
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000001&
      Height          =   855
      Left            =   0
      Top             =   -120
      Width           =   4575
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public OK As Boolean
Private Sub cmdCancel_Click()
  'End
  Unload Me
  
  
   
End Sub
Private Sub cmdOk_Click()

  Call GetNewConnection2
Set Rs1 = New Recordset


Set Rs1 = DCON.Execute("Select * from users")

If Rs1.RecordCount <= 0 Then

    If Text1.text = "admin" Then
        If Text2.text = "password" Then
            CurUser = Text1.text
            'Unload Me
             OK = True
            Me.Hide
            Load frmMAIN
            frmMAIN.Show
        Else
                MsgBox "Access Denied!   ", vbCritical, "Log In"
                Text2.SetFocus
                Text2.SelStart = 0
                Text2.SelLength = Len(Text2.text)
                
        End If
    Else
                MsgBox "Access Denied!   ", vbCritical, "Log In"
                Text2.SetFocus
                Text2.SelStart = 0
                Text2.SelLength = Len(Text2.text)
                
    End If
    
ElseIf Rs1.RecordCount >= 1 Then
       If Text1.text = "admin" Then
        If Text2.text = "password" Then
            CurUser = Text1.text
           ' Unload Me
            OK = True
            Me.Hide
            Load frmMAIN
            frmMAIN.Show
            Exit Sub
        End If
        End If
        
    Set Rs1 = DCON.Execute("Select * from users where username1='" & Text1.text _
                 & "' And password1='" & Text2.text & "'")
    
    If Rs1.RecordCount > 0 Then
            CurUser = Text1.text
            'Unload Me
             OK = True
            Me.Hide
'            Load frmMAIN
Load frmSplash
frmSplash.Show
            'frmMAIN.Show
            
    ElseIf Rs1.RecordCount = 0 Then
         MsgBox "Access Denied!   ", vbCritical, "Log In"
            Text1.SetFocus
                Text1.SelStart = 0
                Text1.SelLength = Len(Text1.text)
              
    End If

        
End If

    

Set Rs1 = Nothing
Set DCON = Nothing
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call cmdOk_Click
End If
End Sub
