VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPR 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ProsVent Inventory Manager 2005"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      TabIndex        =   10
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&New"
      Height          =   375
      Left            =   4200
      TabIndex        =   8
      Top             =   7320
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Left            =   600
      Top             =   2640
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BackColor       =   &H00F9F0EB&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Left            =   0
      ScaleHeight     =   2175
      ScaleWidth      =   11850
      TabIndex        =   13
      Top             =   0
      Width           =   11910
      Begin VB.ComboBox cmbSupp 
         Height          =   315
         ItemData        =   "frmPR.frx":0000
         Left            =   9360
         List            =   "frmPR.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox txtDis 
         Enabled         =   0   'False
         Height          =   375
         Left            =   5040
         TabIndex        =   6
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox text5 
         Enabled         =   0   'False
         Height          =   375
         Left            =   5640
         TabIndex        =   26
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtRate 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3960
         TabIndex        =   5
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtQTY 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2925
         TabIndex        =   4
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         Height          =   375
         Left            =   255
         TabIndex        =   3
         Top             =   1560
         Width           =   2415
      End
      Begin VB.ComboBox cmbCust 
         Height          =   315
         Left            =   9360
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   840
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DTPICKER4 
         Height          =   330
         Left            =   9360
         TabIndex        =   1
         Top             =   1185
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   20381697
         CurrentDate     =   38530
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Enabled         =   0   'False
         Height          =   405
         Left            =   7125
         TabIndex        =   7
         Top             =   1635
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase OrderID"
         Height          =   195
         Left            =   7800
         TabIndex        =   29
         Top             =   1680
         Width           =   1545
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Disc"
         Height          =   255
         Left            =   5040
         TabIndex        =   28
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         Height          =   195
         Left            =   4200
         TabIndex        =   24
         Top             =   1350
         Width           =   390
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Registry:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F9F0EB&
         Height          =   240
         Left            =   7395
         TabIndex        =   23
         Top             =   240
         Width           =   2040
      End
      Begin VB.Label lblHead 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lipa Solid Auto Supply"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   345
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   3300
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         Height          =   195
         Left            =   7920
         TabIndex        =   14
         Top             =   960
         Width           =   705
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "xxxx"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F9F0EB&
         Height          =   240
         Left            =   9510
         TabIndex        =   17
         Top             =   255
         Width           =   2235
      End
      Begin VB.Label lblTop 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Registry"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F9F0EB&
         Height          =   345
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   2550
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Ordered"
         Height          =   195
         Left            =   7920
         TabIndex        =   15
         Top             =   1320
         Width           =   1170
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H80000004&
         BorderColor     =   &H80000001&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   1320
         Left            =   7755
         Shape           =   4  'Rounded Rectangle
         Top             =   765
         Width           =   3930
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   195
         Left            =   5745
         TabIndex        =   20
         Top             =   1380
         Width           =   660
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "R/C QTY"
         Height          =   195
         Left            =   2985
         TabIndex        =   19
         Top             =   1380
         Width           =   735
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name "
         Height          =   195
         Left            =   195
         TabIndex        =   18
         Top             =   1395
         Width           =   1260
      End
      Begin VB.Image imgTop 
         Height          =   720
         Left            =   0
         Picture         =   "frmPR.frx":0004
         Stretch         =   -1  'True
         Top             =   0
         Width           =   12330
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000004&
         BorderColor     =   &H80000001&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   780
         Left            =   60
         Top             =   1290
         Width           =   6900
      End
   End
   Begin MSComctlLib.ListView lvMain 
      Height          =   4455
      Left            =   2520
      TabIndex        =   11
      Top             =   2280
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   7858
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      Enabled         =   0   'False
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "QNTY"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "RATE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "DISCOUNT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "AMOUNT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "TOTAL"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvLook 
      Height          =   4455
      Left            =   120
      TabIndex        =   25
      Top             =   2280
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   7858
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16380139
      BorderStyle     =   1
      Appearance      =   0
      Enabled         =   0   'False
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Product"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblwords 
      BackStyle       =   0  'Transparent
      Height          =   855
      Left            =   240
      TabIndex        =   30
      Top             =   7080
      Width           =   3735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rate"
      Height          =   195
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   390
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H80000004&
      BorderColor     =   &H80000001&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   1080
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   6840
      Width           =   7170
   End
   Begin MSForms.TextBox TextAmount 
      Height          =   375
      Left            =   9600
      TabIndex        =   12
      Top             =   7080
      Width           =   1905
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "3360;661"
      BorderColor     =   -2147483647
      SpecialEffect   =   0
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   7980
      TabIndex        =   21
      Top             =   7140
      Width           =   1530
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H80000003&
      BorderColor     =   &H80000001&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00F9F0EB&
      FillStyle       =   0  'Solid
      Height          =   510
      Left            =   7920
      Shape           =   4  'Rounded Rectangle
      Top             =   7005
      Width           =   3690
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000004&
      BorderColor     =   &H80000001&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   1080
      Left            =   7815
      Shape           =   4  'Rounded Rectangle
      Top             =   6840
      Width           =   3945
   End
   Begin VB.Menu mnuLook 
      Caption         =   "Look"
      Visible         =   0   'False
      Begin VB.Menu mnuUnit 
         Caption         =   "Unit In Stock"
      End
   End
End
Attribute VB_Name = "frmPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PurchReg As Boolean
Dim PurchOrder As Boolean
Dim PurchRet As Boolean
Dim SuppName As String
Dim SuppID As String
Dim PurchID As String
Dim ReturnID As String


Dim TXTLEN As Integer
Dim STRT As Integer
Dim PRILEN As Integer

Private Sub cmbCust_Click()
Call GetNewConnection2
Set Rs1 = New Recordset
Set Rs1 = DCON.Execute("Select suppliersid from supplier where businessname='" & cmbCust & "'")

If Rs1.RecordCount <> 0 Then
    SuppID = Rs1!suppliersid
   
End If
Set Rs1 = Nothing
Set DCON = Nothing

'Call CMB1("v", "PurchaseOrderID", cmbSupp, "where BusinessName='" & cmbCust.text & "' and Deliver=0", True)
    
Call CMB2("Select DIstinct PurchaseOrderID from V where Businessname='" & cmbCust.text & "'", cmbSupp)
cmbSupp.AddItem ("<ALL>")

lvMain.ListItems.clear
Text4.text = ""
text5.text = ""
txtQTY.text = ""
txtRate.text = ""
txtDis.text = ""

EDT = False

If cmbCust.text <> "" Then
    Text4.Enabled = True
Else
    Text4.Enabled = False
End If



End Sub

Private Sub cmbSupp_click()
Dim lst As ListItem
Dim tempSql As String
GetNewConnection2
Set Rs1 = New ADODB.Recordset
lvMain.ListItems.clear
If cmbSupp.text = "<ALL>" Then
     tempSql = "select * from v Where businessName='" & Trim(cmbCust.text) & "'and deliver=0"
Else
     tempSql = "select * from v Where PurchaseOrderID='" & Trim(cmbSupp.text) & "'and deliver=0"
End If

Set Rs1 = DCON.Execute(tempSql)
    While Rs1.EOF <> True
       Set lst = lvMain.ListItems.Add(, , Rs1!ProductID)
        lst.SubItems(1) = Rs1!Name
        lst.SubItems(2) = Rs1!Quantity
        lst.SubItems(3) = Rs1!Rate
         lst.SubItems(4) = 0
        lst.SubItems(5) = Val(Rs1!Quantity) * Val(Rs1!Rate)
        lst.SubItems(6) = Val(Rs1!Quantity) * Val(Rs1!Rate)
        
        Rs1.MoveNext
    Wend

    Set DCON = Nothing
    Set Rs1 = Nothing
    


End Sub
Private Sub RegModify()
Dim lst As ListItem
Dim tempSql As String
GetNewConnection2
Set Rs1 = New ADODB.Recordset
lvMain.ListItems.clear

'If cmbSupp.text = "<ALL>" Then
'
'     tempSql = "select * from v Where businessName='" & Trim(cmbCust.text) & "'and deliver=0"
'Else
     tempSql = "select * from v Where PurchaseOrderID='" & MODIFYID & "' and deliver=0"
'End If

Set Rs1 = DCON.Execute(tempSql)
    While Rs1.EOF <> True
       Set lst = lvMain.ListItems.Add(, , Rs1!ProductID)
        lst.SubItems(1) = Rs1!Name
        lst.SubItems(2) = Rs1!Quantity
        lst.SubItems(3) = Rs1!Rate
         lst.SubItems(4) = 0
        lst.SubItems(5) = Val(Rs1!Quantity) * Val(Rs1!Rate)
        lst.SubItems(6) = Val(Rs1!Quantity) * Val(Rs1!Rate)
        
        Rs1.MoveNext
    Wend

    Set DCON = Nothing
    Set Rs1 = Nothing
End Sub
Private Sub cmdAdd_Click()
Dim LSTITEM As ListItem
Dim CNT As Boolean
Dim DD As Integer
If text5.text <> "" Then
   
    CNT = False
  

    Call GetNewConnection2
        Set Rs1 = New Recordset
        Set Rs1 = DCON.Execute("Select * from Product where ProductID like'" & Text4 & "%' OR Name like'" & Text4 & "%'")

If Not Rs1.EOF Then
    
     
'Set LSTITEM = ListView1.FindItem(RS1!productid, lvwText, , lvwPartial)
'       If LSTITEM Is Nothing Then
            
       
       'LBL_DES.Caption = RS1!ProductID & ", " & RS1!Name & ""
       txtRate = Rs1!UnitCostPrice
            
  With lvMain
        TextAmount.text = ""
        
        If .ListItems.Count <> 0 Then
          
            For DD = 1 To .ListItems.Count
               
                If InStr(1, .ListItems(DD).text, Rs1!ProductID) = 1 Then
                    If InStr(1, .ListItems(DD).SubItems(1), Rs1!Name) = 1 Then
              
                        If EDT = True Then
                            .ListItems(DD).Selected = True
                            .ListItems(DD).SubItems(2) = Val(txtQTY.text)
                            .ListItems(DD).SubItems(3) = Val(txtRate.text)
                            .ListItems(DD).SubItems(4) = Format(Val(txtDis) / 100, ".00%")
                            .ListItems(DD).SubItems(5) = text5.text
                            '.ListItems(DD).SubItems(6) = Val(Val(txtRate.text) * Val(txtqty.text) - Val(txtRate.text) * Val(txtqty.text) * Val(txtDis.text)) '* 'Val(.ListItems(DD).SubItems(5)) - Val(Val(Val(.ListItems(DD).SubItems(5)) * Val(.ListItems(DD).SubItems(4))))
                        '  .ListItems(DD).SubItems(6) = Val(.ListItems(DD).SubItems(5)) - Val(Val(Val(.ListItems(DD).SubItems(5))) * Val(.ListItems(DD).SubItems(4)))
                        
                             '     .ListItems(DD).SubItems(6) = Val(Val(txtRate.text) * Val(txtQTY.text) - Val(txtRate.text) * Val(txtQTY.text) * Val(txtDis) / 100) '* 'Val(.ListItems(DD).SubItems(5)) - Val(Val(Val(.ListItems(DD).SubItems(5)) * Val(.ListItems(DD).SubItems(4))))
                          .ListItems(DD).SubItems(6) = Val(Val(txtRate.text) * Val(txtQTY.text) - Val(txtRate.text) * Val(txtQTY.text) * Val(txtDis) / 100) '* 'Val(.ListItems(DD).SubItems(5)) - Val(Val(Val(.ListItems(DD).SubItems(5)) * Val(.ListItems(DD).SubItems(4))))
                     
                        Else
                            .ListItems(DD).Selected = True
                            .ListItems(DD).SubItems(2) = Val(.ListItems(DD).SubItems(2)) + Val(txtQTY.text)
                            .ListItems(DD).SubItems(3) = Val(txtRate.text)
                             .ListItems(DD).SubItems(4) = Format(Val(txtDis) / 100, ".00%")
                           .ListItems(DD).SubItems(5) = Val(.ListItems(DD).SubItems(2)) * Val(.ListItems(DD).SubItems(3))
                           ' .ListItems(DD).SubItems(6) = Val(.ListItems(DD).SubItems(5)) - Val(Val(Val(.ListItems(DD).SubItems(5))) * Val(.ListItems(DD).SubItems(4)))
                          .ListItems(DD).SubItems(6) = Val(Val(txtRate.text) * Val(txtQTY.text) - Val(txtRate.text) * Val(txtQTY.text) * Val(txtDis) / 100) '* 'Val(.ListItems(DD).SubItems(5)) - Val(Val(Val(.ListItems(DD).SubItems(5)) * Val(.ListItems(DD).SubItems(4))))
            
            
                     
                    
                        End If
                             
                    CNT = True
                    
                    End If
                End If
                   
            Next
       
         End If
            
        If CNT = False Then

         .ListItems.Add , , Rs1!ProductID
            .ListItems(.ListItems.Count).SubItems(1) = Rs1!Name
            .ListItems(.ListItems.Count).SubItems(2) = txtQTY.text
            .ListItems(.ListItems.Count).SubItems(3) = txtRate.text
            .ListItems(.ListItems.Count).SubItems(4) = Format(Val(txtDis) / 100, ".00%")
                           
            .ListItems(.ListItems.Count).SubItems(5) = text5.text
             
             
             ' TextAmount.Text = Val(TextAmount.Text) + Val(TXT_AMT.Text)
        
        End If
           
            
             
        
'        If .ListItems.Count <= 0 Then
'
'            .ListItems.Add 1, , RS1!ProductID
'            .ListItems(.ListItems.Count).SubItems(1) = RS1!Name
'            .ListItems(.ListItems.Count).SubItems(2) = txtqty.text
'            .ListItems(.ListItems.Count).SubItems(3) = txtrate.text
'            .ListItems(.ListItems.Count).SubItems(4) = "1"
'            .ListItems(.ListItems.Count).SubItems(5) = Text5.text
'
'
'        End If


            For DD = 1 To .ListItems.Count
                  TextAmount.text = Val(.ListItems(DD).SubItems(6)) + Val(TextAmount.text)
            Next
            
        ' lblunit.Caption = RS1!UnitsInStock
       
        
       Set Rs1 = DCON.Execute("Select * from Product")
      '  Set DataGrid1.DataSource = RS1
         
        Set Rs1 = Nothing
        Set DCON = Nothing
   
  End With

Else
    MsgBox "Product Not Found", vbInformation, "Product"
    
    
End If
    
Text4.text = ""
text5.text = ""
txtQTY.text = ""
txtRate.text = ""
txtDis.text = ""
EDT = False

'If Me.Enabled = True Then txtQTY.SetFocus:
Else


'
'txtQTY.SetFocus


End If


End Sub

Private Sub cmdAdd_GotFocus()
'Call GetNewConnection2
'
'Set RS1 = New Recordset
'SQL = "Select TOP 5 * from PRODUCT where PRODUCTID like '" & Text4 & "%' OR NAME like'" & Text4 & "%'"
'
'Set RS1 = DCON.Execute(SQL)
'
'
'
'
'    SQL = "UPDATE PRODUCT set UnitSellingPrice=" & Val(txtrate.text) & " where (PRODUCTID='" & RS1!ProductID & "' AND UnitCostPrice <" & Val(txtrate.text) & ")"
'    'MsgBox SQL
'    DCON.Execute SQL
'
'
'Set RS1 = Nothing
'Set DCON = Nothing
'
'Text5.text = Val(txtqty.text) * Val(txtrate.text)
End Sub

Private Sub cmdClear_Click()
EDT = False
lvMain.ListItems.clear

picTop.Enabled = False

lvLook.Enabled = False
lvMain.Enabled = False
Text4.text = ""
text5.text = ""
txtQTY.text = ""
txtRate.text = ""
txtDis.text = ""


cmdOk.Enabled = False
cmdClear.Enabled = False
Command1.Enabled = True

cmbCust.ListIndex = -1

End Sub

Private Sub cmdOk_Click()
If lvMain.ListItems.Count > 0 Then
cmdClear.Enabled = False
cmdOk.Enabled = False
Command1.Enabled = True
picTop.Enabled = False
lvLook.Enabled = False
lvMain.Enabled = False

Call PurchaseReg
    MsgBox "Record has been Saved", vbInformation
    
Else
    MsgBox "There is no Product to record", vbInformation
End If



End Sub





'Private Sub Combo5_Click()
'On Error Resume Next
'
'Call GetNewConnection2
'
'Set Rs1 = New Recordset
'Set Rs1 = DCON.Execute("Select * from PurchaseRegistryDetail where PurchaseRegistryID='" & Combo5.text & "'")
'
'While Not Rs1.EOF
'
'    Dim LVITEM As ListItem
'    Dim ProdID As String
'
'    With lvMain
'    .ListItems.clear
'
'    ProdID = Rs1!productid
'    Set LVITEM = .ListItems.Add(, , Rs1!productid)
'        LVITEM.SubItems(2) = Rs1!Quantity
'        LVITEM.SubItems(4) = Rs1!discount
'        LVITEM.SubItems(3) = Rs1!Rate
'
'    Set RS2 = New Recordset
'    Set RS2 = DCON.Execute("SElect * from product where productID='" & ProdID & "'")
'        LVITEM.SubItems(1) = RS2!Name
'
'
'    End With
'
'    Rs1.MoveNext
'Wend
'
'
'
'Set Rs1 = Nothing
'Set RS2 = Nothing
'Set DCON = Nothing
'End Sub


Private Sub Command1_Click()
Dim CDb As CDbase
Dim CIns As New CInsert


Call GetNewConnection(CIns)
Set CDb = CIns



PurchID = CIns.AUTONUM(CDb.OpenDb, "PurchaseRegistryHeader", "PurchaseRegistryID", "PR", Label1)


Set CIns = Nothing

cmbSupp.ListIndex = -1
cmdClear.Enabled = True
lvMain.ListItems.clear
cmdOk.Enabled = True
Command1.Enabled = False
picTop.Enabled = True
lvLook.Enabled = True
lvMain.Enabled = True

Text4.text = ""
text5.text = ""
txtQTY.text = ""
txtRate.text = ""
txtDis.text = ""
cmbCust.SetFocus


End Sub

Private Sub DTPICKER1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub Form_Load()



PurchOrder = True
PurchReg = False
PurchRet = False
Timer2.Enabled = True
Timer2.Interval = 100
 '   cmbCust.AddItem "Cash"
    
    
    Call CMB1("Supplier", "BusinessName", cmbCust)
cmbCust.ListIndex = -1
    
End Sub

Private Sub Form_Resize()
'picBody.Height = Me.ScaleHeight - picTop.Height - picBottom

End Sub

Private Sub Form_Unload(Cancel As Integer)

       'frmMAIN.WindowState = 0
End Sub


Private Sub LvHeads()
    lvMain.ColumnHeaders(1).Width = lvMain.Width * 0.1
    lvMain.ColumnHeaders(2).Width = lvMain.Width * 0.2
    lvMain.ColumnHeaders(3).Width = lvMain.Width * 0.2

End Sub


Private Sub ReOrder()
Call GetNewConnection2

With lvLook

SQL = "Select * from Product where UnitsInStock <= ReOrderLevel"

Set Rs1 = New Recordset
Set Rs1 = DCON.Execute(SQL)

While Not Rs1.EOF

 .ListItems.Add , , Rs1!Name
 


Rs1.MoveNext
Wend






End With

End Sub


Private Function GetProduct(ProdID As String) As String


Call GetNewConnection2

Set Rs1 = New Recordset
Set Rs1 = DCON.Execute("Select name from product where productid='" & ProdID & "'")

If Not Rs1.EOF Then

GetProduct = Rs1!Name

Else
    MsgBox "PRODUCT NOT FOUND"
    Exit Function
    
End If


Set Rs1 = Nothing
Set DCON = Nothing



End Function


Private Sub PurchaseReg()

Dim CNT1 As Integer

With lvMain
If .ListItems.Count > 0 Then
Call GetNewConnection2
   
   ' If CRED = True Then
  '' If cmbCust.text <> "Cash" Then
   
        'PurchaseRegistryHeader
        'PurchaseRegistryDetail
        
        SQL = "Insert into PurchaseRegistryHeader values('" & PurchID & "','" _
                                                  & cmbSupp & "','" _
                                                  & SuppID & "','" & CStr(DTPicker4.Value) & "')"
                        
       ' CRED = False

      
        DCON.Execute SQL
     
   '' Else
     '  '     SQL = "Insert into PurchaseRegistryHeader values('" & PurchID & "','" _
                                                  & "Cash" & "'," _
                                                  & 1 & ",'" & CStr(DTPicker4.Value) & "')"
                        
                    
    
     '  ' DCON.Execute SQL
      
        
  ' ' End If

        For CNT1 = 1 To .ListItems.Count
            
               Dim discount As String
discount = Replace(.ListItems(CNT1).SubItems(3), "%", 1, Len(.ListItems(CNT1).SubItems(3)), 1, vbTextCompare)
               
     
        SQL = "Insert into PurchaseRegistryDetail values('" & PurchID & "','" _
        & .ListItems(CNT1).text & "'," & .ListItems(CNT1).SubItems(2) & "," & discount & "," & .ListItems(CNT1).SubItems(3) & ")"
        DCON.Execute SQL

    Set Rs1 = New Recordset

        SQL = "Select * from Product where productid='" & .ListItems(CNT1).text & "'"

       Set Rs1 = DCON.Execute(SQL)

       SQL = "update Product set UnitsInStock=" & Val(Val(Rs1!UnitsInStock) + Val(.ListItems(CNT1).SubItems(2))) _
                    & " WHERE ProductID='" & .ListItems(CNT1).text & "'"

       
       DCON.Execute SQL
       
           If cmbSupp.text = "<ALL>" Then
      

               SQL = "Update v set deliver=1 where businessname='" & cmbCust.text & "' AND Productid='" & .ListItems(CNT1).text & "'"
               
                DCON.Execute SQL
      Else
              SQL = "Update v set deliver=1 where purchaseorderid='" & cmbSupp.text & "' AND Productid='" & .ListItems(CNT1).text & "'"
                DCON.Execute SQL
        
       End If
       
        
    Next
        
        
    
  
       
        
Set Rs1 = Nothing
Set DCON = Nothing

End If

End With

End Sub


Private Sub lvLook_DblClick()
'Dim LVindex As Integer
'
'Dim LSTITEM As ListItem
'
'
'
'Dim CNT As Boolean
'Dim DD As Integer
'
'''' query in quantity is not yet included
'
'
'
'    CNT = False
'
'
'    Call GetNewConnection2
'        Set Rs1 = New Recordset
'        Set Rs1 = DCON.Execute("Select * from Product where ProductID='" & lvLook.SelectedItem.text & "'")
'
'If Not Rs1.EOF Then
'
'
'
'       'LBL_DES.Caption = RS1!ProductID & ", " & RS1!Name & ""
'       txtRate = Rs1!UnitCostPrice
'     '  txtQTY.text = RS1!ReOrderQuantity
'
'  With lvMain
'        TextAmount.text = ""
'
'        If .ListItems.Count <> 0 Then
'
'            For DD = 1 To .ListItems.Count
'
'              If InStr(1, .ListItems(DD).text, Rs1!ProductID) = 1 Then
'                If InStr(1, .ListItems(DD).SubItems(1), Rs1!Name) = 1 Then
'
'                        If EDT = True Then
'                            .ListItems(DD).Selected = True
'                            .ListItems(DD).SubItems(2) = Val(txtQTY.text)
'                            .ListItems(DD).SubItems(3) = Val(txtRate.text)
'                            .ListItems(DD).SubItems(5) = text5.text
'                        Else
'                            .ListItems(DD).Selected = True
'                            .ListItems(DD).SubItems(2) = Val(.ListItems(DD).SubItems(2)) + Val(txtQTY.text)
'                            .ListItems(DD).SubItems(3) = Val(txtRate.text)
'                            .ListItems(DD).SubItems(5) = Val(.ListItems(DD).SubItems(2)) * Val(.ListItems(DD).SubItems(3))
'
'                        End If
'
'                    CNT = True
'
'                End If
'
'               End If
'
'            Next
'
'         End If
'
'        If CNT = False Then
'
'         .ListItems.Add , , Rs1!ProductID
'            .ListItems(.ListItems.Count).SubItems(1) = Rs1!Name
'            .ListItems(.ListItems.Count).SubItems(2) = txtQTY.text
'            .ListItems(.ListItems.Count).SubItems(3) = txtRate.text
'            .ListItems(.ListItems.Count).SubItems(4) = "1"
'            .ListItems(.ListItems.Count).SubItems(5) = text5.text
'
'
'             ' TextAmount.Text = Val(TextAmount.Text) + Val(TXT_AMT.Text)
'
'        End If
'
'
'
'
''        If .ListItems.Count <= 0 Then
''
''            .ListItems.Add 1, , RS1!ProductID
''            .ListItems(.ListItems.Count).SubItems(1) = RS1!Name
''            .ListItems(.ListItems.Count).SubItems(2) = txtqty.text
''            .ListItems(.ListItems.Count).SubItems(3) = txtrate.text
''            .ListItems(.ListItems.Count).SubItems(4) = "1"
''            .ListItems(.ListItems.Count).SubItems(5) = Text5.text
''
''
''        End If
'
'
'            For DD = 1 To .ListItems.Count
'                  TextAmount.text = Val(.ListItems(DD).SubItems(5)) + Val(TextAmount.text)
'            Next
'
'        ' lblunit.Caption = RS1!UnitsInStock
'
'
'       Set Rs1 = DCON.Execute("Select * from Product")
'      '  Set DataGrid1.DataSource = RS1
'
'        Set Rs1 = Nothing
'        Set DCON = Nothing
'
'  End With
'
'Else
'    MsgBox "Product Not Found", vbInformation, "Product"
'
'
'End If
End Sub

Private Sub lvLook_ItemClick(ByVal Item As MSComctlLib.ListItem)
Timer2.Enabled = False
Call GetNewConnection2
Set Rs1 = New Recordset
Set Rs1 = DCON.Execute("Select * from product where productid='" & lvLook.SelectedItem.text & "'")

    If Not Rs1.EOF Then
        txtRate.text = Rs1!UnitCostPrice
    End If
    

If lvLook.ListItems.Count > 0 Then
    Text4.text = lvLook.SelectedItem.text
    
End If

End Sub

Private Sub lvLook_LostFocus()
Timer2.Enabled = True
End Sub

Private Sub lvLook_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    PopupMenu mnuLook
End If
End Sub

Private Sub lvMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
If PurchRet = True Then
    Timer2.Enabled = False
    

Else

End If
EDT = True
If lvMain.ListItems.Count <> 0 Then
      Text4.text = lvMain.ListItems(lvMain.SelectedItem.Index).text
    txtQTY.text = lvMain.ListItems(lvMain.SelectedItem.Index).SubItems(2)
    txtRate.text = lvMain.ListItems(lvMain.SelectedItem.Index).SubItems(3)
    txtDis = lvMain.ListItems(lvMain.SelectedItem.Index).SubItems(4)
    text5.text = lvMain.ListItems(lvMain.SelectedItem.Index).SubItems(5)
    
        
End If




End Sub

Private Sub lvMain_KeyDown(KeyCode As Integer, Shift As Integer)
Dim DD As Integer

With lvMain
If .ListItems.Count <> 0 Then
If KeyCode = vbKeyDelete Then
TextAmount.text = ""
       
  .ListItems.Remove (.SelectedItem.Index)
  
Text4.text = ""

  'lblcat.Caption = ""
EDT = False
txtRate.text = ""
text5.text = ""
txtQTY.text = ""
'lblselling.Caption = ""
'lblunit.Caption = ""

            For DD = 1 To .ListItems.Count
                  TextAmount.text = Val(.ListItems(DD).SubItems(5)) + Val(TextAmount.text)
            Next
 End If
End If
End With

End Sub

Private Sub mnuUnit_Click()
Call GetNewConnection2
Set Rs1 = New Recordset
Set Rs1 = DCON.Execute("Select * from Product where ProductID='" & lvLook.SelectedItem.text & "'")

If Rs1.RecordCount <> 0 Then
    MsgBox "Available Stock: " & Rs1!UnitsInStock
End If
Set Rs1 = Nothing
Set DCON = Nothing
End Sub

Private Sub Text4_Change()
If Len(Text4.text) = 0 Then
    Timer2.Enabled = True
End If

'If EDT = False Then
Timer2.Interval = 100


TXTLEN = Len(Text4.text)
STRT = 0

'End If
'EDT = True
txtQTY.text = ""

If Len(Text4.text) <> 0 Then
    txtQTY.Enabled = True
Else
    txtQTY.Enabled = False
    
End If
End Sub

Private Sub TextAmount_Change()
lblwords.Caption = NumToWord(TextAmount.text)
End Sub





Private Sub Timer2_Timer()

'Static c As Integer


STRT = STRT + 1



If STRT = 3 Then
Timer2.Interval = 0
  




Dim FVAL As String
Dim DD As Integer
Dim LISTITM As ListItem

Call GetNewConnection2

Set Rs1 = New Recordset

'SQL = "Select TOP 10 * from PRODUCT where PRODUCTID like '" & Text4 & "%' OR NAME like'" & Text4 & "%'"
SQL = "Select TOP 10 *,(UnitsInStock + UnitsInOrder) as Total from PRODUCT where (PRODUCTID='" & Text4 & "' OR NAME like'" & Text4 & "%')"
'SQL = "Select Top 20 * from lowINstock order by Total"
Set Rs1 = DCON.Execute(SQL)
 Set RS2 = New Recordset
        Set RS2 = DCON.Execute(SQL)
        lvLook.ListItems.clear
        While Not RS2.EOF
        
            Set LISTITM = lvLook.ListItems.Add(, , RS2!ProductID)
            
                LISTITM.SubItems(1) = RS2!Name
                
                If RS2!UnitsInStock <= 0 Then
                    LISTITM.ForeColor = vbRed
                    LISTITM.ListSubItems(1).ForeColor = vbRed
                Else
                    If RS2!UnitsInStock <= RS2!ReorderLevel Then
                    LISTITM.ForeColor = vbBlue
                    LISTITM.ListSubItems(1).ForeColor = vbBlue
                    End If
                End If
        
            RS2.MoveNext
        Wend

If Text4.text <> "" Then

    If Not Rs1.EOF Then
'        TXT_CODE.SelStart = PRILEN
'        TXT_CODE.text = RS1!Name
'        TXT_CODE.SelLength = Len(TXT_CODE.text)
'
      
        FVAL = Rs1!ProductID
        
      
       txtRate.text = Rs1!UnitCostPrice
       
        
       text5.text = Val(txtQTY.text) * Val(txtRate.text)
        
        'lblselling.Caption = RS1!UnitSellingPrice
        'lblunit.Caption = RS1!UnitsInStock

  With lvMain
        If .ListItems.Count <> 0 Then
          
            For DD = 1 To .ListItems.Count
                  
                If InStr(1, .ListItems(DD).SubItems(1), Rs1!ProductID) = 1 Then
                  
                 
                            .ListItems(DD).Selected = True
                         '   lblunit.Caption = Val(lblunit.Caption) - Val(.ListItems(DD).SubItems(3))
                    
                    
                End If
            
               
            Next
         End If
    End With

  

    Else
      txtRate.text = ""
        text5.text = ""
'        TXT_QTY.text = ""
'        lblselling.Caption = ""
'        lblunit.Caption = ""
'        lblcat.Caption = ""
'        PRILEN = 0

    End If

    
    Set Rs1 = Nothing
    Set RS2 = Nothing
    Set DCON = Nothing


ElseIf Text4.text = "" Then
   txtRate.text = ""
        text5.text = ""
        txtQTY.text = ""
'lblcat.Caption = ""
'EDT = False
'TXT_RATE.text = ""
'TXT_AMT.text = ""
'TXT_QTY.text = ""
'lblselling.Caption = ""
'lblunit.Caption = ""
'TXT_CODE.SetFocus
End If

End If



End Sub

Private Sub txtqty_Change()
text5.text = Val(txtQTY.text) * Val(txtRate.text)

If Len(txtQTY.text) <> 0 Then
    txtRate.Enabled = True
Else
    txtRate.Enabled = False
    
End If

If Len(txtRate.text) <> 0 And Val(txtRate.text) <> 0 And Val(txtQTY.text) <> 0 Then
    cmdAdd.Enabled = True
Else
    cmdAdd.Enabled = False
    
End If
End Sub

Private Sub txtqty_KeyPress(KeyAscii As Integer)
Call OFFCHar(KeyAscii, txtQTY)

End Sub

Private Sub txtRate_Change()
text5.text = Val(txtQTY.text) * Val(txtRate.text)
If Len(txtRate.text) <> 0 And Val(txtRate.text) <> 0 And Val(txtQTY.text) <> 0 Then
    cmdAdd.Enabled = True
Else
    cmdAdd.Enabled = False
    
End If
If Len(txtRate.text) <> 0 Then
    txtDis.Enabled = True
Else
    txtDis.Enabled = False
    
End If
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)

Call Decimals(KeyAscii, txtRate, 2)

If KeyAscii = 13 Then

    Call GetNewConnection2

        Set Rs1 = New Recordset
            SQL = "Select * from PRODUCT where PRODUCTID='" & Text4 & "' OR NAME='" & Text4 & "'"

'        Set RS1 = DCON.Execute(SQL)
'
'            SQL = "Select * from Product where (PRODUCTID='" & RS1!productid & "' AND UnitCostPrice <" & Val(txtrate.text) & ")"
'
            Set Rs1 = DCON.Execute(SQL)
'
          If Rs1.RecordCount <> 0 Then
                                     If Val(Rs1!UnitSellingPrice) < Val(txtRate.text) Then
                    
                            If MsgBox("The unit cost price has increase, do you want to update " & vbCrLf _
                            & "unit selling price?   ", vbQuestion + vbYesNo) = vbYes Then
                                Dim SelPrice As String
                                     SelPrice = InputBox("Input New Unit Selling Price for this Product:", , Val(txtRate.text) * 1.1)
                                
                                While IsNumeric(SelPrice) = False Or Val(txtRate.text) > Val(SelPrice)
                                    SelPrice = InputBox("Input New Unit Selling Price for this Product:", , Val(txtRate.text) * 1.1)
                                          
                                Wend
                                  SQL = "UPDATE PRODUCT set UnitSellingPrice=" & Val(SelPrice) & " where (PRODUCTID='" & Rs1!ProductID & "')"
   
                                            DCON.Execute SQL
                             Else
                                txtRate.SetFocus
                                
                           End If
                    End If
                            
'                SQL = "UPDATE PRODUCT set UnitCostPrice=" & Val(txtrate.text) & " where (PRODUCTID='" & RS1!productid & "' AND UnitCostPrice <" & Val(txtrate.text) & ")"
                    SQL = "UPDATE PRODUCT set UnitCostPrice=" & Val(txtRate.text) & " where PRODUCTID='" & Rs1!ProductID & "'"
              
                
                DCON.Execute SQL
'            Else
                
'                SQL = "Select * from PRODUCT where PRODUCTID like '" & Text4 & "%' OR NAME like'" & Text4 & "%'"
'                Set RS1 = DCON.Execute(SQL)
'
'                      If RS1!unitcostprice <> txtrate.text Then
'
'                         MsgBox "Cannot update UnitCostPrice" & vbTab, vbInformation, "UnitCostPrice"
'
'                      End If
'
'                      If RS1.RecordCount <> 0 Then
'                         txtrate.text = RS1!unitcostprice
'                         txtrate.SetFocus
'                      End If
                   
                    
            End If
         

Set Rs1 = Nothing
Set DCON = Nothing



End If
End Sub

Private Sub txtrate_LostFocus()
  
    Call GetNewConnection2

        Set Rs1 = New Recordset
            SQL = "Select * from PRODUCT where PRODUCTID='" & Text4 & "' OR NAME='" & Text4 & "'"

'        Set RS1 = DCON.Execute(SQL)
'
'            SQL = "Select * from Product where (PRODUCTID='" & RS1!productid & "' AND UnitCostPrice <" & Val(txtrate.text) & ")"
'
        Set Rs1 = DCON.Execute(SQL)
'
          If Rs1.RecordCount <> 0 Then
                       If Val(Rs1!UnitSellingPrice) < Val(txtRate.text) Then
                    
                            If MsgBox("The unit cost price has increase, do you want to update " & vbCrLf _
                            & "unit selling price?   ", vbQuestion + vbYesNo) = vbYes Then
                                Dim SelPrice As String
                                     SelPrice = InputBox("Input New Unit Selling Price for this Product:", , Val(txtRate.text) * 1.1)
                                
                                While IsNumeric(SelPrice) = False Or Val(txtRate.text) > Val(SelPrice)
                                    SelPrice = InputBox("Input New Unit Selling Price for this Product:", , Val(txtRate.text) * 1.1)
                                          
                                Wend
                                  SQL = "UPDATE PRODUCT set UnitSellingPrice=" & Val(SelPrice) & " where (PRODUCTID='" & Rs1!ProductID & "')"
   
                                            DCON.Execute SQL
                             Else
                                txtRate.SetFocus
                                
                           End If
                    End If
'                SQL = "UPDATE PRODUCT set UnitCostPrice=" & Val(txtrate.text) & " where (PRODUCTID='" & RS1!productid & "' AND UnitCostPrice <" & Val(txtrate.text) & ")"
               SQL = "UPDATE PRODUCT set UnitCostPrice=" & Val(txtRate.text) & " where PRODUCTID='" & Rs1!ProductID & "'"
              
                
                DCON.Execute SQL
'            Else
                
'                SQL = "Select * from PRODUCT where PRODUCTID like '" & Text4 & "%' OR NAME like'" & Text4 & "%'"
'                Set RS1 = DCON.Execute(SQL)
'
'                      If RS1!unitcostprice <> txtrate.text Then
'
'                         MsgBox "Cannot update UnitCostPrice" & vbTab, vbInformation, "UnitCostPrice"
'
'                      End If
'
'                      If RS1.RecordCount <> 0 Then
'                         txtrate.text = RS1!unitcostprice
'                         txtrate.SetFocus
'                      End If
                   
                    
            End If
         

Set Rs1 = Nothing
Set DCON = Nothing


End Sub


