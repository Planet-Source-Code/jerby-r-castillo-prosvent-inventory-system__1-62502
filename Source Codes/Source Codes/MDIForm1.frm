VERSION 5.00
Begin VB.MDIForm frmMAIN 
   BackColor       =   &H8000000C&
   Caption         =   "ProsVent Inventory Manager 2005"
   ClientHeight    =   7680
   ClientLeft      =   165
   ClientTop       =   -60
   ClientWidth     =   11010
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H8000000D&
      FillColor       =   &H00E0E0E0&
      Height          =   7080
      Left            =   0
      ScaleHeight     =   7020
      ScaleWidth      =   2235
      TabIndex        =   0
      Top             =   600
      Width           =   2295
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   6960
         Left            =   60
         ScaleHeight     =   6960
         ScaleWidth      =   2070
         TabIndex        =   3
         Top             =   75
         Width           =   2070
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H80000013&
            BackStyle       =   0  'Transparent
            Caption         =   "Transaction"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   240
            Left            =   255
            TabIndex        =   15
            ToolTipText     =   "Choose Tasks"
            Top             =   2220
            Width           =   1155
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Location"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Index           =   3
            Left            =   255
            MouseIcon       =   "MDIForm1.frx":0000
            MousePointer    =   99  'Custom
            TabIndex        =   14
            Top             =   1800
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Purchase Orders"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Index           =   4
            Left            =   255
            MouseIcon       =   "MDIForm1.frx":030A
            MousePointer    =   99  'Custom
            TabIndex        =   13
            Top             =   2640
            Width           =   1575
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Index           =   0
            Left            =   255
            MouseIcon       =   "MDIForm1.frx":0614
            MousePointer    =   99  'Custom
            TabIndex        =   12
            Top             =   720
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Suppliers"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Index           =   1
            Left            =   255
            MouseIcon       =   "MDIForm1.frx":091E
            MousePointer    =   99  'Custom
            TabIndex        =   11
            Top             =   1080
            Width           =   870
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Category"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Index           =   2
            Left            =   255
            MouseIcon       =   "MDIForm1.frx":0C28
            MousePointer    =   99  'Custom
            TabIndex        =   10
            Top             =   1440
            Width           =   870
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Purchase Return"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Index           =   5
            Left            =   255
            MouseIcon       =   "MDIForm1.frx":0F32
            MousePointer    =   99  'Custom
            TabIndex        =   9
            Top             =   4080
            Width           =   1560
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Purchase Registry"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Index           =   6
            Left            =   255
            MouseIcon       =   "MDIForm1.frx":123C
            MousePointer    =   99  'Custom
            TabIndex        =   8
            Top             =   3000
            Width           =   1695
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Return"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Index           =   7
            Left            =   255
            MouseIcon       =   "MDIForm1.frx":1546
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   3720
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Registry"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Index           =   8
            Left            =   255
            MouseIcon       =   "MDIForm1.frx":1850
            MousePointer    =   99  'Custom
            TabIndex        =   6
            Top             =   3375
            Width           =   1335
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H80000013&
            BackStyle       =   0  'Transparent
            Caption         =   "Catalogue"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   240
            Left            =   255
            TabIndex        =   5
            ToolTipText     =   "Choose Tasks"
            Top             =   150
            Width           =   1005
         End
         Begin VB.Label lblTask 
            BackStyle       =   0  'Transparent
            Caption         =   "Home"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   240
            Left            =   165
            MouseIcon       =   "MDIForm1.frx":1B5A
            MousePointer    =   99  'Custom
            TabIndex        =   4
            Top             =   5040
            Width           =   1710
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00C7BDAD&
            BorderStyle     =   6  'Inside Solid
            FillColor       =   &H00E0E0E0&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   135
            Shape           =   4  'Rounded Rectangle
            Top             =   120
            Width           =   1845
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00C7BDAD&
            BorderStyle     =   6  'Inside Solid
            FillColor       =   &H00E0E0E0&
            FillStyle       =   0  'Solid
            Height          =   315
            Left            =   135
            Top             =   660
            Visible         =   0   'False
            Width           =   1845
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H00C7BDAD&
            BorderStyle     =   6  'Inside Solid
            FillColor       =   &H00E0E0E0&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   135
            Shape           =   4  'Rounded Rectangle
            Top             =   2160
            Width           =   1845
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00C7BDAD&
            BorderStyle     =   6  'Inside Solid
            FillColor       =   &H00E0E0E0&
            FillStyle       =   0  'Solid
            Height          =   1260
            Left            =   120
            Top             =   4920
            Width           =   1845
         End
      End
   End
   Begin VB.PictureBox iml16 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   10980
      TabIndex        =   1
      Top             =   0
      Width           =   11010
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "ProsVent Inventory Manager 2005"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   480
         Left            =   1800
         TabIndex        =   2
         ToolTipText     =   "Choose Tasks"
         Top             =   0
         Width           =   1965
      End
      Begin VB.Image Image1 
         Height          =   630
         Left            =   15
         Picture         =   "MDIForm1.frx":1E64
         Top             =   -45
         Width           =   4500
      End
   End
   Begin VB.Menu mnuTop 
      Caption         =   "File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New        "
         Begin VB.Menu mnuNew 
            Caption         =   "Customer...."
            Index           =   0
            Shortcut        =   {F1}
         End
         Begin VB.Menu mnuNew 
            Caption         =   "Vendor....."
            Index           =   1
            Shortcut        =   {F2}
         End
         Begin VB.Menu mnuNew 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnuNew 
            Caption         =   "Product"
            Index           =   3
            Shortcut        =   {F3}
         End
         Begin VB.Menu mnuNew 
            Caption         =   "Category"
            Index           =   4
            Shortcut        =   {F4}
         End
         Begin VB.Menu mnuNew 
            Caption         =   "Location"
            Index           =   10
            Shortcut        =   {F8}
         End
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuPageSetup 
         Caption         =   "Page SetUp"
      End
      Begin VB.Menu mnuPrintPrv 
         Caption         =   "Print Preview"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Visible         =   0   'False
      Begin VB.Menu mnuModify 
         Caption         =   "Modify         "
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete           "
      End
      Begin VB.Menu mnuDetails 
         Caption         =   "Details       "
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu spc 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "Find"
      End
   End
   Begin VB.Menu mnuTransaction 
      Caption         =   "Transaction"
      Begin VB.Menu mnuPO 
         Caption         =   "Purchase Order"
      End
      Begin VB.Menu mnuPreg 
         Caption         =   "Purchase Registry"
      End
      Begin VB.Menu mnuSReg 
         Caption         =   "Sales Registry"
      End
      Begin VB.Menu mnPret 
         Caption         =   "Purchase Return"
      End
      Begin VB.Menu mnuSRet 
         Caption         =   "Sales Return"
      End
   End
   Begin VB.Menu mnuAB 
      Caption         =   "View"
      Begin VB.Menu mnuSalesRegister 
         Caption         =   "Sales Registry"
      End
      Begin VB.Menu mnuPurchaseRegister 
         Caption         =   "Purchase Registry"
      End
      Begin VB.Menu mnuStockItem 
         Caption         =   "Stock Analysis"
      End
   End
   Begin VB.Menu mnUsers 
      Caption         =   "Users"
      Begin VB.Menu mnuAddUser 
         Caption         =   "Add User"
      End
      Begin VB.Menu mnuDeleteUser 
         Caption         =   "Delete User"
      End
      Begin VB.Menu mnuChangeUsername 
         Caption         =   "Change Username"
      End
      Begin VB.Menu mnuChangePassword 
         Caption         =   "Change Password"
      End
      Begin VB.Menu mnuViewall 
         Caption         =   "View All User"
      End
   End
   Begin VB.Menu mnuUtility 
      Caption         =   "Utility"
      Begin VB.Menu mnuBackup 
         Caption         =   "Backup"
      End
   End
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_SYSCOMMAND = &H112
Private Const SC_CLOSE = &HF060&
Public Sub size(frm As Form)
    frm.Width = Me.ScaleWidth
    frm.Height = Me.ScaleHeight
End Sub

Private Sub Command1_Click(Index As Integer)
    Dim tempSql As String
    frmControlMain.DataGrid1.Visible = True
    frmControlMain.WBrow.Visible = False
    frmControlMain.Show
End Sub



Private Sub Command2_Click()
'MODIFYID = "100"
'ADDING = False
'C_frmProduct.Show vbModal

End Sub

Private Sub Command3_Click()
Form2.Show vbModal
End Sub

Private Sub Label1_Click(Index As Integer)
    frmControlMain.WBrow.Visible = False
    frmControlMain.DataGrid1.Visible = True
    Select Case Index
    
    Case 0
         SQL = "Select * from customer WHERE CustomerID<>'CASH'"
        CatalogueName = "Customer"
    Case 1
        SQL = "Select * from Supplier WHERE SuppliersID<>'CASH'"
        CatalogueName = "Supplier"
    Case 2
     SQL = "Select * from category"
     CatalogueName = "Category"
    Case 3
        SQL = "Select * from location"
        CatalogueName = "Location"
    Case 4
'        SQL = "Select * from PurchaseOrderHeader"
        SQL = "Select * from TotalOrder"
        CatalogueName = "Purchase Order"
    Case 5
'        SQL = "Select * from PurchaseReturnHeader"
        SQL = "Select * from TotalReturn"
        CatalogueName = "Purchase Return"
    Case 6
'        SQL = "Select * from PurchaseRegistryHeader"
        SQL = "Select * from TotalPurchase"
        CatalogueName = "Purchase Registry"
    Case 7
        SQL = "Select * from SalesReturnHeader"
        CatalogueName = "Sales Return"
    Case 8
        SQL = "Select * from TotalSales"
        CatalogueName = "Sales Registry"
    End Select
    

Call GetNewConnection2
Set Rs1 = New Recordset
Set Rs1 = DCON.Execute(SQL)
If Rs1.RecordCount <= 0 Then
    frmControlMain.DataGrid1.Visible = False
Else
    Set frmControlMain.DataGrid1.DataSource = Rs1
End If

Set Rs1 = Nothing
Set DCON = Nothing

    
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call moveShape(Shape1, Label1(Index))
End Sub
Public Function moveShape(shape As Object, Cntrl As Object)
        shape.Visible = True
        shape.Move Cntrl.Left - 150, Cntrl.Top - 60, 1845, 300
End Function


Private Sub lblTask_Click()
    Load frmControlMain
    frmControlMain.DataGrid1.Visible = False
    frmControlMain.WBrow.Visible = True
    
    
    Dim SqLargs As String
    SqLargs = "SELECT Product.ProductID, Product.Name, Product.UnitsInStock, Product.UnitCostPrice From Product WHERE ((Product.UnitsInStock)<=0) Order by UnitsInStock DESC"
    Call frmControlMain.CreateStartPage(SqLargs)
End Sub

Private Sub MDIForm_Resize()
    Call size(frmControlMain)
End Sub
Private Sub mnuCashBook_Click()
    Form1.Show
End Sub
Private Sub mnuDaybook_Click()
    GetNewConnection2
   Call frmControlMain.CreateDataPage("Select  TOP 2 * From V1", "Day book")
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
End
End Sub

Private Sub mnPret_Click()
'ADDING = True
Load tempPRET
tempPRET.Show vbModal

End Sub

Private Sub mnuAddUser_Click()
    Load frmAddUser
    frmAddUser.Show vbModal
End Sub

Private Sub mnuBackup_Click()
Load frmBackup
frmBackup.Show vbModal

End Sub

Private Sub mnuChangePassword_Click()
Load frmChange
frmChange.Show vbModal

End Sub

Private Sub mnuChangeUsername_Click()
Load frmuser
frmuser.Show vbModal

End Sub

Private Sub mnuDelete_Click()
Call GetNewConnection2
Set Rs1 = New Recordset
Select Case CatalogueName



  Case "Customer"

        Set Rs1 = DCON.Execute("Select * From CustExist where CustomerID='" & frmControlMain.DataGrid1.Columns(0).text & "'")
        
        If Rs1.RecordCount = 0 Then
          If MsgBox("Do You Want to Delete the Information of Customer", vbInformation + vbYesNo) = vbYes Then
          
            SQL = "Delete From Customer Where CustomerId='" & frmControlMain.DataGrid1.Columns(0).text & "'"
            End If
        Else
            MsgBox "Transaction Exist, cannot delete request", vbInformation
        End If
       
  Case "Supplier"
        Set Rs1 = DCON.Execute("Select DISTINCT SupplierID From SuppExist where SupplierID='" & frmControlMain.DataGrid1.Columns(0).text & "'")
        If Rs1.RecordCount = 0 Then
        If MsgBox("Do You Want to Delete the Information of Customer", vbInformation + vbYesNo) = vbYes Then
            SQL = "Delete From Supplier Where SuppliersId='" & frmControlMain.DataGrid1.Columns(0).text & "'"
         End If
        Else
            MsgBox "Transaction Exist, cannot delete request", vbInformation
          
        End If
        
  Case "Category"
      Set Rs1 = DCON.Execute("Select * From CatExist where Category='" & frmControlMain.DataGrid1.Columns(0).text & "'")
        If Rs1.RecordCount = 0 Then
        If MsgBox("Do You Want to Delete the Information of Customer", vbInformation + vbYesNo) = vbYes Then
        SQL = "Delete from Category Where Category='" & frmControlMain.DataGrid1.Columns(0).text & "'"
        End If
       Else
            MsgBox "Transaction Exist, cannot delete request", vbInformation
           
        End If
        
  Case "Location"
    If MsgBox("Do You Want to Delete the Information of Customer", vbInformation + vbYesNo) = vbYes Then
        SQL = "Delete from Location Where location='" & frmControlMain.DataGrid1.Columns(0).text & "'"
    End If
  Case "Purchase Order"
     MsgBox "Transaction Exist, cannot delete request", vbInformation
  Case "Purchase Return"
     MsgBox "Transaction Exist, cannot delete request", vbInformation
         
        'Delete from PurchaseReturnDetail Where PurchaseReturnID='" & "'"
        'Delete from PurchaseOrderHeader Where PurchaseOrderID='" & "'"
  Case "Purchase Registry"
     MsgBox "Transaction Exist, cannot delete request", vbInformation
         
       'Delete from PurchaseOrderDetail Where PurchaseOrderID='" & "'"
       'Delete from PurchaseOrderHeader Where PurchaseOrderID='" & "'"
  Case "Sales Return"
     MsgBox "Transaction Exist, cannot delete request", vbInformation
         
       'Delete from PurchaseOrderDetail Where PurchaseOrderID='" & "'"
        'Delete from PurchaseOrderHeader Where PurchaseOrderID='" & "'"
  Case "Sales Registry"
     MsgBox "Transaction Exist, cannot delete request", vbInformation
         
        'Delete from SalesRegistryDetail Where SalesRegistryID='" & "'"
        'Delete from SalesRegistryHeader Where SalesRegistryID='" & "'"
End Select

DCON.Execute SQL

Call GridRefresh


Set DCON = Nothing
End Sub

Private Sub mnuDeleteUser_Click()
SQL = "Select username1 from users where username1 <> '" & CurUser & "'"
Load frmDelUser
frmDelUser.Show vbModal

End Sub

Private Sub mnuDetails_Click()
On Error GoTo adder:
If frmControlMain.DataGrid1.VisibleRows >= 1 Then
Select Case (CatalogueName)
  Case "Purchase Order"
    CreateH_Page "Select [Product/Category]," _
                    & "PurchaseOrderDetail.Quantity," _
                    & "PurchaseOrderDetail.Rate  from DprodOrder" _
                    & " where PurchaseOrderID='" _
                    & frmControlMain.DataGrid1.Columns(0).text & "'", "" _
                    & " Total Amount " & frmControlMain.DataGrid1.Columns(3).text
  Case "Purchase Return"
        CreateH_Page "Select * from PurchaseReturnDetail where PurchaseReturnID='" & frmControlMain.DataGrid1.Columns(0).text & "'", " Details "
  Case "Purchase Registry"
        CreateH_Page "Select * from PurchaseRegistryDetail where PurchaseRegistryID='" & frmControlMain.DataGrid1.Columns(0).text & "'", " Details "
   Case "Sales Return"
        CreateH_Page "Select * from SalesReturnDetail where SalesReturnID='" & frmControlMain.DataGrid1.Columns(0).text & "'", " Details "
   Case "Sales Registry"
        CreateH_Page "Select * from SalesRegistryDetail where SalesRegistryID='" & frmControlMain.DataGrid1.Columns(0).text & "'", " Details "
   Case "Category"
        CreateH_Page "select * from Category", "Category"
  Case Else
           MakeShortReport SQL, " Details "
End Select
End If
Exit Sub
adder:
End Sub

Private Sub mnuExit_Click()
On Error GoTo adder
   PostMessage Me.hWnd, WM_SYSCOMMAND, SC_CLOSE, 0
   
'little bit of blabla this method is very widely used in Scripting Lanuguage very powerful code but very gentle
adder:
Exit Sub
End Sub

Private Sub mnuFind_Click()
Dim sFind As String

sFind = InputBox("Find a record", "Record")
sFind = Replace(sFind, "'", "", 1, Len(sFind), vbTextCompare)


If sFind <> "" Then
  Select Case CatalogueName
  
  Case "Customer"
         Call GRIDBIND("Customer", frmControlMain.DataGrid1, " Where customerid like'" & sFind & "%' OR Company like'" & sFind & "%'")

  Case "Supplier"
         Call GRIDBIND("Supplier", frmControlMain.DataGrid1, " Where suppliersid like'" & sFind & "%' Or BusinessName like'" & sFind & "%'")

  Case "Category"
         Call GRIDBIND("Category", frmControlMain.DataGrid1, " Where category like'" & sFind & "%'")
  Case "Location"
         Call GRIDBIND("location", frmControlMain.DataGrid1, " Where location like'" & sFind & "%'")
  Case "Purchase Order"
        Call GRIDBIND("PurchaseOrderHeader", frmControlMain.DataGrid1, " Where PurchaseOrderID like'" & sFind & "%'")

  Case "Purchase Return"
            Call GRIDBIND("PurchaseReturnHeader", frmControlMain.DataGrid1, " Where PurchaseReturnID like'" & sFind & "%'")
  Case "Purchase Registry"
      Call GRIDBIND("PurchaseRegistryHeader", frmControlMain.DataGrid1, " Where PurchaseRegistryID like'" & sFind & "%'")
  Case "Sales Return"
            Call GRIDBIND("SalesReturnHeader", frmControlMain.DataGrid1, " Where SalesReturnID like'" & sFind & "%'")
  Case "Sales Registry"
             Call GRIDBIND("SalesRegistryHeader", frmControlMain.DataGrid1, " Where SalesRegistryID like'" & sFind & "%'")
    
End Select

End If

End Sub

Private Sub mnuModify_Click()
ADDING = False
MODIFYID = frmControlMain.DataGrid1.Columns(0).text


Select Case CatalogueName


  Case "Customer"
    Load C_frmCustomer
    C_frmCustomer.Show vbModal
  Case "Supplier"
    Load C_frmSupplier
    C_frmSupplier.Show vbModal
  Case "Category"
    Load C_frmCategory
    C_frmCategory.Show vbModal
  Case "Location"
    Load C_frmLocation
    C_frmLocation.Show vbModal
 
  Case "Purchase Order"
    
     MsgBox "Transaction has Already been Recieved:" & vbCrLf & "Please Use Return Modules For Returning.", vbInformation
  
   ' tempPO.Show vbModal
   

''
''       SQL = "Delete from PurchaseOrderDetail Where PurchaseOrderID='" & "'"
''       SQL = "Delete from PurchaseOrderHeader Where PurchaseOrderID='" & "'"
 Case "Purchase Return"
   ' tempPRET.Show vbModal
    MsgBox "Transaction has Already been Recieved:" & vbCrLf & "Please Use Return Modules For Returning.", vbInformation
  
'        'Delete from PurchaseReturnDetail Where PurchaseReturnID='" & "'"
'        'Delete from PurchaseOrderHeader Where PurchaseOrderID='" & "'"
  Case "Purchase Registry"
        MsgBox "Transaction has Already been Recieved:" & vbCrLf & "Please Use Return Modules For Returning.", vbInformation
  
'       'Delete from PurchaseOrderDetail Where PurchaseOrderID='" & "'"
'       'Delete from PurchaseOrderHeader Where PurchaseOrderID='" & "'"
 Case "Sales Return"
 MsgBox "Transaction has Already been Recieved:" & vbCrLf & "Please Use Return Modules For Returning.", vbInformation
  
'       'Delete from PurchaseOrderDetail Where PurchaseOrderID='" & "'"
'        'Delete from PurchaseOrderHeader Where PurchaseOrderID='" & "'"
 Case "Sales Registry"
 MsgBox "Transaction has Already been Recieved:" & vbCrLf & "Please Use Return Modules For Returning.", vbInformation
 
  
'        'Delete from SalesRegistryDetail Where SalesRegistryID='" & "'"
        'Delete from SalesRegistryHeader Where SalesRegistryID='" & "'"
End Select
    


    
End Sub


Private Sub mnuPO_Click()
'ADDING = True
Load tempPO
tempPO.Show vbModal

End Sub

Private Sub mnuPreg_Click()
' ADDING = True
Load frmPR
frmPR.Show vbModal

End Sub

Private Sub mnuPrint_Click()
On Error GoTo adder:
    If frmControlMain.WBrow.Visible = True Then
        frmControlMain.WBrow.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER
    End If
Exit Sub
adder:
    Exit Sub
End Sub

Private Sub mnuPrintPrv_Click()
On Error GoTo adder:
    If frmControlMain.WBrow.Visible = True Then
        frmControlMain.WBrow.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER
    End If
Exit Sub
adder:
    Exit Sub
End Sub
Private Sub mnuPageSetup_Click()
On Error GoTo adder:
If frmControlMain.WBrow.Visible = True Then
        frmControlMain.WBrow.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_PROMPTUSER
End If
    Exit Sub
adder:
    Exit Sub
End Sub

Private Sub mnuPurchaseRegister_Click()
 frmControlMain.WBrow.Visible = True
    frmControlMain.DataGrid1.Visible = False
    
rptState = "PurchaseRegistry"

Load Form1
Form1.Show vbModal

End Sub

Private Sub mnuRefresh_Click()
    Call GridRefresh
End Sub

Private Sub mnuSalesRegister_Click()
 frmControlMain.WBrow.Visible = True
    frmControlMain.DataGrid1.Visible = False
rptState = "SalesRegistry"

Load Form1
Form1.Show vbModal



End Sub

Private Sub mnuSave_Click()
'    On Error GoTo adder:
'If frmControlMain.wbrow.Visible = True Then
        frmControlMain.WBrow.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_PROMPTUSER, 1
        
'End If
    'Exit Sub
'adder:
'    Exit Sub
End Sub

Private Sub mnuNew_Click(Index As Integer)
'0   customer
'1   vendor
'2   (space)
'3   product
'4   cat
'5   engmod
'6   (space)
'7   uom
'8   (space)
'9   bank
'10   Location
Select Case Index
    Case 0
    ADDING = True
    C_frmCustomer.Show vbModal
    
    Case 1
    ADDING = True
    C_frmSupplier.Show vbModal
    
    Case 3
    ADDING = True
    C_frmProduct.Show vbModal
    
    Case 4
     ADDING = True
    C_frmCategory.Show vbModal
   
    Case 10
    ADDING = True
    C_frmLocation.Show vbModal
    
End Select
End Sub



Private Sub mnuSReg_Click()
' ADDING = True
Load tempSalesReg
tempSalesReg.Show vbModal

End Sub

Private Sub mnuSRet_Click()
' ADDING = True
Load frmSalesReturn
frmSalesReturn.Show vbModal

End Sub

'Private Sub MDIForm_Load()
''setBitmaps
''smakebar
''LoadtabLeft
'
'End Sub
'Sub LoadtabLeft()
'tabLeft.Pinned = False
'tabLeft.ImageList = ImageList1
'   Dim tabX As cTab
'   Set tabX = tabLeft.Tabs.Add("Main", , "Main", 8)
'        tabX.Panel = Picture1
'    Set tabX = tabLeft.Tabs.Add("Catalogue", , "Catalogue", 2)
'        tabX.Panel = picSearch
'
''   Set tabX = tabLeft.Tabs.Add("StockINQ", , "Stock Inquiry", 3)
''   Set tabX = tabLeft.Tabs.Add("EXPLORER2", , "Explorer", 4)
''        'tabX.Panel = Picture1
''   Set tabX = tabLeft.Tabs.Add("EXPLORER3", , "Explorer", 5)
'End Sub
'
'Sub setBitmaps()
'    With PopMenu1
'    .SubClassMenu Me
'    .ImageList = ImageList1
'    .ItemIcon("mnuFileNew") = ImageList1.ListImages("C9").Index - 1
'    .ItemIcon("mnuNew(0)") = ImageList1.ListImages("C1").Index - 1
'    .ItemIcon("mnuNew(1)") = ImageList1.ListImages("C2").Index - 1
'    .ItemIcon("mnuNew(3)") = ImageList1.ListImages("C3").Index - 1
'    .ItemIcon("mnuNew(4)") = ImageList1.ListImages("C4").Index - 1
'    .ItemIcon("mnuNew(5)") = ImageList1.ListImages("C5").Index - 1
'    .ItemIcon("mnuNew(7)") = ImageList1.ListImages("C5").Index - 1
'    .ItemIcon("mnuFind") = ImageList1.ListImages("C7").Index - 1
'    .ItemIcon("mnuDelete") = ImageList1.ListImages("C6").Index - 1
'    .ItemIcon("mnuModify") = ImageList1.ListImages("C8").Index - 1
'    .ItemIcon("mnuDetails") = ImageList1.ListImages("C9").Index - 1
'    .ItemIcon("mnuDetails") = ImageList1.ListImages("C9").Index - 1
'.ItemIcon("mnuCashBook") = ImageList1.ListImages("C9").Index - 1
'.ItemIcon("mnuLedger") = ImageList1.ListImages("C9").Index - 1
'.ItemIcon("mnuGSummary") = ImageList1.ListImages("C9").Index - 1
'.ItemIcon("mnuSalesRegister") = ImageList1.ListImages("C9").Index - 1
'.ItemIcon("mnuPurchaseRegister") = ImageList1.ListImages("C9").Index - 1
'.ItemIcon("mnuDaybook") = ImageList1.ListImages("C9").Index - 1
'.ItemIcon("mnuSOA") = ImageList1.ListImages("C9").Index - 1
'    .ItemIcon("mnuRecievable") = ImageList1.ListImages("C9").Index - 1
'    .ItemIcon("mnuRecievable") = ImageList1.ListImages("C9").Index - 1
'    .ItemIcon("mnuPayables") = ImageList1.ListImages("C9").Index - 1
'    .ItemIcon("mnuSALedger") = ImageList1.ListImages("C9").Index - 1
'
'    .ItemIcon("mnuStockItem") = ImageList1.ListImages("C9").Index - 1
'    .ItemIcon("mnuGroupSummary") = ImageList1.ListImages("C9").Index - 1
'    .ItemIcon("mnuMovementAnalysis") = ImageList1.ListImages("C9").Index - 1
'    .ItemIcon("mnuPhysicalStockRegister") = ImageList1.ListImages("C9").Index - 1
'    .ItemIcon("mnuInventoryStatement") = ImageList1.ListImages("C9").Index - 1
'        .ItemIcon("mnuLocation") = ImageList1.ListImages("C9").Index - 1
'        .ItemIcon("mnuReorderStatus") = ImageList1.ListImages("C9").Index - 1
'        .ItemIcon("mnuCategories") = ImageList1.ListImages("C9").Index - 1
'        .ItemIcon("mnuPendings") = ImageList1.ListImages("C9").Index - 1
'    .UnsubclassMenu
'    End With
'End Sub
''///////////////////////////////////////////////
'Sub makebar()
'Dim barX As cListBar
'Dim itmX As cListBarItem
'Dim i As Long
'    With listbar1
'        .ImageList(evlbLargeIcon) = iml16 ' ilsIcons32
'        .ImageList(evlbSmallIcon) = iml16 'ilsIcons32
''//////////Catalogue Entries
'      Set barX = .Bars.Add("Catalogue", , "Catalogue")
'        Set itmX = barX.Items.Add("AutoCompany", , "AutoCompany", 16)
'            itmX.HelpText = "Catalogue Your Information By Categorizing record for AutoCompany"
'        Set itmX = barX.Items.Add("EngineModel", , "EngineModel", 2)
'            itmX.HelpText = "Cataglogue your Information By Assigning EngineModel"
'        Set itmX = barX.Items.Add("Location", , "Location", 3)
'            itmX.HelpText = "Add/Edit/Remove Location for Stock and Inventory"
'        Set itmX = barX.Items.Add("Category", , "Category", 4)
'            itmX.HelpText = "Add/Edit/Remove Category For easy Groupings"
'        Set itmX = barX.Items.Add("Product", , "Product", 5)
'            itmX.HelpText = "Add/Edit/Remove Products and Details"
'        Set itmX = barX.Items.Add("UOM", , "Units of Measures", 6)
'                itmX.HelpText = "Enter Units of Measures for Packaging"
''//////////Accounts
'      Set barX = .Bars.Add("Account", , "Accounts")
'            Set itmX = barX.Items.Add("Supplier", , "Supplier", 7)
'                itmX.HelpText = ""
'            Set itmX = barX.Items.Add("Customer", , "Customer", 8)
'                itmX.HelpText = ""
'            Set itmX = barX.Items.Add("Bank", , "Banks", 8)
'                itmX.HelpText = ""
''//////////Vouchers
'        Set barX = .Bars.Add("Voucher", , "Vouchers")
'            Set itmX = barX.Items.Add("Payments", , "Payments", 10)
'                itmX.HelpText = ""
'            Set itmX = barX.Items.Add("Receipts", , "Receipts", 11)
'                itmX.HelpText = ""
'            Set itmX = barX.Items.Add("Deposit", , "Deposit", 12)
'                itmX.HelpText = ""
'            Set itmX = barX.Items.Add("DNote", , "Delivery Notes", 13)
'                itmX.HelpText = ""
'            Set itmX = barX.Items.Add("CNote", , "Counter Notes", 14)
'                itmX.HelpText = ""
'            Set itmX = barX.Items.Add("ExpBook", , "Expense Books", 15)
'                itmX.HelpText = ""
'        Set barX = .Bars.Add("Transaction", , "Transaction")
'            Set itmX = barX.Items.Add("SalesReg", , "Sales Registry", 1)
'                itmX.HelpText = "Record Purchases of Items"
'            Set itmX = barX.Items.Add("SalesRet", , "Sales Return", 17)
'                itmX.HelpText = "Make Vocher and Do Payments to Vendors"
'            Set itmX = barX.Items.Add("SalesOrd", , "Sales Order", 18)
'                itmX.HelpText = ""
'            Set itmX = barX.Items.Add("PurReg", , "Purchase Registry", 19)
'                itmX.HelpText = ""
'            Set itmX = barX.Items.Add("PurRet", , "Purchase Return", 20)
'                itmX.HelpText = ""
'            Set itmX = barX.Items.Add("PurOrd", , "Purchase Order", 21)
'                itmX.HelpText = ""
'        Set barX = .Bars.Add("IBooks", , "Inventory Books")
'            Set itmX = barX.Items.Add("Stock Item", , "Stock Item", 22)
'                itmX.HelpText = "Record Purchases of Items"
'            Set itmX = barX.Items.Add("Group Summary", , "Group Summary", 23)
'                itmX.HelpText = "Make Vocher and Do Payments to Vendors"
'            Set itmX = barX.Items.Add("Physical Stock register", , "Physical Stock register", 24)
'                itmX.HelpText = ""
'            Set itmX = barX.Items.Add("Inventory Statement", , "Inventory Statement", 25)
'                itmX.HelpText = ""
'            Set itmX = barX.Items.Add("Movement Analysis", , "Movement Analysis", 26)
'                itmX.HelpText = ""
'
'        Set barX = .Bars.Add("ABooks", , "Accounts Books")
'            Set itmX = barX.Items.Add("Cash Book", , "Cash Book", 21)
'                itmX.HelpText = "Record Purchases of Items"
'            Set itmX = barX.Items.Add("Ledger", , "Ledger", 3)
'                itmX.HelpText = "Make Vocher and Do Payments to Vendors"
'            Set itmX = barX.Items.Add("Group Summary", , "Group Summary", 4)
'                itmX.HelpText = ""
'            Set itmX = barX.Items.Add("Sales Register", , "Sales Register", 5)
'                itmX.HelpText = ""
'            Set itmX = barX.Items.Add("Purchase Register", , "Purchase Register", 6)
'                itmX.HelpText = ""
'            Set itmX = barX.Items.Add("Statement of Accounts", , "Statement of Accounts", 7)
'                itmX.HelpText = ""
'            Set itmX = barX.Items.Add("Day Book", , "Day Book", 8)
'                itmX.HelpText = ""
'
'
'
'        .Bars(1).OfficeXpStyle = True
'        .Bars(2).OfficeXpStyle = True
'        .Bars(3).OfficeXpStyle = True
'        .Bars(4).OfficeXpStyle = True
'        .Bars(5).OfficeXpStyle = True
'        .Bars(6).OfficeXpStyle = True
'
'
'
'
'   End With
'   Set itmX = Nothing
'   Set barX = Nothing
'End Sub
'
'
Private Sub mnuStockItem_Click()
Load Form2
Form2.Show vbModal

'Load Form1
' frmControlMain.WBrow.Visible = True
'    frmControlMain.DataGrid1.Visible = False
''SQL = "SELECT Product.Name," _
'    & " prodOut.Quantity as [QntyOut]," _
'    & " prodIn.Quantity as [QntyIN]," _
'    & " prodIn.TotalPurchase," _
'    & " prodOut.Total AS TotalSales" _
'    & " FROM prodOut RIGHT JOIN " _
'    & " (prodIn RIGHT JOIN Product ON prodIn.ProductID = Product.ProductID) ON prodOut.ProductID = Product.ProductID"
'
'SQL = "Select * From dprodstata"
'
'Call frmControlMain.CreateSubPage(SQL, "Stock Analysis")
End Sub

Private Sub mnuViewall_Click()
SQL = "Select username1 from users"

frmDelUser.Caption = "View Users"
frmDelUser.Command1.Visible = False
Load frmDelUser
frmDelUser.Show vbModal

End Sub
