VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frmControlMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8250
   ClientLeft      =   -45
   ClientTop       =   -435
   ClientWidth     =   10545
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   FillStyle       =   2  'Horizontal Line
   FontTransparent =   0   'False
   ForeColor       =   &H00404040&
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   10545
   ShowInTaskbar   =   0   'False
   Begin SHDocVwCtl.WebBrowser Wbrow 
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   4215
      ExtentX         =   7435
      ExtentY         =   4260
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   5318
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   15657958
      BorderStyle     =   0
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAcrossSplits =   -1  'True
      TabAction       =   2
      WrapCellPointer =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   720
      Left            =   0
      Picture         =   "frmControlMain.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10530
   End
End
Attribute VB_Name = "frmControlMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Private Sub DataGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If Button = 2 Then
        If DataGrid1.Columns(0).text <> "" Then
        If DataGrid1.Columns(0).text = "CASH" Then
            frmMAIN.mnuEdit.Enabled = False
        Else
            DataGrid1.SetFocus
            frmMAIN.mnuEdit.Enabled = True
            frmMAIN.mnuModify.Enabled = True
            PopupMenu frmMAIN.mnuEdit
        End If
        End If
    End If
    
End Sub



Private Sub Form_Load()

Dim SqLargs As String

    SqLargs = "SELECT Product.ProductID, Product.Name, Product.UnitsInStock, Product.UnitCostPrice From Product WHERE ((Product.UnitsInStock)<=0) Order by UnitsInStock DESC"
    Call CreateStartPage(SqLargs)

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub
Private Sub Form_Resize()
On Error Resume Next
    Image1.Width = Me.ScaleWidth
    If DataGrid1.Visible = True Then
        DataGrid1.Move 0, Image1.Height, Me.ScaleWidth - 100, Me.ScaleHeight - (Image1.Height + 100)
    End If
        WBrow.Move 0, Image1.Height, Me.ScaleWidth, Me.ScaleHeight - (Image1.Height + 150) '- 100
End Sub
Sub CreateStartPage(strSqry As String)
On Error GoTo adder:
Dim Rs1 As New ADODB.Recordset
    Rs1.CursorLocation = adUseClient
    GetNewConnection2
    Call Rs1.Open(strSqry, DCON, adOpenForwardOnly, adLockReadOnly)
Dim i As Integer
Dim data2 As Variant
WBrow.Navigate2 "about:blank"
        Do While WBrow.ReadyState <> READYSTATE_COMPLETE
            DoEvents
        Loop
        With WBrow.Document
        .Write ("<HTML><head></head><style type='text/css'> body,td{font-family: Arial;} body,td{font-size:11px;}</style>") 'Style
        .Write ("<BODY Scroll=Yes oncontextmenu='return false';>") '
        .Write ("Welcome To ProsVent 2005 Beta Version")
        .Write ("<table border=0 Width=100% height=80%>")
        .Write ("<tr><td valign=TOP width=80%><table Width=100% border=0>")
        'FIRST TITLE
        .Write ("<tr><td bgcolor=#B4C0DC Height=20>" & "Product Name")
        .Write ("<td bgcolor=#B15C0DC Height=20>" & "Number of Items")
        .Write ("<td bgcolor=#B15C0D0 Height=20>" & "Estimated Cost Unit Price")
                ''DATA COLUMN
        While Rs1.EOF <> True
                .Write ("<tr><td><li><A href='ID?" & Rs1.Collect(0) & "'>")  ''' this thing here is so bull shit
                .Write (Rs1.Collect(1)) ''Product Status
                .Write ("<td> <font color=Red>**" & Rs1.Collect(2) & "</td>")
                If Rs1.Collect(0) <= 0 Then
                         .Write ("<td></td>")
                    Else
                                .Write ("<td>" & Format(Rs1.Collect(3), "P .00") & "</td>")
                End If
                .Write ("</a></li></td></tr>")
                Rs1.MoveNext
        Wend
                .Write ("</td></tr></table></td><td><td valign=TOP>")
       .Write ("</td></table></BODY></HTML>")
       WBrow.Document.script.Document.clear
        WBrow.Document.script.Document.Close
End With
adder:
Exit Sub
End Sub

Private Sub wbrow_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
On Error GoTo adder:
    Dim pos As Integer
    Dim newString As String
    pos = InStr(URL, "?")
    
     If pos > 0 Then
        Cancel = True
        newString = Mid(URL, pos + 1)
        
        newString = Replace(newString, "%20", " ", 1, Len(URL), vbTextCompare)
            SQL = "SELECT  Category, UnitsInOrder, ReorderLevel, ReorderQuantity, UnitSellingPrice,UnitCostPrice FROM Product where ProductID='" & newString & "'"
        '    rptState = "Product Details "
    
        '"Select  *  from Product where ProductID='" & newString & "'"
        Select Case rptState
           
        Case "SalesRegistry"
                SQL = "Select * from SalesRegistryDetail where salesregistryID='" & newString & "'"
        Case "PurchaseRegistry"
          SQL = "Select * , Quantity * rate as Amount from PurchaseRegistryDetail where PurchaseregistryID='" & newString & "'"
    '    Case Else
    '        Call CreateSubPage(SQL, rptState)
        End Select
       Call CreateSubPage(SQL, rptState)
        
    End If
    Exit Sub
adder:
     Exit Sub
End Sub

Sub CreateSubPage(strSqry As String, title As String)
On Error Resume Next
Dim tempRs As New ADODB.Recordset
Dim fld As ADODB.Field
Dim i As Integer
Dim data2 As Variant

Call GetNewConnection2
    Set Rs1 = New Recordset
    Set Rs1 = DCON.Execute(strSqry)
        
        WBrow.Navigate2 "about:blank"
        Do While WBrow.ReadyState <> READYSTATE_COMPLETE
            DoEvents
        Loop
     With WBrow.Document
        .Write ("<HTML><head></head><style type='text/css'> body,td{font-family: Arial;} body,td{font-size:11px;}</style>") 'Style
        .Write ("<BODY Scroll=Yes oncontextmenu='return false';>") '
        .Write (title)
        .Write ("<table Width=100% border=1><tr>")
        .Write ("</tr></table>")
        .Write ("<table Width=100% border=0><tr>")
        ''Headings
        For Each fld In Rs1.Fields
            .Write ("<td bgcolor=#B4C0DC Height=10>" & fld.Name & "</td>")
        Next fld
        'First row
            .Write ("<tr>")
        'Make Data Cells and Loop to Another Row
        While Rs1.EOF <> True
        i = i + 1
            For Each fld In Rs1.Fields
            If i Mod 2 <> 0 Then
                .Write ("<td>" & fld.Value & "</td>")
            Else
                .Write ("<td bgcolor=#CCCCC2>" & fld.Value & "</td>")
            End If
            Next fld
            .Write ("</tr>")
            Rs1.MoveNext
        Wend
            .Write ("</td></tr></table></BODY></HTML>")

        WBrow.Document.script.Document.clear
        WBrow.Document.script.Document.Close
End With
'adder:
End Sub
'

Private Sub wbrow_NavigateError(ByVal pDisp As Object, URL As Variant, Frame As Variant, StatusCode As Variant, Cancel As Boolean)
    MsgBox URL
End Sub
Sub CreateDataPage(strSqry As String, titiles As String)
On Error Resume Next
Dim fld As ADODB.Field
Dim i As Integer
Dim j As Integer
Dim data2 As Variant
Call GetNewConnection2
    Set Rs1 = New ADODB.Recordset
    Set Rs1 = DCON.Execute(strSqry)
    WebSQL = strSqry
    
    WBrow.Navigate2 "about:blank"
    'Wbrow.Navigate2 "about:blank"
        Do While WBrow.ReadyState <> READYSTATE_COMPLETE
            DoEvents
        Loop
    With WBrow.Document
        .Write ("<HTML><head></head><style type='text/css'> body,td{font-family: Arial;} body,td{font-size:11px;}</style>") 'Style
        .Write ("<BODY Scroll=Yes oncontextmenu='return false';>") '
        .Write (titiles)
        .Write ("<table Width=100% border=1><tr>")
        .Write ("</tr></table>")
        .Write ("<table Width=100% border=0><tr>")
        ''Headings
        For Each fld In Rs1.Fields
            .Write ("<td bgcolor=#B4C0DC Height=10>" & fld.Name & "</td>")
        Next fld
        'First row
            .Write ("<tr>")
        'Make Data Cells and Loop to Another Row
        While Rs1.EOF <> True
        i = i + 1
            For Each fld In Rs1.Fields
            
            If j = 0 Then
                .Write ("<td><A href='ID?" & fld.Value & "'>")  ''' this thing here is so bull shit
            Else
                .Write ("<td>")  ''' this thing here is so bull shit
            End If
            'If i Mod 2 <> 0 Then '' making facncy look here
                .Write ("" & fld.Value & "</a></td>")
                j = j + 1
            'Else
            '    .Write ("<td bgcolor=#CCCCC2>" & fld.Value & "</td>")
           ' End If
            Next fld
            .Write ("</tr>")
            Rs1.MoveNext
        Wend
            .Write ("</td></tr></table></BODY></HTML>")

        WBrow.Document.script.Document.clear
        WBrow.Document.script.Document.Close
End With
'adder:

Set DCON = Nothing

End Sub
'


