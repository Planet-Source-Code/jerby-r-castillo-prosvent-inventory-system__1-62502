Attribute VB_Name = "Mod_CON_DISP"
Option Explicit
Public Rs1 As New ADODB.Recordset
Public RS2 As New ADODB.Recordset
Public CON As New CDbase
Public DCON As New ADODB.Connection
Public WebSQL As String
Public SQL As String
Public EDT As Boolean
Public CurUser As String
Public CURRDATE As String
Public CURRTIME As String
Public DESTINATION As String
Public fsystem As FileSystemObject
Public Source As String
Public rptState As String
Public CatalogueName As String
Public ADDING As Boolean
Public MODIFYID As String

Public Sub Main()
Dim frmLog         As frmLogin
Dim frmMd          As frmMAIN
Dim COUNT1      As Long
If App.PrevInstance = True Then
        MsgBox "The System Is In Use." & vbTab, vbInformation
Else
If App.PrevInstance = True Then Exit Sub
    Set frmLog = New frmLogin
        frmLog.Show vbModal
If Not frmLog.OK Then
'        MsgBox "Unauthorized Validataion"
End
End If
        
'frmSplash.Show
Unload frmLog
Load frmMAIN
'frmMAIN.Show
frmSplash.Timer1.Enabled = True
End If
End Sub
Public Sub GetNewConnection2()
Dim sCNSTR As String

Set DCON = New Connection
DCON.CursorLocation = adUseClient

    sCNSTR = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\DATABASE\Thesis.mdb;"
    DCON.Open sCNSTR

'If DCON.State = adStateOpen Then
 '   Set GetNewConnection = DCON
'End If


End Sub
Public Sub GetNewConnection(ClassVar As Object)
'If TypeOf Classname Is CUpdate Then
   ' Set a = Classname
Set DCON = New ADODB.Connection
Set CON = ClassVar

CON.DBPath = App.Path & "\Database\Thesis.mdb"
CON.OpenDb

End Sub

Public Function CMB1(ByVal TABLE As String, ByVal Field As String, CMB As ComboBox, Optional Clause As String, Optional ItemClear As Boolean)

If ItemClear = True Then

CMB.clear
End If

Call GetNewConnection2


Set Rs1 = New Recordset

If Clause = "" Then
Set Rs1 = DCON.Execute("Select * from " & TABLE & "")

Else
Set Rs1 = DCON.Execute("Select * from " & TABLE & " " & Clause)

End If

If Rs1.RecordCount > 0 Then
    While Not Rs1.EOF
       
        CMB.AddItem Rs1.Fields(Field)
       
        Rs1.MoveNext
    Wend

    
    
End If

Set Rs1 = Nothing
Set DCON = Nothing


End Function
Public Function CMB3(ByVal TABLE As String, ByVal Field As String, CMB As ComboBox, Optional Clause As String, Optional ItemClear As Boolean)

If ItemClear = True Then

CMB.clear
End If

Call GetNewConnection2


Set Rs1 = New Recordset

If Clause = "" Then
Set Rs1 = DCON.Execute("Select DISTINCT " & Field & " from " & TABLE & "")

Else
Set Rs1 = DCON.Execute("Select DISTINCT " & Field & "  from " & TABLE & " " & Clause)

End If

If Rs1.RecordCount > 0 Then
    While Not Rs1.EOF
       
        CMB.AddItem Rs1.Fields(Field)
       
        Rs1.MoveNext
    Wend

    
    
End If

Set Rs1 = Nothing
Set DCON = Nothing


End Function
Public Function CMB2(ByVal sqlArg As String, CMB As ComboBox)

CMB.clear
Call GetNewConnection2
Set Rs1 = New Recordset
Set Rs1 = DCON.Execute(sqlArg)
If Rs1.RecordCount > 0 Then
    While Not Rs1.EOF
        CMB.AddItem Rs1.Collect(0)
       Rs1.MoveNext
    Wend
End If
Set Rs1 = Nothing
Set DCON = Nothing
End Function
Public Function Decimals(Key_Ascii As Integer, ByVal ControlName As Object, ByVal DecimalPlace As Integer)
On Error GoTo DECERR

Static DecPlace As Integer
If InStr(1, ControlName, ".") Then
    If Key_Ascii <> 13 And Key_Ascii <> 8 Then
    If Key_Ascii < 48 Or Key_Ascii > 57 Then Key_Ascii = 0
    
    End If
    If DecPlace = 0 Then
    DecPlace = Val(Len(ControlName) + DecimalPlace)
    ControlName.MaxLength = DecPlace
   
    End If
Else
    DecPlace = 0
    If Key_Ascii <> 13 And Key_Ascii <> 8 And Key_Ascii <> 46 Then
    If Key_Ascii < 48 Or Key_Ascii > 57 Then Key_Ascii = 0
    End If
End If
Exit Function

DECERR:
    MsgBox Err.Description & vbTab, vbInformation
    
End Function
Public Function OFFCHar(Key_Ascii As Integer, ByVal ControlName As Object)
On Error GoTo DECERR


    If Key_Ascii <> 13 And Key_Ascii <> 8 Then
    If Key_Ascii < 48 Or Key_Ascii > 57 Then Key_Ascii = 0
    
    End If
  
 
Exit Function

DECERR:
    MsgBox Err.Description & vbTab, vbInformation
    
End Function
Public Function offDefine(Key_Ascii As Integer, ByVal ControlName As Object, sFilter As String)

If InStr(sFilter, Chr(Key_Ascii)) = 0 Then
    Key_Ascii = 0
End If

End Function
Private Function WordTens(ByVal SNUM As Long) As String
Select Case SNUM
    Case 1
        WordTens = " One"
    Case 2
        WordTens = " Two"
    Case 3
        WordTens = " Three"
    Case 4
        WordTens = " Four"
    Case 5
        WordTens = " Five"
    Case 6
        WordTens = " Six"
    Case 7
        WordTens = " Seven"
    Case 8
        WordTens = " Eight"
    Case 9
        WordTens = " Nine"
    Case 10
        WordTens = " Ten"
    Case 11
        WordTens = " Eleven"
    Case 12
        WordTens = " Twelve"
    Case 13
        WordTens = " Thirteen"
    Case 14
        WordTens = " Fourteen"
    Case 15
        WordTens = " Fifteen"
    Case 16
        WordTens = " Sixteen"
    Case 17
        WordTens = " Seventeen"
    Case 18
        WordTens = " Eighteen"
    Case 19
        WordTens = " Nineteen"
    Case 20
        WordTens = " Twenty"
    Case 30
        WordTens = " Thirty"
    Case 40
        WordTens = " Fourty"
    Case 50
        WordTens = " Fifty"
    Case 60
        WordTens = " Sixty"
    Case 70
        WordTens = " Seventy"
    Case 80
        WordTens = " Eighty"
    Case 90
        WordTens = " Ninty"
End Select
End Function


Public Function NumToWord(ByVal src_num As String) As String
Dim SNUM  As Double
SNUM = Val(src_num)
If SNUM > 999999999999999# Then
    NumToWord = "Error: To much number."
    Exit Function
End If
Dim WHOLE As String
Dim EXTRA As String
Dim WORD  As String
Dim NWHOLE As Double

If InStr(1, Str$(SNUM), ".", vbTextCompare) <> 0 Then
   WHOLE = Split(Str$(SNUM), ".")(0)
    EXTRA = Split(src_num, ".")(1)
Else
    WHOLE = SNUM
End If

If SNUM < 1 Then WORD = "Zero"

NWHOLE = Val(WHOLE)
'Check for One and Tens
If Val(Right(NWHOLE, 2)) > 0 And Val(Right(NWHOLE, 2)) < 21 Or Val(Right(NWHOLE, 2)) = 30 Or Val(Right(NWHOLE, 2)) = 40 Or Val(Right(NWHOLE, 2)) = 50 Or Val(Right(NWHOLE, 2)) = 60 Or Val(Right(NWHOLE, 2)) = 70 Or Val(Right(NWHOLE, 2)) = 80 Or Val(Right(NWHOLE, 2)) = 90 Then
    WORD = WORD & WordTens(Val(Right(NWHOLE, 2)))
ElseIf Val(Right(NWHOLE, 2)) > 20 Then
    WORD = WORD & WordTens(Left(Right(NWHOLE, 2), 1) & "0")
    WORD = WORD & WordTens(Right(NWHOLE, 1))
End If
'Check for Hundred
If NWHOLE > 99 Then
   If Left(Right(NWHOLE, 3), 1) <> "0" Then WORD = WordTens(Left(Right(NWHOLE, 3), 1)) & " Hundred" & WORD
End If
'Check for Thousand
If NWHOLE > 999 Then
    If Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) > 0 And Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) < 21 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 30 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 40 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 50 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 60 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 70 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 80 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 90 Then
        WORD = WordTens(Val(Left(NWHOLE, Len("" & NWHOLE) - 3))) & " Thousand" & WORD
    ElseIf Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) > 20 And Right(Left(NWHOLE, Len("" & NWHOLE) - 3), 3) <> "000" Then
        WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 3)), 2), 2, 1)) & " Thousand" & WORD
        WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 3)), 2), 1, 1) & "0") & WORD
        If Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) > 99 Then
            If Left(Right(NWHOLE, 6), 1) <> "0" Then WORD = WordTens(Left(Right(NWHOLE, 6), 1)) & " Hundred" & WORD
        End If
    End If
End If
'Check for Million
If NWHOLE > 999999 Then
    If Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) > 0 And Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) < 21 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) = 30 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) = 40 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) = 50 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) = 60 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) = 70 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) = 80 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) = 90 Then
        WORD = WordTens(Val(Left(NWHOLE, Len("" & NWHOLE) - 6))) & " Million" & WORD
    ElseIf Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) > 20 And Right(Left(NWHOLE, Len("" & NWHOLE) - 6), 3) <> "000" Then
        WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 6)), 2), 2, 1)) & " Million" & WORD
        WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 6)), 2), 1, 1) & "0") & WORD
        If Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) > 99 Then
            If Left(Right(NWHOLE, 9), 1) <> "0" Then WORD = WordTens(Left(Right(NWHOLE, 9), 1)) & " Hundred" & WORD
        End If
    End If
End If
'Check for Billion
If NWHOLE > 999999999 Then
    If Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) > 0 And Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) < 21 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 30 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 40 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 50 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 60 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 70 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 80 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 90 Then
        WORD = WordTens(Val(Left(NWHOLE, Len("" & NWHOLE) - 9))) & " Billion" & WORD
    ElseIf Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) > 20 And Right(Left(NWHOLE, Len("" & NWHOLE) - 9), 3) <> "000" Then
        WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 9)), 2), 2, 1)) & " Billion" & WORD
        WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 9)), 2), 1, 1) & "0") & WORD
        If Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) > 99 Then
            If Left(Right(NWHOLE, 12), 1) <> "0" Then WORD = WordTens(Left(Right(NWHOLE, 12), 1)) & " Hundred" & WORD
        End If
    End If
End If
'Check for Trillion
If NWHOLE > 999999999999# Then
    If Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) > 0 And Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) < 21 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) = 30 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) = 40 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) = 50 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) = 60 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) = 70 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) = 80 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) = 90 Then
        WORD = WordTens(Val(Left(NWHOLE, Len("" & NWHOLE) - 12))) & " Trillion" & WORD
    ElseIf Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) > 20 And Right(Left(NWHOLE, Len("" & NWHOLE) - 12), 3) <> "000" Then
        WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 12)), 2), 2, 1)) & " Trillion" & WORD
        WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 12)), 2), 1, 1) & "0") & WORD
        If Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) > 99 Then
            If Left(Right(NWHOLE, 15), 1) <> "0" Then WORD = WordTens(Left(Right(NWHOLE, 15), 1)) & " Hundred" & WORD
        End If
    End If
End If
If EXTRA = "" Then
    WORD = WORD & "   and   00/100"
Else
    If Val(EXTRA) < 10 Then EXTRA = "0" & EXTRA
    WORD = WORD & "   and   " & EXTRA & "/100"
End If
NumToWord = WORD

NWHOLE = 0
WORD = ""
EXTRA = ""
WHOLE = ""
End Function

Public Function GRIDBIND(ByVal TABLE As String, ByVal Grid2 As DataGrid, Optional Clause As String)
Call GetNewConnection2
Set Rs1 = New Recordset
If Clause <> "" Then
Set Rs1 = DCON.Execute("Select * from " & TABLE & " " & Clause)
SQL = "Select * from " & TABLE & " " & Clause
Else
Set Rs1 = DCON.Execute("Select * from " & TABLE)
SQL = "Select * from " & TABLE
'Rs1.Open "Select * from " + Table, DCON, adOpenDynamic, adLockPessimistic
End If
  
Set Grid2.DataSource = Rs1



Set Rs1 = Nothing
Set DCON = Nothing


End Function


Public Sub GridRefresh()
Select Case CatalogueName
  
  Case "Customer"
         Call GRIDBIND("Customer", frmControlMain.DataGrid1, "WHERE CustomerID<>'CASH'")

  Case "Supplier"
         Call GRIDBIND("Supplier", frmControlMain.DataGrid1, "WHERE SuppliersID<>'CASH'")
  Case "Category"
         Call GRIDBIND("Category", frmControlMain.DataGrid1)
  Case "Location"
         Call GRIDBIND("location", frmControlMain.DataGrid1)
  Case "Purchase Order"
        Call GRIDBIND("PurchaseOrderHeader", frmControlMain.DataGrid1)
  Case "Purchase Return"
            Call GRIDBIND("PurchaseReturnHeader", frmControlMain.DataGrid1)
  Case "Purchase Registry"
      Call GRIDBIND("PurchaseRegistryHeader", frmControlMain.DataGrid1)
  Case "Sales Return"
            Call GRIDBIND("SalesReturnHeader", frmControlMain.DataGrid1)
  Case "Sales Registry"
             Call GRIDBIND("SalesRegistryHeader", frmControlMain.DataGrid1)
    
End Select
End Sub
