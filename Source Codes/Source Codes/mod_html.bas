Attribute VB_Name = "mod_html"
Option Explicit
Public Function MakeShortReport(sqlString As String, header As String) As Boolean
'On Error GoTo adder:
Dim tempRs As New ADODB.Recordset
Dim i As Integer
Dim data2 As String
 Call GetNewConnection2
    Set Rs1 = New Recordset
       
    Set Rs1 = DCON.Execute(sqlString)
    
        frmView.Wbrow.Navigate "about:blank"
        
        Do While frmView.Wbrow.ReadyState <> READYSTATE_COMPLETE
            DoEvents
        Loop
   frmView.Wbrow.Document.Write ("<body oncontextmenu='return false;'>")

While Rs1.EOF <> True
    frmView.Wbrow.Document.Write ("<font face=arial>" & header & Rs1.Collect(0) & "<table width=100%>")
    For i = 0 To Rs1.Fields.Count - 1
     
            If IsNull(Rs1.Collect(i)) = True Then
                data2 = "N/A"
            Else
                data2 = Rs1.Collect(i)
            End If
           
        frmView.Wbrow.Document.writeln ("<TR><Td bgcolor=#cccccc><font size=2>" & Rs1.Fields(i).Name & "</td><td bgcolor=#CBC7B6> <font size=2>" & data2 & "</td></Tr>")
    Next i

Rs1.MoveNext
    frmView.Wbrow.Document.Write ("</font></table><BR><BR>")
Wend

   Set Rs1 = Nothing
    Set DCON = Nothing
 
  
    frmView.Show
    
   
'Exit Function
'adder:
'    MakeShortReport = False
End Function



Public Sub CreateH_Page(strSqry As String, titiles As String)
Dim fld As ADODB.Field
Dim i As Integer
Dim data2 As Variant
Call GetNewConnection2
    Set Rs1 = New ADODB.Recordset
    Set Rs1 = DCON.Execute(strSqry)
    WebSQL = strSqry
    frmView.Wbrow.Navigate2 "about:blank"
        Do While frmView.Wbrow.ReadyState <> READYSTATE_COMPLETE
            DoEvents
        Loop
    With frmView.Wbrow.Document
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

        frmView.Wbrow.Document.script.Document.clear
        frmView.Wbrow.Document.script.Document.Close
End With


Set DCON = Nothing
frmView.Show

End Sub

