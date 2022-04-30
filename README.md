https://github.com/bgoonz/WEB-DEV-NOTES/blob/47df2e93fee49054074dc0cffba970f5d6096124/00-4-all-time/general-ref/VBA-master/VBA-master/All%20BAS%20Files/HTML_VBA_Excel.bas


Private Sub HTML_VBA_Excel()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'If you run the code again without clearing the cache, then old data will be displayed again. To avoid this, use the below code.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Shell "RunDll32.exe InetCpl.Cpl, ClearMyTracksByProcess 11"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim oXMLHTTP As Object
    Dim sPageHTML  As String
    Dim sURL As String
 
    'Change the URL before executing the code
    sURL = "http://WWW.WebSiteName.com"
 
    'Extract data from website to Excel using VBA
    Set oXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP")
	
'Detect Broken URL or dead links in webpage
	
    oXMLHTTP.send
 
    If oXMLHTTP.Status <> 200 Then
        MsgBox sURL & ": URL Link is not Active"
    End If
	
	
	
	
	
	
	
    oXMLHTTP.Open "GET", sURL, False
    oXMLHTTP.send
    sPageHTML = oXMLHTTP.responseText
 
    'Get webpage data into Excel
    ThisWorkbook.Sheets(1).Cells(1, 1) = sPageHTML
 
    MsgBox "XMLHTML Fetch Completed"
 
End Sub
