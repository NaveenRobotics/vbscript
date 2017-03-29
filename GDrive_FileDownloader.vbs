' Set your settings
strFileURL = "https://docs.google.com/spreadsheets/d/[FILE_ID]/export?format=xlsx"
strHDLocation = "C:\Users\naveen.sakthivel\Desktop\Baby\file.xls"

' Fetch the file
Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP.3.0")

objXMLHTTP.open "GET", strFileURL, false
objXMLHTTP.send()

'Response 200 is OK, now download sheet
If objXMLHTTP.Status = 200 Then
  Set objADOStream = CreateObject("ADODB.Stream")
  objADOStream.Open
  objADOStream.Type = 1 'adTypeBinary

  objADOStream.Write objXMLHTTP.ResponseBody
  objADOStream.Position = 0    'Set the stream position to the start

  objADOStream.SaveToFile strHDLocation
  objADOStream.Close
  Set objADOStream = Nothing
End if
