<%
dim act, fileUrl, fileName 
act=request.QueryString("city") 
if(act="") then
 response.write("Wrong call parameters.<br/>Please contact the Xiaoding Studio.QQ:954759397<br/>If you want to use GUI Version,please <a href='GetWeatherGUI.asp'>click here</a>.")
 else
  fileUrl="http://121.41.93.32/road/weather.html?keyword=" & act
 if request.Form("fileUrl")="http://121.41.93.32/road/weather.html?keyword=" then
 response.Redirect("error.html")
 else 
 fileName=mid(fileUrl,instrrev(fileUrl,"/")+1) 
 extPos=instrrev(fileName,"?") 
 if(extPos>0) then 
  fileName=left(fileName,extPos-1) 
 end if 
 call SaveRemoteFile(fileUrl, fileName) 
 response.Redirect("weather.html")
 end if
end if 

Function SaveRemoteFile(RemoteFileUrl, LocalFileName)
SaveRemoteFile=True
  dim Ads,Retrieval,GetRemoteData
  Set Retrieval = Server.CreateObject("Microsoft.XMLHTTP")
  With Retrieval
    .Open "Get", RemoteFileUrl, False, "", ""
    .Send
If .Readystate<>4 then
SaveRemoteFile=False
Exit Function
End If
    GetRemoteData = .ResponseBody
  End With
  Set Retrieval = Nothing
  Set Ads = Server.CreateObject("Adodb.Stream")
  With Ads
    .Type = 1
    .Open
    .Write GetRemoteData
    .SaveToFile server.MapPath(LocalFileName),2
    .Cancel()
    .Close()
  End With
  Set Ads=nothing
end Function
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>静默天气查询系统</title>
</head>

<body>

</body>
</html>
