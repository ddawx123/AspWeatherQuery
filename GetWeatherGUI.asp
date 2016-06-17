<%
dim act, fileUrl, fileName 
act=request.QueryString("act") 
if(act="do") then  
 fileUrl="http://121.41.93.32/road/weather.html?keyword=" & request.Form("fileUrl")
 if request.Form("fileUrl")="" then
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
<title>天气查询系统</title>
</head>

<body>
<form name="form1" method="post" action="?act=do"> 
  <p align="center"><strong style="font-size: 36px">内网天气查询系统_v1.0</strong><span style="font-size: 36px">_alpha</span></p>
  <p align="center">请输入您要查询的城市名称：
  <input name="fileUrl" type="text" size="30"> 
  </p>
  <p align="center">使用说明：请输入您所要查询的市级名称。诸如：绍兴、杭州、温州、南京等等，切勿输入省级名称，如：浙江、江苏等等！</p>
  <p align="center">注意：如果处于机房网络且尚未开放外网的情况下，本系统依旧可用。但是会无法显示天气图标！这是一个bug，敬请谅解！</p> 
  <p align="center"> 
    <input type="submit" name="Submit" value="提交"> 
    <input type="reset" name="Submit2" value="重写">
  </p>
  <p align="center"><a href="api.html" target="_self">如果需要天气小控件，请查看公共调用接口使用方法</a>
  </p>
  <p align="center">本系统由<a href="http://ddawx123.blog.163.com" target="_blank">小丁工作室</a>设计(15电脑高工1班-<a href="http://954759397.qzone.qq.com" target="_blank">丁鼎</a>-20152707)，版权所有，侵权必究！</p> 
</form> 
</body>
</html>
