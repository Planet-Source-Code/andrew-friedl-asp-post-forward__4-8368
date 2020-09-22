<div align="center">

## ASP Post Forward


</div>

### Description

This simple ASP script uses WinHTTP to forward an ASP POST to an entirely different server and retun the content to the original ASP page. It is useful for serving up client content from another http source. In my case it is used to serve PDF reports from Java Server Pages to an ASP front end website. With this script you can server the content to the external user while the origin of the content itself appears to be the internet server. (People are unaware of the Apacha Tomcat Server running JSP and wont try to hack into it.)
 
### More Info
 
An HTML form post.

You webserver can result the IP address of the forwarded server.

The exact response from the forwarded server to the original ASP page.

My not be suitable for posting binary data, form posts work fine though.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Andrew Friedl](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/andrew-friedl.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Server Side](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/server-side__4-31.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/andrew-friedl-asp-post-forward__4-8368/archive/master.zip)

### API Declarations

Free as long as copyright info remains intact.


### Source Code

```
<%@ LANGUAGE = "VBScript" %><%
' Module : ASP POST Forward
' Author : Andrew Friedl
' Copyright: 2003.04.15 - Andrew Friedl
' License : License is hereby granted for commercial and
'   non-commercial use as long as the authors
'   information and copyright remains intact.
Dim http: Set http=Server.CreateObject("WinHttp.WinHttpRequest.5")
Private Function ConvertBin(bsString)
	Dim nIndex
	For nIndex = 1 to LenB(bsString)
	 ConvertBin = ConvertBin & Chr(AscB(MidB(bsString,nIndex,1)))
	Next
End Function
http.Open "POST", "http://TargetServer/TargetPage.ASP", False
http.SetTimeouts 60000, 60000, 60000, 60000
http.SetRequestHeader "Content-Type", Request.ServerVariables("CONTENT_TYPE")
http.Send Trim(ConvertBin(Request.BinaryRead( Request.TotalBytes )))
Response.Clear
Response.ContentType = http.GetResponseHeader("content-type")
Response.BinaryWrite http.ResponseBody
Response.End
%>
```

