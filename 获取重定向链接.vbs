Const WinHttpRequestOption_EnableRedirects = 6
Dim url
If WScript.Arguments.Count>0 Then url = WScript.Arguments(0)
url = InputBox("输入需要检测重定向的网址","输入",url)
If url = False Then WScript.Quit

'获取重定向
Dim WinHttp
Set WinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
WinHttp.Open "GET", url, False
WinHttp.Option(WinHttpRequestOption_EnableRedirects) = False
WinHttp.Send
If WinHttp.Status = 302 Or WinHttp.Status = 301 Or WinHttp.Status = 303 Or WinHttp.Status = 307 Then
	Dim result
	result = WinHttp.GetResponseHeader("Location")
	x = InputBox("该网址重定向到","结果",result)
Else
	WScript.Echo "没有重定向，网址返回 " & WinHttp.Status & " 状态"
End If
 
Set WinHttp = Nothing