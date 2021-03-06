'====================================
'变量定义区
'====================================
Dim FromCharset,ToCharset,Mode,cLog,LogName,cRecover,RecoverName
Recurrence = True '是否开启子文件夹递归操作
cLog = True '是否生成日志文件
cRecover = True '是否生成改名恢复文件
PartRename = 0 '是否分片段判断（0为完整文件名判断，1为由ascii字符隔断的各部分分别判断是否需要转码，2为每个字符都进行判断）
NoCovertChar = "優黒鳥雛勝雜圖鰤" '不要进行转换的字符，作为少量的无法判断是否是乱码的补充

FromCharset = "GBK" '乱码所在系统编码
ToCharset = "Shift-JIS" '乱码原始系统编码
LogName = WScript.ScriptFullName & "_" & Replace(Replace(FormatDateTime(Now(),vbGeneralDate),"/","-"),":","-") & ".log" '日志名称
RecoverName = WScript.ScriptFullName & "_" & Replace(Replace(FormatDateTime(Now(),vbGeneralDate),"/","-"),":","-") & ".恢复.bat" '恢复批处理名称

Dim Appname,Ver(2),fso,DebugMode
Appname = "日文乱码文件名修正" '程序名称
Ver(0) = 3:Ver(1) = 1:Ver(2) = 0 '程序版本
Set fso = CreateObject("scripting.FileSystemObject") '文件操作系统对象
Set osh = CreateObject("WScript.Shell")
DebugMode = False
'====================================
'程序启动判定区
'====================================
If WScript.Arguments.Count<1 Then
	WScript.Echo Appname & " V" & Version & vbcrlf & "请把需改名的文件拖到本脚本上运行（既使用参数方式提供路径）"
	WScript.Quit
ElseIf LCase(Right(WScript.FullName,11)) = "wscript.exe" Then
	Dim ai,argStr
	argStr = """" & WScript.Arguments(0) & """"
	For ai = 1 To WScript.Arguments.Count-1
		argStr = argStr & " """ & WScript.Arguments(ai) & """"
	Next
    osh.run "cmd /c cscript.exe //nologo """ & WScript.ScriptFullName & """ " & argStr
    WScript.quit
End If

'If InStr(1,WScript.FullName,"WScript.exe",vbTextCompare)>1 Then
'	nocmd = MsgBox("本版程序建议使用命令行模式运行，请将文件拖到bat引导脚本上面运行，否则会产生大量对话框。"& CharRepeat(vbCrLf,2) &"若不慎运行可结束“wscript.exe”进程快速关闭。" & CharRepeat(vbCrLf,2) & "是否结束本次运行？（推荐选是）",vbYesNo + vbExclamation,Appname & " V" & Version & " 提示")
'	If nocmd = vbyes Then WScript.Quit
'End If
'If WScript.Arguments.Count<1 Then
'	WScript.Echo Appname & " V" & Version & vbcrlf & "请把需改名的文件拖到bat脚本上运行"
'	WScript.Quit
'End If
'====================================
'函数区
'====================================
'生成版本号
Function Version()
	Version = Join(Ver,".")
End Function
'生成指定数字字符
Function CharRepeat(char,count)
	CharRepeat = ""
	For ci = 1 To count
		CharRepeat = CharRepeat & char
	Next
End Function

'转换字符串编码
Function gTs(str,charset1,charset2)
	Set adostream = CreateObject("ADODB.Stream")
	With adostream
		.Type = 2
		.Open
		.Charset = charset1
		.WriteText str
		.Position = 0
		.Charset = charset2
		gTs = .ReadText
		.close
	End With
	Set adostream = Nothing
End Function

'转换文件编码格式
Function uTu8(FilePath1,FilePath2,charset1,charset2)
	dim str
	str = ""
	Set adoStream = CreateObject("ADODB.Stream")
	Set newStream = CreateObject("ADODB.Stream")
	With adoStream
		.Mode = 3
		.Type = 2
		.Open
		.LoadFromFile FilePath1
		.Position = 0
		.Charset = charset1
		str = .ReadText
		.Position = 0
		.SetEOS
		.Charset = charset2
		.writetext str
		.SaveToFile FilePath2, 2
		
		.Position = 3
		newStream.Mode = 3
		newStream.Type = 1
		newStream.Open()
		.CopyTo(newStream)
		newStream.SaveToFile FilePath2,2
		newStream.Close
		
		.Flush
		.Close
	End With
	Set adoStream = Nothing
	Set newStream = Nothing
End Function

'16进制转正
Function pHex(Hnum)
	pHex = CLng("&H" & Hnum)
End Function

'检测单个字符串是否已被转换过
Function Converted(strc,newstr)
	Converted = False
	newstr = gTs(strc,FromCharset,ToCharset)
	
	'检测是否含有不要进行转换的字符
	If Len(NoCovertChar) > 0 Then
		If RegExpTest(strc,"[" & NoCovertChar & "]") Then
			Converted = True
'			If DebugMode Then ShowLog "含有不要进行转换的字符"
			Exit Function
		End If
	End If
		
	'检测是否为英文等转码后与原文一致
	If strc = newstr Then
		Converted = True
'		If DebugMode Then ShowLog "与原文一致 " & newstr
		Exit Function
	End If
	'转换前后长度不一致则为转换过的
	If Len(strc) <> Len(newstr) Then
		Converted = True
'		If DebugMode Then ShowLog "长度不一致 " & strc & " > " & newstr
		Exit Function
	End If
	'单字符Unicode码大于FF10的也算作转换过的
'	Dim si,ci,w
'	For si = 1 To Len(newstr)
'		w=Mid(newstr,si,1)
'		If pHex(Hex(AscW(w))) > pHex("FF10") Then
'			Converted = True
'			If DebugMode Then ShowLog "大字符 " & Mid(strc,si,1) & " > " & w & " " & Hex(AscW(w))
'			Exit For
'		End If
'	Next
End Function

'是否符合正则表达式
Function RegExpTest(strng, patrn) 
	Dim regEx      ' 创建变量。
	Set regEx = New RegExp         ' 创建正则表达式。
	regEx.Pattern = patrn         ' 设置模式。
	regEx.IgnoreCase = True         ' 设置是否区分大小写，True为不区分。
	regEx.Global = True         ' 设置全程匹配。
	RegExpTest = regEx.Test(strng)   ' 执行搜索。
	Set regEx = Nothing
End Function

'正则表达式搜索
Function RegExpSearch(strng, patrn) 
	Dim regEx      ' 创建变量。
	Set regEx = New RegExp         ' 创建正则表达式。
	regEx.Pattern = patrn         ' 设置模式。
	regEx.IgnoreCase = True         ' 设置是否区分大小写，True为不区分。
	regEx.Global = True         ' 设置全程匹配。
	Set RegExpSearch  = regEx.Execute(strng)
	
	'Debug Start
'	If RegExpSearch.Count > 0 Then
'		Dim txttmp
'		For t1 = 0 To RegExpSearch.Count -1
'			txttmp = txttmp & (t1+1) & ":" & RegExpSearch.Item(t1) & "; "
'			If RegExpSearch.Item(t1).Submatches.Count > 0 Then
'				For t2 = 0 To RegExpSearch.Item(t1).Submatches.Count -1
'					txttmp = txttmp & (t1+1) & "," & (t2+1) & ":" & RegExpSearch.Item(t1).Submatches.Item(t2) & "; "
'				Next
'			End If
'		Next
'		ShowLog txttmp
'	End If
	'Debug End
	
	Set regEx = Nothing
End Function

'新文件名
Function NewFilename(fname)
	Dim SearchPatrn
	NewFilename = fname
	If PartRename = 1 Then
		SearchPatrn = "[^\x20-\xFF]+"
	ElseIf PartRename = 2 Then
		SearchPatrn = "[^\x20-\xFF]"
	Else
		'完整文件名搜寻模式
		If Not Converted(fname,NewChar) Then
			NewFilename = NewChar
		End If
		Exit Function
	End If
	
	'片段搜寻模式
	Dim fStr,is1
	Set fStr = RegExpSearch(fname,"[^\x20-\xFF]+") '搜索所有非ASCLL字符片段
	For is1 = 0 To fStr.Count -1
		If Converted(fStr.Item(is1),NewChar) Then
			'不需要做什么
		Else
			If DebugMode Then ShowLog fStr.Item(is1).FirstIndex & "  """ & fStr.Item(is1) & """ |> """ & NewChar & """"
			NewFilename = Left(NewFilename,fStr.Item(is1).FirstIndex) & NewChar & Mid(NewFilename,fStr.Item(is1).FirstIndex + fStr.Item(is1).length + 1)
		End If
	Next
	Set fStr = Nothing
End Function

'显示日志
Function ShowLog(str)
	WScript.Echo Replace(str,vbTab," ")
	If cLog then
		Set fLog = fso.opentextfile(LogName,8,True,-1)
		fLog.WriteLine(str)
		Set fLog = Nothing
	End If
End Function

'保存恢复BAT
Function RecoverBat(str)
	If cLog then
		Set fRecover = fso.opentextfile(RecoverName,8,True,-1)
		fRecover.WriteLine(str)
		Set fRecover = Nothing
	End If
End Function

'改变文件名
Function ChangeFilename(file,ByVal LogPre,level)
	ChangeFilename = file.path
	Dim NewFlname
	NewFlname = NewFilename(file.name) '新名称
	If NewFlname = file.name Then
		ShowLog LogPre & vbTab & "-" & vbTab &  "名称不需转码" & vbTab & """" & file.Name & """" 
		skipnum = skipnum + 1
	Else
		oldname = file.name
		newname = NewFlname
		newfpath = file.ParentFolder & "\" & newname
		If fso.FileExists(newfpath) Or fso.FolderExists(newfpath) Then
			ShowLog LogPre & vbTab & "错误" & vbTab & """" & oldname & """" & vbTab & "同名目录已存在"  & vbTab & """" & newname& """" 
			errornum = errornum + 1
		Else
			If fso.FolderExists(file.path) Then
				If cRecover Then RecoverBat "Rem 文件夹改名 "
				If Not DebugMode Then
					fso.MoveFolder file,newfpath
					ChangeFilename = newfpath
				End if
			Else
				If Not DebugMode Then
					fso.MoveFile file,newfpath
					ChangeFilename = newfpath
				End if
			End If
			changenum = changenum + 1
			ShowLog LogPre & vbTab & "成功" & vbTab & """" & oldname & """" & vbTab & "已更名为" & vbTab & """" & newname& """"
			If cRecover Then RecoverBat "echo 将 """ & newname & """ 还原为 """ & oldname & """" & vbCrLf & "Rename """ & newname & """ """ & oldname & """"
		End If
	End If
	
End Function

'子文件夹递归
Function SubfolderRecurrence(ByVal folder,ByVal LogPre,level)
	Dim filet,fileid,folderid,nextlevel
	fileid = 0:folderid = 0
	ShowLog CharRepeat(vbTab,level-1) & LogPre & vbTab & """" & folder.Path  & """" & vbTab & "路径下有" & vbTab & folder.Files.Count & "个文件，" & folder.SubFolders.Count & "个子文件夹。"

	'子文件夹递归
	Set Filest = folder.SubFolders
	For Each filet In Filest
		folderid = folderid + 1
		LogPre = CharRepeat(vbTab,level) & folderid & vbTab & " / " & Filest.Count & " 子文件夹"
		nextlevel = level + 1
		oldFileName = filet.Name
		newFile = ChangeFilename(filet,LogPre,level) '改名
		If cRecover Then RecoverBat "cd """ & oldFileName & """" '到父文件夹路径恢复命名
		SubfolderRecurrence fso.GetFolder(newFile),LogPre,nextlevel '递归
		If cRecover Then RecoverBat "cd .." '到父文件夹路径恢复命名
	Next

	'文件
	Set Filest = folder.Files
	For Each filet In Filest
		fileid = fileid + 1
		LogPre = CharRepeat(vbTab,level) & fileid & vbTab & " / " & Filest.Count '& " 文件"
		ChangeFilename filet,LogPre,level
	Next
	Set filest = Nothing
	Set filet = Nothing
End Function
'====================================
'主代码
'====================================
Dim Files,changenum,errornum,skipnum, _
	LogTmp

Set Files = WScript.Arguments '将参数（文件列表）存入类
changenum = 0 '改名数
skipnum = 0 '跳过数
errornum = 0 '错误数
If DebugMode Then ShowLog "Debug Mod"
ShowLog "==========================================" & vbCrLf _
	& Date() & " " & Time() & " start" & vbCrLf _
	& Appname & " V" & Version & vbCrLf
If cRecover Then RecoverBat "@echo off" & vbcrlf & "chcp 65001"
For i = 0 To Files.Count-1
	LogTmp = (i+1) & " / " & Files.Count & " 初始提交文件"
	If fso.FolderExists(Files(i)) Then
		Set file = fso.getFolder(Files(i))
		If cRecover Then RecoverBat "cd /d """ & file.ParentFolder.Path & """" '到父文件夹路径恢复命名
		oldFileName = file.Name
		newFile = ChangeFilename(file,LogTmp,0) '改名
		'子文件夹递归
		If Recurrence Then
			If cRecover Then RecoverBat "cd """ & oldFileName & """" '到父文件夹路径恢复命名
			SubfolderRecurrence fso.GetFolder(newFile),LogTmp,1
			If cRecover Then RecoverBat "cd .." '到父文件夹路径恢复命名
		End If
	ElseIf fso.FileExists(Files(i)) Then
		Set file = fso.getFile(Files(i))
		If cRecover Then RecoverBat "cd /d """ & file.ParentFolder.Path & """" '到父文件夹路径恢复命名
		ChangeFilename file,LogTmp,0
	Else
		ShowLog LogTmp & vbTab & "错误" & vbTab & "文件" & vbTab & "不存在" & vbTab & """" & Files(i) & """" 
		errornum = errornum + 1
	End If
Next
If cRecover Then
	RecoverBat "echo 更名还原完毕 " & vbCrLf & "pause"
	uTu8 RecoverName,RecoverName,"UTF-16LE","UTF-8"
End If
ShowLog vbCrLf & "总共改名" & changenum & "个，错误" & errornum & "个，跳过" & skipnum & "个。" & vbCrLf _
	& Date() & " " & Time() & " end" & vbCrLf  _
	& "==========================================" & vbCrLf
Set fso=Nothing
WScript.Quit