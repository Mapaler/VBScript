'====================================
'变量定义区
'====================================
Dim TempPath,TempSpace,Mode,_
	p_libpath,p_ncvt,p_7zip,p_rar,_
	Charset,JpegQuality,DelOriginalPic,Overwrite,DeleteZip,_
	cLog,LogName
TempPath = "R:\picTemp" '解压缩模式临时文件位置
TempSpace = 2 * 10 ^ 9 '临时文件位置磁盘空余大小，单位Byte。10的9次方就是GB，这里是2GB（小于我内存盘的2GiB）
p_libpath = parDir(WScript.ScriptFullName) '存放本脚本调用其他程序的位置
p_ncvt = p_libpath & "\nconvert.exe" 'NConvert程序路径
p_7zip = p_libpath & "\7z.exe" '7-zip程序路径
p_rar = p_libpath & "\Rar.exe" 'RAR程序路径
Charset = "UTF-8" '图片列表txt文件编码
JpegQuality = 90 '导出jpg质量
DelOriginalPic = True '是否删除原始图片
Overwrite = True '是否覆盖已经有的图片
DeleteZip = False '解压后是否删除压缩包
cLog = True '是否生成日志文件
LogName = WScript.ScriptFullName & "_" & Replace(Replace(FormatDateTime(Now(),vbGeneralDate),"/","-"),":","-") & ".log" '日志名称

'格式设置
'---------------
Dim Txte,Imge,Zipe,Rare
'▼认为是图片列表的格式
Txte = "txt"
'▼需要转换成JPG的格式
Imge = "bmp,png,tif"
'▼7-zip支持的压缩包格式
Zipe = "7z,XZ,BZIP2,GZIP,TAR,ZIP,WIM,AR,ARJ,CAB,CHM,CPIO,CramFS,DMG,FAT,HFS,ISO,LZH,LZMA,MBR,MSI,NSIS,NTFS,RAR,RPM,SquashFS,UDF,UEFI,VHD,WIM,XAR,Z"
'▼RAR才能解压的压缩包格式（暂未支持）
Rare = "rar5"

'程序信息设置（仅供作者修改）
'---------------
Dim Appname,Ver(2),fso,osh,oExec
Appname = "固实7-Zip压缩包快速转JPG" '程序名称
Ver(0) = 2:Ver(1) = 8:Ver(2) = 0 '程序版本
Set fso = CreateObject("Scripting.FileSystemObject") '文件操作系统对象
Set osh = CreateObject("WScript.Shell")
'====================================
'程序启动判定区
'====================================
If WScript.Arguments.Count<1 Then
	WScript.Echo Appname & " V" & Version & vbcrlf & "请把需要处理的压缩包拖到本脚本上运行（既使用参数方式提供路径）"
	WScript.Quit
ElseIf Not fso.DriveExists(fso.GetDriveName(TempPath)) Then
	WScript.Echo TempPath & " ←该路径所在驱动器不存在。"
	WScript.Quit
ElseIf fso.GetDrive(fso.GetDriveName(TempPath)).FreeSpace < TempSpace Then
	WScript.Echo TempPath & " ←该路径所在磁盘可用空间不足，达不到本程序目前设置要求。"
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
'	WScript.Echo Appname & " V" & Version & vbcrlf & "请把需处理的文件拖到bat脚本上运行（既使用参数方式提供路径）"
'	WScript.Quit
'End If
'If fso.DriveExists(fso.GetDriveName(TempPath)) Then
'	If fso.GetDrive(fso.GetDriveName(TempPath)).FreeSpace < TempSpace Then
'		WScript.Echo TempPath & " ←该路径所在磁盘可用空间不足，达不到本程序目前设置要求。"
'		WScript.Quit
'	End If
'Else
'	WScript.Echo TempPath & " ←该路径所在驱动器不存在。"
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
'按编码读取文本
Function LoadText(FilePath,charset)
	Set adostream = CreateObject("ADODB.Stream")
	With adostream
		.Type = 2
		.Open
		.Charset = charset
		.Position = 0
		.LoadFromFile FilePath
		LoadText = .readtext
		.close
	End With
	Set adostream = Nothing
End Function
'按编码保存文本
Function SaveText(FilePath,str,charset)
	Set adostream = CreateObject("ADODB.Stream")
	With adostream
		.Type = 2
		.Open
		.Charset = charset
		.Position = 0
		.writetext str
		.SaveToFile FilePath, 2
		.flush
		.close
	End With
	Set adostream = Nothing
End Function
'显示日志
Function ShowLog(str)
	WScript.Echo str
	If cLog then
		Set fLog = fso.opentextfile(LogName,8,True,-1)
		fLog.WriteLine(str)
		Set fLog = Nothing
	End If
End Function
'获取路径父文件夹
Function parDir(path)
	Dim fsot
	Set fsot = CreateObject("scripting.filesystemobject") '文件操作系统对象
	If fsot.FolderExists(path) Then
		parDir = fsot.GetParentFolderName(fsot.GetFolder(path))
	ElseIf fsot.FileExists(path) Then
		parDir = fsot.GetParentFolderName(fsot.GetFile(path))
	End If
	Set fsot = Nothing
End Function
'图片格式转换
Function ConvertImage(FilePath)
	'仅转换这些格式的图片
	If ReturnMode(fso.GetExtensionName(FilePath)) <> 2 Then
		ConvertImage = False
		Exit Function
	End If
	Dim wsPath
	wsPath = osh.CurrentDirectory '记录原始路径
	osh.CurrentDirectory = parDir(FilePath) '跳转到指定文件夹
	Dim command
	'转格式的命令行
	command = """" & p_ncvt & """ -out jpeg -q " & JpegQuality & " -truecolors"
	If DelOriginalPic Then command = command & " -D" '删除原始文件
	If Overwrite Then command = command & " -overwrite" '覆盖存在文件
	command = command & " """ & fso.GetFileName(FilePath) & """"
	Set oExec = osh.Exec(command)
	'ShowLog command
	'strOut = oExec.StdOut.ReadAll()
	strErr = oExec.StdErr.ReadAll() '加上是为了保持完成后再继续
	
	osh.CurrentDirectory = wsPath '恢复原始路径
	
	'标准输出▼
	'ShowLog strOut
	'标准错误▼
	If Not RegExpTest(strErr,"^\s*Error:") Then
		'ShowLog strErr
		ConvertImage = True
	Else
		ConvertImage = False
	End If
End Function
'通过扩展名判断文件类型
Function ReturnMode(Extension)
	'文本
	Aext=split(Txte,",")
	For i = 0 To UBound(Aext)
		If RegExpTest(Extension,"^" & Aext(i) & "$") Then
			ReturnMode = 1
			Exit Function
		End If
	Next
	'需转换格式的图片
	Aext=split(Imge,",")
	For i = 0 To UBound(Aext)
		If RegExpTest(Extension,"^" & Aext(i) & "$") Then
			ReturnMode = 2
			Exit Function
		End If
	Next
	'压缩包
	Aext=split(Zipe,",")
	For i = 0 To UBound(Aext)
		If RegExpTest(Extension,"^" & Aext(i) & "$") Then
			ReturnMode = 3
			Exit Function
		End If
	Next
	'rar才能解压的压缩包
	Aext=split(Rare,",")
	For i = 0 To UBound(Aext)
		If RegExpTest(Extension,"^" & Aext(i) & "$") Then
			ReturnMode = 4
			Exit Function
		End If
	Next
End Function
'解压压缩包
Function Unzip_7z(FilePath)
	Unzip_7z = True	'解压是否正确
	ShowLog FilePath
	CreateFolder(TempPath) '创建总临时文件夹
	Dim zipName,TempflName,TempTextName,AllListTextName,target,targetLong
	zipName = fso.GetFileName(FilePath)
	TempflName = TempPath & "\" & zipName	'解压文件夹路径
	TempTextName = TempPath & "\" & zipName & ".txt"	'分段内容路径
	AllListTextName = TempPath & "\" & zipName & "_all.txt"	'完整内容路径
	target = parDir(FilePath) & "\" & DelSpace(fso.GetBaseName(FilePath)) '最终目标路径
	targetLong = False	'目标路径是否过长
		
	Dim command,oExec,strOut,strErr
	'7-Zip列出文件列表的命令行
	
'	command = """" & p_7zip & """ l "
'	command = command & " -sccUTF-8 "
'	command = command & " """ & FilePath & """>""" & AllListTextName & """"
	
	command = "cmd /c"
	command = command & " """"" & p_7zip & """ "
	command = command & " l -sccUTF-8 "
	command = command & " """ & FilePath & """"
	command = command & ">""" & AllListTextName & """"" "
	'WScript.echo command
	
	ShowLog "读取压缩包内文件列表"
	Set oExec = osh.Exec(command)
	'osh.Run command,0,True
	'Set oExec = osh.Exec("cmd")
	'oExec.StdIn.WriteLine """D:\CTFX\OneDrive\文档\VB6\小程序\VBS & BAT脚本\VBScript\7z.exe"" l -sccUTF-8  ""f:\System Sevice\会错误压缩包\超大、字符错误\[＊Cherish＊ (西村 にけ)] ユユカン2 (東方project).zip"">""Z:\picTemp\[＊Cherish＊ (西村 にけ)] ユユカン2 (東方project).zip_all.txt"""
	'oExec.StdIn.WriteLine command
	'oExec.StdIn.Close
	'得到标准输出流
	strOut = oExec.StdOut.ReadAll() '这个留着是让程序执行完毕后再继续后面的
	'MsgBox  strOut
	strOut = LoadText(AllListTextName,Charset)
	'strOut = gTs(strOut,"GB18030",Charset) '转换为UTF8
	'得到标准错误流
	'strErr = oExec.StdErr.ReadAll()
	
	'建立文件夹
	Dim Folderexist
	Folderexist = CreateFolder(TempflName)
	If Not Folderexist Then
		ShowLog "临时文件夹驱动器不存在，无法解压。"
		Exit Function
	End If

	Dim strOLine,CfgN
	Dim Matches
	Dim Cfgs,Items
	Set Cfgs = CreateObject("Scripting.Dictionary")'设置集合
	Set Items = CreateObject("Scripting.Dictionary")'文件集合
	
'---------------------------------
'取出设置段
'---------------------------------
	Set Matches = RegExpSearch(strOut,"^--$[\r\n]+((?:^\s*[\w\s]+?\s*=\s*.+?\s*$[\r\n]+)+)")
	If Matches.count > 0 Then
		'在设置段搜索每项设值
		Set Matches = RegExpSearch(Matches(0).SubMatches(0),"^\s*([\w\s]+?)\s*=\s*(.+?)\s*$")
		If Matches.count > 0 Then
			For li = 0 To Matches.count - 1
				'添加基本信息到集合
				Cfgs.Add Matches(li).SubMatches(0) , Matches(li).SubMatches(1)
			Next
		End if
	End If
	'查看设置DEBUG用
'		for each key in Cfgs.Keys
'		    Wscript.Echo "a ---" & key & "<>" & Cfgs.Item(key)
'		Next
	If Cfgs.Count > 0 Then ShowLog "压缩包格式：" & Cfgs.Item("Type")
'---------------------------------
'取出文件列表段
'---------------------------------
	ShowLog "建立文件列表数组"
	Dim sMatches,Attrs
	Set Matches = RegExpSearch(strOut,"(^\-+\s+\-+\s+\-+\s+\-+\s+\-+$)[\r\n]+([\s\S]+)[\r\n]+\1")
	If Matches.count > 0 Then
'			WScript.Echo "全部横杠" & Matches(0).SubMatches(0)
'			WScript.Echo "文件列表" & Matches(0).SubMatches(1)
		'获取横杠数字
		Set sMatches = RegExpSearch(Matches(0).SubMatches(0),"\-+")
		'查看各长度DEBUG用
'			For li = 0 To sMatches.count - 1
'				ShowLog sMatches(li).FirstIndex & " " & Len(sMatches(li))
'			Next
		strOLine = Split(Matches(0).SubMatches(1),vbCrLf) '将标准输出流按行分割
		Dim perTxt,longPath
		For li = 0 To UBound(strOLine)
			If Len(strOLine(li)) > 0 Then
				Set Attrs = CreateObject("Scripting.Dictionary")'文件集合
				Items.Add li,Attrs
				Items.Item(li).Add "DateTime" , DelSpace(Mid(strOLine(li),sMatches(0).FirstIndex + 1,Len(sMatches(0))))
				Items.Item(li).Add "Attr" , DelSpace(Mid(strOLine(li),sMatches(1).FirstIndex + 1,Len(sMatches(1))))
				Items.Item(li).Add "Size" , CLng(0 & DelSpace(Mid(strOLine(li),sMatches(2).FirstIndex + 1,Len(sMatches(2)))))
				Items.Item(li).Add "Compressed" , CLng(0 & DelSpace(Mid(strOLine(li),sMatches(3).FirstIndex + 1,Len(sMatches(3)))))
				Items.Item(li).Add "Name" , DelSpace(Mid(strOLine(li),sMatches(4).FirstIndex + 1))
				
				longPath = target & "\" & Items(fi).Item("Name")
				If Len(longPath) > 255 And targetLong = False Then '发现文件名过长，只保留左边5个字符
					target = parDir(FilePath) & "\" & Left(DelSpace(fso.GetBaseName(FilePath)),5)
					targetLong = True
				End If
			End If
			
			WScript.StdOut.Write CharRepeat(Chr(8),LenB(perTxt))
			perTxt = li+1 & "/" & UBound(strOLine)+1
			WScript.StdOut.Write perTxt
		Next
	End If
	WScript.StdOut.WriteBlankLines 1
	
	If Items.Count = 0 Then
		ShowLog "错误：压缩包中未发现文件列表"
		Unzip_7z = False
		Exit Function
	End If
	'查看内容DEBUG用
'		For Each keya In Items.Keys
'			Wscript.Echo keya
'			For Each key2 In Items(keya).Keys
'				Wscript.Echo "a ---" & key2 & "<>" & Items(keya).Item(key2)
'			Next 
'		Next

	ShowLog "建立完成"

'---------------------------------
'开始解压段
'---------------------------------
	Dim fi,TempSize
	Dim TempList
	
	Do Until fi >= Items.Count - 1
		TempSize = 0
		TempList = ""
'			If fso.FileExists(TempTextName) Then fso.DeleteFile(TempTextName)
		For fi = fi To Items.Count - 1
			If Not RegExpTest(Items(fi).Item("Attr"),"D.{4}") Then
				TempSize = TempSize + Items(fi).Item("Size")
				If Items(fi).Item("Size") > TempSpace Then
					ShowLog "错误：“" & Items(fi).Item("Name") & "” 文件大小超出临时文件夹设定空间"
					fi = fi + 1
				End if
				If TempSize > TempSpace Then
					Exit for
				End If
				'添加到列表
				TempList = TempList & Items(fi).Item("Name") & vbCrLf
			End If
			SaveText TempTextName,TempList,Charset
		Next
		
'			ShowLog TempList
		ShowLog "大小达到设定值，目前文件序号：" & (fi) & "/" & Items.Count & "。"
		ShowLog "开始解压。"
		
		'7-Zip解压文件的命令行
		command = """" & p_7zip & """ x "
		command = command & " """ & FilePath & """ " '压缩包地址
		command = command & " -bb1 " '日志显示文件。
		command = command & " -bsp1 " '日志显示文件。
		command = command & " -aoa " '直接覆盖现有文件，而没有任何提示。
		command = command & " -o""" & TempflName & """ " '输出位置
		command = command & " @""" & TempTextName & """ " '列表文件位置
		'ShowLog "命令行：" & command
'			oExec = osh.Run(command,6,True)
		Dim ReadLine,ReadLine2,NumSearch
		Set oExec = osh.Exec(command)
		Do While oExec.StdOut.AtEndOfStream <> True
			'可加入删除符，解压状态保留在同一行
			ReadLine = oExec.StdOut.ReadLine
			If RegExpTest(ReadLine,"^\s*\d+%") Then
				WScript.StdOut.Write CharRepeat(Chr(8),LenB(ReadLine))
				'ReadLine = Replace(ReadLine,vbCr,"")
				ReadLine2 = ReadLine
				Set NumSearch = RegExpSearch(ReadLine,"^(\s*\d+%\s*\d*)")
				If NumSearch.Count > 0 Then
					ReadLine = NumSearch.Item(0).Submatches.Item(0)
				End If
				WScript.StdOut.Write ReadLine
			Else
				'ShowLog ReadLine
			End If
		Loop
		WScript.StdOut.Write CharRepeat(Chr(8),LenB(ReadLine2))
		ShowLog ReadLine2
		'WScript.StdOut.WriteBlankLines 1
		Do While oExec.StdErr.AtEndOfStream <> True
			'可加入删除符，解压状态保留在同一行
			ReadLine = oExec.StdErr.ReadLine
			ShowLog ReadLine
			If RegExpTest(ReadLine,"^ERRORS?:") Then Unzip_7z = False
		loop
		ShowLog "开始图片转换"
		'得到标准输出流
'			strOut = oExec.StdOut.ReadAll()
		'得到标准错误流
'			strErr = oExec.StdErr.ReadAll()
'			ShowLog "strOut--" & strOut
'			If Len(strErr) > 0 Then ShowLog "解压出错" & vbCrLf & strErr

		'开始图片转换
		'*****************
		Dim PicList,pi
		'存入图片列表
		PicList = Split(TempList,vbCrLf)
		Dim WriteLine
		For pi = 0 To UBound(PicList)
			PicList(pi) = TempflName & "\" & PicList(pi)
			If fso.FileExists(PicList(pi)) Then
				WScript.StdOut.Write CharRepeat(Chr(8),LenB(WriteLine))
				'WriteLine = "转换当前压缩包内第" & vbTab & (fi - UBound(PicList) + pi + 1) & "/" & Items.Count & vbTab & "个文件。"
				WriteLine = (fi - UBound(PicList) + pi + 1) & "/" & Items.Count
				WScript.StdOut.Write WriteLine
				ConvertImage(PicList(pi))
			End If
		Next
		WScript.StdOut.Write CharRepeat(Chr(8),LenB(WriteLine))
		ShowLog WriteLine
		
		'开始移动回去
		'*****************
		'fso只支持同盘移动，需先复制再删除
		If targetLong Then
			ShowLog "由于目标路径过长，目标路径将缩减至：" & vbCrLf & target & vbCrLf & "如缩减后仍然过长，本程序将会因为出错停止运行"
		End If
		ShowLog "从临时文件夹移动回来"
		fso.CopyFolder TempflName, target
		fso.DeleteFolder TempflName,True '删除解压文件
		fso.DeleteFile TempTextName,True '删除临时文件列表
	Loop
	
	fso.DeleteFile AllListTextName,True '删除完整文件列表
	If DeleteZip Then fso.DeleteFile FilePath,True '删除压缩包
	
	'删除临时所在文件夹
	'fso.DeleteFolder(TempPath),True
End Function

'建立多层文件夹
Function CreateFolder(path)
	CreateFolder = True
	If fso.FolderExists(path) Then
		Exit function
	ElseIf Not fso.DriveExists(fso.GetDriveName(path)) Then
		CreateFolder = False
		Exit function
	End If
	If Not fso.FolderExists(fso.GetParentFolderName(path)) Then
		CreateFolder fso.GetParentFolderName(path)   
	End If
	fso.CreateFolder(path)
	CreateFolder = True
End Function

'去除两端空格
Function DelSpace(str) 
	Dim search      ' 创建变量。
	Set search = RegExpSearch(str, "^\s*(\S+|\S.+\S)\s*$")
	If search.count > 0 Then
		DelSpace = search(0).SubMatches(0)
	Else
		DelSpace = str
	End if
End Function

'是否符合正则表达式
Function RegExpTest(strng, patrn) 
	Dim regEx      ' 创建变量。
	Set regEx = New RegExp         ' 创建正则表达式。
	regEx.Pattern = patrn         ' 设置模式。
	regEx.IgnoreCase = True         ' 设置是否区分大小写，True为不区分。
	regEx.Global = True         ' 设置全程匹配。
	regEx.MultiLine = True
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
	regEx.MultiLine = True
	Set RegExpSearch  = regEx.Execute(strng)
'	If RegExpSearch.Count > 0 Then
'		MsgBox RegExpSearch.Item(0)
'		If RegExpSearch.Item(0).Submatches.Count > 0 Then
'			Set SubMatches = RegExpSearch.Item(0).Submatches
'			MsgBox SubMatches.Item(0)
'		End If
'	End If
	Set regEx = Nothing
End Function
'====================================
'主代码
'====================================
Dim Files,PicList,ErrNum
Set Files = WScript.Arguments '将参数（文件列表）存入类
Dim FirstFileType
ShowLog "==========================================" & vbCrLf _
	& Date() & " " & Time() & " start" & vbCrLf _
	& Appname & " V" & Version & vbCrLf
ErrNum = 0
If fso.FileExists(Files(0)) Then
	FirstFileType = fso.GetExtensionName(Files(0))
	Mode = ReturnMode(FirstFileType)
	
	If Mode = 1 Then 'Text模式
		ShowLog "当前为图片列表模式"
		For ii = 0 To Files.Count - 1
			'存入图片列表
			PicList = Split(LoadText(Files(ii),Charset),vbCrLf)
			For fi = 0 To UBound(PicList)
				ShowLog "当前第" & vbTab & fi+1 & "/" & UBound(PicList)+1 & vbTab & "个文件，当前列表" & vbTab & ii+1 & "/" & Files.Count & vbTab & fso.GetFileName(Files(ii))
				If fso.FileExists(PicList(fi)) Then
					ConvertImage(PicList(fi))
				Else
					If Len(PicList(fi)) > 0 Then
						ShowLog "错误：图片" & vbTab & fi+1 & "/" & UBound(PicList)+1 & vbTab & PicList(fi) & vbTab & "文件不存在。"
					Else
						ShowLog "跳过：空行。"
					End If
				End If
			Next
		Next
	ElseIf Mode = 2 Then '直接放上图片文件模式
		ShowLog "当前为直接图片模式"
		For fi = 0 To Files.Count - 1
			ShowLog "当前第" & vbTab & fi+1 & "/" & Files.Count & vbTab & "个文件。"
			If fso.FileExists(Files(fi)) Then
				ConvertImage(Files(fi))
			Else
				If Len(Files(fi)) > 0 Then
					ShowLog "错误：图片" & vbTab & fi+1 & "/" & Files.Count & vbTab & Files(fi) & vbTab & "文件不存在。"
				End If
			End If
		Next
	ElseIf Mode = 3 Or Mode = 4 Then '压缩包模式
		ShowLog "当前为压缩包模式"
		
		If Not fso.DriveExists(fso.GetDriveName(TempPath)) then
			ShowLog "临时文件夹驱动器不存在，无法解压。"
		Else
		'开始解压
			Dim FileType,ZipMode
			For ii = 0 To Files.Count - 1
				FileType = fso.GetExtensionName(Files(ii))
				ZipMode = ReturnMode(FileType)
				ShowLog vbCrLf & "当前第" & vbTab & ii+1 & "/" & Files.Count & vbTab & "个文件，为" & vbTab & FileType & vbTab & "扩展名。"
				If ZipMode = 3 Or ZipMode = 4 Then
					'7-zip模式
					ShowLog "采用 7-ZIP 解压。"
					If Not Unzip_7z(Files(ii)) Then ErrNum=ErrNum+1
				ElseIf ZipMode = 4 Then
					'RAR模式
					ShowLog "采用 WinRAR 解压。"
				Else
					ShowLog "错误：未识别的压缩包格式。"
				End If
			Next
		
		End If
		ShowLog vbCrLf & "压缩包总错误：" & ErrNum & vbTab & "次"
	Else
		ShowLog "错误：未识别的首文件格式"
	End If
Else
	ShowLog "错误：首文件不存在。"
End If

ShowLog vbCrLf & Date() & " " & Time() & " end" & vbCrLf  _
	& "==========================================" & vbCrLf 
Set fso=Nothing
WScript.Quit
