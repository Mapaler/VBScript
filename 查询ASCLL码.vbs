Do
	s=InputBox("ÊäÈëÄãÒª²éÑ¯µÄ×Ö·û´®","ASCLLÂë²éÑ¯")
	res = ""
	For i = 1 To Len(s)
		w=Mid(s,i,1)
		res= res & w & "(" & Hex(Asc(w)) & ":" & Asc(w) & ") "
	Next
	If Len(s)>0 Then MsgBox res
Loop Until(Len(s)<1)