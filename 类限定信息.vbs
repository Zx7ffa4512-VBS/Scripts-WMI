'------------------------------------------------------------------------
'http://www.wayaa.com.cn/post/computer/62/62/62.html
'------------------------------------------------------------------------
strComputer = "."
strNameSpace = "root\cimv2"
WScript.Echo String(20,"*") & "���������ڲ�ѯ���޶���Ϣ" & String(20,"*")
WScript.StdOut.Write "�����:" & strComputer & vbCrLf & "�����ռ�:" & strNameSpace & vbCrLf & vbCrLf & "��������Ҫ��ѯ����:"
strClass=WScript.StdIn.Readline()
Set objClass = GetObject("winmgmts:\\" & strComputer & "\" & strNameSpace & ":" & strClass)

WScript.Echo strClass & " �����޶���Ϣ���£�"
WScript.Echo "------------------------------"
i = 1
For Each objClassQualifier In objClass.Qualifiers_
	If VarType(objClassQualifier.Value) = (vbVariant + vbArray) Then '���� VBVariant ֻ�� VBArray һ�𷵻أ��Ա��� VarType �����Ĳ�����һ�� Variant ���͵����顣
		strQualifier = i & ". " & objClassQualifier.Name & " = " & _
		Join(objClassQualifier.Value, ",")
	Else
		strQualifier = i & ". " & objClassQualifier.Name & " = " & _
		objClassQualifier.Value
	End If
	WScript.Echo strQualifier
	strQualifier = ""
	i = i + 1
Next

WScript.Echo
WScript.Echo strClass & " ������Ժ������޶���Ϣ"
WScript.Echo "-------------------------------------------------"
i = 1 : j = 1
For Each objClassProperty In objClass.Properties_
	WScript.Echo i & ". " & objClassProperty.Name
	For Each objPropertyQualifier In objClassProperty.Qualifiers_
		If VarType(objPropertyQualifier.Value) = (vbVariant + vbArray) Then
			strQualifier = i & "." & j & ". " & _
			objPropertyQualifier.Name & " = " & _
			Join(objPropertyQualifier.Value, ",")
		Else
			strQualifier = i & "." & j & ". " & _
			objPropertyQualifier.Name & " = " & _
			objPropertyQualifier.Value
		End If
		WScript.Echo strQualifier
		strQualifier = ""
		j = j + 1
	Next
	WScript.Echo
	i = i + 1 : j = 1
Next

WScript.Echo
WScript.Echo strClass & " ��ķ����ͷ����޶���Ϣ"
WScript.Echo "-------------------------------------------------"
i = 1 : j = 1
For Each objClassMethod In objClass.Methods_
	WScript.Echo i & ". " & objClassMethod.Name
	For Each objMethodQualifier In objClassMethod.Qualifiers_
		If VarType(objMethodQualifier.Value) = (vbVariant + vbArray) Then
			strQualifier = i & "." & j & ". " & _
			objMethodQualifier.Name & " = " & _
			Join(objMethodQualifier.Value, ",")
		Else
			strQualifier = i & "." & j & ". " & _
			objMethodQualifier.Name & " = " & _
			objMethodQualifier.Value
		End If
		WScript.Echo strQualifier
		strQualifier = ""
		j = j + 1
	Next
	WScript.Echo
	i = i + 1 : j = 1
Next