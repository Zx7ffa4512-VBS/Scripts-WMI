'------------------------------------------------------------------------
'http://www.wayaa.com.cn/post/computer/62/62/62.html
'------------------------------------------------------------------------
strComputer = "."
strNameSpace = "root\cimv2"
WScript.Echo String(20,"*") & "本工具用于查询类限定信息" & String(20,"*")
WScript.StdOut.Write "计算机:" & strComputer & vbCrLf & "命名空间:" & strNameSpace & vbCrLf & vbCrLf & "请输入你要查询的类:"
strClass=WScript.StdIn.Readline()
Set objClass = GetObject("winmgmts:\\" & strComputer & "\" & strNameSpace & ":" & strClass)

WScript.Echo strClass & " 的类限定信息如下："
WScript.Echo "------------------------------"
i = 1
For Each objClassQualifier In objClass.Qualifiers_
	If VarType(objClassQualifier.Value) = (vbVariant + vbArray) Then '常数 VBVariant 只与 VBArray 一起返回，以表明 VarType 函数的参数是一个 Variant 类型的数组。
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
WScript.Echo strClass & " 类的属性和属性限定信息"
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
WScript.Echo strClass & " 类的方法和方法限定信息"
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