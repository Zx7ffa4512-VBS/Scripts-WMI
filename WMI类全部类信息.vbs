On Error Resume Next
'���������ռ�
'************************************************************************************
'Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
'------------------------------------------------------------------------------------
Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set objSWbemServices = objSWbemLocator.ConnectServer()   '�����������ӱ���
'************************************************************************************

Dim fso,File,Line
Set fso = WScript.CreateObject("Scripting.Filesystemobject")
Set File=fso.OpenTextFile("ClassName.txt",1,True)
Do Until File.AtEndOfStream
	Line=File.ReadLine
	ClassName=Line
	Set WMIObjectSet=objSWbemServices.get(ClassName) 				'�õ�һ������ʵ��
	Set WMIObjectSets=objSWbemServices.InstancesOf(ClassName)		'�õ����е�ʵ��
	
	WriteFile Line,"����:"
		For Each m In WMIObjectSet.methods_
			WriteFile Line, m.Name
		Next 
	WriteFile Line, String(79,"-")
	
	For Each ins In WMIObjectSets
		For Each f In WMIObjectSet.Properties_
			WriteFile Line, StringN(30,f.Name,-1) & ":" & Eval("ins." & f.Name)
		Next
		WriteFile Line, String(79,"*")
	Next 
	WScript.Echo Line
Loop


Function WriteFile(FileStr,DataStr)
	Dim File
	Set File=fso.OpenTextFile(FileStr & ".txt",8,True)
	File.WriteLine DataStr
	File.Close
End Function 


'------------------------------------------------------------------------
'����ָ�����ȣ��������ȷ���ԭ�ַ���,LeftCenterRight,-1�����,0����,1�Ҷ���
'------------------------------------------------------------------------
Function StringN(Num,Str,LeftCenterRight)
	If LenEx(Str)<Num Then
		Select Case LeftCenterRight
			Case -1 
				StringN=Str & String(Num-Len(Str)," ")
			Case 0
				Dim nYushu,nShang
				nShang=(Num-Len(Str))/2
				nYushu=(Num-Len(Str)) Mod 2
				If nYushu = 0 Then 
					StringN=String(nShang," ") & Str & String(nShang," ")
				Else
					StringN=String(nShang-0.5," ") & Str & String(nShang+0.5," ")
				End If 
			Case 1
				StringN=String(Num-Len(Str)," ") & Str
		End Select
	Else
		StringN=Str
	End If 
End Function

'------------------------------------------------------------------------
'��ȡ�ַ�������,������Ч,len("����")=2,LenEx("����")=4
'------------------------------------------------------------------------
Function LenEx(Str)
    Dim singleStr,i,iCount
    iCount = 0
    For i = 1 To Len(Str)
        singleStr = Mid(Str,i,1)
        If Asc(singleStr) < 0 Then
            iCount = iCount + 2
        Else 
            iCount = iCount + 1
        End If   
    Next
    LenEx = iCount
End Function

