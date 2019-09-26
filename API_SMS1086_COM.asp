<%
'====================================
'��������
'====================================
Const smsUsername = "xxx"
Const smsPassword = "xxx"

Dim SMS1086, Result
	Set SMS1086 = New API_SMS1086_COM
	SMS1086.Username = smsUsername
	SMS1086.Password = smsPassword
	SMS1086.Mobiles = "13602020202"
	SMS1086.Content = "����һ��"
	SMS1086.f = "1"
	
	SMS1086.SendEnc()
	
Response.Write("������Ϊ:" & SMS1086.GetInfo("result") & "<br>")
Response.Write("��������Ϊ:" & SMS1086.GetInfo("description") & "<br>")
Response.Write("����ʧ�ܺ����б�Ϊ:" & SMS1086.GetInfo("faillist") & "<br>")

'====================================
'====================================
Class API_SMS1086_COM
	
	Public SendUrl
	Public Username
	Public Password
	Public Mobiles
	Public Content
	Public f
	Private LastInfo
	
	
	'��ʼ��
	Private Sub Class_Initialize()
		Call Initialize()
	End Sub
	'����
	Private Sub Class_Terminate()
		Call Initialize()
	End Sub
	'��ʼ������
	Private Sub Initialize()
		SendUrl = "http://api.sms1086.com/api/Send.aspx"
		Username = ""
		Password = ""
		Mobiles = ""
		Content = ""
		f = ""
		LastInfo = ""
	End Sub
	
	'���ð�����תΪUrlCode
	Private Function URLEncoding(Byval vStrIn) 
		Dim strReturn : strReturn = ""
		Dim ThisChr, innerCode, Hight8, Low8
		For i = 1 To Len(vStrIn)
			ThisChr = Mid(vStrIn, i, 1) 
			If Abs(Asc(ThisChr)) < &HFF Then
				strReturn = strReturn & ThisChr
			Else
				innerCode = Asc(ThisChr) 
				If innerCode < 0 Then 
					innerCode = innerCode + &H10000 
				End If
				Hight8 = (innerCode And &HFF00)\ &HFF
				Low8 = innerCode And &HFF
				strReturn = strReturn & "%" & Hex(Hight8) & "%" & Hex(Low8)
			End If
		Next
		URLEncoding = strReturn
	End Function
	
	Private Function URLDecode(enStr) 
		Dim deStr,strSpecial 
		Dim c,i,v
		deStr = "" 
		strSpecial = "!""#$%&'()*+,.-_/:;<=>?@[\]^`{|}~%" 
		For i = 1 To Len(enStr) 
			c = Mid(enStr,i,1)
			If c = "%" Then
   				v = Eval("&h" + Mid(enStr, i + 1, 2)) 
   				If InStr(strSpecial, Chr(v)) > 0 Then 
    				deStr = deStr & Chr(v) 
    				i = i + 2 
   				Else 
    				v = Eval("&h" + Mid(enStr, i + 1, 2) + Mid(enStr, i + 4, 2)) 
    				deStr = deStr & Chr(v) 
    				i = i + 5 
     			End If 
			Else 
   				If c = "+" Then 
    				deStr = deStr & " " 
   				Else 
    				deStr = deStr & c 
   				End If 
			End If 
		Next 
		URLDecode = deStr
	End Function
	
	'XMLHTTPץȡ
	Private Function HttpGet(Byval StrUrl) 
		On Error Resume Next
		Set Http = Server.CreateObject("Microsoft.XMLHTTP")
		Http.Open "GET",StrUrl, False
		Http.setRequestHeader "Content-type:", "text/xml;charset=GB2312"
		Http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 
		Http.Send()
		HttpGet = Http.ResponseText
		Set Http = Nothing
		If Err Then
			Err.Clear
		End If
	End Function
	
	' ��ʽ��ʱ��(��ʾ)
	' ������n_Flag
	' 1:"yyyy-mm-dd hh:mm:ss"
	' 2:"yyyy-mm-dd"
	' 3:"hh:mm:ss"
	' 4:"yyyy��mm��dd��"
	' 5:"yyyymmdd"
	' 6:"yy-mm-dd"
	' 7:mm-dd hh:mm
	' 8:mm-dd 
	Private Function FormatTime(Byval s_Time, Byval n_Flag)
		Dim y, m, d, h, mi, s
		FormatTime = ""
		If IsDate(s_Time) = False Then Exit Function
		y = cstr(year(s_Time))
		m = cstr(month(s_Time))
		If len(m) = 1 And n_Flag<>8 Then m = "0" & m
		d = cstr(day(s_Time))
		If len(d) = 1 And n_Flag<>8 Then d = "0" & d
		h = cstr(hour(s_Time))
		If len(h) = 1 Then h = "0" & h
		mi = cstr(minute(s_Time))
		If len(mi) = 1 Then mi = "0" & mi
		s = cstr(second(s_Time))
		If len(s) = 1 Then s = "0" & s
		Select Case n_Flag
			Case 1
				' yyyy-mm-dd hh:mm:ss
				FormatTime = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s
			Case 2
				' yyyy-mm-dd
				FormatTime = y & "-" & m & "-" & d
			Case 3
				' hh:mm:ss
				FormatTime = h & ":" & mi & ":" & s
			Case 4
				' yyyy��mm��dd��
				FormatTime = y & "��" & m & "��" & d & "��"
			Case 5
				' yyyymmdd
				FormatTime = y & m & d
			Case 6
				If Len(y)=4 Then y=Right(y,2)
				' yyyy-mm-dd
				FormatTime = y & "-" & m & "-" & d
			Case 7
				' mm-dd hh:mm
				FormatTime = m & "-" & d & " " & h & ":" & mi
			Case 8
				' mm-dd
				FormatTime = m & "-" & d
		End Select
	End Function
	
	'�ַ���תByte����
	Private Function StringToBytes(Byval varString)
		 Dim i, strLen
		 strLen = Len(varString)
		 Dim byStr()
		 For i = 1 To strLen
			  ReDim Preserve byStr(i - 1)
			  byStr(i - 1) = CByte(Asc(Mid(varString, i, 1)))
		 Next
		 StringToBytes = byStr
	End Function
	
	'Byte����ת�ַ���
	Private Function BytesToString(Byval Bytes)
		Dim Result : Result = ""
		Dim i
		For i = 0 To UBound(Bytes)
			Result = Result & Chr(Bytes(i))
		Next
		BytesToString = Result
	End Function
	
	'������
	'RefBytes ���
	Private Sub CalcXor(Byval Bytes, ByRef RefBytes)
		Dim i, k : k = 0
		For i = 0 To UBound(RefBytes)
			RefBytes(i) = (RefBytes(i) Xor Bytes(k))
			k = k + 1
			If k = (UBound(Bytes) + 1) Then
				k = 0
			End If
		Next
	End Sub

	'��̬���뷢�Ͷ���
	'����cls_Items����
	Public Sub SendEnc()
		Dim vResponse	: vResponse		= ""
		Dim UrlString	: UrlString		= SendUrl
		Dim strResult	: strResult		= ""
		Dim strPass		: strPass		= ""
		Dim byUsername	: byUsername	= StringToBytes(Username)
		Dim byPassword	: byPassword	= StringToBytes(Password)
		Dim byBytes		: byBytes		= StringToBytes(FormatTime(Now, 1))

		Call CalcXor(byUsername, byBytes)
		vResponse = BytesToString(byBytes)
		
		Call CalcXor(byBytes, byPassword)
		strPass = BytesToString(byPassword)
		
		UrlString = UrlString & "?username=" & URLEncoding(Username)
		UrlString = UrlString & "&password=" & URLEncoding(strPass)
		UrlString = UrlString & "&mobiles=" & URLEncoding(Mobiles)
		UrlString = UrlString & "&content=" & URLEncoding(Content)
		UrlString = UrlString & "&f=" & URLEncoding(f)
		UrlString = UrlString & "&rp=" & URLEncoding(vResponse)
		
		strResult = HttpGet(UrlString)
		LastInfo = strResult
	End Sub
	
	Public Sub ChangePassword()
		
	End Sub
	
	Public Function GetInfo(Byval ParamName)
		Dim Result : Result = ""
		If LastInfo = "" Then
			GetInfo = ""
			Exit Function
		ElseIf InStr(LastInfo, "&") <= 0 Then
			GetInfo = ""
			Exit Function
		End If
		Dim Arys : Arys = Split(LastInfo, "&")
		Dim Strs : Strs = ""
		Dim StrAry
		For Each Strs In Arys
			If InStr(Strs, "=") > 0 Then
				StrAry = Split(Strs, "=")
			Else
				StrAry = Split(Strs & "=", "=")
			End If
			If LCase(StrAry(0)) = LCase(ParamName) Then
				Result = StrAry(1)
				Exit For
			End If
		Next
		GetInfo = URLDecode(Result)
	End Function
	
End Class
%>