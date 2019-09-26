<%@  language="VBScript" %>
<%
'∑¢ÀÕ∂Ã–≈
Function SendSms(sname,spwd,smobiles,scontent)
    dim ht,pos
    set ht=Server.createobject("Microsoft.XMLHTTP") 
    ht.open "GET","http://api.sms1086.com/api/Send.aspx?username="&Server.URLEncode(sname)&"&password="&Server.URLEncode(spwd)&"&mobiles="&smobiles&"&content="&Server.URLEncode(scontent)&"",false
    ht.setRequestHeader "Content-type:", "text/xml;charset=GB2312"
    ht.send
    ret=urldecode(ht.ResponseText)
    'response.Write ret
    'response.End
    set ht=nothing
    pos = InStr(ret,"result=0")
    if pos > 0 then
        SendSms = true
    else
        SendSms = false
    end if
End Function

'–ﬁ∏ƒ√‹¬Î
function chgpwd(sname,spwd,new_password)
    dim ht
    set ht=Server.createobject("Microsoft.XMLHTTP") 
    ht.open "GET","http://api.sms1086.com/api/Send.aspx?username="&Server.URLEncode(sname)&"&password="&Server.URLEncode(spwd)&"&newpws="&Server.URLEncode(new_password)&"",false
    ht.setRequestHeader "Content-type:", "text/xml;charset=GB2312"
    ht.send
    ret=urldecode(ht.ResponseText)
    set ht=nothing
    chgpwd=true
end function

'≤È—Ø”‡∂Ó
function Query(sname,spwd)
    dim ht
    set ht=Server.createobject("Microsoft.XMLHTTP") 
    ht.open "GET","http://api.sms1086.com/api/Send.aspx?username="&Server.URLEncode(sname)&"&password="&Server.URLEncode(spwd)&"",false
    ht.setRequestHeader "Content-type:", "text/xml;charset=GB2312"
    ht.send
    ret=urldecode(ht.ResponseText)
    set ht=nothing
    Query=true
end function

'url±‡¬Î
function urldecode(encodestr)    
newstr=""    
havechar=false    
lastchar=""    
for i=1 to len(encodestr)    
char_c=mid(encodestr,i,1)    
if char_c="+" then    
newstr=newstr & " "    
elseif char_c="%" then    
next_1_c=mid(encodestr,i+1,2)    
next_1_num=cint("&H" & next_1_c)    
if havechar then    
havechar=false    
newstr=newstr & chr(cint("&H" & lastchar & next_1_c))    
else    
if abs(next_1_num)<=127 then    
newstr=newstr & chr(next_1_num)    
else    
havechar=true    
lastchar=next_1_c    
end if    
end if    
i=i+2    
else    
newstr=newstr & char_c    
end if    
next    
urldecode=newstr    
end function   
function urldecode(encodestr) 
newstr="" 
havechar=false 
lastchar="" 
for i=1 to len(encodestr) 
char_c=mid(encodestr,i,1) 
if char_c="+" then 
newstr=newstr & " " 
elseif char_c="%" then 
next_1_c=mid(encodestr,i+1,2) 
next_1_num=cint("&H" & next_1_c) 
if havechar then 
havechar=false 
newstr=newstr & chr(cint("&H" & lastchar & next_1_c)) 
else 
if abs(next_1_num)<=127 then 
newstr=newstr & chr(next_1_num) 
else 
havechar=true 
lastchar=next_1_c 
end if 
end if 
i=i+2 
else 
newstr=newstr & char_c 
end if 
next 
urldecode=newstr 
end function 

  dim sname,spwd,smobiles,scontent
  sname=trim(request.form("username")) 
  spwd =trim(request.form("pwd"))  
  smobiles=trim(request.form("mobiles"))  
  scontent=trim(request.form("contents")) 

if SendSms(sname,spwd,smobiles,scontent)then 
   response.write "∑¢ÀÕ≥…π¶"
else
    response.write "∑¢ÀÕ ß∞‹"
end if

%>