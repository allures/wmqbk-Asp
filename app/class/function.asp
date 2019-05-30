<%
'数据库连接
Public Function OpenConn()
	'On Error Resume Next
	If Conn<>Empty Then Exit Function
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath(rPath&DBPath) & ";"
	Conn.CommandTimeout = 30
	Conn.ConnectionTimeout = 30
	Conn.Open()
	If Err.Number <> 0 Then
    Err.Clear
    Conn.Close
    Set Conn = Nothing
    Response.Write "对不起,数据库连接出错,请与管理员联系!"&Err.Number
    Response.End
    End If
End Function
'关闭数据库
Public Function CloseConn()
    On Error Resume Next
    If IsObject(Conn) Then
        Conn.Close
        Set Conn = Nothing
    End If
End Function
'转换html标签
Function Clstr(ByVal str)
     If str = "" Then Exit Function
		str = replace(str, ">", "&gt;")
		str = replace(str, "<", "&lt;")
		str = Replace(str, CHR(32), " ")		
		str = Replace(str, CHR(34), "&quot;")
		str = Replace(str, CHR(39), "&#39;")
		str = Replace(str, CHR(9), "&nbsp;")
		str = Replace(str, CHR(13), "")
		str = Replace(str, CHR(10) & CHR(10), " ")
		str = Replace(str, CHR(10), " ")
     Clstr = str
End Function 
'转换html标签
Function fStr(ByVal str)
       If str = "" Then Exit Function
       str = replace(str, "{", "")
	   str = replace(str, "}", "")
	   str = Replace(str, CHR(39), "&#39;")
	   fStr = str
End Function
'移除html标签
Function reHtml(ByVal str)
     Dim re
	 Set re=new RegExp   
	 re.IgnoreCase =true   
	 re.Global=True   
	 re.Pattern="<\/*[^<>]*>"   
	 str = re.replace(str," ") 
	 Set re=nothing
	 reHtml = str	 
End Function
'数字格式化
Function fNum(ByVal num)
      If isNumeric(num) Then 
	    fNum = CLng(num)
	  Else
	    fNum = 0
	  End If 
End Function
'getDesc
Function getSumm(ByVal str,ByVal p)
 If p <> "" Then 
  getSumm = "<span style=""color:red;"">这是一篇密码日志！</span>"
  Exit Function
 End If
 Dim re
 Set re=new RegExp   
 re.IgnoreCase =true   
 re.Global=True   
 re.Pattern="(<(?!/?a(\s|>))[^>]*>|[\n\t\r|])"   
 str = re.replace(str," ")   
 str = Left(str,150) 
 getSumm = str
 Set re=nothing
End Function
'最新评论输出
Function NewCom(num)
		Dim ComList,plurl
		ComList=Application("Com_List"&Flag) 
        If Not IsArray(ComList) Then
		Application.Lock()
		Set Rs=Conn.Execute("Select top "&num&" pid,cid,pname,pcontent,rcontent from [pl] where isn=0 order by pid desc")
		If Rs.Eof Then 
		Rs.Close
		NewCom = "<li>暂无评论！</li>"
		Exit Function
		End If
		Application("Com_List"&Flag)=Rs.GetRows
		Rs.Close
		Application.UnLock()
		ComList=Application("Com_List"&Flag)
		End If
		For i=0 to Ubound(ComList,2)
			If Rewrite=1 Then
				plurl="post-"&ComList(1,i)&".html"
			Else
				plurl=iFile&"?act=pl&id="&ComList(1,i)
			End If
		   NewCom =NewCom & "<li id=""Com-"&ComList(0,i)&"""><b>"&Clstr(ComList(2,i))&"：</b><a href="""&plurl&"#Com-"&ComList(0,i)&""">"&Clstr(ComList(3,i))&"</a>"  
		   If ComList(4,i)<>"" Then NewCom = NewCom & " <img width=""10"" title=""已回复"" height=""10"" src="""&TEMPLATE_PATH&"/style/reply.gif"" />"
		  NewCom =NewCom & "</li>"
		Next
End Function
'热门文章
Function TopIc(num)
		Dim TopList,title,plurl
		TopList=Application("Top_List"&Flag) 
        If Not IsArray(TopList) Then
		Application.Lock()
		Set Rs=Conn.Execute("Select top "&num&" id,title,summ,num from [Log] where hide=0 order by num desc,atime desc")
		If Rs.Eof Then 
		Rs.Close
		TopIc = "<li>暂无评论！</li>"
		Exit Function
		End If
		Application("Top_List"&Flag)=Rs.GetRows
		Rs.Close
		Application.UnLock()
		TopList=Application("Top_List"&Flag)
		End If
		For i=0 to Ubound(TopList,2)
			If Rewrite=1 Then
				plurl="post-"&TopList(0,i)&".html"
			Else
				plurl=iFile&"?act=pl&id="&TopList(0,i)
			End If
			title = TopList(1,i)
			'Response.write (title="")
		   If title="" Then title = Left(reHtml(TopList(2,i)),16)
		   TopIc =TopIc & "<li><a href="""&plurl&""">"&title&"</a> ["&TopList(3,i)&"]"		    
		   TopIc =TopIc & "</li>"
		Next
		'Response.write TopIc
End Function
'评论框
Function pFrom(id)
	Dim I1,I2
	If admin = 1 Then 
	  I1=WebUser
	Else
	  I1=Request.Cookies("4jax-nick")
	End If
	If SafeCode = 1 Then
	  I2=" onfocus=""$('#codep').show();"""
	End If
	pFrom  = "<a name=""pl""></a>"&_
	 "<p><input name=""pname"" tabindex=""1"" placeholder=""您的昵称"" id=""pname"" type=""text"" class=""log"" value="""&I1&""" maxlength=""10"" /></p>"&_
	"<p><textarea tabindex=""2"" placeholder=""随便说点什么吧..."" name=""plog"" rows=""3"" id=""plog"" class=""log"""&I2&"></textarea></p><p id=""codep""><input type=""text"" id=""safecode"" placeholder=""右侧计算答案"" name=""safecode"" autocomplete=""off"" class=""log"" value=""""/> <img src=""libs/class/codes.asp"" id=""codeimg"" style=""cursor:pointer"" alt=""更换一道题！"" onclick=""reloadcode()""/></p>"&_
	 "<p><button name=""add"" onClick=""addpl('"&id&"','"&SafeCode&"');"" id=""add"" class=""btn""> 提 交 </button> "&_
	 "<button name=""bck"" onClick=""history.back();"" id=""bck"" class=""btn""> 返 回 </button></p>" 
End Function
'登录
Function doLogin()
    If Pass = Request.Form("pass") Then
		Response.Cookies("4jax-"&Flag)=pass
		Response.Cookies("4jax-"&Flag).Expires = Date()+7
		'Response.Cookies("4jax-"&Flag).domain=Domain
    End If
	Response.Redirect("./")
	Response.End()	
End Function
'退出登录
Function logOut()
    Response.Cookies("4jax-"&Flag)=""		
	Response.Cookies("4jax-"&Flag).Expires = Now()-1
	'Response.Cookies("4jax-"&Flag).domain=Domain
	Response.Redirect("./")
	Response.End()	
End Function
'检查权限
Function isLogin()
    If admin = 0  Then	 
		Call Message("500","无权限操作","")
    End If
End Function
'消息提示
Function Message(result,msg,redirect)
   If redirect = "" Then 
      msg = replace(msg,"""","\""")
      Response.Write "{""result"":"""&result&""",""message"":"""&msg&"""}"
	  Response.End()
   Else 
      Response.Redirect(redirect)
	  Response.End()
   End If
End Function
'判断客户端
Function isMobile()
Dim reg,UA
Set reg = New RegExp
UA = Request.ServerVariables("HTTP_USER_AGENT")
reg.pattern=".+?(iphone|ipad|ipod|android).*"
reg.IgnoreCase = True
If reg.test(UA) Then
  isMobile = True 
  Else
  isMobile = False 
End If 
Set reg = Nothing
End Function 

Function DeleteFiles(FilePath) 
    Dim oFilePath
	On Error Resume Next
    oFilePath=Server.MapPath(Replace(FilePath,"..",""))    
    Dim FSO
    Set FSO = Server.CreateObject("Scripting.FileSystemObject")
    If FSO.FileExists(oFilePath) Then
        FSO.DeleteFile oFilePath, True
        If Err Then
            Set FSO = Nothing
            Err.Clear
            Exit Function
        End If
    End If
    Set FSO = Nothing
End Function


Function vurl(id)
  If Rewrite = 1 Then
     vurl = "post-"&id&".html"
  Else 
     vurl = iFile&"?act=pl&id="&id
  End if 
End Function

Function vmenu(menu)
  If Rewrite = 1 Then
     menu = Replace(menu,"@index","index.html")
	 menu = Replace(menu,"@comment","comment.html")
  Else 
     menu = Replace(menu,"@index",iFile)
	 menu = Replace(menu,"@comment",iFile&"?act=plist")
  End If
  vmenu = menu
End Function
%>