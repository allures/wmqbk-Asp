<!--#include file="config.asp"-->
<!--#include file="function.asp"-->
<%
Response.Buffer = True
Dim I8,I9,I0,action,id,pic
action = Request.Querystring("act")
id = fNum(Request.Querystring("id"))
pic = Request.Querystring("pic")
I8=Cstr(Request.ServerVariables("HTTP_REFERER"))
I9=Cstr(Request.ServerVariables("SERVER_NAME"))
If Mid(I8,8,len(I9))<>I9 then
   Call Message("500","禁止直接访问","")
End If
I0=",delpic,ckcode,add,login"
If Instr(I0,action)=0 Then 
   Call OpenConn() '按需链接数据库   
End If
'审核 删除 置顶 添加评论
Select Case action
Case "login"
        Call lfrom()
	Case "add"
		Call afrom()
	Case "savelog"
		Call saveLog()
	Case "saveset"
	    Call saveSet()
	Case "savewid"
	    Call saveWid()
	Case "delwid"
	    Call delWid()
	Case "editsave"
		Call editsave()				
    Case "edit"
        Call edit()
        Call afrom()
    Case "dellog"
        Call delLog()
    Case "delpl"
        Call delpl()
    Case "delpic"
        Call delPic(pic)		
    Case "shpl"
        Call shPl()	
    Case "addpl"
        Call addpl()
    Case "editpl"
        Call editpl()
        Call pfrom()	
    Case "plsave"
        Call plsave()		
    Case "zdlog"
        Call zdLog()			
	Case "ckpass"	
		Call ckpass() 
	Case "ckcode"	
		Call ckcode() 		
	Case "upcache"	
		Call upCache()			
    Case Else
End Select
CloseConn() 
'删除日志
Function delLog()
    Call isLogin()
	'If id<5 Then Call Message("500","为了演示方便这部分日志被保护了起来！","")
	Dim I1,I2,A1
    Set Rs=Conn.Execute("select pic,pics from [Log] where id="&id)
	I1=Rs("pic")
	I2=Rs("pics")
	If I1<>"" Then DeleteFiles(rPath&I1)
	If I2<>"" Then 
      A1 = Split(I2,",")
	  For I = 0 to UBound(A1) 
          DeleteFiles(rPath&A1(I))
	  Next 
	End if
    Conn.Execute("delete from [Log] where id="&id)
	Application.Contents.Remove("Top_List"&Flag) '删除缓存
	Application.Contents.Remove("Com_List"&Flag) '删除缓存
    Call Message("200","删除日志成功","")
End Function
'日志添加/编辑
Function  saveLog()
    Call isLogin()
    Dim c
	c = Request.Form("c")
    If c = "add" Then
       Call addLog()
    ElseIf c = "edit"  Then
	   'If id<8 Then Call Message("500","为了演示方便这部分日志被保护了起来！","")
	   Call editLog()
	Else
	End If
End Function
'日志添加
Function  addLog()   
    Dim tit,desc,logs,pic,fm,pas,nid
	tit = fStr(Request.Form("tit"))
	summ = fStr(Request.Form("summ"))
    logs = fStr(Request.Form("logs"))
	pic = fStr(Request.Form("pic"))
	pics = fStr(Request.Form("pics"))
	pas = fStr(Request.Form("pass"))
	hide = fNum(Request.Form("hide"))
	lock = fNum(Request.Form("lock"))
	If isMobile Then 
		fm="手机"
	Else
		fm="网页"
	End if
    If logs<>"" Then
	   If Trim(reHtml(summ)) = "" Then  summ = getSumm(logs,pas)
	   If Trim(summ)="" Then summ = "#分享"
        Conn.Execute("insert into [Log](title,summ,content,fm,pic,pics,pass,hide,lock)values('"&tit&"','"&summ&"','"&logs&"','"&fm&"','"&pic&"','"&pics&"','"&pas&"',"&hide&","&lock&")") 
		'Response.Write "ok" 
		nid=Conn.Execute("select @@IDENTITY")(0)
		Call Message("200",nid,"")
    End If	
End Function
'日志编辑保存
Function editLog()    
    Dim tit,desc,logs,pic,opic,picu,atime,pas
    tit = fStr(Request.Form("tit"))
	summ = fStr(Request.Form("summ"))
    logs = fStr(Request.Form("logs"))
	pic = fStr(Request.Form("pic"))
	pics = fStr(Request.Form("pics"))
	atime = Request.Form("atime")
	pas = fStr(Request.Form("pass"))
	hide = fNum(Request.Form("hide"))
	lock = fNum(Request.Form("lock"))
	If IsDate(atime) = False Then atime= Now() 
    If logs<>"" Then
	    If Trim(reHtml(summ)) = "" Or pas<>"" Then  summ = getSumm(logs,pas)
        Conn.Execute("update [Log] set title='"&tit&"',summ='"&summ&"',content='"&logs&"',pic='"&pic&"',pics='"&pics&"',atime='"&atime&"',pass='"&pas&"',hide="&hide&",lock="&lock&" where id="&id)
    End If
	Application.Contents.Remove("Top_List"&Flag)
	Call Message("200",id,"")
End Function
'评论回复保存
Sub plsave()
    Call isLogin()
    Dim rlog,pid
	pid = CLng(Request.Querystring("pid"))
    rlog = replace(Request.Form("rlog"),"'","&#39;")   
    'pname = replace(Request.Form("pname"),"'","&#39;")
    If rlog<>"" Then
        Conn.Execute("update [Pl] set rcontent='"&rlog&"',isn=0 where pid="&pid)
    End If
	Application.Contents.Remove("Com_List"&Flag)
	'Application.Contents.Remove("Top_List"&Flag)	
	Call Message("200",rlog,"")
	'oStr= "<P><b>"&Clstr(pname)&"</b>："&Clstr(plog)	
	'If x=1 Then oStr=oStr&" <a href="""&plurl&"#p"&pid&""">[查看]</a></p>" Else oStr=oStr&"</p>" End If
	'If rlog<>"" Then oStr=oStr&"<p>&nbsp;&nbsp;<b style=""color:#C00"">回复</b>："&Clstr(rlog)&"</p>"
	'Response.Write oStr
End Sub
'置顶日志
Function zdLog()
    Call isLogin()
	Dim x
	x = CLng(Request.Querystring("x"))
    Conn.Execute("update [Log] set ist="&x&" where id="&id)
	If x=1 then	 
	  Call Message("200","取消","")
	Else	  
	  Call Message("200","置顶","")
	End If
End Function
'审核评论
Function shPl()
    Call isLogin()
	Dim pid
	pid = CLng(Request.Querystring("pid"))
    Conn.Execute("update [Pl] set isn=0 where pid="&pid)
	Application.Contents.Remove("Com_List"&Flag) '删除缓存
	Call Message("200","审核成功！","")
End Function
'删除评论
Function delpl()
    Call isLogin()
	Dim pid
	pid = CLng(Request.Querystring("pid"))
    Conn.Execute("delete from [pl] where pid="&pid)
	Conn.Execute("update [Log] set num=num-1 where id="&id)
	Application.Contents.Remove("Com_List"&Flag)
	Call Message("200","评论删除成功！","")
End Function
'添加评论
Sub addPl()
    Dim plog,pname,Temp,I1,I2,I3,ooRs
	Set ooRs = Conn.Execute("select id,hide,lock from [Log] where id=" & id)
	If (ooRs.Eof or ooRs("lock")=1 or ooRs("hide")=1) Then 
	   ooRs.Close()
	   Call Message("500","文章不存在或禁止评论！","")
	end if
	I2=Request.Form("scode")
	If IsNumeric(I2) Then I2=Clng(Request.Form("scode"))
	If Session("safecode")<>I2 And SafeCode="1" Then 
	  Call Message("500","验证码错误！","")
	End If
    plog = left(Request.Form("plog"),200)
	plog = Clstr(plog)
    pname = Clstr(Request.Form("pname"))	
    If admin = 0 Then 
	  pname=replace(pname,WebUser,"网友") '替换冒充
	  I1=PLsh
	Else
	  I1=0 '管理无须审核
	End If
    If plog<>"" And pname<>"" Then
        Conn.Execute("insert into [Pl](cid,pname,pcontent,isn)values("&id&",'"&pname&"','"&plog&"',"&I1&")")	    
		Conn.Execute("update [log] set num=num+1 where id="&id)
		Response.Cookies("4jax-nick") = pname
	    Response.Cookies("4jax-nick").Expires = Date()+7
    End If	
	I3 = "<div class=""comlist""><p><b>"&Clstr(pname)&"</b>："&Clstr(plog)&"</p><p class=""time"">"&Now()&"</p></div>"		
	Application.Contents.Remove("Com_List"&Flag)
	Application.Contents.Remove("Top_List"&Flag)
	Call Message("200",I3,"")
End Sub
'设置保存
Function saveSet()
    Call isLogin()
	'Call Message("200","设置更新成功！演示系统不更新","")
	Dim swebuser,swebtitle,swebdesc,sist,ssafecode,smenu,srewrite
	swebuser= fStr(Request.Form("webuser"))
	swebtitle= fStr(Request.Form("webtitle"))
	swebdesc= fStr(Request.Form("webdesc"))
	icp= fStr(Request.Form("icp"))
	splsh= CInt(Request.Form("plsh"))
	ssafecode= CInt(Request.Form("safecode"))
	smenu= fStr(Request.Form("menu"))
	srewrite= CInt(Request.Form("rewrite"))
	If swebuser<>"" And swebtitle<>"" And swebdesc<>"" Then 
		Conn.Execute("update [Set] set webuser='"&swebuser&"',webtitle='"&swebtitle&"',webdesc='"&swebdesc&"',icp='"&icp&"',plsh="&splsh&",safecode="&ssafecode&",menu='"&smenu&"',rewrite="&srewrite&" where id=1")
		'Conn.Execute("update [Set] set rewrite="&srewrite&" where id=1")
		Application.Contents.Remove("Web_Set"&Flag) '删除缓存
		Call Message("200","设置更新成功！","")
	Else
		Call Message("500","表单数据不完整！","")
	End If 
End Function
'边栏保存
Function saveWid() 
	Call isLogin()
	'Call Message("200","设置更新成功！演示系统不更新","")
	Dim title,html,ord
	title= fStr(Request.Form("title"))
	html= fStr(Request.Form("html"))
	ord= fNum(Request.Form("ord"))
	If title<>"" And html<>"" And ord<>"" Then
	  If id = 0 Then
	    Conn.Execute("insert into [Wid] (title,html,ord)values('"&title&"','"&html&"',"&ord&")")
	    Call Message("200","添加侧栏成功！","")
	  Else 
		Conn.Execute("update [Wid] set title='"&title&"',html='"&html&"',ord="&ord&" where id="&id)
		Application.Contents.Remove("Web_Set"&Flag) '删除缓存
		Call Message("200","更新侧栏成功！","")
		End If 
	Else
		Call Message("500","表单数据不完整！","")
	End If 
End Function
'删除边栏
Function delWid() 
     Call isLogin()
	 'Call Message("200","设置更新成功！演示系统不更新","")
     'Conn.Execute("delete * from  [Wid] where id="&id)
	 Call Message("200","删除侧栏成功！","")
End Function
'检查密码
Function ckPass()
	Dim ps,txt
	ps = fStr(Request.Form("ps"))
	Set Rs=Conn.Execute("select pass,content from [log] where id="&id)
	If ps = Rs(0) Then
	    txt = Replace(Rs(1),"""","\""")
		txt = Replace(txt,vbLf," ")	
		Call Message("200",txt,"")
	Else
		Call Message("500","密码不正确！","")
	End If
End Function
'更新缓存
Function  upCache()
	Call isLogin()
	Dim i,CacheList
	CacheList= Split(GetallCache(),",")
	If UBound(CacheList)>0 Then
		For i=0 to UBound(CacheList)-1
			Application.Lock
			Application.Contents.Remove(CacheList(i))
			Application.unLock
		Next
	End If
	Call Message("200","缓存更新完成！","")
End Function
'获取缓存
Function GetallCache()
	Dim Cacheobj
	For Each Cacheobj in Application.Contents
	If InStr(Cacheobj,Flag)>0 Then
		GetallCache = GetallCache & Cacheobj & ","
	End If
	Next
End Function
%>