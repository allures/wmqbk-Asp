<!--#include file="config.asp" -->
<!--#include file="function.asp" -->
<!--#include file="tpl.class.asp"-->
<!--#include file="page.class.asp" -->
<%
Class wmBlog

Private  id,s,tpl,oRs,sRs,cRs,templete,action,oDbPager,nick

Private Sub Class_Initialize
	Set tpl  = New sTemplate  '使用tpl模板引擎
	s        = Request.Querystring("s")
	action   = Request.Querystring("act")
	templete = ""
	Set sRs  = Server.CreateObject("ADODB.RecordSet")
	sRs.Open "select * from [Wid] order by ord",Conn,1,1
	tpl.assign "ifile",iFile '当前文件名
	tpl.assign "webtitle",WebTitle'标题
	tpl.assign "webdesc",WebDesc'描述
	tpl.assign "rewrite",Rewrite'伪静态
	tpl.assign "admin",admin
End Sub

Private Sub Class_Terminate()
	sRs.Close()
	cRs.Close()
	Set Tpl      = Nothing
	Set oDbPager = Nothing
	CloseConn()
End Sub

Public Sub run()

	Select Case action
		Case "dologin"
			Call doLogin()
		Case "set"
			Call isLogin()
			Set oRs  = Conn.Execute("select * from [Set] where id=1")
			templete = "setting.htm"
			tpl.assign "v",oRs
		Case "wid"
			Call isLogin()
			'Set oRs = Conn.Execute("select * from [Wid]")
			templete = "widget.htm"
			tpl.assign "widset", sRs
		Case "add"
			Call isLogin()
			templete = "post.htm"
			tpl.assign "id",0
			tpl.assign "c",action
			tpl.assign "btn","发 布"
		Case "edit"
			Call isLogin()
			id       = CLng(Request.QueryString("id"))
			Set oRs  = Conn.Execute("select * from [Log] where id=" & id)
			templete = "post.htm"
			tpl.assign "v",oRs
			tpl.assign "c",action
			tpl.assign "btn","编 辑"
		Case "login"
			templete = "login.htm" '模板文件
		Case "logout"
			Call logOut()
		Case "pl"
			id      = CLng(Request.QueryString("id"))
			Set oRs = Conn.Execute("select * from [Log] where id=" & id)
			If(oRs.Eof) Then Response.Status = "404 Not Found":Response.End
			if (admin=0 and oRs("hide")=1) then Response.Status = "404 Not Found":Response.End
			'Set cRs = Conn.Execute("select * from [Pl] where cid="&id)
		
			If admin = 1 Then
				nick = WebUser
			Else
				nick = Request.Cookies("4jax-nick")
			End If

			if oRs("lock")=0 Then
			Set cRs  = Server.CreateObject("ADODB.RecordSet")
			cRs.Open "select * from [Pl] where cid=" & id,Conn,1,1
			tpl.assign "list",cRs
			tpl.assign "nick",nick
			tpl.assign "safecode",SafeCode '验证码
			end if



	

			templete = "view.htm" '模板文件			
			tpl.assign "v",oRs			
			'tpl.assign "pform",pFrom(id)
			'Call pfrom()
		Case "plist"
			'tpl.setCache        = "cache,4,600" '缓存名称,缓存方式,缓存时间(默认是秒)        
			Set oDbPager       = New Kin_db_Pager
			'页面参数
			templete           = "plist.htm" '模板文件
			oDbPager.Connect(Conn)
			oDbPager.PageParam = "page"
			'//指定数据库类型.默认值:"MSSQL"
			oDbPager.DbType    = "ACCESS"
			oDbPager.TableName = "[Pl]"
			'//选择列 用逗号分隔 默认值:"*"
			oDbPager.Fields    = "*"
			'//指定该表的主键
			oDbPager.PKey      = "pid"
			'//指定排序条件
			oDbPager.OrderBy   = "pid DESC"
			'读取模板标签 
			If Rewrite = 1 Then oDbPager.RewritePath = "comment-*.html"
			iPageSize         = 20
			oDbPager.PageSize = iPageSize
			'//指定当前页数
			oDbPager.Page     = Request.QueryString(oDbPager.PageParam)
			Set oRs           = oDbPager.Recordset
			'//获取分页信息
			sPageInfo         = oDbPager.PageInfo
			sPager            = oDbPager.Pager
			'sJumpPage = oDbPager.JumpPage
			tpl.assign "list",oRs'分页后的记录集
			tpl.assign "pagelist",sPager'分页列表
			tpl.assign "pagenum", oDbPager.RecordCount'记录总数
			tpl.assign "pagesize",iPageSize'记录总数  
		Case Else
			Set oDbPager       = New Kin_db_Pager
			'页面参数
			templete           = "index.htm" '模板文件
			oDbPager.Connect(Conn)
			oDbPager.PageParam = "page"
			'//指定数据库类型.默认值:"MSSQL"
			oDbPager.DbType    = "ACCESS"
			oDbPager.TableName = "[Log]"
			'//选择列 用逗号分隔 默认值:"*"
			oDbPager.Fields    = "*"
			'//指定该表的主键
			oDbPager.PKey      = "id"
			'//指定排序条件
			oDbPager.OrderBy   = "ist DESC,atime DESC"
			'读取模板标签
			If admin = 0 Then  oDbPager.AddCondition "hide=0"

			If Len(s) > 1 And  Len(s) < 11 Then
				oDbPager.AddCondition "title like '%" & s & "%'"
			Else
				If Rewrite = 1 Then oDbPager.RewritePath = "index-*.html"
			End If

			iPageSize = 20
			oDbPager.PageSize = iPageSize
			'//指定当前页数
			oDbPager.Page = Request.QueryString(oDbPager.PageParam)
			'//也可以直接使用自定义的SQL语句选取记录集
			Set oRs = oDbPager.Recordset
			'//获取分页信息
			sPageInfo = oDbPager.PageInfo
			sPager = oDbPager.Pager
			'sJumpPage = oDbPager.JumpPage
			tpl.assign "list",oRs'分页后的记录集
			tpl.assign "pagelist",sPager'分页列表
			tpl.assign "pagenum",oDbPager.RecordCount'记录总数
			tpl.assign "pagesize",iPageSize'记录总数  
			tpl.assign "isjpeg", isjpeg
	End Select

	tpl.assign "widget",sRs
	tpl.assign "comment", NewCom(10)
	tpl.assign "topic", TopIc(10)
	tpl.display templete
End Sub

End Class %>