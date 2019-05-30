<%
'/*------------------------------------------------ Jorkin 自定义类 翻页优化代码
' *********************************************************************
' * 来源: KinJAVA日志 (http://jorkin.reallydo.com/article.asp?id=534)
' * 最后更新: 2009-03-22
' * 当前版本: Ver: 1.09
' *********************************************************************
Class Kin_Db_Pager

    '//-------------------------------------------------------------------------
    '// 定义变量 开始

    Private oConn '//连接对象
    Private sDbType '//数据库类型
    Private sTableName '//表名
    Private sPKey '//主键
    Private sFields '//输出的字段名
    Private sOrderBy '//排序字符串
    Private sSql '//当前的查询语句
    Private sSqlString '//自定义Sql语句
    Private aCondition() '//查询条件(数组)
    Private sCondition '//查询条件(字符串)
    Private iPage '//当前页码
    Private iPageSize '//每页记录数
    Private iPageCount '//总页数
    Private iRecordCount '//当前查询条件下的记录数
    Private sPage '//当前页 替换字符串
    Private sPageCount '//总页数 替换字符串
    Private sRecordCount '//当前查询条件下的记录数 替换字符串
    Private sProjectName '//项目名
    Private sVersion '//版本号
    Private bShowError '//是否显示错误信息
    Private bDistinct '//是否显示唯一记录
    Private sPageInfo '//记录数、页码等信息
    Private sPageParam '//page参数名称
    Private iStyle '//翻页的样式
    Private iPagerSize '//翻页按钮的数值
    Private iCurrentPageSize '//当前页面记录数量
    Private sReWrite '//用ISAP REWRITE做的路径,可用Javascript函数实现AJAX翻页
    Private iTableKind '//表的类型, 是否需要强制加 [ ]
    Private sFirstPage '//首页链接 样式
    Private sPreviewPage '//上一页链接 样式
    Private sCurrentPage '//当前页链接 样式
    Private sListPage '//分页列表链接 样式
    Private sNextPage '//下一页链接 样式
    Private sLastPage '//末页链接 样式
    Private iPagerTop '//分页列表头尾数量
    Private iPagerGroup '//多少页做为一组
    Private sJumpPage '//分页跳转功能
    Private sJumpPageType '//分页跳转类型(可选SELECT或INPUT)
    Private sJumpPageAttr '//分页跳转其他HTML属性
    Private sUrl, sQueryString, x, y
    Private sSpaceMark '//链接之前间隔符

    '//定义变量 结束
    '//-------------------------------------------------------------------------

    '//-------------------------------------------------------------------------
    '//事件、方法: 类初始化事件 开始

    Private Sub Class_Initialize()
        ReDim aCondition( -1)
        sProjectName = "Jorkin &#25968;&#25454;&#24211;&#20998;&#39029;&#31867;  Kin_Db_Pager"
        sDbType = "MSSQL"
        sVersion = "Ver: 1.09 Build: 090322"
        sPKey = "ID"
        sFields = "*"
        sCondition = ""
        sOrderBy = ""
        sSqlString = ""
        iPageSize = 20
        iPage = 1
        iRecordCount = Null
        iPageCount = Null
        bShowError = True
        bDistinct = False
        iPagerTop = 0
        sPage = "{$Kin_Page}"
        sPageCount = "{$Kin_PageCount}"
        sRecordCount = "{$Kin_RecordCount}"
        sPageInfo = "&#20849;&#26377;  {$Kin_RecordCount} &#26465;&#35760;&#24405;  &#39029;&#27425; : {$Kin_Page}/{$Kin_PageCount}"
        sPageParam = "page"
        setPageParam(sPageParam)
        iStyle = 0
        iTableKind = 0
        iPagerSize = 6
        sFirstPage = "&lt;&lt;"
        sPreviewPage = "&lt;"
        sCurrentPage = "{$CurrentPage}"
        sListPage = "{$ListPage}"
        sNextPage = "&gt;"
        sLastPage = "&gt;&gt;"
        sJumpPage = ""
        sJumpPageType = "SELECT"
        sSpaceMark = " "
    End Sub

    '//类结束事件

    Private Sub Class_Terminate()
        Set oConn = Nothing
    End Sub

    '//事件、方法: 类初始化事件 结束
    '//-------------------------------------------------------------------------

    '//-------------------------------------------------------------------------
    '//函数、方法 开始

    '功能:ASP里的IIF
    '来源:http://jorkin.reallydo.com/article.asp?id=26

    Private Function IIf(bExp1, sVal1, sVal2)
        If (bExp1) Then
            IIf = sVal1
        Else
            IIf = sVal2
        End If
    End Function

    '功能:只取数字
    '来源:http://jorkin.reallydo.com/article.asp?id=395

    Private Function Bint(sValue)
        On Error Resume Next
        Bint = 0
        Bint = Fix(CDbl(sValue))
    End Function

    '功能:判断是否是空值
    '来源:http://jorkin.reallydo.com/article.asp?id=386

    Private Function IsBlank(byref TempVar)
        IsBlank = False
        Select Case VarType(TempVar)
            Case 0, 1
                IsBlank = True
            Case 8
                If Len(TempVar) = 0 Then
                    IsBlank = True
                End If
            Case 9
                tmpType = TypeName(TempVar)
                If (tmpType = "Nothing") Or (tmpType = "Empty") Then
                    IsBlank = True
                End If
            Case 8192, 8204, 8209
                If UBound(TempVar) = -1 Then
                    IsBlank = True
                End If
        End Select
    End Function

    '//检查数据库连接是否可用

    Public Function Connect(o)
        If TypeName(o) <> "Connection" Then
            doError "无效的数据库连接。"
        Else
            If o.State = 1 Then
                Set oConn = o
                sDbType = GetDbType(oConn)
            Else
                doError "数据库连接已关闭。"
            End If
        End If
    End Function

    '//处理错误信息

    Public Sub doError(s)
        On Error Resume Next
		If Not bShowError Then Exit Sub
        Dim nRnd
        Randomize()
        nRnd = CLng(Rnd() * 29252888)
        With Response
            .Clear
            .Expires = 0
            .Write "<br />"
            .Write "<div style=""width:100%; font-size:12px; cursor:pointer;line-height:150%"">"
            .Write "<label onClick=""ERRORDIV" & nRnd & ".style.display=(ERRORDIV" & nRnd & ".style.display=='none'?'':'none')"">"
            .Write "<span style=""background-color:820222;color:#FFFFFF;height:23px;font-size:14px;"">〖 Kin_Db_Pager &#25552;&#31034;&#20449;&#24687;  ERROR 〗</span><br />"
            .Write "</label>"
            .Write "<div id=""ERRORDIV" & nRnd & """ style=""width:100%;border:1px solid #820222;padding:5px;overflow:hidden;"">"
            .Write "<span style=""color:#FF0000;"">Description</span> " & Server.HTMLEncode(s) & "<br />"
            .Write "<span style=""color:#FF0000;"">Provider</span> " & sProjectName & "<br />"
            .Write "<span style=""color:#FF0000;"">Version</span> " & sVersion & "<br />"
            .Write "<span style=""color:#FF0000;"">Information</span> Coding By <a href=""http://jorkin.reallydo.com"">Jorkin</a>.<br />"
            .Write "<img width=""0"" height=""0"" src=""http://img.users.51.la/2782986.asp"" style=""display:none"" /></div>"
            .Write "</div>"
            .Write "<br />"
            .End()
        End With
    End Sub

    '//产生分页的SQL语句

    Public Function getSql()
        If Not IsBlank(sSqlString) Then
            getSql = sSqlString
            Exit Function
        End If
        Dim iStart, iEnd
        Call makeCondition()
        iStart = ( iPage - 1 ) * iPageSize
        iEnd = iStart + iPageSize
        Select Case sDbType
            Case "MSSQL"
                getSql = " SELECT " & IIf(bDistinct, "DISTINCT", "") & " " & sFields & " FROM " & TableFormat(sTableName) & " " _
                         & " WHERE [" & sPKey & "] IN ( " _
                         & "   SELECT TOP " & iEnd & " [" & sPKey & "] FROM " & TableFormat(sTableName) & " " & sCondition & " " & sOrderBy & " " _
                         & " )"
                If iPage>1 Then
                    getSql = getSql & " AND [" & sPKey & "] NOT IN ( " _
                             & "   SELECT TOP " & iStart & " [" & sPKey & "] FROM " & TableFormat(sTableName) & " " & sCondition & " " & sOrderBy & " " _
                             & " )"
                End If
                getSql = getSql & " " & sOrderBy
            Case "MYSQL"
                getSql = "SELECT " & sFields & " FROM " & TableFormat(sTableName)& " " & sCondition & " " & sOrderBy & " LIMIT "&(iPage -1) * iPageSize&"," & iPageSize
            Case "MSSQLPRODUCE"
            Case "ACCESS"
                getSql = "SELECT " & IIf(bDistinct, "DISTINCT ", " ") & " Top " & iPage * iPageSize & " " & sFields & " FROM " & TableFormat(sTableName) & " " & sCondition & " " & sOrderBy
            Case Else
                getSql = "SELECT " & sFields & " FROM " & TableFormat(sTableName) & " " & sCondition & " " & sOrderBy
        End Select
    End Function

    '//产生条件字符串

    Private Sub makeCondition()
        If Not IsBlank(sCondition) Then Exit Sub
        If UBound(aCondition)>= 0 Then
            sCondition = " WHERE " & Join(aCondition, " AND ")
        End If
    End Sub

    '//计算记录数

    Private Sub CaculateRecordCount()
        On Error Resume Next
        Dim oRs
        If Not IsBlank(sSqlString) Then
            sSql = "SELECT COUNT(0) FROM (" & sSqlString & ")"
        Else
            Call makeCondition()
            sSql = "SELECT COUNT(0) FROM " & TableFormat(sTableName) & " " & IIf(IsBlank(sCondition), "", sCondition)
        End If
        Set oRs = oConn.Execute( sSql )
        If Err Then
            doError Err.Description
        End If
        iRecordCount = oRs.Fields.Item(0).Value
        Set oRs = Nothing
    End Sub

    '//计算页数

    Private Sub CaculatePageCount()
        If IsNull(iRecordCount) Then CaculateRecordCount()
        If iRecordCount = 0 Then
            iPageCount = 0
            Exit Sub
        End If
        iPageCount = Abs( Int( 0 - (iRecordCount / iPageSize) ) )
    End Sub

    '//设置页码

    Private Function setPage(n)
        iPage = Bint(n)
        If iPage < 1 Then iPage = 1
    End Function

    '//增加条件

    Public Sub AddCondition(s)
        If IsBlank(s) Then Exit Sub
        ReDim Preserve aCondition(UBound(aCondition) + 1)
        aCondition(UBound(aCondition)) = s
    End Sub

    '//判断页面连接

    Private Function ReWrite(n)
        n = Bint(n)
        If Not IsBlank(sRewrite) Then
            ReWrite = Replace(sReWrite, "*", n)
        Else
            ReWrite = sUrl & IIf(n>0, n, "")
        End If
    End Function

    '//数据库表加 []

    Private Function TableFormat(s)
        Select Case iTableKind
            Case 0
                TableFormat = "[" & s & "]"
            Case 1
                TableFormat = " " & s & " "
        End Select
    End Function

    '//按Where In顺序进行排序

    Public Function OrderIn(s, sOrderIn)
        OrderIn = " "
        If Not IsBlank(s) And Not IsBlank(sOrderIn) Then
            sOrderIn = Replace(sOrderIn, " ", "")
            sOrderIn = Replace(sOrderIn, "'", "")
            sOrderIn = "'" & sOrderIn & "'"
            Select Case sDbType
                Case "MYSQL"
                    OrderIn = "FIND_IN_SET(" & s & ", " & sOrderIn & ")"
                Case "ACCESS"
                    OrderIn = "INSTR(','+CStr(" & sOrderIn & ")+',',','+CStr(" & s & ")+',')"
                Case Else
                    OrderIn = "PATINDEX('% ' + CONVERT(nvarchar(820222), " & s & ") + ' %',' ' + CONVERT(nvarchar(820222), Replace(" & sOrderIn & ", ',', ' , ')) + ' ')"
            End Select
        End If
        OrderIn = OrderIn & " "
    End Function

    '//根据数据库连接判断数据库类型

    Private Function GetDbType(o)
        Select Case (o.Provider)
            Case "MSDASQL.1", "SQLOLEDB.1", "SQLOLEDB"
                GetDbType = "MSSQL"
            Case "MSDAORA.1", "OraOLEDB.Oracle"
                GetDbType = "ORACLE"
            Case "Microsoft.Jet.OLEDB.4.0"
                GetDbType = "ACCESS"
        End Select
    End Function

    '//设定分页变量的名称

    Private Function setPageParam(s)
        sQueryString = ""
        For Each x In Request.QueryString
            If x <> sPageParam Then
                For Each y In Request.QueryString(x)
                    sQueryString = "&" & x & "=" & Server.URLEncode(y) & sQueryString
                Next
            End If
        Next
        sUrl = Request.ServerVariables("URL") & "?" & IIf(IsBlank(sQueryString), "", Mid(sQueryString, 2) & "&") & sPageParam & "="
    End Function

    '//函数、方法 结束
    '//-------------------------------------------------------------------------

    '//-------------------------------------------------------------------------
    '//输入属性 开始

    '//定义连接对象

    Public Property Set ActiveConnection(o)
        Set oConn = o
        sDbType = GetDbType(oConn)
    End Property

    '//连接字符串

    Public Property Let ConnectionString(s)
        Set oConn = Server.CreateObject("ADODB.Connection")
        oConn.ConnectionString = s
        oConn.Open()
        sDbType = GetDbType(oConn)
    End Property

    '//定义数据库类型

    Public Property Let DBType(s)
        sDBType = UCase(s)
        Select Case sDBType
            Case "ACCESS", "ACC", "AC"
                sDBType = "ACCESS"
            Case "MSSQL", "SQL"
                sDBType = "MSSQL"
            Case "MYSQL"
                sDBType = "MYSQL"
            Case "ORACLE"
                sDBType = "ORACLE"
            Case "PGSQL"
                sDBType = "PGSQL"
            Case "MSSQLPRODUCE", "MSSQLPR", "MSSQL_PR", "PR"
                sDBType = "MSSQLPRODUCE"
            Case Else
                If TypeName(oConn) = "Connection" Then
                    sDBType = GetDbType(oConn)
                End If
        End Select
    End Property

    '//定义 首页 样式

    Public Property Let FirstPage(s)
        sFirstPage = s
    End Property

    '//定义 上一页 样式

    Public Property Let PreviewPage(s)
        sPreviewPage = s
    End Property

    '//定义 当前页 样式

    Public Property Let CurrentPage(s)
        sCurrentPage = s
    End Property

    '//定义 分页列表页 样式

    Public Property Let ListPage(s)
        sListPage = s
    End Property

    '//定义 下一页 样式

    Public Property Let NextPage(s)
        sNextPage = s
    End Property

    '//定义 末页 样式

    Public Property Let LastPage(s)
        sLastPage = s
    End Property

    '//定义间隔符，默认半角空格

    Public Property Let SpaceMark(s)
        sSpaceMark = s
    End Property

    '//定义 列表前后多加几页

    Public Property Let PagerTop(n)
        iPagerTop = Bint(n)
    End Property

    '//定义查询表名

    Public Property Let TableName(s)
        sTableName = s
        '//如果发现表名包含 ([. ，那么就不要用 []
        If InStr(s, "(")>0 Then iTableKind = 1
        If InStr(s, "[")>0 Then iTableKind = 1
        If InStr(s, ".")>0 Then iTableKind = 1
    End Property

    '//定义需要输出的字段名

    Public Property Let Fields(s)
        sFields = s
    End Property

    '//定义主键

    Public Property Let PKey(s)
        If Not IsBlank(s) Then sPKey = s
    End Property

    '//定义排序规则

    Public Property Let OrderBy(s)
        If Not IsBlank(s) Then sOrderBy = " ORDER BY " & s & " "
    End Property

    '//定义每页的记录条数

    Public Property Let PageSize(s)
        iPageSize = Bint(s)
        iPageSize = IIf(iPageSize<1, 1, iPageSize)
    End Property

    '//定义当前页码

    Public Property Let Page(n)
        setPage Bint(n)
    End Property

    '//定义当前页码(同Property Page)

    Public Property Let AbsolutePage(n)
        setPage Bint(n)
    End Property

    '//自定义查询语句

    Public Property Let Sql(s)
        sSqlString = s
    End Property

    '//是否DISTINCT

    Public Property Let Distinct(b)
        bDistinct = b
    End Property

    '//设定分页变量的名称

    Public Property Let PageParam(s)
        sPageParam = LCase(s)
        If IsBlank(sPageParam) Then sPageParam = "page"
        setPageParam(sPageParam)
    End Property

    '//选择分页的样式,可以后面自己添加新的

    Public Property Let Style(s)
        iStyle = Bint(s)
    End Property

    '//分页列表显示数量

    Public Property Let PagerSize(n)
        iPagerSize = Bint(n)
    End Property

    '//自定义ISAPI_REWRITE路径 * 将被替换为当前页数
    '//使用Javascript时请注意本分页类用双引号引用字符串,请先处理.

    Public Property Let ReWritePath(s)
        sReWrite = s
    End Property

    '//强制TABLE类型

    Public Property Let TableKind(n)
        iTableKind = n
    End Property

    '//自定义分页信息

    Public Property Let PageInfo(s)
        sPageInfo = s
    End Property

    '//定义页面跳转类型

    Public Property Let JumpPageType(s)
        sJumpPageType = UCase(s)
        Select Case sJumpPageType
            Case "INPUT", "SELECT"
            Case Else
                sJumpPageType = "SELECT"
        End Select
    End Property

    '//定义页面跳转链接其他HTML属性

    Public Property Let JumpPageAttr(s)
        sJumpPageAttr = s
    End Property

    '//输入属性 结束
    '//-------------------------------------------------------------------------

    '//-------------------------------------------------------------------------
    '//输出属性 开始

    '//输出连接语句

    Public Property Get ConnectionString()
        ConnectionString = oConn.ConnectionString
    End Property

    '//输出连接对象

    Public Property Get Conn()
        Set Conn = oConn
    End Property

    '//输出数据库类型

    Public Property Get DBType()
        DBType = sDBType
    End Property

    '//输出查询表名

    Public Property Get TableName()
        TableName = sTableName
    End Property

    '//输出需要输出的字段名

    Public Property Get Fields()
        Fields = sFields
    End Property

    '//输出主键

    Public Property Get PKey()
        PKey = sPKey
    End Property

    '//输出排序规则

    Public Property Get OrderBy()
        OrderBy = sOrderBy
    End Property

    '//取得当前条件下的记录数

    Public Property Get RecordCount()
        If IsNull(iRecordCount) Then CaculateRecordCount()
        RecordCount = iRecordCount
    End Property

    '//取得每页记录数

    Public Property Get PageSize()
        PageSize = iPageSize
    End Property

    '//取得当前查询的条件

    Public Property Get Condition()
        If IsBlank(sCondition) Then makeCondition()
        Condition = sCondition
    End Property

    '//取得当前页码

    Public Property Get Page()
        Page = iPage
    End Property

    '//取得当前页码

    Public Property Get AbsolutePage()
        AbsolutePage = iPage
    End Property

    '//取得总页数

    Public Property Get PageCount()
        If IsNull(iPageCount) Then CaculatePageCount()
        PageCount = iPageCount
    End Property

    '//取得当前页记录数

    Public Property Get CurrentPageSize()
        If IsNull(iRecordCount) Then CaculateRecordCount()
        If IsNull(iPageCount) Then CaculatePageCount()
        CurrentPageSize = IIf(iRecordCount>0, IIf(iPage = iPageCount, iRecordCount - (iPage -1) * iPageSize, iPageSize), 0)
    End Property

    '//得到分页后的记录集

    Public Property Get RecordSet()
        On Error Resume Next
        Select Case sDbType
            Case "MSSQL" '// MSSQL2000
                sSql = getSql()
                Set RecordSet = oConn.Execute( sSql )
            Case "MSSQLPRODUCE" '// SqlServer2000数据库存储过程版, 可使用叶子的SQL。
                Set oRs = Server.CreateObject("ADODB.RecordSet")
                Set oCommand = Server.CreateObject("ADODB.Command")
                oCommand.CommandType = 4
                oCommand.ActiveConnection = oConn
                oCommand.CommandText = "sp_Util_Page"
                oCommand.Parameters(1) = 0
                oCommand.Parameters(2) = iPage
                oCommand.Parameters(3) = iPageSize
                oCommand.Parameters(4) = sPkey
                oCommand.Parameters(5) = sFields
                oCommand.Parameters(6) = sTableName
                oCommand.Parameters(7) = Join(aCondition, " AND ")
                oCommand.Parameters(8) = Mid(sOrderBy, 11)
                oRs.CursorLocation = 3
                oRs.LockType = 1
                oRs.Open oCommand
            Case "MYSQL" 'MYSQL数据库，不会，暂时空着。
                sSql = getSql()
                Set oRs = oConn.Execute(sSql)
            Case Else '其他情况按最原始的ADO方法处理，包括ACCESS。
                sSql = getSql()
                Set RecordSet = Server.CreateObject ("ADODB.RecordSet")
                RecordSet.Open sSql, oConn, 1, 1, &H0001
                RecordSet.PageSize = iPageSize
                If RecordSet.AbsolutePage <> -1 Then
                    iPage = IIf(iPage > RecordSet.PageCount, RecordSet.PageCount, iPage)
                    RecordSet.AbsolutePage = iPage
                End If
        End Select
        If Err Then
            doError Err.Description
            If Not IsBlank(sSql) Then
                Set RecordSet = oConn.Execute( sSql )
                If Err Then doError Err.Description
            Else
                doError Err.Description
            End If
        End If
        Err.Clear()
    End Property

    '//版本信息

    Public Property Get Version()
        Version = sVersion
    End Property

    '//输出页码及记录数等信息

    Public Property Get PageInfo()
        CaculatePageCount()
        PageInfo = Replace(sPageInfo, sRecordCount, iRecordCount)
        PageInfo = Replace(PageInfo, sPageCount, iPageCount)
        PageInfo = Replace(PageInfo, sPage, iPage)
    End Property

    '//输出分页样式

    Public Property Get Style()
        Style = iStyle
    End Property

    '//输出分页变量

    Public Property Get PageParam()
        PageParam = sPageParam
    End Property

    '//输出翻页按钮

    Public Property Get Pager()
        Dim ii, iStart, iEnd
        Pager = ""
        ii = (iPagerSize \ 2)
        iEnd = iPage + ii
        iStart = iPage - (ii + (iPagerSize Mod 2)) + 1
        If iEnd > iPageCount Then
            iEnd = iPageCount
            iStart = iPageCount - iPagerSize + 1
        End If
        If iStart < 1 Then
            iStart = 1
            iEnd = iStart + iPagerSize -1
        End If
        If iEnd > iPageCount Then
            iEnd = iPageCount
        End If 
                If iPageCount>0 Then
                    If iPage>1 Then
                        Pager = Pager & IIf(IsBlank(sFirstPage), "", "<a href=""" & Rewrite(1) & """>" & sFirstPage & "</a>" & sSpaceMark)
                        Pager = Pager & IIf(IsBlank(sPreviewPage), "", "<a href=""" & Rewrite((iPage -1)) & """>" & sPreviewPage & "</a>" & sSpaceMark)
                    Else
                        Pager = Pager & IIf(IsBlank(sFirstPage), "", "<span class=""disabled"">" & sFirstPage & "</span>" & sSpaceMark)
                        Pager = Pager & IIf(IsBlank(sPreviewPage), "", "<span class=""disabled"">" & sPreviewPage & "</span>" & sSpaceMark)
                    End If
                    If iPagerTop > 0 Then
                        If iPagerTop < iStart Then
                            ii = iPagerTop
                        Else
                            ii = iStart - 1
                        End If
                        For i = 1 To ii
                            Pager = Pager & "<a href=""" & ReWrite(i) & """>" & Replace(sListPage, "{$Listpage}", i, 1, -1, 1) & "</a>" & sSpaceMark
                        Next
                        If iPagerTop < iStart -1 Then Pager = Pager & "..." & sSpaceMark
                    End If
                    If iPagerSize >0 Then
                        For i = iStart To iEnd
                            If i = iPage Then
                                Pager = Pager & "<span class=""current"">" & Replace(sCurrentPage, "{$Currentpage}", i, 1, -1, 1) & "</span>" & sSpaceMark
                            Else
                                Pager = Pager & "<a href=""" & ReWrite(i) & """>" & Replace(sListPage, "{$Listpage}", i, 1, -1, 1) & "</a>" & sSpaceMark
                            End If
                        Next
                    End If
                    If iPagerTop > 0 Then
                        If iPageCount - iPagerTop > iEnd Then Pager = Pager & "..." & sSpaceMark
                        If iPageCount - iPagerTop > iEnd Then
                            ii = iPageCount - iPagerTop + 1
                        Else
                            ii = iEnd + 1
                        End If
                        For i = ii To iPageCount
                            Pager = Pager & "<a href=""" & ReWrite(i) & """>" & Replace(sListPage, "{$Listpage}", i, 1, -1, 1) & "</a>" & sSpaceMark
                        Next
                    End If
                    If iPageCount>iPage Then
                        Pager = Pager & IIf(IsBlank(sNextPage), "", "<a href=""" & Rewrite(iPage + 1) & """>" & sNextPage & "</a>" & sSpaceMark)
                        Pager = Pager & IIf(IsBlank(sLastPage), "", "<a href=""" & Rewrite(iPageCount) & """>" & sLastPage & "</a>" & sSpaceMark)
                    Else
                        Pager = Pager & IIf(IsBlank(sNextPage), "", "<span class=""disabled"">" & sNextPage & "</span>" & sSpaceMark)
                        Pager = Pager & IIf(IsBlank(sLastPage), "", "<span class=""disabled"">" & sLastPage & "</span>")
                    End If
                End If           
    End Property
    '//输出属性 结束
    '//-------------------------------------------------------------------------
End Class
%>