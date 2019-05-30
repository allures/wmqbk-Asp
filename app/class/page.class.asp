<%
'/*------------------------------------------------ Jorkin �Զ����� ��ҳ�Ż�����
' *********************************************************************
' * ��Դ: KinJAVA��־ (http://jorkin.reallydo.com/article.asp?id=534)
' * ������: 2009-03-22
' * ��ǰ�汾: Ver: 1.09
' *********************************************************************
Class Kin_Db_Pager

    '//-------------------------------------------------------------------------
    '// ������� ��ʼ

    Private oConn '//���Ӷ���
    Private sDbType '//���ݿ�����
    Private sTableName '//����
    Private sPKey '//����
    Private sFields '//������ֶ���
    Private sOrderBy '//�����ַ���
    Private sSql '//��ǰ�Ĳ�ѯ���
    Private sSqlString '//�Զ���Sql���
    Private aCondition() '//��ѯ����(����)
    Private sCondition '//��ѯ����(�ַ���)
    Private iPage '//��ǰҳ��
    Private iPageSize '//ÿҳ��¼��
    Private iPageCount '//��ҳ��
    Private iRecordCount '//��ǰ��ѯ�����µļ�¼��
    Private sPage '//��ǰҳ �滻�ַ���
    Private sPageCount '//��ҳ�� �滻�ַ���
    Private sRecordCount '//��ǰ��ѯ�����µļ�¼�� �滻�ַ���
    Private sProjectName '//��Ŀ��
    Private sVersion '//�汾��
    Private bShowError '//�Ƿ���ʾ������Ϣ
    Private bDistinct '//�Ƿ���ʾΨһ��¼
    Private sPageInfo '//��¼����ҳ�����Ϣ
    Private sPageParam '//page��������
    Private iStyle '//��ҳ����ʽ
    Private iPagerSize '//��ҳ��ť����ֵ
    Private iCurrentPageSize '//��ǰҳ���¼����
    Private sReWrite '//��ISAP REWRITE����·��,����Javascript����ʵ��AJAX��ҳ
    Private iTableKind '//�������, �Ƿ���Ҫǿ�Ƽ� [ ]
    Private sFirstPage '//��ҳ���� ��ʽ
    Private sPreviewPage '//��һҳ���� ��ʽ
    Private sCurrentPage '//��ǰҳ���� ��ʽ
    Private sListPage '//��ҳ�б����� ��ʽ
    Private sNextPage '//��һҳ���� ��ʽ
    Private sLastPage '//ĩҳ���� ��ʽ
    Private iPagerTop '//��ҳ�б�ͷβ����
    Private iPagerGroup '//����ҳ��Ϊһ��
    Private sJumpPage '//��ҳ��ת����
    Private sJumpPageType '//��ҳ��ת����(��ѡSELECT��INPUT)
    Private sJumpPageAttr '//��ҳ��ת����HTML����
    Private sUrl, sQueryString, x, y
    Private sSpaceMark '//����֮ǰ�����

    '//������� ����
    '//-------------------------------------------------------------------------

    '//-------------------------------------------------------------------------
    '//�¼�������: ���ʼ���¼� ��ʼ

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

    '//������¼�

    Private Sub Class_Terminate()
        Set oConn = Nothing
    End Sub

    '//�¼�������: ���ʼ���¼� ����
    '//-------------------------------------------------------------------------

    '//-------------------------------------------------------------------------
    '//���������� ��ʼ

    '����:ASP���IIF
    '��Դ:http://jorkin.reallydo.com/article.asp?id=26

    Private Function IIf(bExp1, sVal1, sVal2)
        If (bExp1) Then
            IIf = sVal1
        Else
            IIf = sVal2
        End If
    End Function

    '����:ֻȡ����
    '��Դ:http://jorkin.reallydo.com/article.asp?id=395

    Private Function Bint(sValue)
        On Error Resume Next
        Bint = 0
        Bint = Fix(CDbl(sValue))
    End Function

    '����:�ж��Ƿ��ǿ�ֵ
    '��Դ:http://jorkin.reallydo.com/article.asp?id=386

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

    '//������ݿ������Ƿ����

    Public Function Connect(o)
        If TypeName(o) <> "Connection" Then
            doError "��Ч�����ݿ����ӡ�"
        Else
            If o.State = 1 Then
                Set oConn = o
                sDbType = GetDbType(oConn)
            Else
                doError "���ݿ������ѹرա�"
            End If
        End If
    End Function

    '//���������Ϣ

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
            .Write "<span style=""background-color:820222;color:#FFFFFF;height:23px;font-size:14px;"">�� Kin_Db_Pager &#25552;&#31034;&#20449;&#24687;  ERROR ��</span><br />"
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

    '//������ҳ��SQL���

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

    '//���������ַ���

    Private Sub makeCondition()
        If Not IsBlank(sCondition) Then Exit Sub
        If UBound(aCondition)>= 0 Then
            sCondition = " WHERE " & Join(aCondition, " AND ")
        End If
    End Sub

    '//�����¼��

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

    '//����ҳ��

    Private Sub CaculatePageCount()
        If IsNull(iRecordCount) Then CaculateRecordCount()
        If iRecordCount = 0 Then
            iPageCount = 0
            Exit Sub
        End If
        iPageCount = Abs( Int( 0 - (iRecordCount / iPageSize) ) )
    End Sub

    '//����ҳ��

    Private Function setPage(n)
        iPage = Bint(n)
        If iPage < 1 Then iPage = 1
    End Function

    '//��������

    Public Sub AddCondition(s)
        If IsBlank(s) Then Exit Sub
        ReDim Preserve aCondition(UBound(aCondition) + 1)
        aCondition(UBound(aCondition)) = s
    End Sub

    '//�ж�ҳ������

    Private Function ReWrite(n)
        n = Bint(n)
        If Not IsBlank(sRewrite) Then
            ReWrite = Replace(sReWrite, "*", n)
        Else
            ReWrite = sUrl & IIf(n>0, n, "")
        End If
    End Function

    '//���ݿ��� []

    Private Function TableFormat(s)
        Select Case iTableKind
            Case 0
                TableFormat = "[" & s & "]"
            Case 1
                TableFormat = " " & s & " "
        End Select
    End Function

    '//��Where In˳���������

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

    '//�������ݿ������ж����ݿ�����

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

    '//�趨��ҳ����������

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

    '//���������� ����
    '//-------------------------------------------------------------------------

    '//-------------------------------------------------------------------------
    '//�������� ��ʼ

    '//�������Ӷ���

    Public Property Set ActiveConnection(o)
        Set oConn = o
        sDbType = GetDbType(oConn)
    End Property

    '//�����ַ���

    Public Property Let ConnectionString(s)
        Set oConn = Server.CreateObject("ADODB.Connection")
        oConn.ConnectionString = s
        oConn.Open()
        sDbType = GetDbType(oConn)
    End Property

    '//�������ݿ�����

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

    '//���� ��ҳ ��ʽ

    Public Property Let FirstPage(s)
        sFirstPage = s
    End Property

    '//���� ��һҳ ��ʽ

    Public Property Let PreviewPage(s)
        sPreviewPage = s
    End Property

    '//���� ��ǰҳ ��ʽ

    Public Property Let CurrentPage(s)
        sCurrentPage = s
    End Property

    '//���� ��ҳ�б�ҳ ��ʽ

    Public Property Let ListPage(s)
        sListPage = s
    End Property

    '//���� ��һҳ ��ʽ

    Public Property Let NextPage(s)
        sNextPage = s
    End Property

    '//���� ĩҳ ��ʽ

    Public Property Let LastPage(s)
        sLastPage = s
    End Property

    '//����������Ĭ�ϰ�ǿո�

    Public Property Let SpaceMark(s)
        sSpaceMark = s
    End Property

    '//���� �б�ǰ���Ӽ�ҳ

    Public Property Let PagerTop(n)
        iPagerTop = Bint(n)
    End Property

    '//�����ѯ����

    Public Property Let TableName(s)
        sTableName = s
        '//������ֱ������� ([. ����ô�Ͳ�Ҫ�� []
        If InStr(s, "(")>0 Then iTableKind = 1
        If InStr(s, "[")>0 Then iTableKind = 1
        If InStr(s, ".")>0 Then iTableKind = 1
    End Property

    '//������Ҫ������ֶ���

    Public Property Let Fields(s)
        sFields = s
    End Property

    '//��������

    Public Property Let PKey(s)
        If Not IsBlank(s) Then sPKey = s
    End Property

    '//�����������

    Public Property Let OrderBy(s)
        If Not IsBlank(s) Then sOrderBy = " ORDER BY " & s & " "
    End Property

    '//����ÿҳ�ļ�¼����

    Public Property Let PageSize(s)
        iPageSize = Bint(s)
        iPageSize = IIf(iPageSize<1, 1, iPageSize)
    End Property

    '//���嵱ǰҳ��

    Public Property Let Page(n)
        setPage Bint(n)
    End Property

    '//���嵱ǰҳ��(ͬProperty Page)

    Public Property Let AbsolutePage(n)
        setPage Bint(n)
    End Property

    '//�Զ����ѯ���

    Public Property Let Sql(s)
        sSqlString = s
    End Property

    '//�Ƿ�DISTINCT

    Public Property Let Distinct(b)
        bDistinct = b
    End Property

    '//�趨��ҳ����������

    Public Property Let PageParam(s)
        sPageParam = LCase(s)
        If IsBlank(sPageParam) Then sPageParam = "page"
        setPageParam(sPageParam)
    End Property

    '//ѡ���ҳ����ʽ,���Ժ����Լ�����µ�

    Public Property Let Style(s)
        iStyle = Bint(s)
    End Property

    '//��ҳ�б���ʾ����

    Public Property Let PagerSize(n)
        iPagerSize = Bint(n)
    End Property

    '//�Զ���ISAPI_REWRITE·�� * �����滻Ϊ��ǰҳ��
    '//ʹ��Javascriptʱ��ע�Ȿ��ҳ����˫���������ַ���,���ȴ���.

    Public Property Let ReWritePath(s)
        sReWrite = s
    End Property

    '//ǿ��TABLE����

    Public Property Let TableKind(n)
        iTableKind = n
    End Property

    '//�Զ����ҳ��Ϣ

    Public Property Let PageInfo(s)
        sPageInfo = s
    End Property

    '//����ҳ����ת����

    Public Property Let JumpPageType(s)
        sJumpPageType = UCase(s)
        Select Case sJumpPageType
            Case "INPUT", "SELECT"
            Case Else
                sJumpPageType = "SELECT"
        End Select
    End Property

    '//����ҳ����ת��������HTML����

    Public Property Let JumpPageAttr(s)
        sJumpPageAttr = s
    End Property

    '//�������� ����
    '//-------------------------------------------------------------------------

    '//-------------------------------------------------------------------------
    '//������� ��ʼ

    '//����������

    Public Property Get ConnectionString()
        ConnectionString = oConn.ConnectionString
    End Property

    '//������Ӷ���

    Public Property Get Conn()
        Set Conn = oConn
    End Property

    '//������ݿ�����

    Public Property Get DBType()
        DBType = sDBType
    End Property

    '//�����ѯ����

    Public Property Get TableName()
        TableName = sTableName
    End Property

    '//�����Ҫ������ֶ���

    Public Property Get Fields()
        Fields = sFields
    End Property

    '//�������

    Public Property Get PKey()
        PKey = sPKey
    End Property

    '//����������

    Public Property Get OrderBy()
        OrderBy = sOrderBy
    End Property

    '//ȡ�õ�ǰ�����µļ�¼��

    Public Property Get RecordCount()
        If IsNull(iRecordCount) Then CaculateRecordCount()
        RecordCount = iRecordCount
    End Property

    '//ȡ��ÿҳ��¼��

    Public Property Get PageSize()
        PageSize = iPageSize
    End Property

    '//ȡ�õ�ǰ��ѯ������

    Public Property Get Condition()
        If IsBlank(sCondition) Then makeCondition()
        Condition = sCondition
    End Property

    '//ȡ�õ�ǰҳ��

    Public Property Get Page()
        Page = iPage
    End Property

    '//ȡ�õ�ǰҳ��

    Public Property Get AbsolutePage()
        AbsolutePage = iPage
    End Property

    '//ȡ����ҳ��

    Public Property Get PageCount()
        If IsNull(iPageCount) Then CaculatePageCount()
        PageCount = iPageCount
    End Property

    '//ȡ�õ�ǰҳ��¼��

    Public Property Get CurrentPageSize()
        If IsNull(iRecordCount) Then CaculateRecordCount()
        If IsNull(iPageCount) Then CaculatePageCount()
        CurrentPageSize = IIf(iRecordCount>0, IIf(iPage = iPageCount, iRecordCount - (iPage -1) * iPageSize, iPageSize), 0)
    End Property

    '//�õ���ҳ��ļ�¼��

    Public Property Get RecordSet()
        On Error Resume Next
        Select Case sDbType
            Case "MSSQL" '// MSSQL2000
                sSql = getSql()
                Set RecordSet = oConn.Execute( sSql )
            Case "MSSQLPRODUCE" '// SqlServer2000���ݿ�洢���̰�, ��ʹ��Ҷ�ӵ�SQL��
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
            Case "MYSQL" 'MYSQL���ݿ⣬���ᣬ��ʱ���š�
                sSql = getSql()
                Set oRs = oConn.Execute(sSql)
            Case Else '�����������ԭʼ��ADO������������ACCESS��
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

    '//�汾��Ϣ

    Public Property Get Version()
        Version = sVersion
    End Property

    '//���ҳ�뼰��¼������Ϣ

    Public Property Get PageInfo()
        CaculatePageCount()
        PageInfo = Replace(sPageInfo, sRecordCount, iRecordCount)
        PageInfo = Replace(PageInfo, sPageCount, iPageCount)
        PageInfo = Replace(PageInfo, sPage, iPage)
    End Property

    '//�����ҳ��ʽ

    Public Property Get Style()
        Style = iStyle
    End Property

    '//�����ҳ����

    Public Property Get PageParam()
        PageParam = sPageParam
    End Property

    '//�����ҳ��ť

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
    '//������� ����
    '//-------------------------------------------------------------------------
End Class
%>