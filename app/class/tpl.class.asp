<%
'**********************************
'ASP模板引擎
'用法:	Set var=new sTemplate
'		[var.prop=vars]
'		[var.assign name,value]
'		var.display tplpath
'作者:	shirne
'修改完善:4jax
'日期:	2019/02/10
'**********************************
Class sTemplate
	Private oData, oType, oReg, oSql, oStm, oFso
	Private sApp, sTpl, sExt, sHtm, sFmt
	Private iStart,iQuery		'开始运行时间

	Private	htmPath,aCache,chePath
	Public  bHtm,filePath,iChe,sChr

	Private Sub Class_Initialize
		iStart	= Timer()
		sApp	= AppPath
		sTpl	= AppPath & TEMPLATE_PATH & "/"
		sExt	= ".html"
		sChr	= "utf-8"				'编码
		sFmt	= "\w\d\/\\\-\[\]\.\u00A1-\uFFFF"		'变量格式化允许的字符,不能有}
		iChe	= CACHE_TIME		'缓存时间以秒计
		'bHtm	= HTML_OPEN		'是否生成静态,生成静态时必须指定filepath
		iQuery	= 0				'自定义的sql查询次数

		'htmPath	= AppPath&"html/"		'静态文件路径
		chePath	= AppPath&"app/cache/"		'缓存文件路径

		Set oData	= Server.CreateObject("Scripting.Dictionary")	'存放注册数据
		Set oType	= Server.CreateObject("Scripting.Dictionary")	'存放数据类型
		Set oStm	= Server.CreateObject("ADODB.Stream")
		Set oFso	= Server.CreateObject("Scripting.FileSystemObject")
		Set oReg	= new RegExp
		oReg.Global = True
		'CheckPath htmPath
		'Response.write chePath
		'CheckPath chePath
	End Sub

	Private Sub Class_Terminate
		oData.RemoveAll
		oType.RemoveAll
		sHtm		= ""
		Set oData	= Nothing
		Set oType	= Nothing
		Set oStm	= Nothing
		Set oFso	= Nothing
		Set oReg	= Nothing
	End Sub

	'注册变量或obj或数组
	Public Sub assign(sName,obj)
		If oData.Exists(sName) Then
			oData(sName)=obj
			oType(sName)=vType(obj)
		Else
			oData.Add sName,obj
			oType.Add sName,vType(obj)
		End If
	End Sub

	'显示
	Public Sub Display(fTpl)	 
		Dim n,i,j,k,fPathfPath,iTmp
		j		= -1
		fPath	= chePath&Server.URLEncode(GetFileStr)&".cache"
		If iChe>0 Then	'获取缓存
			If oFso.FileExists(Server.MapPath(fPath)) Then
				Set f=oFso.GetFile(Server.MapPath(fPath))
				If DateDiff("s",f.DateLastModified,Now)<iChe Then
					sHtm=ReadFile(fPath)
				End If
			End If
		End If
		If sHtm="" Then
			sHtm	= ReadFile(sTpl&fTpl)
			sHtm	= include(sHtm)

			If InStr(sHtm,"<nocache>")>0 Then
				i=InStr(sHtm,"<nocache>")
				j=0
				ReDim aCache(0)
				Do Until i<1
					ReDim Preserve aCache(j)
					k=InStr(i,sHtm,"</nocache>")
					If k<1 Then cErr(15)
					aCache(j)=Mid(sHtm,i+9,k-i-10)
					i=InStr(k,"<nocache>")
					If i>0 Then j=j+1
				Loop
			End If

			sHtm	= getCache(sHtm)

			sHtm	= iReplace(sHtm)
			sHtm	= analyTpl(sHtm)
			'sHtm	= iReplace(sHtm)
			If iChe>0 Then
				iTmp=sHtm
				If j>-1 Then
					i=1
					For k=0 To j
						i=InStr(i,iTmp,"<nocache>")
						n=InStr(i,iTmp,"</nocache>")
						If i<0 Or n<0 Then Exit For
						iTmp=Replace(iTmp,Mid(iTmp,i+9,n-i-10),aCache(k))
						i=n
					Next
					sHtm	= Replace(sHtm,"<nocache>","")
					sHtm	= Replace(sHtm,"</nocache>","")
				End If
				SaveFile fPath,iTmp
			End If
		Else
			If InStr(sHtm,"<nocache>")>0 Then
				sHtm	= iReplace(sHtm)
				sHtm	= analyTpl(sHtm)
				'sHtm	= iReplace(sHtm)
				sHtm	= Replace(sHtm,"<nocache>","")
				sHtm	= Replace(sHtm,"</nocache>","")
			End If
		End If
		'If CBol(bHtm) Then
			'CheckPath(getDir(htmPath&filePath))
			'SaveFile htmPath&filePath,sHtm
		'End If

		j=CCur(Timer()-iStart)
		'If j<1 Then j="0"&j
		sHtm=Replace(sHtm,"{runtime}","Processed "&j&"s")
		Echo sHtm
	End Sub

	Public Sub ClearCache
		On Error Resume Next
		If oFso.FolderExists(Server.MapPath(chePath)) Then
			oFso.DeleteFolder Server.MapPath(chePath)
		End If
		If Err Then cErr 32
	End Sub

	Private Function getCache(sCont)
		Dim i,ii,iii
		i=InStr(sCont,"<cache")
		If i<1 Then
			getCache=sCont
		Else
			Dim j,sLabel,sTmp,oAtt,cPath,sTemp
			Do
				ii=InStr(i,sCont,"</cache>")
				If ii<1 Then cErr 16
				j=InStr(i,sCont,">")
				sLabel=Mid(sCont,i+6,j-i-6)
				sTemp=Mid(sCont,j+1,ii-j-1)
				Set oAtt=analyLabel(sLabel)
				If oAtt.Exists("name") Then
					CheckPath chePath&"global/"
					cPath=chePath&"global/"&oAtt("name")&".cache"
					If oFso.FileExists(Server.MapPath(cPath)) Then
						If oAtt.Exists("time") Then
							If DateDiff("h",(oFso.getFile(Server.MapPath(cPath))).DateLastModified,Now)<oAtt("time") Then
								sTmp=ReadFile(cPath)
							End If
						Else
							sTmp=ReadFile(cPath)
						End If
					End If
					If sTmp="" Then
						sTmp=sTemp

						sTmp	= iReplace(sTmp)
						sTmp	= analyTpl(sTmp)
						SaveFile cPath,sTmp
					End If
					sCont=Replace(sCont,"<cache"&sLabel&">"&sTemp&"</cache>",sTmp)
					i=InStr(i+Len(sTmp),sCont,"<cache")
					sTmp=""
				Else
					i=InStr(ii,sCont,"<cache")
				End If
			Loop Until i<1

			getCache=sCont
		End If
	End Function

	Private Function GetFileStr() 
		Dim strTemps 
		strTemps = strTemps & Request.ServerVariables("URL") 
		If Trim(Request.QueryString) <> "" Then 
			strTemps = strTemps & "?" & Trim(Request.QueryString) 
		Else
			strTemps = strTemps 
		End If
		GetFileStr = strTemps 
	End Function

	Private Function include(sContent)
		Dim Matches, Match, i
		include=sContent
		i=0
		oReg.Pattern="\{include\s*\(([\'\""])?([\w\.\d\/\\]+)\1\)\}"
		Do
			Set Matches=oReg.Execute(sContent)
			For Each Match In Matches
				include=Replace(include,Match.Value,ReadFile(sTpl&Match.SubMatches(1)))
			Next
			i=i+1
		Loop While Matches.Count>0 And i<5	'最深5层包含
		If Matches.Count>0 Then
			include=oReg.Replace(include,"")
		End If
	End Function

	Private Sub SaveFile(ByVal tpl,html)
		tpl = Server.MapPath(tpl)
		oStm.Type	= 2
		oStm.Mode	= 3
		oStm.CharSet= sChr
		oStm.Open
		oStm.WriteText html
		oStm.SetEOS
		oStm.SaveToFile tpl,2
		oStm.Close
	End Sub

	Private Function parseInt(n)
	  parseInt = CLng(n)
	End Function

	Private Function ReadFile(ByVal tpl)
		tpl = Server.MapPath(tpl)
		oStm.Type	= 2
		oStm.Mode	= 3
		oStm.CharSet= sChr
		oStm.Open
		If oFso.FileExists(tpl) Then
			oStm.LoadFromFile tpl
			ReadFile=oStm.ReadText
			oStm.Flush
			oStm.Close
		Else
			cErr 1
		End If
	End Function

	Private Function iReplace(sHtm)
		Dim n, oMth, Match, iTmp

		oReg.Pattern="\{\$apppath\}":sHtm=oReg.Replace(sHtm,AppPath)
		oReg.Pattern="\{\$filepath\}":sHtm=oReg.Replace(sHtm,AppPath & FILE_UP_PATH)
		oReg.Pattern="\{\$template\}":sHtm=oReg.Replace(sHtm,sTpl)
		oReg.Pattern="\{\$source\}":sHtm=oReg.Replace(sHtm,sTpl&"resource/")
		oReg.Pattern="\{\$SiteName\}":sHtm=oReg.Replace(sHtm,SiteName)
		oReg.Pattern="\{\$webtitle\}":sHtm=oReg.Replace(sHtm,webtitle)
		oReg.Pattern="\{\$webdesc\}":sHtm=oReg.Replace(sHtm,webdesc)
		oReg.Pattern="\{\$SiteKeyWords\}":sHtm=oReg.Replace(sHtm,SiteWords)
		oReg.Pattern="\{\$Icp\}":sHtm=oReg.Replace(sHtm,Icp)
		oReg.Pattern="\{\$ifile\}":sHtm=oReg.Replace(sHtm,ifile)
		oReg.Pattern="\{\$menu\}":sHtm=oReg.Replace(sHtm,Menu)

		oReg.Pattern="(\{[^{]+)\$apppath([^}]*\})":sHtm=oReg.Replace(sHtm,"$1"&AppPath&"$2")
		oReg.Pattern="(\{[^{]+)\$filepath([^}]*\})":sHtm=oReg.Replace(sHtm,"$1"&AppPath & FILE_UP_PATH&"$2")
		oReg.Pattern="(\{[^{]+)\$template([^}]*\})":sHtm=oReg.Replace(sHtm,"$1"&sTpl&"$2")
		oReg.Pattern="(\{[^{]+)\$source([^}]*\})":sHtm=oReg.Replace(sHtm,"$1"&sTpl&"resource/"&"$2")
		oReg.Pattern="(\{[^{]+)\$SiteName([^}]*\})":sHtm=oReg.Replace(sHtm,"$1"&SiteName&"$2")
		oReg.Pattern="(\{[^{]+)\$webtitle([^}]*\})":sHtm=oReg.Replace(sHtm,"$1"&webtitle&"$2")
		oReg.Pattern="(\{[^{]+)\$webdesc([^}]*\})":sHtm=oReg.Replace(sHtm,"$1"&webdesc&"$2")
		oReg.Pattern="(\{[^{]+)\$SiteKeyWords([^}]*\})":sHtm=oReg.Replace(sHtm,"$1"&SiteWords&"$2")
		oReg.Pattern="(\{[^{]+)\$Icp([^}]*\})":sHtm=oReg.Replace(sHtm,"$1"&Icp&"$2")
		oReg.Pattern="(\{[^{]+)\$ifile([^}]*\})":sHtm=oReg.Replace(sHtm,"$1"&ifile&"$2")
		For Each n In oData		
			If oType(n)=0 Then
				oReg.Pattern="\{var\:"&n&"((?:\|["& sFmt &"]+)*)?\}"
				Set oMth=oReg.Execute(sHtm)
				For Each Match In oMth
					If Match.SubMatches.Count>0 Then
						sHtm=Replace(sHtm,Match.Value,fmtVar(oData(n),Match.SubMatches(0)))
					Else
						sHtm=Replace(sHtm,Match.Value,oData(n))
					End If
				Next
				'替换标签内变量
				oReg.Pattern="\{[^{]+@var:"&n&"[^}]*\}"
				Set oMth=oReg.Execute(sHtm)
				For Each Match In oMth
					sHtm=Replace(sHtm,Match.Value,Replace(Match.Value,"@var:"&n,oData(n)))
				Next
			End If
		Next
		'Response.write sHtm:Response.end
		oReg.Pattern="\{\$([\d\w]+)\.([\d\w]+)((?:\|["& sFmt &"]+)*)?\}"
		Set oMth=oReg.Execute(sHtm)
		For Each Match In oMth
			If Match.SubMatches.Count<=2 Then iTmp="" Else iTmp=Match.SubMatches(2)
			sHtm=Replace(sHtm,Match.Value,getValue(Match.SubMatches(0),Match.SubMatches(1),iTmp))
		Next
		'替换标签内变量
		oReg.Pattern="\{[^{]+\$([\d\w]+)\.([\d\w]+)[^}]*\}"
		Set oMth=oReg.Execute(sHtm)
		For Each Match In oMth
			If Match.SubMatches.Count<=2 Then iTmp="" Else iTmp=Match.SubMatches(2)
			sHtm=Replace(sHtm,Match.Value,_
			Replace(Match.Value,"$"&Match.SubMatches(0)&"."&Match.SubMatches(1),_
			getValue(Match.SubMatches(0),Match.SubMatches(1),iTmp)))
		Next 
		iReplace=sHtm
	End Function

	'解析模板
	Private Function analyTpl(ByVal sCont)
		Dim i,sTag,sLabel,iEnd,iDiv,sTemp,ilayer
		Dim iPos,iRtn,iTmp,j,k,l,ii,iii,oAtt,sTmp,sLbl
		i=InStr(sCont,"{")
		Do While i>0
			'标签的内容
			sLabel=Mid(sCont,i+1,InStr(i,sCont,"}")-i-1)
			ii=InStr(sLabel,":")
			If ii>0 Then	'跳过其它标签
				'标签名
				sTag=Left(sLabel,ii-1)
				If InStr("|if|fn|for|foreach|loop|sql|","|"&sTag&"|")>0 Then
					'标签结束位置
					iEnd=InStr(i,sCont,"{/"&sTag&"}")
					If iEnd <1 Then cErr(3)
					'标签模板
					sTemp=Mid(sCont,i+Len(sLabel)+2,iEnd-i-Len(sLabel)-2)
					'是否存在嵌套
					iDiv=InStr(sTemp,"{"&sTag&":")
					ilayer=0
					Do While iDiv>0
						ilayer=ilayer+1  '层数加1
						iEnd=InStr(iEnd+1,sCont,"{/"&sTag&"}")
						If iEnd<1 Then cErr 3
						sTemp=Mid(sCont,i+Len(sLabel)+2,iEnd-i-Len(sLabel)-2)
						iDiv=InStr(iDiv+1,sTemp,"{"&sTag&":")
					Loop

					'将变量缓存,以防后期被改变
					sTmp=sTemp
					sLbl=sLabel
				End If

				iRtn=""	'解析返回值
				Select Case sTag
				Case "if"
					If ilayer=0 Then	'无嵌套时执行解析	
						If InStr(sTemp,"{elseif:")>0 Then			
							iTmp=Split(sTemp,"{elseif:")
							k=UBound(iTmp)
							If judge(Mid(sLabel,4))  Then
								iRtn=iTmp(0)			
							Else
								For j=1 To k					
									If judge(Left(iTmp(j),InStr(iTmp(j),"}")-1)) = True Then
										iRtn=Mid(iTmp(j),InStr(iTmp(j),"}")+1)
										If InStr(iRtn,"{else")>0 then
										   iRtn=Split(iRtn,"{else")(0)
										End If	
										Exit For
									End If
								Next				
							End If
							If iRtn="" And InStr(iTmp(k),"{else}")>0 Then
								iRtn=analyTpl(Split(iTmp(k),"{else}")(1))
							Else
								iRtn=analyTpl(iRtn)
							End If

						ElseIf InStr(sTemp,"{else}")>0 Then						
							iTmp=Split(sTemp,"{else}")
							s = judge(Mid(sLabel,4))
							If judge(Mid(sLabel,4)) = True Then
								iRtn=analyTpl(iTmp(0))
							Else
								iRtn=analyTpl(iTmp(1))
							End If
						Else
							If judge(Mid(sLabel,4)) = True Then
								iRtn=analyTpl(sTemp)								
							End If
						End If
					Else		'有嵌套时循环解析
						sTemp=Replace(sTemp,"{else}","{elseif:1=1}")
						ii=InStr(sTemp,"{elseif:")
						k=InStr(sTemp,"{if:")
						If judge(Mid(sLabel,4)) = True Then
							If ii<=0 Then
								iRtn=analyTpl(sTemp)
							ElseIf k>ii Then		'隐含条件 ii>0						  
								iRtn=analyTpl(Mid(sTemp,ii-1))								 
							Else		'隐含条件ii>0,k<ii
								iDiv=InStr(sTemp,"{/if}")
								Do Until InStr(k+1,Left(sTemp,iDiv),"{if:")<1
									k=InStr(k+1,sTemp,"{if:")
									iDiv=InStr(iDiv+1,sTemp,"{/if}")
									If iDiv<1 Then cErr(12)
								Loop
								iDiv=InStr(iDiv,sTemp,"{elseif:")
								If iDiv>0 Then
									iRtn=analyTpl(Left(sTemp,iDiv-1))
								Else
									iRtn=analyTpl(sTemp)
								End If
							End If
						ElseIf ii>0 Then	'不存在else或elseif,则整段已经被抛弃
							If k<ii Then	'隐含条件k>0
								iDiv=InStr(sTemp,"{/if}")
								Do Until InStr(k+1,Left(sTemp,iDiv),"{if:")<1
									k=InStr(k+1,sTemp,"{if:")
									iDiv=InStr(iDiv+1,sTemp,"{/if}")
									If iDiv<1 Then cErr(12)
								Loop
								ii=InStr(iDiv,sTemp,"{elseif:")
							End If
							If ii>0 Then	'与上面ii>0不同,如果首段if排除后已经没有else,也抛弃
								sLabel=Mid(sTemp,ii+8,InStr(ii,sTemp,"}")-ii-8)

								Do Until judge(sLabel)	'当前elseif内标签不为真
									k=InStr(ii,sTemp,"{if:")
									iDiv=InStr(ii,sTemp,"{/if}")
									ii=InStr(ii+1,sTemp,"{elseif:")
									If k>0 And k<ii Then	'下一个else前有if
										Do Until InStr(k+1,Left(sTemp,iDiv),"{if:")<1
											k=InStr(k+1,sTemp,"{if:")
											iDiv=InStr(iDiv+1,sTemp,"{/if}")
											If iDiv<1 Then cErr(12)
										Loop
										ii=InStr(iDiv,sTemp,"{elseif:")
									End If
									If ii<1 Then Exit Do
									sLabel=Mid(sTemp,ii+8,InStr(ii,sTemp,"}")-ii-8)
								Loop

								'寻找当前内容段作为返回
								If ii>0 Then
									iii=InStr(ii,sTemp,"}")	'定位当前标签结束位置
									k=InStr(ii,sTemp,"{if:")
									iDiv=InStr(ii,sTemp,"{/if}")
									ii=InStr(ii,sTemp,"{elseif:")
									If k>0 And k<ii Then	'下一个else前有if
										Do Until InStr(k+1,Left(sTemp,iDiv),"{if:")<1
											k=InStr(k+1,sTemp,"{if:")
											iDiv=InStr(iDiv+1,sTemp,"{/if}")
											If iDiv<1 Then cErr(12)
										Loop
										ii=InStr(iDiv,sTemp,"{elseif:")
									End If
									If ii<1 Then
										iRtn=analyTpl(Mid(sTemp,iii+1))
									Else
										iRtn=analyTpl(Mid(sTemp,iii+1,ii-2))
									End If
								End If
							End If
						End If
					End If
				Case "fn"
					Set oAtt=analyLabel(sLabel)
					If oAtt.Exists("func") Then
						Set k=GetRef(oAtt("func"))
						If oAtt.Exists("args") Then
							ii=Split(oAtt("args"),",")
							If oAtt.Exists("argtype") Then
								iii=Split(oAtt("argtype")&",,,,,",",")
							Else
								iii=Split(",,,,,",",")
							End If
							For j=0 To UBound(ii)
								Select Case LCase(iii(5))
								Case "i"
									ii(j)=parseInt(ii(j))
								Case "f"
									If IsNumeric(ii(j)) Then ii(j)=CDbl(ii(j)) Else ii(j)=0
								Case "b"
									ii(j)=CBol(ii(j))
								Case Else
									ii(j)=decode(ii(j),True)
								End Select
								If j>4 Then Exit For
							Next
							Select Case UBound(ii)
							Case 0
								iRtn=k(sTemp,ii(0))
							Case 1
								iRtn=k(sTemp,ii(0),ii(1))
							Case 2
								iRtn=k(sTemp,ii(0),ii(1),ii(2))
							Case 3
								iRtn=k(sTemp,ii(0),ii(1),ii(2),ii(3))
							Case 4
								iRtn=k(sTemp,ii(0),ii(1),ii(2),ii(3),,ii(4))
							End Select
						Else
							iRtn=k(sTemp)
						End If
						iRtn=analyTpl(iRtn)
					End If
				Case "for"
					Set oAtt=analyLabel(sLabel)
					If oAtt.Exists("var") And oAtt.Exists("to") Then
						oAtt("to")=parseInt(oAtt("to"))
						If oAtt.Exists("from") Then oAtt("from")=parseInt(oAtt("from")) Else oAtt.Add "from",1
						If oAtt.Exists("step") Then k=ParseInt(oAtt("step")) Else k=1
						For j=ParseInt(oAtt("from")) To ParseInt(oAtt("to")) Step k
							k = Replace(sTemp,"{@"&oAtt("var")&"}",j)
							oReg.Pattern="(\{[^\{]+)@"&oAtt("var")&"([^\.\}]*\})"
							iRtn = iRtn & oReg.Replace(k,"$1"&j&"$2")
						Next
						iRtn=analyTpl(iRtn)
					End If
				Case "foreach"
					Set oAtt=analyLabel(sLabel)
					If oAtt.Exists("var") And oAtt.Exists("name") Then
						If oData.Exists(oAtt("name")) Then
							If oType(oAtt("name"))=2 Or oType(oAtt("name"))=4 Then
								For Each j In oData(oAtt("name"))
									k=Replace(sTemp,"{@"&oAtt("var")&".name}",j)
									k=Replace(k,"{@"&oAtt("var")&".value}",j)

									oReg.Pattern="(\{[^\{]+)@"&oAtt("var")&"\.name([^\}]*\})"
									k = oReg.Replace(k,"\1"&j&"\2")
									oReg.Pattern="(\{[^\{]+)@"&oAtt("var")&"\.value([^\}]*\})"
									iRtn = iRtn & oReg.Replace(k,"$1"&oData(oAtt("name"))(j)&"$2")
								Next
								iRtn=analyTpl(iRtn)
							End If
						End If
					End If
				Case "loop"
					Set oAtt=analyLabel(sLabel)					
					If oAtt.Exists("name") Then
						If oData.Exists(oAtt("name")) Then 
						'oData(oAtt("name")).MoveFirst
							For ii=1 To Len(sTemp)
								l=InStr(ii,sTemp,"{loopelse}")
								If l>0 Then
									iDiv=InStr(ii,sTemp,"{loop:")
									If iDiv>l Or iDiv<1 Then
										sTemp=Left(sTemp,l-1)&Replace(sTemp,"{loopelse}","{loopelseMARK}",l,1)
										Exit For
									Else
										ii=InStr(ii,sTemp,"{/loop}")
										Do Until iDiv<1
											If ii<1 Then cErr(13)
											iDiv=InStr(iDiv+1,sTemp,"{loop:")
											If iDiv>0 Then ii=InStr(ii+1,sTemp,"{/loop}")
										Loop
									End If
								End If
							Next

							If oType(oAtt("name"))=3 Then
								If oAtt.Exists("limit") Then
									If InStr(oAtt("limit"),",")<1 Then oAtt("limit")="1,"&oAtt("limit")
									oAtt("limit")=Split(oAtt("limit"),",")
									oAtt("limit")(0)=parseInt(oAtt("limit")(0))
									k=parseInt(oAtt("limit")(1))
								Else
									k=oData(oAtt("name")).RecordCount	
									'Response.write k
								End If
								If oAtt.Exists("count") Then k=ParseInt(oAtt("count"))
								If k>100 Then k=100	'最多输出100条
								iii=Split(sTemp&"{loopelseMARK}","{loopelseMARK}")
								If oData(oAtt("name")).EOF Then
									iRtn=iii(1)
								Else
								    'Response.write oAtt("name")&iii(0)
									ii=oData(oAtt("name")).AbsolutePosition	'记录rscordset起始位置
									If oAtt.Exists("limit") Then
										If oData(oAtt("name")).RecordCount>oAtt("limit")(0) Then
											oData(oAtt("name")).AbsolutePosition=oAtt("limit")(0)
										Else
											oData(oAtt("name")).AbsolutePosition=oData(oAtt("name")).RecordCount
										End If
									End If
									For j=1 To k
										iRtn=iRtn & Replace(Replace(subReplace(iii(0),oData(oAtt("name")),oAtt("name")),"{@"&oAtt("name")&".@index}",j),"@"&oAtt("name")&".@index",j)
										oData(oAtt("name")).MoveNext
										If oData(oAtt("name")).EOF Then oData(oAtt("name")).AbsolutePosition=ii:Exit For
									Next
								End If	
								'Response.write iRtn
								iRtn=analyTpl(iRtn)
								
							End If
						End If
					End If
				Case "sql"
					Set oAtt=analyLabel(sLabel)
					If oAtt.Exists("name") And oAtt.Exists("table") Then
						If LCase(oAtt("table"))<>"admin" Then

							For ii=1 To Len(sTemp)
								l=InStr(ii,sTemp,"{sqlelse}")
								If l>0 Then
									iDiv=InStr(ii,sTemp,"{sql:")
									If iDiv>l Or iDiv<1 Then
										sTemp=Left(sTemp,l-1)&Replace(sTemp,"{sqlelse}","{sqlelseMARK}",l,1)
										Exit For
									Else
										ii=InStr(ii,sTemp,"{/sql}")
										Do Until iDiv<1
											If ii<1 Then cErr(14)
											iDiv=InStr(iDiv+1,sTemp,"{sql:")
											If iDiv>0 Then ii=InStr(ii+1,sTemp,"{/sql}")
										Loop
									End If
								End If
							Next

							Set k=New MakeSQL
							k.Table(oAtt("table"))
							If oAtt.Exists("field") Then k.field(Split(oAtt("field"),","))
							If oAtt.Exists("where") Then k.where(Array(decode(oAtt("where"),True)))
							If oAtt.Exists("limit") Then
								If InStr(oAtt("limit"),",")<1 Then oAtt("limit")="1,"&oAtt("limit")
								oAtt("limit")=Split(oAtt("limit"),",")
								k.limit oAtt("limit")(0),oAtt("limit")(1)
							End If
							If oAtt.Exists("order") Then k.order(Split(oAtt("order"),","))
							Set l=k.CreateSQL("select",True)
							iQuery=iQuery+1
							iii=Split(sTemp&"{sqlelseMARK}","{sqlelseMARK}")
							If l.EOF Then
								iRtn=iii(1)
							Else
								If oAtt.Exists("count") Then ii=ParseInt(oAtt("count")) Else ii=l.RecordCount
								If ii>100 Then ii=100	'最多输出100条
								For j=1 To ii
									iRtn=iRtn & Replace(Replace(subReplace(iii(0),l,oAtt("name")),"{@"&oAtt("name")&".@index}",j),"@"&oAtt("name")&".@index",j)
									l.MoveNext
									If l.EOF Then Exit For
								Next
							End If
							iRtn=analyTpl(iRtn)
						End If
					End If
				Case Else
					iRtn="{"
				End Select
				'sCont= Replace(sCont,"{"&sLbl&"}"&sTmp&"{/"&sTag&"}",iRtn)
				sCont= Left(sCont,i-1)& Replace(sCont,"{"&sLbl&"}"&sTmp&"{/"&sTag&"}",iRtn,i,1)
				i=i+Len(iRtn)
			Else
				i=i+Len(sLabel)+1
			End If
			i=InStr(i,sCont,"{")
		Loop
		analyTpl=sCont
	End Function

	'获取obj健值
	Private Function getValue(sObj,sKey,sFlt)
		getValue=""
		Select Case sObj
		Case "query"
			getValue=Request.QueryString(sKey)
		Case "form"
			getValue=Request.Form(sKey)
		Case "cookie"
			getValue=Request.Cookies(sKey)
		Case "server"
			getValue=Request.ServerVariables(sKey)
		Case "session"
			getValue=Session(sKey)
		Case Else
			If oData.Exists(sObj) Then
				If oType(sObj)=2 Then
					If oData(sObj).Exists(sKey) Then getValue=oData(sObj)(sKey)
				ElseIf oType(sObj)=4 Then
					getValue=oData(sObj)(sKey)
				ElseIf oType(sObj)=3 Then
					If Not IsEmpty(oData(sObj)(sKey)) Then getValue=oData(sObj)(sKey)
				End If
			End If
			If IsNull(getValue) Then getValue=""
		End Select
		If sFlt<>"" Then
			getValue=fmtVar(getValue,sFlt)
		End If
	End Function

	'替换obj值
	Private Function subReplace(ByVal Tpl,obj,oName)
		Dim oMth,Match
		oReg.Pattern="\{@"& oName &"\.([\w\d]+)((?:\|["& sFmt &"]+)*)?\}"		
		Set oMth=oReg.Execute(Tpl)
		oReg.Global = True
		For Each Match In oMth		
			If Match.SubMatches.Count<2 Then
				Tpl=Replace(Tpl,Match.Value,obj(Match.SubMatches(0)))
			Else
				Tpl=Replace(Tpl,Match.Value,fmtVar(obj(Match.SubMatches(0)),Match.SubMatches(1)))
			End If
		Next
		'替换标签内变量
		oReg.Pattern="\{[^{]+@"& oName &"\.([\w\d]+)[^}]*\}"
		Set oMth=oReg.Execute(Tpl)
		For Each Match In oMth		
			Tpl=Replace(Tpl,Match.Value,_
			Replace(Match.Value,"@"&oName&"."&Match.SubMatches(0),_
			obj(Match.SubMatches(0))))
		Next
		'Response.write "kkkkkkkkk"&tpl&"kkkkkkkkk"
		subReplace=Tpl
	End Function

	'判断if条件
	Private Function judge(str)
	
		Dim oMth,a,b,c
		judge=True
		oReg.Pattern="^\s*(.+?)\s*(\=|\<\>|\<|\>|\>=|\<=|\!\=|\=\=)\s*(.+?)\s*$"
		Set oMth=oReg.Execute(str)
	  'Response.write Str
   'Response.write  oMth.Count:Response.end
		If oMth.Count<1 Then
			judge=CBol(str)
		Else
			a=Trim(oMth(0).SubMatches(0))
			b=oMth(0).SubMatches(1)
			c=Trim(oMth(0).SubMatches(2))
			 
			'If (IsNumeric(a) Or a="") And (IsNumeric(c) Or c="") Then
				'a=parseInt(a)
				'c=ParseInt(c)
			'End If
			 ' Response.write a
			 'Response.end

			Select Case b
			Case "=","=="
				If a<>c Then judge=False
			Case "<>","!="
				If a=c Then judge=False
			Case ">"
				If a<=c Then judge=False
			Case "<"
				If a>=c Then judge=False
			Case ">="
				If a<c Then judge=False
			Case "<="
				If a>c Then judge=False
			End Select
		End If
	End Function

	'格式化变量
	Private Function fmtVar(var,fmt)
		Dim iTmp,d,f
		iTmp=Split(fmt&"|||||","|")
		fmtVar=var
		Select Case LCase(iTmp(1))
		Case "fmtdate"	'格式化日期"YYYY"
			If IsDate(var) Then
				d=CDate(var)
				If LCase(iTmp(2))="kindly" Then
					f = Replace(LCase(iTmp(2)),"kindly",FmtTime(d,False))
				Else
					f = Replace(LCase(iTmp(2)),"yyyy",Year(d))
					f = Replace(f, "yy",	Right(Year(d),2))
					f = Replace(f, "mm",	Right("00"&Month(d),2))
					f = Replace(f, "m",		Month(d))
					f = Replace(f, "dd",	Right("00"&Day(d),2))
					f = Replace(f, "d",		Day(d))
					f = Replace(f, "hh",	Right("00"&Hour(d),2))
					f = Replace(f, "h",		Hour(d))
					f = Replace(f, "nn",	Right("00"&Minute(d),2))
					f = Replace(f, "n",		Minute(d))
					f = Replace(f, "ss",	Right("00"&Second(d),2))
					f = Replace(f, "s",		Second(d))
					f = Replace(f, "www",	weekdayname(weekday(d)))
					f = Replace(f, "ww",	Right(weekdayname(weekday(d)),1))
					f = Replace(f, "w",		weekday(d))
				End If
				fmtVar=f
			End If
		Case "cutstr"
			d=parseInt(iTmp(2))
			fmtVar=left(reHtml(fmtVar),d)
		Case "lcase"
			fmtVar=LCase(fmtVar)
		Case "ucase"
			fmtVar=UCase(fmtVar)
        Case "vurl"
		    fmtVar=vurl(fmtVar)
		Case "fmtnum"
			iTmp(3)=ParseInt(iTmp(3))
			If iTmp(2)="1" Then
				fmtVar=parseInt(fmtVar)
				If iTmp(3)=0 Or (iTmp(3)<Len(fmtVar) And CBol(iTmp(4))) Then iTmp(3)=Len(fmtVar)
				fmtVar=Right(String("0",iTmp(3))&fmtVar,iTmp(3))
			ElseIf iTmp(2)="2" Then
				If iTmp(3)=0 Or (iTmp(3)<Len(fmtVar) And CBol(iTmp(4))) Then iTmp(3)=Len(fmtVar)
				fmtVar=Left(fmtVar&String("0",iTmp(3)),iTmp(3))
			ElseIf iTmp(2)="3" Then
				fmtVar=Hex(parseInt(fmtVar))
				If iTmp(3)=0 Or (iTmp(3)<Len(fmtVar) And CBol(iTmp(4))) Then iTmp(3)=Len(fmtVar)
				fmtVar=Right(String("0",iTmp(3))&fmtVar,iTmp(3))
			ElseIf iTmp(2)="4" Then
				fmtVar=dHex(fmtVar)
				If iTmp(3)=0 Or (iTmp(3)<Len(fmtVar) And CBol(iTmp(4))) Then iTmp(3)=Len(fmtVar)
				fmtVar=Right(String("0",iTmp(3))&fmtVar,iTmp(3))
			End If
		Case "nohtml"
			fmtVar=reHtml(fmtVar)
		Case "html"
			fmtVar=HTMDecode(fmtVar)
		Case "escape"		
			If IsNull(fmtVar)  Then fmtVar ="":Exit Function  else fmtVar=Server.URLEncode(fmtVar)
		Case "unescape"
			fmtVar=URLDecode(fmtVar)
		Case "jscode"
			fmtVar=UTFEncode(fmtVar)
		Case "replace"
			fmtVar=Replace(fmtVar,iTmp(2),iTmp(3))
		Case "trip"
			fmtVar=html2txt(fmtVar)
		Case "filesize"
			fmtVar=convertSize(fmtVar)
		Case "fpic"
		    fmtVar=Mid(fmtVar,instr(fmtVar,"com/")+4)
		Case "url"
			fmtVar=HTMDecode(fmtVar)
		Case "default"
			If fmtVar="" Or IsEmpty(fmtVar) Or IsNull(fmtVar) Then fmtVar=iTmp(2)
		Case "iif"
			If CBol(fmtVar)  Then			
				fmtVar=iTmp(2)
			Else
				fmtVar=iTmp(3)
			End If
		End Select
		If IsNull(fmtVar) Then fmtVar=""
	End Function

	'解析标签属性
	Private Function analyLabel(sCont)
		Dim oTag,oMatch,oMth
		Set oTag=Server.CreateObject("Scripting.Dictionary")
		oReg.Pattern="\b([\w\d]+)\s*=\s*(['""])([\w\d\-\,\.\s\%\=\<\>\$]+)\2"
		Set oMatch=oReg.Execute(sCont)
		For Each oMth In oMatch
			If Not oTag.Exists(oMth.SubMatches(0)) Then
				oTag.Add oMth.SubMatches(0),decode(oMth.SubMatches(2),False)
			End If
		Next
		Set analyLabel=oTag
		Set oMatch=Nothing
	End Function

	Private Function decode(str,deep)
		decode=str
		If InStr(str,"%")<1 Then Exit Function
		decode=Replace(decode,"%22","""")
		decode=Replace(decode,"%27","'")
		If deep Then
			decode=Replace(decode,"%2C",",")
			decode=Replace(decode,"%25","%")
		End If
	End Function

	Private Function CheckPath(fPath)
		On Error Resume Next
		Dim path,i,cpath,rpath
		If oFso.FolderExists(fpath) Then'//看看是否已经存在目录
			CheckPath = True
			Exit Function
		Else
		cpath = ""
	    rpath = Server.MapPath("/")
		path=Split(Replace(Server.MapPath(fpath),rpath,""),"\")
		For i=0 To Ubound(path)
			If cPath="" Then
				cPath=rpath&path(i)
			Else
				cPath=cPath & "\" & path(i)
			End If
			If Not oFso.FolderExists(cPath) Then
				oFso.CreateFolder(cPath)
			End If
			If Err Then
				Err.Clear
				cErr 31
				CheckPath=False
			End If
		Next
		End If
		CheckPath=True
	End Function

	Private Function vType(obj)
		Select Case TypeName(obj)
		Case "Recordset"
			vType=3
		Case "Dictionary"
			vType=2
		Case "Variant()"
			vType=1
		Case Else
			If VarType(obj)=9 Then
				vType=4
			Else
				vType=0
			End If
		End Select
	End Function

	Function CBol(s)
		If s = "true" Or s = "1" Then
		    CBol = True 
		Else
		    CBol = False 
		End If
	End Function

	Function Echo(s)
	    Response.write s
	End Function

	Sub Die (s)
	    Response.write s:Response.End
	End Sub

	Private Sub cErr(Num)
		If IsNumeric(Num) Then
			Select Case Num
			Case 1:Die "模板不存在"
			Case 2:Die "标签不匹配"
			Case 3:Die "标签未闭合"
			Case 4:Die "标签嵌套错误"
			Case 12:Die "if标签未闭合"
			Case 13:Die "loop标签未闭合"
			Case 14:Die "sql标签未闭合"
			Case 15:Die "nocache标签未闭合"
			Case 16:Die "cache标签未闭合"
			Case 31:Die "创建文件夹失败,请检查权限"
			Case 32:Die "清除缓存失败,请检查权限"
			Case Else:Die "未知错误"
			End Select
		Else
			Die Num&"标签未闭合"
		End If
	End Sub
End Class
%>