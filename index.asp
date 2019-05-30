<!--#include file="app/class/blog.class.asp" -->
<%
OpenConn()
Dim W
Set  W = New wmBlog
W.run
%>