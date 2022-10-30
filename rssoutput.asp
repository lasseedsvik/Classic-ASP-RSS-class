<!--#include file="RSS.asp"-->
<%
    Set objRSS = new RSS
    objRSS.OutputRSS("https://blogs.microsoft.com/feed/")
    Set objRSS = Nothing
%>