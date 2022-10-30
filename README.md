
# RSS Class for ASP Classic



## Authors

- [@lasseedsvik](https://www.github.com/lasseedsvik)

## Usage

**To print out channel information and looping through items**

```asp
...
<%
    Dim objRSS: Set objRSS = GetRSS("https://blogs.microsoft.com/feed/", 10)

    If (IsObject(objRSS)) Then
        Dim chn: Set chn = objRSS.Channel

        ' Output channel info
        Response.Write("Title: " & chn.Title & "<br>" & vbCrlf)
        Response.Write("Description: " & chn.Description & "<br>" & vbCrlf)
        Response.Write("Link: " & chn.Link & "<br>" & vbCrlf)
        Response.Write("LastBuildDate: " & chn.LastBuildDate & "<br>" & vbCrlf)
        Response.Write("Generator: " & chn.Generator & "<br>" & vbCrlf)
        Response.Write("<hr>" & vbCrlf)

        ' Loop through items
        Dim lnk        
        For Each lnk in chn.Items
            Response.Write("<p>" & vbCrlf)            
            Response.Write("<strong>" & vbCrlf)
            Response.Write("<a href=""" & lnk.Link & """ target=""_blank"">" & vbCrlf)
            Response.Write(lnk.Title & vbCrlf)
            Response.Write("</a> " & vbCrlf)    
            Response.Write("</strong>" & vbCrlf)
            Response.Write(lnk.PubDate & vbCrlf)
            Response.Write("</p>" & vbCrlf)
            Response.Write("<p>" & vbCrlf & lnk.Description & "</p>" & vbCrlf)
            Response.Write("<hr>" & vbCrlf)
        Next		
        Set objRSS = Nothing
    Else
        Response.Write("Could not read RSS from " & objRSS.Url)
    End If
%>
...
```

**To print out the RSS from an url.** 

```asp
<!--#include file="RSS.asp"-->
<%
    Set objRSS = new RSS
    objRSS.OutputRSS("https://blogs.microsoft.com/feed/")
    Set objRSS = Nothing
%>
```

**NOTE: It's important when specifying the url to have a trailing "/" at the end.**

## The Do What the Fuck You Want to Public License

[![License: WTFPL](https://img.shields.io/badge/License-WTFPL-brightgreen.svg)](http://www.wtfpl.net/about/)    

