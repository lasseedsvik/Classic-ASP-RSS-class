<%
	' Get RSS feed
	Function GetRSS(url, limit)
        'Limit 0 = All

		Dim objRSS: Set objRSS = new RSS
		objRSS.Url = url
			
		Set GetRSS = objRSS		
	End Function
	
	'Class RSS	
	Class RSS
		Private url_
		Private channel_
		
		Private Sub Class_Initialize()
			Set channel_ = Nothing
		End Sub
		
		public Property Get Url
			Url = url_
		End Property

        public Property Get Limit 
            Limit = limit_
        End Property
		
		Public Property Let Url(v)
			url_ = v

            On Error Resume Next

            Set xml = Server.CreateObject("Microsoft.XMLHTTP")
            xml.Open "GET", url_, False
            xml.Send
            
            If Err.Number <> 0 Then
                Err.Raise vbObjectError + 2, "Xml Data", "Unable to parse Xml: " & 	xml.ParseError.Reason & " ErrorCode:" & xml.ParseError.ErrorCode, "", 0
            End if
            ResponseXML = xml.ResponseText

            Set doc = CreateObject("MSXML2.DOMDocument")
            doc.LoadXML(ResponseXML)

           ' Read channel info
            Set channel_ = New RSSChannel
            channel_.Title = channelNode.SelectSingleNode("title").Text
            channel_.Description = channelNode.SelectSingleNode("description").Text
            channel_.Link = channelNode.SelectSingleNode("link").Text

            Set doc = CreateObject("MSXML2.DOMDocument")
            doc.LoadXML(ResponseXML)

            Set channelNode = doc.GetElementsByTagName("channel")            
            Set chanitem = channelNode(0)
            
            ' Set channel properties
            Set channel_ = New RSSChannel
            channel_.Title = chanitem.SelectSingleNode("title").Text
            channel_.Description = chanitem.SelectSingleNode("description").Text
            channel_.Link = chanitem.SelectSingleNode("link").Text
            channel_.Category = chanitem.SelectSingleNode("category").Text
            channel_.Language = chanitem.SelectSingleNode("language").Text
            channel_.LastBuildDate = chanitem.SelectSingleNode("lastBuildDate").Text
            channel_.Generator = chanitem.SelectSingleNode("generator").Text
            Set channelNode = Nothing
            
            ' Set items
            Set items = doc.GetElementsByTagName("channel/item")
            
            For inum = 0 To items.Length-1 
                Set curitem = items(inum)

                title = curitem.SelectSingleNode("title").Text
                description = curitem.SelectSingleNode("description").Text
                link = curitem.SelectSingleNode("link").Text
                pubDate = curitem.SelectSingleNode("pubDate").Text
                image = curitem.SelectSingleNode("image").Text
                
                Dim ri: Set ri = New RSSItem
                ri.Title = curitem.SelectSingleNode("title").Text
                ri.Description = curitem.SelectSingleNode("description").Text
                ri.Link = curitem.SelectSingleNode("link").Text
                ri.PubDate = curitem.SelectSingleNode("pubDate").Text
                
                Count = Count + 1

                If Count >= limit_ And limit_ > 0 Then : Exit For

                channel_.AddItem(ri)
                Set ri = Nothing                
            Next

            Set doc = Nothing
		End Property
		
		Public Property Get Channel
			Set Channel = channel_
		End Property

        Public Sub OutputRSS(url)
            On Error Resume Next

            Dim objSrvHTTP
            Set objSrvHTTP = Server.CreateObject ("Msxml2.ServerXMLHTTP.6.0") 
            objSrvHTTP.Open "GET", url, False
            objSrvHTTP.Send
            
            Response.ContentType = "text/xml"            
            Response.CharSet = "UTF-8"
            Response.CodePage = 65001

            Response.Write(objSrvHTTP.ResponseText)
            
            If Err.Number <> 0 Then ShowError()
        End Sub
		
	End Class 
	
	' Class RSSChannel
	Class RSSChannel 
		Private items_
		
		Public Title
		Public Description
		Public Link
        Public Category
        Public Language
        Public LastBuildDate
        Public Generator
		
		Private Sub Class_Initialize()
			Set items_ = Server.CreateObject("Scripting.Dictionary")
		End Sub
		
		Private Sub Class_Terminate()
			Set items_ = Nothing
		End Sub
		
		Public Sub AddItem(v)
			items_.Add items_.Count, v
		End Sub
		
		Public Property Get Items
			Items = items_.Items
		End Property
		
		Public Property Let Items(v)
			Set items_ = v
		End Property
	
	End Class 
	
	' Class CRSSItem
	Class RSSItem
		Public Title
		Public Description
		Public Link
        Public PubDate	
	End Class 
%>