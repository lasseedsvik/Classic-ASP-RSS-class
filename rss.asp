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
		Private m_url
		Private m_channel
		
		Private Sub Class_Initialize()
			Set m_channel = Nothing
		End Sub
		
		public Property Get Url
			Url = m_url
		End Property

        public Property Get Limit 
            Limit = m_limit
        End Property

		Public Property Let Url(v)
			m_url = v

            On Error Resume Next

            Set xml = Server.CreateObject("Microsoft.XMLHTTP")
            xml.Open "GET", m_url, False
            xml.Send
            
            If Err.Number <> 0 Then
                Err.Raise vbObjectError + 2, "Xml Data", "Unable to parse Xml: " & 	xml.ParseError.Reason & " ErrorCode:" & xml.ParseError.ErrorCode, "", 0
            End if
            ResponseXML = xml.ResponseText

            Set doc = CreateObject("MSXML2.DOMDocument")
            doc.LoadXML(ResponseXML)

            Set channelNode = doc.GetElementsByTagName("channel")            
            Set chanitem = channelNode(0)
            
            ' Set channel properties
            Set m_channel = New RSSChannel
            m_channel.Title = chanitem.SelectSingleNode("title").Text
            m_channel.Description = chanitem.SelectSingleNode("description").Text
            m_channel.Link = chanitem.SelectSingleNode("link").Text
            m_channel.Category = chanitem.SelectSingleNode("category").Text
            m_channel.Language = chanitem.SelectSingleNode("language").Text
            m_channel.LastBuildDate = chanitem.SelectSingleNode("lastBuildDate").Text
            m_channel.Generator = chanitem.SelectSingleNode("generator").Text
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

                If Count >= m_limit And m_limit > 0 Then : Exit For

                m_channel.AddItem(ri)
                Set ri = Nothing                
            Next

            Set doc = Nothing
		End Property
		
		Public Property Get Channel
			Set Channel = m_channel
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
            
            If Err.Number <> 0 Then
                Err.Raise vbObjectError + 2, "Xml Data", "Unable to parse Xml: " & 	xml.ParseError.Reason & " ErrorCode:" & xml.ParseError.ErrorCode, "", 0
            End if
        End Sub
		
	End Class 
	
	' Class RSSChannel
	Class RSSChannel 
		Private m_items
		
		Public Title
		Public Description
		Public Link
        Public Category
        Public Language
        Public LastBuildDate
        Public Generator
		
		Private Sub Class_Initialize()
			Set m_items = Server.CreateObject("Scripting.Dictionary")
		End Sub
		
		Private Sub Class_Terminate()
			Set m_items = Nothing
		End Sub
		
		Public Sub AddItem(v)
			m_items.Add m_items.Count, v
		End Sub
		
		Public Property Get Items
			Items = m_items.Items
		End Property
		
		Public Property Let Items(v)
			Set m_items = v
		End Property
	
	End Class 
	
	' Class RSSItem
	Class RSSItem
		Public Title
		Public Description
		Public Link
        Public PubDate	
	End Class 
%>