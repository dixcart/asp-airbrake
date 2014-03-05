<%
Class AirBrakeClass

	private error
	private constructed
	private apiKey
	private xml
	private environment

	Private Sub Class_Initialize()

		constructed = false
		Set xml = CreateObject("MSXML2.DOMDocument.3.0")

	End Sub

	Public Default Function Construct(strAPIKey, strEnvironment)

		constructed = true
		apiKey = strAPIKey
		environment = strEnvironment
		Set error = Server.GetLastError()

		Set objProcessing = xml.createProcessingInstruction("xml","version=""1.0""")
		xml.appendChild(objProcessing)
		Set objRoot = xml.createElement("notice")
		objRoot.setAttribute "version", "2.3"

		objRoot.appendChild(AddTextNode("api-key", apiKey))

		'Notifier node
		Set objNotifierNode = xml.createElement("notifier")
		objNotifierNode.appendChild(AddTextNode("name", "Airbrake ASP Notifier"))
		objNotifierNode.appendChild(AddTextNode("version", "0.1"))
		objNotifierNode.appendChild(AddTextNode("url", "http://itinsurrey.co.uk"))
		objRoot.appendChild(objNotifierNode)
		Set objNotifierNode = Nothing

		'error node
		Set objErrorNode = xml.createElement("error")
		objErrorNode.appendChild(AddTextNode("class", "ASP Error"))
		objErrorNode.appendChild(AddTextNode("message", error.Description))
		Set objBacktraceNode = xml.createElement("backtrace")
		Set objLineNode = xml.createElement("line")
		objLineNode.setAttribute "file", error.File
		objLineNode.setAttribute "number", error.Line
		objBacktraceNode.appendChild(objLineNode)
		Set objLineNode = Nothing
		objErrorNode.appendChild(objBacktraceNode)
		Set objBacktraceNode = Nothing
		objRoot.appendChild(objErrorNode)
		Set objErrorNode = Nothing

		'request node
		Set objRequestNode = xml.createElement("request")
		If Request.ServerVariables("HTTP_X_ORIGINAL_URL") <> "" Then
			strUrl = Request.ServerVariables("HTTP_X_ORIGINAL_URL")
		Else
			strUrl = Request.ServerVariables("URL")
		End If
		strUrl = Request.ServerVariables("HTTP_HOST") & strUrl
		If (Request.ServerVariables("HTTPS") = "off") Then
			strUrl = "http://" & strUrl
		Else
			strUrl = "https://" & strUrl
		End If
		objRequestNode.appendChild(AddTextNode("url", strUrl))
		Set objParams = xml.createElement("params")
		For each item in Request.QueryString
			objParams.appendChild(VarNode(item, Request.QueryString(item)))
		Next
		For each item in Request.Form
			objParams.appendChild(VarNode(item, Request.Form(item)))
		Next
		objRequestNode.appendChild(objParams)
		Set objParams = Nothing
		Set objCGI = xml.createElement("cgi-data")
		For each item in Request.ServerVariables
			objCGI.appendChild(VarNode(item, Request.ServerVariables(item)))
		Next
		objRequestNode.appendChild(objCGI)
		Set objCGI = Nothing
		Set objSession = xml.createElement("session")
		For each item in Session.Contents
			objSession.appendChild(VarNode(item, Session.Contents(item)))
		Next
		objRequestNode.appendChild(objSession)
		Set objSession = Nothing
		objRoot.appendChild(objRequestNode)

		'environment node
		Set objEnvNode = xml.createElement("server-environment")
		objEnvNode.appendChild(AddTextNode("environment-name", environment))
		objEnvNode.appendChild(AddTextNode("project-root", Request.ServerVariables("APPL_PHYSICAL_PATH")))
		objRoot.appendChild(objEnvNode)
		Set objEnvNode = Nothing

		xml.appendChild(objRoot)

		'Now send to airbrake
		Set objHTTP = CreateObject("WinHttp.WinHttprequest.5.1")

		objHTTP.SetTimeouts 2000,2000,2000,10000
		objHTTP.Open "POST", "http://api.airbrake.io/notifier_api/v2/notices", false
		objHTTP.SetRequestHeader "Content-Type","text/xml"
		objHTTP.Send xml.xml

		strResponse = objHTTP.ResponseText
		Set objHTTP = Nothing

		Set Construct = Me

		'For debug only
		'Response.Write(Server.HTMLEncode(xml.xml))
		'Response.Write("<hr>")
		'Response.Write(Server.HTMLEncode(strResponse))
		'Response.End

	End Function

	Private Function AddTextNode(strName, strValue)

		Set objNode = xml.createElement(strName)
		Set objText = xml.createTextNode(strValue)
		objNode.appendChild(objText)

		Set AddTextNode = objNode

		Set objText = Nothing
		Set objNode = Nothing

	End Function

	Private Function VarNode(strName, strValue)

		Set objNode = xml.createElement("var")
		objNode.setAttribute "key", strName
		If (strValue <> "") Then
			Set objText = xml.createTextNode(strValue)
			objNode.appendChild(objText)
			Set objText = Nothing
		End If
		Set VarNode = objNode
		Set objNode = Nothing

	End Function

	Public Function GetLastError()

		If (constructed) Then
			GetLastError = error
		Else
			Call Err.Raise(99999, "NotConstructedException", "AirBrakeClass not constructed")
		End If

	End Function

End Class
%>