<%

'GIOVANNI AREVALO X207063 , NEW REST URL IMPLEMENTED'
'Const strCPURL2 =https://apigw-st.tsl.telus.com/customer/businessProfile/v1/organization/'
Const strCPURL2 ="https://apigw-pr.tsl.telus.com/customer/businessProfile/v1/organization/"
'******************************************************************************
'	CheckLogon
'
'	Author: Daniel Nica
'	Date: Aug. 15, 2000
'
'	Returns user's access level to the specified Business Area
'	Param:  BAN - Business Area Name (check smaConstants.inc for details)
'  
'******************************************************************************
public function CheckLogon(strBAN)
dim x
	x = 0
	'if Session("UserAccessLevel").Item(strBAN) = 15 then
	'    Response.Write "Working"
	'else
	'    Response.Write("<script language=VBScript>MsgBox """ + Cstr(Session("UserAccessLevel").Item(strBAN) ) + """</script>") 
	'end if
	if Session("UserAccessLevel").Exists(strBAN)=true then
	    x = Session("UserAccessLevel").Item(strBAN)
 	end if
	CheckLogon = x



	'response.write "CheckLogon has strBAN = "  &strBAN
	'response.end
	'if strBAN <> strConst_Logon then
	
	'	x = CInt(Request.Cookies("UserAccessLevel")(strBAN))
	'	if x < 100 then
	'        x = CInt(Session("UserAccessLevel").Item(strBAN))
	 '       response.write "x=" &x
	 '       if x < 1 then
	'		'if Request.Cookies("UserAccessLevel")(strConst_Logon) = "" then
	'		if IsEmpty(Session("UserAccessLevel").Item(strBAN)) then
	'			response.redirect "default.asp?redir=Y"
	'		end if
	'	end if
	'	CheckLogon = x
	'else
		'if Request.Cookies("UserAccessLevel")(strBAN) = "" the
		'if IsEmpty(Session("UserAccessLevel").Item(strBAN)) then
		'	response.redirect "default.asp?redir=Y"
		'end if
	'end if
end function

' CAPICOM Constants                                                          
			Const CAPICOM_ENCRYPTION_ALGORITHM_RC2 = 0
			Const CAPICOM_ENCRYPTION_ALGORITHM_RC4 = 1
			Const CAPICOM_ENCRYPTION_ALGORITHM_DES = 2
			Const CAPICOM_ENCRYPTION_ALGORITHM_3DES = 3
			Const CAPICOM_ENCRYPTION_ALGORITHM_AES = 4
			Const CAPICOM_ENCRYPTION_KEY_LENGTH_MAXIMUM = 0
			Const CAPICOM_ENCRYPTION_KEY_LENGTH_40_BITS = 1
			Const CAPICOM_ENCRYPTION_KEY_LENGTH_56_BITS = 2
			Const CAPICOM_ENCRYPTION_KEY_LENGTH_128_BITS = 3
			Const CAPICOM_ENCRYPTION_KEY_LENGTH_192_BITS = 4
			Const CAPICOM_ENCRYPTION_KEY_LENGTH_256_BITS = 5 
			Const CAPICOM_SECRET_PASSWORD = 0
			Const CAPICOM_ENCODE_BASE64 = 0
			Const CAPICOM_ENCODE_BINARY = 1
			Const CAPICOM_ENCODE_ANY = -1

Public function DecryptWithKey(strPass, strData)
    dim oEncryptedData
	
			Set oEncryptedData = Server.CreateObject("CAPICOM.EncryptedData")				
			oEncryptedData.Algorithm.KeyLength = CAPICOM_ENCRYPTION_KEY_LENGTH_256_BITS
			oEncryptedData.Algorithm.Name = CAPICOM_ENCRYPTION_ALGORITHM_AES
			' Set the secret to be used when deriving the key
			oEncryptedData.SetSecret strPass, CAPICOM_SECRET_PASSWORD
			' Encrypt the data and return it in Base64 to the txtCipherText object
			DecryptWithKey = oEncryptedData.Decrypt(strData)
			DecryptWithKey = oEncryptedData.Content
	 	  Set oEncryptedData = Nothing
End function

Public function EncryptWithKey(strPass, strData)

if(strData <> "" or strData <> Empty) then
			Set oEncryptedData = Server.CreateObject("CAPICOM.EncryptedData")				
			oEncryptedData.Algorithm.KeyLength = CAPICOM_ENCRYPTION_KEY_LENGTH_256_BITS
			oEncryptedData.Algorithm.Name = CAPICOM_ENCRYPTION_ALGORITHM_AES
			' Set the secret to be used when deriving the key
			oEncryptedData.SetSecret strPass, CAPICOM_SECRET_PASSWORD
			' Provide the Encrypted Data object with the text to encrypt
			oEncryptedData.Content = strData				
			' Encrypt the data and return it in Base64 to the txtCipherText object
			EncryptWithKey = oEncryptedData.Encrypt(CAPICOM_ENCODE_BASE64)
	 	  Set oEncryptedData = Nothing
		  End if
End function


'******************************************************************************
'	routineHtmlString
'
'	Author: Oliver Kim
'	Date: Sept. 24, 1999
'
'	Function will format strings for HTML.
'	Param: inString is String to check
'  Errors: 0301 = unexpected error
'******************************************************************************
Public Function routineHtmlString (inString)
	Dim aString
	
	on error resume next
	
	if Len(inString) > 0 Then
		aString = Replace(inString, ">", "&#62;")
		aString = Replace(aString, "<", "&#60;")
		aString = Replace(aString, "'", "&#39;")
		aString = Replace(aString, """", "&#34;")
		routineHtmlString = aString
	Else
		routineHtmlString = ""
	End If
	' handle any missed errors
	if Err.number <> 0 then
		Err.Raise 0301, "routineHtmlString()" ,"Unexpected Error: " & Err.description
		Exit Function
	end if
End Function

'*******************************************************************************
'	routineJavascriptString
'
'	Author: Oliver Kim
'	Date: don't know
'	Modified: Sept. 24, 1999 - moved to routines.inc file (OK)
'
'	Function will format strings for Javascript.
'	Param: inString is String to check
'  Errors: 0301 = unexpected error
'*******************************************************************************
Public Function routineJavascriptString (inString)
	Dim aString
	
	on error resume next
	
	if Len(inString) > 0 Then
		aString = Replace(inString, Chr(13)+Chr(10), "\n")
		aString = Replace(aString, "'", "\'")
		aString = Replace(aString, """", "&#34;")
		aString = Replace(aString, "+", "\n")
		routineJavascriptString = aString
	Else
		routineJavascriptString = ""
	End If
	' handle any missed errors
	if Err.number <> 0 then
		Err.Raise 0301, "routineJavascriptString()" ,"Unexpected Error: " & Err.description
		Exit Function
	end if
End Function


'******************************************************************************
'	routineOraString
'	
'	Author: Henry Kamstra
'	Date: don't know
'	Modified: Sept. 24, 1999 - moved to routines.inc file (OK)
'
'	Function will format strings for oracle SQL.
'	Param: inString is String to check
'  Errors: 0301 = unexpected error
'******************************************************************************
Public Function routineOraString(inString)
	Dim aString
	
	on error resume next
	
	aString = Replace(inString, "'", "''")
	routineOraString = aString
	' handle any missed errors
	if Err.number <> 0 then
		Err.Raise 0301, "routineOraString()" ,"Unexpected Error: " & Err.description
		Exit Function
	end if
End Function


'******************************************************************************
'	rrRmvSpChr
'
'	Author: Dan Ty
'	Date: Feb 8, 2002
'
'	Function will enclose a given variable name with functions that will remove
'      non-alphabet and non-numberic characters or special characters in the
'      value of the given variable name or return the stripped value itself.
'	Param: inVarName is the variable name to enclose with stripping fuctions or
'              the value of the string.
'          inCode indicate whether the return string is a code to strip the variable
'              of special characters (Y) or a stripped value of variable itself (N).
'  Errors: 0301 = unexpected error
'******************************************************************************
Public Function rtRmvSpChr (inVarName, inCode)

	dim aString
	on error resume next

	if inCode = "Y" then
		aString = "TRIM(UPPER(REPLACE(TRANSLATE(" & inVarName & ", '~`!@#$%^&*()_-+={}[]:,<>,.?/|\" & """" & "', '                                   '), ' ', '')))"
	else
		aString = inVarName

		aString = REPLACE(aString, "~", "")
		aString = REPLACE(aString, "`", "")
		aString = REPLACE(aString, "!", "")
		aString = REPLACE(aString, "@", "")
		aString = REPLACE(aString, "#", "")
		aString = REPLACE(aString, "$", "")
		aString = REPLACE(aString, "%", "")
		aString = REPLACE(aString, "^", "")
		aString = REPLACE(aString, "&", "")
		aString = REPLACE(aString, "*", "")
		aString = REPLACE(aString, "(", "")
		aString = REPLACE(aString, ")", "")
		aString = REPLACE(aString, "_", "")
		aString = REPLACE(aString, "-", "")
		aString = REPLACE(aString, "+", "")
		aString = REPLACE(aString, "=", "")
		aString = REPLACE(aString, "{", "")
		aString = REPLACE(aString, "}", "")
		aString = REPLACE(aString, "[", "")
		aString = REPLACE(aString, "]", "")
		aString = REPLACE(aString, ":", "")
		aString = REPLACE(aString, ";", "")
		aString = REPLACE(aString, "<", "")
		aString = REPLACE(aString, ">", "")
		aString = REPLACE(aString, ",", "")
		aString = REPLACE(aString, ".", "")
		aString = REPLACE(aString, "?", "")
		aString = REPLACE(aString, "/", "")
		aString = REPLACE(aString, "\", "")
		aString = REPLACE(aString, "|", "")
		aString = REPLACE(aString, "'", "")
		aString = REPLACE(aString, " ", "")
		aString = REPLACE(aString, """", "")

	end if

	rtRmvSpChr = aString
	
	' handle any missed errors
	if Err.number <> 0 then
		Err.Raise 0301, "rtRmvSpChr" ,"Unexpected Error: " & Err.description
		Exit Function
	end if
end function

'*****************************************************************************
'	DateToYYYYMMDD
'
'	Author: Daniel Nica
'	Date: April 10, 2000
'
'	Converts a date value in a YYYYMMDD string
'
'	Input:			date
'	Return:			YYYYMMDD string representing entered date
'
'  Errors: 0301 = unexpected error
'******************************************************************************
function DateToYYYYMMDD(anyDate)
	dim strTmp
	
	on error resume next
	
	if isNull(anyDate) then
		DateToYYYYMMDD = ""
	else
		strTmp = CStr(Year(anyDate))
		if Month(anyDate) < 10 then
			strTmp = strTmp + "0"
		end if
		strTmp = strTmp + CStr(Month(anyDate))
		if Day(anyDate) < 10 then
			strTmp = strTmp + "0"
		end if
		strTmp = strTmp + CStr(Day(anyDate))
		DateToYYYYMMDD = strTmp
	end if
	' handle any missed errors
	if Err.number <> 0 then
		Err.Raise 0301, "DateToYYYYMMDD()" ,"Unexpected Error: " & Err.description
		Exit Function
	end if
end function

'--------------------------------------------------------
' Function DateTimeString(dtmDate)
' Input:  A dtmDate value from the database
' Returns: String - "MM/dd/YYYY 12H:MI:SS AM"
' Errors: 0301 = unexpected error
'--------------------------------------------------------
Function DateTimeString(dtmDate)
	on error resume next
			
	DateTimeString = Month(dtmDate) & "/" &_
					 Day(dtmDate) & "/" &_
					 Year(dtmDate) & " " &_
					 FormatDateTime(dtmDate, vbLongTime)
	if  (Err.number <> 0) then 
		Err.Raise 0301, "DateTimeString()" , "Unexpected Error formatting date: " & Err.description
		Exit function
	end if
End Function


'--------------------------------------------------------
' Function DateString(dtmDate)
' Input:  A dtmDate value from the database
' Returns: String - "MM/dd/YYYY"
' Errors: 0301 = unexpected error
'--------------------------------------------------------
Function DateString(dtmDate)
	on error resume next
			
	DateTimeString = Month(dtmDate) & "/" &_
					 Day(dtmDate) & "/" &_
					 Year(dtmDate) & " "
	if  (Err.number <> 0) then 
		Err.Raise 0301, "DateTimeString()" , "Unexpected Error formatting date: " & Err.description
		Exit function
	end if
End Function


'******************************************************************************
'	DisplayError
'
'	Author: Daniel Nica
'	Date: Aug. 22, 2000
'
'	Returns user's access level to the specified Business Area
'	Param:  BAN - Business Area Name (check smaConstants.inc for details)
'  
'******************************************************************************
function DisplayError(errAction, winLocation, errCode, errDesc, errExtra)
	Response.Clear
	if InStr(errExtra, "[Microsoft][ODBC driver for Oracle]") = 1 then
		errExtra = Right(errExtra, Len(errExtra)-35)
	end if
	if InStr(errExtra, "[Oracle]") = 1 then
		errExtra = Right(errExtra, Len(errExtra)-8)
	end if
	Response.Write "<script type=""text/javascript"">"&vbCrLf
	Response.Write "var x;"&vbCrLf
	Response.Write "var winLocation = '"&winLocation&"';"&vbCrLf
	Response.Write "var errAction = '"&errAction&"'"&vbCrLf
	Response.Write "x = 'SMA Error Manager'"&vbCrLf
	Response.Write "x += String.fromCharCode(10);"&vbCrLf
	Response.Write "x += "& "'" & errcode &" - '+unescape('" & escape(errDesc) & "');"&vbCrLf
	Response.Write "x += String.fromCharCode(10);"&vbCrLf
	Response.Write "x += String.fromCharCode(10);"&vbCrLf
	Response.Write "x += unescape('" & escape(errExtra) & "');"&vbCrLf
	Response.Write "x += String.fromCharCode(10);"&vbCrLf
	Response.Write "x += String.fromCharCode(10);"&vbCrLf
	Response.Write "if (errAction == 'REFRESH') {"
	Response.Write "x += String.fromCharCode(10);"&vbCrLf
	'Response.Write "x += 'Action: Click OK to refresh the page from the database or CANCEL to go back to the page without refreshing the data. You can always manually refresh a screen by clicking the RESET button at the bottom of the page.';"&vbCrLf
	'Response.Write "  if (confirm(x)) {"&vbCrLf
	Response.Write "    alert(x);" & vbCrLf
	Response.Write "    top.text.document.location=winLocation;"&vbCrLf
	'Response.Write "  }else{history.go(-1);}"&vbCrLf
	Response.Write "} else if (errAction == 'CLOSE') {" & vbCrLf
	Response.Write "   x += 'Action: Click OK button to close the window. You can try this function again after the window is closed.';"&vbCrLf
	Response.write "   alert(x);" & vbCrLf
	REsponse.write "   window.close();" & vbCrLf
	Response.Write "} else {"&vbCrLf
	Response.Write "alert(x);"&vbCrLf
	Response.Write "history.go(-1);"&vbCrLf
	Response.Write "}"&vbCrLf
	Response.Write "</script>"&vbCrLf
	Response.End
end function

'******************************************************************************
'	SendMail
'
'	Author: Daniel Nica
'	Date: Sept. 22, 2000
'
'	Sends email
'	Param:  see function call below
'  
'******************************************************************************
function SendEmail(strFrom, strTo, strCC, strBCC, strSubject, strBody)
	dim objEmail
	
	set objEmail = Server.CreateObject("CDONTS.NewMail")
	if err.number <> 0 then 
		err.Raise 0601,, "Cannot create new email object."
		exit function
	end if

	objEmail.MailFormat = 0 'CdoMailFormatMime
	if err then
		Err.Raise 0601,, "Error formatting new mail object."
		exit function
	end if
	
	objEmail.From = strFrom
	objEmail.To = strTo
	objEmail.Cc = strCC
	objEmail.Bcc = strBCC
	objEmail.Subject = strSubject
	objEmail.Body = strBody

	objEmail.Send
	if err then
		Err.Raise 0601,, "Error sending new email."
		exit function
	end if

	set objEmail = nothing
	if err then
		Err.Raise err.number, , err.description
		exit function
	end if
end function

'******************************************************************************
'   Below is a reference for xmlHttp Web service
'	
' 	xmlHttp.readyState  (AJAX reference: http://www.cristiandarie.ro/asp-ajax/Async.html)
'	Returns the status of the request:
'
'	0 = uninitialized
'	1 = loading
'	2 = loaded
'	3 = interactive
'	4 = complete
'******************************************************************************
'   function handlexmlHTTPRequest
''	Author: Anthony Cheung
'	Date: June 27, 2013
'
'	Returns error message
'	Param:  xmlHttp, soapReq, wsname
'******************************************************************************
Public function handlexmlHTTPRequest(xmlHttp, soapReq, wsname)
	<!--#include file="smaConstants.inc"-->
	<!--#include file="sma_env.inc"-->
	Response.write "<p>handlexmlHTTPRequest</p>"
  	if xmlHttp.readyState >= 1  then	  
	    ' try to connect to the server
	    'try
	      ' initiate server request
		   on error resume next
	       xmlHTTP.send soapReq
	      ' display an error in case of failure
	       if err then
		   		DisplayError "BACK", "", err.Number, " nc_getVPNListByService - web service is not available! Please check with your service admin.", err.Description
		   end if	
	    'catch (e)    
	else
		Response.write "<p>Cannot connect to server. readyState = " & xmlHttp.readyState & "</p>"
		'		exit
	    '  Response.write "Can't connect to server:" + e.toString()
	end if
end function

'******************************************************************************
'	CP_customers
'
'	Author: Anthony Cheung
'	Date: May 14, 2013
'
'	Returns web service status, record count and a list of the customer profile records
'	Param:  Customer Profile Name, max record return 
'  
'******************************************************************************
Public function CP_customers(customername, recordmax, record_count, cpList)
	<!--#include file="smaConstants.inc"-->

	'Response.write "<p>Inside CP_Customers</p>"	

	'on error resume next
	
	'SOAP variables
	dim wsUrl, xmlHTTP, xmlDoc, strWSStatus, strRetval, soapReq, retCustomers, elem, name, names, key, ExtKeys
	dim custChildNodes, node, nindex, strESDCID,tokenString,fso, jsonObjs, str,obj,accessToken,tokenType,experiesIn,obj2,customer,jsonObjs2

	
		'GIOVANNI AREVALO X207063 COMPLETE REST API URL USING NAME PARAMETER '
		wsUrl = strCPURL2&"/?name="& Server.HTMLEncode(customername)
		
	   'GIOVANNI AREVALO x207063'
		tokenString	 = generatingToken()
		
		Set jsonObjs = New VbsJson	
		Set obj = jsonObjs.decode(tokenString)
		accessToken = obj("access_token")
		tokenType = obj("token_type")
		experiesIn = obj("expires_in")

		Set xmlHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
		Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")

        xmlHTTP.SetOption SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS, SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS

        xmlHTTP.open "GET", wsUrl, false
		
		xmlHTTP.setRequestHeader "Content-Type", "text/xml;charset=UTF-8" 
	  	xmlHTTP.setRequestHeader "Authorization",tokenType &" "&accessToken
		
	  	if xmlHttp.readyState >= 1  then	  
		    ' try to connect to the server
		    'try
		      ' initiate server request
			   on error resume next
		       xmlHTTP.send 
		      ' display an error in case of failure
		       if err then
			   		DisplayError "BACK", "", err.Number, " searchCustomer - web service is not available! Please check with your service admin.", err.Description
			   end if	
		    'catch (e)    
		else
			Response.write "<p>Cannot connect to server. readyState = " & xmlHttp.readyState & "</p>"
		end if

	
		
		strWSStatus = xmlHTTP.Status
		'Response.write "<p>Status = " & strWSStatus & "</p>"
		strRetval = xmlHTTP.responseText
		'Response.write "Return = " & strRetval 

		Dim x,xref, eng : Set eng = CreateObject("MSScriptControl.ScriptControl")
		eng.Language = "JScript"
		eng.AddCode "function json() { return " & strRetval & "; }"
		Dim oResp : Set oResp = eng.Run("json")
	
		nindex=0
		
		record_count = oResp.length
		
		For Each x In oResp
		
			cpList(1,nindex) = x.name
			For Each xref In x.externalReference
				
				If   StrComp(xref.externalReferenceType,"CustomerId",1) = 0 then
				
					cpList(0,nindex) = xref.entitySourceKeyId
				end if
			Next
			nindex = nindex + 1
		Next
		

		CP_customers = strWSStatus
			
		Set xmlHTTP = nothing
		Set xmlDoc = nothing
	
end function

'******************************************************************************
'	CP_GetCustomerID
'
'	Author: Anthony Cheung
'	Date: May 15, 2013
'
'	Returns web service status, record count and a list of the customer id records
'	Param:  Customer Profile ID, max record return
'  
'******************************************************************************
Public function CP_GetCustomerID(CustomerProfileID, recordmax, record_count, cidList)
	<!--#include file="smaConstants.inc"-->
	'Response.write "<p>Inside CP_GetCustomerID</p>"	

	'on error resume next
	
	'SOAP variables
	dim wsUrl, xmlHTTP, xmlDoc, strWS3Status, strRetval, soapReq, retCustomerID, elem, name, names
	dim node, nindex,jsonObjs,obj,accessToken,tokenType,experiesIn,tokenString
  

	wsUrl = strCPURL2&"/?customerProfileId="& Server.HTMLEncode(CustomerProfileID)
		
	'GIOVANNI AREVALO x207063'
	tokenString = generatingToken()
		
	Set jsonObjs = New VbsJson	
	Set obj = jsonObjs.decode(tokenString)
	accessToken = obj("access_token")
	tokenType = obj("token_type")
	experiesIn = obj("expires_in")
		
	Set xmlHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
	Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")	

    xmlHTTP.SetOption SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS, SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS

    xmlHTTP.open "GET", wsUrl, false

	xmlHTTP.setRequestHeader "Content-Type", "text/xml;charset=UTF-8" 
	xmlHTTP.setRequestHeader "Authorization",tokenType &" "&accessToken

	
  	if xmlHttp.readyState >= 1  then	  
	    ' try to connect to the server
	    'try
	      ' initiate server request
		   on error resume next
	        xmlHTTP.send 
	      ' display an error in case of failure
	       if err then
		   		DisplayError "BACK", "", err.Number, " getNetworkList - web service is not available! Please check with your service admin.", err.Description
		   end if	
	    'catch (e)    
	else
		Response.write "<p>Cannot connect to server. readyState = " & xmlHttp.readyState & "</p>"
	end if
	
	strWS3Status = xmlHTTP.Status
	'Response.write "<p>Status = " & strWS3Status & "</p>"
	strRetval = xmlHTTP.responseText
	'Response.write "<p>Response = " & strRetval & "</p>"

	Dim x,xref, eng : Set eng = CreateObject("MSScriptControl.ScriptControl")
		eng.Language = "JScript"
		eng.AddCode "function json() { return " & strRetval & "; }"
		Dim oResp : Set oResp = eng.Run("json")

	record_count = oResp.length
	
	if record_count > recordmax then
		Response.write "<p>Maximum record exceeded. Please contact service admin.</p>"
		Response.end
	end if		
	
	'Response.write "<p>Found CIDs " & record_count & "</p>"

	nindex = 0

	
		For each node in oResp

			For Each xref In node.externalReference
				
				If   StrComp(xref.externalReferenceType,"CustomerId",1) = 0 then
				
					cidList(nindex) = xref.entitySourceKeyId
				end if
			Next
			nindex = nindex + 1
		Next
		
	CP_GetCustomerID = strWS3Status
		

end function

'******************************************************************************
'	CP_GetCPName
'
'	Author: Anthony Cheung
'	Date: May 15, 2013
'
'	Returns web service status and a customer profile name
'	Param:  Customer Profile ID, max record return
'  
'******************************************************************************
Public function CP_GetCPName(CustomerProfileID, record_count, cpname)
	<!--#include file="smaConstants.inc"-->
	'Response.write "<p>Inside CP_GetCPName</p>"	

	'on error resume next
	
	'SOAP variables
	dim wsUrl, xmlHTTP, xmlDoc, strWS3Status, strRetval, soapReq, retCustomerID, elem, name, names, custChildNodes 
	dim node,tokenString,jsonObjs,obj,accessToken,tokenType,experiesIn,nindex
  
    'GIOVANNI AREVALO X207063 COMPLETE REST API URL USING NAME PARAMETER '
    wsUrl = strCPURL2&"/?customerProfileId="& Server.HTMLEncode(CustomerProfileID)
	'Response.write CustomerProfileID
    'GIOVANNI AREVALO x207063 function that retrieve token to be used in API call
    tokenString	 = generatingToken()
    'taking json values
    Set jsonObjs = New VbsJson	
    Set obj = jsonObjs.decode(tokenString)
    accessToken = obj("access_token")
    tokenType = obj("token_type")
    experiesIn = obj("expires_in")

		
	Set xmlHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
	Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")

    xmlHTTP.SetOption SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS, SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
    'GIOVANNI AREVALO X207063 USING GET PROTOCOL
    xmlHTTP.open "GET", wsUrl, false
	
	xmlHTTP.setRequestHeader "Content-Type", "text/xml;charset=UTF-8" 
	xmlHTTP.setRequestHeader "Authorization",tokenType &" "&accessToken
	
  	if xmlHttp.readyState >= 1  then	  
	    ' try to connect to the server
	    'try
	      ' initiate server request
		   on error resume next
	       xmlHTTP.send 
	      ' display an error in case of failure
	       if err then
		   		DisplayError "BACK", "", err.Number, " getCustomer - web service is not available! Please check with your service admin.", err.Description
		   end if	
	    'catch (e)    
	else
		Response.write "<p>Cannot connect to server. readyState = " & xmlHttp.readyState & "</p>"
	end if
	

    strWS3Status = xmlHTTP.Status
    'Response.write "<p>Status = " & strWS3Status & "</p>"
    strRetval = xmlHTTP.responseText
    'Response.write "Return = " & strRetval

    'GIOVANNI AREVALO X207063 READING JSON RESPONSE
    Dim x,xref, eng : Set eng = CreateObject("MSScriptControl.ScriptControl")
    eng.Language = "JScript"
    eng.AddCode "function json() { return " & strRetval & "; }"
    Dim oResp : Set oResp = eng.Run("json")
    
    record_count = oResp.length


		
		For Each x In oResp
		
			cpname = x.name

		Next

		CP_GetCPName = strWS3Status
			
		Set xmlHTTP = nothing
		Set xmlDoc = nothing
end function

'******************************************************************************
'	nc_getCSIDListByVPN
'
'	Author: Anthony Cheung
'	Date: June 6, 2013
'
'	Returns web service status, record count and a list of the CSID records
'	Param:  VPN Name, max record return
'  
'******************************************************************************
Public function nc_getCSIDListByVPN(VPNname, recordmax, record_count, vpnlist)
	<!--#include file="smaConstants.inc"-->
	<!--#include file="sma_env.inc"-->
	'Response.write "<p>Inside nc_getCSIDListByVPN</p>"	

	'on error resume next
	
	'SOAP variables
	dim wsUrl, xmlHTTP, xmlDoc, strWSStatus, strRetval, soapReq, retCSID, elem, name, names
	dim node, nindex, CSIDChildNodes, CSIDs, VPN, VPNs
  
	wsUrl = strNCWSURL
		
	Set xmlHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
	Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")

    xmlHTTP.SetOption SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS, SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS

    xmlHTTP.open "POST", wsUrl, false, strcpwsid, strcppw

'	xmlHTTP.open "POST", wsUrl, false
	
	xmlHTTP.setRequestHeader "Content-Type", "text/xml;charset=UTF-8" 
	xmlHTTP.setRequestHeader "SOAPAction", "getServiceListByVPN"
	
	soapReq = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:vpn=""http://xmlschema.tmi.telus.com/srv/RMO/InventoryMgmt/VPNInventoryServiceRequestResponse_v1"">" &_
    "<soapenv:Header/>" &_
     	"<soapenv:Body>" &_
      		"<vpn:getServiceListByVPN>" &_
         		"<vpnName>" & VPNname & "</vpnName>" &_
      		"</vpn:getServiceListByVPN>" &_
   		"</soapenv:Body>" &_
    "</soapenv:Envelope>"
	
  	if xmlHttp.readyState >= 1  then	  
	    ' try to connect to the server
	    'try
	      ' initiate server request
		   on error resume next
	       xmlHTTP.send soapReq
	      ' display an error in case of failure
	       if err then
		   		DisplayError "BACK", "", err.Number, " nc_getCSIDListByVPN - web service is not available! Please check with your service admin.", err.Description
		   end if	
	    'catch (e)    
	else
		Response.write "<p>Cannot connect to server. readyState = " & xmlHttp.readyState & "</p>"
	end if

	'xmlHTTP.send soapReq
	
	strWSStatus = xmlHTTP.Status
	'Response.write "<p>Status = " & strWSStatus & "</p>"
	strRetval = xmlHTTP.responseText
	xmlDoc.loadxml(xmlHTTP.responseText)
	
	Set retCSID = xmlDoc.getElementsByTagName("serviceList")

	record_count = retCSID.length
	
	'Response.write "<p>Found VPNname " & VPNname & "</p>"
	'Response.write "<p>Found CSIDs " & record_count & "</p>"

	nindex = 0
	
	For each elem in retCSID
		set CSIDChildNodes = elem.childNodes
	    'set CSIDs = elem.getElementsByTagName("serviceId")

		For each node in CSIDChildNodes
	        if node.nodeName = "serviceId" then
				'Response.Write node.nodeName & "  =  " & node.text & "<br />" & vbCrLf
				vpnlist(nindex) = node.text
			end if
		Next
		nindex = nindex + 1

		'if node.nodeName = "serviceId" then
		'	'Response.Write node.nodeName & "  =  " & node.text & "<br />" & vbCrLf
		'	vpnindex(nindex) = node.text
		'	'Response.Write "<hr>" & vbCrLf
		'end if
		'nindex = nindex + 1
	Next
		
	nc_getCSIDListByVPN = strWSStatus
	
end function

'******************************************************************************
'	nc_getVPNListByService
'
'	Author: Anthony Cheung
'	Date: June 11, 2013
'
'	Returns web service status, record count, a list of the VPN names, its vpntype, vpntopology, vrfs, RD, Connect_To VPN name and customer name
'	Param:  CSID
'  
'******************************************************************************
Public function nc_getVPNListByService(CSID, record_count, vpnlist, vpntypelist, vpntoptypelist, keyCSID, vrflist, RDlist, cTvpnlist, cTcustomerName)
	<!--#include file="smaConstants.inc"-->
	<!--#include file="sma_env.inc"-->
' Debug info
'	Response.write "<p>Inside nc_getVPNListByService</p>"	

	'SOAP variables
	dim wsUrl, xmlHTTP, xmlDoc, strWSStatus, strRetval, soapReq, retCSID, elem, name, names, retVpnName
	dim node, nindex, VPNChildNodes, VPNs, VPN, test, wsname, vrfs, vrf, connectedTos, connectedTo 
  
	wsUrl = strNCWSURL
	
	'Response.write	"<p>wsUrl = " & wsUrl & "</p>"

	'on error resume next
		
	Set xmlHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
	Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")

    xmlHTTP.SetOption SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS, SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS

    xmlHTTP.open "POST", wsUrl, false, strcpwsid, strcppw
'	xmlHTTP.open "POST", wsUrl, false
	
	xmlHTTP.setRequestHeader "Content-Type", "text/xml;charset=UTF-8" 
	xmlHTTP.setRequestHeader "SOAPAction", "getServiceListByVPN"
	
	soapReq = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:vpn=""http://xmlschema.tmi.telus.com/srv/RMO/InventoryMgmt/VPNInventoryServiceRequestResponse_v1"">" &_
    "<soapenv:Header/>" &_
     	"<soapenv:Body>" &_
      		"<vpn:getVPNListByService>" &_
      			"<serviceList>" &_      		
	         		"<serviceId>" & CSID & "</serviceId>" &_
      			"</serviceList>" &_	         		
      		"</vpn:getVPNListByService>" &_
   		"</soapenv:Body>" &_
    "</soapenv:Envelope>"
	
	'Response.write	"<p>xmlHttp.readyState = " & xmlHttp.readyState & "</p>"
	
  	if xmlHttp.readyState >= 0  then	  
	    ' try to connect to the server
	    'try
	      ' initiate server request
		   on error resume next
	       xmlHTTP.send soapReq
	      ' display an error in case of failure
	       if err then
		   		DisplayError "BACK", "", err.Number, " nc_getVPNListByService - web service is not available! Please check with your service admin.", err.Description
		   end if	
	    'catch (e)    
	else
		Response.write "<p>Cannot connect to server. readyState = " & xmlHttp.readyState & "</p>"
	end if
	
'	xmlHTTP.send soapReq
					
	strWSStatus = xmlHTTP.Status
	'Response.write "<p>Status = " & strWSStatus & "</p>"
	'Response.write	"<p>xmlHttp.readyState after send = " & xmlHttp.readyState & "</p>"	

	strRetval = xmlHTTP.responseText
	xmlDoc.loadxml(xmlHTTP.responseText)
	
	record_count = 0     ' added by Linda Chen
		
	Set retVpnName = xmlDoc.getElementsByTagName("serviceList")
	
	'Linda Chen record_count = retVpnName.length
		
	' Debug info
	'	Response.write "<p>Found VPNs " & record_count & "</p>"
	
	nindex = 0
		
	For each elem in retVpnName 
		set VPNChildNodes = elem.childNodes
	    'set CSIDs = elem.getElementsByTagName("serviceId")
	
		For each node in VPNChildNodes
	        if node.nodeName = "serviceId" then
	' Debug info
	'			'lc Response.Write node.nodeName & "  =  " & node.text & "<br />" & vbCrLf
				keyCSID(nindex) = node.text
			end if
		Next
	
    	set VPNs = elem.getElementsByTagName("vpn")
    	
    	record_count = record_count + VPNs.length  ' Added by Linda Chen
    	
if VPNs.length > 0 then 
		'lc response.write "VPNs length is " & VPNs.length    	
		For each VPN in VPNs
	' Debug info
	'		Response.Write "Found VPNName : " & VPN.getElementsByTagName("vpnName").[0].text & "<br />" & vbCrLf
	'		Response.Write "Found VPNType : " & VPN.getElementsByTagName("vpnType").[0].text & "<br />" & vbCrLf
	'		Response.Write "Found VPNTopologyType : " & VPN.getElementsByTagName("vpnTopologyType").[0].text & "<br />" & vbCrLf
			vpnlist(nindex) = VPN.getElementsByTagName("vpnName").[0].text
			vpntypelist(nindex) = VPN.getElementsByTagName("vpnType").[0].text	
			vpntoptypelist(nindex) = VPN.getElementsByTagName("vpnTopologyType").[0].text
'VRF
	    	set vrfs = elem.getElementsByTagName("vrfList")
			For each vrf in vrfs
				vrflist(nindex) = vrf.getElementsByTagName("vrfName").[0].text
				RDlist(nindex) = vrf.getElementsByTagName("routeDistinguisherId").[0].text				
				'Response.Write "Found vrfName : " & VPN.getElementsByTagName("vrfName").[0].text & "<br />" & vbCrLf
				'Response.Write "Found routeDistinguisherId : " & VPN.getElementsByTagName("routeDistinguisherId").[0].text & "<br />" & vbCrLf
			next
				
	    	set connectedTos = elem.getElementsByTagName("connectedTo")
			For each connectedTo in connectedTos				
				cTvpnlist(nindex) = connectedTo.getElementsByTagName("vpnName").[0].text
				'cTvpntypelist(nindex) = connectedTo.getElementsByTagName("vpnType").[0].text
				'cTvpntoptypelist(nindex) = connectedTo.getElementsByTagName("vpnTopologyType").[0].text
				cTcustomerName(nindex) = connectedTo.getElementsByTagName("customerName").[0].text			
				'Response.Write "Found connectedTo vpnName : " & connectedTo.getElementsByTagName("vpnName").[0].text & "<br />" & vbCrLf
				'Response.Write "Found connectedTo vpnType : " & connectedTo.getElementsByTagName("vpnType").[0].text & "<br />" & vbCrLf
				'Response.Write "Found connectedTo vpnTopologyType : " & connectedTo.getElementsByTagName("vpnTopologyType").[0].text & "<br />" & vbCrLf
				'Response.Write "Found connectedTo customerName : " & connectedTo.getElementsByTagName("customerName").[0].text & "<br />" & vbCrLf				
			next													
		Next
end if
	
		nindex = nindex + 1
	
	Next
			
	nc_getVPNListByService = strWSStatus
		
end function

'******************************************************************************
'	nc_getVPNListByCustomer
'
'	Author: Anthony Cheung
'	Date: June 11, 2013
'
'	Returns web service status, record count, a list of the VPN names, its vpntype and its vpntopology  
'	Param:  CID
'  
'******************************************************************************
Public function nc_getVPNListByCustomer(CID, record_count, vpnlist, vpntypelist, vpntoptypelist, vrflist, RDlist, cTvpnlist, cTcustomerName)
<!--#include file="smaConstants.inc"-->
<!--#include file="sma_env.inc"-->
' Debug info
	'Response.write "<p>Inside nc_getVPNListByCustomer</p>"	

	'on error resume next
	
	'SOAP variables
	dim wsUrl, xmlHTTP, xmlDoc, strWSStatus, strRetval, soapReq, retCID, elem, name, names, retVpnName
	dim node, nindex, VPNChildNodes, VPNs, VPN, vrfs, vrf, connectedTos, connectedTo, strNCWSURL 
  
	strNCWSURL = "https://soa-mp-rmsk-it03.tsl.telus.com/RMO/InventoryMgmt/Inventory/VPNInventoryService_v1_0_1_vs0?wsdl"

	'Response.write "<p>Before request</p>"

	'Response.write strNCWSURL

	wsUrl = strNCWSURL

	'Response.write "<p>Before request</p>"	

	Set xmlHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
	Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")
    xmlHTTP.SetOption SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS, SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS

    xmlHTTP.open "POST", wsUrl, false, strcpwsid, strcppw
'	xmlHTTP.open "POST", wsUrl, false
	
	xmlHTTP.setRequestHeader "Content-Type", "text/xml;charset=UTF-8" 
	xmlHTTP.setRequestHeader "SOAPAction", "getVPNListByCustomer"
	
	'Response.write "<p>Before request</p>"

	soapReq = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:vpn=""http://xmlschema.tmi.telus.com/srv/RMO/InventoryMgmt/VPNInventoryServiceRequestResponse_v1"">" &_
    "<soapenv:Header/>" &_
     	"<soapenv:Body>" &_
      		"<vpn:getVPNListByCustomer>" &_
         		"<customerId>" & CID & "</customerId>" &_
      		"</vpn:getVPNListByCustomer>" &_
   		"</soapenv:Body>" &_
    "</soapenv:Envelope>"
	
  	if xmlHttp.readyState >= 1  then	  
	    ' try to connect to the server
	    'try
	      ' initiate server request
		   on error resume next
	       xmlHTTP.send soapReq
	      ' display an error in case of failure
	       if err then
		   		DisplayError "BACK", "", err.Number, " nc_getVPNListByCustomer - web service is not available! Please check with your service admin.", err.Description
		   end if	
	    'catch (e)    
	else
		Response.write "<p>Cannot connect to server. readyState = " & xmlHttp.readyState & "</p>"
	end if

'	xmlHTTP.send soapReq
	
	strWSStatus = xmlHTTP.Status
	'Response.write "<p>Status = " & strWSStatus & "</p>"
	strRetval = xmlHTTP.responseText
	xmlDoc.loadxml(xmlHTTP.responseText)
	
	'Set retVpnName = xmlDoc.getElementsByTagName("serviceList")

	Set retVpnName = xmlDoc.getElementsByTagName("vpnList")

	record_count = retVpnName.length

' Debug info	
	'Response.write "<p>Found VPNs " & record_count & "</p>"

	nindex = 0
	
	For each elem in retVpnName 
		set VPNChildNodes = elem.childNodes
	    'set CIDs = elem.getElementsByTagName("serviceId")

'		For each node in VPNChildNodes
'	        if node.nodeName = "serviceId" then
'				Response.Write node.nodeName & "  =  " & node.text & "<br />" & vbCrLf
'				keyCID(nindex) = node.text
'			end if
'		Next

    	set VPNs = elem.getElementsByTagName("vpnList")
	    	

' Debug info
		'Response.Write "Found VPNName : " & elem.getElementsByTagName("vpnName").[0].text & "<br />" & vbCrLf
		'Response.Write "Found VPNType : " & elem.getElementsByTagName("vpnType").[0].text & "<br />" & vbCrLf
		'Response.Write "Found VPNTopologyType : " & elem.getElementsByTagName("vpnTopologyType").[0].text & "<br />" & vbCrLf
		vpnlist(nindex) = elem.getElementsByTagName("vpnName").[0].text
		vpntypelist(nindex) = elem.getElementsByTagName("vpnType").[0].text	
		vpntoptypelist(nindex) = elem.getElementsByTagName("vpnTopologyType").[0].text


    	set vrfs = elem.getElementsByTagName("vrfList")
		For each vrf in vrfs
			vrflist(nindex) =  vrf.getElementsByTagName("vrfName").[0].text
			RDlist(nindex) = vrf.getElementsByTagName("routeDistinguisherId").[0].text
			'Response.Write "Found vrfName : " & vrf.getElementsByTagName("vrfName").[0].text & "<br />" & vbCrLf
			'Response.Write "Found routeDistinguisherId : " & vrf.getElementsByTagName("routeDistinguisherId").[0].text & "<br />" & vbCrLf
		next
		
		set connectedTos = elem.getElementsByTagName("connectedTo")
		For each connectedTo in connectedTos				
			cTvpnlist(nindex) = connectedTo.getElementsByTagName("vpnName").[0].text
			'cTvpntypelist(nindex) = connectedTo.getElementsByTagName("vpnType").[0].text
			'cTvpntoptypelist(nindex) = connectedTo.getElementsByTagName("vpnTopologyType").[0].text
			cTcustomerName(nindex) = connectedTo.getElementsByTagName("customerName").[0].text			
			'Response.Write "Found connectedTo vpnName : " & connectedTo.getElementsByTagName("vpnName").[0].text & "<br />" & vbCrLf
			'Response.Write "Found connectedTo vpnType : " & connectedTo.getElementsByTagName("vpnType").[0].text & "<br />" & vbCrLf
			'Response.Write "Found connectedTo vpnTopologyType : " & connectedTo.getElementsByTagName("vpnTopologyType").[0].text & "<br />" & vbCrLf
			'Response.Write "Found connectedTo customerName : " & connectedTo.getElementsByTagName("customerName").[0].text & "<br />" & vbCrLf				
		next

	nindex = nindex + 1
	Next

	
	nc_getVPNListByCustomer = strWSStatus
	
end function


'******************************************************************************
'	getUserRoleList
'
'	Author: Linda Chen
'	Date: January 2015
'
'	Returns web service status, record count, a list of the role names  
'	Param:  appId, userId
'  
'******************************************************************************
Public function getUserRoleList(appId, userId, retRoleNames)
<!--#include file="smaConstants.inc"-->
' Debug info
 	Response.write "<p>getUserRoleList</p>"	

	'on error resume next
	
	'SOAP variables
	dim wsUrl, xmlHTTP, xmlDoc, strWSStatus, strRetval, soapReq
	dim nindex, node

	wsUrl = strPermItWSURL    'strPermItWSURL needs be defined in smaConstants.inc
		
	Set xmlHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
	Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")
       ' xmlHTTP.SetOption SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS, SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS

        xmlHTTP.open "POST", strPermItWSURL, false, strPermItwsid, strPermItpw
 	xmlHTTP.open "POST", wsUrl, false
	
	xmlHTTP.setRequestHeader "Content-Type", "text/xml;charset=UTF-8" 
	xmlHTTP.setRequestHeader "SOAPAction", "getUserRoleList"
	'soapReq = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:vpn=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:aut=""http://xmlschema.tmi.telus.com/srv/EO/SecurityMgmt/AuthorizationPolicyServiceRequestResponse_v1"">" &_

	soapReq = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:vpn=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:aut=""http://xmlschema.tmi.telus.com/srv/EO/SecurityMgmt/AuthorizationPolicyServiceRequestResponse_v1"">" 
        soapReq = soapReq + "<soapenv:Header/>" 

soapReq = soapReq +"<soapenv:Body>" 

soapReq = soapReq +"<aut:getUserRoleList>"  
soapReq = soapReq + "<appId>"       
soapReq = soapReq + appId + "</appId>"  
soapReq = soapReq + "<userId>" + userId + "</userId>"  
soapReq = soapReq +  "</aut:getUserRoleList>"  
soapReq = soapReq +"</soapenv:Body>"  
soapReq = soapReq + "</soapenv:Envelope>"
'response.write soapReq

if xmlHttp.readyState >= 1  then	  
	    ' try to connect to the server
	    'try
	      ' initiate server request
	       on error resume next
	      	   
	       xmlHTTP.send soapReq
	      '  response.write "Called WS"
	      ' display an error in case of failure
	       if err then
		   		DisplayError "BACK", "", err.Number, " getUserRoleList - web service is not available! Please check with your service admin." & err.ASPDescription, err.Description
	       end if	
	    'catch (e)    
	else
		Response.write "<p>Cannot connect to server. readyState = " & xmlHttp.readyState & "</p>"
	end if

'	xmlHTTP.send soapReq
	
	strWSStatus = xmlHTTP.Status
	'Response.write "<p>HTTP Status = " & strWSStatus & "</p>"
strRetval = xmlHTTP.responseText
xmlDoc.loadxml(xmlHTTP.responseText)

'response.write xmlDoc.getElementsByTagName("roleList").length
'response.write xmlDoc.getElementsByTagName("roleList").[1].text +"<br>"

Set strRetval = xmlDoc.getElementsByTagName("roleList")
'response.write strRetval.length
'response.write strRetval.[1].text +"<br>"
'response.end

redim Preserve retRoleNames(strRetval.length)
nindex=0 
  
For each node in strRetval
	 
 	retRoleNames(nindex)= node.text
   
  	nindex = nindex + 1
Next
  	 
  
  	

	getUserRoleList = strWSStatus
	
end function







'******************************************************************************
'	getPolicyList
'
'	Author: Linda Chen
'	Date: January 2015
'
'	Returns web service status, record count, a list of the policy  
'	Param:  appI
'  
'******************************************************************************
Public function getPolicyList(appId, Policies)

 'response.write "<p>inside getPolicyList</p>"
 'response.end

dim nindex, i,rcount,accessLvl
dim policyList,policy 
 
 
dim wsUrl, xmlHTTP, xmlDoc, strWSStatus, strRetval, soapReq
 
 	wsUrl = strPermItWSURL    'strPermItWSURL needs be defined in smaConstants.inc
 		
 	Set xmlHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
 	Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")
        xmlHTTP.SetOption SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS, SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
 
  	xmlHTTP.open "POST", strPermItWSURL, false, strPermItwsid, strPermItpw
 	
 	xmlHTTP.setRequestHeader "Content-Type", "text/xml;charset=UTF-8" 
 	xmlHTTP.setRequestHeader "SOAPAction", "getUserRoleList"
 	'soapReq = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:vpn=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:aut=""http://xmlschema.tmi.telus.com/srv/EO/SecurityMgmt/AuthorizationPolicyServiceRequestResponse_v1"">" &_
 
 	soapReq = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:vpn=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:aut=""http://xmlschema.tmi.telus.com/srv/EO/SecurityMgmt/AuthorizationPolicyServiceRequestResponse_v1"">" 
         soapReq = soapReq + "<soapenv:Header/>" 
 
 soapReq = soapReq +"<soapenv:Body>" 
 
 soapReq = soapReq + "<aut:getAuthorizationPolicyList>"
 soapReq = soapReq + "<appId>"
 soapReq = soapReq + appId + "</appId>"  
 soapReq = soapReq + "</aut:getAuthorizationPolicyList>"
 soapReq = soapReq +"</soapenv:Body>"  
 soapReq = soapReq + "</soapenv:Envelope>"
 
 if xmlHttp.readyState >= 1  then	  
 	    ' try to connect to the server
 	    'try
 	      ' initiate server request
 	       on error resume next
 	      	   
 	       xmlHTTP.send soapReq
 	       ' response.write "Called WS"
 	      ' display an error in case of failure
 	       if err then
 		   		DisplayError "BACK", "", err.Number, " getAuthorizationPolicyList - web service is not available! Please check with your service admin.", err.Description
 	       end if	
 	    'catch (e)    
 	else
 		Response.write "<p>Cannot connect to server. readyState = " & xmlHttp.readyState & "</p>"
 	end if
 
 '	xmlHTTP.send soapReq
 	
 	strWSStatus = xmlHTTP.Status
 	'Response.write "<p>HTTP Status = " & strWSStatus & "</p>"
 	
 strRetval = xmlHTTP.responseText
 'response.write strRetval
 'response.end
 xmlDoc.loadxml(xmlHTTP.responseText)
  
  
 Set policyList = xmlDoc.getElementsByTagName("policyList")
 'response.write "policyList length is " 
 'response.write policyList.length
 if policyList.length=0 then
    getPolicyList = 0
 else 
	 nindex = 0 
	 for each policy in policyList
		'response.write " <BR> item " & nindex &":"

		'response.write policy.getElementsByTagName("resourceId")(0).text
		'response.write policy.getElementsByTagName("resourceId").[0].text
		'response.write policy.getElementsByTagName("expression").[0].text
		'response.write "<BR>"
		'response.end
		ReDim Preserve Policies(3,nindex + 1) 
		Policies(0,nindex) = policy.getElementsByTagName("expression").[0].text

		accessLvl=0
		rcount =policy.getElementsByTagName("actionList").length
		rcount=rcount-1
		for i = 0 to rcount
			  if strComp(UCase(policy.getElementsByTagName("actionList")(i).text),"SELECT")=0 then	     	  	 
				accessLvl = accessLvl+1
			  elseif strComp(UCase(policy.getElementsByTagName("actionList")(i).text),"INSERT")=0 then	 	           
				    accessLvl = accessLvl+2	   	 	

			  elseif strComp(UCase(policy.getElementsByTagName("actionList")(i).text), "UPDATE")=0 then	 	          
					accessLvl = accessLvl+4
			  elseif strComp(UCase(policy.getElementsByTagName("actionList")(i).text),"DELETE")=0 then	 	        
				 accessLvl=accessLvl+8
			  else
			  end if	 	    
		next 
		Policies(1,nindex) = accessLvl




		rcount =policy.getElementsByTagName("roleList").length

		for i = 0 to rcount-1
			'Policies(nindex,i) = policy.getElementsByTagName("roleList")(i).text
			'response.write "PolicyRoles("&nindex &"," &i &")=" &PolicyRoles(nindex,i)
			if i= 0 then
			   Policies(2,nindex) = policy.getElementsByTagName("roleList")(i).text
			else
			   nindex = nindex + 1
			   ReDim Preserve Policies(3,nindex + 1)
			   Policies(0,nindex) = policy.getElementsByTagName("expression").[0].text
			   Policies(1,nindex) = accessLvl
			   Policies(2,nindex) = policy.getElementsByTagName("roleList")(i).text
			end if
		next 	

		nindex = nindex+1
	 next



	'Response.Write(" <TABLE border=0>") 
	'Response.Write("<TR><TD>Row</TD> <TD>PolicyID</TD>") 
	'Response.Write("<TD>AccessLevel</TD><TD>RoleName</TD></TR>") 

	'The UBound function will return the 'index' of the highest element in an array.
	'For i = 0 to UBound(Policies, 2) -1
	'Response.Write("<TR><TD>#" & i & "</TD>") 
	'Response.Write("<TD>" & Policies(0,i) & "</TD>") 
	'Response.Write("<TD>" & Policies(1,i) & "</TD>") 
	'Response.Write("<TD>" & Policies(2,i) & "</TD>") 
	'Next 

	'Response.Write("</TABLE>") 
	'response.end

	  getPolicyList = strWSStatus
  end if
 end function
 
 
 
 
 
 
 
 
 
 Public function getApplicationRoleList(appId,  appRoleNbr)
 <!--#include file="smaConstants.inc"-->
 
 ' Debug info
  	Response.write "<p>getApplicationRoleList</p>"	
 
 	'on error resume next
 	
 	'SOAP variables
 	dim wsUrl, xmlHTTP, xmlDoc, strWSStatus, strRetval, soapReq
 	dim nindex, node
 
 	wsUrl = strPermItWSURL    'strPermItWSURL needs be defined in smaConstants.inc
 		
 	Set xmlHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
 	Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")
         xmlHTTP.SetOption SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS, SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
  	xmlHTTP.open "POST", strPermItWSURL, false, strPermItwsid, strPermItpw
 	
 	xmlHTTP.setRequestHeader "Content-Type", "text/xml;charset=UTF-8" 
 	xmlHTTP.setRequestHeader "SOAPAction", "getUserRoleList"
 	'soapReq = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:vpn=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:aut=""http://xmlschema.tmi.telus.com/srv/EO/SecurityMgmt/AuthorizationPolicyServiceRequestResponse_v1"">" &_
 
 	soapReq = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:vpn=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:aut=""http://xmlschema.tmi.telus.com/srv/EO/SecurityMgmt/AuthorizationPolicyServiceRequestResponse_v1"">" 
         soapReq = soapReq + "<soapenv:Header/>" 
 
 soapReq = soapReq +"<soapenv:Body>" 
 
 soapReq = soapReq +"<aut:getApplicationRoleList>"  
 soapReq = soapReq + "<appId>"       
 soapReq = soapReq + appId + "</appId>"  

 soapReq = soapReq +  "</aut:getApplicationRoleList>"  
 soapReq = soapReq +"</soapenv:Body>"  
 soapReq = soapReq + "</soapenv:Envelope>"
 response.write soapReq
 
 if xmlHttp.readyState >= 1  then	  
 	    ' try to connect to the server
 	    'try
 	      ' initiate server request
 	       on error resume next
 	      	   
 	       xmlHTTP.send soapReq
 	      '  response.write "Called WS"
 	      ' display an error in case of failure
 	       if err then
 		   		DisplayError "BACK", "", err.Number, " getApplilcationRoleList - web service is not available! Please check with your service admin.", err.Description
 	       end if	
 	    'catch (e)    
 	else
 		Response.write "<p>Cannot connect to server. readyState = " & xmlHttp.readyState & "</p>"
 	end if
 
 '	xmlHTTP.send soapReq
 	
 	strWSStatus = xmlHTTP.Status
 	'Response.write "<p>HTTP Status = " & strWSStatus & "</p>"
 strRetval = xmlHTTP.responseText
 xmlDoc.loadxml(xmlHTTP.responseText)
 
 'response.write xmlDoc.getElementsByTagName("roleList").length
 'response.write xmlDoc.getElementsByTagName("roleList").[1].text +"<br>"
 
 Set strRetval = xmlDoc.getElementsByTagName("roleList")
  
' response.write strRetval.[1].text +"<br>"
 'response.end
 
 nindex=0 
 appRoleNbr = 0  
 For each node in strRetval
      if isEmpty(node.text) then
        break
      else  	 
   	nindex = nindex + 1
   	appRoleNbr = appRoleNbr + 1
      end if
 Next
   	 
 
getApplicationRoleList = strWSStatus
 	
end function













'******************************************************************************
'	getUserRoleList
'
'	Author: Linda Chen
'	Date: January 2015
'
'	Returns web service status, record count, a list of the role names  
'	Param:  appId, userId
'  
'******************************************************************************
Public function getUserRolePolicyList(appId, userId, retRoleNames,retPolicyNames)
<!--#include file="smaConstants.inc"-->
' Debug info
 	'response.write "<p> In getUserRolePolicyList</p>"	
 	 

	'on error resume next
	
	'SOAP variables
	dim wsUrl, xmlHTTP, xmlDoc, strWSStatus, strRetval, soapReq
	dim rnindex,pnindex, node

	wsUrl = strPermItWSURL    'strPermItWSURL needs be defined in smaConstants.inc
		
	Set xmlHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
	Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")
        xmlHTTP.SetOption SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS, SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS

        xmlHTTP.open "POST", wsUrl, false, strPermItwsid, strPermItpw     'strPermItwsid, strPermItpw needs be defined in smaconstants.inc
 	'xmlHTTP.open "POST", wsUrl, false
	
	xmlHTTP.setRequestHeader "Content-Type", "text/xml;charset=UTF-8" 
	xmlHTTP.setRequestHeader "SOAPAction", "getUserRoleList"
	'soapReq = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:vpn=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:aut=""http://xmlschema.tmi.telus.com/srv/EO/SecurityMgmt/AuthorizationPolicyServiceRequestResponse_v1"">" &_

	soapReq = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:vpn=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:aut=""http://xmlschema.tmi.telus.com/srv/EO/SecurityMgmt/AuthorizationPolicyServiceRequestResponse_v1"">" 
        soapReq = soapReq + "<soapenv:Header/>" 

soapReq = soapReq +"<soapenv:Body>" 

soapReq = soapReq +"<aut:getUserRoleList>"  
soapReq = soapReq + "<appId>"       
soapReq = soapReq + appId + "</appId>"  
soapReq = soapReq + "<userId>" + userId + "</userId>"  
soapReq = soapReq +  "</aut:getUserRoleList>"  
soapReq = soapReq +"</soapenv:Body>"  
soapReq = soapReq + "</soapenv:Envelope>"
'response.write soapReq
 

if xmlHttp.readyState >= 1  then	  
	    ' try to connect to the server
	    'try
	      ' initiate server request
	       on error resume next
	      	   
	       xmlHTTP.send soapReq
	      '  response.write "Called WS"
	      ' display an error in case of failure
	       if err then
		   		DisplayError "BACK", "", err.Number, " getUserRoleList - web service is not available! Please check with your service admin.", err.Description
	       end if	
	    'catch (e)    
	else
		Response.write "<p>Cannot connect to server. readyState = " & xmlHttp.readyState & "</p>"
	end if

'	xmlHTTP.send soapReq
	
	strWSStatus = xmlHTTP.Status
	'Response.write "<p>HTTP Status = " & strWSStatus & "</p>"
strRetval = xmlHTTP.responseText
  'response.write strRetval



xmlDoc.loadxml(xmlHTTP.responseText)

 'response.write xmlDoc.getElementsByTagName("roleList").length
'response.write xmlDoc.getElementsByTagName("roleList").[0].text +"<br>"
'response.end

Set strRetval = xmlDoc.getElementsByTagName("roleList")
  'response.write strRetval.length
  'response.write strRetval.[0].text +"<br>"
  'response.end
  if strRetval.length=0 then
  	getUserRolePolicyList=0
  else
	dim policyLength, roleLength 

	policyLength=0
	roleLength=0


	For each node in strRetval	 
		if InStr(node.text,"SMA2")=1 then
		   policyLength=policyLength+1
		else
		   roleLength = roleLength + 1
		end if
	Next
	 ' response.write policyLength 
	 ' response.write "<BR>"
	 ' response.write roleLength 
	 ' response.end


	redim Preserve retRoleNames(rolelength)
	redim Preserve retPolicyNames(policyLength)
	rnindex=0 
	pnindex=0
	  'response.write policyLength 
	  'response.write roleLength 
	 'response.end
	For each node in strRetval
		if InStr(node.text,"SMA2")=1 then
		   retPolicyNames(pnindex)= node.text
		   pnindex = pnindex+1
		else
		   retRoleNames(rnindex)= node.text 	   
	 '	   response.write "Role " & rnindex & "=" &retRoleNames(rnindex)
		   rnindex = rnindex+1
		end if
	Next
	  'response.write rnindex
	  'response.write "policy index = "
	  'response.write pnindex
	  'response.end	 

	getUserRolePolicyList = strWSStatus
  end if	
end function

Public function Decrypt(strData)
 dim oEncryptedData, objWMIService, colItems
		  ' CAPICOM Constants                                                          
			Const CAPICOM_ENCRYPTION_ALGORITHM_RC2 = 0
			Const CAPICOM_ENCRYPTION_ALGORITHM_RC4 = 1
			Const CAPICOM_ENCRYPTION_ALGORITHM_DES = 2
			Const CAPICOM_ENCRYPTION_ALGORITHM_3DES = 3
			Const CAPICOM_ENCRYPTION_ALGORITHM_AES = 4
			Const CAPICOM_ENCRYPTION_KEY_LENGTH_MAXIMUM = 0
			Const CAPICOM_ENCRYPTION_KEY_LENGTH_40_BITS = 1
			Const CAPICOM_ENCRYPTION_KEY_LENGTH_56_BITS = 2
			Const CAPICOM_ENCRYPTION_KEY_LENGTH_128_BITS = 3
			Const CAPICOM_ENCRYPTION_KEY_LENGTH_192_BITS = 4
			Const CAPICOM_ENCRYPTION_KEY_LENGTH_256_BITS = 5 
			Const CAPICOM_SECRET_PASSWORD = 0
			Const CAPICOM_ENCODE_BASE64 = 0
			Const CAPICOM_ENCODE_BINARY = 1
			Const CAPICOM_ENCODE_ANY = -1
		
 			Set oEncryptedData = Server.CreateObject("CAPICOM.EncryptedData")
			oEncryptedData.Algorithm.KeyLength = CAPICOM_ENCRYPTION_KEY_LENGTH_256_BITS
			oEncryptedData.Algorithm.Name = CAPICOM_ENCRYPTION_ALGORITHM_AES
			
			' Set the secret to be used when deriving the key based on the unique server id
			Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
			Set colItems = objWMIService.ExecQuery ( "SELECT UUID FROM Win32_ComputerSystemProduct")
			oEncryptedData.SetSecret colItems.ItemIndex(0).UUID, CAPICOM_SECRET_PASSWORD
'response.write "<br>"
'response.write "strData: " & strData
'response.end
			oEncryptedData.Decrypt(strData)
			Decrypt = oEncryptedData.Content

'response.write "<br>"
'response.write "Decrypt: " & Decrypt
'response.end
end function

public function getConnString(str)
  dim connstring
  connstring=""
  Select case str
	case "strConstConnectString"
		connstring=strConstConnectString 
	case "strConstSConnectString"
		connstring=strConstSConnectString
	case Else
		connstring=eval(str)
  end select
  getConnString=connstring 
end function

'GIOVANNI AREVALO X207063 CREATING TOKEN FUNCTION'

public function generatingToken()

	dim xmlHTTP,xmlDoc,json,err,strWSStatus,strRetval,base64String

	'GIOVANNI AREVALO X207063 ADDING TOKEN VALUES'

	const client_id="17385b9f-82be-47b5-a596-08928b51f613"
	const client_secret="83378c17-29f8-4746-a938-5efa4b2ff9ffe5f29069-9239-41c9-ae59-884645aa50d1"
	const client_scope="1910"
	const token_URL="https://apigw-pr.tsl.telus.com/token"
	const client_credentials="client_credentials"
   
    'GIOVANNI AREVALO X207063 CALLING BASE64 function'
   	base64String = Base64Encode(client_id&":"&client_secret)
	

	Set xmlHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
	Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")

	xmlHTTP.SetOption SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS, SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
	xmlHTTP.open "POST",token_URL, false,client_id,client_secret
	xmlHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 
	xmlHTTP.setRequestHeader "Authorization","Basic "&base64String
    
	  
    
	json = "grant_type="&client_credentials &_
		   "&scope="&client_scope

	if xmlHttp.readyState >= 1  then	  
		    ' try to connect to the server
		    'try
		      ' initiate server request
			   on error resume next
		       xmlHTTP.send json
		      ' display an error in case of failure
		       if err then
			   		DisplayError "BACK", "", err.Number, " searchCustomer - web service is not available! Please check with your service admin.", err.Description
			   end if	
		    'catch (e)    
		else
			Response.write "<p>Cannot connect to server. readyState = " & xmlHttp.readyState & "</p>"
		end if

    strWSStatus = xmlHTTP.Status
    strRetval = xmlHTTP.responseText
	
    generatingToken= strRetval
	
end function

'GIOVANNI AREVALO x207063 adding BASE64 encode and decode functions'
public function Base64Encode(sText)
  Dim oXML, oNode
  Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
  Set oNode = oXML.CreateElement("base64")
  oNode.dataType = "bin.base64"
  oNode.nodeTypedValue = Stream_StringToBinary(sText)
  Base64Encode = Replace(Replace(oNode.text, Chr(10), ""), Chr(13), "")
  Set oNode = Nothing
  Set oXML = Nothing
end function

public function Base64Decode(ByVal vCode)
  Dim oXML, oNode
  Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
  Set oNode = oXML.CreateElement("base64")
  oNode.dataType = "bin.base64"
  oNode.text = vCode
  Base64Decode = Stream_BinaryToString(oNode.nodeTypedValue)
  Set oNode = Nothing
  Set oXML = Nothing
end function

private function Stream_StringToBinary(Text)
  Const adTypeText = 2
  Const adTypeBinary = 1
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")
  BinaryStream.Type = adTypeText
  BinaryStream.CharSet = "us-ascii"
  BinaryStream.Open
  BinaryStream.WriteText Text
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeBinary
  BinaryStream.Position = 0
  Stream_StringToBinary = BinaryStream.Read
  Set BinaryStream = Nothing
end function

private function Stream_BinaryToString(Binary)
  Const adTypeText = 2
  Const adTypeBinary = 1
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")
  BinaryStream.Type = adTypeBinary
  BinaryStream.Open
  BinaryStream.Write Binary
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeText
  BinaryStream.CharSet = "us-ascii"
  Stream_BinaryToString = BinaryStream.ReadText
  Set BinaryStream = Nothing
end function

'GIOVANNI AREVALO X207063 VBS JSON CLASS (THIS CLASS SCAN JSON STRING TO BE CONVERTED IN OBJECTS'
Class VbsJson
   
    Private Whitespace, NumberRegex, StringChunk
    Private b, f, r, n, t

    Private Sub Class_Initialize
        Whitespace = " " & vbTab & vbCr & vbLf
        b = ChrW(8)
        f = vbFormFeed
        r = vbCr
        n = vbLf
        t = vbTab

        Set NumberRegex = New RegExp
        NumberRegex.Pattern = "(-?(?:0|[1-9]\d*))(\.\d+)?([eE][-+]?\d+)?"
        NumberRegex.Global = False
        NumberRegex.MultiLine = True
        NumberRegex.IgnoreCase = True

        Set StringChunk = New RegExp
        StringChunk.Pattern = "([\s\S]*?)([""\\\x00-\x1f])"
        StringChunk.Global = False
        StringChunk.MultiLine = True
        StringChunk.IgnoreCase = True
    End Sub
    
    'Return a JSON string representation of a VBScript data structure
    'Supports the following objects and types
    '+-------------------+---------------+
    '| VBScript          | JSON          |
    '+===================+===============+
    '| Dictionary        | object        |
    '+-------------------+---------------+
    '| Array             | array         |
    '+-------------------+---------------+
    '| String            | string        |
    '+-------------------+---------------+
    '| Number            | number        |
    '+-------------------+---------------+
    '| True              | true          |
    '+-------------------+---------------+
    '| False             | false         |
    '+-------------------+---------------+
    '| Null              | null          |
    '+-------------------+---------------+
    Public Function Encode(ByRef obj)
        Dim buf, i, c, g
        Set buf = CreateObject("Scripting.Dictionary")
        Select Case VarType(obj)
            Case vbNull
                buf.Add buf.Count, "null"
            Case vbBoolean
                If obj Then
                    buf.Add buf.Count, "true"
                Else
                    buf.Add buf.Count, "false"
                End If
            Case vbInteger, vbLong, vbSingle, vbDouble
                buf.Add buf.Count, obj
            Case vbString
                buf.Add buf.Count, """"
                For i = 1 To Len(obj)
                    c = Mid(obj, i, 1)
                    Select Case c
                        Case """" buf.Add buf.Count, "\"""
                        Case "\"  buf.Add buf.Count, "\\"
                        Case "/"  buf.Add buf.Count, "/"
                        Case b    buf.Add buf.Count, "\b"
                        Case f    buf.Add buf.Count, "\f"
                        Case r    buf.Add buf.Count, "\r"
                        Case n    buf.Add buf.Count, "\n"
                        Case t    buf.Add buf.Count, "\t"
                        Case Else
                            If AscW(c) >= 0 And AscW(c) <= 31 Then
                                c = Right("0" & Hex(AscW(c)), 2)
                                buf.Add buf.Count, "\u00" & c
                            Else
                                buf.Add buf.Count, c
                            End If
                    End Select
                Next
                buf.Add buf.Count, """"
            Case vbArray + vbVariant
                g = True
                buf.Add buf.Count, "["
                For Each i In obj
                    If g Then g = False Else buf.Add buf.Count, ","
                    buf.Add buf.Count, Encode(i)
                Next
                buf.Add buf.Count, "]"
            Case vbObject
                If TypeName(obj) = "Dictionary" Then
                    g = True
                    buf.Add buf.Count, "{"
                    For Each i In obj
                        If g Then g = False Else buf.Add buf.Count, ","
                        buf.Add buf.Count, """" & i & """" & ":" & Encode(obj(i))
                    Next
                    buf.Add buf.Count, "}"
                Else
                    Err.Raise 8732,,"None dictionary object"
                End If
            Case Else
                buf.Add buf.Count, """" & CStr(obj) & """"
        End Select
        Encode = Join(buf.Items, "")
    End Function

    'Return the VBScript representation of ``str(``
    'Performs the following translations in decoding
    '+---------------+-------------------+
    '| JSON          | VBScript          |
    '+===============+===================+
    '| object        | Dictionary        |
    '+---------------+-------------------+
    '| array         | Array             |
    '+---------------+-------------------+
    '| string        | String            |
    '+---------------+-------------------+
    '| number        | Double            |
    '+---------------+-------------------+
    '| true          | True              |
    '+---------------+-------------------+
    '| false         | False             |
    '+---------------+-------------------+
    '| null          | Null              |
    '+---------------+-------------------+
    Public Function Decode(ByRef str)
        Dim idx
        idx = SkipWhitespace(str, 1)

        If Mid(str, idx, 1) = "{" Then
            Set Decode = ScanOnce(str, 1)
        Else
            Decode = ScanOnce(str, 1)
        End If
    End Function
    
    Private Function ScanOnce(ByRef str, ByRef idx)
        Dim c, ms

        idx = SkipWhitespace(str, idx)
        c = Mid(str, idx, 1)

        If c = "{" Then
            idx = idx + 1
            Set ScanOnce = ParseObject(str, idx)
            Exit Function
        ElseIf c = "[" Then
            idx = idx + 1
            ScanOnce = ParseArray(str, idx)
            Exit Function
        ElseIf c = """" Then
            idx = idx + 1
            ScanOnce = ParseString(str, idx)
            Exit Function
        ElseIf c = "n" And StrComp("null", Mid(str, idx, 4)) = 0 Then
            idx = idx + 4
            ScanOnce = Null
            Exit Function
        ElseIf c = "t" And StrComp("true", Mid(str, idx, 4)) = 0 Then
            idx = idx + 4
            ScanOnce = True
            Exit Function
        ElseIf c = "f" And StrComp("false", Mid(str, idx, 5)) = 0 Then
            idx = idx + 5
            ScanOnce = False
            Exit Function
        End If
        
        Set ms = NumberRegex.Execute(Mid(str, idx))
        If ms.Count = 1 Then
            idx = idx + ms(0).Length
            ScanOnce = CDbl(ms(0))
            Exit Function
        End If
        
        Err.Raise 8732,,"No JSON object could be ScanOnced"
    End Function

    Private Function ParseObject(ByRef str, ByRef idx)
        Dim c, key, value
        Set ParseObject = CreateObject("Scripting.Dictionary")
        idx = SkipWhitespace(str, idx)
        c = Mid(str, idx, 1)
        
        If c = "}" Then
            Exit Function
        ElseIf c <> """" Then
            Err.Raise 8732,,"Expecting property name"
        End If

        idx = idx + 1
        
        Do
            key = ParseString(str, idx)

            idx = SkipWhitespace(str, idx)
            If Mid(str, idx, 1) <> ":" Then
                Err.Raise 8732,,"Expecting : delimiter"
            End If

            idx = SkipWhitespace(str, idx + 1)
            If Mid(str, idx, 1) = "{" Then
                Set value = ScanOnce(str, idx)
            Else
                value = ScanOnce(str, idx)
            End If
            ParseObject.Add key, value

            idx = SkipWhitespace(str, idx)
            c = Mid(str, idx, 1)
            If c = "}" Then
                Exit Do
            ElseIf c <> "," Then
                Err.Raise 8732,,"Expecting , delimiter"
            End If

            idx = SkipWhitespace(str, idx + 1)
            c = Mid(str, idx, 1)
            If c <> """" Then
                Err.Raise 8732,,"Expecting property name"
            End If

            idx = idx + 1
        Loop

        idx = idx + 1
    End Function
    
    Private Function ParseArray(ByRef str, ByRef idx)
        Dim c, values, value
        Set values = CreateObject("Scripting.Dictionary")
        idx = SkipWhitespace(str, idx)
        c = Mid(str, idx, 1)

        If c = "]" Then
            ParseArray = values.Items
            Exit Function
        End If

        Do
            idx = SkipWhitespace(str, idx)
            If Mid(str, idx, 1) = "{" Then
                Set value = ScanOnce(str, idx)
            Else
                value = ScanOnce(str, idx)
            End If
            values.Add values.Count, value

            idx = SkipWhitespace(str, idx)
            c = Mid(str, idx, 1)
            If c = "]" Then
                Exit Do
            ElseIf c <> "," Then
                Err.Raise 8732,,"Expecting , delimiter"
            End If

            idx = idx + 1
        Loop

        idx = idx + 1
        ParseArray = values.Items
    End Function
    
    Private Function ParseString(ByRef str, ByRef idx)
        Dim chunks, content, terminator, ms, esc, char
        Set chunks = CreateObject("Scripting.Dictionary")

        Do
            Set ms = StringChunk.Execute(Mid(str, idx))
            If ms.Count = 0 Then
                Err.Raise 8732,,"Unterminated string starting"
            End If
            
            content = ms(0).Submatches(0)
            terminator = ms(0).Submatches(1)
            If Len(content) > 0 Then
                chunks.Add chunks.Count, content
            End If
            
            idx = idx + ms(0).Length
            
            If terminator = """" Then
                Exit Do
            ElseIf terminator <> "\" Then
                Err.Raise 8732,,"Invalid control character"
            End If
            
            esc = Mid(str, idx, 1)

            If esc <> "u" Then
                Select Case esc
                    Case """" char = """"
                    Case "\"  char = "\"
                    Case "/"  char = "/"
                    Case "b"  char = b
                    Case "f"  char = f
                    Case "n"  char = n
                    Case "r"  char = r
                    Case "t"  char = t
                    Case Else Err.Raise 8732,,"Invalid escape"
                End Select
                idx = idx + 1
            Else
                char = ChrW("&H" & Mid(str, idx + 1, 4))
                idx = idx + 5
            End If

            chunks.Add chunks.Count, char
        Loop

        ParseString = Join(chunks.Items, "")
    End Function

    Private Function SkipWhitespace(ByRef str, ByVal idx)
        Do While idx <= Len(str) And _
            InStr(Whitespace, Mid(str, idx, 1)) > 0
            idx = idx + 1
        Loop
        SkipWhitespace = idx
    End Function

End Class
%>