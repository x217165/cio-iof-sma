<%@ Language=VBScript %>
<% Option Explicit 
  on error resume next 
%>
<% Response.Buffer = true %>
<!--#include file="SmaConstants.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--#include file="smaProcs.inc"-->
<%



dim cmdUpdateObj,strEmailSubject,strEmailBody

  strEmailSubject = unescape(Request("subject"))
  strEmailBody =  unescape(Request("body"))
  

	set cmdUpdateObj = server.CreateObject("ADODB.Command")
	set cmdUpdateObj.ActiveConnection = objConn
	cmdUpdateObj.CommandType = adCmdStoredProc
	cmdUpdateObj.CommandText = "sma_sp_userid.spk_sma_fac_inter.sp_fac_send_email"
			'create parameters
			
	   cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_circuit_id",adNumeric , adParamInput,, Clng(Request("CircuitID")))
	   
	   if Request("CustServID") <>"" then
	   cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_customer_service_id",adNumeric , adParamInput,, Clng(Request("CustServID")))	
	   else
	   cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_customer_service_id",adNumeric , adParamInput,, null)	
	   end if	
	   			
	   cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_mail_list", adVarChar, adParamOutput, 2000)
	   cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_alias_list", adVarChar, adParamOutput, 2000)
	   cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_order_no", adVarChar, adParamOutput, 10)
	        
       cmdUpdateObj.Execute
       
	if objConn.Errors.Count <> 0 then
		DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT UPDATE OBJECT", objConn.Errors(0).Description
		objConn.Errors.Clear
	else
		dim strEmailFrom, strEmailTo,strOrderNo,strAliasList,strAdslMainStream,strDOCSorder
		
		if strEmailSubject <> "" then
			'it's time to send an email
			strOrderNo = cmdUpdateObj.Parameters("p_order_no").Value
		    strAliasList = escape(cmdUpdateObj.Parameters("p_alias_list").Value)
			strEmailTo = escape(cmdUpdateObj.Parameters("p_mail_list").Value)
			strAdslMainStream ="ADSL MainStream Circuit Number: " & unescape(strAliasList) &vbCrLf
			strDOCSorder = "DOCS Order Number: " & strOrderNo &vbCrLf
			
			strEmailBody = strEmailBody & strAdslMainStream & strDOCSorder
			
			Response.Cookies("txtEmailTo") = unescape(strEmailTo)
			Response.Cookies("txtEmailSubject") = strEmailSubject
			Response.Cookies("txtEmailBody") = escape(strEmailBody)
		end if
	end if

%>



<script type="text/javascript">
<%if strEmailSubject <> "" then%>
//pop-up the email window  
 document.location = 'email.asp';
<%end if%>
</script>
 