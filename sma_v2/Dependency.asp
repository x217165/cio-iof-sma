<%@ Language=VBScript %>
<% option explicit %>
<% Response.Buffer = true %>
<!--#include file="smaConstants.inc"-->
<!--#include file="smaProcs.inc"-->
<!--#include file="databaseconnect.asp" -->
<%


dim strOwner, strTableName, strRecordID, strResults

strOwner = Request.QueryString("Owner")
strTableName = Request.QueryString("TableName")
strRecordID = Request.QueryString("RecordID")   

	dim cmdViewObj
	set cmdViewObj = server.CreateObject("ADODB.Command")
	set cmdViewObj.ActiveConnection = objConn
	cmdViewObj.CommandType = adCmdStoredProc
	cmdViewObj.CommandText = "sma_sp_userid.spk_sma_library.sp_web_dependencies"
		
	'create params 
	cmdViewObj.Parameters.Append cmdViewObj.CreateParameter("p_owner", adVarChar, adParamInput, 30, strOwner) 							
	cmdViewObj.Parameters.Append cmdViewObj.CreateParameter("p_table_name", adVarChar, adParamInput, 30, strTableName) 							
	cmdViewObj.Parameters.Append cmdViewObj.CreateParameter("p_record_id", adNumeric  , adParamInput, 20, strRecordID) 							
	cmdViewObj.Parameters.Append cmdViewObj.CreateParameter("p_results", adVarChar, adParamOutput, 30000, strResults) 							
	
			
  	'dim objparm
  	'for each objparm in cmdViewObj.Parameters
  	'	  Response.Write "<b>" & objparm.name & "</b>"
  	'	  Response.Write " has size:  " & objparm.Size & " "
  	'	  Response.Write " and value:  " & objparm.value & " "
  	'	  Response.Write " and datatype:  " & objparm.Type & "<br> "
  	'next
  	'Response.Write "<b> count = " & cmdViewObj.Parameters.count & "<br>"
  	'dim nx
  	'for nx=0 to cmdViewObj.Parameters.count-1
  	'   Response.Write  " parm " & nx + 1 &  " value= " & cmdViewObj.Parameters.Item(nx).Value  & "<br>"
  	'next 
	on error resume next
		cmdViewObj.Execute

	if objConn.Errors.Count <> 0 then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "Can not view dependencies", objConn.Errors(0).Description
			objConn.Errors.Clear
	else
			strResults = cmdViewObj.Parameters("p_results").Value
	end if
	
	
	Response.Write "<TITLE>References</TITLE>"
	Response.Write "<BODY>"
	Response.Write "<TABLE border=1 cellPadding=2 cellSpacing=0>"
	Response.Write "<THEAD><TR ><TH colspan=2>The " & strTableName & " is referenced by the following records." & "</TH></TR><THEAD>"
	Response.Write "<TBODY>"
	Response.Write strResults 
	Response.Write "</TBODY></TABLE>"
	Response.Write "</BODY>"
	
	
	' get details on the records.
	'set cmdViewObj = server.CreateObject("ADODB.Command")
	'set cmdViewObj.ActiveConnection = objConn
	'cmdViewObj.CommandType = adCmdStoredProc
	'cmdViewObj.CommandText = "sma_sp_userid.spk_sma_library.sp_check_dependencies"
		
	'create params 
	'cmdViewObj.Parameters.Append cmdViewObj.CreateParameter("p_owner", adVarChar, adParamInput, 30, strOwner) 							
	'cmdViewObj.Parameters.Append cmdViewObj.CreateParameter("p_table_name", adVarChar, adParamInput, 30, strTableName) 							
	'cmdViewObj.Parameters.Append cmdViewObj.CreateParameter("p_record_id", adNumeric  , adParamInput, 20, strRecordID) 							
	'cmdViewObj.Parameters.Append cmdViewObj.CreateParameter("p_results", adVarChar, adParamOutput, 4000, strResults) 							
	
			
  	'dim objparm
  	'for each objparm in cmdViewObj.Parameters
  	'	  Response.Write "<b>" & objparm.name & "</b>"
  	'	  Response.Write " has size:  " & objparm.Size & " "
  	'	  Response.Write " and value:  " & objparm.value & " "
  	'	  Response.Write " and datatype:  " & objparm.Type & "<br> "
  	'next
  	'Response.Write "<b> count = " & cmdViewObj.Parameters.count & "<br>"
  	'dim nx
  	'for nx=0 to cmdViewObj.Parameters.count-1
  	'   Response.Write  " parm " & nx + 1 &  " value= " & cmdViewObj.Parameters.Item(nx).Value  & "<br>"
  	'next 
	'on error resume next
	'	cmdViewObj.Execute

	'if objConn.Errors.Count <> 0 then
	'		DisplayError "BACK", "", objConn.Errors(0).NativeError, "Can not view dependencies", objConn.Errors(0).Description
	'		objConn.Errors.Clear
	'else
	'		strResults = cmdViewObj.Parameters("p_results").Value
	'end if
	
	'Response.Write "Below is the datailed information on dependent records." & "<br>"
	'Response.Write "--------------------------------------------------------------" & "<br>"
	'Response.Write strResults
	
	
			
%>


<HTML>
<HEAD>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="stylesheets/styles.css">
<script type="text/javascript" SRC="AccessLevels.js"></script> 
<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>

<SCRIPT LANGUAGE=javascript>
<!--
//******************************************** End of Java Functions *****************************
//-->
</SCRIPT>
</HEAD>
</HTML>
