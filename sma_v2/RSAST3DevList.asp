<%@ Language=VBScript %>
<%  OPTION EXPLICIT
on error resume next
%>

<!-- #include file="smaConstants.inc" -->
<!-- #include file="smaProcs.inc" -->
<!-- #include file="databaseconnect.asp" -->

<!--
*************************************************************************************
* File Name:	RSAST3DevList.asp
*
* Purpose:		This page lists the devices for the iframe on the TC Detail page.
*
* In Param:
*
* Out Param:
*
* Created By:
**************************************************************************************

        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       10-23-01	     DTy		Change field names and variables.

								Local DNA to Local X25 DNA
								LOCAL_DNA   to LOCAL_X25_DNA

								Host DNA  to Host X25 DNA
								HOST_DNA_ID to HOST_X25_DNA_ID
								Add X25 Mnemonic column
								Add No. of POS Devices
**************************************************************************************
-->


<%
Const ASP_NAME = "RSAST3DevList.asp" 'only need to change this value when changing the filename

dim strTailCircuitID
dim rsDev
dim strSQL

Dim intAccessLevel
intAccessLevel = CInt(CheckLogon(strConst_RSAS))



'get the tail circuit id from the TC Detail Page
strTailCircuitID = Request("TailCircuitID")

'Response.Write "hdntailcircuitid =" & strTailCircuitID & "<BR>"
'Response.end

 
select case Request("action")
	
case "DELETE" 
	 
		if (Request("TailCircuitID") <> "") then
		    if (intAccessLevel and intConst_Access_Delete) <> intConst_Access_Delete then
				DisplayError "BACK", "", 0, "DELETE DENIED", "You don't have access to delete Tail Circuit Devices. Please contact your system administrator."
			end if
			
			dim cmdDeleteObj
			set cmdDeleteObj = server.CreateObject("ADODB.Command")
			set cmdDeleteObj.ActiveConnection = objConn
			cmdDeleteObj.CommandType = adCmdStoredProc
			cmdDeleteObj.CommandText = "sma_sp_userid.spk_sma_rsas_inter.sp_device_delete"
			
			'create the delete parameters 
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_device_id", adNumeric , adParamInput,, CInt(Request("hdnDeviceID")))
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_last_update", adDBTimeStamp, adParamInput,, CDate(Request("hdnUpdateDateTime")))
		
			cmdDeleteObj.Execute
			
			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			end if
		
		end if	
 
 end select
  			
			
if strTailCircuitID <> "" then
  
	strSQL =  " SELECT D.tail_circuit_id" &_
				" ,      D.device_id" &_
				" ,      D.local_x25_dna" &_
				" ,      D.poll_code" &_
				" ,      H.host_x25_dna_id" &_
				" ,      H.host_dna" &_
				" ,      H.host_dna_mnemonic" &_
				" ,		 D.update_date_time" &_ 
				" FROM   " &_
				"        CRP.RSAS_DEVICE D" &_
				" ,      CRP.RSAS_HOST_DNA H" &_
				" WHERE  D.TAIL_CIRCUIT_ID = " & strTailCircuitID &_
				" AND D.HOST_X25_DNA_ID = H.HOST_X25_DNA_ID (+)" 
				
					
		'Response.Write "<BR>" & strSQL	 & "<br>"
		'Response.end
		
		set rsDev=server.CreateObject("ADODB.Recordset")
	rsDev.CursorLocation = adUseClient
	rsDev.Open strSQL, objConn
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
	end if
	set rsDev.ActiveConnection = nothing
	
	'if  rsDev.EOF then
		'strTailCircuitID = ""
	'end if

	if err then
		DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
	end if

end if

%>

<HTML>

<HEAD>

<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK href="stylesheets/styles.css" rel="stylesheet" type="text/css">

<STYLE>
.regularItem {
	cursor: hand;
}
.whiteItem {
	cursor: hand;
	background-color: white;
}
.Highlight {
	cursor: hand; 
	background-color: #00974f;
	color: white;
}
</STYLE>

<script type="text/javascript">

var oldHighlightedElement;
var oldHighlightedElementClassName;

function cell_onClick(dtUpdate, intDeviceID, intTailCircuitID){
	
	document.frmIFR.hdnDeviceID.value = intDeviceID;
	document.frmIFR.hdnTailCircuitID.value = intTailCircuitID;
	document.frmIFR.hdnUpdateDateTime.value = dtUpdate; 
		
	//highlight current record
	if (oldHighlightedElement != null) 
	{
		oldHighlightedElement.className = oldHighlightedElementClassName
	}
	oldHighlightedElement = window.event.srcElement.parentElement;
	oldHighlightedElementClassName = oldHighlightedElement.className;
	oldHighlightedElement.className = "Highlight";
}


</script>

</HEAD>

<BODY>

<form name="frmIFR" action="<%=ASP_NAME%>" method="POST">

<input type="hidden" name="hdnDeviceID" value="">
<input type="hidden" name="hdnTailCircuitID" value="">
<input type="hidden" name="hdnUpdateDateTime" value="">
<input type="hidden" name="hdnDevCount" value="<%if strTailCircuitID <> "" then Response.write (rsDev.RecordCount) else Response.write (0)%>">

<TABLE border=0 cellspacing=0 frame=void cellpadding=0 width="100%">
 <thead>
	No. of POS Devices: <%if strTailCircuitID <> "" then Response.write (rsDev.RecordCount) else Response.write (0)%>
 </thead>
</TABLE>

<TABLE border=1 cellspacing=0 cellpadding=2 width="100%">
	<thead>
	
		<th nowrap>Local X25 DNA</th>
		<th nowrap>Poll Code</th>
		<th nowrap>Host X25 DNA</th>
		<th nowrap>X25 MNEMONIC</th>
			
	</thead>
	<tbody>
		<%
		
		dim k
		k = 0
		
		if strTailCircuitID <> "" then
			while not rsDev.EOF 
				if Int(k/2) = k/2 then
					Response.Write "<tr class=""regularItem"">"
				else
					Response.Write "<tr class=""whiteItem"">"
				end if
				k = k+1
				
					
			%>
					<!--td nowrap onClick="cell_onClick('<%=rsDev("UPDATE_DATE_TIME")%>',<%=strTailCircuitID%>);"><%=routineHTMLString(rsDev("tail_circuit_id"))%>&nbsp;</td-->
					<td nowrap onClick="cell_onClick('<%=rsDev("UPDATE_DATE_TIME")%>',<%=routineHTMLString(rsDev("device_id"))%>,<%=strTailCircuitID%>);"><%=routineHTMLString(rsDev("local_x25_dna"))%>&nbsp;</td>
					<td nowrap onClick="cell_onClick('<%=rsDev("UPDATE_DATE_TIME")%>',<%=routineHTMLString(rsDev("device_id"))%>,<%=strTailCircuitID%>);"><%=routineHTMLString(rsDev("poll_code"))%>&nbsp;</td>
					<td nowrap onClick="cell_onClick('<%=rsDev("UPDATE_DATE_TIME")%>',<%=routineHTMLString(rsDev("device_id"))%>,<%=strTailCircuitID%>);"><%=routineHTMLString(rsDev("host_dna"))%>&nbsp;</td>
					<td nowrap onClick="cell_onClick('<%=rsDev("UPDATE_DATE_TIME")%>',<%=routineHTMLString(rsDev("device_id"))%>,<%=strTailCircuitID%>);"><%=routineHTMLString(rsDev("host_dna_mnemonic"))%>&nbsp;</td>
				</tr>
			<%
				rsDev.MoveNext
			wend
			rsDev.Close
			set rsDev = Nothing
		end if
		%>
	</tbody>
</table>
</FORM>
</BODY>
</HTML>
