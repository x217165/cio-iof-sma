<%@ Language=VBScript %>
<%  
OPTION EXPLICIT
on error resume next
%>
<!-- #include file="smaConstants.inc" -->
<!-- #include file="smaProcs.inc" -->
<!-- #include file="databaseconnect.asp" -->
<%
'Get Circuit Id?

dim strCircuitID,objRsFacilityAlias,StrSql
Dim StrAliasID,StrCircuitTyp

strCircuitID = Request("CircuitID")
StrCircuitTyp = Request("FacType")

dim intAccessLevel

IF StrCircuitTyp = "ATMPVC" THEN
intAccessLevel = CInt(CheckLogon(strConst_PVC))
ELSE
 intAccessLevel = CInt(CheckLogon(strConst_Facilities))
END IF

if intAccessLevel and intConst_Access_ReadOnly <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to PVC/Facilities. Please contact your system administrator"
end if


 select case Request("txtFrmAction")
	case "DELETE"  
	if ((intAccessLevel and intConst_Access_Delete) <> intConst_Access_Delete) then
	  DisplayError "BACK", "", 0, "DELETE DENIED", "You don't have access to delete Facility Alias. Please contact your system administrator"
	end if
	 
    if (Request("AliasID") <>"") then
	    	
			dim cmdDeleteObj
			set cmdDeleteObj = server.CreateObject("ADODB.Command")
			set cmdDeleteObj.ActiveConnection = objConn
			cmdDeleteObj.CommandType = adCmdStoredProc
			cmdDeleteObj.CommandText = "sma_sp_userid.spk_sma_fac_inter.sp_fac_alias_delete"
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_circuit_alias_id", adNumeric, adParamInput,,Clng(Request("AliasID")))	
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , CDate(Request("hdnUpdateDateTime")))
			
			cmdDeleteObj.Execute
			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE FACILITY", objConn.Errors(0).Description
				objConn.Errors.Clear
			end if
		    'strCircuitID = 0 'Request("CircuitID")
			'strWinMessage = "Record deleted successfully."
			'else
		      'DisplayError "BACK", "", err.Number, "ACCESS DENIED. PLEASE CONTACT YOUR SYSTEM ADMINISTRATOR.", err.Description
	      'end if
		
	end if	
 end select 			
			
			
			

if strCircuitID = "" then
	Response.End
end if



StrSql = "SELECT CIRCUIT_NUMBER_ALIAS_ID,CIRCUIT_ID,CIRCUIT_NUMBER_ALIAS,CIRCUIT_PROVIDER_CODE,UPDATE_DATE_TIME FROM CRP.CIRCUIT_NUMBER_ALIAS WHERE RECORD_STATUS_IND = 'A' AND CIRCUIT_ID =" & StrCircuitID & "  ORDER BY UPPER(CIRCUIT_NUMBER_ALIAS)"
 
set objRsFacilityAlias = objConn.Execute(StrSql)

if err then
	DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET 32132", err.Description
end if
'release connection


%>

<HTML>
<HEAD>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
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

function cell_onClick(dtUpdate,intAliasID,intCircuitID){
	document.frmIFR.txtAliasID.value = intAliasID;
	document.frmIFR.txtCircuitID.value = intCircuitID;
	document.frmIFR.hdnUpdateDateTime.value = dtUpdate; 
	//highlight current record
	if (oldHighlightedElement != null) {oldHighlightedElement.className = oldHighlightedElementClassName}
	oldHighlightedElement = window.event.srcElement.parentElement;
	oldHighlightedElementClassName = oldHighlightedElement.className;
	oldHighlightedElement.className = "Highlight";
}

</script>

</HEAD>
<BODY>
<form name="frmIFR" action="FacilityAlias.asp" method="POST">
<input type="hidden" name="txtAliasID" value="">
<input type="hidden" name="txtCircuitID" value="">
<input type="hidden" name="hdnUpdateDateTime" value="">

<TABLE border=1 cellspacing=0 cellpadding=2 width="100%">
	<thead>
		<th nowrap title="Last update user id">Circuit Alias</th>
		<th nowrap title="Record status indicator">Provider Code</th>
	</thead>
	<tbody>
		<%
		dim k
		k = 0
		while not objRsFacilityAlias.EOF
			if Int(k/2) = k/2 then
				Response.Write "<tr class=""regularItem"">"
			else
				Response.Write "<tr class=""whiteItem"">"
			end if
			k = k+1
		%>
			<td nowrap onClick="cell_onClick('<%=objRsFacilityAlias("UPDATE_DATE_TIME")%>',<%=objRsFacilityAlias("CIRCUIT_NUMBER_ALIAS_ID")%>, <%=objRsFacilityAlias("CIRCUIT_ID")%>);"><%=objRsFacilityAlias("CIRCUIT_NUMBER_ALIAS")%>&nbsp;</td>
			<td nowrap onClick="cell_onClick('<%=objRsFacilityAlias("UPDATE_DATE_TIME")%>',<%=objRsFacilityAlias("CIRCUIT_NUMBER_ALIAS_ID")%>, <%=objRsFacilityAlias("CIRCUIT_ID")%>);"><%=objRsFacilityAlias("CIRCUIT_PROVIDER_CODE")%>&nbsp;</td>
			
		</tr>
		<%
		objRsFacilityAlias.MoveNext
		wend
		objRsFacilityAlias.Close
		set objRsFacilityAlias = Nothing
		%>
	</tbody>
</table>
</FORM>
</BODY>
</HTML>


