<%@ Language=VBScript %>
<%   OPTION EXPLICIT
on error resume next %>
<!-- #include file="smaConstants.inc" -->
<!-- #include file="smaProcs.inc" -->
<!-- #include file="databaseconnect.asp" -->
<!--
*************************************************************************************************
* Page name:	STypeSLAList.asp																*	
* Purpose:		To display the default SLA for each region										*														*
*																								*
* Created by:					Date															*
* Sara Sangha					02/15/2000														*
* ====================================															*
* Modifications By				Date				Modifcations								*
*
*************************************************************************************************
-->
<%

Dim objRs, strSQL, strWinMessage
Dim strServiceTypeID, strXRefID, strUpdateDateTime, strNOCRegionLcode
Dim intAccessLevel


strServiceTypeID = Request("ServiceTypeID")
'Response.Write "Service type : " & strServiceTypeID
'Response.End

strXRefID = Request("XRefID")
strUpdateDateTime = Request("UpdateDateTime")

intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
if intAccessLevel and intConst_Access_ReadOnly <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access Service Type. Please contact your system administrator"
end if

select case Request("txtFrmAction")

	
	case "DELETE"  
	
	if ((intAccessLevel and intConst_Access_Delete) <> intConst_Access_Delete) then
	  DisplayError "BACK", "", 0, "DELETE DENIED", "You don't have access to delete Facility Alias. Please contact your system administrator"
	end if
	 
   if strXRefID <> "" then
		
			dim cmdDeleteObj
			set cmdDeleteObj = server.CreateObject("ADODB.Command")
			set cmdDeleteObj.ActiveConnection = objConn
			
			cmdDeleteObj.CommandType = adCmdStoredProc
			cmdDeleteObj.CommandText = "sma_sp_userid.spk_sma_admin_inter.sp_servtype_region_xref_delete"
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_servtype_region_xref_id", adNumeric, adParamInput,,Clng(strXRefID))	
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , CDate(strUpdateDateTime))
						
			cmdDeleteObj.Execute
			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE THE RECORD.", objConn.Errors(0).Description
				objConn.Errors.Clear
			else
				strWinMessage = "Record deleted successfully."
			end if
		
	end if	
 end select 			
			

		 
if isnumeric(strServiceTypeID)  then
	
	'Response.Write "Service Type :" & strServiceTypeID & "<P>"	
	
	StrSql = " SELECT  " &_
				 "X.SERVICE_TYPE_ID, " &_
				 "X.SERVICETYPE_REGION_XREF_ID, " &_
				 "R.REVENUE_REGION_DESC,   " &_
				 "A.SERVICE_LEVEL_AGREEMENT_DESC,"  &_
				 "X.UPDATE_DATE_TIME " &_
		  " FROM  CRP.SERVICETYPE_REGION_XREF X,  " &_
				"CRP.SERVICE_TYPE T,  " &_
				"CRP.SERVICE_LEVEL_AGREEMENT A, " &_
				"SO.LCODE_REVENUE_REGION R  " &_
		  " WHERE  T.SERVICE_TYPE_ID = X.SERVICE_TYPE_ID  " &_
		  " AND	  X.REGION_LCODE = R.REVENUE_REGION_LCODE  " &_
		  " AND	   X.SERVICE_LEVEL_AGREEMENT_ID = A.SERVICE_LEVEL_AGREEMENT_ID " &_	   
		  " AND	  X.RECORD_STATUS_IND = 'A'  " &_
		  " AND	  T.SERVICE_TYPE_ID = " & strServiceTypeID &_
		  " ORDER BY REVENUE_REGION_DESC"
 
	'Response.Write(strSQL)
	'Response.End 
 
	set objRs = objConn.Execute(strSQL)
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET 32132", err.Description
	end if

end if
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK href="stylesheets/styles.css" rel="stylesheet" type="text/css">
<STYLE>
.regularItem {
	cursor: hand;
}
.whiteItem {
	cursor: hand;
	background-color: white; }
	
.Highlight {
	cursor: hand; 
	background-color: #00974f;
	color: white;
}
</STYLE>

<script type="text/javascript">

var oldHighlightedElement;
var oldHighlightedElementClassName;

function cell_onClick(dtUpdate,intXRefID,intServiceType){

	document.frmIFR.txtXRefID.value = intXRefID;
	document.frmIFR.hdnUpdateDateTime.value = dtUpdate; 
	document.frmIFR.hdnServiceTypeID.value = intServiceType;
	//highlight current record

	if (oldHighlightedElement != null) {
		oldHighlightedElement.className = oldHighlightedElementClassName
	}
	oldHighlightedElement = window.event.srcElement.parentElement;
	oldHighlightedElementClassName = oldHighlightedElement.className;
	oldHighlightedElement.className = "Highlight";

}

</script>

</HEAD>
<BODY>
<form name="frmIFR" action="STypeSLAList.asp" method="POST">

		<input type=hidden name=hdnServiceTypeID value="">
		<input type=hidden name=txtXRefID value="">
		<input type=hidden name=txtSLADesc value="">
		<input type=hidden name=hdnUpdateDateTime value="">

<TABLE border=1 cellspacing=0 cellpadding=2 width="100%">
	<thead>
		<th nowrap title="Last update user id">Region</th>
		<th nowrap title="Record status indicator">SLA Description</th>
	</thead>
	<tbody>
		
<%	if isnumeric(strServiceTypeID)  then

		dim k
		k = 0
		while not objRs.EOF
			if Int(k/2) = k/2 then
				Response.Write "<tr class=""regularItem"">"
			else
				Response.Write "<tr class=""whiteItem"">"
			end if
			k = k+1 %>
			<td nowrap onClick="cell_onClick('<%=objRs("UPDATE_DATE_TIME")%>',<%=objRs("SERVICETYPE_REGION_XREF_ID")%>, <%=objRs("SERVICE_TYPE_ID")%>);"><%=objRs("REVENUE_REGION_DESC")%>&nbsp;</td>
			<td nowrap onClick="cell_onClick('<%=objRs("UPDATE_DATE_TIME")%>',<%=objRs("SERVICETYPE_REGION_XREF_ID")%>, <%=objRs("SERVICE_TYPE_ID")%>);"><%=objRs("SERVICE_LEVEL_AGREEMENT_DESC")%>&nbsp;</td>
			</tr>
			<% objRs.MoveNext
		wend
		
		objRs.Close
		set objRs = Nothing
		
 end if %>
</tbody>
</table>
</FORM>
</BODY>
</HTML>


