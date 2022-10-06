<%@ Language=VBScript %>
<%  
OPTION EXPLICIT
on error resume next
%>
<!-- #include file="smaConstants.inc" -->
<!-- #include file="smaProcs.inc" -->
<!-- #include file="databaseconnect.asp" -->
<!--        
  ***************************************************************************************************
  * Name:		ServLocContact.asp
  * Purpose:	This page list the service location contacts.
  * Created By:	?
  ***************************************************************************************************

        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       23-Feb-01	 DTy		List contacts based on active status selected by the user.
	                            Exclude service location contacts that are:
                                    - Marked as deleted in CONTACT, i.e.,
                                      RECORD_STATUS_IND='D' or STAFF_STATUS_LCODE<>'Departed'.
                                    - Marked as deleted in SERVICE_LOCATION_CONTACT, i.e.,
                                      RECORD_STATUS_IND='D'.
--> 
<%
'Get Circuit Id?
Const ASP_NAME = "ServLocContact.asp" 'only need to change this value when changing the filename

dim strServLocID, objRsServiceContact, StrSql
Dim intAccessLevel, bolActiveOnly
Dim strRealUserID
strRealUserID = Session("username")

intAccessLevel = CInt(CheckLogon(strConst_ServiceLocation))

strServLocID = Request("ServLocID")
bolActiveOnly = Request("ActiveOnly")

 select case Request("txtFrmAction")
	case "DELETE" 
	 
		if (Request("ContactID") <> "") then
		    if (intAccessLevel and intConst_Access_Delete) <> intConst_Access_Delete then
				DisplayError "BACK", "", 0, "DELETE DENIED", "You don't have access to delete Service Location Contacts. Please contact your system administrator."
			end if
			
			dim cmdDeleteObj
			set cmdDeleteObj = server.CreateObject("ADODB.Command")
			set cmdDeleteObj.ActiveConnection = objConn

			cmdDeleteObj.CommandType = adCmdStoredProc
			cmdDeleteObj.CommandText = "sma_sp_userid.spk_sma_cust_inter.sp_sl_cont_delete"

			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_service_location_contact_id", adNumeric, adParamInput, , clng(Request("ContactID")))					'number(9)	
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_last_update", adDBTimeStamp, adParamInput, ,Cdate(Request("hdnUpdateDateTime")))		'Date
            cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("P_USER_ID", adVarChar , adParamInput, 30, strRealUserID)

		
			cmdDeleteObj.Execute
			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			end if
				
			strWinMessage = "Record deleted successfully."
		end if	
 end select 			
			
if strServLocID <> "" then
	strSQL = " select con.contact_id" &_
				   ", c.serv_loc_contact_type_lcode" &_
				   ", l.serv_loc_contact_type_desc" &_
				   ", c.contact_priority" &_
				   ", con.contact_name" &_
				   ", con.work_number" &_
				   ", con.work_number_ext" &_
				   ", con.cell_number" &_
				   ", con.pager_number" &_
				   ", con.fax_number" &_
				   ", con.email_address" &_ 
				   ", c.service_location_contact_id" &_
				   ", c.update_date_time" &_
			   " from crp.service_location s " &_
				   ", crp.service_location_contact c " &_
				   ", crp.contact con " &_
				   ", crp.lcode_serv_loc_contact_type l  " &_
			" where s.service_location_id = c.service_location_id " &_ 
			" and c.contact_id = con.contact_id " &_
			" and c.serv_loc_contact_type_lcode = l.serv_loc_contact_type_lcode " &_
			" and s.service_location_id = "  &  strServLocID
	
	If bolActiveOnly = "YES" then
		strSQL = strSQL & " and c.record_status_ind = 'A' and con.record_status_ind = 'A' " & _
			    " and (con.staff_status_lcode is null " & _
			    " or (con.staff_status_lcode is not null and con.staff_status_lcode <> 'Departed'))" & _
				" order by serv_loc_contact_type_lcode, contact_priority" 
	end if
	
	set objRsServiceContact = objConn.Execute(StrSql)
	if  objRsServiceContact.EOF then
		strServLocID = ""
	end if

end if

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

function cell_onClick(dtUpdate, intServLocID, intContactID){
	
	document.frmIFR.hdnContactID.value = intContactID;
	document.frmIFR.hdnServLocID.value = intServLocID;
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
<input type="hidden" name="hdnContactID" value="">
<input type="hidden" name="hdnServLocID" value="">
<input type="hidden" name="hdnUpdateDateTime" value="">

<TABLE border=1 cellspacing=0 cellpadding=2 width="100%">
	<thead>
		<th nowrap>Role</th>
		<th nowrap>Role Desc</th>
		<th nowrap>Priority</th>
		<th nowrap>Contact Name</th>
		<th nowrap>Work Number</th>
		<th nowrap>Ext</th>
		<th nowrap>Cell Number</th>
		<th nowrap>Pager</th>
		<th nowrap>Fax Number</th>
		<th nowrap>E-mail Address</th>
	</thead>
	<tbody>
		<%
		Dim strWPArea,strWPMid,strWPEnd,strWP
		Dim strCPArea,strCPMid,strCPEnd,strCP
		Dim strPPArea,strPPMid,strPPEnd,strPP
		Dim strFPArea,strFPMid,strFPEnd,strFP
		
		dim k
		k = 0
		if strServLocID <> "" then
			while not objRsServiceContact.EOF
				if Int(k/2) = k/2 then
					Response.Write "<tr class=""regularItem"">"
				else
					Response.Write "<tr class=""whiteItem"">"
				end if
				k = k+1
	
		
				'format the work phone number
	 			strWPArea = mid(objRsServiceContact("work_number"),1,3)
	 			strWPMid = mid(objRsServiceContact("work_number"),4,3)
	 			strWPEnd = mid(objRsServiceContact("work_number"),7,4)
	 			strWP = "(" & strWPArea & ") " & strWPMid & "-" & strWPEnd
	 			If strWP = "() -" then
	 				strWP = ""
	 			End If
	
				'format the cell phone number
				strCPArea = mid(objRsServiceContact("cell_number"),1,3)
	 			strCPMid = mid(objRsServiceContact("cell_number"),4,3)
	 			strCPEnd = mid(objRsServiceContact("cell_number"),7,4)
	 			strCP = "(" & strCPArea & ") " & strCPMid & "-" & strCPEnd
	 			If strCP = "() -" then
	 				strCP = ""
	 			End If

	
				'format the pager number
				strPPArea = mid(objRsServiceContact("pager_number"),1,3)
	 			strPPMid = mid(objRsServiceContact("pager_number"),4,3)
	 			strPPEnd = mid(objRsServiceContact("pager_number"),7,4)
	 			strPP = "(" & strPPArea & ") " & strPPMid & "-" & strPPEnd
	 			If strPP = "() -" then
	 				strPP = ""
	 			End If
					
				'format the fax number
	 			strFPArea = mid(objRsServiceContact("fax_number"),1,3)
	 			strFPMid = mid(objRsServiceContact("fax_number"),4,3)
	 			strFPEnd = mid(objRsServiceContact("fax_number"),7,4)
	 			strFP = "(" & strFPArea & ") " & strFPMid & "-" & strFPEnd
	 			If strFP = "() -" then
	 				strFP = ""
	 			End If
					
			%>
					<td nowrap onClick="cell_onClick('<%=routineJavaScriptString(t("UPDATE_DATE_TIME"))%>',<%=strServLocID%>, <%=objRsServiceContact("SERVICE_LOCATION_CONTACT_ID")%>);"><%=routineHTMLString(objRsServiceContact("serv_loc_contact_type_lcode"))%>&nbsp;</td>
					<td nowrap onClick="cell_onClick('<%=routineJavaScriptString(objRsServiceContact("UPDATE_DATE_TIME"))%>',<%=strServLocID%>, <%=objRsServiceContact("SERVICE_LOCATION_CONTACT_ID")%>);"><%=routineHTMLString(objRsServiceContact("serv_loc_contact_type_desc"))%>&nbsp;</td>
					<td nowrap onClick="cell_onClick('<%=routineJavaScriptString(objRsServiceContact("UPDATE_DATE_TIME"))%>',<%=strServLocID%>, <%=objRsServiceContact("SERVICE_LOCATION_CONTACT_ID")%>);"><%=objRsServiceContact("contact_priority")%>&nbsp;</td>
					<td nowrap onClick="cell_onClick('<%=routineJavaScriptString(oobjRsServiceContacbjRsServiceContact("UPDATE_DATE_TIME"))%>',<%=strServLocID%>, <%=objRsServiceContact("SERVICE_LOCATION_CONTACT_ID")%>);"><%=routineHTMLString(objRsServiceContact("contact_name"))%>&nbsp;</td>
					<td nowrap onClick="cell_onClick('<%=routineJavaScriptString(objRsServiceContact("UPDATE_DATE_TIME"))%>',<%=strServLocID%>, <%=objRsServiceContact("SERVICE_LOCATION_CONTACT_ID")%>);"><%=routineHTMLString(strWP)%>&nbsp;</td>
					<td nowrap onClick="cell_onClick('<%=routineJavaScriptString(objRsServiceContact("UPDATE_DATE_TIME"))%>',<%=strServLocID%>, <%=objRsServiceContact("SERVICE_LOCATION_CONTACT_ID")%>);"><%=routineHTMLString(objRsServiceContact("work_number_ext"))%>&nbsp;</td>
					<td nowrap onClick="cell_onClick('<%=routineJavaScriptString(objRsServiceContact("UPDATE_DATE_TIME"))%>',<%=strServLocID%>, <%=objRsServiceContact("SERVICE_LOCATION_CONTACT_ID")%>);"><%=routineHTMLString(strCP)%>&nbsp;</td>
					<td nowrap onClick="cell_onClick('<%=routineJavaScriptString(objRsServiceContact("UPDATE_DATE_TIME"))%>',<%=strServLocID%>, <%=objRsServiceContact("SERVICE_LOCATION_CONTACT_ID")%>);"><%=routineHTMLString(strPP)%>&nbsp;</td>
					<td nowrap onClick="cell_onClick('<%=routineJavaScriptString(objRsServiceContact("UPDATE_DATE_TIME"))%>',<%=strServLocID%>, <%=objRsServiceContact("SERVICE_LOCATION_CONTACT_ID")%>);"><%=routineHTMLString(strFP)%>&nbsp;</td>
					<td nowrap onClick="cell_onClick('<%=routineJavaScriptString(objRsServiceContact("UPDATE_DATE_TIME"))%>',<%=strServLocID%>, <%=objRsServiceContact("SERVICE_LOCATION_CONTACT_ID")%>);"><%=routineHTMLString(objRsServiceContact("email_address"))%>&nbsp;</td>
				</tr>
			<%
				objRsServiceContact.MoveNext
			wend
			objRsServiceContact.Close
			set objRsServiceContact = Nothing
		end if
		%>
	</tbody>
</table>
</FORM>
</BODY>
</HTML>


