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
  * Name:		CustServContList.asp
  * Purpose:	This page list the customer service contact.
  * Created By:	?
  ***************************************************************************************************

        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       23-Feb-01	 DTy		Exclude customer service contacts are:
                                    - Marked as deleted in CONTACT, i.e.,
                                      RECORD_STATUS_IND='D' or STAFF_STATUS_LCODE<>'Departed'.
                                    - Marked as deleted in CUSTOMER_SERVICE_CONTACT, i.e.,
                                      RECORD_STATUS_IND='D'.
  ***************************************************************************************************
-->
<%
    stop
Const ASP_NAME = "CustServContList.asp" 'only need to change this value when changing the filename

dim strCustServID, objRsServiceContact, StrSql
Dim intAccessLevel, bolActiveOnly
    
Dim strRealUserID
strRealUserID = Session("username")
stop
intAccessLevel = CInt(CheckLogon(strConst_CustomerServiceContact))

strCustServID = Request("CustServID")

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
			cmdDeleteObj.CommandText = "sma_sp_userid.spk_sma_cserv_inter.sp_cs_cont_delete"

			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_customer_service_contact_id", adNumeric, adParamInput, , clng(Request("ContactID")))					'number(9)
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_last_update", adDBTimeStamp, adParamInput, ,Cdate(Request("hdnUpdateDateTime")))		'Date
            cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("P_USER_ID", adVarChar , adParamInput, 30, strRealUserID)

			cmdDeleteObj.Execute
			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			end if

		end if
 end select
  

if strCustServID <> "" then
  strSQL =  " SELECT con.contact_id" &_
			" ,      c.cust_serv_contact_type_lcode" &_
			" ,      l.cust_serv_contact_type_desc" &_
			" ,      c.contact_priority" &_
			" ,      con.contact_name" &_
			" ,      con.work_number" &_
			" ,      con.work_number_ext" &_
			" ,      con.cell_number" &_
			" ,      con.pager_number" &_
			" ,      con.fax_number" &_
			" ,      con.email_address" &_
			" ,      c.customer_service_contact_id" &_
			" ,      c.update_date_time " &_
			" FROM   crp.customer_service s" &_
			" ,      crp.customer_service_contact c" &_
			" ,      crp.contact con" &_
			" ,      crp.lcode_cust_serv_contact_type l " &_
			" WHERE  s.customer_service_id = c.customer_service_id " &_
			" AND    c.contact_id = con.contact_id " &_
			" AND    c.cust_serv_contact_type_lcode = l.cust_serv_contact_type_lcode " &_
			" AND    s.customer_service_id =  " &  CLng(strCustServID)

	if bolActiveOnly = "YES" then
		strSQL = strSQL & " AND c.record_status_ind = 'A' AND con.record_status_ind = 'A' " & _
			    " AND (con.staff_status_lcode is null " & _
			    " OR  (con.staff_status_lcode is not null AND con.staff_status_lcode <> 'Departed'))"
	end if

	strSQL = strSQL & " ORDER BY cust_serv_contact_type_lcode" &_
				", contact_priority "

	set objRsServiceContact = objConn.Execute(StrSql)
	if  objRsServiceContact.EOF then
		strCustServID = ""
	end if

	if err then
		DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET 32132", err.Description
	end if

end if

%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
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

function cell_onClick(dtUpdate, intCustServID, intContactID){

	document.frmIFR.hdnContactID.value = intContactID;
	document.frmIFR.hdnCustServID.value = intCustServID;
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
<input type="hidden" name="hdnCustServID" value="">
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
		if strCustServID <> "" then
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
					<td nowrap onClick="cell_onClick('<%=objRsServiceContact("UPDATE_DATE_TIME")%>',<%=strCustServID%>, <%=objRsServiceContact("CUSTOMER_SERVICE_CONTACT_ID")%>);"><%=routineHTMLString(objRsServiceContact("cust_serv_contact_type_lcode"))%>&nbsp;</td>
					<td nowrap onClick="cell_onClick('<%=objRsServiceContact("UPDATE_DATE_TIME")%>',<%=strCustServID%>, <%=objRsServiceContact("CUSTOMER_SERVICE_CONTACT_ID")%>);"><%=routineHTMLString(objRsServiceContact("cust_serv_contact_type_desc"))%>&nbsp;</td>
					<td nowrap onClick="cell_onClick('<%=objRsServiceContact("UPDATE_DATE_TIME")%>',<%=strCustServID%>, <%=objRsServiceContact("CUSTOMER_SERVICE_CONTACT_ID")%>);"><%=routineHTMLString(objRsServiceContact("contact_priority"))%>&nbsp;</td>
					<td nowrap onClick="cell_onClick('<%=objRsServiceContact("UPDATE_DATE_TIME")%>',<%=strCustServID%>, <%=objRsServiceContact("CUSTOMER_SERVICE_CONTACT_ID")%>);"><%=routineHTMLString(objRsServiceContact("contact_name"))%>&nbsp;</td>
					<td nowrap onClick="cell_onClick('<%=objRsServiceContact("UPDATE_DATE_TIME")%>',<%=strCustServID%>, <%=objRsServiceContact("CUSTOMER_SERVICE_CONTACT_ID")%>);"><%=routineHTMLString(strWP)%>&nbsp;</td>
					<td nowrap onClick="cell_onClick('<%=objRsServiceContact("UPDATE_DATE_TIME")%>',<%=strCustServID%>, <%=objRsServiceContact("CUSTOMER_SERVICE_CONTACT_ID")%>);"><%=routineHTMLString(objRsServiceContact("work_number_ext"))%>&nbsp;</td>
					<td nowrap onClick="cell_onClick('<%=objRsServiceContact("UPDATE_DATE_TIME")%>',<%=strCustServID%>, <%=objRsServiceContact("CUSTOMER_SERVICE_CONTACT_ID")%>);"><%=routineHTMLString(strCP)%>&nbsp;</td>
					<td nowrap onClick="cell_onClick('<%=objRsServiceContact("UPDATE_DATE_TIME")%>',<%=strCustServID%>, <%=objRsServiceContact("CUSTOMER_SERVICE_CONTACT_ID")%>);"><%=routineHTMLString(strPP)%>&nbsp;</td>
					<td nowrap onClick="cell_onClick('<%=objRsServiceContact("UPDATE_DATE_TIME")%>',<%=strCustServID%>, <%=objRsServiceContact("CUSTOMER_SERVICE_CONTACT_ID")%>);"><%=routineHTMLString(strFP)%>&nbsp;</td>
					<td nowrap onClick="cell_onClick('<%=objRsServiceContact("UPDATE_DATE_TIME")%>',<%=strCustServID%>, <%=objRsServiceContact("CUSTOMER_SERVICE_CONTACT_ID")%>);"><%=routineHTMLString(objRsServiceContact("email_address"))%>&nbsp;</td>
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


