<%@ LANGUAGE=VBSCRIPT %>
<%
OPTION EXPLICIT
on error resume next
%>
<!-- #include file="smaConstants.inc" -->
<!-- #include file="smaProcs.inc" -->
<!-- #include file="databaseconnect.asp" -->
<%
'got new sort criteria?
dim strSortCriteria
strSortCriteria = Request("txtSortCriteria")
'get real userid
dim strRealUserID
strRealUserID = Session("username")
strCustomerServiceID = Request("CustomerServiceID")

if strSortCriteria = "" then
'this is not a sort request, continue...

'get Customer Service ID
dim strCustomerServiceID
'strCustomerServiceID = Request("CustomerServiceID")
if strCustomerServiceID = "" then
	Response.End
end if

dim strNCflag
	strNCflag = "Y"

dim strDisable
strDisable = ""

dim sql

'get action item
dim strAction
strAction = Request("action")
if strAction = "add" then
	'add a new element
	dim strNewType, strNewID
	strNewType = Request("newType")
	strNewID = Request("newID")
	if strNewType <> "" and strNewID <> "" then
		'create command object for Delete stored proc
		dim cmdInsertObj
		set cmdInsertObj = server.CreateObject("ADODB.Command")
		set cmdInsertObj.ActiveConnection = objConn
		cmdInsertObj.CommandType = adCmdStoredProc
		cmdInsertObj.CommandText = "sma_sp_userid.spk_sma_corr_inter.sp_correlation_insert"
		'create params
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_userid", adVarChar , adParamInput, 20, strRealUserID)
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_cs_id", adNumeric , adParamInput,, CLng(strCustomerServiceID))
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_sort_order", adNumeric , adParamInput,, null)
		if strNewType <> "MO" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_ne_id", adNumeric , adParamInput,, null)
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_ne_id", adNumeric , adParamInput,, CLng(strNewID))
		end if
		if strNewType <> "FAC" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_circuit_id", adNumeric , adParamInput,, null)
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_circuit_id", adNumeric , adParamInput,, CLng(strNewID))
		end if
		if strNewType <> "Root" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_root_id", adNumeric , adParamInput,, null)
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_root_id", adNumeric , adParamInput,, CLng(strNewID))
		end if
'		if objConn.Errors.Count <> 0 then
'			DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT INSERT OBJECT - PARAMETER ERROR", objConn.Errors(0).Description
'			objConn.Errors.Clear
'		end if
		'call the Insert stored proc
		cmdInsertObj.Execute
		if objConn.Errors.Count <> 0 then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT INSERT OBJECT", objConn.Errors(0).Description
			objConn.Errors.Clear
		end if
	end if

elseif strAction = "delete" then
	'delete an existing element
	dim strDelObjID
	strDelObjID = Request("delObjID")
	if strDelObjID <> "" then
		'create command object for Delete stored proc
		dim cmdDeleteObj
		set cmdDeleteObj = server.CreateObject("ADODB.Command")
		set cmdDeleteObj.ActiveConnection = objConn
		cmdDeleteObj.CommandType = adCmdStoredProc
		cmdDeleteObj.CommandText = "sma_sp_userid.spk_sma_corr_inter.sp_correlation_delete"
		'create params
		cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_customer_service_id", adNumeric , adParamInput,, CLng(Request("delObjID")))
		cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , CDate(Request("txtUpdateDateTime")))
'		if objConn.Errors.Count <> 0 then
'			DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT - PARAMETER ERROR", objConn.Errors(0).Description
'			objConn.Errors.Clear
'		end if
		'call the Delete stored proc
		cmdDeleteObj.Execute
		if objConn.Errors.Count <> 0 then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", objConn.Errors(0).Description
			objConn.Errors.Clear
		end if
	end if

elseif strAction = "move" then
	'move selected element up or down
	dim strDirection
	strDirection = Request("direction")
	dim strObjID
	strObjID = Request("corrid")
	if strObjID <> "" then
		'create command object for Move stored proc
		dim cmdMoveObj
		set cmdMoveObj = server.CreateObject("ADODB.Command")
		set cmdMoveObj.ActiveConnection = objConn
		cmdMoveObj.CommandType = adCmdStoredProc
		cmdMoveObj.CommandText = "sma_sp_userid.spk_sma_corr_inter.sp_corr_move"
		'create params
		cmdMoveObj.Parameters.Append cmdMoveObj.CreateParameter("p_userid", adVarChar, adParamInput, 20, strRealUserID)
		cmdMoveObj.Parameters.Append cmdMoveObj.CreateParameter("p_man_corr_id", adNumeric , adParamInput,, CLng(strObjID))
		cmdMoveObj.Parameters.Append cmdMoveObj.CreateParameter("p_direction", adVarChar, adParamInput, 10, ucase(Request("direction")))
'		if objConn.Errors.Count <> 0 then
'			DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT MOVE OBJECT - PARAMETER ERROR", objConn.Errors(0).Description
'			objConn.Errors.Clear
'		end if
		'call the Move stored proc
		cmdMoveObj.Execute
		if objConn.Errors.Count <> 0 then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT MOVE OBJECT", objConn.Errors(0).Description
			objConn.Errors.Clear
		end if
	end if
end if


'get Correlated Elements and display the list
sql = "SELECT "&_
			"MC.MANAGED_CORRELATION_ID, "&_
			"MC.CUSTOMER_SERVICE_ID, "&_
			"MC.SORT_ORDER, "&_
			"MC.NETWORK_ELEMENT_ID				OBJ_ID, "&_
			"NE.NETWORK_ELEMENT_TYPE_CODE		OBJ_TYPE, "&_
			"NE.NETWORK_ELEMENT_NAME			OBJ_NAME, "&_
			"UPPER(NE.NETWORK_ELEMENT_NAME)		OBJ_NAME_UPPER, "&_
			"''									OBJ_NO, "&_
			"MC.RECORD_STATUS_IND, "&_
			"NE.OWNED_BY_NC 	OBJ_OWNED_BY_NC, "&_
			"sma_sp_userid.spk_sma_library.sf_get_full_username(MC.UPDATE_REAL_USERID) as UPDATE_REAL_USERID, "&_
			"TO_CHAR(MC.UPDATE_DATE_TIME,'MON-DD-YYYY HH24:MI:SS') UPDATE_DATE_TIME, "&_
			"'MO'								OBJ_CLASS "&_
		"FROM "&_
			"CRP.MANAGED_CORRELATION			MC, "&_
			"CRP.NETWORK_ELEMENT				NE "&_
		"WHERE "&_
			"MC.CUSTOMER_SERVICE_ID = " & strCustomerServiceID &_
			" AND MC.NETWORK_ELEMENT_ID = NE.NETWORK_ELEMENT_ID "&_
	"UNION "&_
		"SELECT "&_
			"MC.MANAGED_CORRELATION_ID, "&_
			"MC.CUSTOMER_SERVICE_ID, "&_
			"MC.SORT_ORDER, "&_
			"MC.CIRCUIT_ID						OBJ_ID, "&_
			"CI.CIRCUIT_TYPE_CODE				OBJ_TYPE, "&_
			"CI.CIRCUIT_NAME					OBJ_NAME, "&_
			"UPPER(CI.CIRCUIT_NAME)				OBJ_NAME_UPPER, "&_
			"CI.CIRCUIT_NUMBER					OBJ_NO, "&_
			"MC.RECORD_STATUS_IND, "&_
			"''						OBJ_OWNED_BY_NC,  "&_
			"sma_sp_userid.spk_sma_library.sf_get_full_username(MC.UPDATE_REAL_USERID) as UPDATE_REAL_USERID, "&_
			"TO_CHAR(MC.UPDATE_DATE_TIME,'MON-DD-YYYY HH24:MI:SS') UPDATE_DATE_TIME, "&_
			"'CIRCUIT'							OBJ_CLASS "&_
		"FROM "&_
			"CRP.MANAGED_CORRELATION			MC, "&_
			"CRP.CIRCUIT						CI "&_
		"WHERE "&_
			"MC.CUSTOMER_SERVICE_ID = " & strCustomerServiceID &_
			" AND MC.CIRCUIT_ID = CI.CIRCUIT_ID "&_
	"UNION "&_
		"SELECT "&_
			"MC.MANAGED_CORRELATION_ID, "&_
			"MC.CUSTOMER_SERVICE_ID, "&_
			"MC.SORT_ORDER, "&_
			"MC.ROOT_CUSTOMER_SERVICE_ID		OBJ_ID, "&_
			"'ROOT'								OBJ_TYPE, "&_
			"RT.CUSTOMER_SERVICE_DESC			OBJ_NAME, "&_
			"UPPER(RT.CUSTOMER_SERVICE_DESC)	OBJ_NAME_UPPER, "&_
			"''									OBJ_NO, "&_
			"MC.RECORD_STATUS_IND, "&_
			"''						OBJ_OWNED_BY_NC,  "&_
			"sma_sp_userid.spk_sma_library.sf_get_full_username(MC.UPDATE_REAL_USERID) as UPDATE_REAL_USERID, "&_
			"TO_CHAR(MC.UPDATE_DATE_TIME,'MON-DD-YYYY HH24:MI:SS') UPDATE_DATE_TIME, "&_
			"'ROOT'								OBJ_CLASS "&_
		"FROM "&_
			"CRP.MANAGED_CORRELATION			MC, "&_
			"CRP.CUSTOMER_SERVICE				RT "&_
		"WHERE "&_
			"MC.CUSTOMER_SERVICE_ID = " & strCustomerServiceID &_
			" AND MC.ROOT_CUSTOMER_SERVICE_ID = RT.CUSTOMER_SERVICE_ID "
	dim sqlNoSort
	sqlNoSort = sql
	sql = sql & "ORDER BY SORT_ORDER"

else
	'this is a sort request, restore the sql body and start from there
	sqlNoSort = unescape(Request("sqlNoSort"))
	select case	strSortCriteria
		case "sort"		sql = sqlNoSort & "ORDER BY SORT_ORDER"
		case "type"		sql = sqlNoSort & "ORDER BY OBJ_TYPE"
		case "name"		sql = sqlNoSort & "ORDER BY OBJ_NAME_UPPER"
		case "number"	sql = sqlNoSort & "ORDER BY OBJ_NO"
	end select
end if

if strSortCriteria = "" then strSortCriteria="sort"

'Response.Write sql
'Response.End

dim rsCorr
set rsCorr = Server.CreateObject("ADODB.Recordset")
rsCorr.CursorLocation = adUseClient
rsCorr.Open sql, objConn
if err then
	DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
end if
'release connection
set rsCorr.ActiveConnection = nothing

'If  rsCorr("OBJ_OWNED_BY_NC") = strNCflag Then
'	strDisable = "DISABLED"
'End If

'Response.Write rsCorr("OBJ_OWNED_BY_NC")
'Response.End

dim secondsql, secondCSID

secondCSID = Request("CustomerServiceID")
strDisable = Request("hdnDisable")

secondsql = "SELECT "&_
		"CS.CUSTOMER_SERVICE_ID, "&_
		"CS.SERVICE_TYPE_ID, "&_
		"ADDR.PROVINCE_STATE_LCODE, "&_
		"CUS.NOC_REGION_LCODE, "&_
		"ST.SEND_TO_NC_LCODE "&_
      "FROM "&_
		"CRP.CUSTOMER_SERVICE			CS, "&_
		"CRP.CUSTOMER				CUS, "&_
		"CRP.ADDRESS				ADDR, " &_
		"CRP.SERVICE_LEVEL_AGREEMENT	        SLA, "&_
		"CRP.SERVICE_LOCATION			SL, "&_
		"CRP.V_REMEDY_SUPPORT_GROUP		SG, " &_
		"CRP.SERVICE_TYPE			ST "&_
      "WHERE "&_
		"CS.SERVICE_TYPE_ID = ST.SERVICE_TYPE_ID "&_
		"AND CS.CUSTOMER_ID = CUS.CUSTOMER_ID "&_
		"AND CS.SERVICE_LEVEL_AGREEMENT_ID = SLA.SERVICE_LEVEL_AGREEMENT_ID "&_
		"AND CS.SERVICE_LOCATION_ID = SL.SERVICE_LOCATION_ID(+) "&_
		"AND CS.REMEDY_SUPPORT_GROUP_ID = SG.REMEDY_SUPPORT_GROUP_ID(+) " &_
		"AND SL.ADDRESS_ID = ADDR.ADDRESS_ID(+) " &_
		"AND CS.CUSTOMER_SERVICE_ID = '" & secondCSID & "'"

dim rsStype
set rsStype = Server.CreateObject("ADODB.Recordset")
rsStype.CursorLocation = adUseClient
rsStype.Open secondsql, objConn
if err then
	DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
end if
'release connection
set rsStype.ActiveConnection = nothing


'Response.Write strDisable
'Response.End

%>
<LINK href="stylesheets/styles.css" rel="stylesheet" type="text/css">
<HTML>
<TITLE>In-line Frame Page</TITLE>
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

function cell_onClick(intCorrID, intObjID, strObjType, strObjName, strObjClass, strObjLastUpdate){
	document.frmIFR.txtCorrID.value = intCorrID;
	document.frmIFR.txtObjID.value = intObjID;
	document.frmIFR.txtObjType.value = strObjType;
	document.frmIFR.txtObjName.value = strObjName;
	document.frmIFR.txtObjClass.value = strObjClass;
	document.frmIFR.txtLastUpdate.value = strObjLastUpdate;
	//highlight current record
	if (oldHighlightedElement != null) {oldHighlightedElement.className = oldHighlightedElementClassName}
//	alert(window.event.srcElement.id);
	oldHighlightedElement = window.event.srcElement.parentElement;
	oldHighlightedElementClassName = oldHighlightedElement.className;
	oldHighlightedElement.className = "Highlight";
//	If (rsStype("province_state_lcode") = "QC" and rsStype("NOC_REGION_LCODE") = "QUEBEC") then
//		strDisable = ""
//	else If  ((rsStype("SEND_TO_NC_LCODE") = "2" and rsCorr("OBJ_OWNED_BY_NC") = strNCflag) or (rsStype("SEND_TO_NC_LCODE") = "2" and rsCorr("OBJ_CLASS") <> "CIRCUIT") or  (rsStype("SEND_TO_NC_LCODE") = "2" and rsCorr("OBJ_CLASS") = "CIRCUIT" and rsCorr("OBJ_TYPE") = "ATMPVC")) Then
//		strDisable = "DISABLED"
//	else
//		strDisable = ""
//	  End If
//	End If
}

function go_sort(strCriteria){
	switch (strCriteria) {
		case 'sort':
			document.frmIFR.txtSortCriteria.value = 'sort';
			document.frmIFR.submit();
			break;
		case 'type':
			document.frmIFR.txtSortCriteria.value = 'type';
			document.frmIFR.submit();
			break;
		case 'name':
			document.frmIFR.txtSortCriteria.value = 'name';
			document.frmIFR.submit();
			break;
		case 'number':
			document.frmIFR.txtSortCriteria.value = 'number';
			document.frmIFR.submit();
			break;
	}
}
</script>

<body>
<form name="frmIFR" action="iFrmCorr.asp" method="POST">
<input type="hidden" name="txtCorrID" value="">
<input type="hidden" name="txtObjID" value="">
<input type="hidden" name="txtObjType" value="">
<input type="hidden" name="txtObjName" value="">
<input type="hidden" name="txtObjClass" value="">
<input type="hidden" name="sqlNoSort" value="<%=escape(sqlNoSort)%>">
<input type="hidden" name="txtSortCriteria" value="">
<input type="hidden" name="txtLastUpdate" value="">
<input type="hidden" name="CustomerServiceID" value="<%=strCustomerServiceID%>">
<input type="hidden" name="hdnDisable" value="<%=strDisable%>">

<TABLE border=1 cellspacing=0 cellpadding=2 width="100%">
	<thead>
		<th nowrap title="Click on the header to sort by this column." <%=strDisable%> onClick="go_sort('sort');">Sort<%if strSortCriteria = "sort" then Response.Write "<IMG SRC=""images/SORT_BY.GIF"">"%></th>
		<th nowrap title="Click on the header to sort by this column." <%=strDisable%> onClick="go_sort('type');">Type<%if strSortCriteria = "type" then Response.Write "<IMG SRC=""images/SORT_BY.GIF"">"%></th>
		<th nowrap title="Click on the header to sort by this column." <%=strDisable%> onClick="go_sort('name');">Object Name<%if strSortCriteria = "name" then Response.Write "<IMG SRC=""images/SORT_BY.GIF"">"%></th>
		<th nowrap title="Click on the header to sort by this column." <%=strDisable%> onClick="go_sort('number');">Circuit Number<%if strSortCriteria = "number" then Response.Write "<IMG SRC=""images/SORT_BY.GIF"">"%></th>
		<th nowrap title="Last update date/time">Update date</th>
		<th nowrap title="Last update user id">Updated By</th>
		<th nowrap title="Record status indicator">Status</th>
	</thead>
	<tbody>
		<%
		dim k
		k = 0
		while not rsCorr.EOF
			if strObjID = "" then strObjID = 0
			if Int(k/2) = k/2 then
 				Response.Write "<tr class=""regularItem"">"
			else
				Response.Write "<tr class=""whiteItem"">"
			end if
			k = k+1
			If (rsStype("province_state_lcode") = "QC" and rsStype("NOC_REGION_LCODE") = "QUEBEC") then
//Response.Write secondsql
				strDisable = ""
			else If  ((rsStype("SEND_TO_NC_LCODE") = "2" and rsCorr("OBJ_OWNED_BY_NC") = strNCflag) or (rsStype("SEND_TO_NC_LCODE") = "2" and rsCorr("OBJ_CLASS") <> "CIRCUIT") or  (rsStype("SEND_TO_NC_LCODE") = "2" and rsCorr("OBJ_CLASS") = "CIRCUIT" and rsCorr("OBJ_TYPE") = "ATMPVC")) Then
//Response.Write "Here 2"
					strDisable = "DISABLED"
				else
//Response.Write "Here 3"
					strDisable = ""
			     End If
			End If
			%>
			<td id="t<%=rsCorr("MANAGED_CORRELATION_ID")%>" nowrap <%=strDisable%> onClick="cell_onClick(<%=rsCorr("MANAGED_CORRELATION_ID")%>, <%=rsCorr("OBJ_ID")%>, '<%=routineJavaScriptString(rsCorr("OBJ_TYPE"))%>', '<%=routineJavaScriptString(rsCorr("OBJ_NAME"))%>', '<%=rsCorr("OBJ_CLASS")%>', '<%=rsCorr("UPDATE_DATE_TIME")%>');"><%=routineHTMLString(rsCorr("SORT_ORDER"))%>&nbsp;</td>
			<td nowrap <%=strDisable%> onClick="cell_onClick(<%=rsCorr("MANAGED_CORRELATION_ID")%>, <%=rsCorr("OBJ_ID")%>, '<%=routineJavaScriptString(rsCorr("OBJ_TYPE"))%>', '<%=routineJavaScriptString(rsCorr("OBJ_NAME"))%>', '<%=rsCorr("OBJ_CLASS")%>', '<%=rsCorr("UPDATE_DATE_TIME")%>');"><%=routineHTMLString(rsCorr("OBJ_TYPE"))%>&nbsp;</td>
			<td nowrap <%=strDisable%> onClick="cell_onClick(<%=rsCorr("MANAGED_CORRELATION_ID")%>, <%=rsCorr("OBJ_ID")%>, '<%=routineJavaScriptString(rsCorr("OBJ_TYPE"))%>', '<%=routineJavaScriptString(rsCorr("OBJ_NAME"))%>', '<%=rsCorr("OBJ_CLASS")%>', '<%=rsCorr("UPDATE_DATE_TIME")%>');"><%=routineHTMLString(rsCorr("OBJ_NAME"))%>&nbsp;</td>
			<td nowrap <%=strDisable%> onClick="cell_onClick(<%=rsCorr("MANAGED_CORRELATION_ID")%>, <%=rsCorr("OBJ_ID")%>, '<%=routineJavaScriptString(rsCorr("OBJ_TYPE"))%>', '<%=routineJavaScriptString(rsCorr("OBJ_NAME"))%>', '<%=rsCorr("OBJ_CLASS")%>', '<%=rsCorr("UPDATE_DATE_TIME")%>');"><%=routineHTMLString(rsCorr("OBJ_NO"))%>&nbsp;</td>
			<td nowrap <%=strDisable%> onClick="cell_onClick(<%=rsCorr("MANAGED_CORRELATION_ID")%>, <%=rsCorr("OBJ_ID")%>, '<%=routineJavaScriptString(rsCorr("OBJ_TYPE"))%>', '<%=routineJavaScriptString(rsCorr("OBJ_NAME"))%>', '<%=rsCorr("OBJ_CLASS")%>', '<%=rsCorr("UPDATE_DATE_TIME")%>');"><%=routineHTMLString(rsCorr("UPDATE_DATE_TIME"))%>&nbsp;</td>
			<td nowrap <%=strDisable%> onClick="cell_onClick(<%=rsCorr("MANAGED_CORRELATION_ID")%>, <%=rsCorr("OBJ_ID")%>, '<%=routineJavaScriptString(rsCorr("OBJ_TYPE"))%>', '<%=routineJavaScriptString(rsCorr("OBJ_NAME"))%>', '<%=rsCorr("OBJ_CLASS")%>', '<%=rsCorr("UPDATE_DATE_TIME")%>');"><%=routineHTMLString(rsCorr("UPDATE_REAL_USERID"))%>&nbsp;</td>
			<td nowrap <%=strDisable%> align=center onClick="cell_onClick(<%=rsCorr("MANAGED_CORRELATION_ID")%>, <%=rsCorr("OBJ_ID")%>, '<%=routineJavaScriptString(rsCorr("OBJ_TYPE"))%>', '<%=routineJavaScriptString(rsCorr("OBJ_NAME"))%>', '<%=rsCorr("OBJ_CLASS")%>', '<%=rsCorr("UPDATE_DATE_TIME")%>');"><%=routineHTMLString(rsCorr("RECORD_STATUS_IND"))%>&nbsp;</td>
			<%if CLng(strObjID) = CLng(rsCorr("MANAGED_CORRELATION_ID")) then%>
			<script type="text/javascript">
				window.t<%=rsCorr("MANAGED_CORRELATION_ID")%>.click();
			</script>
			<%end if%>
		</tr>
		<%
			rsCorr.MoveNext
		wend
		rsCorr.Close
		%>
	</tbody>
</table>
<%if strObjID <> "" then%>
<script type="text/javascript">
try {window.scroll(0,oldHighlightedElement.offsetTop);}
catch(e){}
</script>
<%end if%>
</FORM>
</BODY>
</HTML>