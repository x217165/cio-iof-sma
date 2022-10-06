<%@ Language=VBScript %>
<%   OPTION EXPLICIT
on error resume next %>
<!-- #include file="smaConstants.inc" -->
<!-- #include file="smaProcs.inc" -->
<!-- #include file="databaseconnect.asp" -->
<!--
*************************************************************************************************
* Page name:	STypeAttrList.asp																*
* Purpose:		To display the service instance attributes and its values inside a frame 										*														*
*																								*
* Created by:					Date															*
* Sara Sangha					02/15/2000														*
* ====================================															*
* Modifications By				Date				Modifcations								*
*
*************************************************************************************************
-->
<%

Dim objRs, strSQL, strORDER
Dim strServiceTypeID, strXRefID, strAction
Dim intAccessLevel

strServiceTypeID = Request("ServiceTypeID")
strXRefID = Request("Xrefid")
dim strRealUserID
strRealUserID = Session("username")

intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
if intAccessLevel and intConst_Access_ReadOnly <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access Service Type Attribute. Please contact your system administrator"
end if

strAction = Request("txtFrmAction")

''select case Request("txtFrmAction")


''	case "DELETE"
if strAction = "DELETE" then
' The following 3 lines temp commented for my test LC
	if ((intAccessLevel and intConst_Access_Delete) <> intConst_Access_Delete) then
	  DisplayError "BACK", "", 0, "DELETE DENIED", "You don't have access to delete Service Type Attribute. Please contact your system administrator"
	end if

   		if strXRefID <> "" then
			dim cmdDeleteObj
			set cmdDeleteObj = server.CreateObject("ADODB.Command")
			set cmdDeleteObj.ActiveConnection = objConn
			cmdDeleteObj.CommandType = adCmdStoredProc
			cmdDeleteObj.CommandText = "SMA_SP_USERID.Sp_Srvtype_Att_Val_Xref_Delete"

			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_userid", adVarChar , adParamInput, 20, Session("username"))
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_servicetype_Att_Val_xref_id",adNumeric , adParamInput,, Clng(strXRefID))

			'****************************
			'check parameter values
  			'****************************

  			'dim objparm
  			'for each objparm in cmdDeleteObj.Parameters
  			'	  Response.Write "<b>" & objparm.name & "</b>"
  			'	  Response.Write " has size:  " & objparm.Size & " "
  			'	  Response.Write " and value:  " & objparm.value & " "
  			'	  Response.Write " and datatype:  " & objparm.type & "<br> "
  			'next

  			'Response.Write "<b> count = " & cmdDeleteObj.Parameters.count & "<br>"
  			'dim nx
  			'for nx=0 to cmdDeleteObj.Parameters.count-1
  			'   Response.Write nx+1 & " parm value= " & cmdDeleteObj.Parameters.Item(nx).Value  & "<br>"
  			'next

  			'response.write (cmdDeleteObj.CommandText)
			'response.end
			cmdDeleteObj.Execute
			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT Delete RECORD", objConn.Errors(0).Description
				objConn.Errors.Clear
			end if
	end if

elseif strAction = "move" then
''	case "move"
	'move selected element up or down
	dim strDirection
	strDirection = Request("direction")
	if strXRefID <> "" then
		'create command object for Move stored proc
		dim cmdMoveObj
		set cmdMoveObj = server.CreateObject("ADODB.Command")
		set cmdMoveObj.ActiveConnection = objConn
		cmdMoveObj.CommandType = adCmdStoredProc
		cmdMoveObj.CommandText = "sma_sp_userid.SPK_SMA_ADMIN_INTER.sp_srvc_type_xref_move"
		'create params
		cmdMoveObj.Parameters.Append cmdMoveObj.CreateParameter("p_userid", adVarChar, adParamInput, 20, strRealUserID)
		cmdMoveObj.Parameters.Append cmdMoveObj.CreateParameter("p_srvc_type_att_val_xref_id", adNumeric , adParamInput,, CLng(strXRefID))
		cmdMoveObj.Parameters.Append cmdMoveObj.CreateParameter("p_direction", adVarChar, adParamInput, 10, ucase(Request("direction")))
'		if objConn.Errors.Count <> 0 then
'			DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT MOVE OBJECT - PARAMETER ERROR", objConn.Errors(0).Description
'			objConn.Errors.Clear
'		end if

		'****************************
		'check parameter values
		'****************************

		'dim objparm
		'for each objparm in cmdMoveObj.Parameters
		'     Response.Write "<b>" & objparm.name & "</b>"
		'	  Response.Write " has size:  " & objparm.Size & " "
		'	  Response.Write " and value:  " & objparm.value & " "
		'	  Response.Write " and datatype:  " & objparm.type & "<br> "
		'next

		'Response.Write "<b> count = " & cmdMoveObj.Parameters.count & "<br>"
		'dim nx
		'for nx=0 to cmdMoveObj.Parameters.count-1
		'   Response.Write nx+1 & " parm value= " & cmdMoveObj.Parameters.Item(nx).Value  & "<br>"
		'next

		'response.write (cmdMoveObj.CommandText)
		'response.end


		'call the Move stored proc
		cmdMoveObj.Execute
		if objConn.Errors.Count <> 0 then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT MOVE OBJECT", objConn.Errors(0).Description
			objConn.Errors.Clear
		end if
	end if
'' end select
end if


if isnumeric(strServiceTypeID)  then

	'Response.Write "Service Type :" & strServiceTypeID & "<P>"
	'Response.Write "Xref ID :" & strXRefID & "<P>"

	StrSql =" SELECT b.srvc_type_att_name, " &_
					"b.srvc_type_att_id, " &_
				    "c.SRVC_TYPE_ATT_VAL_NAME, " &_
				    "c.SRVC_TYPE_ATT_VAL_ID, " &_
				    "a.srvc_type_att_val_usage_id, " &_
				    "d.srvc_type_att_val_xref_id, " &_
				    "d.display_order " &_
			" FROM crp.srvc_type_att_val_usage a," &_
				   "crp.srvc_type_att b, " &_
				   "crp.srvc_type_att_val c, " &_
				   "crp.srvc_type_att_val_xref d " &_
			" WHERE a.srvc_type_att_id = b.srvc_type_att_id  AND " &_
				   "a.srvc_type_att_val_id = c.srvc_type_att_val_id " &_
			" AND  a.srvc_type_att_val_usage_id = d.srvc_type_att_val_usage_id " &_
			" AND  d.service_type_id = " & strServiceTypeID &_
			" AND  d.RECORD_STATUS_IND  = 'A' "

	strORDER = " order by d.display_order "

	StrSql = StrSql + strORDER

	'Response.Write "SQL :" & StrSql & "<P>"

	set objRs = objConn.Execute(strSQL)

	if err then
		'DisplayError "BACK", "", err.Number, "ERROR IN SELECTING SERVICE TYPE ATTRIBUTES", err.Description
		DisplayError "BACK", "", err.Number, "ERROR IN SQL", StrSql
		objConn.Errors.Clear
	end if

	'release connection
	set objRs.ActiveConnection = nothing

end if
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
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

function cell_onClick(intXRefID, intUsgeID, intatID, intatvID, intServiceType){

	document.frmIFR.txtXRefID.value = intXRefID;
	document.frmIFR.hdnUsageID.value = intUsgeID;
	document.frmIFR.txtattID.value = intatID;
	document.frmIFR.txtattvID.value = intatvID;
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
<form name="frmIFR" action="STypeAttrList.asp" method="POST">

		<input type=hidden name=hdnServiceTypeID value="">
		<input type=hidden name=txtXRefID value="">
	    <input type=hidden name=txtattID value="">
		<input type=hidden name=txtattvID value="">
		<input type=hidden name=hdnUsageID value="">
		<input type=hidden name=UpdateDateTime value="">


<TABLE border=1 cellspacing=0 cellpadding=2 width="100%">
	<thead>
		<th>Service Type Attribute</th>
		<th>Value</th>
		<th>Display Order</th>
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
		<td nowrap onClick="cell_onClick(<%=objRs(5)%>, <%=objRs(4)%>, <%=objRs(1)%>, <%=objRs(3)%>, <%=strServiceTypeID%>);"><%=objRs(0)%>&nbsp;</td>
		<td nowrap onClick="cell_onClick(<%=objRs(5)%>, <%=objRs(4)%>, <%=objRs(1)%>, <%=objRs(3)%>, <%=strServiceTypeID%>);"><%=objRs(2)%>&nbsp;</td>
		<td nowrap onClick="cell_onClick(<%=objRs(5)%>, <%=objRs(4)%>, <%=objRs(1)%>, <%=objRs(3)%>, <%=strServiceTypeID%>);"><%=objRs(6)%>&nbsp;</td>
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


