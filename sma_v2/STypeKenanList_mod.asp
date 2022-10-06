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

Dim objRs, strSQL, kenanRS, strComponentID. sql
Dim strServiceTypeID, strXRefID
Dim intAccessLevel

strServiceTypeID = Request("ServiceTypeID")
strXRefID = Request("XRefID")

'set kenanRS = server.createobject("ADODB.RecordSet")

intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
if intAccessLevel and intConst_Access_ReadOnly <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access Service Type Attribute. Please contact your system administrator"
end if

select case Request("txtFrmAction")

	
	case "DELETE"  
' The following 3 lines temp commented for my test LC	
	if ((intAccessLevel and intConst_Access_Delete) <> intConst_Access_Delete) then
	  DisplayError "BACK", "", 0, "DELETE DENIED", "You don't have access to delete Service Type Attribute. Please contact your system administrator"
	end if
	 
   if strXRefID <> "" then
		
			
		strSQL = "DELETE FROM CRP.SRVC_TYPE_ATT_VAL_XREF" &_
				" WHERE SERVICE_TYPE_ID = " & strServiceTypeID  &_
				" AND SRVC_TYPE_ATT_VAL_XREF_ID = " & strXRefID
		set objRs = objConn.Execute(strSQL)
		if err then
			DisplayError "BACK", "", err.Number, "CANNOT DELETE RECORD", err.Description
		end if
	end if	
 end select
		 
if isnumeric(strServiceTypeID)  then
	
'	Response.Write "Service Type :" & strServiceTypeID & "<P>"	
	
'	StrSql =" SELECT ST.SERVICE_TYPE_ID, " &_
'				    "ST.SERVICE_TYPE_DESC, " &_
'				    "STK.COMPONENT_ID, " &_
'				    "STK.REP_HELP_TEXT, " &_
'				    "STK.UPDATE_DATE_TIME " &_					
'			" FROM CRP.SERVICE_TYPE_KENAN_XREF STK," &_
'				   "CRP.SERVICE_TYPE ST " &_
'			" WHERE ST.SERVICE_TYPE_ID = STK.SERVICE_TYPE_ID " &_
'			" AND  ST.service_type_id = " & strServiceTypeID 

	StrSql =" SELECT b.srvc_type_att_name, " &_
					"b.srvc_type_att_id, " &_
				    "c.SRVC_TYPE_ATT_VAL_NAME, " &_
				    "c.SRVC_TYPE_ATT_VAL_ID, " &_
				    "a.srvc_type_att_val_usage_id, " &_
				    "d.srvc_type_att_val_xref_id " &_
			" FROM crp.srvc_type_att_val_usage a," &_
				   "crp.srvc_type_att b, " &_
				   "crp.srvc_type_att_val c, " &_
				   "crp.srvc_type_att_val_xref d " &_
			" WHERE a.srvc_type_att_id = b.srvc_type_att_id  AND " &_
				   "a.srvc_type_att_val_id = c.srvc_type_att_val_id " &_
			" AND  a.srvc_type_att_val_usage_id = d.srvc_type_att_val_usage_id " &_
			" AND  d.service_type_id = " & strServiceTypeID 
 
	set objRs = objConn.Execute(strSQL)
	if err then
		DisplayError "BACK", "", err.Number, "ERROR IN SELECTING Kenan Attributes", err.Description
	end if
	
end if

'Response.Write (objRs(2))
'Response.End 

'strComponentID = objRs(2)

'if isnumeric(strComponentID) then	
'	sql = "SELECT PACKAGE_NAME, COMPONENT_NAME FROM ARBOR.V_PKG_COMPONENTS WHERE COMPONENT_ID = " & strComponentID

'Response.Write (sql)
'Response.End 

'	kenanRS.Open sql, objKenanConn, adOpenForwardOnly, adLockReadOnly, adCmdText
'	set kenanRS = objConn.Execute(sql)


'	if err then
'		DisplayError "BACK", "", err.Number, "ERROR IN SELECTING Kenan data", err.Description
'	end if
	
'end if					



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
<form name="frmIFR" action="STypeKenanList.asp" method="POST">

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


