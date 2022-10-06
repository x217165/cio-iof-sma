<%@ Language=VBScript %>
<%   OPTION EXPLICIT
on error resume next %>
<!-- #include file="smaConstants.inc" -->
<!-- #include file="kenanconnect.asp" -->
<!-- #include file="smaProcs.inc" -->
<!-- #include file="databaseconnect.asp" -->
<!--
*************************************************************************************************
* Page name:	STypeKenanList.asp																*	
* Purpose:		To display the Kenan Component ID, Component Name, Package Name and the help text for each service type										*														*
*																								*
* Created by:					Date															*
* Sara Sangha					02/15/2000														*
* ====================================															*
* Modifications By				Date				Modifcations
* Anthony Cheung				10/06/2008			Adopting the Kenan names						*
*
*************************************************************************************************
-->
<%

Dim objRs, strSQL, kenanRS, strComponentID, sql
Dim strServiceTypeID, strXRefID
Dim intAccessLevel

strServiceTypeID = Request("hdnServiceTypeID")
strXRefID = Request("hdnXRefID")  'this is the component id when updating and deleting - AC- no this is the SERVICE_TYPE_KENAN_XREF_ID 10/13/08

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
		
			
		strSQL = "DELETE FROM CRP.service_type_kenan_XREF" &_
				" WHERE SERVICE_TYPE_ID = " & strServiceTypeID  &_
				" AND SERVICE_TYPE_KENAN_XREF_ID = " & strXRefID
				
	'	response.write(strSQL)
	'	response.end
		set objRs = objConn.Execute(strSQL)
		if err then
			DisplayError "BACK", "", err.Number, "CANNOT DELETE RECORD", err.Description
		end if
	end if	
 end select 			
			
		 
if isnumeric(strServiceTypeID)  then
	
	'Response.Write "Service Type :" & strServiceTypeID & "<P>"	

'Replace line 76 with STK.PACKAGE_ID after this field is added to STK table
		StrSql =" SELECT ST.SERVICE_TYPE_ID, " &_
				    "ST.SERVICE_TYPE_DESC, " &_
				    "STK.COMPONENT_ID, " &_
				    "STK.REP_HELP_TEXT, " &_
				    "STK.CREATE_DATE_TIME, " &_	
				    "STK.CREATE_REAL_USERID, " &_	
				    "STK.UPDATE_DATE_TIME, " &_	
				    "STK.UPDATE_REAL_USERID, " &_
				    "STK.SERVICE_TYPE_KENAN_XREF_ID, " &_	
   				    "STK.PACKAGE_ID " &_
   				    " FROM CRP.SERVICE_TYPE_KENAN_XREF STK," &_
				   "CRP.SERVICE_TYPE ST " &_
			" WHERE ST.SERVICE_TYPE_ID = STK.SERVICE_TYPE_ID " &_
			" AND  ST.service_type_id = " & strServiceTypeID 
	'		response.write strSql
	'		response.end
	set objRs = objConn.Execute(strSql)
	if err then
		DisplayError "BACK", "", err.Number, "ERROR IN SELECTING Kenan ATTRIBUTES", err.Description
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

function cell_onClick(intXRefID, intServiceType, intKenanCID, intKenanPID){

	document.frmIFR.txtXRefID.value = intXRefID;
//	document.frmIFR.hdnUsageID.value = intUsgeID; 
//	document.frmIFR.txtattID.value = intatID;
//	document.frmIFR.txtattvID.value = intatvID;
	document.frmIFR.hdnServiceTypeID.value = intServiceType;
	document.frmIFR.hdnKenanCompID.value = intKenanCID;
	document.frmIFR.hdnKenanPackID.value = intKenanPID;

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
		<input type=hidden name=hdnKenanCompID value="">
		<input type=hidden name=hdnKenanPackID value="">
		<input type=hidden name=txtXRefID value="">
		<input type=hidden name=UpdateDateTime value="">


<TABLE border=1 cellspacing=0 cellpadding=4 width="100%">
	<thead>
		<th>Package ID</th>
		<th>Package Name</th>
		<th>Component ID</th>
		<th>Component Name</th>
		<th>Help Text</th>
	</thead>
	<tbody>
		
<%	if isnumeric(strServiceTypeID)  then

		dim k
		k = 0
		while not objRs.EOF
			sql = "SELECT PACKAGE_NAME, COMPONENT_NAME FROM ARBOR.V_PKG_COMPONENTS WHERE COMPONENT_ID = " & objRs(2)
			if (objRs(9) <> "") then
				sql = sql + "AND PACKAGE_ID = " & objRs(9)				
			end if
			set kenanRS = objKenanConn.Execute(sql)
			if err then
				DisplayError "BACK", "", err.Number, "ERROR IN SELECTING Kenan data", err.Description
			end if

			if Int(k/2) = k/2 then
				Response.Write "<tr class=""regularItem"">"
			else
				Response.Write "<tr class=""whiteItem"">"
			end if
			k = k+1 %>
		<td nowrap onClick="cell_onClick(<%=objRs(8)%>, <%=strServiceTypeID%>, <%response.write objRs(2)%>, 
		<% if objRs(9)<>"" then response.write objRs(9) else response.write 0 %>);"><%=objRs(9)%>&nbsp;</td>
		<td nowrap onClick="cell_onClick(<%=objRs(8)%>, <%=strServiceTypeID%>, <%=objRs(2)%>, 
		<% if objRs(9)<>"" then response.write objRs(9) else response.write 0 %>);"><%=kenanRS(0)%>&nbsp;</td>
		
		<td nowrap onClick="cell_onClick(<%=objRs(8)%>, <%=strServiceTypeID%>, <%=objRs(2)%>, 
		<% if objRs(9)<>"" then response.write objRs(9) else response.write 0 %>);"><%=objRs(2)%>&nbsp;</td>

		<td nowrap onClick="cell_onClick(<%=objRs(8)%>, <%=strServiceTypeID%>, <%=objRs(2)%>, 
		<% if objRs(9)<>"" then response.write objRs(9) else response.write 0 %>);"><%=kenanRS(1)%>&nbsp;</td>
		<td nowrap onClick="cell_onClick(<%=objRs(8)%>, <%=strServiceTypeID%>, <%=objRs(2)%>, 
		<% if objRs(9)<>"" then response.write objRs(9) else response.write 0 %>);"><%=objRs(3)%>&nbsp;</td>
		</tr>
			<% kenanRS.Close
			set kenanRS = Nothing
			objRs.MoveNext
		wend
		
		objRs.Close
		set objRs = Nothing
		
 end if %>
</tbody>
</table>
</FORM>
</BODY>
</HTML>


