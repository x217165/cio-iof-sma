<%@ LANGUAGE=VBSCRIPT %>
<%
OPTION EXPLICIT
on error resume next
%>
<!-- #include file="smaConstants.inc" -->
<!-- #include file="smaProcs.inc" -->
<!-- #include file="databaseconnect.asp" -->
<%

'get strNE_ID - Network Element ID
dim strNE_ID
strNE_ID = Request("ne_id")

dim sql

'get action item
'dim strAction
'strAction = Request("action")
'if strAction = "add" then
'	'add a new element
'	dim strNewType, strNewID
'	strNewType = Request("newType")
'	strNewID = Request("newID")
'	if strNewType <> "" and strNewID <> "" then
'		sql = "INSERT INTO CRP.MANAGED_CORRELATION (CUSTOMER_SERVICE_ID, SORT_ORDER, USER_ENTERED_FLAG, "
'		if strNewType = "Root" then
'			sql = sql & "ROOT_CUSTOMER_SERVICE_ID"
'		elseif (strNewType = "MO") then
'			sql = sql & "NETWORK_ELEMENT_ID"
'		elseif (strNewType = "FAC") then
'			sql = sql & "CIRCUIT_ID"
'		end if
'		sql = sql & ") VALUES (" & strCustomerServiceID & ", 0, 'N', " & strNewID & ")"
'	end if
''	Response.Write "@"&strNewType&"@"
''	Response.Write sql
'
'	objConn.BeginTrans
'	objConn.execute sql
'
'	if objConn.errors.count <> 0 then
'		Response.Write "There was an error on insert"
'		objConn.Rollback
'	else
'		objConn.Rollback
'	end if
'
'elseif strAction = "delete" then
'	'delete an existing element
'	dim strDelObjID
'	strDelObjID = Request("delObjID")
'	if strDelObjID <> "" then
'		sql = "delete from crp.managed_correlation where managed_correlation_id = " & strDelObjID
'
'		objConn.BeginTrans
'		objConn.execute sql
'
'		if objConn.errors.count <> 0 then
'			Response.Write "There was an error on delete"
'			objConn.Rollback
'		else
'			objConn.Rollback
'		end if
'
'	end if
'end if

'get the name alias recordset
if strNE_ID <> "" then
	dim rsAlias
	sql = "SELECT NETWORK_ELEMENT_NAME_ALIAS_ID, NETWORK_ELEMENT_NAME_ALIAS, UPDATE_DATE_TIME FROM CRP.NETWORK_ELEMENT_NAME_ALIAS WHERE NETWORK_ELEMENT_ID = " & strNE_ID
	set rsAlias=server.CreateObject("ADODB.Recordset")
	rsAlias.CursorLocation = adUseClient
	rsAlias.Open sql, objConn
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
	end if
	set rsAlias.ActiveConnection = nothing
end if
%>
<HTML>
<LINK href="stylesheets/styles.css" rel="stylesheet" type="text/css">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
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

function cell_onClick(intAliasID, strLastUpdate){
	document.frmIFR.hdnNameAliasID.value = intAliasID;
	document.frmIFR.hdnLastUpdate.value = strLastUpdate;
	//highlight current record
	if (oldHighlightedElement != null) {oldHighlightedElement.className = oldHighlightedElementClassName}
	oldHighlightedElement = window.event.srcElement.parentElement;
	oldHighlightedElementClassName = oldHighlightedElement.className;
	oldHighlightedElement.className = "Highlight";
}

</script>

<body>
<form name="frmIFR" action="manobjalias.asp" method="POST">
<input type="hidden" name="hdnNameAliasID" value="">
<input type="hidden" name="hdnLastUpdate" value="">

<TABLE border=1 cellspacing=0 cellpadding=2 width="100%">
	<thead>
		<th nowrap title="Network element name alias">Name Alias</th>
	</thead>
	<tbody>
		<%
		if strNE_ID <> "" then
		dim k
		k = 0
		while not rsAlias.EOF
			if Int(k/2) = k/2 then
				Response.Write "<tr class=""regularItem"">"
			else
				Response.Write "<tr class=""whiteItem"">"
			end if
			k = k+1
		%>
			<td nowrap onClick="cell_onClick(<%=rsAlias(0)%>, '<%=rsAlias(2)%>')"><%=rsAlias(1)%>&nbsp;</td>
		</tr>
		<%
		rsAlias.MoveNext
		wend
		rsAlias.Close
		end if
		%>
	</tbody>
</TABLE>

</FORM>
</BODY>
</HTML>