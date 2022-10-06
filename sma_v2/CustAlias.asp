<%@ LANGUAGE=VBSCRIPT %>
<%  
OPTION EXPLICIT
on error resume next
Response.CacheControl="Private"
%>
<!-- #include file="smaConstants.inc" -->
<!-- #include file="smaProcs.inc" -->
<!-- #include file="databaseconnect.asp" -->
<%
dim strCustomerID, sql
strCustomerID = Request("CustomerID")

'get the name alias recordset
if isNumeric(strCustomerID)then
	dim rsAlias
	sql = "SELECT CMA.CUSTOMER_NAME_ALIAS_ID, CMA.CUSTOMER_NAME_ALIAS_UPPER, CMA.UPDATE_DATE_TIME FROM CRP.CUSTOMER_NAME_ALIAS CMA WHERE CMA.CUSTOMER_ID = " & strCustomerID &_
	      " AND CMA.RECORD_STATUS_IND = 'A' " &_
	      " ORDER BY CUSTOMER_NAME_ALIAS_UPPER "
	      
	'response.write Session("ConnectString")
	'response.write "<BR>"
	'response.write sql
	'response.end	

	set rsAlias=server.CreateObject("ADODB.Recordset")
	rsAlias.CursorLocation = adUseClient
	rsAlias.Open sql, objConn
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
	end if
	set rsAlias.ActiveConnection = nothing
end if
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
<form name="frmIFR" action="CustAlias.asp" method="POST">
<input type="hidden" name="hdnNameAliasID" value="">
<input type="hidden" name="hdnLastUpdate" value="">

<TABLE border=1 cellspacing=0 cellpadding=2 width="100%">
	<thead>
		<th nowrap title="Customer Name Alias">Name Alias</th>
	</thead>
	<tbody>
		<%
		if isNumeric(strCustomerID) then
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