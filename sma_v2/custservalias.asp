<%@  language="VBSCRIPT" %>
<%
OPTION EXPLICIT
on error resume next
%>
<!-- #include file="smaConstants.inc" -->
<!-- #include file="smaProcs.inc" -->
<!-- #include file="databaseconnect.asp" -->
<!--
********************************************************************************************
* Page name:	CustServAlias.asp
* Purpose:		To display Customer Service Name alias.
*
* Created by:	Dan S. Ty	02/27/2002
********************************************************************************************
        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
***************************************************************************************************
-->

<%

'get strCS_ID - Customer Service ID
dim strCS_ID
strCS_ID = Request("cs_id")
   
dim sql

'end if
    
'get the name alias recordset
if strCS_ID <> "" then
	dim rsAlias
	sql = "SELECT CUSTOMER_SERVICE_DESC_ALIAS_ID, CUSTOMER_SERVICE_DESC_ALIAS, UPDATE_DATE_TIME FROM CRP.CUSTOMER_SERVICE_DESC_ALIAS WHERE CUSTOMER_SERVICE_ID = " & strCS_ID & "ORDER BY 2"
	set rsAlias=server.CreateObject("ADODB.Recordset")
	rsAlias.CursorLocation = adUseClient
	rsAlias.Open sql, objConn
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
	end if
	set rsAlias.ActiveConnection = nothing
end if
%>
<html>
<link href="stylesheets/styles.css" rel="stylesheet" type="text/css">
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>In-line Frame Page</title>
<style>
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
</style>

<script type="text/javascript">
    var oldHighlightedElement;
    var oldHighlightedElementClassName;

    function cell_onClick(intAliasID, strLastUpdate) {
        document.frmIFR.hdnNameAliasID.value = intAliasID;
        document.frmIFR.hdnLastUpdate.value = strLastUpdate;
        //highlight current record
        if (oldHighlightedElement != null) { oldHighlightedElement.className = oldHighlightedElementClassName }
        oldHighlightedElement = window.event.srcElement.parentElement;
        oldHighlightedElementClassName = oldHighlightedElement.className;
        oldHighlightedElement.className = "Highlight";
    }

</script>

<body>
    <form name="frmIFR" action="CustServAlias.asp" method="POST">
        <input type="hidden" name="hdnNameAliasID" value="">
        <input type="hidden" name="hdnLastUpdate" value="">

        <table border="1" cellspacing="0" cellpadding="2" width="100%">
            <thead>
                <th nowrap title="Customer Service name alias">CS Name Alias</th>
            </thead>
            <tbody>
                <%
            
		if strCS_ID <> "" then
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
                <td nowrap onclick="cell_onClick(<%=rsAlias(0)%>, '<%=rsAlias(2)%>')"><%=rsAlias(1)%>&nbsp;</td>
                </tr>
		<%
		rsAlias.MoveNext
		wend
		rsAlias.Close
		end if
        %>
            </tbody>
        </table>

    </form>
</body>
</html>
