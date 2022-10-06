<%@ Language=VBScript %>
<%  
OPTION EXPLICIT
on error resume next
Response.CacheControl="Private"
%>
<!-- #include file="smaConstants.inc" -->
<!-- #include file="smaProcs.inc" -->
<!-- #include file="databaseconnect.asp" -->
<!--        
  ***************************************************************************************************
  * Name:		CorrUsageList.asp
  * Purpose:	This page list the service usage information.
  * Created By:	?
  ***************************************************************************************************

        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       02-Aug-07	ACheung		Provide service usage information and technical questions & answers
                                    - LCODE_ SERVICE_BASE and LCODE_SERVICE_CONCENTRATION are new tables that 
				      contain the service usage info. 
                                    - SO_ELEMENT and SO_DETAIL_ELEMENT are the technical questions & answers
				      related tables.	
. 
  ***************************************************************************************************
-->
<%
Const ASP_NAME = "CorrUsageList_3.asp" 'only need to change this value when changing the filename

dim strCustServID, strServTypeID
dim objRsServiceOrderInfo, StrSql, objRs, objRsSTAtt, objRsSTAvalue

strServTypeID = Request("ServiceTypeID")
'strCustServID = Request("CustomerServiceID")

if isnumeric(strServTypeID)  then
	
	'Response.Write "Service :" & strServTypeID & "<P>"	
				  
	'StrSql =" SELECT b.srvc_type_att_name, " &_
	'			    "c.SRVC_TYPE_ATT_VAL_NAME, " &_
	'			    "a.srvc_type_att_val_usage_id, " &_
	'			    "st.SERVICE_TYPE_DESC, " &_
	'			    "a.update_date_time, " &_
	'	                    "cs.SERVICE_TYPE_ID, " &_
        '                            "d.UPDATE_DATE_TIME as xref_time " &_
	'		" FROM crp.srvc_type_att_val_usage a," &_
	'			   "crp.srvc_type_att b, " &_
	'			   "crp.srvc_type_att_val c, " &_
	'			   "crp.srvc_type_att_val_xref d, " &_
	'			   "crp.customer_service cs, " &_
	'			   "crp.service_type st " &_
	'		" WHERE a.srvc_type_att_id = b.srvc_type_att_id  AND " &_
	'			   "a.srvc_type_att_val_id = c.srvc_type_att_val_id " &_
	'		" AND a.srvc_type_att_val_usage_id = ( " &_
	'			   " SELECT srvc_type_att_val_usage_id FROM crp.srvc_type_att_val_xref d" &_
	'			   " WHERE service_type_id = ( " &_ 
	'			   " SELECT service_type_id FROM crp.customer_service cs " &_ 
	'			   " WHERE cs.CUSTOMER_SERVICE_ID = " & strCustServID &_ 
	'			   "))" &_
	'		" AND a.SRVC_TYPE_ATT_VAL_USAGE_ID = d.SRVC_TYPE_ATT_VAL_USAGE_ID " &_
	'		" AND cs.SERVICE_TYPE_ID = st.SERVICE_TYPE_ID " &_
	'		" AND d.SERVICE_TYPE_ID = cs.SERVICE_TYPE_ID " &_
	'		" AND cs.CUSTOMER_SERVICE_ID = " & strCustServID
	
 	StrSql =" SELECT b.srvc_type_att_name, " &_
				    "b.srvc_type_att_id, " &_
				    "c.SRVC_TYPE_ATT_VAL_NAME, " &_
				    "c.SRVC_TYPE_ATT_VAL_ID, " &_
				    "a.srvc_type_att_val_usage_id, " &_
				    "d.UPDATE_DATE_TIME, " &_
				    "d.srvc_type_att_val_xref_id " &_
			" FROM crp.srvc_type_att_val_usage a," &_
				   "crp.srvc_type_att b, " &_
				   "crp.srvc_type_att_val c, " &_
				   "crp.srvc_type_att_val_xref d " &_
			" WHERE a.srvc_type_att_id = b.srvc_type_att_id  AND " &_
				   "a.srvc_type_att_val_id = c.srvc_type_att_val_id " &_
			" AND  a.srvc_type_att_val_usage_id = d.srvc_type_att_val_usage_id " &_
			" AND  d.service_type_id = " & strServTypeID 



 	'Response.Write(StrSql)
	'Response.End 

	set objRs = objConn.Execute(StrSql)
	if  objRs.EOF then
		strServTypeID = ""
	end if
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET 32132", err.Description
	end if

end if

	

%>

<HTML>
<HEAD>
<LINK href="stylesheets/styles.css" rel="stylesheet" type="text/css">
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

function cell_onClick(dtUpdate, intServID, intRefID){
	
	document.frmIFR.hdnRefID.value = intRefID;
	document.frmIFR.hdnServID.value = intServID;
	document.frmIFR.hdnUpdateDateTime.value = dtUpdate; 
		
	//highlight current record
	//if (oldHighlightedElement != null) 
	//{
	//	oldHighlightedElement.className = oldHighlightedElementClassName
	//}
	//oldHighlightedElement = window.event.srcElement.parentElement;
	//oldHighlightedElementClassName = oldHighlightedElement.className;
	//oldHighlightedElement.className = "Highlight";
}
</script>

</HEAD>
<BODY>
<form name="frmIFR" action="<%=ASP_NAME%>" method="POST">
<input type="hidden" name="hdnRefID" value="">
<input type="hidden" name="hdnServID" value="">
<input type="hidden" name="hdnUpdateDateTime" value="">
<TABLE border=2 cellspacing=0 cellpadding=2 width="100%">
	<thead>
		<th nowrap title="Service Type Attribute">Service Type Attribute</th>
		<th nowrap title="Service Type Attribute Value">Value</th>
	</thead>
	<tbody>
		<%
		dim k
		k = 0
		if strServTypeID <> "" then
			if Int(k/2) = k/2 then
				Response.Write "<tr class=""regularItem"">"
			else
				Response.Write "<tr class=""whiteItem"">"
			end if
			k = k+1

			while not objRs.EOF
			%> 
				<td nowrap onClick="cell_onClick('<%=objRs("UPDATE_DATE_TIME")%>',<%=strServTypeID%>, <%=objRs("srvc_type_att_val_xref_id")%>);"><%=routineHTMLString(objRs("srvc_type_att_name"))%>&nbsp;</td>
				<td nowrap onClick="cell_onClick('<%=objRs("UPDATE_DATE_TIME")%>',<%=strServTypeID%>, <%=objRs("srvc_type_att_val_xref_id")%>);"><%=routineHTMLString(objRs("SRVC_TYPE_ATT_VAL_NAME"))%>&nbsp;</td>
			</tr>
			<%
				objRs.MoveNext
			wend
			objRs.Close
			set objRs = Nothing
		end if
		%>
	</tbody>
</table>

</FORM>
</BODY>
</HTML>


