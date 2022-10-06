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
  * Name:		CorrSOInstList.asp
  * Purpose:	This page list the service usage information.
  * Created By:	?
  ***************************************************************************************************

        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       02-Aug-07	ACheung		Provide service usage information and technical questions & answers
       updated:                     - SRVC_TYPE_ATT and SRVC_TYPE_ATT_VAL are new tables that 
       03-Oct-08		      contain the serice type attribute and its values (or service usage info.) 
                                    - SO.SRVC_INSTNC_ATT and SO.SRVC_INSTNC_ATT_VAL are the 
				      service instance attribute and its values (or technical questions & answers
				      related) tables.	
. 
  ***************************************************************************************************
-->
<%
Const ASP_NAME = "CorrSOInstList.asp" 'only need to change this value when changing the filename

dim strCustServID, strServTypeID
dim objRsServiceOrderInfo, StrSql, objRs, objRsSTAtt, objRsSTAvalue, StrSql2, objRs2

strServTypeID = Request("ServiceTypeID")
strCustServID = Request("CustomerServiceID")

if isnumeric(strCustServID)  then
	
	'Response.Write "Service :" & strServTypeID & "<P>"	
	'Response.Write "CSID :" & strCustServID & "<P>"				  

        StrSql = "SELECT serv_inst.SRVC_INSTNC_ATT_NAME, " &_
                 "decode(DET_INST_XREF.SRVC_INSTNC_ATT_USR_DEF_VAL, NULL,  att_val.SRVC_INSTNC_ATT_VAL, DET_INST_XREF.SRVC_INSTNC_ATT_USR_DEF_VAL) "&_
                 "AS SRVC_INSTNC_ATT_VAL " &_       
                        " FROM SO.SO_DETAIL_INSTNC_ATT_VAL_XREF det_inst_xref, " &_
                                " SO.SRVC_INSTNC_ATT_VAL_USAGE inst_val_xref, " &_
                                " SO.SRVC_INSTNC_ATT_XREF serv_inst_xref, " &_
                                " SO.SRVC_INSTNC_ATT serv_inst, " &_
                                " SO.SRVC_INSTNC_ATT_VAL att_val, " &_
                                " SO.SO_DETAIL det " &_
                    " WHERE " &_ 
                                 " det.SO_DETAIL_ID = det_inst_xref.SO_DETAIL_ID " &_                        
                                 " AND det_inst_xref.SRVC_INSTNC_ATT_VAL_USAGE_ID = inst_val_xref.SRVC_INSTNC_ATT_VAL_USAGE_ID " &_ 
                                 " AND inst_val_xref.SRVC_INSTNC_ATT_XREF_ID = serv_inst_xref.SRVC_INSTNC_ATT_XREF_ID " &_ 
                                 " AND serv_inst_xref.SERVICE_TYPE_ID = det.SERVICE_TYPE_ID " &_ 
                                 " AND  serv_inst_xref.SRVC_INSTNC_ATT_ID = serv_inst.SRVC_INSTNC_ATT_ID " &_ 
                                 " AND att_val.SRVC_INSTNC_ATT_VAL_ID =  inst_val_xref.SRVC_INSTNC_ATT_VALUE_ID " &_ 
                                 " AND det_inst_xref.RECORD_STATUS_IND = 'A' " &_ 
                                 " AND det.SO_DETAIL_ID =  " &_ 
                                          "( SELECT MAX(dtl.SO_DETAIL_ID) " &_
                                		  " FROM SO.SO_DETAIL dtl, SO.V_ECOPS_ORDER ec " &_
                                		  "	WHERE dtl.SO_DETAIL_ID = ec.SO_DETAIL_ID " &_
                                		  " AND dtl.CUSTOMER_SERVICE_ID = " & strCustServID &_
                                		  " AND ec.ORDER_COMPLETE_DATE IS NULL " &_
						  " AND ec.ORDER_STATUS <> 'Cancelled' ) " &_
						  " order by SERV_INST_XREF.DISPLAY_ORDER "
                                

' remove this from line 69    " AND ec.ORDER_COMPLETE_DATE IS NOT NULL " 
	'Response.Write(StrSql)
	'Response.End 

	set objRs = objConn.Execute(StrSql)
	if  objRs.EOF then
		strCustServID = ""
	end if
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET 32132", err.Description
	end if


'	Set cmdObjSP_GET_LATEST_SERV_INST = Nothing

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

function cell_onClick(dtUpdate, intServID ){
	
	document.frmIFR.hdnRefID.value = dtUpdate;
	document.frmIFR.hdnServID.value = intServID;
	//document.frmIFR.hdnUpdateDateTime.value = dtUpdate; 
		
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
		<th nowrap title="Service Instance Attribute">Service Instance Attribute</th>
		<th nowrap title="Service Instance Attribute Value">Value</th>
	</thead>
	<tbody>
		<%
		dim k
		k = 0
		if strServTypeID <> "" and strServTypeID <> 0 then    'LC add strServTypeID<>0 on March 2015 
			if Int(k/2) = k/2 then
				Response.Write "<tr class=""regularItem"">"
			else
				Response.Write "<tr class=""whiteItem"">"
			end if
			k = k+1

			while not objRs.EOF
			%> 
				<td nowrap onClick="cell_onClick('<%=objRs("SRVC_INSTNC_ATT_NAME")%>',<%=strServTypeID%>);"><%=routineHTMLString(objRs("srvc_instnc_att_name"))%>&nbsp;</td>
				<td nowrap onClick="cell_onClick('<%=objRs("SRVC_instnc_att_val")%>',<%=strServTypeID%>);"><%=routineHTMLString(objRs("SRVC_instnc_att_val"))%>&nbsp;</td>
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


