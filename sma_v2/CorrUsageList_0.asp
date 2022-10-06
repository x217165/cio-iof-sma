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
  * Name:		CorrUsage.asp
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
Const ASP_NAME = "CorrUsageList.asp" 'only need to change this value when changing the filename

dim strCustServID, objRsServiceContact, strServTypeID, StrSql
Dim intAccessLevel, bolActiveOnly

strCustServID = Request("CustomerServiceID")
strServTypeID = Request("ServiceTypeID")

<!--if strCustServID <> "" then
  strSql =  "select sd.SO_DETAIL_ID," &_ 
			"     se.SO_ELEMENT_DESCRIPTION," &_
			"     sde.DETAIL_ELEMENT_TEXT," &_
			"     SE.ELEMENT_TYPE_LCODE," &_
			"     sde.SO_ELEMENT_ID," &_
			"     se.SO_ELEMENT_SMA_DISPLAY," &_
			"     st.SERVICE_TYPE_ID," &_
			"     sc.SERVICE_CONCENTRATION_NAME," &_
			"     sb.SERVICE_BASE_NAME" &_
			"from so.so_detail sd," &_
			"     so.so_detail_element sde," &_
	                "     so.so_element se," &_
                        "     crp.service_type st," &_
                        "     crp.service_con_base_xref scbx," &_
                        "     crp.lcode_service_concentration sc," &_
			"     crp.lcode_service_base sb" &_
			"where sd.SO_DETAIL_ID = sde.SO_DETAIL_ID" &_
			"and   sde.SO_ELEMENT_ID = se.SO_ELEMENT_ID" &_
			"and   st.service_type_id = scbx.service_type_id" &_
			"and   scbx.SERVICE_CONCENTRATION_LCODE = sc.SERVICE_CONCENTRATION_LCODE" &_
			"and   scbx.SERVICE_BASE_LCODE = sb.SERVICE_BASE_LCODE" &_
			"and   se.SO_ELEMENT_SMA_DISPLAY = 'Y'" &_
			"and   st.service_type_id = " &strServTypeID
			"and   sd.CUSTOMER_SERVICE_ID = " &strCustServID
	
	set objRsServiceOrderInfo = objConn.Execute(StrSql)
	if  objRsServiceOrderInfo.EOF then
		strCustServID = ""
	end if

	if err then
		DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET 32132", err.Description
	end if

end if
-->

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

</script>

</HEAD>
<BODY>
<form name="frmaifr2" action="<%=ASP_NAME%>" method="POST">
<input type="hidden" name="hdnServTypeID" value="">
<input type="hidden" name="hdnCustServID" value="">
<input type="hidden" name="hdnUpdateDateTime" value="">

<TABLE border=1 cellspacing=0 cellpadding=2 width="100%">
	<thead>
		<th nowrap>Service Concentration</th>
		<th nowrap>Service Base</th>
		<th nowrap>Technical Question</th>
		<th nowrap>Technical Answer</th>
	</thead>
	<tbody>
		<%
		dim k
		k = 0
		if strCustServID <> "" then
			while not objRsServiceOrderInfo.EOF 
				if Int(k/2) = k/2 then
					Response.Write "<tr class=""regularItem"">"
				else
					Response.Write "<tr class=""whiteItem"">"
				end if
				k = k+1
					
			%>
					Response.Write "<TD NOWRAP >"&routineHtmlString(objRsServiceOrderInfo("SERVICE_CONCENTRATION_NAME"))&"&nbsp;</TD>"&vbCrLf
					Response.Write "<TD NOWRAP >"&routineHtmlString(objRsServiceOrderInfo("SERVICE_BASE_NAME"))&"&nbsp;</TD>"&vbCrLf
					Response.Write "<TD NOWRAP >"&routineHtmlString(objRsServiceOrderInfo("SO_ELEMENT_DESCRIPTION"))&"&nbsp;</TD>"&vbCrLf
					Response.Write "<TD NOWRAP >"&routineHtmlString(objRsServiceOrderInfo("DETAIL_ELEMENT_TEXT"))&"&nbsp;</TD>"&vbCrLf
				</tr>
			<%
				objRsServiceOrderInfo.MoveNext
			wend
			objRsServiceOrderInfo.Close
			set objRsServiceOrderInfo = Nothing
		end if
		%>
	</tbody>
</table>
</FORM>
</BODY>
</HTML>


