<%@ Language=VBScript %>
<%  
OPTION EXPLICIT
on error resume next
Response.CacheControl="Private"
%>
<% Response.Buffer = true %>
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
Const ASP_NAME = "CorrSOWInstList.asp" 'only need to change this value when changing the filename

dim strCustServID, strServTypeID
dim objRsServiceOrderInfo, StrSql, objRs, objRsSTAtt, objRsSTAvalue, StrWhereClasuse, objRs2

strServTypeID = Request("ServiceTypeID")
strCustServID = Request("CustomerServiceID")

if isnumeric(strCustServID)  then
	
	'Response.Write "Service :" & strServTypeID & "<P>"	
	'Response.Write "CSID :" & strCustServID & "<P>"				  

        'StrSql = "SELECT csidSI.CUSTOMER_SERVICE_ID, descSI.SRVC_INSTNC_ATT_NAME, valSI.SRVC_INSTNC_ATT_VAL " &_ 
	'			" FROM SO.CUST_SRVC_INST_ATT_VAL_XREF csidSI,  " &_
	'			" SO.SRVC_INSTNC_ATT_VAL_USAGE valueSI, " &_
	'			" SO.SRVC_INSTNC_ATT_XREF xrefSI, " &_
	'			" SO.SRVC_INSTNC_ATT_VAL valSI, " &_
	'			" SO.SRVC_INSTNC_ATT descSI " &_
	'			" WHERE csidSI.SRVC_INSTNC_ATT_VAL_USAGE_ID = valueSI.SRVC_INSTNC_ATT_VAL_USAGE_ID " &_
	'			" and valueSI.SRVC_INSTNC_ATT_VAL_USAGE_ID = csidSI.SRVC_INSTNC_ATT_VAL_USAGE_ID" &_
	'			" and valueSI.SRVC_INSTNC_ATT_VALUE_ID = valSI.SRVC_INSTNC_ATT_VAL_ID" &_
	'			" and valueSI.SRVC_INSTNC_ATT_XREF_ID  = xrefSI.SRVC_INSTNC_ATT_XREF_ID" &_
	'			" and xrefSI.SRVC_INSTNC_ATT_ID = descSI.SRVC_INSTNC_ATT_ID" &_
	'			" and csidSI.RECORD_STATUS_IND = 'A'" &_
	'			" and csidSI.CUSTOMER_SERVICE_ID = " & strCustServID

        StrSql = "SELECT csidSI.CUSTOMER_SERVICE_ID, CS.CUSTOMER_SERVICE_DESC, ST.SERVICE_TYPE_DESC, CUST.CUSTOMER_NAME, CUST.CUSTOMER_SHORT_NAME, CUST.CUSTOMER_ID, descSI.SRVC_INSTNC_ATT_NAME, " &_
			 	 " decode (csidSI.SRVC_INSTNC_ATT_USR_DEF_VAL, NULL,  valSI.SRVC_INSTNC_ATT_VAL, CSIDSI.SRVC_INSTNC_ATT_USR_DEF_VAL) AS SRVC_INSTNC_ATT_VAL " &_
 				" FROM SO.CUST_SRVC_INST_ATT_VAL_XREF csidSI, " &_
 				" SO.SRVC_INSTNC_ATT_VAL_USAGE valueSI, " &_
 				" SO.SRVC_INSTNC_ATT_XREF xrefSI, " &_
 				" SO.SRVC_INSTNC_ATT_VAL valSI, " &_
 				" SO.SRVC_INSTNC_ATT descSI, " &_
 	 			" CRP.CUSTOMER_SERVICE cs, " &_
				" CRP.service_type st, " &_
				" CRP.CUSTOMER cust " &_
				" WHERE csidSI.SRVC_INSTNC_ATT_VAL_USAGE_ID = valueSI.SRVC_INSTNC_ATT_VAL_USAGE_ID" &_
				" and valueSI.SRVC_INSTNC_ATT_VAL_USAGE_ID = csidSI.SRVC_INSTNC_ATT_VAL_USAGE_ID" &_
				" and valueSI.SRVC_INSTNC_ATT_VALUE_ID = valSI.SRVC_INSTNC_ATT_VAL_ID" &_
				" and valueSI.SRVC_INSTNC_ATT_XREF_ID  = xrefSI.SRVC_INSTNC_ATT_XREF_ID" &_
				" and xrefSI.SRVC_INSTNC_ATT_ID = descSI.SRVC_INSTNC_ATT_ID" &_
				" and csidSI.RECORD_STATUS_IND = 'A'" &_
				" and CS.SERVICE_TYPE_ID = ST.SERVICE_TYPE_ID" &_
				" and CS.CUSTOMER_ID = CUST.CUSTOMER_ID" &_
				" and csidSI.CUSTOMER_SERVICE_ID = CS.CUSTOMER_SERVICE_ID" &_
				" and csidSI.CUSTOMER_SERVICE_ID = " & strCustServID &_
				" order by XREFSI.DISPLAY_ORDER"
					
	'Response.Write(StrSql)
	'Response.End 

	Dim aList, strExportenabled
	
	set objRs = objConn.Execute(StrSql)

' with empty row display
	if  objRs.EOF then
		strCustServID = ""
	else
		aList = objRs.GetRows
	end if
	
	if aList(0,0) = "" then 
		strExportenabled = "return false;" 'to disable the excel export icon for archor 
	else	
		strExportenabled = ""
	end if

'Response.write "<p>strExportenabled = " & strExportenabled & "</p>"
	

' without empty row display
''''	if  not objRs.EOF then
''''		aList = objRs.GetRows
''		for k = 0 to UBound(aList, 2)
''			Response.Write "element 0= " & aList(0,k) & vbCrLf	
''			Response.Write "element 1= " & aList(1,k) & vbCrLf
''			Response.Write "element 2= " & aList(2,k) & vbCrLf
''			Response.Write "element 3= " & aList(3,k) & vbCrLf
''		next
''		Response.Write "upper bound = " & UBound(aList, 2)
''		Response.end	
''''	else
''''		strCustServID = ""
''''		Response.Write "0 records found"
''''		Response.end
''''	end if

	objRs.Close
	set objRs = nothing
	
	objConn.close
	set objConn = nothing

	select case Request("Action")
		case "" if objRs <> "" then set objRs = nothing
		case else if Request("hdnExport") <> "" then
				'get real userid
				dim strRealUserID
				strRealUserID = Request.Cookies("UserInformation")("username")
				'determine export path
				dim strExportPath, liLength
				strExportPath = Request.ServerVariables("PATH_TRANSLATED")
				While (Right(strExportPath, 1) <> "\" And Len(strExportPath) <> 0)
					liLength = Len(strExportPath) - 1
					strExportPath = Left(strExportPath, liLength)
				Wend
				strExportPath = strExportPath & "export\"
		
				'create scripting object
				dim objFSO, objTxtStream
				set objFSO = server.CreateObject("Scripting.FileSystemObject")
				'create export file (overwrite if exists)
				set objTxtStream = objFSO.CreateTextFile(strExportPath&strRealUserID&"-SIA.xls", true, false)
		
				if err then
						DisplayError "CLOSE", "", err.Number, "CorrSOWInstList.asp - Cannot create Excel spreadsheet file due to the following reasons. Please contact your website administrator.", err.Description
				end if
		
				with objTxtStream
					.WriteLine "<table border=2>"
		
				'export the header
					.WriteLine "<TR>"
					.WriteLine "<TH>CSID</TD>"
					.WriteLine "<TH>CS Name</TD>"
					.WriteLine "<TH>Service Type Description</TD>"
					.WriteLine "<TH>Customer Name</TD>"
					.WriteLine "<TH>Customer Short Name</TD>"
					.WriteLine "<TH>Customer ID</TD>"
					.WriteLine "<TH>Service Instance Attribute</TD>"
					.WriteLine "<TH>Value</TD>"
					.WriteLine "<TD NOWRAP>&nbsp;&nbsp;&nbsp;</TD>"
					.WriteLine "</TR>"
									
					'export the body
					for k = 0 to UBound(aList, 2)
						.WriteLine "<TR>"
						.WriteLine "<TD NOWRAP colspan=1>"&routineHtmlString(aList(0,k))&"&nbsp;</TD>"
						.WriteLine "<TD NOWRAP colspan=1>"&routineHtmlString(aList(1,k))&"&nbsp;</TD>"
						.WriteLine "<TD NOWRAP colspan=1>"&routineHtmlString(aList(2,k))&"&nbsp;</TD>"
						.WriteLine "<TD NOWRAP colspan=1>"&routineHtmlString(aList(3,k))&"&nbsp;</TD>"
						.WriteLine "<TD NOWRAP colspan=1>"&routineHtmlString(aList(4,k))&"&nbsp;</TD>"
						.WriteLine "<TD NOWRAP colspan=1>"&routineHtmlString(aList(5,k))&"&nbsp;</TD>"
						.WriteLine "<TD NOWRAP colspan=1>"&routineHtmlString(aList(6,k))&"&nbsp;</TD>"
						.WriteLine "<TD NOWRAP colspan=1>"&routineHtmlString(aList(7,k))&"&nbsp;</TD>"
						.WriteLine "<TD NOWRAP>&nbsp;&nbsp;&nbsp;</TD>"
						.WriteLine "</TR>"
					next
						.WriteLine "</table>"
				end with
				objTxtStream.Close
				StrSql = "<script type=""text/javascript"">document.location=""export/"&strRealUserID&"-SIA.xls"";</script>"
				Response.Write StrSql
				Response.End
				'Response.redirect "export/"&strRealUserID&"-SIA.xls"
			end if
	end select


	if response.isclientconnected = false then
		Response.End
	end if

	'if err then
	'	DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET 32132", err.Description
	'end if

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
<form method=post name=frmIFR action="<%=ASP_NAME%>">
<input type="hidden" name="hdnRefID" value="">
<input type="hidden" name="hdnServID" value="">
<input type="hidden" name="hdnUpdateDateTime" value="">
<input name="hdnExport" type=hidden value>

<TABLE border=2 cellspacing=0 cellpadding=2 width="100%">
	<thead>
		<th nowrap title="Service Instance Attribute">Service Instance Attribute</th>
		<th nowrap title="Service Instance Attribute Value">Value</th>
	</thead>
	<tbody>
		<%
		dim k
		if strServTypeID <> "" then
			if Int(k/2) = k/2 then
				Response.Write "<tr class=""regularItem"">"
			else
				Response.Write "<tr class=""whiteItem"">"
			end if
			k = k+1

			for k = 0 to UBound(aList,2)
			%>
			<tr> 
				<td nowrap onClick="cell_onClick('<%=hdnCustServID%>',<%=strServTypeID%>);"><%=routineHTMLString(aList(6,k))%>&nbsp;</td>
				<td nowrap onClick="cell_onClick('<%=hdnCustServID%>',<%=strServTypeID%>);"><%=routineHTMLString(aList(7,k))%>&nbsp;</td>
			</tr>
			<%
			objRs.MoveNext
			next
			'objRs.Close
			'set objRs = Nothing
		end if
		%>
		</tbody>
		<TFOOT>
		<TR>
			<TD align=left colSpan=2>
				<a  onclick="<%=strExportenabled%>" href ="CorrSOWInstList.asp?ServiceTypeID=<%=strServTypeID%>&CustomerServiceID=<%=strCustServID%>&hdnExport=xls&Action=1" target="_blank" ><img SRC="images/excel.gif"/></a>
			</TD>
		</TR>		
		</TFOOT>
</table>
</FORM>
</BODY>
</HTML>


