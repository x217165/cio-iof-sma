<%@ Language=VBScript %>
<% option explicit%>
<% Response.Buffer = true %>
<% 'on error resume next %>
<!--#include file="SmaConstants.inc"-->
<!--#include file="SmaProcs.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--
  ***************************************************************************************************
  * Name:		CustServList.asp i.e. Customer Service List**
  * Purpose:	This page reads users's search critiera and bring back a list of matching Customer
  *					Service location records.
  * Created By:	Sara Sangha 08/01/00
  ***************************************************************************************************

		 Date		Author			Changes/enhancements made
         -----		------		------------------------------------------------------
		07-Feb-01	 DTy		Added 'DISTINCT' to variable strSQL to prevent service
								location records from appearing more than once when the
								search name appears both in the CUSTOMER and
								CUSTOMER_ALIAS tables.
		06-Mar-01	 DTy		Save 'ActiveOnly' cookie for use by ServLocContact.asp.
        	20-Jul-01	 DTy		When 'Active Only' is selected:
								  Exclude customers that are marked as soft deleted
								  Exclude addresses that are marked as soft deleted
       		18-Feb-02	 DTy		Active customers are those whose status is either
                		                'Prospect', 'OnHold' or 'Current'.
		07-May-08   	ACheung 	Add CLLI Code (Geocode)
********************************************************************************************
-->

<html>
<head>
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>SMA - Service Location List</title>
</head>
<link rel="stylesheet" type="text/css" href="stylesheets/styles.css">
<script type="text/javascript">

function go_back(strServiceEnd,lngServLocID, strServLocName, strCilliCode, strProvinceCode, strAddress, strStreet, strBuilding){
//*******************************************************************************************
//
//
//
//
//
//******************************************************************************************

        if (strServiceEnd == 'A'){
			parent.opener.document.forms[0].hdnServiceLocIdA .value = lngServLocID ;
	        parent.opener.document.forms[0].txtsrvloca.value = strServLocName ;
	        parent.opener.document.forms[0].txtaaddressa.value = strAddress ;

		  }
		 else if (strServiceEnd == 'B'){
		    parent.opener.document.forms[0].hdnServiceLocIdB.value = lngServLocID ;
	        parent.opener.document.forms[0].txtsrvlocb.value = strServLocName ;
	        parent.opener.document.forms[0].txtaaddressb.value = strAddress ;
		   }

		 else if (strServiceEnd == 'Z') {
			parent.opener.document.forms[0].hdnServLocID.value = lngServLocID;
	        parent.opener.document.forms[0].txtServLocName.value = strServLocName;
	        parent.opener.document.forms[0].txtServLocAddress.value = strAddress;
			try {parent.opener.document.forms[0].hdnClliCode.value = strCilliCode;
				parent.opener.document.forms[0].hdnProvinceCode.value = strProvinceCode;
				parent.opener.document.forms[0].hdnBuildingName.value = strBuilding;
				parent.opener.document.forms[0].hdnStreetName.value = strStreet;
				}
			catch(e){}//do nothing, the caller don't need that info
		   }
		 else {
		    parent.opener.document.forms[0].hdnServLocID.value = lngServLocID ;
	        parent.opener.document.forms[0].txtServLocName.value = strServLocName ;
	        parent.opener.document.forms[0].txtServLocAddress.value = strAddress ;
		 }

	 parent.window.close ();
}


function window_onload() {
//******************************************************************************************
//
//
//
//
//******************************************************************************************
	// clear the value so that new search will not consider it
	//top.text.fraCriteria.frmServLocCriteria.hdnAddressID.value = "" ;


}

function btnExcel_onClick()
{
	document.forms[0].target='new';
	document.forms[0].hdnExport.value='xls';
	document.forms[0].submit();
	document.forms[0].hdnExport.value='';
	document.forms[0].target='_self';

}
//*****************************************End of Java Functions ****************************
</script>
<body LANGUAGE="javascript" onload="return window_onload()">

<%

 dim aList, intPageNumber, intPageCount
 dim strCustomerName, strStreetName, strServiceLocationName, strAddressID
 dim strCity, strSpecificLocationDesc, strProvince, bolActiveOnly

 dim strSQL, strWhereClause, strRecordStatus, strOrderBy,strServiceEnd
 dim strServLocID, strWinName

	strServLocID = Request("ServLocID")
	strWinName = Request("hdnWinName")&Request("WinName")
	strCustomerName = UCase(trim(Request.Form("txtCustomerName")))
	'strAddressID = trim(Request.Form("hdnAddressID"))
	strStreetName = UCase(trim(Request.Form("txtStreetName")))
	strServiceLocationName = UCase(trim(Request.Form("txtServiceLocationName")))
	strCity = UCase(trim(Request.Form("txtCity")))
	strSpecificLocationDesc = UCase(trim(Request.Form("txtSpecificLocationDesc")))
	strProvince = UCase(trim(Request.Form("selProvince")))
	bolActiveOnly = UCase(trim(Request.Form("chkActiveOnly")))
	strServiceEnd = Request("hdnServiceEnd")

	if strServiceEnd = "" then
	 strServiceEnd = "OTHER"
	END IF

	if len(strCustomerName) = 0 then

	  strSQL = "select l.service_location_id, " &_
					 "c.customer_name, " &_
					 "l.service_location_name, " &_
					 "a.street, " &_
					 "a.municipality_name, " &_
					 "a.province_state_lcode, " &_
					 "m.clli_code, " &_
					 "NVL(A.BUILDING_NAME,'NO BUILDING NAME') ||" &_
					 "CHR(13)||CHR(10)||NVL(A.STREET,'NO STREET ADDRESS')||" &_
					 "CHR(13)||CHR(10)||NVL(A.MUNICIPALITY_NAME,'NO MUNICIPALITY')||' '||" &_
                                         "NVL(A.PROVINCE_STATE_LCODE,'NO PROVINCE')||" &_
                                         "CHR(13)||CHR(10)||NVL(A.POSTAL_CODE_ZIP,'NO POSTAL CODE') ADDRESS, " &_
                                         "A.BUILDING_NAME, " &_
                                         "g.clli_code AS full_clli_code " &_
			 "from crp.service_location l, " &_
					"crp.customer c, " &_
					"crp.v_address_consolidated_street a,  " &_
                                        "crp.service_location_geocode slg,  " &_
                                        "crp.lcode_geocodeid g,  " &_
					"crp.municipality_lookup m "

	  strWhereClause = "where c.customer_id = l.customer_id " &_
					 "and   l.address_id = a.address_id " &_
					 "and   a.municipality_name = m.municipality_name " &_
					 "and a.province_state_lcode = m.province_state_lcode " &_
					 "and a.country_lcode = m.country_lcode " &_
   					 "and l.service_location_id = slg.service_location_id(+) " &_
					 "and slg.geocodeid_lcode = g.geocodeid_lcode(+) "
	else

	  strSQL = "select distinct l.service_location_id, " &_
					 "c.customer_name, " &_
					 "l.service_location_name, " &_
					 "a.street, " &_
					 "a.municipality_name, " &_
					 "a.province_state_lcode, " &_
					 "m.clli_code, " &_
					 "NVL(A.BUILDING_NAME,'NO BUILDING NAME') ||" &_
					 "CHR(13)||CHR(10)||NVL(A.STREET,'NO STREET ADDRESS')||" &_
					 "CHR(13)||CHR(10)||NVL(A.MUNICIPALITY_NAME,'NO MUNICIPALITY')||' '||" &_
                                         "NVL(A.PROVINCE_STATE_LCODE,'NO PROVINCE')||" &_
                                         "CHR(13)||CHR(10)||NVL(A.POSTAL_CODE_ZIP,'NO POSTAL CODE') ADDRESS, " &_
                                         "A.BUILDING_NAME, " &_
                                         "g.clli_code " &_
			 "from crp.service_location l, " &_
					"crp.customer c, " &_
					"crp.customer_name_alias c1, " &_
					"crp.v_address_consolidated_street a, "  &_
                                        "crp.service_location_geocode slg,  " &_
                                        "crp.lcode_geocodeid g,  " &_
					"crp.municipality_lookup m "
	  strWhereClause = "where c.customer_id = l.customer_id " &_
	   				 "and	c.customer_id = c1.customer_id " &_
	   				 "and   l.address_id = a.address_id " &_
	   				 "and   a.municipality_name = m.municipality_name " &_
					 "and a.province_state_lcode = m.province_state_lcode " &_
					 "and a.country_lcode = m.country_lcode " &_
   					 "and l.service_location_id = slg.service_location_id(+) " &_
					 "and slg.geocodeid_lcode = g.geocodeid_lcode(+) "

		strWhereClause = strWhereClause & "and   (c1.customer_name_alias_upper like '" & routineOraString(strCustomerName) &  "%' "
		if Request("chkIncludeTelus") = "YES" then
			strWhereClause = strWhereClause & " OR c1.customer_name_alias_upper LIKE '" & UCase(strTelus) & "%' "
		end if
		strWhereClause = strWhereClause & ") "

	end if

	if len(strStreetName) > 0 then
		strWhereClause = strWhereClause & " and upper(a.street) like '" & routineOraString(strStreetName) &  "%' "
	end if

	if len(strServiceLocationName) > 0 then
		strWhereClause = strWhereClause & " and upper(l.service_location_name) like '" & routineOraString(strServiceLocationName) &  "%' "
	end if

	if len(strCity) > 0 then
		strWhereClause = strWhereClause & " and upper(a.municipality_name) like '" & routineOraString(strCity) &  "%' "
	end if

	if len(strSpecificLocationDesc) > 0 then
		strWhereClause = strWhereClause & " and upper(l.specific_location_desc) like '" & routineOraString(strSpecificLocationDesc) &  "%' "
	end if

	if len(strProvince) > 0 then
		strWhereClause = strWhereClause & " and upper(a.province_state_lcode) = '" & routineOraString(strProvince) &  "' "
	end if

    Response.Cookies ("ActiveOnly")=bolActiveOnly
	if bolActiveOnly = "YES" THEN
		strRecordStatus = " and c.customer_status_lcode IN ('Prospect', 'Current', 'OnHold')" &_
                          " and l.record_status_ind (+) = 'A' and a.record_status_ind (+) = 'A' and c.record_status_ind = 'A'"
	   if len(strCustomerName) <> 0 then
		  strRecordStatus = strRecordStatus + " and c1.record_status_ind = 'A'"
	   end if
	else
		strRecordStatus = " "
	end if

	strOrderBy = " order by UPPER(c.customer_name), l.service_location_name"
	strSQL = strSQL & strWhereClause & strRecordStatus & strOrderBy

	Dim objRs,Recordcnt,strbgcolor

	set objRS = objConn.Execute(StrSql)
	if not objRS.EOF then
		aList = objRS.GetRows
	else
		Response.Write "0 records found"
		Response.end
	end if

   'release and kill the recordset and the connection objects
	objRS.Close
	set objRS = nothing

	objConn.close
	set objConn = nothing

   'calculate page number
	intPageCount = Int(UBound(aList, 2) / intConstDisplayPageSize) + 1
	select case Request("Action")
		case "<<"		intPageNumber = 1
		case "<"		intPageNumber = Request("txtPageNumber") - 1
					if intPageNumber < 1 then intPageNumber = 1
		case ">"		intPageNumber = Request("txtPageNumber") + 1
					if intPageNumber > intPageCount then intPageNumber = intPageCount
		case ">>"		intPageNumber = intPageCount
		'Case "Export"
		case else
				if Request("hdnExport") <> "" then
					Dim strRealUserID
					Dim strExportPath
					Dim liLength
					Dim objFSO
					Dim objTxtStream
					strRealUserID = Session("username")
					'determine export path
					strExportPath = Request.ServerVariables("PATH_TRANSLATED")
					Do While (Right(strExportPath, 1) <> "\" And Len(strExportPath) <> 0)
						liLength = Len(strExportPath) - 1
						strExportPath = Left(strExportPath, liLength)
					Loop
					strExportPath = strExportPath & "export\"

					'create scripting object
					Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
					'create export file (overwrite if exists)
					Set objTxtStream = objFSO.CreateTextFile(strExportPath & strRealUserID & "-servloc.xls", True, False)
					if err then
						DisplayError "CLOSE", "", err.Number, ASP_NAME & " - Cannot create Excel spreadsheet file due to the following reasons.  Please contact your website administrator.", err.Description
					end if

					With objTxtStream
						.WriteLine "<TABLE border=1>"
						.WriteLine "<THEAD>"
						.WriteLine "<TH>Customer Name</TH>"
						.WriteLine "<TH>Service Location</TH>"
						.WriteLine "<TH>Street Name</TH>"
						.WriteLine "<TH>Municipality</TH>"
						.WriteLine "<TH>Prov/State</TH>"
						.WriteLine "<TH>CLLI Code</TH>"
						.WriteLine "</THEAD>"

						'export the body
						For k = 0 To UBound(aList, 2)
							.WriteLine "<TR>"
							.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(1, k)) & "&nbsp;</TD>"
							.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(2, k)) & "&nbsp;</TD>"
							.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(3, k)) & "&nbsp;</TD>"
							.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(4, k)) & "&nbsp;</TD>"
							.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(5, k)) & "&nbsp;</TD>"
							.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(9, k)) & "&nbsp;</TD>"
							.WriteLine "</TR>"
						Next
						.WriteLine "</TABLE>"
					End With
					objTxtStream.Close
					Set objTxtStream = Nothing
					Set objFSO = Nothing
					'Response.Write "<SCRIPT type='text/javascript' language='javascript'>"
					'Response.Write "window.open('" & "export/" & strRealUserID & ".xls" & "');"
					'Response.Write "</SCRIPT>"
						strsql = "<script type=""text/javascript"">document.location=""export/"&strRealUserID&"-servloc.xls"";</script>"
						Response.Write strsql
						Response.End
					Response.redirect "export/"&strRealUserID&".xls"
		'case else
				elseif Request("txtGoToPageNo") <> "" then
					intPageNumber = CInt(Request("txtGoToPageNo"))
				else
					intPageNumber = 1
				end if
	end select

	if intPageNumber < 1 then
		intPageNumber = 1
	end if
	if intPageNumber > intPageCount then
		intPageNumber = intPageCount
	end if

	dim k, m, n
	m = (intPageNumber - 1 ) * intConstDisplayPageSize
	n = (intPageNumber) * intConstDisplayPageSize - 1
	if n > UBound(aList, 2) then
		n = UBound(aList, 2)
	end if

	'check if the client is still connected just before sending any html to the browser
	if response.isclientconnected = false then
		Response.End
	end if

	'catch any unexpected error
	if err then
		DisplayError "BACK", "", err.Number, "Unexpected error", err.Description
	end if
%>
<form method="post" name="frmAddressList" action="ServLocList.asp">

	<input type="hidden" name="hdnWinName" value="<%=strWinName%>">
	<input type="hidden" name="txtCustomerName" value="<%=routineHTMLString(strCustomerName)%>">
    <input type="hidden" name="txtStreetName" value="<%=routineHTMLString(strStreetName)%>">
    <input type="hidden" name="txtServiceLocationName" value="<%=routineHTMLString(strServiceLocationName)%>">
    <input type="hidden" name="txtCity" value="<%=routineHTMLString(strCity)%>">
    <input type="hidden" name="txtSpecificLocation" value="<%=routineHTMLString(strSpecificLocationDesc)%>">
    <input type="hidden" name="selProvince" value="<%=routineHTMLString(strProvince)%>">
    <input type="hidden" name="chkActiveOnly" value="<%=bolActiveOnly%>">
    <input type="hidden" name="hdnServiceEnd" value="<%=routineHTMLString(strServiceEnd)%>">
    <input type="hidden" name="chkIncludeTelus" value="<%=routineHTMLString(Request("chkIncludeTelus"))%>">

    <input type="hidden" name="hdnExport" value>

<table border="1" cellPadding="2" cellSpacing="0" width="100%">
    <thead>
    <tr>
        <th align="left">Customer Name</th>
        <th align="left">Service Location</th>
        <th align="left">Street Name</th>
        <th align="left">Municipality</th>
        <th align="left">Prov/State</th>
        <th align="left">CLLI Code</th>
        </tr>
    </thead>
    <tbody>

<%
	for k = m to n
	'Alternate row background colour
	if Int(k/2) = k/2 then
		Response.write "<TR>"
	else
		Response.write "<TR bgcolor=White>"
	end if


	if strWinName = "Popup" then

	Response.Write "<td NOWRAP><a href=""#"" onClick=""return go_back('" &routineJavaScriptString(strServiceEnd)& "','"&routineJavaScriptString(aList(0,k))& "', '" &routineJavascriptString(aList(2,k))&  "', '" &routineJavascriptString(aList(6,k))& "', '" &routineJavascriptString(aList(5,k))& "', '" &routineJavascriptString(aList(7,k))&  "', '" &routineJavascriptString(aList(3,k))&  "', '" &routineJavascriptString(aList(8,k))&  "')"">" &routineHTMLString(aList(1,k))& "</a></td>"&vbCrLf
	Response.Write "<td NOWRAP><a href=""#"" onClick=""return go_back('" &routineJavaScriptString(strServiceEnd)& "','"&routineJavaScriptString(aList(0,k))& "', '" &routineJavascriptString(aList(2,k))&  "', '" &routineJavascriptString(aList(6,k))& "', '" &routineJavascriptString(aList(5,k))& "', '" &routineJavascriptString(aList(7,k))&  "', '" &routineJavascriptString(aList(3,k))&  "', '" &routineJavascriptString(aList(8,k))&  "')"">" &routineHTMLString(aList(2,k))& "</a></td>"&vbCrLf
	Response.Write "<td NOWRAP><a href=""#"" onClick=""return go_back('" &routineJavaScriptString(strServiceEnd)& "','"&routineJavaScriptString(aList(0,k))& "', '" &routineJavascriptString(aList(2,k))&  "', '" &routineJavascriptString(aList(6,k))& "', '" &routineJavascriptString(aList(5,k))& "', '" &routineJavascriptString(aList(7,k))&  "', '" &routineJavascriptString(aList(3,k))&  "', '" &routineJavascriptString(aList(8,k))&  "')"">" &routineHTMLString(aList(3,k))& "</a></td>"&vbCrLf
	Response.Write "<td NOWRAP><a href=""#"" onClick=""return go_back('" &routineJavaScriptString(strServiceEnd)& "','"&routineJavaScriptString(aList(0,k))& "', '" &routineJavascriptString(aList(2,k))&  "', '" &routineJavascriptString(aList(6,k))& "', '" &routineJavascriptString(aList(5,k))& "', '" &routineJavascriptString(aList(7,k))&  "', '" &routineJavascriptString(aList(3,k))&  "', '" &routineJavascriptString(aList(8,k))&  "')"">" &routineHTMLString(aList(4,k))& "</a></td>"&vbCrLf
	Response.Write "<td NOWRAP><a href=""#"" onClick=""return go_back('" &routineJavaScriptString(strServiceEnd)& "','"&routineJavaScriptString(aList(0,k))& "', '" &routineJavascriptString(aList(2,k))&  "', '" &routineJavascriptString(aList(6,k))& "', '" &routineJavascriptString(aList(5,k))& "', '" &routineJavascriptString(aList(7,k))&  "', '" &routineJavascriptString(aList(3,k))&  "', '" &routineJavascriptString(aList(8,k))&  "')"">" &routineHTMLString(aList(5,k))& "</a></td>"&vbCrLf
	Response.Write "<td NOWRAP><a href=""#"" onClick=""return go_back('" &routineJavaScriptString(strServiceEnd)& "','"&routineJavaScriptString(aList(0,k))& "', '" &routineJavascriptString(aList(2,k))&  "', '" &routineJavascriptString(aList(6,k))& "', '" &routineJavascriptString(aList(5,k))& "', '" &routineJavascriptString(aList(7,k))&  "', '" &routineJavascriptString(aList(3,k))&  "', '" &routineJavascriptString(aList(8,k))&  "')"">" &routineHTMLString(aList(9,k))& "</a></td>"&vbCrLf
	Response.Write "</TR>"


else

	Response.Write "<TD NOWRAP><a target=""_parent"" href=""ServLocDetail.asp?ServLocID="&aList(0,k)&""">"&routineHtmlString(aList(1,k))&"</a></TD>"&vbCrLf
	Response.Write "<TD NOWRAP><a target=""_parent"" href=""ServLocDetail.asp?ServLocID="&aList(0,k)&""">"&routineHtmlString(aList(2,k))&"</a></TD>"&vbCrLf
	Response.Write "<TD NOWRAP><a target=""_parent"" href=""ServLocDetail.asp?ServLocID="&aList(0,k)&""">"&routineHtmlString(aList(3,k))&"</a></TD>"&vbCrLf
	Response.Write "<TD NOWRAP><a target=""_parent"" href=""ServLocDetail.asp?ServLocID="&aList(0,k)&""">"&routineHtmlString(aList(4,k))&"</a></TD>"&vbCrLf
	Response.Write "<TD NOWRAP><a target=""_parent"" href=""ServLocDetail.asp?ServLocID="&aList(0,k)&""">"&routineHtmlString(aList(5,k))&"</a></TD>"&vbCrLf
	Response.Write "<TD NOWRAP><a target=""_parent"" href=""ServLocDetail.asp?ServLocID="&aList(0,k)&""">"&routineHtmlString(aList(9,k))&"</a></TD>"&vbCrLf
	Response.Write "</TR>"

end if
next
%>

</tbody>
<tfoot>
<tr>
<td align="left" colSpan="6">
	<input type="hidden" name="txtPageNumber" value="<%=intPageNumber%>">
	<input type="submit" name="action" value="&lt;&lt;">
	<input type="submit" name="action" value="&lt;">
	<input type="text" name="txtGoToPageNo" title="You can jump to a specific page by typing the page number in this box" value="page <%=intPageNumber%> of <%=intPageCount%>" style="HEIGHT: 22px; WIDTH: 150px">
	<input type="submit" name="action" value="&gt;">
	<input type="submit" name="action" value="&gt;&gt;">
	<!--<INPUT type="submit" name="action" value="Export" title="Export this list to Excel"> -->
	<img src="images/excel.gif" onclick="btnExcel_onClick();" WIDTH="32" HEIGHT="32">
</td>
</tr>
</tfoot>
<caption>Records <%=m+1%> to <%=n+1%> of <%=UBound(aList, 2)+1 & " records"%></caption>
</table>
</form>
</body>
</html>
