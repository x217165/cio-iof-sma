<%@ Language=VBScript %>
<% option explicit %>
<!--#include file="smaConstants.inc"-->
<!--#include file="smaProcs.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--
********************************************************************************************
* Page name:	AddressList.asp
* Purpose:		To display the results of an address search.
*				Search criteria are chosen via AddressCriteria.asp
*
* Created by:	Sara Sangha	08/01/2000
*
********************************************************************************************
        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       20-Jul-01	 DTy		Exclude customers that are marked as soft-deleted and
                                  addressess that are marked as soft-deleted when
                                  'Active Only' is selected.
       12-Feb-02	 DTy		Provide 'Simplified' address listing.
                                Active customers are those whose status is either
                                  'Prospect', 'OnHold' or 'Current'.
***************************************************************************************************
-->
 <%

 dim aList, intPageNumber, intPageCount
 dim strCustomerName, strRegion, strStreet, strCity, strPostal, strProvince, bolAddressType, bolActiveOnly, strWinName
 dim strSQL, strSelectClause,strFromClause, strWhereClause, strRecordStatus, strOrderBy

 ' read submitted values and create an sql
	strWinName = trim(Request("hdnWinName"))
	strCustomerName = UCase(trim(Request("txtCustomerName")))
	strRegion = UCase(trim(Request("selRegion")))
	strStreet = UCase(trim(Request("txtStreet")))
	strCity = UCase(trim(Request("txtCity")))
	strPostal = UCase(trim(Request("txtPostal")))
	strProvince	= trim(Request("selProvince"))
	bolAddressType = trim(Request("radAddressType"))
	bolActiveOnly = trim(Request("chkActiveOnly"))

	strSQL = "select distinct(a.address_id), " &_
			"c.customer_name, " &_
			"a.billing_address_flag as billing, " &_
			"a.primary_address_flag as primary, " &_
			"a.mailing_address_flag as mailing, " &_
			"a.country_lcode, " &_
			"a.province_state_lcode, " &_
			"a.municipality_name, " &_
			"a.street, " &_
			"nvl(a.building_name, '<NO BUILDING SPECIFIED>' ), " &_
			"p.province_state_name, " &_
			"cl.country_Desc, " &_
			"m.clli_code " & _
		"from crp.customer c, " &_
			 "crp.V_ADDRESS_CONSOLIDATED_STREET a, " &_
			 "crp.customer_name_alias c1, " &_
			 "crp.lcode_country cl, " &_
			 "crp.lcode_province_state p, " &_
			 "crp.municipality_lookup m "


	strWhereClause =    "where c.customer_id = a.customer_id (+) " &_
						"and   c.customer_id = c1.customer_id " &_
						"and   a.province_state_lcode = p.province_state_lcode " &_
						"and   a.country_lcode	= cl.country_lcode " &_
						"and   a.municipality_name = m.municipality_name " &_
						"and   a.province_state_lcode = m.province_state_lcode " &_
						"and   a.country_lcode = m.country_lcode " &_
						"and   p.country_lcode = cl.country_lcode "


	if len(strCustomerName) > 0 then
		strWhereClause = strWhereClause & "and c1.customer_name_alias_upper like '" & routineOraString(strCustomerName) & "%'"
	end IF


	if len(strRegion) > 0 then
		strWhereClause = strWhereClause & "and c.noc_region_lcode = '" & strRegion & "'"
	end if

	if len(strStreet) > 0 then
		strWhereClause = strWhereClause & "and UPPER(a.street) like '" & routineOraString(strStreet) & "%'"
	end if

	if len(strCity) > 0 then
		strWhereClause = strWhereClause & "and UPPER(a.municipality_name) like '" & routineOraString(strCity) & "%'"
	end if

	if len(strPostal) > 0 then
		strWhereClause = strWhereClause & "and UPPER(a.postal_code_zip) like '" & strPostal & "%'"
	end if


	if len(strProvince) > 0 then
		strWhereClause = strWhereClause & "and (a.province_state_lcode) = '" & routineOraString(strProvince) & "'"
	end if

	select case bolAddressType
		case  "billing"
			strWhereClause = strWhereClause & " and a.billing_address_flag = 'Y' "

		case "mailing"
			strWhereClause = strWhereClause & " and a.mailing_address_flag = 'Y' "

		case "primary"
			strWhereClause = strWhereClause & " and a.primary_address_flag = 'Y' "
	end select

	'see all records?
	If bolActiveOnly = "yes" then
		strRecordStatus = " and a.record_status_ind (+) = 'A' and c.record_status_ind='A' and c1.record_status_ind = 'A'" &_
		  " and c.customer_status_lcode IN ('Prospect', 'Current', 'OnHold')"
	Else 'no
		strRecordStatus = " "
	End If

	strOrderBy = " order by c.customer_name, " &_
				 " decode(primary_address_flag,'Y',0,1)  + decode(billing_address_flag,'Y',0,1) +  decode(mailing_address_flag,'Y',0,1), "  &_
				 " a.province_state_lcode, a.municipality_name, a.street "

	strsql = strsql & strWhereClause & strRecordStatus & strOrderBy

	'Response.Write strsql			'write sql for debugging
'	Response.End


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

		case else	if Request("hdnExport") <> "" then
						'get real userid
						dim strRealUserID
						strRealUserID = Session("username")
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
						set objTxtStream = objFSO.CreateTextFile(strExportPath&strRealUserID&"-Address.xls", true, false)

						with objTxtStream
							.WriteLine "<table border=1>"

							'export the header
							.WriteLine "<TR>"
'							.WriteLine "<TR bgcolor=#ffcc99>"
							.WriteLine "<TH>Customer Name</TD>"
							.WriteLine "<TH>Primary</TD>"
							.WriteLine "<TH>Mailing</TD>"
							.WriteLine "<TH>Billing</TD>"
							.WriteLine "<TH>Province</TD>"
							.WriteLine "<TH>City</TD>"
							.WriteLine "<TH>Street</TD>"
							.WriteLine "<TH>Building Name</TD>"
							.WriteLine "<TD NOWRAP>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
							.WriteLine "</TR>"

							'export the body
							for k = 0 to UBound(aList, 2)
								'Alternate row background colour
								if Int(k/2) = k/2 then
'									.WriteLine "<TR bgcolor=#ffffcc>"
									.WriteLine "<TR>"
								else
'									.WriteLine "<TR bgcolor=#ffffff>"
									.WriteLine "<TR>"
								end if

								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(1,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(3,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(4,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(2,k))&" &nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(6,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(7,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(8,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(9,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
								.WriteLine "</TR>"
							next
							.WriteLine "</table>"
						end with
						objTxtStream.Close
						strsql = "<script type=""text/javascript"">document.location=""export/"&strRealUserID&"-Address.xls"";</script>"
						Response.Write strsql
						Response.End

						'Response.redirect "export/"&strRealUserID&"-Address.xls"

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


<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<LINK rel="stylesheet" type="text/css" href="stylesheets/styles.css">
<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
<SCRIPT LANGUAGE=javascript>
<!--
//**********************************************Java Functions ***********************************
function go_back(lngAddressID, strBuilding, strStreet, strCity, strProvince, strCountry, strClliCode, strProvinceCode){
//************************************************************************************************
// Function:	go_back
//
// Purpose:		To write the values of selected row into the base window that called the lookup
//				function. In addition, this function closes the Popup window.
//
// Created By:	Sara Sangha Aug 29th, 2000
//
// Updated By:
//************************************************************************************************

var strFullAddress ;
var exception;
	strFullAddress = strBuilding + '\n' + strStreet + '\n' +  strCity + '\n' + strProvince + '\n'+  strCountry ;

	//alert (strProvinceCode);
	parent.opener.document.forms[0].hdnAddressID.value = lngAddressID ;
	parent.opener.document.forms[0].textAddress.value = strFullAddress ;
	DeleteCookie("strSimple");

	try
	{
		//the following fields were added because they needed to be updated in ServLocDetail.asp
		parent.opener.document.forms[0].hdnProvinceCode.value = strProvinceCode;
		parent.opener.document.forms[0].hdnStreetName.value = strStreet;
		parent.opener.document.forms[0].hdnClliCode.value = strClliCode;
	}
	catch(exception)
	{}
	DeleteCookie("WinName");
	DeleteCookie("strSimple");
	parent.window.close ();

	}
//-->
//*********************************************** End of Java Functions****************************
</SCRIPT>

</HEAD>
<BODY>

<FORM method=post name=frmAddressList action="AddressList.asp">

	<input type=hidden name=hdnWinName value="<%=strWinName%>">
    <input type=hidden name=txtCustomerName value="<%=strCustomerName%>">
    <input type=hidden name=txtStreet value="<%=strStreet%>">
    <input type=hidden name=txtCity value="<%=strCity%>">
    <input type=hidden name=txtPostal value="<%=strPostal%>">
    <input type=hidden name=strProvince value="<%=strProvince%>">
    <input type=hidden name=radAddressType value="<%=bolAddressType%>">
    <input type=hidden name=chkActiveOnly value="<%=bolActiveOnly%>">
    <input type=hidden name=selRegion value="<%=strRegion%>">
    <input type=hidden name="hdnExport" value>

<TABLE  border=1 cellPadding=2 cellSpacing=0 width="100%">
 <THEAD>
    <TR>
		<TH align=left>Customer Name</TH>
		<TH align=center>Primary</TH>
		<TH align=center>Mailing</TH>
		<TH align=center>Billing</TH>
		<TH align=left>Province</TH>
	    <TH align=left>City</TH>
		<TH align=left>Street</TH>
        <TH align=left>Building</TH>
     </TR>
  </THEAD>
  <TBODY>
<%

dim strBilling, strPrimary, strMailing
'display the table
for k = m to n
	'Alternate row background colour
	if Int(k/2) = k/2 then
		Response.write "<TR>"
	else
		Response.write "<TR bgcolor=White>"
	end if

	if alist(2,k) = "Y" then
		strBilling = "=yes checked"
	else
		strBilling = ""
	end if

	if alist(3,k) = "Y" then
		strPrimary = "=yes checked"
	else
		strPrimary = ""
	end if

	if alist(4, k) = "Y" then
		strMailing = "=yes checked"
	else
		strMailing = ""
	end if

	if (strWinName = "Popup" or strWinName = "Simple") then

	Response.Write "<td><a href=""#"" onClick=""return go_back('"&aList(0,k)& "', '" &routineJavascriptString(aList(9,k))& "','" &routineJavascriptString(aList(8,k))& "','" &routineJavascriptString(aList(7,k))& "','" &routineJavascriptString(aList(10,k))& "','" &routineJavascriptString(aList(11,k))& "','" &routineJavascriptString(aList(12,k))& "','" &routineJavascriptString(aList(6,k))& "')"">" &aList(1,k)& "</a></td>"&vbCrLf
	Response.Write "<TD NOWRAP align=""center""><INPUT ID=""Primary""  name=""primary"" type=""checkbox"" style=""HEIGHT: 22px; WIDTH: 22px"" DISABLED VALUE" &strPrimary& "></TD>" &vbCrLf
	Response.Write "<TD NOWRAP align=""center""><INPUT ID=""Mailing""  name=""mailing"" type=""checkbox"" style=""HEIGHT: 22px; WIDTH: 22px"" DISABLED VALUE" &strMailing& "></TD>" &vbCrLf
	Response.Write "<TD NOWRAP align=""center""><INPUT ID=""billing""  name=""billing"" type=""checkbox"" style=""HEIGHT: 22px; WIDTH: 22px"" DISABLED VALUE" &strBilling& "></TD>" &vbCrLf
	Response.Write "<td><a href=""#"" onClick=""return go_back('"&aList(0,k)& "', '" &routineJavascriptString(aList(9,k))& "','" &routineJavascriptString(aList(8,k))& "','" &routineJavascriptString(aList(7,k))& "','" &routineJavascriptString(aList(10,k))& "','" &routineJavascriptString(aList(11,k))& "','" &routineJavascriptString(aList(12,k))& "','" &routineJavascriptString(aList(6,k))& "')"">" &aList(6,k)& "</a></td>"&vbCrLf
	Response.Write "<td><a href=""#"" onClick=""return go_back('"&aList(0,k)& "', '" &routineJavascriptString(aList(9,k))& "','" &routineJavascriptString(aList(8,k))& "','" &routineJavascriptString(aList(7,k))& "','" &routineJavascriptString(aList(10,k))& "','" &routineJavascriptString(aList(11,k))& "','" &routineJavascriptString(aList(12,k))& "','" &routineJavascriptString(aList(6,k))& "')"">" &aList(7,k)& "</a></td>"&vbCrLf
	Response.Write "<td><a href=""#"" onClick=""return go_back('"&aList(0,k)& "', '" &routineJavascriptString(aList(9,k))& "','" &routineJavascriptString(aList(8,k))& "','" &routineJavascriptString(aList(7,k))& "','" &routineJavascriptString(aList(10,k))& "','" &routineJavascriptString(aList(11,k))& "','" &routineJavascriptString(aList(12,k))& "','" &routineJavascriptString(aList(6,k))& "')"">" &aList(8,k)& "</a></td>"&vbCrLf
	Response.Write "<td><a href=""#"" onClick=""return go_back('"&aList(0,k)& "', '" &routineJavascriptString(aList(9,k))& "','" &routineJavascriptString(aList(8,k))& "','" &routineJavascriptString(aList(7,k))& "','" &routineJavascriptString(aList(10,k))& "','" &routineJavascriptString(aList(11,k))& "','" &routineJavascriptString(aList(12,k))& "','" &routineJavascriptString(aList(6,k))& "')"">" &aList(9,k)& "</a></td>"&vbCrLf
	Response.Write "</TR>"


	else

	Response.Write "<TD NOWRAP><a target=""_parent"" href=""AddressDetail.asp?AddressID="&aList(0,k)&""">"&routineHtmlString(aList(1,k))&"</a></TD>"&vbCrLf
	Response.Write "<TD NOWRAP align=""center""><INPUT ID=""Primary""  name=""primary"" type=""checkbox"" style=""HEIGHT: 22px; WIDTH: 22px"" DISABLED  VALUE" &strPrimary& "></TD>" &vbCrLf
	Response.Write "<TD NOWRAP align=""center""><INPUT ID=""Mailing""  name=""mailing"" type=""checkbox"" style=""HEIGHT: 22px; WIDTH: 22px"" DISABLED VALUE" &strMailing& "></TD>" &vbCrLf
	Response.Write "<TD NOWRAP align=""center""><INPUT ID=""billing""  name=""billing"" type=""checkbox"" style=""HEIGHT: 22px; WIDTH: 22px"" DISABLED VALUE" &strBilling& "></TD>" &vbCrLf
	Response.Write "<TD NOWRAP><a target=""_parent"" href=""AddressDetail.asp?AddressID="&aList(0,k)&""">"&routineHtmlString(aList(6,k))&"</a></TD>"&vbCrLf
	Response.Write "<TD NOWRAP><a target=""_parent"" href=""AddressDetail.asp?AddressID="&aList(0,k)&""">"&routineHtmlString(aList(7,k))&"</a></TD>"&vbCrLf
	Response.Write "<TD NOWRAP><a target=""_parent"" href=""AddressDetail.asp?AddressID="&aList(0,k)&""">"&routineHtmlString(aList(8,k))&"</a></TD>"&vbCrLf
	Response.Write "<TD NOWRAP><a target=""_parent"" href=""AddressDetail.asp?AddressID="&aList(0,k)&""">"&routineHtmlString(aList(9,k))&"</a></TD>"&vbCrLf
	Response.Write "</TR>"

	end if
next
%>
</TBODY>
<TFOOT>
<TR>
<TD align=left colSpan=8>
	<input type=hidden name=txtPageNumber value=<%=intPageNumber%>>
	<input type="submit" name="action" value="&lt;&lt;">
	<input type="submit" name="action" value="&lt;">
	<input type="text" name="txtGoToPageNo" title="You can jump to a specific page by typing the page number in this box" onclick="document.frmAddressList.txtGoToPageNo.value=''" value="page <%=intPageNumber%> of <%=intPageCount%>" style="HEIGHT: 22px; WIDTH: 150px">
	<input type="submit" name="action" value="&gt;">
	<input type="submit" name="action" value="&gt;&gt;">
	<img SRC="images/excel.gif" onclick="document.frmAddressList.target='new'; document.frmAddressList.hdnExport.value='xls';document.frmAddressList.submit();document.frmAddressList.hdnExport.value='';document.frmAddressList.target='_self';" WIDTH="32" HEIGHT="32">
</TD>
</TR>
</TFOOT>
<CAPTION>Records <%=m+1%> to <%=n+1%> of <%=UBound(aList, 2)+1 & " records"%></CAPTION>
</TABLE>
</FORM>
</BODY>
</HTML>
