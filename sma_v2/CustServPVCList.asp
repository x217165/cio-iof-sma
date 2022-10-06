<%@ Language=VBScript %>
<% option explicit
   'on error resume next
 %>
 <% Response.Buffer = true %>
<!--#include file="SmaConstants.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--#include file="smaProcs.inc"-->

<!--
***************************************************************************************************
* Name:		CustServPVCList.asp i.e. Customer Service List for PVC
*
* Purpose:	This page reads users's search critiera and bring back a list of matching Customer
*			Service records with extra features for PVC detail.
*
* Created By:	Original CustServList Sara Sangha 08/01/00
* Edited by:    Adam Haydey 01/25/01
*               Added Customer Service City and  Customer Service Address search fields.
***************************************************************************************************

		 Date		Author			Changes/enhancements made
		06-Mar-01	 DTy		Save 'ActiveOnly' cookie for use by CustServContList.asp.
		12-Mar-01	A Haydey		Created the CustServPVCList page as a separate page to allow for
									managed object name to be part of the search criteria and as well
									displaying project code when called from the PVC Detail screen.
       18-Feb-02	 DTy		Active customers are those whose status is either
                                  'Prospect', 'OnHold' or 'Current'.
***************************************************************************************************
-->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<LINK rel="stylesheet" type="text/css" href="stylesheets/styles.css">
<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
<script type="text/javascript">

	function go_back(lngCustomerID,lngServLocID,strServiceEnd,lngCustomerServID, strCustomerServDesc,strCustomer,StrServLoc,StrAddress)
     {

	  if (strServiceEnd == 'A'){
	   //populates facility detail
		parent.opener.document.forms[0].hdnCustomerServIDA.value = lngCustomerServID;
		parent.opener.document.forms[0].hdnCustomerIdA.value = lngCustomerID;
		parent.opener.document.forms[0].hdnServiceLocIdA.value = lngServLocID;
		parent.opener.document.forms[0].txtcusserva.value = strCustomerServDesc;
		parent.opener.document.forms[0].txtcustomera.value = strCustomer;
		parent.opener.document.forms[0].txtsrvloca.value = StrServLoc;
		parent.opener.document.forms[0].txtaaddressa.value =StrAddress;
       }
	  else if (strServiceEnd == 'B'){
       //populates facility detail
        parent.opener.document.forms[0].hdnCustomerServIDB.value = lngCustomerServID;
        parent.opener.document.forms[0].hdnCustomerIdB.value = lngCustomerID;
		parent.opener.document.forms[0].hdnServiceLocIdB.value = lngServLocID;
		parent.opener.document.forms[0].txtcusservb.value = strCustomerServDesc;
		parent.opener.document.forms[0].txtcustomerb.value = strCustomer;
		parent.opener.document.forms[0].txtsrvlocb.value = StrServLoc;
		parent.opener.document.forms[0].txtaaddressb.value =StrAddress;

       }
	  else if (strServiceEnd == 'C'){
       //populates fields in correlation detail screen
        parent.opener.document.forms[0].hdnNewElementID.value = lngCustomerServID;
		parent.opener.document.forms[0].hdnNewElementType.value = 'Root';
		parent.opener.btn_iFrmAddNewElement();
       }

		DeleteCookie("WinName");
	    parent.window.close ();

      }


// End of script hiding -->
</script>
</HEAD>
 <%

 dim aList, intPageNumber, intPageCount
 dim strCustomerServiceDesc, intSupportGroupID, strCustomerName, strServiceLocationName, strOrderNo
 dim strStatusCode, intCustomerServiceID, strRegionLcode, strServiceType, bolActiveOnly
 dim strSQL, strWhereClause, strRecordStatus,strOrderBy,strMyWinName,strServiceEnd
 dim strServiceCity, strServiceAddress, strManObjName

	strMyWinName = Request.Form("hdnWinName")
	strServiceEnd = Request.Form("hdnServiceEnd")
	strCustomerServiceDesc = UCase(trim(Request.Form("txtCustomerServiceDesc")))
	strServiceLocationName = UCase(trim(Request.Form("txtServiceLocationName")))
	intSupportGroupID = trim(Request.Form("selSupportGroup"))
	strCustomerName = UCase(trim(Request.Form("txtCustomerName")))
	strStatusCode = trim(Request.Form("SelStatus"))
	intCustomerServiceID = trim(Request.Form("txtCustomerServiceID"))
	strOrderNo = trim(Request.Form("txtOrderNo"))
	strServiceType = UCase(trim(Request.Form("txtServiceType")))
	strRegionLcode = trim(Request.Form("selRegion"))
	strServiceCity = UCase(Request.Form("txtServiceCity"))
	strServiceAddress = UCase(Request.Form("txtServiceAddress"))
	bolActiveOnly = trim(Request.Form("chkActiveOnly"))
	strManObjName = UCase(Request.Form("txtManObjName"))

	'Response.Write("Here:")
	'Response.Write(Request.Form("txtManObjName"))
	'Response.Write(strManObjName)
	'Response.End

	IF len(strCustomerName) = 0 then

		strSQL = "select s.customer_service_id, " &_
					"s.customer_service_desc, " &_
					"s.service_status_code, " &_
					"s.customer_service_id, " &_
					"l.service_location_name, " &_
					"c.customer_name, " &_
					"c.noc_region_lcode, " &_
					"g.group_name, " &_
					"NVL(F.BUILDING_NAME,'NO BUILDING NAME') ||CHR(13)||CHR(10)|| " &_
					"decode(F.APARTMENT_NUMBER, null, null, rtrim(F.APARTMENT_NUMBER) || ' ') || " &_
					"decode(F.HOUSE_NUMBER, null, null, rtrim(f.house_number)  || ' ') || " &_
					"decode(F.STREET_VECTOR, null, null, rtrim(F.STREET_VECTOR) || ' ') || " &_
					"NVL(F.LONG_STREET_NAME,'NO STREET ADDRESS')||CHR(13)||CHR(10)|| " &_
					"NVL(F.MUNICIPALITY_NAME,'NO MUNICIPALITY')||' '|| " &_
					"NVL(F.PROVINCE_STATE_LCODE,'NO PROVINCE')||CHR(13)||CHR(10)|| " &_
					"NVL(F.POSTAL_CODE_ZIP,'NO POSTAL CODE') ADDRESS, " &_
                    "c.customer_id,  " &_
                    "s.service_location_id,  " &_
                    "t.service_type_desc, " &_
                    "s.project_code "
                    if len(strManObjName)> 0 then
						strSQL = strSQL & ", ne.network_element_name "
					end if
		strSQL = strSQL & "from    crp.customer_service s, " &_
					"crp.customer c,  " &_
					"crp.service_location l, " &_
					"crp.v_remedy_support_group g," &_
					"crp.address f,  " &_
					"crp.service_type t "
		if len(strManObjName)> 0 then
			strSQL = strSQL & ", crp.managed_correlation mc " &_
							  ", crp.network_element ne "
		end if

		strWhereClause = "where s.customer_id = c.customer_id " &_
						"and	  s.remedy_support_group_id = g.remedy_support_group_id(+) " &_
						"and	  s.service_type_id = t.service_type_id " &_
						"and	  s.service_location_id = l.service_location_id(+) " &_
						"and      L.ADDRESS_ID = F.ADDRESS_ID(+) "
		if len(strManObjName)> 0 then
			strWhereClause = strWhereClause & " and s.customer_service_id = mc.customer_service_id " &_
											  " and mc.network_element_id = ne.network_element_id "
		end if

	else
		strSQL = "select distinct(s.customer_service_id), " &_
					"s.customer_service_desc, " &_
					"s.service_status_code, " &_
					"s.customer_service_id, " &_
					"l.service_location_name, " &_
					"c.customer_name, " &_
					"c.noc_region_lcode, " &_
					"g.group_name,  " &_
					"NVL(F.BUILDING_NAME,'NO BUILDING NAME') ||CHR(13)||CHR(10)|| " &_
					"decode(F.APARTMENT_NUMBER, null, null, rtrim(F.APARTMENT_NUMBER) || ' ') || " &_
					"decode(F.HOUSE_NUMBER, null, null, rtrim(f.house_number)  || ' ') || " &_
					"decode(F.STREET_VECTOR, null, null, rtrim(F.STREET_VECTOR) || ' ') || " &_
					"NVL(F.LONG_STREET_NAME,'NO STREET ADDRESS')||CHR(13)||CHR(10)|| " &_
					"NVL(F.MUNICIPALITY_NAME,'NO MUNICIPALITY')||' '|| " &_
					"NVL(F.PROVINCE_STATE_LCODE,'NO PROVINCE')||CHR(13)||CHR(10)|| " &_
					"NVL(F.POSTAL_CODE_ZIP,'NO POSTAL CODE') ADDRESS, " &_
                    "c.customer_id,  " &_
                    "s.service_location_id,  " &_
                    "t.service_type_desc, " &_
                    "s.project_code "
              if len(strManObjName)> 0 then
						strSQL = strSQL & ", ne.network_element_name "
				end if
			 strSQL = strSQL & "from crp.customer_service s,  " &_
					"crp.customer c,  " &_
					"crp.service_location l, " &_
					"crp.v_remedy_support_group g, " &_
					"crp.address f,  " &_
					"crp.customer_name_alias a, " &_
					"crp.service_type t "
			if len(strManObjName)> 0 then
				strSQL = strSQL & ", crp.managed_correlation mc " &_
							  ", crp.network_element ne "
			end if
		strWhereClause = "where s.customer_id = c.customer_id " &_
						"and	  s.remedy_support_group_id = g.remedy_support_group_id(+) " &_
						"and	  s.service_type_id = t.service_type_id " &_
						"and	  s.service_location_id = l.service_location_id(+) " &_
						"and	  a.customer_id = c.customer_id " &_
						"and      L.ADDRESS_ID = F.ADDRESS_ID(+) "

		strWhereClause = strWhereClause & " AND a.customer_name_alias_upper LIKE '" & routineOraString(strCustomerName) &"%'"
		if len(strManObjName)> 0 then
			strWhereClause = strWhereClause & " and s.customer_service_id = mc.customer_service_id " &_
											  " and mc.network_element_id = ne.network_element_id "
		end if
	end if


	'add other search parameters to the where clause
	IF  LEN(strCustomerServiceDesc) > 0 THEN
      strWhereClause = strWhereClause & " AND UPPER(s.customer_service_desc) LIKE '" & routineOraString(strCustomerServiceDesc) &"%'"
	END IF

	IF  LEN(strServiceLocationName) > 0 THEN
      strWhereClause = strWhereClause & " AND UPPER(l.service_location_name) LIKE '" & routineOraString(strServiceLocationName) &"%'"
	END IF

	IF  LEN(intSupportGroupID) > 0 THEN
      strWhereClause = strWhereClause & " AND g.remedy_support_group_id = " &  intSupportGroupID
	END IF

	IF  LEN(strStatusCode) > 0 THEN
		if strStatusCode = "AllExceptTerm" then
			strWhereClause = strWhereClause & " AND s.service_status_code <> 'TERM'"
		else
			strWhereClause = strWhereClause & " AND s.service_status_code = '" & routineOraString(strStatusCode) & "'"
		end if
    END IF

	IF  LEN(intCustomerServiceID) > 0 THEN
      strWhereClause = strWhereClause & " AND s.customer_service_id =" & intCustomerServiceID
	END IF

	IF  LEN(strServiceType) > 0 THEN
      strWhereClause = strWhereClause & " AND Upper(t.service_type_desc)  LIKE '" & routineOraString(strServiceType) & "%' "
	END IF

	IF  LEN(strServiceCity) > 0 THEN
      strWhereClause = strWhereClause & " AND Upper(f.municipality_name)  LIKE '" & routineOraString(strServiceCity) & "%' "
	END IF

	IF  LEN(strServiceAddress) > 0 THEN
      strWhereClause = strWhereClause & " AND Upper(NVL(F.BUILDING_NAME,'NO BUILDING NAME') ||CHR(13)||CHR(10)|| " &_
					"decode(F.APARTMENT_NUMBER, null, null, rtrim(F.APARTMENT_NUMBER) || ' ') || " &_
					"decode(F.HOUSE_NUMBER, null, null, rtrim(f.house_number)  || ' ') || " &_
					"decode(F.STREET_VECTOR, null, null, rtrim(F.STREET_VECTOR) || ' ') || " &_
					"NVL(F.LONG_STREET_NAME,'NO STREET ADDRESS')||CHR(13)||CHR(10)|| " &_
					"NVL(F.MUNICIPALITY_NAME,'NO MUNICIPALITY')||' '|| " &_
					"NVL(F.PROVINCE_STATE_LCODE,'NO PROVINCE')||CHR(13)||CHR(10)|| " &_
					"NVL(F.POSTAL_CODE_ZIP,'NO POSTAL CODE'))  LIKE '" & routineOraString(strServiceAddress) & "%' "
	END IF

	IF  LEN(strRegionLcode) > 0 THEN
      strWhereClause = strWhereClause & " AND c.noc_region_lcode = '" & routineOraString(strRegionLcode) & "'"
	END IF

	if len(strOrderNo) >  0 then
		strWhereClause = strWhereClause & " AND s.project_code = '" & routineOraString(strOrderNo) & "'"
	end if

	if len(strManObjName)> 0 then
			strWhereClause = strWhereClause  & " AND ne.network_element_name like '" & routineOraString(strManObjName) & "%'"

	end if


    Response.Cookies ("ActiveOnly")=bolActiveOnly
	if bolActiveOnly = "YES" then
		strRecordStatus = " and c.customer_status_lcode IN ('Prospect', 'Current', 'OnHold')" &_
		                  " and s.record_status_ind = 'A' "
	else
		'display all record
		strRecordStatus = " "
	end if

	strOrderBy = " order by upper(s.customer_service_desc)"

	'join all pieces to make a complete query
	strsql = strSQL & strWhereClause & strRecordStatus & strOrderBy

	'Response.Write( strsql )       'display SQL for debugging
	'Response.End

	Dim objRsResult,Recordcnt,strbgcolor

	set objRsResult = objConn.Execute(StrSql)
	if not objRsResult.EOF then
		aList = objRsResult.GetRows
	else
		Response.Write "0 records found"
		Response.end
	end if

   'release and kill the recordset and the connection objects
	objRsResult.Close
	set objRsResult = nothing

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
						set objTxtStream = objFSO.CreateTextFile(strExportPath&strRealUserID&"-CustomerServicePVC.xls", true, false)

						with objTxtStream
							.WriteLine "<table border=1>"

							'export the header
							.WriteLine "<TR>"
'							.WriteLine "<TR bgcolor=#ffcc99>"
							.WriteLine "<TH>Customer Service Name</TH>"
							.WriteLine "<TH>Status</TH>"
							.WriteLine "<TH>Service ID</TH>"
							.WriteLine "<TH>Service Type</TH>"
							.WriteLine "<TH>Service Location</TH>"
							.WriteLine "<TH>Customer Name</TH>"
							.WriteLine "<TH>Region</TH>"
							.WriteLine "<TH>Support Group</TH>"
							.WriteLine "<TH align=left nowrap>Project Code</TH>"
							if len(strManObjName)> 0 then
								.WriteLine "<TH align=left nowrap>Managed Object Name</TH>"
							end if
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
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(2,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(3,k))&"</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(11,k))&" &nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(4,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(5,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(6,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(7,k))&" &nbsp; </TD>"

									.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(12,k))&" &nbsp; </TD>"
									if len(strManObjName)> 0 then
										.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(13,k))&" &nbsp; </TD>"
									end if

								.WriteLine "<TD NOWRAP>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
								.WriteLine "</TR>"
							next
							.WriteLine "</table>"
						end with
						objTxtStream.Close
						strsql = "<script type=""text/javascript"">document.location=""export/"&strRealUserID&"-CustomerServicePVC.xls"";</script>"
						Response.Write strsql
						Response.End
'						Response.redirect "export/"&strRealUserID&"-CustomerServicePVC.xls"
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
<BODY>
<FORM method=post name=frmCustServPVCList action="CustServPVCList.asp">

    <input type=hidden name=hdnWinName value="<%=strMyWinName%>">
    <input type=hidden name=txtCustomerServiceDesc value="<%=strCustomerServiceDesc%>">
    <input type=hidden name=txtServiceLocationName value="<%=strServiceLocationName%>">
    <input type=hidden name=selSupportGroup value="<%=intSupportGroupID%>">
    <input type=hidden name=txtCustomerName value="<%=strCustomerName%>">
    <input type=hidden name=SelStatus value="<%=strStatusCode%>">
    <input type=hidden name=txtCustomerServiceID value="<%=intCustomerServiceID%>">
    <input type=hidden name=txtOrderNo value="<%=strOrderNo%>">
    <input type=hidden name=selRegion value="<%=strRegionLcode%>">
    <input type=hidden name=chkActiveOnly value="<%=bolActiveOnly%>">
    <input type=hidden name=hdnServiceEnd value="<%=strServiceEnd%>">
    <input type=hidden name=txtServiceType value="<%=strServiceType%>"  >
    <input type=hidden name=txtServiceCity value="<%=strServiceCity%>"  >
    <input type=hidden name=txtServiceAddress value="<%=strServiceAddress%>"  >
    <input type=hidden name=txtManObjName value="<%=strManObjName%>"  >
    <input type=hidden name="hdnExport" value>


<TABLE  border=1 cellPadding=2 cellSpacing=0 width="100%">
  <THEAD>
    <TR>
        <TH align=left nowrap>Customer Service Name</TH>
        <TH align=left nowrap>Status</TH>
        <TH align=left nowrap>Service ID</TH>
        <TH align=left nowrap>Service Type</TH>
        <TH align=left nowrap>Service Location</TH>
        <TH align=left nowrap>Customer Name</TH>
        <TH align=left nowrap>Region</TH>
        <TH align=left nowrap>Support Group</TH>
        <TH align=left nowrap>Project Code</TH>
        <%if len(strManObjName)> 0 then
				Response.Write "<TH align=left nowrap>Managed Object Name</TH>"
		end if %>
     </TR>
 </THEAD>
 <TBODY>
<%for k = m to n
	'Alternate row background colour
	if Int(k/2) = k/2 then
		Response.write "<TR>"
	else
		Response.write "<TR bgcolor=White>"
	end if

	if strMyWinName = "Popup" then

		Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(0,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(1, k))& "&nbsp;</a></TD>" &vbCrLf
		Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(0,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(2, k))& "&nbsp;</a></TD>" &vbCrLf
		Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(0,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(3, k))& "&nbsp;</a></TD>" &vbCrLf
		Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(0,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(11, k))& "&nbsp;</a></TD>" &vbCrLf
		Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(0,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(4, k))& "&nbsp;</a></TD>" &vbCrLf
		Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(0,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(5, k))& "&nbsp;</a></TD>" &vbCrLf
		Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(0,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(6, k))& "&nbsp;</a></TD>" &vbCrLf
		Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(0,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(7, k))& "&nbsp;</a></TD>" &vbCrLf
		Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(0,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(12, k))& "&nbsp;</a></TD>" &vbCrLf
		if len(strManObjName)> 0 then
			Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(0,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(13, k))& "&nbsp;</a></TD>" &vbCrLf
		end if
		Response.Write "</TR>"

	else

		Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServDetail.asp?CustServID="&aList(0,k)&""">"&routineHtmlString(aList(1,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
		Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServDetail.asp?CustServID="&aList(0,k)&""">"&routineHtmlString(aList(2,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
		Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServDetail.asp?CustServID="&aList(0,k)&""">"&routineHtmlString(aList(3,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
		Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServDetail.asp?CustServID="&aList(0,k)&""">"&routineHtmlString(aList(11,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
		Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServDetail.asp?CustServID="&aList(0,k)&""">"&routineHtmlString(aList(4,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
		Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServDetail.asp?CustServID="&aList(0,k)&""">"&routineHtmlString(aList(5,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
		Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServDetail.asp?CustServID="&aList(0,k)&""">"&routineHtmlString(aList(6,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
		Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServDetail.asp?CustServID="&aList(0,k)&""">"&routineHtmlString(aList(7,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf

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
	<input type="text" name="txtGoToPageNo" title="You can jump to a specific page by typing the page number in this box" onclick="document.frmCustServPVCList.txtGoToPageNo.value=''" value="page <%=intPageNumber%> of <%=intPageCount%>" style="HEIGHT: 22px; WIDTH: 150px">
	<input type="submit" name="action" value="&gt;">
	<input type="submit" name="action" value="&gt;&gt;">
	<img SRC="images/excel.gif" onclick="document.frmCustServPVCList.target='new'; document.frmCustServPVCList.hdnExport.value='xls';document.frmCustServPVCList.submit();document.frmCustServPVCList.hdnExport.value='';document.frmCustServPVCList.target='_self';" WIDTH="32" HEIGHT="32">
</TD>
</TR>
</TFOOT>
<CAPTION>Records <%=m+1%> to <%=n+1%> of <%=UBound(aList, 2)+1 & " records"%></CAPTION>
</TABLE>
</FORM>
</BODY>
</HTML>
