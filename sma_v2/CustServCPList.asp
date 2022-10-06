﻿<%@  language="VBScript" %>
<% option explicit
   'on error resume next
%>
<% Response.Buffer = true %>
<!--#include file="SmaConstants.inc"-->
<!--#include file="SMA_Env.inc"-->
<!--#include file="smaProcs.inc"-->
<!--#include file="databaseconnect.asp"-->

<!--
***************************************************************************************************
* Name:		CustServList.asp i.e. Customer Service List
*
* Purpose:	This page reads users's search critiera and bring back a list of matching Customer
*			Service records.
*
* Created By:	Sara Sangha 08/01/00
* Edited by:    Adam Haydey 01/25/01
*               Added Customer Service City and  Customer Service Address search fields.
***************************************************************************************************
		 Date		Author			Changes/enhancements made
		06-Mar-01	 DTy		Save 'ActiveOnly' cookie for use by CustServContList.asp.
		20-Jul-01	 DTy		When 'Active Only' is selected:
		                          Exclude Service Locations that are marked as soft deleted.
		                          Exclude Customers that are marked as soft deleted.
		                          Exclude Addresses that are marked as soft deleted.
       18-Feb-02	 DTy		Active customers are those whose status is either
                                  'Prospect', 'OnHold' or 'Current'.
       28-Feb-02	 DTy		Include Customer Service Desc Alias when searching for Customer
                                  Service names.
       26-Oct-03     DTy        Add Customer Service selection from ManObjPortDetail.asp
	   13-Sept-04	  MW    	Add Lynx default severity as search fields.
       10-Aug-12    ACheung		Add Customer ID and Customer Shortname
       27-May-13	ACheung		Add Customer Profile (adapted from CustServList.asp)
       27-Jun-13	ACheung		Add NetCracker VPN web services (adapted from CustServList.asp)
***************************************************************************************************
-->

<html>
<head>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <meta http-equiv="Pragma" content="no-cache">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <link rel="stylesheet" type="text/css" href="stylesheets/styles.css">
    <script type="text/javascript" src="GeneralJavaFunctions.js"></script>
    <script type="text/javascript">

        function go_back(lngCustomerID, lngServLocID, strServiceEnd, lngCustomerServID, strCustomerServDesc, strCustomer, StrServLoc, StrAddress) {

            if (strServiceEnd == 'A') {
                //populates facility detail
                parent.opener.document.forms[0].hdnCustomerServIDA.value = lngCustomerServID;
                parent.opener.document.forms[0].hdnCustomerIdA.value = lngCustomerID;
                parent.opener.document.forms[0].hdnServiceLocIdA.value = lngServLocID;
                parent.opener.document.forms[0].txtcusserva.value = strCustomerServDesc;
                parent.opener.document.forms[0].txtcustomera.value = strCustomer;
                parent.opener.document.forms[0].txtsrvloca.value = StrServLoc;
                parent.opener.document.forms[0].txtaaddressa.value = StrAddress;
            }
            else if (strServiceEnd == 'B') {
                //populates facility detail
                parent.opener.document.forms[0].hdnCustomerServIDB.value = lngCustomerServID;
                parent.opener.document.forms[0].hdnCustomerIdB.value = lngCustomerID;
                parent.opener.document.forms[0].hdnServiceLocIdB.value = lngServLocID;
                parent.opener.document.forms[0].txtcusservb.value = strCustomerServDesc;
                parent.opener.document.forms[0].txtcustomerb.value = strCustomer;
                parent.opener.document.forms[0].txtsrvlocb.value = StrServLoc;
                parent.opener.document.forms[0].txtaaddressb.value = StrAddress;

            }
            else if (strServiceEnd == 'C') {
                //populates fields in correlation detail screen
                parent.opener.document.forms[0].hdnNewElementID.value = lngCustomerServID;
                parent.opener.document.forms[0].hdnNewElementType.value = 'Root';
                parent.opener.btn_iFrmAddNewElement();
            }

            else if (strServiceEnd == 'D') {
                //populates fields in Port Information detail screen
                parent.opener.document.forms[0].lngCSID.value = lngCustomerServID;
                parent.opener.document.forms[0].txtCSName.value = strCustomerServDesc;
            }
            DeleteCookie("WinName");
            parent.window.close();

        }


        // End of script hiding -->
    </script>
</head>
<%
 dim aList, intPageNumber, intPageCount
 dim strCustomerServiceDesc, intSupportGroupID, strCustomerName, strServiceLocationName, strOrderNo
 dim strStatusCode, intCustomerServiceID, strRegionLcode, strServiceType, bolActiveOnly, bolPrefLangOnly, strLANG
 dim strSQL, strWhereClause, strRecordStatus,strOrderBy,strMyWinName,strServiceEnd
 dim strSTypeTable, strLangPref, strLangWhere
 dim strServiceCity, strServiceAddress
 dim strLynxSeverity, intCustomerID, strCustomerShortName

 'SOAP variables
dim strwsStatus,record_count,cidList(100), strCustomerProfileName, strCustomerProfileID
'dim strvpnwsStatus, vpn_record_count, vpnList(int_maxNCWSLength), strVpnName, vpntypelist(int_maxNCWSLength), vpntoptypelist(int_maxNCWSLength), keyCSID(int_maxNCWSLength), vrflist(int_maxNCWSLength), RDlist(int_maxNCWSLength), cTvpnlist(int_maxNCWSLength), cTcustomerName(int_maxNCWSLength)
dim strvpnwsStatus, vpn_record_count, strVpnName
ReDim vpnList(int_max_csidpervpn), vpntypelist(int_max_csidpervpn), vpntoptypelist(int_max_csidpervpn), keyCSID(int_max_csidpervpn), vrflist(int_max_csidpervpn), RDlist(int_max_csidpervpn), cTvpnlist(int_max_csidpervpn), cTcustomerName(int_max_csidpervpn)

'TQ_INOSS
	strLANG = Request.Cookies("UserInformation")("language_preference")
	if (Len(strLANG) = 0) then strLANG = "EN"

	' The view is slightly slower than the table, so we speed up the
	' query by skipping the view when it isn't needed (i.e. English-only searches).
	IF (strLANG = "EN" and trim(Request.Form("chkPrefLangOnly")) = "YES") THEN
		strSTypeTable = " crp.service_type t "
		strLangPref = " 'EN' language_preference_lcode "
		strLangWhere = ""
	ELSE
		strSTypeTable = " crp.v_service_type t "
		strLangPref = " t.language_preference_lcode "
		if (trim(Request.Form("chkPrefLangOnly")) = "YES") THEN
			strLangWhere = " and t.language_preference_lcode like '" & strLANG & "' "
		else
			strLangWhere = ""
		end if
	END IF


	' Response.Write( strLANG & "<br/>" & trim(Request.Form("chkPrefLangOnly")) & "<p/>" & strSTypeTable & "<br/>" & strLangPref & "<br/>" & strLangWhere & "<p/>" )       'for debugging

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
	bolPrefLangOnly = trim(Request.Form("chkPrefLangOnly"))
	strLynxSeverity = Request("selRepairPriority")
	intCustomerID = trim(Request.Form("txtCustomerID"))
	strCustomerShortName = UCase(trim(Request.Form("txtCustomerShortName")))

	strCustomerProfileName = UCase((trim(Request.Form("txtCustomerProfileName"))))
	strCustomerProfileID = UCase((trim(Request.Form("txtCustomerProfileID"))))
	'strVpnName = UCase((trim(Request.Form("txtVpnName"))))
	strVpnName = trim(Request.Form("txtVpnName"))
     'Response.Write("debugger")
	IF (len(intCustomerID) = 0 and len(strCustomerShortName) = 0) then

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
                    "c.customer_short_name,  " &_
					"t.vpn_type_lcode,  " &_
					strLangPref &_
			"from    crp.customer_service s, " &_
					"crp.customer c,  " &_
					"crp.service_location l, " &_
					"crp.v_remedy_support_group g," &_
					"crp.address f,  " &_
					strSTypeTable

		strWhereClause = "where s.customer_id = c.customer_id " &_
						"and s.remedy_support_group_id = g.remedy_support_group_id(+) " &_
						"and s.service_type_id = t.service_type_id " &_
						"and s.service_location_id = l.service_location_id(+) " &_
						"and L.ADDRESS_ID = F.ADDRESS_ID(+) "
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
                    "c.customer_short_name,  " &_
                    "t.vpn_type_lcode,  " &_
					strLangPref &_
			 "from crp.customer_service s,  " &_
					"crp.customer c,  " &_
					"crp.service_location l, " &_
					"crp.v_remedy_support_group g, " &_
					"crp.address f,  " &_
					"crp.customer_name_alias a, " &_
				    strSTypeTable

		strWhereClause = "where s.customer_id = c.customer_id " &_
						"and s.remedy_support_group_id = g.remedy_support_group_id(+) " &_
						"and s.service_type_id = t.service_type_id " &_
						"and s.service_location_id = l.service_location_id(+) " &_
						"and a.customer_id = c.customer_id " &_
						"and L.ADDRESS_ID = F.ADDRESS_ID(+) "
	end if

'	IF len(strCustomerName) > 0 THEN
'		strWhereClause = strWhereClause & " AND a.customer_name_alias_upper LIKE '" & routineOraString(strCustomerName) &"%'"
'	END IF

	IF  LEN(intCustomerID) > 0 THEN
      strWhereClause = strWhereClause & " AND c.customer_id =" & intCustomerID
	END IF

	IF  LEN(strCustomerShortName) > 0 THEN
      strWhereClause = strWhereClause & " AND Upper(c.customer_short_name)  LIKE '" & routineOraString(strCustomerShortName) & "%' "
	END IF

	'add other search parameters to the where clause
	IF LEN(strCustomerServiceDesc) > 0 THEN
	  strWhereClause = strWhereClause & " AND s.customer_service_id in (" &_
		            " select customer_service_id from crp.customer_service where " & rtRmvSpChr("customer_service_desc", "Y") & " like '%" & rtRmvSpChr(strCustomerServiceDesc, "N") & "%' union" &_
                    " select customer_service_id from crp.customer_service_desc_alias where " & rtRmvSpChr("customer_service_desc_alias", "Y") & " like '%" & rtRmvSpChr(strCustomerServiceDesc, "N") & "%')"

	END IF

	IF LEN(strServiceLocationName) > 0 THEN
      strWhereClause = strWhereClause & " AND UPPER(l.service_location_name) LIKE '" & routineOraString(strServiceLocationName) &"%'"
	END IF

	IF LEN(intSupportGroupID) > 0 THEN
      strWhereClause = strWhereClause & " AND g.remedy_support_group_id = " &  intSupportGroupID
	END IF

	IF LEN(strStatusCode) > 0 THEN
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

	if strLynxSeverity <> "" then
		strWhereClause = strWhereClause & "AND s.lynx_def_sev_lcode = '" & routineOraString(strLynxSeverity) & "'"
	end if

	'CPID entered
	If strCustomerProfileID <> "" then
		'Response.Write strCustomerProfileID & "<br />" & vbCrLf

		If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

			strwsStatus = CP_GetCustomerID(strCustomerProfileID,100,record_count,cidList)
			'Response.write "<p>Status = " & strwsStatus & "</p>"
			'Response.write "<p>Size = " & record_count & "</p>"
			if 	strwsStatus <> 200 or record_count = 0 then
				strWhereClause = strWhereClause &  " and " & _
				"c.customer_id in (-999)"   'assuming there is no customer id = -999
			else
				Dim  cidindex
				'for cidindex = 0 to record_count
				'	Response.write "<p>CID " & cidindex & " = " & routineOraString(cidList(cidindex)) & "</p>"
				'next

				for cidindex = 0 to record_count
					'Response.write "<p>in cid loop " & cidindex & " = " & routineOraString(cidList(cidindex)) &"</p>"
					if cidindex = 0 Then
						strWhereClause = strWhereClause &  " and " & _
						"c.customer_id in (" & int(cidList(cidindex))
					elseif cidindex < record_count Then
						strWhereClause = strWhereClause & ", " & int(cidList(cidindex))
					elseif cidindex = record_count Then
						strWhereClause = strWhereClause & ") "
					end if
				next
			end if 'WSstatus
		end if 'WS
	End if	'CPID is not null

	'VPNname entered
	If strVpnName <> "" then
		'Response.Write strVpnName & "<br />" & vbCrLf

		If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

			'Response.write "<p>Before ws call strVpnName = " & strVpnName & "</p>"
			strvpnwsStatus = nc_getCSIDListByVPN(strVpnName,int_max_csidpervpn,vpn_record_count,vpnList)
			'Response.write "<p>Status = " & strvpnwsStatus & "</p>"
			'Response.write "<p>Size = " & vpn_record_count & "</p>"
			if 	strvpnwsStatus <> 200 or vpn_record_count = 0 then
				strWhereClause = strWhereClause &  " and " & _
				"s.customer_service_id in (-999)"   'assuming there is no customer id = -999
			else
				Dim  vpnindex

				for vpnindex = 0 to vpn_record_count
					'Response.write "<p>in cid loop " & vpnindex & " = " & routineOraString(vpnList(vpnindex)) &"</p>"
					if vpnindex = 0 Then
						strWhereClause = strWhereClause &  " and " & _
						"s.customer_service_id in (" & int(vpnList(vpnindex))
					elseif vpnindex < vpn_record_count Then
						strWhereClause = strWhereClause & ", " & int(vpnList(vpnindex))
					elseif vpnindex = vpn_record_count Then
						strWhereClause = strWhereClause & ") "
					end if
				next
			end if 'WSstatus
		end if 'WS
	End if	'VPNname is not null

    Response.Cookies ("ActiveOnly")=bolActiveOnly

	if bolActiveOnly = "YES" then
		strRecordStatus = " and c.customer_status_lcode IN ('Prospect', 'Current', 'OnHold')" &_
                          " and s.record_status_ind (+) = 'A' and l.record_status_ind (+) = 'A'" & _
		                  " and c.record_status_ind = 'A' and f.record_status_ind (+) = 'A' "

'        if strLynxSeverity <> "ALL" then
'               strRecordStatus = strRecordStatus &_
'			  "AND s.lynx_def_sev_lcode = '"  &routineOraString(strLynxSeverity)& "' "
'	    end if
	else
		'display all record
		strRecordStatus = " "
	end if

	if bolPrefLangOnly = "YES" then
	strWhereClause = strWhereClause & strLangWhere
	end if

	strOrderBy = " order by upper(s.customer_service_desc)"

	'join all pieces to make a complete query
	strsql = strSQL & strWhereClause & strRecordStatus & strOrderBy

	'Response.Write( strsql )       'display SQL for debugging
	'response.end
	'Response.Write "strCustomerName ="
	'Response.Write (strCustomerName)
	'Response.Write "intCustomerID ="
	'Response.Write (intCustomerID)
	'Response.Write "strCustomerShortName="
	'Response.Write (strCustomerShortName)
	'Response.End

	Dim objRsResult,Recordcnt,strbgcolor

	set objRsResult = objConn.Execute(strSql)
	if not objRsResult.EOF then
		aList = objRsResult.GetRows
	else
		Response.Write "0 records found"
		Response.end
	end if

	'VPN Name at Column
	' aList(13,k) is the VPN name/VPN_LCODE, aList(3,k) is CSID
	If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

		Dim  vindex, vpnname(10), vpn_map
'		set vpn_map = CreateObject("Scripting.Dictionary")
		for k = 0 to UBound(aList, 2)
			if routineHtmlString(aList(13,k)) <> 0 then 'VPN_TYPE_LCODE <> 0
  
				'Response.Write(routineHtmlString(aList(3,k)))
				aList(13,k) = ""	'No VPN Name Assigned for this VPN capable Service Type
				strvpnwsStatus =nc_getVPNListByService(routineHtmlString(aList(3,k)), vpn_record_count, vpnlist, vpntypelist, vpntoptypelist, keyCSID, vrflist, RDlist, cTvpnlist, cTcustomerName)
				'Response.write "<p>Status = " & strvpnwsStatus & "</p>"
				'Response.write "<p>Size = " & vpn_record_count & "</p>"
				'Response.write "<p>k = " & k & "</p>"
				for vindex = 0 to vpn_record_count-1			'donot comment
					'Response.write "<p>vindex = " & vindex & "</p>"
					'Response.write "<p>keyCSID(vindex) = " & keyCSID(vindex) & "</p>"
					'Response.write "<p>aList(3,k) = " & routineHtmlString(aList(3,k)) & "</p>"
					'Response.write "<p>aList(13,k) = " & routineHtmlString(aList(13,k)) & "</p>"
					'Response.write "<p>vpnlist(vindex) = " & routineHtmlString(vpnlist(vindex)) & "</p>"
					if routineHtmlString(aList(3,k)) = keyCSID(vindex) then
					  if vpnlist(vindex) <> "" and IsEmpty(vpnlist(vindex)) = false then
						 if vindex = 0  then
						   aList(13,k) = vpnlist(vindex)
    					 else                     
					       aList(13,k) = aList(13,k) + ";" + vpnlist(vindex)
					     end if
					  end if
					end if
				next
			else
				aList(13,k) = ""		'Not VPN Service Type or VPN_TYPE_LCODE = 0
			end if
		next
	End if

   'release and kill the recordset and the connection objects
	objRsResult.Close
	set objRsResult = nothing

	objConn.close
	set objConn = nothing

    ' Response.Write("Debugger")
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
						set objTxtStream = objFSO.CreateTextFile(strExportPath&strRealUserID&"-CustomerService.xls", true, false)

						with objTxtStream
							.WriteLine "<table border=1>"

							'export the header
							.WriteLine "<TR>"
							.WriteLine "<TH>Customer Service Name</TD>"
							.WriteLine "<TH>Status</TD>"
							.WriteLine "<TH>Service ID</TD>"
							.WriteLine "<TH>Service Type</TD>"
							.WriteLine "<TH>Language Code</TD>"
							.WriteLine "<TH>Service Location</TD>"
							.WriteLine "<TH>Customer Name</TD>"
							.WriteLine "<TH>CSN</TD>"
							.WriteLine "<TH>CID</TD>"
							.WriteLine "<TH>Region</TD>"
							.WriteLine "<TH>Support Group</TD>"
							.WriteLine "<TH>VPN name</TD>"
							.WriteLine "<TD NOWRAP>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
							.WriteLine "</TR>"

							'export the body
							for k = 0 to UBound(aList, 2)
								.WriteLine "<TR>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(1,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(2,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(3,k))&"</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(11,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(14,k))&"</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(4,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(5,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(12,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(9,k))&"</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(6,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(7,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(13,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
								.WriteLine "</TR>"
							next
							.WriteLine "</table>"
						end with
						objTxtStream.Close
						strsql = "<script type=""text/javascript"">document.location=""export/"&strRealUserID&"-CustomerService.xls"";</script>"
						Response.Write strsql
						Response.End
'						Response.redirect "export/"&strRealUserID&"-CustomerService.xls"
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
<body>
    <form method="post" name="frmCustServCPList" action="CustServCPList.asp">

        <input type="hidden" name="hdnWinName" value="<%=strMyWinName%>">
        <input type="hidden" name="txtCustomerServiceDesc" value="<%=strCustomerServiceDesc%>">
        <input type="hidden" name="txtServiceLocationName" value="<%=strServiceLocationName%>">
        <input type="hidden" name="selSupportGroup" value="<%=intSupportGroupID%>">
        <input type="hidden" name="txtCustomerName" value="<%=strCustomerName%>">
        <input type="hidden" name="SelStatus" value="<%=strStatusCode%>">
        <input type="hidden" name="txtCustomerServiceID" value="<%=intCustomerServiceID%>">
        <input type="hidden" name="txtOrderNo" value="<%=strOrderNo%>">
        <input type="hidden" name="selRegion" value="<%=strRegionLcode%>">
        <input type="hidden" name="chkActiveOnly" value="<%=bolActiveOnly%>">
        <input type="hidden" name="chkPrefLangOnly" value="<%=bolPrefLangOnly%>">
        <input type="hidden" name="hdnServiceEnd" value="<%=strServiceEnd%>">
        <input type="hidden" name="txtServiceType" value="<%=strServiceType%>">
        <input type="hidden" name="txtServiceCity" value="<%=strServiceCity%>">
        <input type="hidden" name="txtServiceAddress" value="<%=strServiceAddress%>">
        <input type="hidden" name="selRepairPriority" value="<%=strLynxSeverity%>">
        <input type="hidden" name="hdnExport" value>
        <input type="hidden" name="txtCustomerID" value="<%=intCustomerID%>">
        <input type="hidden" name="txtCustomerShortName" value="<%=strCustomerShortName%>">
        <input type="hidden" name="txtCustomerProfileName" value="<%=strCustomerProfileName%>">
        <input type="hidden" name="txtCustomerProfileID" value="<%=strCustomerProfileID%>">
        <input type="hidden" name="txtVpnName" value="<%=strVpnName%>">

        <table border="1" cellpadding="2" cellspacing="0" width="100%">
            <thead>
                <tr>
                    <th align="left" nowrap>Customer Service Name</th>
                    <th align="left" nowrap>Status</th>
                    <th align="left" nowrap>Service ID</th>
                    <th align="left" nowrap>Service Type</th>
                    <th align="left" nowrap title="Service Type Language Code">LC</th>
                    <th align="left" nowrap>Service Location</th>
                    <th align="left" nowrap>Customer Name</th>
                    <th align="left" nowrap>CSN</th>
                    <th align="left" nowrap>CID</th>
                    <th align="left" nowrap>Region</th>
                    <th align="left" nowrap>Support Group</th>
                    <th align="left" nowrap>VPN Name</th>
                </tr>
            </thead>
            <tbody>
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
		Response.Write "<TD nowrap><a href=""#""  title=""Service Type Language Code"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(0,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(14, k))& "</a></TD>" &vbCrLf
		Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(0,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(4, k))& "&nbsp;</a></TD>" &vbCrLf
		Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(0,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(5, k))& "&nbsp;</a></TD>" &vbCrLf
		Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(0,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(12, k))& "&nbsp;</a></TD>" &vbCrLf
		Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(0,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(9, k))& "&nbsp;</a></TD>" &vbCrLf
		Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(0,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(6, k))& "&nbsp;</a></TD>" &vbCrLf
		Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(0,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(7, k))& "&nbsp;</a></TD>" &vbCrLf
		Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(0,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(13, k))& "&nbsp;</a></TD>" &vbCrLf
		Response.Write "</TR>"

	else

		Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServCPDetail.asp?CustServID="&aList(0,k)&""">"&routineHtmlString(aList(1,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
		Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServCPDetail.asp?CustServID="&aList(0,k)&""">"&routineHtmlString(aList(2,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
		Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServCPDetail.asp?CustServID="&aList(0,k)&""">"&routineHtmlString(aList(3,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
		Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServCPDetail.asp?CustServID="&aList(0,k)&""">"&routineHtmlString(aList(11,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
		Response.write "<TD nowrap><a TARGET=""_parent"" title=""Service Type Language Code"" href=""CustServCPDetail.asp?CustServID="&aList(0,k)&""">"&routineHtmlString(aList(14,k))&"</a>&nbsp;</TD>" &vbCrLf
		Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServCPDetail.asp?CustServID="&aList(0,k)&""">"&routineHtmlString(aList(4,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
		Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServCPDetail.asp?CustServID="&aList(0,k)&""">"&routineHtmlString(aList(5,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
		Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServCPDetail.asp?CustServID="&aList(0,k)&""">"&routineHtmlString(aList(12,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
		Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServCPDetail.asp?CustServID="&aList(0,k)&""">"&routineHtmlString(aList(9,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
		Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServCPDetail.asp?CustServID="&aList(0,k)&""">"&routineHtmlString(aList(6,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
		Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServCPDetail.asp?CustServID="&aList(0,k)&""">"&routineHtmlString(aList(7,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
		Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServCPDetail.asp?CustServID="&aList(0,k)&""">"&routineHtmlString(aList(13,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
		Response.Write "</TR>"


	end if
next
                %>
            </tbody>
            <tfoot>
                <tr>
                    <td align="left" colspan="12">
                        <input type="hidden" name="txtPageNumber" value="<%=intPageNumber%>">
                        <input type="submit" name="action" value="&lt;&lt;">
                        <input type="submit" name="action" value="&lt;">
                        <input type="text" name="txtGoToPageNo" title="You can jump to a specific page by typing the page number in this box" onclick="document.frmCustServCPList.txtGoToPageNo.value = ''" value="page <%=intPageNumber%> of <%=intPageCount%>" style="height: 22px; width: 150px">
                        <input type="submit" name="action" value="&gt;">
                        <input type="submit" name="action" value="&gt;&gt;">
                        <img src="images/excel.gif" onclick="document.frmCustServCPList.target='new'; document.frmCustServCPList.hdnExport.value='xls';document.frmCustServCPList.submit();document.frmCustServCPList.hdnExport.value='';document.frmCustServCPList.target='_self';" width="32" height="32">
                    </td>
                </tr>
            </tfoot>
            <caption>Records <%=m+1%> to <%=n+1%> of <%=UBound(aList, 2)+1 & " records"%></caption>
        </table>
    </form>
</body>
</html>
