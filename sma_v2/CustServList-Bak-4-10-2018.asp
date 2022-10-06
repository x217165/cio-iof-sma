<%@  language="VBScript" %>
<% option explicit
   'on error resume next
%>
<% Response.Buffer = true %>
<!--#include file="SmaConstants.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--#include file="smaProcs.inc"-->

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
***************************************************************************************************
-->

<html>
<head>

    <style>
        /* The Modal (background) */
        .modal {
            display: none; /* Hidden by default */
            position: fixed; /* Stay in place */
            z-index: 1; /* Sit on top */
            padding-top: 70px; /* Location of the box */
            left: 0;
            top: -5px;
            width: 100%; /* Full width */
            height: 100%; /* Full height */
            overflow: auto; /* Enable scroll if needed */
            background-color: rgb(0,0,0); /* Fallback color */
            background-color: rgba(0,0,0,0.4); /* Black w/ opacity */
            filter: alpha(opacity=50);
        }

        /* Modal Content */
        .modal-content {
            background-color: #fefefe;
            margin: auto;
            padding: 20px;
            border: 1px solid #888;
            width: 50%;
            height: 160px;
        }

        /* The Close Button */
        .close {
            color: #aaaaaa;
            float: right;
            font-size: 28px;
            font-weight: bold;
        }

            .close:hover,
            .close:focus {
                color: #000;
                text-decoration: none;
                cursor: pointer;
            }
    </style>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <meta http-equiv="Pragma" content="no-cache">
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
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

        //function onExport()
        //{

        //    document.frmCustServList.target = 'new';
        //    document.frmCustServList.hdnExport.value = 'OSS';
        //    document.frmCustServList.submit();
        //    document.frmCustServList.hdnExport.value = 'OSS';
        //    document.frmCustServList.target = '_self';
        //}

        function onExport() {

            // Get the modal
            var modal = document.getElementById('myModal');

            modal.style.display = "block";
            document.getElementsByName("_txtStartDate")[0].value = "";
            document.getElementsByName("_txtEndDate")[0].value = "";
            // Get the <span> element that closes the modal
            var span = document.getElementsByClassName("close")[0];
            // When the user clicks on <span> (x), close the modal
            span.onclick = function () {
                var modal = document.getElementById('myModal');
                modal.style.display = "none";
            }

            return false;

        }
        //-->

        function OnExportFilter() {
            document.frmCustServList.target = 'new';
            var expElement = document.getElementsByName("hdnExport");
            expElement[0].value = 'OSS';
            document.getElementsByName("txtGoToPageNo")[0].value = "";
            document.frmCustServList.submit();
            // document.frmCustList.detailhdnexport.value = '';
            // document.frmCustList.target = '_self';

        }




        // When the user clicks anywhere outside of the modal, close it
        window.onclick = function (event) {
            var modal = document.getElementById('myModal');
            if (event.target == modal) {
                modal.style.display = "none";
            }
        }

        // End of script hiding -->
    </script>
</head>
<%
        Function SimpleBinaryToString(Binary)
    
    Dim DecryptedData
    if (Binary <> "" and Binary <> Empty) then
    
SimpleBinaryToString = DecryptWithKey("Constant",Binary)
    
    
    else
     SimpleBinaryToString=""
   
    end if
   
End Function
 dim aList, intPageNumber, intPageCount
 dim strCustomerServiceDesc, intSupportGroupID, strCustomerName, strServiceLocationName, strOrderNo
 dim strStatusCode, intCustomerServiceID, strRegionLcode, strServiceType, bolActiveOnly, bolPrefLangOnly, strLANG ,intNCServiceID
 dim strSQL, strWhereClause, strRecordStatus,strOrderBy,strMyWinName,strServiceEnd , strFromClause
 dim strSTypeTable, strLangPref, strLangWhere
 dim strServiceCity, strServiceAddress
 dim strLynxSeverity, intCustomerID, strCustomerShortName
 dim color
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
    intNCServiceID = trim(Request.Form("txtNCServiceID"))



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


	Dim strSTAttName, strSTAttValue,strSIAttName, strSIAttValue
	strSTAttName = Request.Form("txtSTAttName")
	strSTAttValue = Request.Form("txtSTAttValue")
	strSIAttName = Request.Form("txtSIAttName")
	strSIAttValue = Request.Form("txtSIAttValue")

	dim outcolumns
	outcolumns = 13
	if (trim(strSTAttName) <> "" and trim(strSTAttName) <> "%") or (trim(strSTAttValue) <> "" and trim(strSTAttValue) <> "%") then
		outcolumns = outcolumns + 2
	else
		'if (trim(strSTAttName) = "" and trim(strSTAttValue) = "" and trim(strSIAttName) = "%" and trim(strSIAttValue) = "%")  then
		if (trim(strSIAttName) = "%" and trim(strSIAttValue) = "%" ) then
				if trim(strSTAttName) <> "%" or trim(strSTAttValue) <> "%" then
		        	outcolumns = outcolumns + 2
		        end if
		end if
	end if

	if (trim(strSIAttName) <> "" and trim(strSIAttName) <> "%") or (trim(strSIAttValue) <> "" and trim(strSIAttValue) <> "%")  then
		outcolumns = outcolumns + 2
	else
		'if (trim(strSTAttName) = "%" and trim(strSTAttValue) = "%" and trim(strSIAttName) = "" and trim(strSIAttValue) = "")  then
		if (trim(strSTAttName) = "%" and trim(strSTAttValue) = "%" ) then
		    if (trim(strSIAttName) <> "%" or trim(strSIAttValue) <> "%") then
		     outcolumns = outcolumns + 2
		    end if
		end if
	end if

	IF (len(intCustomerID) = 0 and len(strCustomerShortName) = 0) then
    
		strSQL = "select  distinct(s.customer_service_id), " &_
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
                    "s.NC_customer_service_id, " &_
					strLangPref


		if (trim(strSTAttName) <> "" and trim(strSTAttName) <> "%") or (trim(strSTAttValue) <> "" and  trim(strSTAttValue) <> "%") then
			strSQL = strSQL + ",matt.SRVC_TYPE_ATT_NAME ,mattv.SRVC_TYPE_ATT_VAL_name"
		else
		    'if (trim(strSTAttName) = "" and trim(strSTAttValue) = "" and trim(strSIAttName) = "%" and trim(strSIAttValue) = "%")  then
		    if (trim(strSIAttName) = "%" and trim(strSIAttValue) = "%") then
		       if (trim(strSTAttName) <> "%" or trim(strSTAttValue) <> "%") then
		        'strSQL = strSQL + ",msiaatt.SRVC_INSTNC_ATT_NAME, msiaattv.SRVC_INSTNC_ATT_VAL  "
		        strSQL = strSQL + ",msiaatt.SRVC_INSTNC_ATT_NAME, decode(siavxref.SRVC_INSTNC_ATT_USR_DEF_VAL, NULL,  MSIAATTV.SRVC_INSTNC_ATT_VAL, siavxref.SRVC_INSTNC_ATT_USR_DEF_VAL) "
		       end if
			end if
		end if


		if (trim(strSIAttName) <> "" and trim(strSIAttName) <> "%") or (trim(strSIAttValue) <> "" and trim(strSIAttValue) <> "%")  then
			'strSQL = strSQL + ",msiaatt.SRVC_INSTNC_ATT_NAME, msiaattv.SRVC_INSTNC_ATT_VAL  "
		     strSQL = strSQL + ",msiaatt.SRVC_INSTNC_ATT_NAME, decode(siavxref.SRVC_INSTNC_ATT_USR_DEF_VAL, NULL,  MSIAATTV.SRVC_INSTNC_ATT_VAL, siavxref.SRVC_INSTNC_ATT_USR_DEF_VAL) "
		else
		   'if (trim(strSTAttName) = "%" and trim(strSTAttValue) = "%" and trim(strSIAttName) = "" and trim(strSIAttValue) = "")  then
		   if (trim(strSTAttName) = "%" and trim(strSTAttValue) = "%" ) then
		    if (trim(strSIAttName) <> "%" or trim(strSIAttValue) <> "%") then
		     strSQL = strSQL + ",matt.SRVC_TYPE_ATT_NAME ,mattv.SRVC_TYPE_ATT_VAL_name"
		    end if
		   end if
		end if




		strSQL = strSQL + " from crp.customer_service s, " &_
					"crp.customer c,  " &_
					"crp.service_location l, " &_
					"crp.v_remedy_support_group g," &_
					"crp.address f,  " &_
					strSTypeTable

       	if (trim(strSTAttName) <> "" and trim(strSTAttName) <> "%") or (trim(strSTAttValue) <> "" and  trim(strSTAttValue) <> "%") then
			strSQL = strSQL + ",CRP.SRVC_TYPE_ATT matt,CRP.SRVC_TYPE_ATT_val mattv,CRP.SRVC_TYPE_ATT_VAL_XREF xref,CRP.SRVC_TYPE_ATT_VAL_USAGE usage "
		else
			'if (trim(strSTAttName) = "" and trim(strSTAttValue) = "" and trim(strSIAttName) = "%" and trim(strSIAttValue) = "%")  then
			if (trim(strSIAttName) = "%" and trim(strSIAttValue) = "%")  then
			   if (trim(strSTAttName) <> "%" or trim(strSTAttValue) <> "%") then
				strSQL = strSQL + ", so.SRVC_INSTNC_ATT msiaatt, so.SRVC_INSTNC_ATT_val msiaattv " &_
			         ",SO.CUST_SRVC_INST_ATT_VAL_XREF siavxref, sO.SRVC_INSTNC_ATT_VAL_USAGE siavusage, SO.SRVC_INSTNC_ATT_XREF siaattxref "
			   end if
			end if
		end if



        if (trim(strSIAttName) <> "" and trim(strSIAttName) <> "%") or (trim(strSIAttValue) <> "" and trim(strSIAttValue) <> "%")  then
			strSQL = strSQL + ", so.SRVC_INSTNC_ATT msiaatt, so.SRVC_INSTNC_ATT_val msiaattv " &_
			         ",SO.CUST_SRVC_INST_ATT_VAL_XREF siavxref, sO.SRVC_INSTNC_ATT_VAL_USAGE siavusage, SO.SRVC_INSTNC_ATT_XREF siaattxref "
	    else
	    	'if (trim(strSTAttName) = "%" and trim(strSTAttValue) = "%" and trim(strSIAttName) = "" and trim(strSIAttValue) = "")  then
	    	if (trim(strSTAttName) = "%" and trim(strSTAttValue) = "%") then
	    	  if (trim(strSIAttName) <> "%" or trim(strSIAttValue) <> "%")  then
    			strSQL = strSQL + ",CRP.SRVC_TYPE_ATT matt,CRP.SRVC_TYPE_ATT_val mattv,CRP.SRVC_TYPE_ATT_VAL_XREF xref,CRP.SRVC_TYPE_ATT_VAL_USAGE usage "
    		  end if
			end if
		end if



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
                    "s.customer_service_id,s.NC_customer_service_id, " &_
					strLangPref

		if (trim(strSTAttName) <> "" and trim(strSTAttName) <> "%") or (trim(strSTAttValue) <> "" and  trim(strSTAttValue) <> "%") then
			strSQL = strSQL + ",matt.SRVC_TYPE_ATT_NAME ,mattv.SRVC_TYPE_ATT_VAL_name"
		else
		    'if (trim(strSTAttName) = "" and trim(strSTAttValue) = "" and trim(strSIAttName) = "%" and trim(strSIAttValue) = "%")  then
		    if (trim(strSIAttName) = "%" and trim(strSIAttValue) = "%")  then
		      if (trim(strSTAttName) <> "%" or trim(strSTAttValue) <> "%") then
				'strSQL = strSQL + ",msiaatt.SRVC_INSTNC_ATT_NAME,msiaattv.SRVC_INSTNC_ATT_VAL  "
		        strSQL = strSQL + ",msiaatt.SRVC_INSTNC_ATT_NAME, decode(siavxref.SRVC_INSTNC_ATT_USR_DEF_VAL, NULL,  MSIAATTV.SRVC_INSTNC_ATT_VAL, siavxref.SRVC_INSTNC_ATT_USR_DEF_VAL) "
			  end if
		    end if
		end if


        if (trim(strSIAttName) <> "" and trim(strSIAttName) <> "%") or (trim(strSIAttValue) <> "" and trim(strSIAttValue) <> "%")  then
			'strSQL = strSQL + ",msiaatt.SRVC_INSTNC_ATT_NAME,msiaattv.SRVC_INSTNC_ATT_VAL  "
	        strSQL = strSQL + ",msiaatt.SRVC_INSTNC_ATT_NAME, decode(siavxref.SRVC_INSTNC_ATT_USR_DEF_VAL, NULL,  MSIAATTV.SRVC_INSTNC_ATT_VAL, siavxref.SRVC_INSTNC_ATT_USR_DEF_VAL) "
		else
			'if trim(strSIAttName) = "" and trim(strSIAttValue) = "" and trim(strSTAttName) = "%" and trim(strSTAttValue) = "%" then
			if (trim(strSTAttName) = "%" and trim(strSTAttValue) = "%") then
			  if (trim(strSIAttName) <> "%" or trim(strSIAttValue) <> "%") then
				strSQL = strSQL + ",matt.SRVC_TYPE_ATT_NAME ,mattv.SRVC_TYPE_ATT_VAL_name"
			  end if
			end if
		end if


		strSQL = strSQL + " from crp.customer_service s,  " &_
					"crp.customer c,  " &_
					"crp.service_location l, " &_
					"crp.v_remedy_support_group g, " &_
					"crp.address f,  " &_
					"crp.customer_name_alias a, " &_
				    strSTypeTable
		if (trim(strSTAttName) <> "" and trim(strSTAttName) <> "%") or (trim(strSTAttValue) <> "" and  trim(strSTAttValue) <> "%") then
			strSQL = strSQL + ",CRP.SRVC_TYPE_ATT matt,CRP.SRVC_TYPE_ATT_val mattv,CRP.SRVC_TYPE_ATT_VAL_XREF xref,CRP.SRVC_TYPE_ATT_VAL_USAGE usage "
		else
		    'if (trim(strSTAttName) = "" and trim(strSTAttValue) = "" and trim(strSIAttName) = "%" and trim(strSIAttValue) = "%")  then
		    if (trim(strSIAttName) = "%" and trim(strSIAttValue) = "%")  then
		      if (trim(strSTAttName) <> "%" or trim(strSTAttValue) <> "%") then
				strSQL = strSQL + ", so.SRVC_INSTNC_ATT msiaatt, so.SRVC_INSTNC_ATT_val msiaattv " &_
			         ",SO.CUST_SRVC_INST_ATT_VAL_XREF siavxref, sO.SRVC_INSTNC_ATT_VAL_USAGE siavusage, SO.SRVC_INSTNC_ATT_XREF siaattxref "
			  end if
		    end if
		end if



        if (trim(strSIAttName) <> "" and trim(strSIAttName) <> "%") or (trim(strSIAttValue) <> "" and trim(strSIAttValue) <> "%")  then
			strSQL = strSQL + ", so.SRVC_INSTNC_ATT msiaatt, so.SRVC_INSTNC_ATT_val msiaattv " &_
			         ",SO.CUST_SRVC_INST_ATT_VAL_XREF siavxref, sO.SRVC_INSTNC_ATT_VAL_USAGE siavusage, SO.SRVC_INSTNC_ATT_XREF siaattxref "
		else
			'if trim(strSIAttName) = "" and trim(strSIAttValue) = "" and trim(strSTAttName) = "%" and trim(strSTAttValue) = "%" then
			if trim(strSTAttName) = "%" and trim(strSTAttValue) = "%" then
			  if trim(strSIAttName) <> "%" or trim(strSIAttValue) <> "%" then
				strSQL = strSQL + ",CRP.SRVC_TYPE_ATT matt,CRP.SRVC_TYPE_ATT_val mattv,CRP.SRVC_TYPE_ATT_VAL_XREF xref,CRP.SRVC_TYPE_ATT_VAL_USAGE usage "
			  end if
			end if
		end if



		strWhereClause = "where s.customer_id = c.customer_id " &_
						"and s.remedy_support_group_id = g.remedy_support_group_id(+) " &_
						"and s.service_type_id = t.service_type_id " &_
						"and s.service_location_id = l.service_location_id(+) " &_
						"and a.customer_id = c.customer_id " &_
						"and L.ADDRESS_ID = F.ADDRESS_ID(+) "
	end if


	IF  LEN(intCustomerID) > 0 THEN
      strWhereClause = strWhereClause & " AND c.customer_id =" & intCustomerID
	END IF

	IF  LEN(strCustomerShortName) > 0 THEN
      strWhereClause = strWhereClause & " AND Upper(c.customer_short_name)  LIKE '%" & routineOraString(strCustomerShortName) & "%' "
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

    IF  LEN(intNCServiceID) > 0 THEN
      strWhereClause = strWhereClause & " AND s.NC_customer_service_id =" & intNCServiceID
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

'	if bolPrefLangOnly = "YES" then
	strWhereClause = strWhereClause & strLangWhere
'	end if

	if (trim(strSTAttName) <> "" and trim(strSTAttName) <> "%") or (trim(strSTAttValue) <> "" and  trim(strSTAttValue) <> "%") then

		        'if trim(strSTAttName) <> "" then
		        if (trim(strSTAttName) <> "" and trim(strSTAttName) <> "%") then
		            'strWhereClause = strWhereClause + " and matt.SRVC_TYPE_ATT_NAME like  upper'%" & strSTAttName &"%'"
		            strWhereClause = strWhereClause + " and upper(matt.SRVC_TYPE_ATT_NAME) like upper('%'||trim('" & strSTAttName &" ')||'%')"

		        end if

		       ' if trim(strSTAttValue) <> "" then
		        if (trim(strSTAttValue) <> "" and  trim(strSTAttValue) <> "%") then
		            'strWhereClause = strWhereClause + " and mattv.SRVC_TYPE_ATT_VAL_NAME like '%" & strSTAttValue &"%'"
		            strWhereClause = strWhereClause + " and upper(mattv.SRVC_TYPE_ATT_VAL_NAME) like upper('%'||trim('" & strSTAttValue &" ')||'%')"

		        end if
		        strWhereClause = strWhereClause +  " and XREF.SERVICE_TYPE_ID = T.SERVICE_TYPE_ID " &_
		  						 "and USAGE.SRVC_TYPE_ATT_VAL_USAGE_ID = XREF.SRVC_TYPE_ATT_VAL_USAGE_ID "&_
		     					 "and mATT.SRVC_TYPE_ATT_ID  =USAGE.SRVC_TYPE_ATT_ID " &_
							     "and MATTV.SRVC_TYPE_ATT_VAL_ID = USAGE.SRVC_TYPE_ATT_VAL_ID " &_
		     					 "and MATT.RECORD_STATUS_IND='A' " &_
		     					 "and MATTV.RECORD_STATUS_IND='A' " &_
		     					 "and XREF.RECORD_STATUS_IND='A' " &_
		    				 	 "and USAGE.RECORD_STATUS_IND='A'"
	else
		'if (trim(strSTAttName) = "" and trim(strSTAttValue) = "" and trim(strSIAttName) = "%" and trim(strSIAttValue) = "%")  then
		if (trim(strSIAttName) = "%" and trim(strSIAttValue) = "%")  then
		   if trim(strSTAttName) <> "%" or trim(strSTAttValue) <> "%"  then
		    'strWhereClause = strWhereClause + " and upper(msiaatt.SRVC_INSTNC_ATT_NAME) like '%' and msiaattv.SRVC_INSTNC_ATT_VAL like '%'"
	         strWhereClause = strWhereClause + " and upper(msiaatt.SRVC_INSTNC_ATT_NAME) like '%' and " &_
	                          "decode(siavxref.SRVC_INSTNC_ATT_USR_DEF_VAL, NULL,  MSIAATTV.SRVC_INSTNC_ATT_VAL, siavxref.SRVC_INSTNC_ATT_USR_DEF_VAL) like '%' "

		    strWhereClause = strWhereClause + " and s.customer_service_id =siaVXREF.CUSTOMER_SERVICE_ID "&_
											" and siaVXREF.RECORD_STATUS_IND='A' and siaVUSAGE.RECORD_STATUS_IND='A' "&_
											" and siaVXREF.CUSTOMER_SERVICE_ID is not null "&_
											" and siaVUSAGE.SRVC_INSTNC_ATT_VAL_USAGE_ID = siaVXREF.SRVC_INSTNC_ATT_VAL_USAGE_ID "&_
											" and siaVUSAGE.SRVC_INSTNC_ATT_XREF_ID=siaATTXREF.SRVC_INSTNC_ATT_XREF_ID " &_
											" and siaATTXREF.RECORD_STATUS_IND='A' and siaATTXREF.SRVC_INSTNC_ATT_ID=msiaatt.SRVC_INSTNC_ATT_ID " &_
											" and SIAVUSAGE.SRVC_INSTNC_ATT_VALUE_ID = MSIAATTV.SRVC_INSTNC_ATT_VAL_ID "
		   end if
		end if
    end if


    if (trim(strSIAttName) <> "" and trim(strSIAttName) <> "%") or (trim(strSIAttValue) <> "" and trim(strSIAttValue) <> "%")  then
	      if trim(strSIAttName) <> "" then
	           ' strWhereClause = strWhereClause + " and msiaatt.SRVC_INSTNC_ATT_NAME like '%" & strSIAttName &"%'"
	        strWhereClause = strWhereClause + " and upper(msiaatt.SRVC_INSTNC_ATT_NAME) like upper('%'||trim('" & strSIAttName &" ')||'%')"
	      end if
	      if trim(strSIAttValue) <> "" then
	          '  strWhereClause = strWhereClause + " and msiaattv.SRVC_INSTNC_ATT_VAL like '%" & strSIAttValue &"%'"
	        'strWhereClause = strWhereClause + " and upper(msiaattv.SRVC_INSTNC_ATT_VAL) like upper('%'||trim('" & strSIAttValue &" ')||'%')"
	        strWhereClause = strWhereClause + " and upper(decode(siavxref.SRVC_INSTNC_ATT_USR_DEF_VAL, NULL,  MSIAATTV.SRVC_INSTNC_ATT_VAL, siavxref.SRVC_INSTNC_ATT_USR_DEF_VAL))" &_
	        			     " like upper('%'||trim('" & strSIAttValue &" ')||'%')"
	      end if
	      strWhereClause = strWhereClause + " and s.customer_service_id =siaVXREF.CUSTOMER_SERVICE_ID "&_
											" and siaVXREF.RECORD_STATUS_IND='A' and siaVUSAGE.RECORD_STATUS_IND='A' "&_
											" and siaVXREF.CUSTOMER_SERVICE_ID is not null "&_
											" and siaVUSAGE.SRVC_INSTNC_ATT_VAL_USAGE_ID = siaVXREF.SRVC_INSTNC_ATT_VAL_USAGE_ID "&_
											" and siaVUSAGE.SRVC_INSTNC_ATT_XREF_ID=siaATTXREF.SRVC_INSTNC_ATT_XREF_ID " &_
											" and siaATTXREF.RECORD_STATUS_IND='A' and siaATTXREF.SRVC_INSTNC_ATT_ID=msiaatt.SRVC_INSTNC_ATT_ID " &_
											" and SIAVUSAGE.SRVC_INSTNC_ATT_VALUE_ID = MSIAATTV.SRVC_INSTNC_ATT_VAL_ID "
	else
		  'if trim(strSIAttName) = "" and trim(strSIAttValue) = "" and trim(strSTAttName) = "%" and trim(strSTAttValue) = "%" then
		  if trim(strSTAttName) = "%" and trim(strSTAttValue) = "%" then
		    if trim(strSIAttName) <> "%" or trim(strSIAttValue) <> "%" then
		 	 strWhereClause = strWhereClause + " and upper(matt.SRVC_TYPE_ATT_NAME) like '%' and upper(mattv.SRVC_TYPE_ATT_VAL_NAME) like '%'"
			 strWhereClause = strWhereClause +  " and XREF.SERVICE_TYPE_ID = T.SERVICE_TYPE_ID " &_
		  						 "and USAGE.SRVC_TYPE_ATT_VAL_USAGE_ID = XREF.SRVC_TYPE_ATT_VAL_USAGE_ID "&_
		     					 "and mATT.SRVC_TYPE_ATT_ID  =USAGE.SRVC_TYPE_ATT_ID " &_
							     "and MATTV.SRVC_TYPE_ATT_VAL_ID = USAGE.SRVC_TYPE_ATT_VAL_ID " &_
		     					 "and MATT.RECORD_STATUS_IND='A' " &_
		     					 "and MATTV.RECORD_STATUS_IND='A' " &_
		     					 "and XREF.RECORD_STATUS_IND='A' " &_
		    				 	 "and USAGE.RECORD_STATUS_IND='A'"
		    end if

		  end if
    end if

	strOrderBy = " order by upper(s.customer_service_desc)"
    
    strFromClause = strSQL
	'join all pieces to make a complete query
	strsql = strSQL & strWhereClause & strRecordStatus & strOrderBy
    
	'Response.Write( strsql )       'display SQL for debugging
	'Response.end
	'Response.Write "strCustomerName ="
	'Response.Write (strCustomerName)
	'Response.Write "intCustomerID ="
	'Response.Write (intCustomerID)
	'Response.Write "strCustomerShortName="
	'Response.Write (strCustomerShortName)
	'Response.End



	Dim objRsResult,Recordcnt,strbgcolor
    'Response.Write(strSql)
	set objRsResult = objConn.Execute(strSql)
	if not objRsResult.EOF then
		aList = objRsResult.GetRows
	else
		Response.Write "0 records found"
		Response.end
	end if

   'release and kill the recordset and the connection objects
	objRsResult.Close
	set objRsResult = nothing

	'objConn.close
	'set objConn = nothing

   'calculate page number
	intPageCount = Int(UBound(aList, 2) / intConstDisplayPageSize) + 1

	select case Request("Action")

		case "<<"		intPageNumber = 1
		case "<"		color = Request("txtcolor")
		                intPageNumber = Request("txtPageNumber") - 1
					    if intPageNumber < 1 then intPageNumber = 1
		case ">"		color = Request("txtcolor")
						intPageNumber = Request("txtPageNumber") + 1
					    if intPageNumber > intPageCount then intPageNumber = intPageCount
		case ">>"	
    	intPageNumber = intPageCount
		case  else  
    
    if Request("hdnExport") = "OSS" then
    dim strSelect

    'strSelect = "SELECT null FEDERATED_ASSET_ID, NE.NETWORK_ELEMENT_NAME DEVICE_NAME, NENA.NETWORK_ELEMENT_NAME_ALIAS ALIAS_NAME, NEP.NETWORK_ELEMENT_PORT_NAME PORT_NAME, NEP.CTR_IN_ID CTR_IN, NEP.CTR_OUT_ID CTR_OUT, NEP.VTR_IN_ID VTR_IN, NEP.VTR_OUT_ID VTR_OUT, NEP.QOS_NAME QOS_NAME, NEP.ETR_IN_ID ETR_IN, NEP.ETR_OUT_ID ETR_OUT, LMSP.MGMT_SPACE_NAME MGMT_SPACE, NE.MANAGED_IP_ADDRESS MGMT_IP_ADDRESS, NEP.NETWORK_ELEMENT_PORT_IP CUST_IP_ADDRESS, NES.SNMP_STRING SNMP_STRING, NES.SNMP_V3_USERNAME SNMP_V3_USERNAME, NES.SNMP_V3_ENGINEID SNMP_V3_ENGINEID, NES.SNMP_V3_CONTEXT_NAME SNMP_V3_CONTEXT_NAME, LSSL.SNMP_SECURITY_LVL_NAME SNMP_SECURITY_LVL_NAME, LSAP.SNMP_AUTH_PROT_NAME SNMP_AUTH_PROT_NAME, LSPP.SNMP_PRIV_PROT_NAME SNMP_PRIV_PROT_NAME, NES.SNMP_V3_AUTH_KEY SNMP_V3_AUTH_KEY, NES.SNMP_V3_PRIV_KEY SNMP_V3_PRIV_KEY, NES.SNMP_PORT SNMP_PORT, LCS.CI_STATUS_VALUE CI_STATUS ,( select listagg(S.MGMT_SYSTEM_NAME,',')   within group (order by S.MGMT_SYSTEM_ID)  MGMT_SYSTEMS " &_
'" from CRP.NETWORK_ELEMENT_MGMT_SYS NEMS, CRP.LCODE_MGMT_SYSTEMS S " &_
 ' "where NEMS.MGMT_SYSTEM_ID = S.MGMT_SYSTEM_ID and  NEMS.NETWORK_ELEMENT_ID = ne.Network_Element_Id), MK.MAKE_DESC DEVICE_VENDOR, MDL.MODEL_DESC DEVICE_MODEL, NE.NETWORK_ELEMENT_TYPE_CODE TECHNOLOGY, t.SERVICE_TYPE_DESC SERVICE_TYPE, null REPORTING_PACKAGE, LTC.TENANT_NAME TENANT, C.CUSTOMER_NAME COMPANY_NAME, C.CUSTOMER_SHORT_NAME COMPANY_CODE, CO.ORGANIZATION_NAME ORGANIZATION_NAME, CO.ORGANIZATION_CODE ORGANIZATION_CODE, lc.COUNTRY_DESC COUNTRY, F.COUNTRY_LCODE COUNTRY_CODE, LPS.PROVINCE_STATE_NAME PROVINCE, F.PROVINCE_STATE_LCODE PROVINCE_CODE, F.MUNICIPALITY_NAME CITY, ML.CLLI_CODE CITY_CODE, SNC.SITE_NAME SITE, SNC.SITE_CODE SITE_CODE, null CUSTOM_1, null CUSTOM_2, null CUSTOM_3, null CUSTOM_4, null CUSTOM_5 "
    
    strSelect = "SELECT DISTINCT NULL FEDERATED_ASSET_ID, NE.NETWORK_ELEMENT_NAME DEVICE_NAME, NPNA.NETWORK_PORT_NAME_ALIAS ALIAS_NAME,"&_
    " NEP.NETWORK_ELEMENT_PORT_NAME PORT_NAME, LCI.CTR_IN_VALUE,     LCO.CTR_OUT_VALUE,  LVI.VTR_IN_VALUE,"&_
    "LVO.VTR_OUT_VALUE,       NEP.QOS_NAME QOS_NAME,        LEI.ETR_IN_VALUE,   "&_
    "     LEO.ETR_OUT_VALUE, LMSP.MGMT_SPACE_NAME MGMT_SPACE, NE.MANAGED_IP_ADDRESS MGMT_IP_ADDRESS, NEP.NETWORK_ELEMENT_PORT_IP CUST_IP_ADDRESS, "&_ 
    "NES.SNMP_STRING SNMP_STRING, NES.SNMP_V3_USERNAME SNMP_V3_USERNAME, NES.SNMP_V3_ENGINEID SNMP_V3_ENGINEID, " &_
    " NES.SNMP_V3_CONTEXT_NAME SNMP_V3_CONTEXT_NAME, LSSL.SNMP_SECURITY_LVL_NAME SNMP_SECURITY_LVL_NAME,"&_
    " LSAP.SNMP_AUTH_PROT_NAME SNMP_AUTH_PROT_NAME, LSPP.SNMP_PRIV_PROT_NAME SNMP_PRIV_PROT_NAME, NES.SNMP_V3_AUTH_KEY SNMP_V3_AUTH_KEY,"&_
    " NES.SNMP_V3_PRIV_KEY SNMP_V3_PRIV_KEY, NES.SNMP_PORT SNMP_PORT, LCS.CI_STATUS_NAME CI_STATUS,( select listagg(S.MGMT_SYSTEM_NAME,',')   within group (order by S.MGMT_SYSTEM_ID)  MGMT_SYSTEMS " &_
" from CRP.NETWORK_ELEMENT_MGMT_SYS NEMS, CRP.LCODE_MGMT_SYSTEMS S " &_
  "where NEMS.MGMT_SYSTEM_ID = S.MGMT_SYSTEM_ID and  NEMS.NETWORK_ELEMENT_Port_ID = nep.NETWORK_ELEMENT_Port_ID), MK.MAKE_DESC DEVICE_VENDOR, MDL.MODEL_DESC DEVICE_MODEL,LNPF.NE_PORT_FUNCTION_NAME, t.SERVICE_TYPE_DESC SERVICE_TYPE, NULL REPORTING_PACKAGE, LTC.TENANT_NAME TENANT, C.CUSTOMER_NAME COMPANY_NAME, C.CUSTOMER_SHORT_NAME COMPANY_CODE, CO.ORGANIZATION_NAME ORGANIZATION_NAME, CO.ORGANIZATION_CODE ORGANIZATION_CODE, lc.COUNTRY_DESC COUNTRY, F.COUNTRY_LCODE COUNTRY_CODE, LPS.PROVINCE_STATE_NAME PROVINCE, F.PROVINCE_STATE_LCODE PROVINCE_CODE, F.MUNICIPALITY_NAME CITY, ML.CLLI_CODE CITY_CODE, SNC.SITE_NAME SITE, SNC.SITE_CODE SITE_CODE, NULL CUSTOM_1, NULL CUSTOM_2, NULL CUSTOM_3, NULL CUSTOM_4, NULL CUSTOM_5 , NEP.VN_NAME VN_NAME "
    
    strFromClause =  " , " & Right(strFromClause, Len(strFromClause) - 6)

    strFromClause = strFromClause & " , CRP.NETWORK_ELEMENT ne, CRP.NETWORK_PORT_NAME_ALIAS NPNA, CRP.NETWORK_ELEMENT_PORT nep ,CRP.LCODE_MGMT_SPACE lmsp,   CRP.NETWORK_ELEMENT_SNMP nes,       CRP.LCODE_CI_STATUS lcs,       CRP.ASSET_CATALOGUE ac,CRP.MAKE mk,CRP.MODEL mdl,CRP.MANAGED_CORRELATION mc,  CRP.LCODE_TENANT_CODE ltc,CRP.CUSTOMER_ORGANIZATION co,CRP.LCODE_COUNTRY lc,CRP.LCODE_PROVINCE_STATE lps,    CRP.MUNICIPALITY_LOOKUP ml,CRP.SITE_NAME_CODE snc, CRP.LCODE_SNMP_SECURITY_LVL lssl, CRP.LCODE_SNMP_AUTH_PROT lsap, CRP.LCODE_SNMP_PRIV_PROT lspp  , crp.LCODE_CTR_IN LCI, crp.LCODE_CTR_OUT LCO,   crp.LCODE_ETR_IN LEI, crp.LCODE_ETR_OUT LEO,       crp.LCODE_VTR_IN LVI,       crp.LCODE_VTR_OUT LVO, CRP.LCODE_NE_PORT_FUNCTION LNPF "

    strWhereClause = strWhereClause & " and ne.customer_id = c.customer_Id  AND ne.Network_Element_Id = nep.Network_Element_Id(+) " &_
  " AND ne.MGMT_SPACE_ID = lmsp.MGMT_SPACE_ID(+)" &_
  " AND ne.Network_Element_Id = nes.Network_Element_Id(+)" &_
  " AND nep.CI_STATUS_ID = lcs.CI_STATUS_ID(+)" &_
  " AND ne.Asset_Catalogue_Id = ac.Asset_Catalogue_Id(+)" &_
  " AND ac.MAKE_ID = mk.MAKE_ID(+)" &_
  " AND ac.MODEL_ID = mdl.MODEL_ID(+)" &_
  " AND NE.NETWORK_ELEMENT_ID = MC.NETWORK_ELEMENT_ID(+)" &_ 
  "and MC.CUSTOMER_SERVICE_ID = s.CUSTOMER_SERVICE_ID" &_
  " AND s.SERVICE_TYPE_ID = t.SERVICE_TYPE_ID" &_
  " AND ne.TENANT_ID = ltc.TENANT_ID(+)" &_
  " AND nep.ORGANIZATION_ID = co.ORGANIZATION_ID(+)" &_
  " and s.SERVICE_LOCATION_ID = l.SERVICE_LOCATION_ID" &_
  "   and f.ADDRESS_ID = l.ADDRESS_ID" &_
  " and f.COUNTRY_LCODE = LC.COUNTRY_LCODE(+)" &_
  " and f.PROVINCE_STATE_LCODE = LPS.PROVINCE_STATE_LCODE" &_
  " and f.MUNICIPALITY_NAME = ML.MUNICIPALITY_NAME(+)" &_
  "   AND NEP.Site_id = snc.Site_id(+)" &_
  " AND NES.SNMP_SECURITY_LVL_ID = LSSL.SNMP_SECURITY_LVL_ID (+) " &_
  " AND NES.SNMP_AUTH_PROT_ID = LSAP.SNMP_AUTH_PROT_ID (+)" &_
  " AND NES.SNMP_PRIV_PROT_ID = lspp.SNMP_PRIV_PROT_ID (+)" &_
     "and  NEP.CTR_IN_ID =  lci.CTR_IN_ID(+)       and NEP.CTR_OUT_ID= LCO.CTR_OUT_ID(+)       and NEP.ETR_IN_ID=LEI.ETR_IN_ID(+)       and NEP.ETR_OUT_ID= LEO.ETR_OUT_ID(+)       and NEP.VTR_IN_ID = LVI.VTR_IN_ID(+)       and NEP.VTR_OUT_ID = LVO.VTR_OUT_ID(+)   and ML.COUNTRY_LCODE = F.COUNTRY_LCODE and ML.PROVINCE_STATE_LCODE = F.PROVINCE_STATE_LCODE AND NEP.NE_PORT_FUNCTION_LCODE = LNPF.NE_PORT_FUNCTION_LCODE(+) AND NEP.NETWORK_ELEMENT_PORT_ID = NPNA.NETWORK_ELEMENT_PORT_ID(+)" 


   ' if Request("hdnDate") <> "" then
   ' strWhereClause = strWhereClause + " AND TRUNC( c.create_date_time) >= to_date('"+ Request("hdnDate")  + "','DD/MM/YYYY')   "
    'AND TRUNC( c.create_date_time) >    TO_DATE ('25/07/2011', 'DD/MM/YYYY')
   ' end if

    if Request("_txtStartDate") <> "" then
    strWhereClause = strWhereClause + " AND TRUNC( NEP.update_date_time) >= to_date('"+ Request("_txtStartDate")  + "','DD/MM/YYYY')   "
    'AND TRUNC( c.create_date_time) >    TO_DATE ('25/07/2011', 'DD/MM/YYYY')
    end if

     if Request("_txtEndDate") <> "" then
    strWhereClause = strWhereClause + " AND TRUNC( NEP.update_date_time) <= to_date('"+ Request("_txtEndDate")  + "','DD/MM/YYYY')   "
    'AND TRUNC( c.create_date_time) >    TO_DATE ('25/07/2011', 'DD/MM/YYYY')
    end if


    strsql = strSelect & strFromClause & strWhereClause
    if strSQL <> "" then 
    strSQL = Replace(strSQL,"distinct(s.customer_service_id)","s.customer_service_id")
    end if
    'Dim objRsResult,Recordcnt,strbgcolor,detailaList
    Dim detailaList

	set objRsResult = objConn.Execute(strSql)
	if not objRsResult.EOF then
		detailaList = objRsResult.GetRows
	else
		Response.Write "0 records found"
		Response.end
	end if

   'release and kill the recordset and the connection objects
	objRsResult.Close
	set objRsResult = nothing

	objConn.close
	set objConn = nothing

    
	intPageCount = Int(UBound(detailaList,2) / intConstDisplayPageSize) + 1

                    dim detailstrRealUserID
						detailstrRealUserID = Session("username")

						'determine export path
						dim detailstrExportPath, detailliLength
						detailstrExportPath = Request.ServerVariables("PATH_TRANSLATED")


						While (Right(detailstrExportPath, 1) <> "\" And Len(detailstrExportPath) <> 0)
							detailliLength = Len(detailstrExportPath) - 1
							detailstrExportPath = Left(detailstrExportPath, detailliLength)
						Wend
						detailstrExportPath = detailstrExportPath & "export\"


						'create scripting object
						dim detailobjFSO, detailobjTxtStream
						set detailobjFSO = server.CreateObject("Scripting.FileSystemObject")
						'create export file (overwrite if exists)
						set detailobjTxtStream = detailobjFSO.CreateTextFile(detailstrExportPath&detailstrRealUserID&"-detailcustomer.xls", true, false)

						if err then
							DisplayError "CLOSE", "", err.Number, "CustList.asp - Cannot create Excel spreadsheet file due to the following reasons. Please contact your website administrator.", err.Description
						end if

						with detailobjTxtStream
							.WriteLine "<table border=1>"

							'export the header

							.WriteLine "<THEAD>"
   .WriteLine "<TH>FEDERATED_ASSET_ID</TH> " 
    .WriteLine "<TH>DEVICE_NAME</TH> " 
    .WriteLine "<TH>CUST_DEVICE_NAME</TH> "
     .WriteLine "<TH>SNC_DEVICE_NAME</TH> "
    .WriteLine "<TH>ALIAS_NAME</TH> " 
    .WriteLine "<TH>PORT_NAME</TH>"
    .WriteLine "<TH>CTR_IN</TH>"
    .WriteLine "<TH>CTR_Out</TH>"
    .WriteLine "<TH>VN_NAME</TH>"
.WriteLine "<TH>VTR_IN</TH>"
    .WriteLine "<TH>VTR_OUT</TH>"
    .WriteLine "<TH>QOS_NAME</TH>"
    .WriteLine "<TH>ETR_IN</TH>"
.WriteLine "<TH>ETR_OUT</TH>"
    .WriteLine "<TH>MGMT_SPACE</TH>"
    .WriteLine "<TH>MGMT_IP_ADDRESS</TH>"
    .WriteLine "<TH>CUST_IP_ADDRESS</TH>"
    .WriteLine "<TH>SNMP_STRING</TH>"
							.WriteLine "<TH>SNMP_V3_USERNAME</TH>"
							.WriteLine "<TH>SNMP_V3_ENGINEID</TH>"
							.WriteLine "<TH>SNMP_V3_CONTEXT_NAME</TH>"
    .WriteLine "<TH>SNMP_V3_SEC_LEVEL</TH>"
							.WriteLine "<TH>SNMP_V3_AUTH_PROTOCOL</TH>"
							.WriteLine "<TH>SNMP_V3_PRIV_PROTOCOL</TH>"
							.WriteLine "<TH>SNMP_V3_AUTH_KEY</TH>"
    .WriteLine "<TH>SNMP_V3_PRIV_KEY</TH>"
							.WriteLine "<TH>SNMP_PORT</TH>"
							.WriteLine "<TH>CI_STATUS</TH>"
                             .Writeline "<TH>MGMT_SYSTEM</TH>"
                            .WriteLine "<TH>DEVICE_VENDOR</TH>"
							.WriteLine "<TH>DEVICE_MODEL</TH>"
							
							.WriteLine "<TH>TECHNOLOGY</TH>"
                           .Writeline "<TH>SERVICE_TYPE</TH>"
                           .Writeline "<TH>REPORTING_PACKAGE</TH>"
							.WriteLine "<TH>TENANT</TH>"
                            .WriteLine "<TH>COMPANY_NAME</TH>"
                             .WriteLine "<TH>COMPANY_CODE</TH>"
                             .WriteLine "<TH>ORGANIZATION_NAME</TH>"
                             .WriteLine "<TH>ORGANIZATION_CODE</TH>"
                             .WriteLine "<TH>COUNTRY</TH>"
                             .WriteLine "<TH>COUNTRY_CODE</TH>"
                               .WriteLine "<TH>PROVINCE</TH>"
                               .WriteLine "<TH>PROVINCE_CODE</TH>"
                               .WriteLine "<TH>CITY</TH>"
                             .WriteLine "<TH>CITY_CODE</TH>"
                             .WriteLine "<TH>SITE</TH>"
                             .WriteLine "<TH>SITE_CODE</TH>"
                               .WriteLine "<TH>CUSTOM_1</TH>"
                                .WriteLine "<TH>CUSTOM_2</TH>"
                                 .WriteLine "<TH>CUSTOM_3</TH>"
                                .WriteLine "<TH>CUSTOM_4</TH>"
                                 .WriteLine "<TH>CUSTOM_5</TH>"
                             .WriteLine "</THEAD>"

							'export the body
							for k = 0 to UBound(detailaList, 2)
								.WriteLine "<TR>"
	.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(0,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(1,k))&"&nbsp;</TD>"
								 .WriteLine "<TD NOWRAP>&nbsp;</TD>"
    .WriteLine "<TD NOWRAP>&nbsp;</TD>"
    .WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(2,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(3,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(4,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(5,k))&"&nbsp;</TD>"
    .WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(49,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(6,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(7,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(8,k))&"&nbsp;</TD>"
    .WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(9,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(10,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(11,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(12,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(13,k))&"&nbsp;</TD>"	
    .WriteLine "<TD NOWRAP>"& SimpleBinaryToString(routineHtmlString(detailaList(14,k))) &"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&  SimpleBinaryToString(routineHtmlString(detailaList(15,k)))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"& SimpleBinaryToString(routineHtmlString(detailaList(16,k)))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"& SimpleBinaryToString(routineHtmlString(detailaList(17,k)))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(18,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(19,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(20,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"& SimpleBinaryToString(routineHtmlString(detailaList(21,k)))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"& SimpleBinaryToString(routineHtmlString(detailaList(22,k)))&"&nbsp;</TD>"
    .WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(23,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(24,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(25,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(26,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(27,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(28,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(29,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(30,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(31,k))&"&nbsp;</TD>"
    .WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(32,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(33,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(34,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(35,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(36,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(37,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(38,k))&"&nbsp;</TD>"
    .WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(39,k))&"&nbsp;</TD>"
    .WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(40,k))&"&nbsp;</TD>"
    .WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(41,k))&"&nbsp;</TD>"
    .WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(42,k))&"&nbsp;</TD>"
    .WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(43,k))&"&nbsp;</TD>"
    .WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(44,k))&"&nbsp;</TD>"
    .WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(45,k))&"&nbsp;</TD>"
    .WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(46,k))&"&nbsp;</TD>"
    .WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(47,k))&"&nbsp;</TD>"
    .WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(48,k))&"&nbsp;</TD>"
								.WriteLine "</TR>"
							next
							.WriteLine "</table>"
						end with
						detailobjTxtStream.Close	
						set detailobjTxtStream = Nothing
						set detailobjFSO = Nothing

						strsql = "<script type=""text/javascript"">document.location=""export/"&detailstrRealUserID&"-detailcustomer.xls"";</script>"
						Response.Write strsql
						Response.End
    
   elseif Request("hdnExport") <> "" then
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
							.WriteLine "<TR>"

							'export the header
							if (trim(strSTAttName) <> "" and trim(strSTAttName) <> "%") or (trim(strSTAttValue) <> "" and trim(strSTAttValue) <> "%") then
								.WriteLine "<TH>ST Name</TD>"
								.WriteLine "<TH>ST Value</TH></TD>"

                                if (trim(strSIAttName) <> "" and trim(strSIAttName) <> "%") or (trim(strSIAttValue) <> "" and trim(strSIAttValue) <> "%")  then
										.WriteLine "<TH>SIA Name</TD>"
										.WriteLine "<TH>SIA Value</TH></TD>"
								end if

							else
                               if (trim(strSIAttName) <> "" and trim(strSIAttName) <> "%") or (trim(strSIAttValue) <> "" and trim(strSIAttValue) <> "%")  then
 							 		.WriteLine "<TH>SIA Name</TD>"
									.WriteLine "<TH>SIA Value</TH></TD>"
							   else
							   		if (trim(strSIAttName) = "%" and trim(strSIAttValue) = "%")  then
			   							if (trim(strSTAttName) <> "%" or trim(strSTAttValue) <> "%") then
			   								.WriteLine "<TH>SIA Name</TD>"
											.WriteLine "<TH>SIA Value</TH></TD>"
			   							end if
			   						end if

							   end if
							end if

							if (trim(strSTAttName) = "%" and trim(strSTAttValue) = "%")  then
						   	  if trim(strSIAttName) <> "%" or trim(strSIAttValue) <> "%" then
									.WriteLine "<TH>ST Name</TD>"
									.WriteLine "<TH>ST Value</TH></TD>"
							  end if
							end if


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



							.WriteLine "<TD NOWRAP>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
                            .WriteLine "</TR>"

							'export the body
							for k = 0 to UBound(aList, 2)
								.WriteLine "<TR>"

    							if (trim(strSTAttName) <> "" and trim(strSTAttName) <> "%") or (trim(strSTAttValue) <> "" and trim(strSTAttValue) <> "%") then
						 			.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(14,k))&"</TD>"
									.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(15,k))&"</TD>"
                                    if (trim(strSIAttName) <> "" and trim(strSIAttName) <> "%") or (trim(strSIAttValue) <> "" and trim(strSIAttValue) <> "%")  then
							 				.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(16,k))&"</TD>"
											.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(17,k))&"</TD>"
									end if

								else
	                                if (trim(strSIAttName) <> "" and trim(strSIAttName) <> "%") or (trim(strSIAttValue) <> "" and trim(strSIAttValue) <> "%")  then
								 			.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(14,k))&"</TD>"
											.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(15,k))&"</TD>"
								    'end if
								    else
								      if (trim(strSIAttName) = "%" and trim(strSIAttValue) = "%")  then
			   							if (trim(strSTAttName) <> "%" or trim(strSTAttValue) <> "%") then
			   					   			.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(14,k))&"</TD>"
											.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(15,k))&"</TD>"
										end if
								      end if
								    end if
								end if
								if (trim(strSTAttName) = "%" and trim(strSTAttValue) = "%") then
								    if (trim(strSIAttName) <> "%" or trim(strSIAttValue) <> "%") then
								  			.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(14,k))&"</TD>"
											.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(15,k))&"</TD>"
								    end if
								end if

								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(1,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(2,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(3,k))&"</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(11,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(13,k))&"</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(4,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(5,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(12,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(9,k))&"</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(6,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(7,k))&"&nbsp;</TD>"



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
    <form method="post" name="frmCustServList" action="CustServList.asp">
        <input type="hidden" name="hdnDate" value="">
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

        <input type="hidden" name="txtSTAttName" value="<%=strSTAttName%>">
        <input type="hidden" name="txtSTAttValue" value="<%=strSTAttValue%>">
        <input type="hidden" name="txtSIAttName" value="<%=strSIAttName%>">
        <input type="hidden" name="txtSIAttValue" value="<%=strSIAttValue%>">

        <table border="1" cellpadding="2" cellspacing="0" width="100%">
            <thead>
                <tr>
                    <%
		'if (trim(strSTAttName) <> "" and trim(strSTAttName) <> "%") or (trim(strSTAttValue) <> "" and trim(strSTAttValue) <> "%") then
'			response.write "<TH align=left nowrap>ST Name</TH>"
'			response.write "<TH align=left nowrap>ST Value</TH>"
'           if (trim(strSIAttName) <> "" and trim(strSIAttName) <> "%") or (trim(strSIAttValue) <> "" and trim(strSIAttValue) <> "%")  then
'			  	 response.write "<TH align=left nowrap>SIA Name</TH>"
'			 	 response.write "<TH align=left nowrap>SIA Value</TH>"
'		    end if
'		 else
'           if (trim(strSIAttName) <> "" and trim(strSIAttName) <> "%") or (trim(strSIAttValue) <> "" and trim(strSIAttValue) <> "%")  then'
'			  	response.write "<TH align=left nowrap>SIA Name</TH>"
'			  	response.write "<TH align=left nowrap>SIA Value</TH>"
'		    end if
'		 end if
		if (trim(strSTAttName) <> "" and trim(strSTAttName) <> "%") or (trim(strSTAttValue) <> "" and trim(strSTAttValue) <> "%") then
			response.write "<TH align=left nowrap>ST Name</TH>"
			response.write "<TH align=left nowrap>ST Value</TH>"
		else

			'if (trim(strSTAttName) = "" and trim(strSTAttValue) = "" and trim(strSIAttName) = "%" and trim(strSIAttValue) = "%")  then
			if (trim(strSIAttName) = "%" and trim(strSIAttValue) = "%")  then
			  if (trim(strSTAttName) <> "%" or trim(strSTAttValue) <> "%") then
				response.write "<TH align=left nowrap>SIA Name</TH>"
				response.write "<TH align=left nowrap>SIA Value</TH>"
			  end if
			end if
		end if

	    if (trim(strSIAttName) <> "" and trim(strSIAttName) <> "%") or (trim(strSIAttValue) <> "" and trim(strSIAttValue) <> "%")  then
				response.write "<TH align=left nowrap>SIA Name</TH>"
				response.write "<TH align=left nowrap>SIA Value</TH>"
		else
	        	'if (trim(strSTAttName) = "%" and trim(strSTAttValue) = "%" and trim(strSIAttName) = "" and trim(strSIAttValue) = "")  then
	       	if (trim(strSTAttName) = "%" and trim(strSTAttValue) = "%") then
	        	if (trim(strSIAttName) <> "%" or trim(strSIAttValue) <> "%")  then
				  	response.write "<TH align=left nowrap>ST Name</TH>"
				  	response.write "<TH align=left nowrap>ST Value</TH>"
				end if
			end if
		end if


                    %>
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

                </tr>
            </thead>
            <tbody>
                <%
   'response.write "color" & color
    for k = m to n
	''Alternate row background colour
	'if Int(k/2) = k/2 then
	'	Response.write "<TR>"
	'else
	'	Response.write "<TR bgcolor=White>"
	'end if


	if (k=0) then
	 color=""
	 Response.write "<TR>"
	else
	  if (StrComp(aList(3,k),aList(3,k-1)) <> 0) then
	    if (color="") then
	    	color="White"
	    	Response.write "<TR bgcolor=White>"
	    else
	        color=""
	    	Response.write "<TR>"
	    end if

	  end if
	end if

	Response.write "<TR bgcolor=" &color & ">"



	if strMyWinName = "Popup" then
'		if (trim(strSTAttName) <> "" and trim(strSTAttName) <> "%") or (trim(strSTAttValue) <> "" and trim(strSTAttValue) <> "%") then
'				   Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(3,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(14, k))& "&nbsp;</a></TD>" &vbCrLf
'				   Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(3,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(15, k))& "&nbsp;</a></TD>" &vbCrLf
'                  if (trim(strSIAttName) <> "" and trim(strSIAttName) <> "%") or (trim(strSIAttValue) <> "" and trim(strSIAttValue) <> "%")  then
'						 	 Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(3,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(16, k))& "&nbsp;</a></TD>" &vbCrLf
'						 	 Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(3,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(17, k))& "&nbsp;</a></TD>" &vbCrLf
'				   end if
'        else

'        	if (trim(strSIAttName) <> "" and trim(strSIAttName) <> "%") or (trim(strSIAttValue) <> "" and trim(strSIAttValue) <> "%")  then
'			  Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(3,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(14, k))& "&nbsp;</a></TD>" &vbCrLf
'			  Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(3,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(15, k))& "&nbsp;</a></TD>" &vbCrLf
'			end if
'		end if
    if (trim(strSTAttName) <> "" and trim(strSTAttName) <> "%") or (trim(strSTAttValue) <> "" and trim(strSTAttValue) <> "%") then
		Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(3,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(14, k))& "&nbsp;</a></TD>" &vbCrLf
		Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(3,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(15, k))& "&nbsp;</a></TD>" &vbCrLf
        if (trim(strSIAttName) <> "" and trim(strSIAttName) <> "%") or (trim(strSIAttValue) <> "" and trim(strSIAttValue) <> "%")  then
			 Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(3,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(16, k))& "&nbsp;</a></TD>" &vbCrLf
			 Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(3,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(17, k))& "&nbsp;</a></TD>" &vbCrLf
		end if
 	else
		'if (trim(strSTAttName) = "" and trim(strSTAttValue) = "" and trim(strSIAttName) = "%" and trim(strSIAttValue) = "%")  then
		if (trim(strSIAttName) <> "" and trim(strSIAttName) <> "%") or (trim(strSIAttValue) <> "" and trim(strSIAttValue) <> "%")  then
			 Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(3,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(14, k))& "&nbsp;</a></TD>" &vbCrLf
			 Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(3,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(15, k))& "&nbsp;</a></TD>" &vbCrLf
		else
			if (trim(strSIAttName) = "%" and trim(strSIAttValue) = "%")  then
			   if (trim(strSTAttName) <> "%" or trim(strSTAttValue) <> "%") then
				  Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(3,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(14, k))& "&nbsp;</a></TD>" &vbCrLf
				  Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(3,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(15, k))& "&nbsp;</a></TD>" &vbCrLf
			   end if
			end if
		end if
	end if

   	if (trim(strSTAttName) = "%" and trim(strSTAttValue) = "%")  then
   	  if trim(strSIAttName) <> "%" or trim(strSIAttValue) <> "%" then
 		Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(3,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(14, k))& "&nbsp;</a></TD>" &vbCrLf
		Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(3,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(15, k))& "&nbsp;</a></TD>" &vbCrLf
	  end if
	end if


		Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(3,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(1, k))& "&nbsp;</a></TD>" &vbCrLf
		Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(3,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(2, k))& "&nbsp;</a></TD>" &vbCrLf
		Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(3,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(3, k))& "&nbsp;</a></style></TD>" &vbCrLf
		Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(3,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(11, k))& "&nbsp;</a></TD>" &vbCrLf
		Response.Write "<TD nowrap><a href=""#""  title=""Service Type Language Code"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(3,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(13, k))& "</a></TD>" &vbCrLf
		Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(3,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(4, k))& "&nbsp;</a></TD>" &vbCrLf
		Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(3,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(5, k))& "&nbsp;</a></TD>" &vbCrLf
		Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(3,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(12, k))& "&nbsp;</a></TD>" &vbCrLf
		Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(3,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(9, k))& "&nbsp;</a></TD>" &vbCrLf
		Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(3,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(6, k))& "&nbsp;</a></TD>" &vbCrLf
		Response.Write "<TD nowrap><a href=""#"" onclick=""return go_back('"& aList(9,k)&"','"&aList(10,k)&"','" &strServiceEnd& "', " &aList(3,k)& ", '" &routineJavascriptString(aList(1,k))& "', '"&routineJavascriptString(aList(5, k))& "', ' " &routineJavascriptString(aList(4, k))& "', ' " &routineJavascriptString(aList(8, k))& "' );"">" &routineJavascriptString(aList(7, k))& "&nbsp;</a></TD>" &vbCrLf

		Response.Write "</TR>"

	else
'		if (trim(strSTAttName) <> "" and trim(strSTAttName) <> "%") or (trim(strSTAttValue) <> "" and trim(strSTAttValue) <> "%") then
'		   Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServDetail.asp?CustServID="&aList(3,k)&""">"&routineHtmlString(aList(14,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
'		   Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServDetail.asp?CustServID="&aList(3,k)&""">"&routineHtmlString(aList(15,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
'          if (trim(strSIAttName) <> "" and trim(strSIAttName) <> "%") or (trim(strSIAttValue) <> "" and trim(strSIAttValue) <> "%")  then
'			      Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServDetail.asp?CustServID="&aList(3,k)&""">"&routineHtmlString(aList(16,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
'			      Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServDetail.asp?CustServID="&aList(3,k)&""">"&routineHtmlString(aList(17,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
'		   end if
'        else
'        	if (trim(strSIAttName) <> "" and trim(strSIAttName) <> "%") or (trim(strSIAttValue) <> "" and trim(strSIAttValue) <> "%")  then
'		      	Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServDetail.asp?CustServID="&aList(3,k)&""">"&routineHtmlString(aList(14,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
'		      	Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServDetail.asp?CustServID="&aList(3,k)&""">"&routineHtmlString(aList(15,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
'			end if
'		end if

    if (trim(strSTAttName) <> "" and trim(strSTAttName) <> "%") or (trim(strSTAttValue) <> "" and trim(strSTAttValue) <> "%") then
		Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServDetail.asp?CustServID="&aList(3,k)&""">"&routineHtmlString(aList(14,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
		Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServDetail.asp?CustServID="&aList(3,k)&""">"&routineHtmlString(aList(15,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
        if (trim(strSIAttName) <> "" and trim(strSIAttName) <> "%") or (trim(strSIAttValue) <> "" and trim(strSIAttValue) <> "%")  then
		   Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServDetail.asp?CustServID="&aList(3,k)&""">"&routineHtmlString(aList(16,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
		   Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServDetail.asp?CustServID="&aList(3,k)&""">"&routineHtmlString(aList(17,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
		end if
 	else
 	   	if (trim(strSIAttName) <> "" and trim(strSIAttName) <> "%") or (trim(strSIAttValue) <> "" and trim(strSIAttValue) <> "%")  then
		   Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServDetail.asp?CustServID="&aList(3,k)&""">"&routineHtmlString(aList(14,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
		   Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServDetail.asp?CustServID="&aList(3,k)&""">"&routineHtmlString(aList(15,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
		else

		'if (trim(strSTAttName) = "" and trim(strSTAttValue) = "" and trim(strSIAttName) = "%" and trim(strSIAttValue) = "%")  then
			if (trim(strSIAttName) = "%" and trim(strSIAttValue) = "%")  then
			  if (trim(strSTAttName) <> "%" or trim(strSTAttValue) <> "%") then
				Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServDetail.asp?CustServID="&aList(3,k)&""">"&routineHtmlString(aList(14,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
				Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServDetail.asp?CustServID="&aList(3,k)&""">"&routineHtmlString(aList(15,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
			  end if
			end if
		end if

	end if

    'if (trim(strSTAttName) = "%" and trim(strSTAttValue) = "%" and trim(strSIAttName) = "" and trim(strSIAttValue) = "")  then
    if (trim(strSTAttName) = "%" and trim(strSTAttValue) = "%") then
      if trim(strSIAttName) <> "%" or trim(strSIAttValue) <> "%"  then
		Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServDetail.asp?CustServID="&aList(3,k)&""">"&routineHtmlString(aList(14,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
		Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServDetail.asp?CustServID="&aList(3,k)&""">"&routineHtmlString(aList(15,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
	  end if
	end if








		Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServDetail.asp?CustServID="&aList(3,k)&""">"&routineHtmlString(aList(1,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
		Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServDetail.asp?CustServID="&aList(3,k)&""">"&routineHtmlString(aList(2,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
		Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServDetail.asp?CustServID="&aList(3,k)&""">"&routineHtmlString(aList(3,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
		Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServDetail.asp?CustServID="&aList(3,k)&""">"&routineHtmlString(aList(11,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
		Response.write "<TD nowrap><a TARGET=""_parent"" title=""Service Type Language Code"" href=""CustServDetail.asp?CustServID="&aList(3,k)&""">"&routineHtmlString(aList(14,k))&"</a>&nbsp;</TD>" &vbCrLf
		Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServDetail.asp?CustServID="&aList(3,k)&""">"&routineHtmlString(aList(4,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
		Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServDetail.asp?CustServID="&aList(3,k)&""">"&routineHtmlString(aList(5,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
		Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServDetail.asp?CustServID="&aList(3,k)&""">"&routineHtmlString(aList(12,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
		Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServDetail.asp?CustServID="&aList(3,k)&""">"&routineHtmlString(aList(9,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
		Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServDetail.asp?CustServID="&aList(3,k)&""">"&routineHtmlString(aList(6,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf
		Response.write "<TD nowrap><a TARGET=""_parent"" href=""CustServDetail.asp?CustServID="&aList(3,k)&""">"&routineHtmlString(aList(7,k))&"&nbsp;</a>&nbsp;</TD>" &vbCrLf





		Response.Write "</TR>"


	end if
next
                %>
            </tbody>
            <tfoot>
                <tr>
                    <td align="left" colspan="<%=outcolumns%>">
                        <input type="hidden" name="txtPageNumber" value="<%=intPageNumber%>">

                        <input type="hidden" name="txtcolor" value="<%=color%>">

                        <input type="submit" name="action" value="&lt;&lt;">
                        <input type="submit" name="action" value="&lt;">
                        <input type="text" name="txtGoToPageNo" title="You can jump to a specific page by typing the page number in this box" onclick="document.frmCustServList.txtGoToPageNo.value = ''" value="page <%=intPageNumber%> of <%=intPageCount%>" style="height: 22px; width: 150px">
                        <input type="submit" name="action" value="&gt;">
                        <input type="submit" name="action" value="&gt;&gt;">
                        <img src="images/excel.gif" onclick="document.frmCustServList.target='new'; document.frmCustServList.hdnExport.value='xls';document.frmCustServList.submit();document.frmCustServList.hdnExport.value='';document.frmCustServList.target='_self';" width="32" height="32">
                        <span>Service List Report</span>
                        <img src="images/excel.gif" style="padding-left: 400px;" onclick="onExport()" width="32" height="32">
                        <span>OSS CPE configuration extract</span>
                    </td>
                </tr>
            </tfoot>
            <caption>Records <%=m+1%> to <%=n+1%> of <%=UBound(aList, 2)+1 & " records"%></caption>
        </table>
        <div id="myModal" class="modal">

            <!-- Modal content -->
            <div class="modal-content">
                <span class="close">×</span>
                <div>Please enter the below dates in DD/MM/YYYY formats</div>
                <div>
                    Start Date :
                    <input type="text" name="_txtStartDate" />
                    End Date : 
                    <input type="text" name="_txtEndDate" />
                </div>
                <button onclick="OnExportFilter()">Submit</button>
            </div>

        </div>
    </form>
</body>
</html>
