<%@  language="VBScript" %>
<%Option Explicit
  on error resume next %>
<!--#include file="smaConstants.inc"-->
<!--#include file="smaProcs.inc"-->
<!--#include file = "databaseconnect.asp"-->
<!--

********************************************************************************************
* Page name:	CustServCriteria.asp
* Purpose:		To dynamically set the criteria to search for a Customer Service.
*				Results are displayed via CustServList.asp
*
* In Param:		This form reads following cookies :
*				 - WinName
*				 - CustomerName
*				 - CustomerService
*				 - ServiceEnd
*
* Created by:	Sara Sangha	Aug. 14th, 2000
* Edited by:    Adam Haydey Jan. 25th, 2000
*               Added Customer Service City and  Customer Service Address search fields.
********************************************************************************************
        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       28-Feb-02	 DTy		Add 'Alias' to 'Customer Service Name' field name.
       09-Sep-04     MW         Add  Repair Priority select list
   	   10-Aug-12     ACheung	Add Customer ID and Customer Shortname
       29-Jul-15   PSmith  Set Cookies in validation so the back key works
       05-Oct-15   PSmith  Only sumbit() for pop-up windows
       03-Feb-16   PSmith  Don't pre-populate search criteria
********************************************************************************************
-->
<%
dim intAccessLevel
dim objRsRegion ,objRsSupportGroup, objRsStatus, objRsRepairPriority, objRsCustomerName, strSQL
dim strCustomerService, strWinName,strServiceEnd, strServLocName, StrCustomerName, strLANG
Dim strServiceTypeID, strServiceTypeName, strCustomerID, strCustomerShortName

'Dim objRsSTNames, objRsSTValues, objRsSINames, objRsSIValues

intAccessLevel = CInt(CheckLogon(strConst_CustomerService))
if (intAccessLevel and intConst_Access_ReadOnly) <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to view customer service. Please contact your system administrator."
end if

	strLANG = Request.Cookies("UserInformation")("language_preference")
	if (Len(strLANG) = 0) then strLANG = "EN"

	strCustomerService = Request.Cookies("CustomerService")
	strWinName	= Request.Cookies("WinName")
    strServiceEnd = Request.Cookies("ServiceEnd")
    strServLocName = Request.Cookies("ServLocName")
    'Response.Write "end=" & strServiceEnd
	strCustomerName = Request.Cookies("CustomerName")
	strCustomerServiceId = Request.Cookies("CustomerServiceID")
	strServiceTypeID = Request.Cookies("ServiceTypeID")
	strServiceTypeName = Request.Cookies("ServiceTypeName")

	strCustomerID = Request.Cookies("hdnCustomerID")
	strCustomerShortName = Request.Cookies("CustomerShortName")

	'get a list of region codes
	strSQL = "select noc_region_lcode, " &_
					"noc_region_desc " &_
			 "from crp.lcode_noc_region " &_
			 "where record_status_ind = 'A' " &_
			 "order by noc_region_desc"

	set objRsRegion = objConn.Execute(strSQL)

	'get a list of service status codes
	strSQL = "SELECT service_status_code, " &_
					"service_status_name " &_
			 "FROM crp.service_status " &_
			 "WHERE record_status_ind = 'A' " &_
			 "order by service_status_name "

	set objRsStatus = objConn.Execute(strSQL)

	'get a list of support groups
	strSQL = "SELECT remedy_support_group_id, " &_
					"group_name " &_
			  "FROM crp.v_remedy_support_group " &_
			  "order by group_name"

	set objRsSupportGroup = objConn.Execute(strSQL)


	'get the LYNX repair priority list

	strSQL = "SELECT lynx_def_sev_lcode " &_
	         ",      lynx_def_sev_desc " &_
			 "FROM crp.lcode_lynx_def_sev " &_
			 "WHERE record_status_ind='A' "  &_
			 "ORDER BY lynx_def_sev_lcode"

	set objRsRepairPriority = objConn.Execute(strSQL)

%>

<html>
<head>
    <link rel="stylesheet" type="text/css" href="stylesheets/styles.css">
    <title>Customer Service Search</title>
    <meta name="GENERATOR" content="Microsoft FrontPage 12.0">
    <meta http-equiv="Pragma" content="no-cache">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <script type="text/javascript" src="GeneralJavaFunctions.js"></script>
    <script type="text/javascript" src="AccessLevels.js"></script>
    <script type="text/javascript">
        //****************************************** Java Functions *****************************
        //set section title
        var intAccessLevel = "<%=intAccessLevel%>";
        setPageTitle("SMA - Customer Service");


        function validate(theForm) {

            var bolConfirm;
            //Dec 13 2013 add STA SIA validation
            var sta, sia;
            sta = theForm.txtSTAttName.value;
            if (isWhitespace(sta)) {
                sta = theForm.txtSTAttValue.value;
            }
            sia = theForm.txtSIAttName.value;
            if (isWhitespace(sia)) {
                sia = theForm.txtSIAttValue.value;
            }
            //alert ("sia is " + sia);
            //return false;
            if (isWhitespace(sta) || isWhitespace(sia)) {
                if (isWhitespace(theForm.txtCustomerServiceDesc.value) && isWhitespace(theForm.selSupportGroup.value)
                   && isWhitespace(theForm.selRepairPriority.value)
                    && isWhitespace(theForm.txtCustomerName.value) && isWhitespace(theForm.selStatus.value)
                    && isWhitespace(theForm.txtCustomerServiceID.value) && isWhitespace(theForm.selRegion.value)
                    && isWhitespace(theForm.txtServiceLocationName.value) && isWhitespace(theForm.txtOrderNo.value)
                    && isWhitespace(theForm.txtServiceType.value) && isWhitespace(theForm.txtServiceCity.value)
                    && isWhitespace(theForm.txtServiceAddress.value) && isWhitespace(theForm.txtCustomerShortName.value)
                    && isWhitespace(theForm.txtCustomerID.value ) && isWhitespace(theForm.txtNCServiceID.value)) {
                    bolConfirm = window.confirm("No Search Criteria have been entered. This search may take a long time..Continue?");
                    if (!bolConfirm) {
                        return false;
                    }
                }

                if (isNaN(Number(theForm.txtCustomerServiceID.value))) {
                    alert("Customer Service ID must be a number");
                    document.frmCSSearch.txtCustomerServiceID.focus();
                    document.frmCSSearch.txtCustomerServiceID.select();
                    return (false);
                }
                else if (isNaN(Number(theForm.txtCustomerID.value))) {
                    alert("Customer ID must be a number");
                    document.frmCSSearch.txtCustomerID.focus();
                    document.frmCSSearch.txtCustomerID.select();
                    return (false);
                }
            }
            else {
                alert("Please perform queries on Service Type Attributes OR Service Instance Attributes, but not both. Correct your query and try again.");
                document.frmCSSearch.txtCustomerServiceID.focus();
                document.frmCSSearch.txtCustomerServiceID.select();
                return (false);
            }

            // Start thinking
            thinking(parent.fraResult);

            return true;

        }

        function btnAdd_onclick() {

            if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
                alert('Access denied. Please contact your system administrator.');
                return;
            }

            //	parent.location = 'CustServDetail.asp?CustServID=0'
            parent.location = 'CustServDetail.asp?CustServID=0&ServiceTypeID=0'
        }

        function fct_lookupCustomer(CustService) {

            //SetCookie("ServiceEnd", CustService);
            var strCustomerName = document.frmCSSearch.txtCustomerName.value;

            //alert (CustService);

            if (CustService != "") {
                SetCookie("ServiceEnd", CustService);
                //alert (CustService);
            }

            if (strCustomerName != "") {
                SetCookie("CustomerName", strCustomerName);
                //alert (strCustomerName);
            }

            /*if (document.frmCSSearch.txtCustomerName.value != "")
                {SetCookie("CustomerName", document.frmCSSearch.txtCustomerName.value);
                }*/

            SetCookie("WinName", 'Popup');
            window.open('SearchFrame.asp?fraSrc=Cust', 'Popup', 'top=50, left=100, height=600, width=800');
            //document.frmCSSearch.txtCustomerName.value = txtCustomerName ;
        }

        function window_onload() {
            //***********************************************************************************************
            // Function: window_onload																		*
            //																								*
            // Purpose:		To submit the form automatically when txtCustomerName has a value.				*
            //				This will happen when this page is called by a lookup function or by the Quick	*
            //			    Navigation drop-down box, which had already saved some values in the cookies	*
            //				and this form's VBScript has read those values and put in the form controls.	*
            //																								*
            // Created By:	Sara Sangha Aug. 25th 2000														*
            //																								*
            // Updated By:																					*
            //***********************************************************************************************
            var strCustomerService = document.frmCSSearch.txtCustomerServiceDesc.value;
            var strCustomerName = window.frmCSSearch.txtCustomerName.value;
            var strServiceLocationName = window.frmCSSearch.txtServiceLocationName.value;
            var intCustomerServiceID = document.frmCSSearch.txtCustomerServiceID.value;
            var strServiceTypeName = document.frmCSSearch.txtServiceType.value;

            var strSTAttName = document.frmCSSearch.txtSTAttName.value;
            var strSTAttValue = document.frmCSSearch.txtSTAttValue.value;
            var strSIAttName = document.frmCSSearch.txtSIAttName.value;
            var strSIAttValue = document.frmCSSearch.txtSIAttValue.value;
            var strWinName = document.frmCSSearch.hdnWinName.value;
            var strNCServiceId = document.frmCSSearch.txtNCServiceID.value;
            DeleteCookie("WinName");
            DeleteCookie("CustomerName");
            DeleteCookie("CustomerService");
            DeleteCookie("ServiceEnd");
            DeleteCookie("ServLocName");
            DeleteCookie("CustomerServiceID");
            DeleteCookie("ServiceTypeID");
            DeleteCookie("ServiceTypeName");
            DeleteCookie("hdnCustomerID");
            DeleteCookie("CustomerShortName");


            if (strWinName == "Popup" && ((intCustomerServiceID != "") || (strCustomerName != "") || (strCustomerService != "") || (strServiceTypeName != ""))) {
                SetCookie("CustomerName", document.frmCSSearch.txtCustomerName.value);
                SetCookie("CustomerServiceID", document.frmCSSearch.txtCustomerServiceID.value);
                SetCookie("CustomerService", document.frmCSSearch.txtCustomerServiceDesc.value);
                SetCookie("ServiceTypeName", document.frmCSSearch.txtServiceType.value);
                thinking(parent.fraResult);
                document.frmCSSearch.submit();
            }
        }

        function btnClear_onclick() {
            document.frmCSSearch.txtCustomerName.value = "";
            document.frmCSSearch.selStatus.selectedIndex = 0;
            document.frmCSSearch.txtCustomerServiceDesc.value = "";
            document.frmCSSearch.selRegion.selectedIndex = 0;
            document.frmCSSearch.txtServiceLocationName.value = "";
            document.frmCSSearch.txtCustomerServiceID.value = "";
            document.frmCSSearch.txtServiceCity.value = "";
            document.frmCSSearch.txtServiceAddress.value = "";
            document.frmCSSearch.selSupportGroup.selectedIndex = 0;
            document.frmCSSearch.txtOrderNO.value = "";
            document.frmCSSearch.txtServiceType.value = "";
            document.frmCSSearch.selRepairPriority.selectedIndex = 0;
            document.frmCSSearch.chkActiveOnly.checked = true;
            document.frmCSSearch.chkPrefLangOnly.checked = true;
            document.frmCSSearch.txtCustomerShortName.value = "";
            document.frmCSSearch.txtNCServiceID.value = "";
            document.frmCSSearch.txtCustomerID.value = "";

            document.frmCSSearch.txtSIAttName.value = "";
            document.frmCSSearch.txtSIAttValue.value = "";
            document.frmCSSearch.txtSTAttName.value = "";
            document.frmCSSearch.txtSTAttValue.value = "";
        }


        //-->
    </script>
</head>
<body language="javascript" onload="return window_onload()">

    <form name="frmCSSearch" method="post" action="CustServList.asp" target="fraResult" onsubmit="return validate(this)">
        <input type="hidden" id="hdnServiceTypeID" name="hdnServiceTypeID" value="<%=strServiceTypeID%>">
        <table>
            <thead>
                <tr>
                    <td colspan="4" align="left">
                    Customer Service Search
                </tr>
                </td>
            </thead>
            <tr>
                <td width="15%" align="right" nowrap>Customer Service ID</td>
                <td>
                    <input id="txtCustomerServiceID" name="txtCustomerServiceID" tabindex="1" style="height: 22px; width: 200px" value="<%=strCustomerServiceId%>"></td>

                <td width="15%" align="right" nowrap>Customer Name<font color="red"></font></td>
                <td colspan="2" nowrap align="left">
                    <input name="txtCustomerName" tabindex="2" type="text" style="width: 250px; height: 26px; color: silver;" readonly maxlength="50"
                        value="<%=unescape(strCustomerName)%>" align="left"><%if strCustomerName <> null then Response.write routineHTMLString(strCustomerName) else Response.Write null end if%>
                    <input name="btnCustomerLookup" type="button" value="..." language="javascript" onclick="fct_lookupCustomer('E')"></td>
            </tr>
            <tr>
                <td width="15%" align="right" nowrap>Customer Service Name/Alias</td>
                <td width="20%">
                    <input id="txtCustomerServiceDesc" name="txtCustomerServiceDesc" tabindex="2" style="height: 22px; width: 270px" value="<%=unescape(strCustomerService)%>">
                <td width="15%" align="right" nowrap>Customer Short Name</td>
                <td width="20%" colspan="2" nowrap align="left">
                    <input id="txtCustomerShortName" name="txtCustomerShortName" tabindex="3" style="height: 22px; width: 270px" value="<%=strCustomerShortName%>"></td>
            </tr>
            <tr>
                <td width="15%" align="right" nowrap>Service Type</td>
                <td width="20%">
                    <input id="txtServiceType" name="txtServiceType" tabindex="6" style="height: 22px; width: 270px" value="<%=strServiceTypeName%>"></td>
                <td width="15%" align="right" nowrap>Customer ID</td>
                <td>
                    <input id="txtCustomerID" name="txtCustomerID" tabindex="4" style="height: 22px; width: 200px" value="<%=strCustomerID%>"></td>
            </tr>
            <tr>
                <td width="15%" align="right" nowrap>Service Location Name</td>
                <td width="20%">
                    <input id="txtServiceLocationName" name="txtServiceLocationName" tabindex="3" style="height: 22px; width: 270px" value="<%=strServLocName%>"></td>
                <td width="15%" align="right" nowrap>Status</td>
                <td>
                    <select id="selStatus" name="selStatus" tabindex="8" style="height: 22px; width: 200px">
                        <option value=" ">
                        <option value="AllExceptTerm" <% if ( strCustomerName = "" and strServLocName = "" and strCustomerService = "" ) then Response.write " selected " end if %>>All Except Terminated</option>
                        <% Do while not objRsStatus.EOF %>
                        <option value="<%=objRsStatus(0)%>"><%=objRsStatus(1)%>
                            <%objRsStatus.MoveNext%>
                            <%Loop%>
                    </select></td>
            </tr>
            <tr>
                <td width="15%" align="right" nowrap>Service Address </td>
                <td width="20%">
                    <input id="txtServiceAddress" name="txtServiceAddress" tabindex="4" style="height: 22px; width: 270px" value="<%=strServAddress%>"></td>
                <td colspan="1" align="right" nowrap>Region </td>
                <td>
                    <select id="selRegion" name="selRegion" tabindex="7" style="height: 22px; width: 200px">
                        <option value=" " selected>
                            <% Do while not objRsRegion.EOF %>
                        <option value="<%=objRsRegion(0)%>"><%=objRsRegion(1)%>
                            <%objRsRegion.MoveNext%>
                            <%Loop%>
                    </select></td>
            </tr>
            <tr>
                <td width="15%" align="right" nowrap>Support Group</td>
                <td width="20%">
                    <select id="selSupportGroup" name="selSupportGroup" style="height: 22px; width: 270px" tabindex="5">
                        <option value=" " selected>
                            <% Do while not objRsSupportGroup.EOF %>
                        <option value="<%=objRsSupportGroup(0)%>"><%=objRsSupportGroup(1)%>
                            <%objRsSupportGroup.MoveNext%>
                            <%Loop%>
                    </select></td>
                <td width="15%" align="right" nowrap>Service Location City</td>
                <td>
                    <input id="txtServiceCity" name="txtServiceCity" tabindex="10" style="height: 22px; width: 200px" value="<%=strServCity%>"></td>
            </tr>
            <tr>
                <td width="15%" align="right" nowrap>Repair Priority</td>
                <td width="20%">
                    <select id="selRepairPriority" name="selRepairPriority" style="height: 22px; width: 160px" tabindex="6">
                        <option selected value="">&nbsp;</option>
                        <%
			while not objRsRepairPriority.EOF
				Response.Write "<option value=" & routineHtmlString(objRsRepairPriority(0)) & ">"
				Response.Write routineHtmlString(objRsRepairPriority(1))
				Response.Write "</option>"
				objRsRepairPriority.MoveNext
			wend
			objRsRepairPriority.Close
                        %>
                    </select></td>
                <td width="15%" align="right" nowrap>Order No.</td>
                <td>
                    <input id="txtOrderNO" name="txtOrderNo" tabindex="11" style="height: 22px; width: 200px"></td>
            </tr>


            <tr>
                <td width="15%" align="right" nowrap>Service Type Attribute (STA) Name</td>
                <td width="20%">
                    <input id="txtSTAttName" name="txtSTAttName" tabindex="6" style="height: 22px; width: 270px" value="<%strSTAttName%>"></td>
                <td width="15%" align="right" nowrap>Service Instance Attribute (SIA) Name</td>
                <td width="20%">
                    <input id="txtSIAttName" name="txtSIAttName" tabindex="6" style="height: 22px; width: 270px" value="<%strSIAttName%>"></td>

            </tr>

            <tr>
                <td width="15%" align="right" nowrap>Service Type Attribute (STA) Value</td>
                <td width="20%">
                    <input id="txtSTAttValue" name="txtSTAttValue" tabindex="6" style="height: 22px; width: 270px" value="<%strSTAttValue%>"></td>

                <td width="15%" align="right" nowrap>Service Instance Attribute (SIA) Value</td>
                <td width="20%">
                    <input id="txtSIAttValue" name="txtSIAttValue" tabindex="6" style="height: 22px; width: 270px" value="<%strSIAttValue%>"></td>

            </tr>




            <tr>
                <td width="15%" align="right" nowrap>NC Service ID</td>
                <td>
                    <input id="txtNCServiceID" name="txtNCServiceID" tabindex="1" style="height: 22px; width: 200px" value="<%=strNCServiceId%>"></td>

                <td width="15%" align="right" nowrap></td>

                <td align="left" nowrap>Active&nbsp;Only&nbsp;<input id="chkActiveOnly" name="chkActiveOnly" tabindex="12" type="checkbox" value="YES" checked style="height: 24px; width: 24px">&nbsp;&nbsp;Pref'd Lang Only&nbsp;<input id="chkPrefLangOnly" name="chkPrefLangOnly" tabindex="12" type="checkbox" value="YES" checked style="height: 24px; width: 24px"></td>
            </tr>
            <tr>
                <td width="15%" align="right" nowrap></td>
                <td></td>
                <td></td>
                <td>
                    <nobr>
		<% if strWinName <> "Popup" then %>
			<INPUT id=btnAdd name=btnAdd type=button value=New style="width: 2cm"   LANGUAGE=javascript onclick="return btnAdd_onclick()">
		<% end if %>
			<INPUT id=btnClear name=btnClear type=button value=Clear style="width: 2cm"  LANGUAGE=javascript onclick="return btnClear_onclick()">
			<INPUT id=btnSearch name=btnSearch type=submit value=Search style="width: 2cm" > </nobr>
                </td>
            </tr>
            <tr>
                <td>
                    <input id="hdnWinName" name="hdnWinName" type="hidden" value="<%=strWinName%>"></td>
                <td>
                    <input id="hdnServiceEnd" name="hdnServiceEnd" type="hidden" value="<%=strServiceEnd%>"></td>
            </tr>
        </table>
    </form>
    <%
        Response.Write "Session: " & Session(strConst_Logon)
    %>
</body>
</html>
