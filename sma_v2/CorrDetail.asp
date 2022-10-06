<%@  language="VBSCRIPT" %>
<%
option explicit
on error resume next
%>
<!-- #include file="smaConstants.inc" -->
<!-- #include file="smaProcs.inc" -->
<!-- #include file="databaseConnect.asp" -->
<!--
***************************************************************************************************
* Name:		CustServDetail.asp i.e. Customer Service List
*
* Purpose:	This page displays information about a customer service and allows the user to update it
*
* Created By:	Sara Sangha 08/01/00
***************************************************************************************************

        Date		Author			Projects/enhancements made
        -----		------		------------------------------------------------------
	27-Nov-01 DTy                     Add number of seats.
	19-Oct-04	ACheung           Add repair priority
	21-OCt-04	MW		  CRC  Change on support group
	11-Apr-05 	MW             	  ASF  Grey out fields for NetCracker Items
	13-Mar-06	ACheung  	  ASF  Reenable the correlation with Quebec NOC and provincial code of QC
				 	  ASF  Reenable the +Facility button
	31-Aug-07	ACheung		  Display Usage Billing for provisioners
***************************************************************************************************
-->
<%
'********************************************************************************************
'* Page name:		CorrDetail.asp
'* Purpose:			To display the elements correlated to a customer service.
'* Created by:		Daniel Nica
'* Last updated by: Nancy Mooney 11/07/2000 - added Expand button which opens CorrReport.asp
'*
'********************************************************************************************
'check user's rights
dim intAccessLevel
intAccessLevel = CInt(CheckLogon(strConst_CorrelationCustomer))
if intAccessLevel and intConst_Access_ReadOnly <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to correlation management. Please contact your system administrator"
end if
    
'get input parameter: customer service id
dim strCustomerServiceID, strServiceTypeID, ServiceTypeID
strCustomerServiceID = Request("CustomerServiceID")

    if strCustomerServiceID = "" or IsEmpty( strCustomerServiceID ) then
    strCustomerServiceID = 0
    end if
strServiceTypeID = Request("ServiceTypeID")
'get real userid
dim strRealUserID
strRealUserID = Session("username")

if err then
	'unexpected error
	DisplayError "BACK", "", err.number, "UNEXPECTED ERROR", err.description
end if

'field disabler default to not disabled
dim strNCid
dim strDisable

strDisable = ""
strNCid = "netcracker"
'strNCid = "CMoore"

dim strWinMessage
    
'is that a save request?
if Request("txtFrmAction") = "SAVE" then
 	if (intAccessLevel and intConst_Access_Update <> intConst_Access_Update) then
		DisplayError "BACK", "", 0, "UPDATE DENIED", "You don't have access to update managed correlation. Please contact your system administrator"
	end if

	'create command object for update stored proc
	dim cmdUpdateObj
	set cmdUpdateObj = server.CreateObject("ADODB.Command")
	set cmdUpdateObj.ActiveConnection = objConn
	cmdUpdateObj.CommandType = adCmdStoredProc
	cmdUpdateObj.CommandText = "sma_sp_userid.spk_sma_cserv_inter.sp_cs_corr_update"
	'create params
	cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_userid", adVarChar, adParamInput, 20, strRealUserID)
	cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_customer_service_id", adNumeric , adParamInput,, CLng(Request("CustomerServiceID")))
	cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_remedy_group", adVarChar, adParamInput, 15, Request("selSupportGroup"))

	cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_service_status", adVarChar, adParamInput, 15, Request("selStatus"))
	dim chkCreateServiceTag, chkCheckServiceDep
	if Request("chkCreateService") <> "" then chkCreateServiceTag = "Y" else chkCreateServiceTag = "N"
	if Request("chkDependency") <> "" then chkCheckServiceDep = "Y" else chkCheckServiceDep = "N"
	cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_create_service_tag", adVarChar,adParamInput, 1, chkCreateServiceTag)
	cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_check_service", adVarChar, adParamInput, 1, chkCheckServiceDep)
	dim strDatePassed
	if Request("selday2") = "" then strDatePassed = "" else strDatePassed = Request("selmonth2")&"/"&Request("selday2")&"/"&Request("selyear2")
	cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_date_inservice", adVarChar, adParamInput, 10, strDatePassed)
	if Request("selday3") = "" then strDatePassed = "" else strDatePassed = Request("selmonth3")&"/"&Request("selday3")&"/"&Request("selyear3")
	cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_date_terminated", adVarChar, adParamInput, 10, strDatePassed)
	if Request("selday") = "" then strDatePassed = "" else strDatePassed = Request("selmonth")&"/"&Request("selday")&"/"&Request("selyear")
	cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_billing_date", adVarChar, adParamInput, 10, strDatePassed)
	cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_comments", adVarChar, adParamInput, 2000, Request("txtComment"))

	IF	Request("txtNoOfSeats") <> "" THEN
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_No_Of_Seats", adNumeric, adParamInput, , clng(Request("txtNoOfSeats")))
	ELSE
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_No_Of_Seats", adNumeric, adParamInput, , NULL)
	END IF

	cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_repair_priority", adNumeric, adParamInput, , clng(Request("hdnLynx_Def_Sev_Lcode")))
	cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , CDate(Request("hdnUpdateDateTime")))
	cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_list", adVarChar, adParamOutput, 4000)
	cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_subject", adVarChar, adParamOutput, 4000)
	cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_message", adVarChar, adParamOutput, 4000)
'	if objConn.Errors.Count <> 0 then
'		DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT UPDATE OBJECT - PARAMETER ERROR", objConn.Errors(0).Description
'		objConn.Errors.Clear
'	end if
	'call the update stored proc
	cmdUpdateObj.Execute
	if objConn.Errors.Count <> 0 then
		DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT UPDATE OBJECT", objConn.Errors(0).Description
		objConn.Errors.Clear
	else
		dim strEmailFrom, strEmailTo, strEmailSubject, strEmailBody
		strEmailSubject = cmdUpdateObj.Parameters("p_subject").Value
		if strEmailSubject <> "" then
			'it's time to send an email
			strEmailTo = cmdUpdateObj.Parameters("p_list").Value
			strEmailBody = cmdUpdateObj.Parameters("p_message").Value
			Response.Cookies("txtEmailTo") = strEmailTo
			Response.Cookies("txtEmailSubject") = strEmailSubject
			Response.Cookies("txtEmailBody") = escape(strEmailBody)
		end if
	end if
	strWinMessage = "Record saved successfully. You can now see the changes you made."

end if


dim sql
'get Customer Service details:
if strCustomerServiceID <> "" then
	sql = "SELECT " &_
					"CUS.CUSTOMER_ID, "&_
					"CUS.CUSTOMER_NAME, "& _
					"CUS.CUSTOMER_SHORT_NAME, "&_
					"CUS.NOC_REGION_LCODE, "&_
	 				"CS.CUSTOMER_SERVICE_ID, "&_
					"CS.CUSTOMER_SERVICE_DESC, "&_
					"CS.SERVICE_TYPE_ID, "&_
					"CS.CHECK_SERVICE_DEPENDENCY_FLAG, "&_
					"CS.CREATE_SERVICE_TAG_FLAG, "&_
					"ST.SERVICE_TYPE_DESC, "& _
					"CS.SERVICE_LEVEL_AGREEMENT_ID, "&_
					"SLA.SERVICE_LEVEL_AGREEMENT_DESC, "&_
					"CS.SERVICE_LOCATION_ID, " & _
					"SL.SERVICE_LOCATION_NAME, " & _
					"ADDR.BUILDING_NAME, " & _
					"ADDR.STREET_NAME, " & _
					"ADDR.MUNICIPALITY_NAME, " & _
					"ADDR.PROVINCE_STATE_LCODE, " & _
					"CS.SERVICE_STATUS_CODE, " & _
					"CS.PROJECT_CODE, " & _
					"TO_CHAR(CS.DATE_IN_SERVICE, 'MON-DD-YYYY') AS date_in_service, " &_
					"TO_CHAR(CS.DATE_TERMINATED, 'MON-DD-YYYY') AS date_terminated, " &_
					"TO_CHAR(CS.DATE_TO_START_BILLING, 'MON-DD-YYYY') AS date_to_start_billing,  " &_
					"SG.REMEDY_SUPPORT_GROUP_ID, "&_
					"SG.GROUP_NAME, "&_
					"CS.COMMENTS, " &_
					"CS.NO_OF_SEATS, " &_
					"CS.LYNX_DEF_SEV_LCODE, " &_
					"CS.RECORD_STATUS_IND, "&_
					"ST.SEND_TO_NC_LCODE," &_
					"sma_sp_userid.spk_sma_library.sf_get_full_username(CS.CREATE_REAL_USERID) as create_real_userid, " &_
					"TO_CHAR(CS.CREATE_DATE_TIME,'MON-DD-YYYY HH24:MI:SS') CREATE_DATE_TIME, "&_
					"sma_sp_userid.spk_sma_library.sf_get_full_username(CS.UPDATE_REAL_USERID) as update_real_userid, "&_
					"TO_CHAR(CS.UPDATE_DATE_TIME,'MON-DD-YYYY HH24:MI:SS') UPDATE_DATE_TIME, "&_
					"CS.UPDATE_DATE_TIME AS LAST_UPDATE_DATE_TIME "&_
				"FROM "&_
					"CRP.CUSTOMER_SERVICE			CS, "&_
					"CRP.CUSTOMER					CUS, "&_
					"CRP.SERVICE_TYPE				ST, "&_
					"CRP.SERVICE_LEVEL_AGREEMENT	SLA, "&_
					"CRP.SERVICE_LOCATION			SL, "&_
					"CRP.V_REMEDY_SUPPORT_GROUP		SG, " &_
					"CRP.ADDRESS					ADDR  " &_

				"WHERE "&_
					"CS.CUSTOMER_ID = CUS.CUSTOMER_ID " &_
					"AND CS.SERVICE_TYPE_ID = ST.SERVICE_TYPE_ID " &_
					"AND CS.SERVICE_LEVEL_AGREEMENT_ID = SLA.SERVICE_LEVEL_AGREEMENT_ID " &_
					"AND CS.SERVICE_LOCATION_ID = SL.SERVICE_LOCATION_ID(+)    " &_
					"AND CS.REMEDY_SUPPORT_GROUP_ID = SG.REMEDY_SUPPORT_GROUP_ID(+) " &_
					"AND SL.ADDRESS_ID = ADDR.ADDRESS_ID(+) " &_
					"AND CS.CUSTOMER_SERVICE_ID = " & strCustomerServiceID


	'get the customer service recordset
	if err then
		DisplayError "BACK", "", err.Number, "UNEXPECTED ERROR", err.Description
	end if
	dim rsCustServ
	set rsCustServ=server.CreateObject("ADODB.Recordset")
	rsCustServ.CursorLocation = adUseClient
	rsCustServ.Open sql, objConn
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
	end if
end if

'if rsCustServ.EOF then
'	DisplayError "BACK", "", 999, "CANNOT FIND REQUESTED OBJECT", "EOF condition occured in rsCustServ recordset."
'end if
'release connection
set rsCustServ.ActiveConnection = nothing

'check if NC
'Response.Write sql
'Response.End

		If  (rsCustServ("send_to_nc_lcode") = "2") Then
			if (rsCustServ("province_state_lcode") = "QC" and rsCustServ("NOC_REGION_LCODE") = "QUEBEC") Then
				strDisable = ""
			else
				strDisable = "DISABLED"
			End If
		End If

'Response.Write "<b>" & rsCustServ("send_to_nc_lcode") & "</b>"
'Response.Write "<b>" & rsCustServ("NOC_REGION_LCODE") & "</b>"
'Response.Write "<b>" & rsCustServ("PROVINCE_STATE_LCODE") & "</b>"
'Response.Write "<b>" & rsCustServ("SERVICE_TYPE_ID") & "</b>"
'Response.Write "<b>" & rsCustServ("CUSTOMER_SERVICE_ID") & "</b>"
strServiceTypeID = rsCustServ("SERVICE_TYPE_ID")
'Response.Write "<b>" & strDisable & "</b>"

'get status list
sql = "select SERVICE_STATUS_CODE, SERVICE_STATUS_NAME " &_
		"from CRP.SERVICE_STATUS " &_
		"where RECORD_STATUS_IND = 'A' " &_
		"order by SERVICE_STATUS_NAME "
dim rsStatus
set rsStatus = Server.CreateObject("ADODB.Recordset")
rsStatus.CursorLocation = adUseClient
rsStatus.Open sql, objConn
if err then
	DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
end if
rsStatus.MoveFirst
'release the active connection, keep the recordset open
set rsStatus.ActiveConnection = nothing

'get support group list
sql = "select REMEDY_SUPPORT_GROUP_ID, GROUP_NAME " &_
		"from CRP.V_REMEDY_SUPPORT_GROUP " &_
		"order by GROUP_NAME "
dim rsSupportGroup
set rsSupportGroup = Server.CreateObject("ADODB.Recordset")
rsSupportGroup.CursorLocation = adUseClient
rsSupportGroup.Open sql, objConn
if err then
	DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
end if
'release the active connection, keep the recordset open
set rsSupportGroup.ActiveConnection = nothing


'pre-format Location Address
dim strServLocAddress
if len(rsCustServ("building_name") ) > 0 then
	strServLocAddress = rsCustServ("building_name") & vbNewLine & rsCustServ("street_name") & vbNewLine&_
				   rsCustServ("municipality_name") & " " & rsCustServ("province_state_lcode")
else
	strServLocAddress = rsCustServ("street_name") & vbNewLine & rsCustServ("municipality_name") & " " & rsCustServ("province_state_lcode")
end if


'create the innerValues for the iFrame
dim intRowCount, intColCount, strInnerValues
intRowCount = 0
intColCount = 5
strInnerValues = ""

if strDisable <> "" Then
Response.Write "<B>NetCracker Item. Update is limited to adding and deleting Facilities.</B>"
End If

%>
<html>
<head>
    <meta http-equiv="Pragma" content="no-cache">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <link href="stylesheets/styles.css" rel="stylesheet" type="text/css">
    <script type="text/javascript" src="GeneralJavaFunctions.js"></script>
    <script type="text/javascript" src="AccessLevels.js"></script>
    <script type="text/javascript">
        //set section title
        //set the heading
        var bolSaveRequired = false;
        var intAccessLevel = <%=intAccessLevel%>;
        var intAccessLevelDetail = <%=CheckLogon(strConst_CorrelationElements)%>
        var intConst_MessageDisplay=<%=intConst_MessageDisplay%>
        var strWinMessage = '<%=strWinMessage%>';

        setPageTitle("SMA - Managed Correlation");

        <%if strEmailSubject <> "" then%>
        //pop-up the email window
        var wndEmail = window.open('email.asp', 'PopupEmail', 'top=50, left=100, height=610, width=800' ) ;
        <%end if%>

        function fct_clearStatus() {
            window.status = "";
        }

            function fct_displayStatus(strMessage){
                window.status = strMessage;
                setTimeout('fct_clearStatus()',intConst_MessageDisplay);
            }

        function body_onLoad(){
            debugger;
            fct_displayStatus(strWinMessage);
            iFrame_display();
            iframe1_display();
            iframe2_display();
            iframe3_display();
        }

        function qlink_onChange(optValue){
            switch (optValue) {
                case 'CustServ':
                    document.frmCorrDetail.selQuickLink.selectedIndex=0;
                    document.location.href = 'custservdetail.asp?CustServID=' + document.frmCorrDetail.txtCustServID.value;
                    break;
                case 'CustServVPN':
                    document.frmCorrDetail.selQuickLink.selectedIndex=0;
                    document.location.href = 'CustServCPDetail.asp?CustServID=' + document.frmCorrDetail.txtCustServID.value;
                    break;
                case 'Customer':
                    document.frmCorrDetail.selQuickLink.selectedIndex=0;
                    document.location.href = 'CustDetail.asp?CustomerID=' + document.frmCorrDetail.txtCustomerID.value;
                    break;
                case 'CorrelationVpn':
                    document.frmCorrDetail.selQuickLink.selectedIndex=0;
                    if (document.frmCorrDetail.txtCustServID.value != ""){SetCookie("CustomerServiceID", document.frmCorrDetail.txtCustServID.value)};
                    document.location.href = 'CorrCPDetail.asp?CustServID=' + document.frmCorrDetail.txtCustServID.value;
                    break;
                case 'OrderHistory':
                    document.frmCorrDetail.selQuickLink.selectedIndex=0;
                    //SetCookie("CustomerServiceName", document.frmCorrDetail.txtCustServName.value);
                    if (document.frmCorrDetail.txtCustServID.value != ""){SetCookie("CustomerServiceID", document.frmCorrDetail.txtCustServID.value)};
                    self.location.href = 'SearchFrame.asp?fraSrc=OrderHistory';
                    break;
                case 'PVC':
                    var doc;
                    var iframeObject = document.getElementById('aifr'); // MUST have an ID
                    if (iframeObject.contentDocument) { // DOM
                        doc = iframeObject.contentDocument;
                    } 
                    else if (iframeObject.contentWindow) { // IE win
                        doc = iframeObject.contentWindow.document;
                    }

                    
                    if (doc.getElementsByName("txtObjClass")[0].value == 'CIRCUIT' && doc.getElementsByName("txtObjType")[0].value == 'ATMPVC') {
                        document.frmCorrDetail.selQuickLink.selectedIndex=0;
                        document.location.href = 'FacilityDetail.asp?CircuitID=' + doc.getElementsByName("txtObjID")[0].value + '&CircuitTyp=' + doc.getElementsByName("txtObjType")[0].value;
                    }
                    else {
                        document.frmCorrDetail.selQuickLink.selectedIndex=0;
                        alert('No PVC is currently selected. Please select one and try again.');
                    }
                    break;
                case 'Facility':
                    var doc;
                    var iframeObject = document.getElementById('aifr'); // MUST have an ID
                    if (iframeObject.contentDocument) { // DOM
                        doc = iframeObject.contentDocument;
                    } 
                    else if (iframeObject.contentWindow) { // IE win
                        doc = iframeObject.contentWindow.document;
                    }


                    if (doc.getElementsByName("txtObjClass")[0].value == 'CIRCUIT' && doc.getElementsByName("txtObjType")[0].valu!= 'ATMPVC') {
                        document.frmCorrDetail.selQuickLink.selectedIndex=0;
                        document.location.href = 'FacilityDetail.asp?CircuitID=' + doc.getElementsByName("txtObjID")[0].value + '&CircuitTyp=' + doc.getElementsByName("txtObjType")[0].value;
                    }
                    else {
                        document.frmCorrDetail.selQuickLink.selectedIndex=0;
                        alert('No facility is currently selected. Please select one and try again.');
                    }
                    break;
                case "ManObj":

                    var doc;
                    var iframeObject = document.getElementById('aifr'); // MUST have an ID
                    if (iframeObject.contentDocument) { // DOM
                        doc = iframeObject.contentDocument;
                    } 
                    else if (iframeObject.contentWindow) { // IE win
                        doc = iframeObject.contentWindow.document;
                    }

                    if (doc.getElementsByName("txtObjClass")[0].value == 'MO') {
                        document.frmCorrDetail.selQuickLink.selectedIndex=0;
                        document.location.href = 'manobjdet.asp?ne_id=' + doc.getElementsByName("txtObjID")[0].value;
                    }
                    else {
                        document.frmCorrDetail.selQuickLink.selectedIndex=0;
                        alert('No managed object is currently selected. Please select one and try again.');
                    }
                    break;
                case "Root":

                    var doc;
                    var iframeObject = document.getElementById('aifr'); // MUST have an ID
                    if (iframeObject.contentDocument) { // DOM
                        doc = iframeObject.contentDocument;
                    } 
                    else if (iframeObject.contentWindow) { // IE win
                        doc = iframeObject.contentWindow.document;
                    }

                    if (doc.getElementsByName("txtObjClass")[0].value == 'ROOT') {
                        document.frmCorrDetail.selQuickLink.selectedIndex=0;
                        document.location.href = 'custservdetail.asp?CustServID=' + doc.getElementsByName("txtObjID")[0].value;
                    }
                    else {
                        document.frmCorrDetail.selQuickLink.selectedIndex=0;
                        alert('No root service is currently selected. Please select one and try again.');
                    }
                    break;
            }
        }

        function btnCalendar_onclick(intDateFieldNo) {
            var NewWin;
            if (intDateFieldNo != ""){SetCookie("Field",intDateFieldNo)};
            NewWin=window.open("calendar.asp","NewWin","toolbar=no,status=no,width=260,height=200,menubar=no,resize=no");
            NewWin.focus();
            fct_onChange();
        }

        function fct_onChange(){
            bolSaveRequired = true;
        }

        function fct_onDelete(){
            if (((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) || (document.frmAlias.txtRecordStatusInd.value == "D")) {alert('Access denied. Please contact your system administrator.'); return;}
            alert("DELETE");
        }

        function fct_onReset(){
            if(confirm('All changes will be lost. Do you really want to reset the page?')){
                bolSaveRequired = false;
                document.location = 'corrdetail.asp?CustomerServiceID=<%=strCustomerServiceID%>';
            }
        }

        //iframe section
        var intCustServID=<%=strCustomerServiceID%>;
        var intRowCount=<%=intRowcount%>;
        var intColCount=<%=intColCount%>;
        var iFrameValues='<%=strInnerValues%>';
        var strDelimiter='<%=strDelimiter%>';
        var intServTypeID='<%=strServiceTypeID%>';
        <% if isnumeric(strServiceTypeID) then %>
                intServiceTypeID = <%=strServiceTypeID%> ;
        <% end if %>

        function iFrame_display(){

                    if(document.location.href.indexOf("corrdetail")>-1)
        {
                    document.getElementById("aifr").src = document.location.href.replace("corrdetail", "iFrmCorr") ;//iFrmCorr
        }
        else{
                    document.getElementById("aifr").src ='iFrmCorr.asp?CustomerServiceID=' + intCustServID;
        }
            
            //document.frames("aifr").document.location.href = 'iFrmCorr.asp?CustomerServiceID=' + intCustServID;
        }

            function iframe1_display(){
                document.getElementById("aifr2").src= 'CorrUsageList.asp?ServiceTypeID=' + intServTypeID;
                //	document.frames("aifr2").document.location.href = 'CorrUsageList.asp?CustomerServiceID=' + intCustServID;
                //	document.frames("aifr2").document.location.href = 'CorrUsageList.asp?CustomerServiceID=' + intCustServID + '&ServiceTypeID=' + intServTypeID;
                //	document.frames("aifr2").document.location.href = 'STypeAttrList.asp?CustomerServiceID=' + intCustServID + '&ServiceTypeID=' + intServTypeID;
            }

        function iframe2_display(){
            document.getElementById("aifr3").src= 'CorrSOInstList.asp?ServiceTypeID=' + intServTypeID + '&CustomerServiceID=' + intCustServID;
        }

        function iframe3_display(){
            document.getElementById("waifr3").src  = 'CorrSOWInstList.asp?ServiceTypeID=' + intServTypeID + '&CustomerServiceID=' + intCustServID;
        }


        function btn_iFrmDelete(){
            //delete selected row
            if ((intAccessLevelDetail & intConst_Access_Delete) != intConst_Access_Delete) {alert('Access denied. Please contact your system administrator.'); return;}
            var doc;
            var iframeObject = document.getElementById('aifr'); // MUST have an ID
            if (iframeObject.contentDocument) { // DOM
                doc = iframeObject.contentDocument;
            } 
            else if (iframeObject.contentWindow) { // IE win
                doc = iframeObject.contentWindow.document;
            }
            var strObjName = doc.getElementsByName("txtObjName")[0].value ;
            if (strObjName != "") {
                if (confirm('Do you want to delete the element "' + strObjName + '"')) {
                    iFrameValues = 'action=delete&delObjID=' +  doc.getElementsByName("txtCorrID")[0].value + '&txtUpdateDateTime=' + doc.getElementsByName("txtLastUpdate")[0].value ;
                    document.getElementById("aifr").src = 'iFrmCorr.asp?CustomerServiceID=' + intCustServID + '&' + iFrameValues;
                }
            }
            else alert('You must select an element first.');
        }//end of btn_iFrmDelete()

        function btn_iFrmNewRoot() {
            //adds a customer service to the correlation list
            if ((intAccessLevelDetail & intConst_Access_Create) != intConst_Access_Create) {alert('Access denied. Please contact your system administrator.'); return;}
            SetCookie("WinName", 'Popup');
            SetCookie("ServiceEnd", 'C');
            window.open('SearchFrame.asp?fraSrc=CustServ', 'Popup', 'top=50, left=100, height=600, width=800' ) ;
        }

        function btn_iFrmNewPVC() {
            //adds a new PVC to the correlation list
            if ((intAccessLevelDetail & intConst_Access_Create) != intConst_Access_Create) {alert('Access denied. Please contact your system administrator.'); return;}
            SetCookie("WinName", 'Popup');
            window.open('SearchFrame.asp?fraSrc=FacilityPVC', 'Popup', 'top=50, left=100, height=600, width=800' ) ;
        }

        function btn_iFrmNewFacility() {
            //adds a new PVC to the correlation list
            if ((intAccessLevelDetail & intConst_Access_Create) != intConst_Access_Create) {alert('Access denied. Please contact your system administrator.'); return;}
            SetCookie("WinName", 'Popup');
            window.open('SearchFrame.asp?fraSrc=Facility', 'Popup', 'top=50, left=100, height=600, width=800' ) ;
        }

        function btn_iFrmNewMO() {
            //adds a new PVC to the correlation list
            if ((intAccessLevelDetail & intConst_Access_Create) != intConst_Access_Create) {alert('Access denied. Please contact your system administrator.'); return;}
            SetCookie("WinName", 'Popup');
            window.open('SearchFrame.asp?fraSrc=ManagedObjects', 'Popup', 'top=50, left=100, height=600, width=800' ) ;
        }

        function btn_iFrmAddNewElement(){
            if ((intAccessLevelDetail & intConst_Access_Create) != intConst_Access_Create) {alert('Access denied. Please contact your system administrator.'); return;}
            iFrameValues = 'action=add&newType=' + document.frmCorrDetail.hdnNewElementType.value + '&newID=' + document.frmCorrDetail.hdnNewElementID.value ;
            document.getElementById("aifr").src   = 'iFrmCorr.asp?CustomerServiceID=' + intCustServID + '&' + iFrameValues;
        }

        function fct_onMoveUp(){
            if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update) {alert('Access denied. Please contact your system administrator.'); return;}

            var doc;
            var iframeObject = document.getElementById('aifr'); // MUST have an ID
            if (iframeObject.contentDocument) { // DOM
                doc = iframeObject.contentDocument;
            } 
            else if (iframeObject.contentWindow) { // IE win
                doc = iframeObject.contentWindow.document;
            }
            var strObjName = doc.getElementsByName("txtObjName")[0].value ;
            if (strObjName != "") {
                strParams = 'action=move&direction=up&corrid=' + doc.getElementsByName("txtCorrID")[0].value ;
                document.getElementById("aifr").src  = 'iFrmCorr.asp?CustomerServiceID=' + intCustServID + '&' + strParams;
            } else alert('You must select an element first.');
        }

        function fct_onMoveDown(){
            if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update) {alert('Access denied. Please contact your system administrator.'); return;}
            var doc;
            var iframeObject = document.getElementById('aifr'); // MUST have an ID
            if (iframeObject.contentDocument) { // DOM
                doc = iframeObject.contentDocument;
            } 
            else if (iframeObject.contentWindow) { // IE win
                doc = iframeObject.contentWindow.document;
            }
            var strObjName = doc.getElementsByName("txtObjName")[0].value ;
            if (strObjName != "") {
                strParams = 'action=move&direction=down&corrid=' +  doc.getElementsByName("txtCorrID")[0].value ;
                document.getElementById("aifr").src = 'iFrmCorr.asp?CustomerServiceID=' + intCustServID + '&' + strParams;
            } else alert('You must select an element first.');
        }

        function body_onBeforeUnload(){
            //	document.frmCorrDetail.btnSave.focus();
            if (bolSaveRequired) {
                if	(((intAccessLevel & intConst_Access_Create == intConst_Access_Create) && (document.frmCorrDetail.CustomerServiceID.value == "")) || ((intAccessLevel & intConst_Access_Update == intConst_Access_Update) && (document.frmCorrDetail.CustomerServiceID.value != ""))) {
                    event.returnValue = "There is unsaved data on the screen. To save changes, click CANCEL below then click SAVE on the main form.";
                }
            }
        }

        function fct_onSave(){
            if	((((intAccessLevel & intConst_Access_Create) == intConst_Access_Create) && (document.frmCorrDetail.CustomerServiceID.value == "")) || (((intAccessLevel & intConst_Access_Update) == intConst_Access_Update) && (document.frmCorrDetail.CustomerServiceID.value != ""))) {

                //validate fields
                if ((document.frmCorrDetail.selday.selectedIndex != 0)||(document.frmCorrDetail.selmonth.selectedIndex != 0)||(document.frmCorrDetail.selyear.selectedIndex != 0))  {
                    if ((document.frmCorrDetail.selday.selectedIndex == 0)||(document.frmCorrDetail.selmonth.selectedIndex == 0)||(document.frmCorrDetail.selyear.selectedIndex == 0)) {
                        alert('Date Start Billing incomplete. Try again.')
                        return (false);}

                }		if ((document.frmCorrDetail.selday2.selectedIndex != 0)||(document.frmCorrDetail.selmonth2.selectedIndex != 0)||(document.frmCorrDetail.selyear2.selectedIndex != 0))  {
                    if ((document.frmCorrDetail.selday2.selectedIndex == 0)||(document.frmCorrDetail.selmonth2.selectedIndex == 0)||(document.frmCorrDetail.selyear2.selectedIndex == 0)) {
                        alert('Date In Service incomplete. Try again.')
                        return (false);}
                }
                if ((document.frmCorrDetail.selday3.selectedIndex != 0)||(document.frmCorrDetail.selmonth3.selectedIndex != 0)||(document.frmCorrDetail.selyear3.selectedIndex != 0))  {
                    if ((document.frmCorrDetail.selday3.selectedIndex == 0)||(document.frmCorrDetail.selmonth3.selectedIndex == 0)||(document.frmCorrDetail.selyear3.selectedIndex == 0)) {
                        alert('Date Terminated incomplete. Try again.')
                        return (false);}
                }
                if (document.frmCorrDetail.selSupportGroup.selectedIndex == 0)  {
                    alert('Please select a support group from the drop-down list.');
                    document.frmCorrDetail.selSupportGroup.focus() ;
                    return (false);
                }
                //save
                bolSaveRequired = false;
                document.frmCorrDetail.txtFrmAction.value = "SAVE";
                document.frmCorrDetail.submit();
                return(true);
            }else{
                alert('Access denied. Please contact your system administrator.');
                return(false);
            }
        }

        function btnExpand_onclick(){
            //if ((intAccessLevel & intConst_Access_ReadOnly) != intConst_Access_ReadOnly) {alert('Access denied. Please contact your system administrator.'); return;}
            var CSID = document.frmCorrDetail.CustomerServiceID.value;
            var CSName = document.frmCorrDetail.CustomerServiceName.value;
            var URL;
            URL='CorrReport.asp?CSID='+CSID+'&CSName='+CSName;
            window.open(URL,'Popup','top=100,left=100,WIDTH=700,HEIGHT=500,scrollbars=yes,resizable=yes');
        }

        function btnNetcrackerweblink_onclick() {
            var strCSID = document.frmCorrDetail.CustomerServiceID.value ;
            var strNetcrackerURL = '<%=strConstNetcrackerURL%>';
            var strNetcrackerURLCSID ;

            strNetcrackerURCSID = strNetcrackerURL + strCSID

            window.open(strNetcrackerURCSID);
        }

    </script>
</head>

<body onload="body_onLoad();" onbeforeunload="body_onBeforeUnload();">
    <form name="frmCorrDetail" action="CorrDetail.asp" method="post">
        <input type="hidden" name="hdnNewElementID" value>
        <input type="hidden" name="hdnNewElementType" value>
        <input type="hidden" name="hdnNewElementName" value>
        <input type="hidden" name="hdnNewCircuitName" value>
        <input name="hdnUpdateDateTime" type="hidden" value="<%=rscustServ("LAST_UPDATE_DATE_TIME")%>">
        <input name="CustomerServiceID" type="hidden" value="<%=strCustomerServiceID%>">
        <input name="CustomerServiceName" type="hidden" value="<%=routineHTMLString(rscustServ("CUSTOMER_SERVICE_DESC"))%>">
        <input name="txtCustomerID" type="hidden" value="<%=rscustServ("CUSTOMER_ID")%>">
        <input name="hdnLynx_Def_Sev_Lcode" type="hidden" value="<%=rscustServ("LYNX_DEF_SEV_LCODE")%>">
        <input name="hdnServiceTypeID" type="hidden" value="<%=rscustServ("SERVICE_TYPE_ID")%>">
        <input name="txtFrmAction" type="hidden" value="">
        <table border="0" width="100%" cols="3">
            <thead>
                <tr>
                    <td colspan="2">Managed Correlation - Details</td>
                    <td align="right">
                        <select name="selQuickLink" size="1" onchange="qlink_onChange(this.value);">
                            <option value>Quickly Goto...</option>
                            <option value="CustServ">Customer Service</option>
                            <option value="CustServVPN">Customer Service VPN</option>
                            <option value="Customer">Customer</option>
                            <option value="CorrelationVpn">Correlation VPN</option>
                            <option value="OrderHistory">Order History</option>
                            <option value="Root">Root Customer Service</option>
                            <option value="Facility">Facility</option>
                            <option value="ManObj">Managed Objects</option>
                            <option value="PVC">PVC</option>
                        </select>
                    </td>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td align="right">Customer Service ID&nbsp;</td>
                    <td>
                        <input disabled name="txtCustServID" style="width: 100%" value="<%=rscustServ("CUSTOMER_SERVICE_ID")%>"></td>
                    <td width="50%" align="left" valign="top" rowspan="14">Correlated Elements&nbsp;&nbsp;<input type="button" name="btnExpand" style="width: 2cm" value="Expand" onclick="btnExpand_onclick();"><br>
                        <iframe id="aifr" width="100%" height="96%" src scrolling="yes" marginheight="1" marginwidth="1"></iframe>
                    </td>
                </tr>
                <tr>
                    <td width="145" align="right">Customer Service Name&nbsp;</td>
                    <td>
                        <input disabled name="txtCustServName" style="width: 100%" value="<%=routineHTMLString(rscustServ("CUSTOMER_SERVICE_DESC"))%>"></td>
                </tr>
                <tr>
                    <td align="right">Customer Name&nbsp;</td>
                    <td>
                        <input disabled name="txtCustomerName" style="width: 100%" value="<%=routineHTMLString(rscustServ("CUSTOMER_NAME"))%>"></td>
                </tr>
                <tr>
                    <td align="right">Service Location&nbsp;</td>
                    <td>
                        <input disabled name="txtServiceLocation" style="width: 100%" value="<%=routineHTMLString(rscustServ("SERVICE_LOCATION_NAME"))%>"></td>
                </tr>
                <tr>
                    <td align="right" valign="top">Location Address&nbsp;</td>
                    <td>
                        <textarea disabled name="txtLocationAddress" style="width: 100%" rows="3"><%=routineHTMLString(strServLocAddress)%></textarea></td>
                </tr>
                <tr>
                    <td align="right">Service Type&nbsp;</td>
                    <td>
                        <input disabled name="txtServiceType" style="width: 100%" value="<%=routineHTMLString(rscustServ("SERVICE_TYPE_DESC"))%>"></td>
                </tr>
                <tr>
                    <td align="right">Customer Region&nbsp;</td>
                    <td>
                        <input disabled name="txtCustomerRegion" style="width: 100%" value="<%=routineHTMLString(rscustServ("NOC_REGION_LCODE"))%>"></td>
                </tr>
                <tr>
                    <td align="right">Order Number&nbsp;</td>
                    <td>
                        <input disabled name="txtOrderNumber" style="width: 100%" value="<%=routineHTMLString(rscustServ("PROJECT_CODE"))%>"></td>
                </tr>
                <tr>
                    <td align="right">SLA&nbsp;</td>
                    <td>
                        <input name="txtSLA" style="width: 100%" value="<%=routineHTMLString(rscustServ("SERVICE_LEVEL_AGREEMENT_DESC"))%>"></td>
                </tr>
                <tr>
                    <td align="right">Support Group&nbsp;<font color="red">*</font></td>
                    <td>
                        <select name="selSupportGroup" onchange="fct_onChange()">
                            <option></option>
                            <%while not rsSupportGroup.EOF
						If rsSupportGroup(1) = rscustServ("GROUP_NAME") Then
							Response.Write "<option value='"&rsSupportGroup(0)&"' selected >" & routineHtmlString(rsSupportGroup(1)) & "</option>" & vbCrLf
			  	        Else
				 			Response.write "<option value='"&rsSupportGroup(0)&"'>" & routineHtmlString(rsSupportGroup(1)) & "</option>" & vbCrLf
						End If
						rsSupportGroup.MoveNext
					wend
					rsSupportGroup.Close
                            %>
                        </select>
                    </td>
                </tr>
                <tr>
                    <td align="right">Service Status<font color="red">*</font></td>
                    <td>
                        <select onchange="fct_onChange()" name="selStatus">
                            <%while not rsStatus.EOF
						If rsStatus(0) = rscustServ("SERVICE_STATUS_CODE") Then
							Response.Write "<option selected value='" & rsStatus(0) & "'>" & routineHtmlString(rsStatus(1)) & "</option>" & vbCrLf
						Else
							Response.write "<option value='" & rsStatus(0)& "'>" & routineHtmlString(rsStatus(1)) & "</option>" & vbCrLf
						End If
						rsStatus.MoveNext
					wend
					rsStatus.Close
                            %>
                        </select>
                    </td>
                </tr>
                <tr>
                    <td align="right">Date Start Billing&nbsp; </td>
                    <td align="left">
                        <select name="selmonth" size="1" onchange="fct_onChange();">
                            <option></option>
                            <%
					dim k
 					for k = 1 to 12
						Response.Write "<option "
						if strCustomerServiceID <> 0 then
							if k = month(rsCustServ("date_to_start_billing")) then
								Response.Write " selected "
							end if
						end if
						if k < 10 then
							k="0"&k
						end if
						Response.write " VALUE ="& k & ">" &ucase(monthName(k,true)) & "</OPTION>"
					next
                            %>
                        </select>
                        <select name="selday" size="1" onchange="fct_onChange();">
                            <option></option>
                            <%
					for k = 1 to 31
						Response.Write "<option "
						if strCustomerServiceID <> 0 then
							if k = day(rsCustServ("date_to_start_billing")) then
								Response.Write " selected "
							end if
						end if
						if k < 10 then
							k="0"&k
						end if
						Response.write " VALUE ="& k & ">" &k & "</OPTION>"
					next
                            %>
                        </select>
                        <select name="selyear" size="1" onchange="fct_onChange();">
                            <option></option>
                            <%
					for k = 1990 to 2050
						Response.Write "<option "
						if strCustomerServiceID <> 0 then
							if k = year(rsCustServ("date_to_start_billing")) then
								Response.Write " selected "
							end if
						end if
						Response.write " VALUE ="& k & ">" &k & "</OPTION>"
					next
                            %>
                        </select>
                        <input type="button" value="..." name="btnCalendar" onclick="return btnCalendar_onclick(1)">
                    </td>
                </tr>
                <tr>
                    <td align="right">Date In Service&nbsp; </td>
                    <td>
                        <select name="selmonth2" size="1" onchange="fct_onChange();">
                            <option></option>
                            <%
 					for k = 1 to 12
						Response.Write "<option "
						if strCustomerServiceID <> 0 then
							if k = month(rsCustServ("DATE_IN_SERVICE")) then
								Response.Write " selected "
							end if
						end if
						if k < 10 then
							k="0"&k
						end if
						Response.write " VALUE ="& k & ">" &ucase(monthName(k,true)) & "</OPTION>"
					next
                            %>
                        </select>
                        <select name="selday2" size="1" onchange="fct_onChange();">
                            <option></option>
                            <%
					for k = 1 to 31
						Response.Write "<option "
						if strCustomerServiceID <> 0 then
							if k = day(rsCustServ("DATE_IN_SERVICE")) then
								Response.Write " selected "
							end if
						end if
						if k < 10 then
							k="0"&k
						end if
						Response.write " VALUE ="& k & ">" &k & "</OPTION>"
					next
                            %>
                        </select>
                        <select name="selyear2" size="1" onchange="fct_onChange();">
                            <option></option>
                            <%
					for k = 1990 to 2050
						Response.Write "<option "
						if strCustomerServiceID <> 0 then
							if k = year(rsCustServ("DATE_IN_SERVICE")) then
								Response.Write " selected "
							end if
						end if
						Response.write " VALUE ="& k & ">" &k & "</OPTION>"
					next
                            %>
                        </select>
                        <input type="button" value="..." name="btnCalendar" onclick="return btnCalendar_onclick(2)">
                        &nbsp;&nbsp;<input type="checkbox" onclick="fct_onChange()" name="chkCreateService" <%if rscustServ("CREATE_SERVICE_TAG_FLAG")="Y" then response.write "checked"%>>
                        Create Service Tag
                    </td>
                </tr>
                <tr>
                    <td align="right">Date Terminated&nbsp; </td>
                    <td>
                        <select name="selmonth3" size="1" onchange="fct_onChange();">
                            <option></option>
                            <%
 					for k = 1 to 12
						Response.Write "<option "
						if strCustomerServiceID <> 0 then
							if k = month(rsCustServ("DATE_TERMINATED")) then
								Response.Write " selected "
							end if
						end if
						if k < 10 then
							k="0"&k
						end if
						Response.write " VALUE ="& k & ">" &ucase(monthName(k,true)) & "</OPTION>"
					next
                            %>
                        </select>
                        <select name="selday3" size="1" onchange="fct_onChange();">
                            <option></option>
                            <%
					for k = 1 to 31
						Response.Write "<option "
						if strCustomerServiceID <> 0 then
							if k = day(rsCustServ("DATE_TERMINATED")) then
								Response.Write " selected "
							end if
						end if
						if k < 10 then
							k="0"&k
						end if
						Response.write " VALUE ="& k & ">" &k & "</OPTION>"
					next
                            %>
                        </select>
                        <select name="selyear3" size="1" onchange="fct_onChange();">
                            <option></option>
                            <%
					for k = 1990 to 2050
						Response.Write "<option "
						if strCustomerServiceID <> 0 then
							if k = year(rsCustServ("DATE_TERMINATED")) then
								Response.Write " selected "
							end if
						end if
						Response.write " VALUE ="& k & ">" &k & "</OPTION>"
					next
                            %>
                        </select>
                        <input type="button" value="..." name="btnCalendar" onclick="return btnCalendar_onclick(3)">
                        &nbsp;&nbsp;<input type="checkbox" onclick="fct_onChange()" name="chkDependency" <%if rscustServ("CHECK_SERVICE_DEPENDENCY_FLAG")="Y" then response.write "checked"%>>
                        Check Service Dep
                    </td>
                </tr>
                <tr>
                    <td align="right" valign="top">Comments&nbsp; </td>
                    <td>
                        <textarea onchange="fct_onChange()" name="txtComment" style="width: 100%"><%=routineHTMLString(rscustServ("COMMENTS"))%></textarea></td>
                    <td align="right" valign="top">
                        <input type="button" name="btnNewMO" style="width: 2cm" value="+ MO" onclick="btn_iFrmNewMO();" <%=strDisable%>>
                        <input type="button" name="btnNewPVC" style="width: 2cm" value="+ PVC" onclick="btn_iFrmNewPVC();" <%=strDisable%>>
                        <input type="button" name="btnNewFacility" style="width: 2cm" value="+ Facility" onclick="btn_iFrmNewFacility();">
                        <input type="button" name="btnNewRoot" style="width: 2cm" value="+ Root" onclick="btn_iFrmNewRoot();" <%=strDisable%>>
                        <input type="button" name="btnDelete" style="width: 2cm" value="Delete" onclick="btn_iFrmDelete();">
                        <img src="images/up.gif" title width="31" height="31" <%=strDisable%> onclick="fct_onMoveUp()">
                        <img src="images/down.gif" title width="31" height="31" <%=strDisable%> onclick="fct_onMoveDown()">
                    </td>
                </tr>
                <tr>
                    <td width="15%" align="right" nowrap>NetCracker Weblink</td>
                    <td width="20%">
                        <input id="btnNetcrackerweblink" name="btnNetcrackerweblink" tabindex="2" style="height: 22px; width: 180px" type="button" value="NetCracker" language="javascript" onclick="return btnNetcrackerweblink_onclick()"></td>
                </tr>
                <tr>
                    <% if rsCustServ("NO_OF_SEATS") <> "" then
		         Response.Write "<td align=""right"">No. of Seats&nbsp;</td>"&vbCrLf
				 Response.Write "<td><input name=""txtNoOfSeats"" style=""WIDTH: 25%"" value="""%><%=rsCustServ("NO_OF_SEATS")%>"" <%=strDisable%>></td>
		    <%  elses
		         Response.Write "<td></td>"
		      end if
            %>
                </tr>
                <!-- 		<TABLE  border=1 cellPadding=2 cellSpacing=0 width="100%">
			<THEAD>
				<TR><td align=left colspan=4>Usage</td></tr>
				<TR>
					<TH align=left>Service Type Attribute</TH>
					<TH align=center>Value</TH>
					<TH align=left>Technical Question</TH>
					<TH align=left>Technical Answer</TH>
				 </TR>
			</THEAD>
			<TBODY><%
			%>
			</tbody>
		</table>-->



                <!-- New Frame begins -->
                <thead>
                    <tr>
                        <td align="left" colspan="2">Service Type Attributes and Values</td>
                        <td width="50%" align="left" valign="top" colspan="2">Working Service Instance Attribute Values for this CSID</td>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td colspan="2">
                            <iframe id="aifr2" width="100%" height="75" src="" scrolling="yes" marginheight="1" marginwidth="1"></iframe>
                        </td>
                        <td width="50%" align="left" valign="top" colspan="2">
                            <iframe id="waifr3" width="100%" height="75" src="" scrolling="yes" marginheight="1" marginwidth="1"></iframe>
                            <br>
                        </td>
                    </tr>
                </tbody>


                <thead>
                    <tr>
                        <td bgcolor="#FFFFCC" align="left" colspan="2"></td>
                        <td width="50%" align="left" valign="top" colspan="2">Service Instance Attribute Values Requested for this CSID in this Order</td>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td colspan="2"></td>
                        <td width="50%" align="left" valign="top" colspan="2">
                            <iframe id="aifr3" width="100%" height="75" src="" scrolling="yes" marginheight="1" marginwidth="1"></iframe>
                            <br>
                        </td>
                    </tr>
                </tbody>

                <!--New Frame ends -->
                <!-- New Frame begins -->
                <!--		<thead><TR><TD align=left colspan=4>Service Order Instances and Values</TD></TR></thead>
		<tbody>
			<TR>
				<TD colspan=4><iframe id=aifr3 width=50% height=75 src="" scrolling=yes marginheight=1 marginwidth=1></iframe>
				<br>
			</TR>
		</tbody>-->
                <!--New Frame ends -->
            </tbody>
            <tfoot>
                <tr>
                    <td colspan="3" align="right">
                        <input <%=strDisable%> name="btnReset" type="reset" style="width: 2cm" value="Reset" onclick="return fct_onReset();">&nbsp;&nbsp;
			<input name="btnSave" type="button" style="width: 2cm" value="Save" onclick="return fct_onSave();">&nbsp;&nbsp;
                    </td>
                </tr>
            </tfoot>
        </table>
        <%
if strDisable <> "" Then
Response.Write "<B>NetCracker Item. Update is limited to adding and deleting Facilities.</B>"
End If
        %>
        <fieldset>
            <legend align="right"><b>Audit Information</b></legend>
            <div size="8pt" align="right">
                Record Status Indicator&nbsp;<input align="left" name="txtRecordStatusInd" style="height: 20px; width: 18px" disabled value="<%=rscustServ("RECORD_STATUS_IND")%>">&nbsp;&nbsp;&nbsp;
	Create Date&nbsp;<input align="center" name="txtCreateDateTime" style="height: 20px; width: 150px" disabled value="<%=rscustServ("CREATE_DATE_TIME")%>">&nbsp;
	Created By&nbsp;<input align="right" name="txtCreateRealUser" style="height: 20px; width: 200px" disabled value="<%=rscustServ("CREATE_REAL_USERID")%>"><br>
                Update Date&nbsp;<input align="center" name="txtUpdateDateTime" style="height: 20px; width: 150px" disabled value="<%=rscustServ("UPDATE_DATE_TIME")%>">
                Updated By&nbsp;<input align="right" name="txtUpdateRealUser" style="height: 20px; width: 200px" disabled value="<%=rscustServ("UPDATE_REAL_USERID")%>">
            </div>
        </fieldset>
    </form>
</body>
</html>
