<%@  language="VBScript" %>
<% option explicit %>
<% Response.Buffer = true %>
<% on error resume next %>
<!--#include file="smaConstants.inc"-->
<!--#include file="smaProcs.inc"-->
<!--#include file="databaseconnect.asp" -->

<!--
********************************************************************************************
* Page name:	ServLocDetail.asp
* Purpose:		To display the detailed information about a Service Location.
*				Customer chosen via ServLocList.asp
*
* In Param:		This page reads Service Location ID from a query string.
*
* Out Param:	Sometimes this Page writes following cookeis
*				Cookie - ServLocName
*				Cookie - CustomerName
*				Cookie - WinName
*
*
* Created by:	Sara Sangha	08/11/2000
*
********************************************************************************************
		 Date		Author			Changes/enhancements made


		07-May-08   	ACheung 	Add CLLI Code (Geocode) as part of the Service Location selection
********************************************************************************************
-->
<%
     
const ASP_NAME = "ServLocDetail.asp"  'if the name of the file changes, you only have to change this constant to update this form.
const NO_ID = "null" 'if the Service Locationis new then the value of the id is manually set to this value

'--- check user's access rights
dim intAccessLevel, intChildAccessLevel
Dim strRealUserID
dim strWinMessage,strSiteNameCodeSelect


intAccessLevel = CInt(CheckLogon(strConst_ServiceLocation))
intChildAccessLevel = CInt(CheckLogon(strConst_ServiceLocationContact))
strRealUserID = Session("username")

'intAccessLevel = intConst_Access_ReadOnly or intConst_Access_Create or intConst_Access_Update or intConst_Access_Delete
if (intAccessLevel and intConst_Access_ReadOnly) <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to view service locations. Please contact your system administrator."
end if

Dim lngServLocID, geoclliid, geocllicode
Dim strNew, strSQL, strGeocode, strSQL2, objCmd
dim objRsSchedule, objRsServiceLocation, objRsServiceContact, objGeoRs, objSlGeoRs,objSites

strNew =Request.QueryString("NewServLoc")
lngServLocID = Request.QueryString("ServLocID")

if  strNew = "NEW" then
   lngServLocID = NO_ID
end if

dim strWinLocation
strWinLocation = ASP_NAME & "?ServLocID="&Request.Form("hdnServiceLocationID")
    
select case Request("hdnFrmAction")
	case "SAVE"
	  if Request.Form("hdnServiceLocationID")  <> "" then  ' it is an existing record so save the changes

		if (intAccessLevel and intConst_Access_Update) <> intConst_Access_Update then
			DisplayError "REFRESH", strWinLocation, 0, "UPDATE DENIED", "You don't have access to update service locations. Please contact your system administrator."
		end if
		dim cmdUpdateObj
		set cmdUpdateObj = server.CreateObject("ADODB.Command")
		set cmdUpdateObj.ActiveConnection = objConn
		cmdUpdateObj.CommandType = adCmdStoredProc
		cmdUpdateObj.CommandText = "sma_sp_userid.spk_sma_cust_inter.sp_sl_update"

		'create params
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("userid", adVarChar, adParamInput, 20, strRealUserID)
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("service_location_id", adNumeric, adParamInput, , Clng(Request("hdnServiceLocationID")))
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("customer_id", adNumeric, adParamInput, , CLng(Request("hdnCustomerID")))
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("sl_name", adVarChar, adParamInput, 50, Request("txtServiceLocationName"))
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("address", adNumeric, adParamInput, , CLng(Request("hdnAddressID")))
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("last_update", adDBTimeStamp, adParamInput, , CDate(Request("hdnUpdateDateTime")))

		if Request.Form("txtSpecificLocation") <> "" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("sl_specific_desc", adVarChar, adParamInput, 80, Request("txtSpecificLocation"))
		else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("sl_specific_desc", adVarChar, adParamInput, 80, null)
		end if

		if Request.Form("selSchedule") <> "" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("schedule_id", adNumeric, adParamInput, , CLng(Request("selSchedule")))
		else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("schedule_id", adNumeric, adParamInput, , null)
		end if

		if Request.Form("txtAccessInfo") <> "" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("access_info", adVarChar, adParamInput, 2000, Request("txtAccessInfo"))
		else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("access_info", adVarChar, adParamInput, 2000, null)
		end if

		if Request.Form("txtComments") <> "" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("comments", adVarChar, adParamInput, 2000, Request("txtComments"))
		else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("comments", adVarChar, adParamInput, 2000, null)
		end if

		cmdUpdateObj.Execute
		if objConn.Errors.Count <> 0 then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT UPDATE OBJECT", objConn.Errors(0).Description
			objConn.Errors.Clear
		end if
	' Need logic here to check how to control LC
		lngServLocID = Clng(Request("hdnServiceLocationID"))
		' ok till now
		if lngServLocID <> NO_ID then

			strGeocode = Request.Form("hdnGeocode")
			strSQL = "select GEOCODEID_LCODE from crp.SERVICE_LOCATION_GEOCODE " &_
				 "where SERVICE_LOCATION_ID = " & lngServLocID &""
			'ok till now
			set objSlGeoRs = objConn.Execute(strSQL)
			geoclliid = objSlGeoRs("GEOCODEID_LCODE")

			if geoclliid = "" then
				strSQL = "INSERT INTO crp.SERVICE_LOCATION_GEOCODE(GEOCODEID_LCODE , SERVICE_LOCATION_ID) " &_
						 "VALUES (" & strGeocode &"," & lngServLocID & ")"
			else
				strSQL ="Update crp.SERVICE_LOCATION_GEOCODE " &_
						"SET GEOCODEID_LCODE = " & strGeocode &", " &_
						"UPDATE_REAL_USERID = '" & strRealUserID & "' " &_
						"where SERVICE_LOCATION_ID = " & lngServLocID & ""
			end if
			'response.write(strSQL)
			'response.end
			objconn.Execute(strSQL)
			'objconn.Execute("commit")

		end if


		strWinMessage = "Record saved successfully. You can now see the changes you made."

	  else 'create a new record
		if (intAccessLevel and intConst_Access_Create) <> intConst_Access_Create then
			DisplayError "REFRESH", strWinLocation, 0, "INSERT DENIED", "You don't have access to create service locations. Please contact your system administrator."
		end if

		dim cmdInsertObj
		set cmdInsertObj = server.CreateObject("ADODB.Command")
		set cmdInsertObj.ActiveConnection = objConn
		cmdInsertObj.CommandType = adCmdStoredProc
		cmdInsertObj.CommandText = "sma_sp_userid.spk_sma_cust_inter.sp_sl_insert"

		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_user_id", adVarChar, adParamInput, 20, strRealUserID)
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_service_location_id", adNumeric , adParamOutput, , null)
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_customer_id", adNumeric, adParamInput, , CLng(Request("hdnCustomerID")))
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_sl_name", adVarChar, adParamInput, 50, Request("txtServiceLocationName"))
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_address", adNumeric, adParamInput, , CLng(Request("hdnAddressID")))

		if Request("txtSpecificLocation") <> "" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_sl_specific_desc", adVarChar, adParamInput, 80, Request("txtSpecificLocation"))
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_sl_specific_desc", adVarChar, adParamInput, 80, null)
		end if

		if Request.Form("selSchedule") <> "" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_schedule_id", adNumeric, adParamInput, , CLng(Request("selSchedule")))
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_schedule_id", adNumeric, adParamInput, , null)
		end if

		if Request("txtAccessInfo") <> "" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_access_info", adVarChar, adParamInput, 2000, Request("txtAccessInfo"))
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_access_info", adVarChar, adParamInput, 2000, null)
		end if

		if Request("txtComments") <> "" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_comments", adVarChar, adParamInput, 2000, Request("txtComments"))
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_comments", adVarChar, adParamInput, 2000, null)
		end if


		cmdInsertObj.Execute
		if objConn.Errors.Count <> 0 then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT CREATE NEW OBJECT", objConn.Errors(0).Description
			objConn.Errors.Clear
		else
			lngServLocID = cmdInsertObj.Parameters("p_service_location_id").Value
		end if

		strGeocode = Request.Form("hdnGeocode")
		strSQL = "INSERT INTO crp.SERVICE_LOCATION_GEOCODE(GEOCODEID_LCODE , SERVICE_LOCATION_ID) " &_
					 "VALUES (" & strGeocode &"," & lngServLocID & ")"
		'response.write(strSQL)
		'response.end

		objconn.Execute(strSQL)
		'objconn.Execute("commit")

		strWinMessage = "Record created successfully. You can now see the new record."

	  end if
	case "DELETE"
			strSQL = "DELETE crp.SERVICE_LOCATION_GEOCODE " &_
					 "where SERVICE_LOCATION_ID = " & lngServLocID & ""
			'response.write(strSQL)
			'response.end
			objconn.Execute(strSQL)
    
			if (intAccessLevel and intConst_Access_Delete) <> intConst_Access_Delete then
				DisplayError "REFRESH", strWinLocation, 0, "DELETE DENIED", "You don't have access to delete service locations. Please contact your system administrator"
			end if
			dim cmdDeleteObj
			set cmdDeleteObj = server.CreateObject("ADODB.Command")
			set cmdDeleteObj.ActiveConnection = objConn
			cmdDeleteObj.CommandType = adCmdStoredProc
			cmdDeleteObj.CommandText = "sma_sp_userid.spk_sma_cust_inter.sp_sl_delete"
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_service_location_id", adNumeric, adParamInput, , clng(lngServLocID))					'number(9)
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_last_update", adDBTimeStamp, adParamInput, ,Cdate(Request("hdnUpdateDateTime")))		'Date
            cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("P_USER_ID", adVarChar , adParamInput, 30, strRealUserID)
			cmdDeleteObj.Execute
			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			end if

			lngServLocID = NO_ID

			StrWinMessage = "Record deleted successfully."
    case "DELETESite"
    
    Dim SiteID
    
    Set SiteID=Request("SiteID")
    strSQL = "DELETE CRP.SITE_NAME_CODE " &_
					 "where SITE_ID = " & SiteID & ""
    set cmdUpdateObj = server.CreateObject("ADODB.Command")
			set cmdUpdateObj.ActiveConnection = objConn
			cmdUpdateObj.CommandType = adCmdText
			cmdUpdateObj.CommandText = strSQL
            cmdUpdateObj.Execute
			'response.write(strSQL)
			'response.end
			
			
			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			end if

			lngSiteID = NO_ID

			StrWinMessage = "Record deleted successfully."
end select
 
if lngServLocID <> NO_ID then
   strsql = "select c.customer_id, " &_
				   "c.customer_name, " &_
				   "s.service_location_id, " &_
	   			   "s.service_location_name, " &_
	   			   "s.specific_location_desc, " &_
	   			   "s.access_information, " &_
	   			   "s.accessible_schedule_id, " &_
				   "s1.schedule_name, " &_
	   			   "s.record_status_ind, " &_
	   			   "to_char(s.create_date_time, 'MON-DD-YYYY HH24:MI:SS') as create_date, " &_
	   			   "sma_sp_userid.spk_sma_library.sf_get_full_username(s.create_real_userid) as create_real_userid, " &_
	   			   "to_char(s.update_date_time, 'MON-DD-YYYY HH24:MI:SS') as update_date , " &_
	   			   "sma_sp_userid.spk_sma_library.sf_get_full_username(s.update_real_userid) as update_real_userid, " &_
	   			   "s.update_date_time as last_update_date_time, " &_
	   			   "a.address_id, " &_
	   			   "a.building_name,  " &_
	   			   "a.street,  " &_
	   			   "a.municipality_name, " &_
	   			   "p.province_state_name, " &_
	   			   "t.country_desc,  " &_
	   			   "m.clli_code, " &_
	   			   "a.province_state_lcode, " &_
	   			   "s.accessible_schedule_id, " &_
	   			   "s.comments " &_
			   "from crp.service_location s, " &_
	 			   "crp.v_address_consolidated_street a, " &_
	 			   "crp.customer c, " &_
	 			   "crp.lcode_province_state p, " &_
	 			   "crp.lcode_country t, " &_
	 			   "crp.municipality_lookup m, " &_
	 			   "crp.schedule s1 " &_
			   "where s.address_id = a.address_id " &_
			   "and	  s.customer_id = c.customer_id  " &_
			   "and	  a.province_state_lcode = p.province_state_lcode " &_
			   "and	  a.country_lcode = t.country_lcode " &_
			   "and	  t.country_lcode = p.country_lcode " &_
			   "and	  a.municipality_name = m.municipality_name " &_
			   "and	  a.province_state_lcode = m.province_state_lcode  " &_
			   "and   s.accessible_schedule_id = s1.schedule_id(+)  " &_
			   "and   s.service_location_id = " &  lngServLocID

   'Create the command object
     
   set objCmd = Server.CreateObject("ADODB.command")
       objCmd.ActiveConnection = objconn
	   objCmd.CommandText = strSql
	   objCmd.CommandType = adCmdText

   'Create Recordset object
     
   set objRsServiceLocation = objCmd.Execute

 	dim address
		if len(objRsServiceLocation("building_name")) > 0 then
			address = objRsServiceLocation("building_name") & vbNewLine &_
			objRsServiceLocation("street") & vbNewLine &_
			objRsServiceLocation("municipality_name") & ", " &_
			objRsServiceLocation("province_state_name") & vbNewLine &_
			objRsServiceLocation("country_desc")
		else
			address = objRsServiceLocation("street") & vbNewLine &_
			objRsServiceLocation("municipality_name")  & ", " &_
			objRsServiceLocation("province_state_name") & vbNewLine &_
			objRsServiceLocation("country_desc")
		end if

		strSQL = "select GEOCODEID_LCODE from crp.SERVICE_LOCATION_GEOCODE " &_
				 "where SERVICE_LOCATION_ID = " & lngServLocID &""
		'response.write(strSQL)
		set objSlGeoRs = objConn.Execute(strSQL)
		geoclliid = objSlGeoRs("GEOCODEID_LCODE")

 end if

 'response.write(objRsServiceLocation("street"))
 'response.end

 if geoclliid <> 0 then

   	strSQL = "select CLLI_CODE, GEOCODEID_LCODE as geocodeid,  DESCRIPTION, "&_
			"ADDRESS, CITY, PROVINCE, POSTAL_CODE " &_
			"FROM CRP.LCODE_GEOCODEID where GEOCODEID_LCODE  = "  & geoclliid &""
    set objGeoRs = objConn.Execute(StrSql)
		geocllicode = objGeoRs("clli_code") & vbNewLine &_
		objGeoRs("geocodeid") & ", " &_
        objGeoRs("description") & vbNewLine &_
		objGeoRs("address") & vbNewLine &_
		objGeoRs("city") & ", " &_
		objGeoRs("province") & vbNewLine &_
		objGeoRs("postal_code")

' 	response.write(geocllicode)
 '	response.end
 end if

strSQL = "select schedule_id, " &_
	            "schedule_name " &_
		 "from crp.schedule " &_
		 "where record_status_ind = 'A' " &_
		 "order by schedule_name"
     

set objRsSchedule = objconn.execute(strSQL)
    Dim aList1
    Dim m
    Dim s
    set s=0
    'set i=0,j=2
     if not objRsSchedule.EOF then
		aList1 = objRsSchedule.GetRows
    m=Ubound(aList1,2)
	else
		
	end if
     
    strSiteNameCodeSelect = "select SITE_NAME,SITE_CODE,Site_Id from CRP.SITE_NAME_CODE where SERVICE_LOCATION_ID = "&lngServLocID
    Dim objRs,Recordcnt,aList,k,n
    'set k=0,n=2
	set objRS = objconn.execute(strSiteNameCodeSelect)
	
    'objSites  = objconn.execute(strSiteNameCodeSelect)
    if not objRS.EOF then
		aList = objRS.GetRows
    n=Ubound(aList)
	else
		
	end if

  
%>

<html>
<head>
    <meta http-equiv="Pragma" content="no-cache">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <link rel="stylesheet" type="text/css" href="stylesheets/styles.css">
    <script type="text/javascript" src="GeneralJavaFunctions.js"></script>
    <script type="text/javascript" src="AccessLevels.js"></script>
    <script id="clientEventHandlersJS" language="javascript">
<!--
    //*******************************************************************************************
    var boolNeedToSave = false;
    var intAccessLevel = '<%=intAccessLevel%>';
    var intChildAccessLevel = '<%=intChildAccessLevel%>';
    var oldHighlightedElement;
    var oldHighlightedElementClassName;

    //*******************************************************************************************

    //set section title
    setPageTitle("SMA - Service Location");

    //*******************************************************************************************

    function iFrame_display()
    {
        //called whenever a refresh of the iFrame is needed
        document.getElementById("aifr").src = 'ServLocContact.asp?ServLocID=' + '<%=lngServLocID%>';
    }

    //function iFrame_display()
    //{
    //    //called whenever a refresh of the iFrame is needed
    //    document.getElementById("aifrsite").src   = 'Site Name and Code Maintenance.asp?SiteID=' + '<%=lngSiteID%>'+'ServLocID=' + '<%=lngServLocID%>';
    //}

    //*******************************************************************************************

    function btn_iFrmAdd()
    {

        if ((intChildAccessLevel & intConst_Access_Create) != intConst_Access_Create)
        {
            alert('Access denied.  Please contact your system administrator.');
            return false;
        }

        var NewWin;
        NewWin=window.open("ServLocContactDetail.asp?NewContact=NEW&ServLocID=" + document.frmServLocDetail.hdnServiceLocationID.value + "&CustName=" + document.frmServLocDetail.txtCustomerName.value ,"NewWin","toolbar=no,status=no,width=700,height=430,menubar=no resize=no");
        NewWin.focus();
    }

    //*******************************************************************************************

    function btn_iFrmUpdate(){

        var NewWin;

        if ((intChildAccessLevel & intConst_Access_Update) != intConst_Access_Update)
        {
            alert('Access denied.  Please contact your system administrator.');
            return false;
        }

        var doc;
        var iframeObject = document.getElementById('aifr'); // MUST have an ID
        if (iframeObject.contentDocument) { // DOM
            doc = iframeObject.contentDocument;
        } 
        else if (iframeObject.contentWindow) { // IE win
            doc = iframeObject.contentWindow.document;
        }

        
        if (doc.getElementsByName("hdnContactID")[0].value !="")
        {

            var strSource ="ServLocContactDetail.asp?ServLocContactID="+doc.getElementsByName("hdnContactID")[0].value;
            NewWin=window.open(strSource ,"NewWin","toolbar=no,status=no,width=700,height=430,menubar=no resize=no");
            NewWin.focus();
        }
        else
        {
            alert('You must select a record to update!');
        }

    }

    //*******************************************************************************************

    function cell_onClick(intSiteID){
       
        document.frmServLocDetail.hdnSiteID.value = intSiteID;
        		
        //highlight current record
        if (oldHighlightedElement != null) 
        {
            oldHighlightedElement.className = oldHighlightedElementClassName
        }
        oldHighlightedElement = window.event.srcElement.parentElement;
        oldHighlightedElementClassName = oldHighlightedElement.className;
        oldHighlightedElement.className = "Highlight";
    }

    //*******************************************************************************************

    function btn_iFrmDelete()
    {

        if ((intChildAccessLevel & intConst_Access_Delete) != intConst_Access_Delete)
        {
            alert('Access denied.  Please contact your system administrator.');
            return false;
        }


        var doc;
        var iframeObject = document.getElementById('aifr'); // MUST have an ID
        if (iframeObject.contentDocument) { // DOM
            doc = iframeObject.contentDocument;
        } 
        else if (iframeObject.contentWindow) { // IE win
            doc = iframeObject.contentWindow.document;
        }

        //  document.getElementById("aifr").src = document.location.href.replace("manobjdet", "manobjalias") ;
        if (doc.getElementsByName("hdnContactID")[0].value  !="")
        {
            if (confirm('Do you really want to delete this Contact?'))
            {
                document.getElementById("aifr").src  = "ServLocContact.asp?txtFrmAction=DELETE&ServLocID=<%=lngServLocID%>&ContactID="+doc.getElementsByName("hdnContactID")[0].value+"&hdnUpdateDateTime="+  doc.getElementsByName("hdnUpdateDateTime")[0].value ;
            }
        }
        else
        {
            alert('You must select a record to delete!');
        }
    }

    //*******************************************************************************************

    function window_onload() {
        iFrame_display();
        fct_displayStatus('<%=routineJavaScriptString(strWinMessage)%>');
    }

    //*******************************************************************************************

    function fct_clearStatus() {
        window.status = "";
    }

    //*******************************************************************************************

    function fct_displayStatus(strWinStatus){
        window.status=strWinStatus;
        setTimeout('fct_clearStatus()',5000);
    }

    //*******************************************************************************************

    function btnReferences_onclick()
    {

        if ('<%=lngServLocID%>' != '<%=NO_ID%>')
        {
            var strOwner = 'CRP' ;			// owner name must be in Uppercase
            var strTableName = 'SERVICE_LOCATION' ;		// replace ADDRESS with your own table name and table name must be in Uppercase
            var strRecordID = document.frmServLocDetail.hdnServiceLocationID.value ;   // insert your record id
            var URL ;

            URL ='Dependency.asp?Owner=' + strOwner + '&TableName=' + strTableName + '&RecordID='+ strRecordID;
            window.open(URL, 'Popup', 'top=100, left=100, WIDTH=500, HEIGHT=300'  ) ;
        }
        else
        {
            alert('This is a new record, therefore there are no references.');
        }

    }

    //*******************************************************************************************

    function btnDelete_onclick() {

        if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete)
        {
            alert('Access denied. Please contact your system administrator.');
            return false;
        }

        var logServLocID = document.frmServLocDetail.hdnServiceLocationID.value ;
        var strUpdateDateTime = document.frmServLocDetail.hdnUpdateDateTime.value ;

        if (logServLocID != "<%=NO_ID%>")
        {
            if (confirm("Do you really want to delete this service location?"))
            {
                boolNeedToSave = false;
                document.location = "<%=ASP_NAME%>?hdnFrmAction=DELETE&ServLocID=" + logServLocID + "&hdnUpdateDateTime=" + strUpdateDateTime ;
            }
        }

    }

    function btnSiteDelete_onclick()
    {
        if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete)
        {
            alert('Access denied. Please contact your system administrator.');
            return false;
        }

        var logSiteID = document.frmServLocDetail.hdnSiteID.value ;
        var logServLocID = document.frmServLocDetail.hdnServiceLocationID.value ;
        var strUpdateDateTime = document.frmServLocDetail.hdnUpdateDateTime.value ;

        if (logSiteID != "")
        {
            if (confirm("Do you really want to delete this site name/code?"))
            {
                boolNeedToSave = false;
                document.location = "<%=ASP_NAME%>?hdnFrmAction=DELETESite&SiteID=" + logSiteID+"&ServLocID=" + logServLocID + "&hdnUpdateDateTime=" + strUpdateDateTime ;
            }
        }
        else{
            alert("You cannot delete a  site name record. You must save the service location object first.")
        }
    }

    function btnSiteUpdate_onclick()
    {
        if ((intChildAccessLevel & intConst_Access_Update) != intConst_Access_Update)
        {
            alert('Access denied.  Please contact your system administrator.');
            return false;
        }
        var logServLocID = document.frmServLocDetail.hdnServiceLocationID.value ;

        if(!logServLocID)
        {
            alert("You cannot update a  site name record. You must save the service location object first.")
        }else{
            if (document.getElementById("hdnSiteID").value !="")
            {
                
                var strSource ="Site Name and Code Maintenance.asp?SiteID="+document.getElementById("hdnSiteID").value + "&ServLocID=" + logServLocID;
                NewWin=window.open(strSource ,"NewWin","toolbar=no,status=no,width=700,height=430,menubar=no resize=no");
                NewWin.focus();
            }
            else
            {
                alert('You must select a record to update!');
            }
        }

    }

    function btnSiteAddNew_onclick()
    {
        var logServLocID = document.frmServLocDetail.hdnServiceLocationID.value ;
        if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create)
        {
            alert('Access denied.  Please contact your system administrator.');
            return false;       }
        if(!logServLocID)
        {
            alert("You cannot create a site name record. You must save the service location object first.")
        }else{
            var strSource ="Site Name and Code Maintenance.asp?ServLocID=" + logServLocID ;
            NewWin=window.open(strSource ,"NewWin","toolbar=no,status=no,width=700,height=430,menubar=no resize=no");
            NewWin.focus();
        }
    }
    //*******************************************************************************************

    function btnReset_onclick()
    {
        var logServLocID = document.frmServLocDetail.hdnServiceLocationID.value ;

        if(confirm('All changes will be lost. Do you really want to reset the page?')){
            boolNeedToSave = false;
            document.location.href = '<%=ASP_NAME%>?ServLocID=' + logServLocID ;
        }
    }

    //*******************************************************************************************

    function form_onsubmit(){
       
        //no need to validate if the user cannot save the record
        if ( ((<%=intAccessLevel%> & <%=intconst_Access_Create%>) == <%=intconst_Access_Create%>) || ( (<%=intAccessLevel%> & <%=intconst_Access_Update%>) == <%=intconst_Access_Update%>) )
        {
            if (document.frmServLocDetail.hdnCustomerID.value == "" )
            {
                alert("Please select a customer using lookup function");
                document.frmServLocDetail.btnCustomerLookup.focus();
                return(false);
            }

            if (document.frmServLocDetail.hdnAddressID.value == "" )
            {
                alert("Please select an address using lookup function");
                document.frmServLocDetail.btnAddressLookup.focus();
                return(false);
            }

            if (document.frmServLocDetail.txtServiceLocationName.value == "" )
            {
                alert("Please type a service location name or generate on using the guess function");
                document.frmServLocDetail.btnGuess.focus();
                return(false);
            }

            if (document.frmServLocDetail.txtAccessInfo.value.length > 2000)
            {
                alert('The specified access information can be at most 2000 characters.\n\nYou entered ' + document.frmServLocDetail.txtAccessInfo.value.length + ' character(s).');
                document.frmServLocDetail.txtAccessInfo.focus();
                return false;
            }

            if (document.frmServLocDetail.txtComments.value.length > 2000)
            {
                alert('The comments for this service location can be at most 2000 characters.\n\nYou entered ' + document.frmServLocDetail.txtComments.value.length + ' character(s).');
                document.frmServLocDetail.txtComments.focus();
                return false;
            }

        }
        else
        {
            alert('Access denied.  Please contact your system administrator.');
            return (false);
        }

        document.frmServLocDetail.hdnFrmAction.value = "SAVE"
        boolNeedToSave = false;
        document.forms[0].submit();
        return(true);

    }

    //*******************************************************************************************

    function btnAddressLookup_onclick() {
        //***************************************************************************************************
        // Function:	btnAddressLookup_onclick															*
        //																									*
        // Purpose:		To display Address Search page with pre-populated customer name and to indicate		*
        //				that the search page is displayed in a popup window. (Note: search pages behave		*
        //				differently when displayed in popup windows verses when displayed in the base window)
        //																									*
        // Created By:	Sara Sangha Aug. 25th, 2000															*
        //																									*
        // Updated By:																						*
        //***************************************************************************************************

        var strCustomerName  = window.frmServLocDetail.txtCustomerName.value ;

        if (strCustomerName != "" ) {
            SetCookie("CustomerName", strCustomerName) ;

        }

        SetCookie("WinName", 'Popup');
        window.open('SearchFrame.asp?fraSrc=Address', 'Popup', 'top=50, left=100, WIDTH=800, HEIGHT=600' ) ;
    }

    function btnGeocodeLookup_onclick() {
        //***************************************************************************************************
        // Function:	btnGeocodeLookup_onclick															*
        //***************************************************************************************************

        //var strGeocodeid  = window.frmServLocDetail.txtCustomerName.value ;
        var strGeocodeid  = document.frmServLocDetail.hdnGeocode.value ;
        var strStreet = document.frmServLocDetail.hdnStreetName.value;
        var strCity = document.frmServLocDetail.hdnmunicipality_name.value;


        if (strStreet != "" ) {
            SetCookie("GeoStreet", strStreet) ;
        }

        if (strCity != "" ) {
            SetCookie("GeoCity", strCity) ;
        }

        SetCookie("WinName", 'Popup');
        window.open('SearchFrame.asp?fraSrc=Geocode', 'Popup', 'top=50, left=100, WIDTH=800, HEIGHT=600' ) ;
    }



    //*******************************************************************************************

    function btnAddNew_onclick()
    {

        if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create)
        {
            alert('Access denied.  Please contact your system administrator.');
            return false;
        }


        self.document.location.href = "<%=ASP_NAME%>?NewServLoc=NEW";

    }

    //*******************************************************************************************

    function btnCustomerLookup_onclick(CustService) {

        var strCustomerName = window.frmServLocDetail.txtCustomerName.value ;

        if (CustService != ""){SetCookie("ServiceEnd",CustService)};

        if (strCustomerName != "" )
        {
            SetCookie("CustomerName", strCustomerName);

        }


        SetCookie("WinName", 'Popup');
        window.open('SearchFrame.asp?fraSrc=Cust', 'Popup', 'top=50, left=100, WIDTH=800, HEIGHT=600'  ) ;

    }

    //*******************************************************************************************

    function selNavigate_onchange(){
        //**********************************************************************************************
        // Function:	selNavigate_onchange
        //
        // Purpose:		To display the page selected by the user from Quick Navigation drop-down box.	*
        //																								*
        // Created By:	Sara Sangha	Aug. 25th, 2000														*
        //																								*
        // Updated By:																					*
        //***********************************************************************************************
        var strPageName = document.frmServLocDetail.selNavigate.item(document.frmServLocDetail.selNavigate.selectedIndex).value ;
        var strCustomerID = document.frmServLocDetail.hdnCustomerID.value ;
        var strCustomerName = document.frmServLocDetail.txtCustomerName.value ;
        var strAddressID = document.frmServLocDetail.hdnAddressID.value ;
        var strServiceLocationName = document.frmServLocDetail.txtServiceLocationName.value ;
        var strGeocode = document.frmServLocDetail.hdnGeocode.value ;


        switch ( strPageName ) {

            case 'Address':
                document.frmServLocDetail.selNavigate.selectedIndex=0;
                self.location.href  = 'AddressDetail.asp?AddressID=' + strAddressID ;
                break ;

            case 'Cust' :
                document.frmServLocDetail.selNavigate.selectedIndex=0;
                self.location.href  = 'CustDetail.asp?CustomerID=' + strCustomerID ;
                break ;

            case 'CustServ':
                document.frmServLocDetail.selNavigate.selectedIndex=0;
                if (strServiceLocationName != ""){SetCookie("ServLocName", strServiceLocationName)};
                self.location.href = 'SearchFrame.asp?fraSrc=CustServ' ;
                break ;

            case 'Facility' :
                //alert("Go to Facility not implemented yet");
                document.frmServLocDetail.selNavigate.selectedIndex=0;
                if (strServiceLocationName != ""){SetCookie("ServLocName", strServiceLocationName)};
                self.location.href = 'SearchFrame.asp?fraSrc=' + strPageName ;
                break ;

            case 'ManagedObjects':  //to a list
                document.frmServLocDetail.selNavigate.selectedIndex=0;
                if (strServiceLocationName != ""){SetCookie("ServLocName", strServiceLocationName)};
                self.location.href = "SearchFrame.asp?fraSrc=" + strPageName  ;
                break;

            case 'FacilityPVC' :
                document.frmServLocDetail.selNavigate.selectedIndex=0;
                if (strServiceLocationName != ""){SetCookie("ServLocName", strServiceLocationName)};
                self.location.href = 'SearchFrame.asp?fraSrc=' + strPageName ;
                break ;

            case 'DEFAULT' :
                // do nothing ;
        }


    }

    //*******************************************************************************************

    function btnGuess_onclick() {

        var strClliCode = document.frmServLocDetail.hdnClliCode.value ;
        var strStreet = document.frmServLocDetail.hdnStreetName.value ;
        var strProvince= document.frmServLocDetail.hdnProvinceCode.value ;
        var strLen ;
        var strSuggestedName ;

        strLen = strStreet.length ;
        if (strLen > 42 ) {
            strStreet = strStreet.substr(0, 41) ;
        }

        strLen = strProvince.length ;
        if (strLen > 2 ) {
            strProvince = strProvince.substr(0, 1) ;
        }

        if (strProvince == "QC") {
            strProvince = "PQ";
        }

        if (strProvince == "NL") {
            strProvince = "NF";
        }

        strSuggestedName = strClliCode + strProvince + '_' + strStreet;
        document.frmServLocDetail.txtServiceLocationName.value  = strSuggestedName ;
        //document.frmServLocDetail.textGeocllicode.value = geocllicode ;
    }

    //*******************************************************************************************

    function on_change()
    {
        boolNeedToSave = true;
    }

    //*******************************************************************************************

    function window_unload()
    {
        //must set focus to save button because if user has changed only one field and has not
        //left it the on_change event will not have fired and the flag that determines whether
        //you need to save or not will be false
        document.frmServLocDetail.btnSave.focus();

        if ((intAccessLevel & intConst_Access_Update) == intConst_Access_Update || (intAccessLevel & intConst_Access_Create) == intConst_Access_Create)
        {
            if (boolNeedToSave == true)
            {
                event.returnValue = "There is unsaved data on the screen. To save changes, click CANCEL below then click SAVE on the main form.";
            }
        }
    }

    //*****************************************End of Java Functions*********************************
    //-->

    </script>
    <style type="text/css">
        .auto-style1 {
            height: 109px;
            width: 356px;
        }

        .auto-style2 {
            width: 356px;
        }

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
    </style>
</head>
<body language="javascript" onload="return window_onload()" onbeforeunload="return window_unload();">
    <form name="frmServLocDetail" action="<%=ASP_NAME%>" method="POST">
        <!-- hidden variables -->
        <input id="hdnCustomerID" name="hdnCustomerID" type="hidden" value="<%if lngServLocID <> NO_ID then Response.Write objRsServiceLocation("customer_id") else Response.Write null end if%>">
        <input id="hdnAddressID" name="hdnAddressID" type="hidden" value="<%if lngServLocID <> NO_ID then Response.Write objRsServiceLocation("address_id") else Response.Write null end if%>">
        <input id="hdnServiceLocationID" name="hdnServiceLocationID" type="hidden" value="<%if lngServLocID <> NO_ID then Response.Write objRsServiceLocation("service_location_id") else Response.Write null end if%>">
        <input id="hdnAccessibleScheduleID" name="hdnAccessibleScheduleID" type="hidden" value="<%if lngServLocID <> NO_ID then Response.Write objRsServiceLocation("accessible_schedule_id") else Response.Write null end if%>">
        <input id="hdnClliCode" name="hdnClliCode" type="hidden" value="<%if lngServLocID <> NO_ID then Response.Write routineHTMLString(objRsServiceLocation("clli_code")) else Response.Write null end if%>">
        <input id="hdnGeocode" name="hdnGeocode" type="hidden" value="<%if lngServLocID <> NO_ID then Response.Write(strGeocode) else Response.Write null end if%>">
        <input id="hdnStreetName" name="hdnStreetName" type="hidden" value="<%if lngServLocID <> NO_ID then Response.Write routineHTMLString(objRsServiceLocation("street")) else Response.Write null end if%>">
        <input id="hdnmunicipality_name" name="hdnmunicipality_name" type="hidden" value="<%if lngServLocID <> NO_ID then Response.Write routineHTMLString(objRsServiceLocation("municipality_name")) else Response.Write null end if%>">
        <input id="hdnProvinceCode" name="hdnProvinceCode" type="hidden" value="<%if lngServLocID <> NO_ID then Response.Write routineHTMLString(objRsServiceLocation("province_state_lcode")) else Response.Write null end if%>">
        <input id="hdnUpdateDateTime" name="hdnUpdateDateTime" type="hidden" value="<%if lngServLocID <> NO_ID then Response.Write objRsServiceLocation("last_update_date_time") else Response.Write null end if%>">
        <input id="hdnFrmAction" name="hdnFrmAction" type="hidden" value="">
        <input id="hdnSiteID" name="hdnSiteID" type="hidden" value="">

        <table border="0">
            <thead>
                <tr>
                    <td colspan="2" align="left">Service Location Detail</td>
                    <td>
                        <select align="right" valign="top" id="selNavigate" name="selNavigate" language="javascript" onchange="return selNavigate_onchange()" <%if lngServLocID = NO_ID then Response.Write " disabled " end if%> tabindex="18">
                            <option value="DEFAULT">Quickly Goto ...</option>
                            <option value="Address">Address</option>
                            <option value="Cust">Customer</option>
                            <option value="CustServ">Customer Service</option>
                            <option value="Facility">Facility</option>
                            <option value="ManagedObjects">Managed Object</option>
                            <option value="FacilityPVC">PVC</option>
                        </select>
                    </td>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td align="right">Customer Name<font color="red">*</font></td>
                    <td align="left">
                        <input id="txtCustomerName" name="txtCustomerName" disabled style="height: 22px; width: 350px" value="<%if lngServLocID <> NO_ID then Response.Write routineHTMLString(objRsServiceLocation("customer_name")) else Response.Write null end if%>">
                        <input id="btnCustomerLookup" name="btnCustomerLookup" style="height: 23px; width: 19px" type="button" value="..." language="javascript" onclick="on_change(); return btnCustomerLookup_onclick('C');" tabindex="1">
                    </td>
                    <td>&nbsp;</td>
                </tr>
                <tr>
                    <td align="right" valign="top">Address<font color="red">*</font></td>
                    <td align="left" valign="top">
                        <textarea style="width: 350px" rows="4" cols="20" id="textAddress" disabled name="textAddress"><%if lngServLocID <> NO_ID then Response.write routineHTMLString(address) else Response.Write null end if%></textarea>
                        <input id="btnAddressLookup" name="btnAddressLookup" style="height: 23px; width: 19px" type="button" value="..." language="javascript" onclick="on_change(); return btnAddressLookup_onclick()" tabindex="2">
                    </td>
                    <td>
                        <table border="1">
                            <thead>


                                <tr>
                                    <th class="auto-style2">SITE_CODE</th>
                                    <th>SITE_NAME</th>
                                </tr>
                            </thead>
                            <tbody>

                                <%   
                                for k=0 to n+1
                                    Response.Write "<tr>"&vbCrLf
                                  
                                     Response.Write "<td onclick= ""cell_onClick(" & routineHTMLString(alist(2,k)) & ")"" >" & routineHTMLString(alist(1,k)) &  "</td>" &vbCrLf
                                    Response.Write  "<td onclick= ""cell_onClick(" & routineHTMLString(alist(2,k)) & ")"" >" & routineHTMLString(alist(0,k)) & "</td>"   &vbCrLf
                                    Response.Write "</tr>"
                                   
                                  next  
                                %>


                                <tr>
                                    <td align="left" style="text-align: left;" colspan="2">
                                        <br />
                                        <br />
                                        <br />
                                        <br />
                                        <!--<input name="btnReferences" tabindex="9" type="button" value="References" style="width: 2.2cm" onclick="return btnReferences_onclick()">&nbsp;&nbsp;-->
                                        <input name="btnDelete" tabindex="10" type="button" value="Delete" style="width: 2cm" onclick="return btnSiteDelete_onclick()">&nbsp;&nbsp;
			<input name="btnReset" tabindex="11" type="button" value="Update" style="width: 2cm" onclick="return btnSiteUpdate_onclick();">&nbsp;&nbsp;
			<input name="btnAddNew" tabindex="12" type="button" value="New" style="width: 2cm" onclick="return btnSiteAddNew_onclick()">&nbsp;&nbsp;                                    </td>
                                </tr>
                            </tbody>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td align="right">Service Location Name<font color="red">*</font></td>
                    <td align="left">
                        <input id="txtServiceLocationName" name="txtServiceLocationName" tabindex="3" style="height: 22px; width: 350px" onchange="return on_change();" value="<% if lngServLocID <> NO_ID then Response.Write routineHTMLString(objRsServiceLocation("service_location_name").value) else Response.write null end if%>">
                        <input id="btnGuess" name="btnGuess" style="height: 23px; width: 50px" type="button" value="Guess" language="javascript" onclick="return btnGuess_onclick()" tabindex="4">
                    </td>
                    </TD>
		<td>&nbsp;</td>
                </tr>
                <tr>
                    <td align="right">Specific Location Desc</td>
                    <td align="left">
                        <input id="txtSpecificLocation" name="txtSpecificLocation" tabindex="5" style="height: 22px; width: 350px" onchange="return on_change();" value="<%if lngServLocID <> NO_ID then Response.write routineHTMLString(objRsServiceLocation("specific_location_desc").value) else Response.Write null end if%>"></td>
                    <td>&nbsp;</td>
                </tr>
                <tr>
                    <td align="right">Access Information</td>
                    <td align="left">
                        <textarea rows="2" cols="20" id="txtAccessInfo" tabindex="6" name="txtAccessInfo" onchange="return on_change();" style="width: 350px"><%if lngServLocID <> NO_ID then Response.write routineHTMLString(objRsServiceLocation("access_information")) else Response.Write null end if%></textarea>
                    <td>&nbsp;</td>
                </tr>
                <tr>
                    <td align="right" valign="top">CLLI CODE</td>
                    <td align="left" valign="top">
                        <textarea style="width: 350px" rows="4" cols="20" id="textGeocllicode" disabled name="textGeocllicode"><%if lngServLocID <> NO_ID then Response.write routineHTMLString(geocllicode) else Response.Write null end if%></textarea>
                        <input id="btnGeocodeLookup" name="btnGeocodeLookup" style="height: 23px; width: 19px" type="button" value="..." language="javascript" onclick="on_change(); return btnGeocodeLookup_onclick()" tabindex="2">
                    </td>
                    <td>&nbsp;</td>
                </tr>

                <tr>
                    <td align="right">Schedule Info</td>
                    <td align="left">
                        <select name="selSchedule" tabindex="7" style="height: 22px; width: 350px" onchange="return on_change();">
                            <option></option>
                            <%
                               for s=0 to m+1
                                                            Response.write "<OPTION "
					if lngServLocID <> NO_ID then 'only select an option if there is an existing service location to edit.
						if objRsServiceLocation("accessible_schedule_id").Value <> "" then
							if Cint(objRsServiceLocation("accessible_schedule_id").Value) = Cint(objRsSchedule("SCHEDULE_ID").Value) then
								Response.Write " SELECTED "
							END IF
						END IF
					end if
					Response.Write 	" VALUE=" & routineHTMLString(aList1(0,s)) & ">" & routineHTMLString(aList1(1,s)) & "</OPTION>" &vbCrLf
					next
                            %>
                        </select>

                    </td>
                    <td>&nbsp;</td>
                </tr>


                <tr>
                    <td align="right">Comments</td>
                    <td align="left">
                        <textarea rows="2" cols="20" id="txtComments" name="txtComments" tabindex="8" style="width: 350px" onchange="return on_change();"><%if lngServLocID <> NO_ID then Response.write routineHTMLString(objRsServiceLocation("comments")) else Response.Write null end if%></textarea></td>
                    <td>&nbsp; </td>
                </tr>
            </tbody>
            <tfoot>
                <tr>
                    <td align="right" colspan="3">
                        <input name="btnReferences" tabindex="9" type="button" value="References" style="width: 2.2cm" onclick="return btnReferences_onclick()">&nbsp;&nbsp;
			<input name="btnDelete" tabindex="10" type="button" value="Delete" style="width: 2cm" onclick="return btnDelete_onclick()">&nbsp;&nbsp;
			<input name="btnReset" tabindex="11" type="button" value="Reset" style="width: 2cm" onclick="return btnReset_onclick();">&nbsp;&nbsp;
			<input name="btnAddNew" tabindex="12" type="button" value="New" style="width: 2cm" onclick="return btnAddNew_onclick()">&nbsp;&nbsp;
			<input name="btnSave" tabindex="13" type="button" value="Save" style="width: 2cm" onclick="return form_onsubmit();">&nbsp;&nbsp;
                    </td>
                </tr>
            </tfoot>
        </table>
        <table>
            <thead>
                <tr>
                    <td colspan="4" align="left">Service Location Contacts</td>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td colspan="4">
                        <iframe id="aifr" width="100%" height="100" src="" scrolling="yes" marginheight="1" marginwidth="1"></iframe>
                        <br>
                        <!-- The following buttons are disabled if this is a new record -->
                        <input type="button" tabindex="14" value="Delete" <%if lngServLocID <> NO_ID then  Response.Write null else Response.Write "DISABLED" end if%> name="btn_iFrameDelete" onclick="btn_iFrmDelete();" style="width: 2cm">&nbsp;&nbsp;
			<input type="button" tabindex="15" value="Refresh" <%if lngServLocID <> NO_ID then  Response.Write null else Response.Write "DISABLED" end if%> name="btn_iFrameRefresh" onclick="iFrame_display();" style="width: 2cm">&nbsp;&nbsp;
			<input type="button" tabindex="16" value="New" <%if lngServLocID <> NO_ID then  Response.Write null else Response.Write "DISABLED" end if%> name="btn_iFrameAdd" onclick="btn_iFrmAdd(); " style="width: 2cm">&nbsp;&nbsp;
			<input type="button" tabindex="17" value="Update" <%if lngServLocID <> NO_ID then  Response.Write null else Response.Write "DISABLED" end if%> name="btn_iFrameupdate" onclick="btn_iFrmUpdate();" style="width: 2cm">&nbsp;&nbsp;

                    </td>
                </tr>
            </tbody>
        </table>
        <fieldset width="100%">
            <legend align="right"><b>Audit Information</b></legend>
            <div size="8pt" align="RIGHT">
                Record Status Indicator
		<input align="left" name="txtRecordStatusInd" type="text" style="height: 20px; width: 18px" disabled value="<%if lngServLocID <> NO_ID then Response.Write objRsServiceLocation("record_status_ind") else Response.Write null end if%>">&nbsp;&nbsp;&nbsp;
		Create Date
		<input align="center" name="txtRecordStatusInd" type="text" style="height: 20px; width: 150px" disabled value="<%if lngServLocID <> NO_ID then Response.Write objRsServiceLocation("create_date") else Response.Write null end if%>">&nbsp;
		Created By
		<input align="right" name="txtRecordStatusInd" type="text" style="height: 20px; width: 200px" disabled value="<%if lngServLocID <> NO_ID then Response.Write objRsServiceLocation("create_real_userid") else Response.Write null end if%>"><br>
                Update Date
		<input align="center" name="txtRecordStatusInd" type="text" style="height: 20px; width: 150px" disabled value="<%if lngServLocID <> NO_ID then Response.Write objRsServiceLocation("update_date") else Response.Write null end if%>">
                Updated By
		<input align="right" name="txtRecordStatusInd" type="text" style="height: 20px; width: 200px" disabled value="<%if lngServLocID <> NO_ID then Response.Write objRsServiceLocation("update_real_userid") else Response.Write null end if%>">
            </div>
        </fieldset>

    </form>
</body>
</html>
<%  'clean up ADO objects
	if lngServLocID <> NO_ID then
		set objRsServiceLocation = nothing
		set objCmd = nothing
		objConn.close
		set objConn = nothing
	end if
%>