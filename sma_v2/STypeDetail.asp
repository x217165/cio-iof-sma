<%@  language="VBScript" %>
<%Option Explicit%>
<%Response.Buffer = True
 on error resume next%>
<!--#include file = "smaConstants.inc"-->
<!--#include file = "smaProcs.inc"-->
<!--#include file = "databaseconnect.asp" -->
<!--
*************************************************************************************************
* Page name:	STypeDetail.asp																	*
* Purpose:		To display the Service Type														*
*				Chosen via STypeList.asp														*
*																								*
* Created by:	Gilles Archer 09/27/2000														*
* Modifications By				Date				Modifcations								*
* Sara Sangha					02/15/2000			- Added an iFrame to display Default SLA for*
*													  different regions
* Anthony Cheung				10/06/2008			- Added an iFrame to display Service Type Attributes
*										  Added an iFrame to display Service Instance Attributes
*										  Added an iFrame to display Kenan Attributes
* Linda Chen					08/06/2009			- Display STID and if owned by NetCracker   *
*
* 																								*
*************************************************************************************************
-->
<%
Dim strServiceTypeID, datUpdateDateTime, strWinMessage, strWinLocation
Dim lRow, arrLOBList, arrCategoryList, arrClassList, arrLevelList

Dim arrVPNList

Dim	objRS, objRSFrench, objRSSelect, objCommand, strSQL, strErrMessage, lIndex
Dim p_service_type_french
Dim strEmailFrom, strEmailTo, strEmailCC, strEmailBCC, strEmailSubject, strEmailBody, strLANG
Dim datServiceStartDate, datServiceEndDate, selServiceCategory
Dim intAccessLevel
Dim strNC

	'Check user's rights
	intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
	' The following 3 lines temp commented for my test LC
	If (intAccessLevel And intConst_Access_ReadOnly) <> intConst_Access_ReadOnly  Then
		DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to Service Type. Please contact your system administrator"
	End If

	strWinMessage = "New Message"
	strServiceTypeID = Request("ServiceTypeID")
	strNC = Request("ncflag")

	'response.write(strNC)
	'response.end
	if len(trim(Request.Form("txtServiceTypeFrench"))) <> 0 Then
		p_service_type_french = trim(Request.Form("txtServiceTypeFrench"))
	Else
		p_service_type_french = "NULL"
	end if

	Select Case Request.Form("selmonth").Count
		Case 1
			datServiceStartDate = Request.Form("selmonth")(1)
		Case 2
			datServiceStartDate = Request.Form("selmonth")(1)
			datServiceEndDate = Request.Form("selmonth")(2)
	End Select

	Select Case Request.Form("selday").Count
		Case 1
			datServiceStartDate = datServiceStartDate & "/" & Request.Form("selday")(1)
		Case 2
			datServiceStartDate = datServiceStartDate & "/" & Request.Form("selday")(1)
			datServiceEndDate = datServiceEndDate & "/" & Request.Form("selday")(2)
	End Select

	Select Case Request.Form("selyear").Count
		Case 1
			datServiceStartDate = datServiceStartDate & "/" & Request.Form("selyear")(1)
		Case 2
			datServiceStartDate = datServiceStartDate & "/" & Request.Form("selyear")(1)
			datServiceEndDate = datServiceEndDate & "/" & Request.Form("selyear")(2)
	End Select

	If Len(datServiceStartDate) <> 10 Then datServiceStartDate = ""
	If Len(datServiceEndDate) <> 10 Then datServiceEndDate = ""

	selServiceCategory = Request("selServiceCategory")
	lIndex = InStr(1, selServiceCategory, "|", 0)
	If lIndex <> 0 Then
		selServiceCategory = Left(selServiceCategory, lIndex - 1)
	End If

	Select Case UCase(Request("hdnFrmAction"))
		Case "SAVE"
			Set objCommand = Server.CreateObject("ADODB.Command")
			Set objCommand.ActiveConnection = objConn
			objCommand.CommandType = adCmdStoredProc


			If IsNumeric(Request("hdnServiceTypeID")) Then	'Save existing Service Type
				If (intAccessLevel And intConst_Access_Update) <> intConst_Access_Update Then
					DisplayError "REFRESH", strWinLocation, 0, "UPDATE DENIED", "You don't have access to update service types. Please contact your system administrator"
				End If

				objCommand.CommandText = "SMA_SP_USERID.Spk_Sma_Admin_Inter.sp_servtype_update"

				objCommand.Parameters.Append objCommand.CreateParameter("p_userid", adVarChar, adParamInput, 20, Session("username"))
				objCommand.Parameters.Append objCommand.CreateParameter("p_service_type_id", adNumeric, adParamInput, , CLng(Request("hdnServiceTypeID")))
				objCommand.Parameters.Append objCommand.CreateParameter("p_service_type_desc", adVarChar, adParamInput, 80, Trim(Request("txtServiceDescription")))
				objCommand.Parameters.Append objCommand.CreateParameter("p_service_type_french", adVarChar, adParamInput, 80, p_service_type_french)
				objCommand.Parameters.Append objCommand.CreateParameter("p_service_cat_id", adNumeric, adParamInput, , selServiceCategory)
				objCommand.Parameters.Append objCommand.CreateParameter("p_sla_id", adNumeric, adParamInput, , 0)
				objCommand.Parameters.Append objCommand.CreateParameter("p_start_dt", adVarChar, adParamInput, 10, datServiceStartDate)
				objCommand.Parameters.Append objCommand.CreateParameter("p_end_dt", adVarChar, adParamInput, 10, datServiceEndDate)
				objCommand.Parameters.Append objCommand.CreateParameter("p_service_class", adVarChar, adParamInput, 8, Request("selServiceClass"))

				if Request("selVPNTypes") <>"" then
					objCommand.Parameters.Append objCommand.CreateParameter("p_vpn_type_code", adNumeric, adParamInput, , CLng(Request("selVPNTypes")))
				else
					objCommand.Parameters.Append objCommand.CreateParameter("p_vpn_type_code", adNumeric, adParamInput, , 0)
				end if

				objCommand.Parameters.Append objCommand.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , CDate(Request("hdnUpdateDateTime")))
				objCommand.Parameters.Append objCommand.CreateParameter("p_old_stype_desc", adVarChar, adParamOutput, 80, Null)
                objCommand.Parameters.Append objCommand.CreateParameter("p_lcode", adNumeric, adParamInput, , CiNt(Request("txtNCFlag")))
                'check parameter values
                'dim objparm
 				'for each objparm in objCommand.Parameters
 				'	  Response.Write "<b>" & objparm.name & "</b>"
 				'	  Response.Write " has size:  " & objparm.Size & " "
 				'	  Response.Write " and value:  " & objparm.value & " "
 				'	  Response.Write " and datatype:  " & objparm.type & "<br> "
 				'next
 				'response.write (objCommand.CommandText)

 				'Response.Write "<b> count = " & objCommand.Parameters.count & "<br>"
  				'dim nx
  				'for nx=0 to objCommand.Parameters.count-1
  				'   Response.Write nx+1 & " parm value= " & objCommand.Parameters.Item(nx).Value  & "<br>"
  				'next
  				'response.write (objCommand.CommandText)
                'response.end





				strErrMessage = "CANNOT UPDATE RECORD"

				On Error Resume Next
				objCommand.Execute
				If objConn.Errors.Count <> 0 Then
					DisplayError "BACK", "", objConn.Errors(0).NativeError, strErrMessage, objConn.Errors(0).Description
					objConn.Errors.Clear
				End If
				strServiceTypeID = CStr(objCommand.Parameters("p_service_type_id").Value)

				If Not IsNull(objCommand.Parameters("p_old_stype_desc").Value) Then
					'it's time to send an email
					strSQL = "SELECT " &_
							"ST.SERVICE_TYPE_DESC, " &_
							"SC.SERVICE_CATEGORY_DESC, " &_
							"SLA.SERVICE_LEVEL_AGREEMENT_DESC, " &_
							"ST.VPN_TYPE_LCODE " &_
							"FROM " &_
							"CRP.SERVICE_TYPE ST, " &_
							"CRP.SERVICE_CATEGORY SC, " &_
							"CRP.SERVICE_LEVEL_AGREEMENT SLA " &_
							"WHERE " &_
							"ST.SERVICE_CATEGORY_ID = SC.SERVICE_CATEGORY_ID AND " &_
							"ST.DEFAULT_SLA_ID = SLA.SERVICE_LEVEL_AGREEMENT_ID AND " &_
							"ST.SERVICE_TYPE_ID = " & strServiceTypeID

					'Create Recordset object
					Set objRS = Server.CreateObject("ADODB.Recordset")
					On Error Resume Next
					objRS.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
					If objConn.Errors.Count <> 0 Then
						DisplayError "BACK", "", objConn.Errors(0).NativeError, "ERROR LOADING DATA (Email Notification)", objConn.Errors(0).Description
						objConn.Errors.Clear
					End If

					strEmailSubject = "Notification of New/Changed Service Type"
					'Changed Service Type
					strEmailBody = "CHANGED" & vbCrLf
					strEmailBody = strEmailBody & "Old Description: " & objCommand.Parameters("p_old_stype_desc").Value & vbCrLf
					strEmailBody = strEmailBody & "New Description: " & objRS.Fields("SERVICE_TYPE_DESC").Value & vbCrLf
					strEmailBody = strEmailBody & "Service Category: " & objRS.Fields("SERVICE_CATEGORY_DESC").Value & vbCrLf
					strEmailBody = strEmailBody & "Default SLA: " & objRS.Fields("SERVICE_LEVEL_AGREEMENT_DESC").Value & vbCrLf

					objRS.Close
					Set objRS = Nothing

					Response.Cookies("txtEmailTo") = strConst_ServiceTypeEmailTo
					Response.Cookies("txtEmailSubject") = escape(strEmailSubject)
					Response.Cookies("txtEmailBody") = escape(strEmailBody)
				End If

				strWinMessage = "Record saved successfully. You can now see the changes you made."
			Else										'Create a new Service Type
				If (intAccessLevel And intConst_Access_Create) <> intConst_Access_Create Then
					DisplayError "REFRESH", strWinLocation, 0, "INSERT DENIED", "You don't have access to create service types. Please contact your system administrator"
				End If

				objCommand.CommandText = "SMA_SP_USERID.Spk_Sma_Admin_Inter.Sp_Servtype_Insert"

				objCommand.Parameters.Append objCommand.CreateParameter("p_userid", adVarChar, adParamInput, 20, Session("username"))
				objCommand.Parameters.Append objCommand.CreateParameter("p_service_type_id", adNumeric, adParamOutput, , Null)
				objCommand.Parameters.Append objCommand.CreateParameter("p_service_type_desc", adVarChar, adParamInput, 80, Trim(Request("txtServiceDescription")))
				objCommand.Parameters.Append objCommand.CreateParameter("p_service_type_french", adVarChar, adParamInput, 80, p_service_type_french)
				objCommand.Parameters.Append objCommand.CreateParameter("p_service_cat_id", adNumeric, adParamInput, , selServiceCategory)
				objCommand.Parameters.Append objCommand.CreateParameter("p_sla_id", adNumeric, adParamInput, , 0)
				objCommand.Parameters.Append objCommand.CreateParameter("p_start_dt", adVarChar, adParamInput, 10, datServiceStartDate)
				objCommand.Parameters.Append objCommand.CreateParameter("p_end_dt", adVarChar, adParamInput, 10, datServiceEndDate)
				objCommand.Parameters.Append objCommand.CreateParameter("p_service_class", adVarChar, adParamInput, 8, Request("selServiceClass"))

				if Request("selVPNTypes") <>"" then
					objCommand.Parameters.Append objCommand.CreateParameter("p_vpn_type_code", adNumeric, adParamInput, , Request("selVPNTypes"))
				else
					objCommand.Parameters.Append objCommand.CreateParameter("p_vpn_type_code", adNumeric, adParamInput, ,0)
				end if
                objCommand.Parameters.Append objCommand.CreateParameter("p_lcode", adNumeric, adParamInput, , CiNt(Request("txtNCFlag")))
				'check parameter values



               ' dim objparm
  				'for each objparm in objCommand.Parameters
  				'	  Response.Write "<b>" & objparm.name & "</b>"
  				'	  Response.Write " has size:  " & objparm.Size & " "
  				'	  Response.Write " and value:  " & objparm.value & " "
  				'	  Response.Write " and datatype:  " & objparm.type & "<br> "
  				'next

  				'Response.Write "<b> count = " & objCommand.Parameters.count & "<br>"
  				'dim nx
  				'for nx=0 to objCommand.Parameters.count-1
  				'   Response.Write nx+1 & " parm value= " & objCommand.Parameters.Item(nx).Value  & "<br>"
  				'next
  				'response.write (objCommand.CommandText)
               ' response.end



				strErrMessage = "CANNOT CREATE OBJECT"

				On Error Resume Next
				objCommand.Execute
				If objConn.Errors.Count <> 0 Then
					DisplayError "BACK", "", objConn.Errors(0).NativeError, strErrMessage, objConn.Errors(0).Description
					objConn.Errors.Clear
				End If
				strServiceTypeID = CStr(objCommand.Parameters("p_service_type_id").Value)

				'it's time to send an email
				strSQL = "SELECT " &_
						"ST.SERVICE_TYPE_DESC, " &_
						"SC.SERVICE_CATEGORY_DESC, " &_
						"SLA.SERVICE_LEVEL_AGREEMENT_DESC " &_
						"FROM " &_
						"CRP.SERVICE_TYPE ST, " &_
						"CRP.SERVICE_CATEGORY SC, " &_
						"CRP.SERVICE_LEVEL_AGREEMENT SLA " &_
						"WHERE " &_
						"ST.SERVICE_CATEGORY_ID = SC.SERVICE_CATEGORY_ID AND " &_
						"ST.DEFAULT_SLA_ID = SLA.SERVICE_LEVEL_AGREEMENT_ID AND " &_
						"ST.SERVICE_TYPE_ID = " & strServiceTypeID

				'Create Recordset object
				Set objRS = Server.CreateObject("ADODB.Recordset")
				On Error Resume Next
				objRS.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
				If objConn.Errors.Count <> 0 Then
					DisplayError "BACK", "", objConn.Errors(0).NativeError, "ERROR LOADING DATA (Email Notification)", objConn.Errors(0).Description
					objConn.Errors.Clear
				End If

				strEmailSubject = "Notification of New/Changed Service Type"
				'New Service Type
				strEmailBody = "NEW" & vbCrLf
				strEmailBody = strEmailBody & "Description: " & objRS.Fields("SERVICE_TYPE_DESC").Value & vbCrLf
				strEmailBody = strEmailBody & "Service Category: " & objRS.Fields("SERVICE_CATEGORY_DESC").Value & vbCrLf
				strEmailBody = strEmailBody & "Default SLA: " & objRS.Fields("SERVICE_LEVEL_AGREEMENT_DESC").Value & vbCrLf
				objRS.Close
				Set objRS = Nothing

				Response.Cookies("txtEmailTo") = strConst_ServiceTypeEmailTo
				Response.Cookies("txtEmailSubject") = escape(strEmailSubject)
				Response.Cookies("txtEmailBody") = escape(strEmailBody)

				strWinMessage = "Record saved successfully. You can now see the changes you made."
			End If


		Case "DELETE"
			If (intAccessLevel And intConst_Access_Delete) <> intConst_Access_Delete Then
				DisplayError "REFRESH", strWinLocation, 0, "DELETE DENIED", "You don't have access to delete service types. Please contact your system administrator"
			End If

			Set objCommand = Server.CreateObject("ADODB.Command")
			Set objCommand.ActiveConnection = objConn
			objCommand.CommandType = adCmdStoredProc
			objCommand.CommandText = "sma_sp_userid.spk_sma_admin_inter.sp_servtype_delete"
			objCommand.Parameters.Append objCommand.CreateParameter("p_service_type_id", adNumeric, adParamInput, , CLng(Request("hdnServiceTypeID")))					'number(9)
			objCommand.Parameters.Append objCommand.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , CDate(Request("hdnUpdateDateTime")))		'Date

  			On Error Resume Next
			objCommand.Execute
			If objConn.Errors.Count <> 0 Then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			End If
			strServiceTypeID = "DEL"
			strWinMessage = "Record deleted successfully."
	End Select

	If IsNumeric(strServiceTypeID) Then
		strSQL = "SELECT ST.SERVICE_TYPE_ID, " &_
			"ST.SERVICE_TYPE_DESC, " &_
			"LOB.LOB_ID, " &_
			"ST.SERVICE_CATEGORY_ID, " &_
			"ST.DEFAULT_SLA_ID, " &_
			"ST.SERVICE_TYPE_START_DATE, " &_
			"ST.SERVICE_TYPE_END_DATE, " &_
			"ST.SERVICE_TYPE_STATUS, " &_
			"ST.SERVICE_TYPE_SUCCESSOR, " &_
			"ST.SERVICE_TYPE_DETAIL_WEB_LINK, " &_
			"ST.CONTRACT_REQUIRED_FLAG, " &_
			"ST.SERVICE_CLASS_LCODE, " &_


			"ST.VPN_TYPE_LCODE,"&_

			"TO_CHAR(ST.CREATE_DATE_TIME, 'MON-DD-YYYY HH24:MI:SS') AS CREATE_DATE_TIME, " &_
			"sma_sp_userid.spk_sma_library.sf_get_full_username(ST.CREATE_REAL_USERID) AS CREATE_REAL_USERID, " &_
			"TO_CHAR(ST.UPDATE_DATE_TIME, 'MON-DD-YYYY HH24:MI:SS') AS UPDATE_DATE_TIME, " &_
			"sma_sp_userid.spk_sma_library.sf_get_full_username(ST.UPDATE_REAL_USERID) AS UPDATE_REAL_USERID, " &_
			"ST.RECORD_STATUS_IND, " &_
			"ST.UPDATE_DATE_TIME AS LAST_UPDATE_DATE_TIME " &_
			"FROM CRP.SERVICE_TYPE ST, CRP.SERVICE_CATEGORY SC, CRP.LOB LOB " &_
			"WHERE ST.SERVICE_CATEGORY_ID = SC.SERVICE_CATEGORY_ID " &_
			"AND SC.LOB_ID = LOB.LOB_ID " &_
			"AND ST.SERVICE_TYPE_ID = " & strServiceTypeID

		'Response.Write strSQL
		'Response.End

		'Create Recordset object
		Set objRS = Server.CreateObject("ADODB.Recordset")
		On Error Resume Next
		objRS.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
		If objConn.Errors.Count <> 0 Then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "ERROR LOADING DATA (Service Type)", objConn.Errors(0).Description
			objConn.Errors.Clear
		End If

		strSQL = " SELECT SERVICE_TYPE_ID, SERVICE_TYPE_LANG_DESC " &_
				 " FROM CRP.SERVICE_TYPE_LANG " &_
				 " WHERE SERVICE_TYPE_ID = " & strServiceTypeID &_
				 " AND RECORD_STATUS_IND = 'A' "

		'Create Recordset object
		Set objRSFrench = Server.CreateObject("ADODB.Recordset")
		On Error Resume Next
		objRSFrench.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
		If objConn.Errors.Count <> 0 Then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "ERROR LOADING DATA", objConn.Errors(0).Description
			objConn.Errors.Clear
		End If


	End If

	'Create Recordset object
	Set objRSSelect = Server.CreateObject("ADODB.Recordset")

	'TQ_INOSS
	' strLANG = Request.Cookies("UserInformation")("language_preference")
	strLANG="EN"
	if (Len(strLANG) = 0) then strLANG = "EN"

	'Get the Line of Business : TQ_INOSS
	strSQL = "SELECT LOB_ID, LOB_CODE, LOB_DESC " &_
			"FROM CRP.V_LOB " &_
			"WHERE lob_id NOT IN" &_
		        	"(SELECT lob_id " &_
		        	"FROM crp.v_lob " &_
		        	"WHERE language_preference_lcode = '" & strLANG & "' ) " &_
			"AND LANGUAGE_PREFERENCE_LCODE = 'EN'" &_
			"AND RECORD_STATUS_IND = 'A'" &_
			"UNION SELECT LOB_ID, LOB_CODE, LOB_DESC " &_
			"FROM crp.v_lob " &_
			"WHERE language_preference_lcode = '" & strLANG & "'" &_
			"AND RECORD_STATUS_IND = 'A' " &_
			"ORDER BY LOB_DESC ASC"


'Response.Write strSQL
'Response.End

	On Error Resume Next
	objRSSelect.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	If objConn.Errors.Count <> 0 Then
		DisplayError "BACK", "", objConn.Errors(0).NativeError, "ERROR LOADING DATA (Line of Business)", objConn.Errors(0).Description
		objConn.Errors.Clear
	End If
	arrLOBList = objRSSelect.GetRows
	objRSSelect.Close

	strSQL = " SELECT SERVICE_CATEGORY_ID " &_
			 " ,      LOB_ID " &_
			 " ,      SERVICE_CATEGORY_DESC " &_
			 " FROM CRP.V_SERVICE_CATEGORY " &_
			 " WHERE SERVICE_CATEGORY_ID NOT IN ( " &_
			 "     SELECT SERVICE_CATEGORY_ID " &_
			 "     FROM CRP.V_SERVICE_CATEGORY " &_
			 "     WHERE LANGUAGE_PREFERENCE_LCODE = '" & strLANG & "' " &_
			 " ) " &_
			 " AND LANGUAGE_PREFERENCE_LCODE = 'EN' " &_
			 " UNION " &_
			 " SELECT SERVICE_CATEGORY_ID " &_
			 " ,      LOB_ID " &_
			 " ,      SERVICE_CATEGORY_DESC " &_
			 "FROM CRP.V_SERVICE_CATEGORY " &_
			 "WHERE LANGUAGE_PREFERENCE_LCODE = '" & strLANG & "' " &_
			 "AND   RECORD_STATUS_IND = 'A' " &_
			 "ORDER BY SERVICE_CATEGORY_DESC ASC"

	On Error Resume Next
	objRSSelect.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	If objConn.Errors.Count <> 0 Then
		DisplayError "BACK", "", objConn.Errors(0).NativeError, "ERROR LOADING DATA (Service Category)", objConn.Errors(0).Description
		objConn.Errors.Clear
	End If
	arrCategoryList = objRSSelect.GetRows
	objRSSelect.Close

	strSQL = "SELECT SERVICE_CLASS_LCODE, SERVICE_CLASS_DESC " &_
			"FROM CRP.LCODE_SERVICE_CLASS " &_
			"WHERE RECORD_STATUS_IND = 'A' " &_
			"ORDER BY SERVICE_CLASS_DESC ASC"

	On Error Resume Next
	objRSSelect.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	If objConn.Errors.Count <> 0 Then
		DisplayError "BACK", "", objConn.Errors(0).NativeError, "ERROR LOADING DATA (Service Class)", objConn.Errors(0).Description
		objConn.Errors.Clear
	End If
	arrClassList = objRSSelect.GetRows
	objRSSelect.Close



	strSQL = "SELECT VPN_TYPE_LCODE, VPN_TYPE_DESC " &_
			"FROM CRP.LCODE_VPN_TYPE " &_
			"WHERE RECORD_STATUS_IND = 'A' " &_
			"ORDER BY VPN_TYPE_LCODE ASC"

	On Error Resume Next
	objRSSelect.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	If objConn.Errors.Count <> 0 Then
		DisplayError "BACK", "", objConn.Errors(0).NativeError, "ERROR LOADING DATA (Service Class)", objConn.Errors(0).Description
		objConn.Errors.Clear
	End If
	arrVPNList = objRSSelect.GetRows
	objRSSelect.Close





	strSQL = "SELECT SERVICE_LEVEL_AGREEMENT_ID, SERVICE_LEVEL_AGREEMENT_DESC " &_
			"FROM CRP.SERVICE_LEVEL_AGREEMENT " &_
			"WHERE RECORD_STATUS_IND = 'A' " &_
			"ORDER BY SERVICE_LEVEL_AGREEMENT_DESC ASC"

	On Error Resume Next
	objRSSelect.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	If objConn.Errors.Count <> 0 Then
		DisplayError "BACK", "", objConn.Errors(0).NativeError, "ERROR LOADING DATA (Service Level)", objConn.Errors(0).Description
		objConn.Errors.Clear
	End If
	arrLevelList = objRSSelect.GetRows
	objRSSelect.Close
	Set objRSSelect = Nothing


dim intRowCount, intColCount,strInnerValues
intRowCount = 0
intColCount = 3

strInnerValues = ""


%>
<html>
<head>
    <meta name="Generator" content="Microsoft Visual Studio 6.0">
    <meta http-equiv="Pragma" content="no-cache">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <link rel="stylesheet" type="text/css" href="stylesheets/styles.css">
    <script type="text/javascript" language="javascript" src="AccessLevels.js"></script>
    <script type="text/javascript" language="javascript" src="GeneralJavaFunctions.js"></script>
    <script type="text/javascript" language="javascript" id="clientEventHandlersJS">
<!-- //Hide Client-Side SCRIPT


    var strWinMessage = "<%=strWinMessage%>" ;
    //var intServiceTypeID  = <%=strServiceTypeID%> ;
    var intServiceTypeID ;
    var intAccessLevel = <%=intAccessLevel%> ;
    var bolSaveRequired = false ;
    var arrLOBList = new Array() ;
    var arrServiceCategoryList = new Array() ;


    setPageTitle("SMA - Service Type");


    <% if isnumeric(strServiceTypeID) then %>
            intServiceTypeID = <%=strServiceTypeID%> ;
    <% end if %>

    <%If strEmailSubject <> "" Then%>
    //pop-up the email window
    var wndEmail = window.open('email.asp', 'PopupEmail', 'top=50, left=100, height=610, width=800');
    <%End If%>



    function iFrame_display(){
        //*********************************************************
        // Purpose:		Called whenever a refresh of the iFrame is needed
        //*************************************************************

        var strURL = 'STypeSLAList.asp?ServiceTypeID=' + intServiceTypeID ;
        document.getElementById("aifr").src = strURL ;
    }
    // ************** End of iFrame_display() **************
    function iSTAFrame_display() {

        var strAttrURL = 'STypeAttrList.asp?ServiceTypeID=' + intServiceTypeID ;
        document.getElementById("aiattrfr").src  = strAttrURL ;

    }

    function iSINSTFrame_display() {

        var strAttrURL = 'STypeInstList.asp?hdnServiceTypeID=' + intServiceTypeID ;
        document.getElementById("aiinstfr").src   = strAttrURL ;//aiinstfr

    }

    function iSKenanFrame_display() {

        var strAttrURL = 'STypeKenanList.asp?hdnServiceTypeID=' + intServiceTypeID ;
        //  var strAttrURL = 'STypeAttrList.asp?ServiceTypeID=' + intServiceTypeID ;
        document.getElementById("aiKenanfr").src  = strAttrURL ;

    }

    function btn_iFrmAdd(){

        var NewWin;
        var strSource ;

        if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
            alert('Access denied. Please contact your system administrator.');
            return;
        }

        strSource = 'STypeSLADetail.asp?XRefID=0&ServiceTypeID=' + document.frmSTypeDetail.hdnServiceTypeID.value ;
        NewWin=window.open(strSource ,"NewWin","toolbar=no,status=no,width=700,height=250,menubar=no resize=no");
        //NewWin=window.open(strSource ,"NewWin") ;
        NewWin.focus();

    }

    function btn_iSTAFrmAdd(){

        var NewWin;
        var strSource ;

        if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
            alert('Access denied. Please contact your system administrator.');
            return;
        }

        strSource = 'STypeAttDetail.asp?hdnXRefID=0&hdnServiceTypeID=' + document.frmSTypeDetail.hdnServiceTypeID.value ;
        NewWin=window.open(strSource ,"NewWin","toolbar=no,status=no,width=700,height=250,menubar=no resize=no");
        //NewWin=window.open(strSource ,"NewWin") ;
        NewWin.focus();

    }

    function btn_KenanFrmAdd(){

        var NewWin;
        var strSource ;

        if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
            alert('Access denied. Please contact your system administrator.');
            return;
        }

        strSource = 'STypeKenanDetail.asp?hdnXRefID=0&hdnServiceTypeID=' + document.frmSTypeDetail.hdnServiceTypeID.value ;
        strSource = strSource + '&hdnselPackID=0' ;
        NewWin=window.open(strSource ,"NewWin","toolbar=no,status=no,width=700,height=250,menubar=no resize=no");
        //NewWin=window.open(strSource ,"NewWin") ;
        NewWin.focus();

    }


    function btn_iSINSFrmAdd(){

        var NewWin;
        var strSource ;

        if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
            alert('Access denied. Please contact your system administrator.');
            return;
        }

        strSource = 'STypeInstDetail.asp?txtXRefID=0&hdnServiceTypeID=' + document.frmSTypeDetail.hdnServiceTypeID.value ;
        //strSource = strSource + '&txtInstID=0&txtInstvID=0&hdnUsageID=0' ;
        NewWin=window.open(strSource ,"NewWin","toolbar=no,status=no,width=700,height=250,menubar=no resize=no");
        //NewWin=window.open(strSource ,"NewWin") ;
        NewWin.focus();

    }


    function btn_iSTAFrmUpdate(){
        var NewWin ;

        var strSource = 'STypeAttDetail.asp?hdnXRefID=' + document.frames("aiattrfr").frmIFR.txtXRefID.value;
        //changed txtAttID to hdnstrAttID in below line in July 6 2009
        strSource = strSource + '&hdnstrAttID=' + document.frames("aiattrfr").frmIFR.txtattID.value;

        strSource = strSource + '&hdnstrattvID=' + document.frames("aiattrfr").frmIFR.txtattvID.value;
        strSource = strSource + '&hdnUsageID=' + document.frames("aiattrfr").frmIFR.hdnUsageID.value;
        strSource = strSource + '&hdnServiceTypeID=' + document.frames("aiattrfr").frmIFR.hdnServiceTypeID.value ;


        if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update) {
            alert('Access denied. Please contact your system administrator.');
            return ;
        }

        if (document.frames("aiattrfr").frmIFR.txtXRefID.value !=""){

            NewWin=window.open(strSource ,"NewWin","toolbar=no,status=no,width=700,height=250,menubar=no resize=no") ;
            //NewWin=window.open(strSource ,"NewWin") ;
            NewWin.focus();
        }

        else {
            alert('You must select a record to update!');
        }

    } // ************* End of btn_iSTAFrmUpdate() ************


    function btn_KenanFrmUpdate(){
        var NewWin ;

        var strSource = 'STypeKenanDetail.asp?hdnXRefID=' + document.frames("aiKenanfr").frmIFR.txtXRefID.value;
        strSource = strSource + '&hdnServiceTypeID=' + document.frames("aiKenanfr").frmIFR.hdnServiceTypeID.value ;
        strSource = strSource + '&hdnKenanCompID=' + document.frames("aiKenanfr").frmIFR.hdnKenanCompID.value;
        strSource = strSource + '&hdnKenanPackID='+ document.frames("aiKenanfr").frmIFR.hdnKenanPackID.value;

        if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update) {
            alert('Access denied. Please contact your system administrator.');
            return ;
        }

        if (document.frames("aiKenanfr").frmIFR.txtXRefID.value !=""){

            NewWin=window.open(strSource ,"NewWin","toolbar=no,status=no,width=700,height=250,menubar=no resize=no") ;
            //NewWin=window.open(strSource ,"NewWin") ;
            NewWin.focus();
        }

        else {
            alert('You must select a record to update!');
        }

    } // ************* End of btn_KenanFrmUpdate() ************



    function btn_iSINSFrmUpdate(){
        var NewWin ;

        var strSource = 'STypeInstDetail.asp?txtXRefID=' + document.frames("aiinstfr").frmIFR.txtXRefID.value;
        strSource = strSource + '&txtInstID=' + document.frames("aiinstfr").frmIFR.txtInstID.value;
        strSource = strSource + '&txtInstvID=' + document.frames("aiinstfr").frmIFR.txtInstvID.value;
        strSource = strSource + '&hdnUsageID=' + document.frames("aiinstfr").frmIFR.hdnUsageID.value;
        strSource = strSource + '&hdnServiceTypeID=' + document.frames("aiinstfr").frmIFR.hdnServiceTypeID.value ;


        if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update) {
            alert('Access denied. Please contact your system administrator.');
            return ;
        }

        if (document.frames("aiinstfr").frmIFR.txtXRefID.value !=0){

            NewWin=window.open(strSource ,"NewWin","toolbar=no,status=no,width=700,height=250,menubar=no resize=no") ;
            //NewWin=window.open(strSource ,"NewWin") ;
            NewWin.focus();
        }

        else {
            alert('You must select a record to update!');
        }

    } // ************* End of btn_iSTAFrmUpdate() ************

    function btn_iSINSSetSIASeq(){
        var NewWin;
        var strSource ;

        if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
            alert('Access denied. Please contact your system administrator.');
            return;
        }

        strSource = 'STypeInstListSeq.asp?hdnServiceTypeID=' + document.frmSTypeDetail.hdnServiceTypeID.value ;
        //strSource = strSource + '&txtInstID=0&txtInstvID=0&hdnUsageID=0' ;
        //20140613 NewWin=window.open(strSource ,"NewWin","toolbar=no,status=no,width=700,height=550,menubar=no resize=no");
        NewWin=window.open(strSource ,"NewWin","scrollbars=1 toolbar=0,status=0,width=700,height=600,menubar=0 resizable=1");

        //NewWin=window.open(strSource ,"NewWin") ;
        NewWin.focus();

    }

    function btn_iFrmUpdate(){
        var NewWin ;
        var strSource = 'STypeSLADetail.asp?XRefID=' + document.frames("aifr").frmIFR.txtXRefID.value + '&ServiceTypeID=' + document.frames("aifr").frmIFR.hdnServiceTypeID.value ;


        if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update) {
            alert('Access denied. Please contact your system administrator.');
            return ;
        }

        if (document.frames("aifr").frmIFR.txtXRefID.value !=""){

            NewWin=window.open(strSource ,"NewWin","toolbar=no,status=no,width=700,height=250,menubar=no resize=no") ;
            //NewWin=window.open(strSource ,"NewWin") ;
            NewWin.focus();
        }

        else {
            alert('You must select a record to update!');
        }

    } // ************* End of btn_iFrmUpdate() ************




    function btn_iFrmDelete() {
        var strURL ;
        // temp commented for testing
        if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) {
            alert('Access denied. Please contact your system administrator.') ;
            return ;
        }

        if (document.frames("aifr").frmIFR.txtXRefID.value !="") {
            if (confirm('Do you really want to delete this record?')){
                strURL = 'STypeSLAList.asp?txtFrmAction=DELETE&ServiceTypeID=' + document.frmSTypeDetail.hdnServiceTypeID.value + '&XRefID=' + document.frames("aifr").frmIFR.txtXRefID.value + '&UpdateDateTime=' + document.frames("aifr").frmIFR.hdnUpdateDateTime.value ;
                document.frames("aifr").document.location.href = strURL ;
            }
        }
        else {
            alert('You must select a record to delete!') ;
        }

    }  // ***************  end of btn_iFrmDelete() ******************


    function iSTAFrame_Delete(){
        var strURL ;
        if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) {
            alert('Access denied. Please contact your system administrator.') ;
            return ;
        }

        if (document.frames("aiattrfr").frmIFR.txtXRefID.value !="") {
            if (confirm('Do you really want to delete this record?')){
                strURL = 'STypeAttrList.asp?txtFrmAction=DELETE&ServiceTypeID=' + document.frmSTypeDetail.hdnServiceTypeID.value + '&XRefID=' + document.frames("aiattrfr").frmIFR.txtXRefID.value;
                document.frames("aiattrfr").document.location.href = strURL ;
            }
        }
        else {
            alert('You must select a record to delete!') ;
        }

    }  // ***************  end of btn_iSTAFrameDelete() ******************

    function iSKenanFrame_Delete(){
        var strURL ;
        if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) {
            alert('Access denied. Please contact your system administrator.') ;
            return ;
        }

        if (document.frames("aiKenanfr").frmIFR.txtXRefID.value !="") {
            if (confirm('Do you really want to delete this record?')){
                strURL = 'STypeKenanList.asp?txtFrmAction=DELETE&hdnServiceTypeID=' + document.frmSTypeDetail.hdnServiceTypeID.value + '&hdnXRefID=' + document.frames("aiKenanfr").frmIFR.txtXRefID.value;
                //			strURL = strURL + '&hdnselPackID=0';
                document.frames("aiKenanfr").document.location.href = strURL ;
            }
        }
        else {
            alert('You must select a record to delete!') ;
        }

    }  // ***************  end of iSKenanFrame_Delete() ******************



    function iSINSFrame_Delete(){
        var strURL ;
        /*  This section temp commented for my test  -- LC  */
        if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) {
            alert('Access denied. Please contact your system administrator.') ;
            return ;
        }

        if (document.frames("aiinstfr").frmIFR.txtXRefID.value !="") {
            if (confirm('Do you really want to delete this record?')){
                strURL = 'STypeInstList.asp?txtFrmAction=DELETE&hdnServiceTypeID=' + document.frmSTypeDetail.hdnServiceTypeID.value + '&txtXRefID=' + document.frames("aiinstfr").frmIFR.txtXRefID.value;
                strURL = strURL + '&hdnUsageID=' + document.frames("aiinstfr").frmIFR.hdnUsageID.value;

                document.frames("aiinstfr").document.location.href = strURL ;
            }
        }
        else {
            alert('You must select a record to delete!') ;
        }

    }  // ***************  end of btn_iSTAFrameDelete() ******************

    function fct_onMoveUp(){
        if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update) {alert('Access denied. Please contact your system administrator.'); return;}
        //	var strURL, strParams;
        //	if (document.frames("aiattrfr").frmIFR.txtXRefID.value !="") {
        //		strParams = 'txtFrmAction=move&direction=down&xrefid=' + document.frames("aiattrfr").frmIFR.txtXRefID.value;
        //		strURL = 'STypeAttrList.asp?txtFrmAction=move&direction=up&xrefid=' + document.frames("aiattrfr").frmIFR.txtXRefID.value;
        //		document.frames("aiattrfr").document.location.href = 'STypeAttrList.asp?hdnServiceTypeID=' + document.frmSTypeDetail.hdnServiceTypeID.value + '&' + strParams;
        //		}
        //	  }
        //    else {
        //		alert('You must select an element first.') ;
        //    }
        var strObjName = document.frames("aiattrfr").frmIFR.txtXRefID.value;
        if (strObjName != "") {
            strParams = 'txtFrmAction=move&direction=up&xrefid=' + document.frames("aiattrfr").frmIFR.txtXRefID.value;
            //document.frames("aiattrfr").document.location.href = 'STypeAttrList.asp?hdnServiceTypeID=' + document.frmSTypeDetail.hdnServiceTypeID.value + '&' + strParams;
            document.frames("aiattrfr").document.location.href = 'STypeAttrList.asp?ServiceTypeID=' + intServiceTypeID + '&' + strParams;
        } else alert('You must select an element first.');
    }  // ***************  end of fct_onMoveUp() ******************

    function fct_onMoveDown(){
        if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update) {alert('Access denied. Please contact your system administrator.'); return;}
        var strObjName = document.frames("aiattrfr").frmIFR.txtXRefID.value;
        if (strObjName != "") {
            strParams = 'txtFrmAction=move&direction=down&xrefid=' + document.frames("aiattrfr").frmIFR.txtXRefID.value;
            document.frames("aiattrfr").document.location.href = 'STypeAttrList.asp?ServiceTypeID=' + intServiceTypeID + '&' + strParams;
        } else alert('You must select an element first.');
    }  // ***************  end of fct_onMoveDown() ******************

    //function body_onLoad(){
    //	iFrame_display();
    //}

    function fct_setDays(iIndex) {
        var intDays = 31;
        var strMonth = document.frmSTypeDetail.item("selmonth", iIndex).options[document.frmSTypeDetail.item("selmonth", iIndex).selectedIndex].value;
        var strYear = document.frmSTypeDetail.item("selyear", iIndex).options[document.frmSTypeDetail.item("selyear", iIndex).selectedIndex].value;
        var intCurrentDay = document.frmSTypeDetail.item("selday", iIndex).options[document.frmSTypeDetail.item("selday", iIndex).selectedIndex].value;
        var intCounter = document.frmSTypeDetail.item("selday", iIndex).options.length;

        switch (strMonth) {
            case "02":						//February
                if (strYear % 4 != 0) { intDays = 28; }
                else if (strYear % 400 == 0) { intDays = 29; }
                else if (strYear % 100 == 0) { intDays = 28; }
                else { intDays = 29; }
                break;
            case "04": intDays = 30; break;	//April
            case "06": intDays = 30; break;	//June
            case "09": intDays = 30; break;	//September
            case "11": intDays = 30; break;	//November
            default: intDays = 31; break;	//January, March, May, July, August, October, December
        }
        if (intCounter <= intDays) {
            while (intCounter <= intDays) {
                var oOption = new Option(intCounter, intCounter);
                document.frmSTypeDetail.item("selday", iIndex).options[intCounter++] = oOption;
            }
        }
        else {
            while (intCounter > intDays) {
                document.frmSTypeDetail.item("selday", iIndex).options[intCounter--] = null;
            }
        }
        if (intCurrentDay > intDays) {
            document.frmSTypeDetail.item("selday", iIndex).selectedIndex = intDays;
        }
        bolSaveRequired = true;
    }

    function btnCalendar_onClick(iIndex) {
        var NewWin;
        SetCookie("Field", iIndex);
        NewWin=window.open("TheCalendar.asp","NewWin","toolbar=no,status=no,width=260,height=225,menubar=no resize=no");
        //NewWin.creator=self;
        NewWin.focus();
        bolSaveRequired = true;
    }

    function fct_selNavigate(){
        //***********************************************************************************************
        // Function:	selNavigate_onChange															*
        //																								*
        // Purpose:		To display the page selected by the user from Quick Navigation drop-down box.	*
        //																								*
        // Created By:	Gilles Archer 09/27/2000														*
        //																								*
        // Updated By:																					*
        //***********************************************************************************************
        var strPageName = document.frmSTypeDetail.selNavigate.item(document.frmSTypeDetail.selNavigate.selectedIndex).value ;

        switch (strPageName) {
            case "CustomerServices":
                document.frmSTypeDetail.selNavigate.selectedIndex = 0;
                var strServiceTypeID = document.frmSTypeDetail.hdnServiceTypeID.value;
                var strServiceTypeName = document.frmSTypeDetail.txtServiceDescription.value;
                SetCookie("ServiceTypeID", strServiceTypeID);
                SetCookie("ServiceTypeName", strServiceTypeName);
                self.location.href = "SearchFrame.asp?fraSrc=CustServ";
                break ;

            case "LOB":
                document.frmSTypeDetail.selNavigate.selectedIndex = 0;
                var strBusinessID = document.frmSTypeDetail.hdnBusinessID.value;
                self.location.href = "LOBDetail.asp?BusinessID=" + strBusinessID;
                break ;

            case "SCategory":
                document.frmSTypeDetail.selNavigate.selectedIndex = 0;
                var strServiceCategoryID = document.frmSTypeDetail.hdnServiceCategoryID.value;
                self.location.href = "SCategoryDetail.asp?ServiceCategoryID=" + strServiceCategoryID;
                break ;

                //case "SLA":
                //	document.frmSTypeDetail.selNavigate.selectedIndex = 0;
                //	var strServiceLevelID = document.frmSTypeDetail.hdnServiceLevelID.value;
                //	self.location.href = "SLADetail.asp?ServiceLevelID=" + strServiceLevelID;
                //	break ;

            case "DEFAULT":
                // do nothing ;
        }
    }

    function fct_onChangeLOB() {
        var intCounter = 1;
        var strBusinessID;

        if (document.frmSTypeDetail.selLOB.selectedIndex != 0) {
            strBusinessID = document.frmSTypeDetail.selLOB.value;


            //Remove all the OPTION tags from the Service Category
            for (intCounter = document.frmSTypeDetail.selServiceCategory.length - 1; intCounter > 0; intCounter--) {
                document.frmSTypeDetail.selServiceCategory.options.remove(intCounter);
            }

            //Add Service Categories that belong to the selected Line of Business
            for (intCounter = 1; intCounter < arrServiceCategoryList.length; intCounter++) {
                var strValue = arrServiceCategoryList[intCounter];
                var arrValue = strValue.split("|");
                if (arrValue[1] == strBusinessID) {
                    //var strElement = "<option value='" + arrValue[0] + "'>" + arrValue[2] + "</option>";
                    //var oOption = document.createElement(strElement);
		    var oOption = document.createElement("option");
   		    oOption.text = arrValue[2];
    		    oOption.value = arrValue[0];

                    document.frmSTypeDetail.selServiceCategory.options.add(oOption);
                    oOption.innerText = arrValue[2];	//SERVICE_CATEGORY_DESC
                    //				oOption.Value = arrValue[1];		//LOB_ID
                    //				oOption.Value = arrValue[0];		//SERVICE_CATEGORY_ID
                }
            }
        }
        else {
            //Remove all the OPTION tags from the Service Category
            for (intCounter = document.frmSTypeDetail.selServiceCategory.length - 1; intCounter > 0; intCounter--) {
                document.frmSTypeDetail.selServiceCategory.options.remove(intCounter);
            }
            //Add all the Service Categories
            for (intCounter = 1; intCounter < arrServiceCategoryList.options.length; intCounter++) {
                var strValue = arrServiceCategoryList[intCounter];
                var arrValue = strValue.split("|");
                //var strElement = "<option value='" + arrValue[0] + "'>" + arrValue[2] + "</option>";
                //var oOption = document.createElement(strElement);
		var oOption = document.createElement("option");
   		oOption.text = arrValue[2];
    		oOption.value = arrValue[0];
                document.frmSTypeDetail.selServiceCategory.options.add(oOption);
                oOption.innerText = arrValue[2];	//SERVICE_CATEGORY_DESC
                //			oOption.Value = arrValue[1];		//LOB_ID
                //			oOption.Value = arrValue[0];		//SERVICE_CATEGORY_ID
            }
        }
    }

    function btnDelete_onClick() {
        //**********************************************************************************************
        // Function:	btnDelete_onClick
        //
        // Purpose:		To delete a service type
        //
        // Created By:	Gilles Archer 09/27/2000
        //
        // Updated By:
        //***********************************************************************************************
        if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) {
            alert('You do not have permission to DELETE a Service Type.  Please contact your System Administrator.');
            return false;
        }

        if (document.frmSTypeDetail.hdnServiceTypeID.value == "") {
            alert('This Service Type does not exist in the database.');
            return false;
        }

        if (confirm('Do you really want to delete this object?')){
            document.frmSTypeDetail.hdnFrmAction.value = "DELETE";
            document.frmSTypeDetail.submit();
        }
    }

    function btnNew_onClick() {
        if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
            alert('You do not have permission to CREATE a Service Type.  Please contact your System Administrator.');
            return false;
        }
        document.location = "STypeDetail.asp?ServiceTypeID=NEW";
    }

    function fct_onChange() {
        bolSaveRequired = true;
    }

    function btnSave_onClick() {
        if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update) {
            alert('You do not have permission to UPDATE a Service Type.  Please contact your System Administrator.');
            return false;
        }

        if (document.frmSTypeDetail.txtServiceDescription.value == "") {
            alert('Please enter the Service Type Description');
            document.frmSTypeDetail.txtServiceDescription.focus();
            return false;
        }

        if (document.frmSTypeDetail.selServiceCategory.selectedIndex == 0) {
            alert('Please select a Service Category');
            document.frmSTypeDetail.selServiceCategory.focus();
            return false;
        }

        if (document.frmSTypeDetail.selServiceClass.selectedIndex == 0) {
            alert('Please select a Service Class');
            document.frmSTypeDetail.selServiceClass.focus();
            return false;
        }


        if (document.frmSTypeDetail.item("selmonth", 0).value == "" || document.frmSTypeDetail.item("selday", 0).value == "" || document.frmSTypeDetail.item("selyear", 0).value == "") {
            alert('Please enter a Service Start Date');
            document.frmSTypeDetail.item("btnCalendar", 0).focus();
            return false;
        }


        document.frmSTypeDetail.hdnFrmAction.value = "SAVE";
        bolSaveRequired = false;
        document.frmSTypeDetail.submit();
        return true;
    }

    function btnReferences_onClick() {
        var strOwner = 'CRP';			// owner name must be in Uppercase
        var strTableName = 'SERVICE_TYPE';		// replace ADDRESS with your own table name and table name must be in Uppercase
        var strRecordID = document.frmSTypeDetail.hdnServiceTypeID.value ;   // insert your record id
        var strURL;

        if (strRecordID == "") {
            alert("No references. This is a new record.");
            return false;
        }

        strURL = "Dependency.asp?Owner=" + strOwner + "&TableName=" + strTableName + "&RecordID=" + strRecordID;
        window.open(strURL, 'Popup', 'top=100, left=100, width=500, height=300');
    }

    function window_onBeforeUnload() {
        //Ensure that fct_onChange() fires for any changed data.
        document.frmSTypeDetail.btnSave.focus();

        if (bolSaveRequired) {
            event.returnValue = "There is unsaved data on the screen. To save changes, click CANCEL below then click SAVE on the main FORM.";
        }
    }

    function window_onUnload() {
        //
    }

    function ClearStatus() {
        window.status = "";
    }

    function DisplayStatus(strWinStatus) {

        var strBusinessID = document.frmSTypeDetail.hdnBusinessID.value;
        var strServiceCategoryID = document.frmSTypeDetail.hdnServiceCategoryID.value;
        var intCounter;

        arrLOBList[0] = "";
        arrServiceCategoryList[0] = "";
        //debugger;
        var lobOptions = document.getElementById("selLOB").getElementsByTagName("option");
        var searchOptions = document.getElementById("selServiceCategory").getElementsByTagName("option");

        for (intCounter = 1; intCounter < lobOptions.length; intCounter++) {
            var oOption = lobOptions[intCounter];
            //Each array element holds LOB_ID|LOB_DESC
            arrLOBList[intCounter] = (oOption.value + "|" + oOption.text);
        }

        for (intCounter = 1; intCounter < searchOptions.length; intCounter++) {
            var oOption = searchOptions[intCounter];
            //Each array element holds SERVICE_CATEGORY_ID|LOB_ID|SERVICE_CATEGORY_DESC
            arrServiceCategoryList[intCounter] = (oOption.value + "|" + oOption.text);
        }
        window.status = strWinStatus;
        setTimeout('ClearStatus()', 5000);

        iFrame_display();
        iSTAFrame_display();
        iSINSTFrame_display();
        iSKenanFrame_display();
    }

    function btnReset_onClick() {
        if(confirm('All changes will be lost. Do you really want to reset the page?')){
            bolSaveRequired = false;
            document.location.href = "STypeDetail.asp?ServiceTypeID=<%=strServiceTypeID%>";
        }
    }


    //function fct_onChange() {
    // some comments

    //}

    // Unhide Client-Side SCRIPT -->
    </script>
</head>

<body language="javascript" onload="DisplayStatus('');" onload="iFrame_display();" onbeforeunload="window_onBeforeUnload();" onunload="window_onUnload();">
    <form id="frmSTypeDetail" name="frmSTypeDetail" action="STypeDetail.asp" method="post">

        <input type="hidden" id="hdnBusinessID" name="hdnBusinessID" value="<%If IsNumeric(strServiceTypeID) Then Response.Write objRS.Fields("LOB_ID").Value%>">
        <input type="hidden" id="hdnServiceCategoryID" name="hdnServiceCategoryID" value="<%If IsNumeric(strServiceTypeID) Then Response.Write objRS.Fields("SERVICE_CATEGORY_ID").Value%>">
        <input type="hidden" id="hdnServiceTypeID" name="hdnServiceTypeID" value="<%If IsNumeric(strServiceTypeID) Then Response.Write objRS.Fields("SERVICE_TYPE_ID").Value%>">
        <input type="hidden" id="hdnServiceLevelID" name="hdnServiceLevelID" value="<%If IsNumeric(strServiceTypeID) Then Response.Write objRS.Fields("DEFAULT_SLA_ID").Value%>">
        <input type="hidden" id="hdnUpdateDateTime" name="hdnUpdateDateTime" value="<%If IsNumeric(strServiceTypeID) Then Response.Write objRS.Fields("LAST_UPDATE_DATE_TIME").Value%>">
        <input type="hidden" id="hdnFrmAction" name="hdnFrmAction" value="">


        <table cols="4" width="100%">
            <thead>
                <tr>
                    <td align="left" colspan="3">Service Type Detail</td>
                    <td align="right">
                        <select valign="top" id="selNavigate" name="selNavigate" onchange="fct_selNavigate();">
                            <option value="DEFAULT" selected>Quickly Goto ...</option>
                            <option value="CustomerServices">Customer Services</option>
                            <option value="LOB">Line of Business</option>
                            <option value="SCategory">Service Category</option>
                        </select></td>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td align="right" valign="top" nowrap>Service Type ID <font color="red">*</font></td>
                    <td align="left" colspan="3" nowrap>
                        <input id="txtServiceTypeID" name="txtServiceTypeID" disabled value="<% =strServiceTypeID%>" maxlength="80" size="13"></td>
                </tr>

                <tr>
                    <td align="right" valign="top" nowrap>English Description<font color="red">*</font></td>
                    <td align="left" colspan="3" nowrap>
                        <input id="txtServiceDescription" name="txtServiceDescription" onchange="return fct_onChange();" value="<%If IsNumeric(strServiceTypeID) Then Response.Write objRS.Fields("SERVICE_TYPE_DESC").Value%>" maxlength="80" size="80"></td>
                </tr>
                <tr>
                    <td align="right" nowrap>Description Française&nbsp;<br />
                        French Description<font color="red">&nbsp</font></td>
                    <td align="left" colspan="2" nowrap>
                        <input id="txtServiceTypeFrench" name="txtServiceTypeFrench" onchange="fct_onChange();" value="<%If IsNumeric(strServiceTypeID) Then Response.Write objRSFrench.Fields("SERVICE_TYPE_LANG_DESC").Value%>" maxlength="80" size="80"></td>
                </tr>
                <tr>
                    <td align="right" valign="top" nowrap>Line of Business<font color="red">*</font></td>
                    <td align="left" nowrap>
                        <select id="selLOB" name="selLOB" style="width: 350px" onchange="fct_onChangeLOB();">
                            <option></option>
                            <%For lRow = LBound(arrLOBList, 2) To UBound(arrLOBList, 2)
			If IsNumeric(strServiceTypeID) Then
				If StrComp(CStr(objRS.Fields("LOB_ID").Value), arrLOBList(0, lRow), 0) = 0 Then%>
                            <option selected value="<%=arrLOBList(0, lRow)%>"><%=arrLOBList(1, lRow) & " - " & arrLOBList(2, lRow)%></option>
                            <%Else%>
                            <option value="<%=arrLOBList(0, lRow)%>"><%=arrLOBList(1, lRow) & " - " & arrLOBList(2, lRow)%></option>
                            <%	End If
			Else%>
                            <option value="<%=arrLOBList(0, lRow)%>"><%=arrLOBList(1, lRow) & " - " & arrLOBList(2, lRow)%></option>
                            <%End If
		Next%>
                        </select></td>

                </tr>
                <tr>
                    <td align="right" valign="top" nowrap>Service Category<font color="red">*</font></td>
                    <td align="left" nowrap>
                        <select id="selServiceCategory" name="selServiceCategory" style="width: 350px" onchange="return fct_onChange();">
                            <option></option>
                            <%For lRow = LBound(arrCategoryList, 2) To UBound(arrCategoryList, 2)
			If IsNumeric(strServiceTypeID) Then
				If StrComp(CStr(objRS.Fields("SERVICE_CATEGORY_ID").Value), arrCategoryList(0, lRow), 0) = 0 Then%>
                            <option selected value="<%=arrCategoryList(0, lRow) & "|" & arrCategoryList(1, lRow)%>"><%=arrCategoryList(2, lRow)%></option>
                            <%Else%>
                            <option value="<%=arrCategoryList(0, lRow) & "|" & arrCategoryList(1, lRow)%>"><%=arrCategoryList(2, lRow)%></option>
                            <%End If
			Else%>
                            <option value="<%=arrCategoryList(0, lRow) & "|" & arrCategoryList(1, lRow)%>"><%=arrCategoryList(2, lRow)%></option>
                            <%End If
		Next%>
                        </select></td>

                </tr>
                <tr>
                    <td align="right" valign="top" nowrap>Service Class<font color="red">*</font></td>
                    <td align="left" nowrap>
                        <select id="selServiceClass" name="selServiceClass" style="width: 250px" onchange="return fct_onChange();">
                            <option></option>
                            <%For lRow = LBound(arrClassList, 2) To UBound(arrClassList, 2)
			If IsNumeric(strServiceTypeID) Then
				If StrComp(objRS.Fields("SERVICE_CLASS_LCODE").Value, arrClassList(0, lRow), 0) = 0 Then%>
                            <option selected value="<%=arrClassList(0, lRow)%>"><%=arrClassList(1, lRow)%></option>
                            <%Else%>
                            <option value="<%=arrClassList(0, lRow)%>"><%=arrClassList(1, lRow)%></option>
                            <%End If
			Else%>
                            <option value="<%=arrClassList(0, lRow)%>"><%=arrClassList(1, lRow)%></option>
                            <%	End If
		Next%>
                        </select></td>
                </tr>

                <tr>



                    <td align="right" valign="top" nowrap>VPN Type<font color="red">*</font></td>
                    <td align="left" nowrap>
                        <select id="selVPNTypes" name="selVPNTypes" style="width: 250px" onchange="return fct_onChange();">

                            <%For lRow = LBound(arrVPNList, 2) To UBound(arrVPNList, 2)
	 		If IsNumeric(strServiceTypeID) Then
				If StrComp(objRS.Fields("VPN_TYPE_LCODE").Value, arrVPNList(0, lRow), 0) = 0 Then%>
                            <option selected value="<%=arrVPNList(0, lRow)%>"><%=arrVPNList(1, lRow)%></option>
                            <%Else%>
                            <option value="<%=arrVPNList(0, lRow)%>"><%=arrVPNList(1, lRow)%></option>
                            <%End If
			Else%>
                            <option value="<%=arrVPNList(0, lRow)%>"><%=arrVPNList(1, lRow)%></option>
                            <%	End If
		Next%>
                        </select></td>





                </tr>


                <tr>
                    <td align="right" valign="top" nowrap>Start Date<font color="red">*</font></td>
                    <td align="left" nowrap>
                        <select id="selmonth" name="selmonth" onchange="fct_setDays(0);return fct_onChange();">
                            <option></option>
                            <%For lIndex = 1 to 12
			Response.Write "<option "
			If IsNumeric(strServiceTypeID) Then
				If lIndex = Month(objRS.Fields("SERVICE_TYPE_START_DATE").Value) Then Response.Write "selected "
			End If
			If lIndex < 10 Then	lIndex = "0" & lIndex
			Response.Write "value='" & lIndex & "'>" & monthName(lIndex, False) & "</option>"
		Next%>
                        </select>
                        <select id="selday" name="selday" onchange="fct_setDays(0);return fct_onChange();">
                            <option></option>
                            <%For lIndex = 1 to 31
			Response.Write "<option "
			If IsNumeric(strServiceTypeID) Then
				If lIndex = Day(objRS.Fields("SERVICE_TYPE_START_DATE").Value) Then Response.Write "selected "
			End If
			If lIndex < 10 Then	lIndex = "0" & lIndex
			Response.Write "value='" & lIndex & "'>" & lIndex & "</option>"
		Next%>
                        </select>
                        <select id="selyear" name="selyear" onchange="fct_setDays(0);return fct_onChange();">
                            <option></option>
                            <%For lIndex = intBaseYear To Year(Now) +  7
			Response.Write "<OPTION "
			If IsNumeric(strServiceTypeID) Then
				If lIndex = Year(objRS.Fields("SERVICE_TYPE_START_DATE").Value) Then Response.Write "selected "
			End If
			Response.Write "value='" & lIndex & "'>" & lIndex & "</OPTION>"
		Next%>
                        </select>
                        <input id="btnCalendar" name="btnCalendar" type="button" value="..." language="javascript" onclick="return btnCalendar_onClick(0);return fct_onChange();"></td>
                </tr>
                <tr>
                    <td align="right" valign="top" nowrap>End Date</td>
                    <td align="left" nowrap>
                        <select id="selmonth" name="selmonth" onchange="fct_setDays(1);fct_onChange();">
                            <option></option>
                            <%For lIndex = 1 to 12
			Response.Write "<OPTION "
			If IsNumeric(strServiceTypeID) Then
				If lIndex = Month(objRS.Fields("SERVICE_TYPE_END_DATE").Value) Then Response.Write "selected "
			End If
			If lIndex < 10 Then	lIndex = "0" & lIndex
			Response.Write "value='" & lIndex & "'>" & monthName(lIndex, False) & "</OPTION>"
		Next%>
                        </select>
                        <select id="selday" name="selday" onchange="fct_setDays(1);return fct_onChange();">
                            <option></option>
                            <%For lIndex = 1 to 31
			Response.Write "<OPTION "
			If IsNumeric(strServiceTypeID) Then
				If lIndex = Day(objRS.Fields("SERVICE_TYPE_END_DATE").Value) Then Response.Write "selected "
			End If
			If lIndex < 10 Then	lIndex = "0" & lIndex
			Response.Write "value='" & lIndex & "'>" & lIndex & "</OPTION>"
		Next%>
                        </select>
                        <select id="selyear" name="selyear" onchange="fct_setDays(1);return fct_onChange();">
                            <option></option>
                            <%For lIndex = intBaseYear To Year(Now) + 7
			Response.Write "<OPTION "
			If IsNumeric(strServiceTypeID) Then
				If lIndex = Year(objRS.Fields("SERVICE_TYPE_END_DATE").Value) Then Response.Write "selected "
			End If
			Response.Write "value='" & lIndex & "'>" & lIndex & "</OPTION>"
		Next%>
                        </select>
                        <input id="btnCalendar" name="btnCalendar" type="button" value="..." language="javascript" onclick="return btnCalendar_onClick(1);return fct_onChange();"></td>
                </tr>
                <tr>
                    <td align="right" valign="top" nowrap> Send to NetCracker </td>

                    <td align="left" colspan="3" nowrap>
                        <select id="txtNCFlag" name="txtNCFlag">
                            <option <% if Clng(strNC) = 0 then response.write "Selected"  %> value="0">No</option>
                            <option <% if Clng(strNC) = 1 then response.write "Selected"  %> value="1">Manual</option>
                            <option <% if Clng(strNC) = 2 then response.write "Selected"  %> value="2">Automated</option>

                        </select>
                    </td>
                </tr>
            </tbody>
        </table>

        <table>
            <thead>
                <tr>
                    <td colspan="2" width="70%">Service Type Attributes</td>
                </tr>
                <tbody>
                    <td>
                        <iframe id="aiattrfr" width="100%" height="100" src="" scrolling="yes" marginheight="1" marginwidth="1"></iframe>
                        <input type="button" style="width: 2cm" value="Delete" <%if strServiceTypeID ="NEW" or strServiceTypeID ="" then Response.Write "DISABLED" end if%> name="btn_iSTAFrameDelete" onclick="iSTAFrame_Delete();">
                        &nbsp;&nbsp;
		<input type="button" style="width: 2cm" value="Refresh" <%if strServiceTypeID ="NEW" or strServiceTypeID ="" then Response.Write "DISABLED" end if%> name="btn_iSTAFrameRefresh" onclick="iSTAFrame_display();">
                        &nbsp;&nbsp;
		<input type="button" style="width: 2cm" value="New" <%if strServiceTypeID ="NEW" or strServiceTypeID ="" then Response.Write "DISABLED" end if%> name="btn_iSTAFrameAdd" onclick="btn_iSTAFrmAdd();fct_onChange();">
                        &nbsp;&nbsp;
		<input type="button" style="width: 2cm" value="Update" <%if strServiceTypeID ="NEW" or strServiceTypeID ="" then Response.Write "DISABLED" end if%> name="btn_iSTAFrameupdate" onclick="btn_iSTAFrmUpdate();fct_onChange();">
                        <img src="images/up.gif" title width="31" height="31" <%if strServiceTypeID ="NEW" or strServiceTypeID ="" then Response.Write "DISABLED" end if%> onclick="fct_onMoveUp();">
                        <img src="images/down.gif" title width="31" height="31" <%if strServiceTypeID ="NEW" or strServiceTypeID ="" then Response.Write "DISABLED" end if%> onclick="fct_onMoveDown()">
                    </td>
                    <td></td>
                    <td></td>
                </tbody>
        </table>

        <table>
            <thead>
                <tr>
                    <td colspan="2" width="70%">Service Instance Attributes</td>
                </tr>
                <tbody>
                    <td>
                        <iframe id="aiinstfr" width="100%" height="100" src="" scrolling="yes" marginheight="1" marginwidth="1"></iframe>
                        <input type="button" style="width: 2cm" value="Delete" <%if strServiceTypeID ="NEW" or strServiceTypeID ="" then Response.Write "DISABLED" end if%> name="btn_iSINSFrameDelete" onclick="iSINSFrame_Delete();">
                        &nbsp;&nbsp;
		<input type="button" style="width: 2cm" value="Refresh" <%if strServiceTypeID ="NEW" or strServiceTypeID ="" then Response.Write "DISABLED" end if%> name="btn_iSINSFrameRefresh" onclick="iSINSTFrame_display();">
                        &nbsp;&nbsp;
		<input type="button" style="width: 2cm" value="New" <%if strServiceTypeID ="NEW" or strServiceTypeID ="" then Response.Write "DISABLED" end if%> name="btn_iSINSFrameAdd" onclick="btn_iSINSFrmAdd();fct_onChange();">
                        &nbsp;&nbsp;
		<input type="button" style="width: 2cm" value="Update" <%if strServiceTypeID ="NEW" or strServiceTypeID ="" then Response.Write "DISABLED" end if%> name="btn_iSINSFrameupdate" onclick="btn_iSINSFrmUpdate();fct_onChange();">
                        &nbsp;&nbsp;
        <input type="button" style="width: 2.1cm" value="Set SIA Seq" <%if strServiceTypeID ="NEW" or strServiceTypeID ="" then Response.Write "DISABLED" end if%> name="btn_iSINSFrameSetSeq" onclick="btn_iSINSSetSIASeq();fct_onChange();">
                    </td>
                    <td></td>
                    <td></td>
                </tbody>
        </table>

        <table>
            <thead>
                <tr>
                    <td colspan="4" align="left">Kenan Package / Component</td>
                </tr>
                <tbody>
                    <td>
                        <iframe id="aiKenanfr" width="100%" height="100" src="" scrolling="yes" marginheight="1" marginwidth="1"></iframe>
                        <input type="button" style="width: 2cm" value="Delete" <%if strServiceTypeID ="NEW" or strServiceTypeID ="" then Response.Write "DISABLED" end if%> name="btn_KenanFrameDelete" onclick="iSKenanFrame_Delete();">
                        &nbsp;&nbsp;
		<input type="button" style="width: 2cm" value="Refresh" <%if strServiceTypeID ="NEW" or strServiceTypeID ="" then Response.Write "DISABLED" end if%> name="btn_KenanFrameRefresh" onclick="iSKenanFrame_display();">
                        &nbsp;&nbsp;
		<input type="button" style="width: 2cm" value="New" <%if strServiceTypeID ="NEW" or strServiceTypeID ="" then Response.Write "DISABLED" end if%> name="btn_KenanFrameAdd" onclick="btn_KenanFrmAdd();fct_onChange();">
                        &nbsp;&nbsp;
		<input type="button" style="width: 2cm" value="Update" <%if strServiceTypeID ="NEW" or strServiceTypeID ="" then Response.Write "DISABLED" end if%> name="btn_KenanFrameupdate" onclick="btn_KenanFrmUpdate();fct_onChange();">
                    </td>
                    <td></td>
                    <td></td>
                </tbody>
        </table>


        <table>
            <thead>
                <tr>
                    <td colspan="4" align="left">Default SLAs</td>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td width="25%" rowspan="5" colspan="2" valign="top" align="left">
                        <iframe id="aifr" width="100%" height="100" src="" scrolling="yes" marginheight="1" marginwidth="1"></iframe>
                        <br>
                        <input type="button" style="width: 2cm" value="Delete" <%if strServiceTypeID ="NEW" or strServiceTypeID ="" then Response.Write "DISABLED" end if%> name="btn_iFrameDelete" onclick="btn_iFrmDelete();fct_onChange();">&nbsp;&nbsp;
			<input type="button" style="width: 2cm" value="Refresh" <%if strServiceTypeID ="NEW" or strServiceTypeID ="" then Response.Write "DISABLED" end if%> name="btn_iFrameRefresh" onclick="iFrame_display();">
                        &nbsp;&nbsp;
			<input type="button" style="width: 2cm" value="New" <%if strServiceTypeID ="NEW" or strServiceTypeID ="" then Response.Write "DISABLED" end if%> name="btn_iFrameAdd" onclick="btn_iFrmAdd();fct_onChange();">
                        &nbsp;&nbsp;
			<input type="button" style="width: 2cm" value="Update" <%if strServiceTypeID ="NEW" or strServiceTypeID ="" then Response.Write "DISABLED" end if%> name="btn_iFrameupdate" onclick="btn_iFrmUpdate();fct_onChange();">
                    </td>
                    <td width="15%"></td>
                </tr>

            </tbody>
        </table>

        <table>
            <tfoot>
                <td colspan="4" align="right">
                    <input id="btnReferences" name="btnReferences" type="button" value="References" style="width: 2.2cm" language="javascript" onclick="return btnReferences_onClick();">&nbsp;
	<input id="btnDelete" name="btnDelete" type="button" value="Delete" style="width: 2cm" language="javascript" onclick="return btnDelete_onClick();">&nbsp;
	<input id="btnReset" name="btnReset" type="button" value="Reset" style="width: 2cm" language="javascript" onclick="return btnReset_onClick();">&nbsp;
	<input id="btnNew" name="btnNew" type="button" value="New" style="width: 2cm" language="javascript" onclick="return btnNew_onClick();">&nbsp;
	<input id="btnSave" name="btnSave" type="button" value="Save" style="width: 2cm" onclick="return btnSave_onClick();">&nbsp;</td>
                </TR>
            </tfoot>
        </table>

        <fieldset width="100%">
            <legend align="right"><b>Audit Information</b></legend>
            <div size="8pt" align="right">
                Record Status Indicator:<input align="left" name="txtRecordStatusInd" type="text" style="width: 18px" disabled value="<%If IsNumeric(strServiceTypeID) Then Response.Write objRS.Fields("RECORD_STATUS_IND").Value%>">&nbsp;&nbsp;&nbsp;
	Create Date:<input align="center" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If IsNumeric(strServiceTypeID) Then Response.Write objRS.Fields("CREATE_DATE_TIME").Value%>">&nbsp;
	Created By:<input align="right" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If IsNumeric(strServiceTypeID) Then Response.Write objRS.Fields("CREATE_REAL_USERID").Value%>"><br>
                Update Date:<input align="center" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If IsNumeric(strServiceTypeID) Then Response.Write objRS.Fields("UPDATE_DATE_TIME").Value%>">&nbsp;
	Updated By:<input align="right" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If IsNumeric(strServiceTypeID) Then Response.Write objRS.Fields("UPDATE_REAL_USERID").Value%>">
            </div>
        </fieldset>
    </form>
    <%
	'Clean up our ADO objects
	Set objRS = Nothing
	objConn.Close
	Set ObjConn = Nothing
    %>
</body>
</html>
