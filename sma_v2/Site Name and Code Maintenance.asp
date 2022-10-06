<%@  language="VBScript" %>
<% Option Explicit %>
<% Response.Buffer = true %>
<% On Error Resume Next %>
<!--#include file="smaConstants.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--#include file="smaProcs.inc"-->

<!-- This is the child detail screen for the service location screen.
     It can have the following values passed into it:

     	Parameter			Details
     	------------------------------------------------------------------------
    	ServLocContactID	the ID from the database of the service location contact
    						required for updates and deletes

    	ServLocID			the ID of the service location from the parent screen
    						required for creates

    	CustName			the Customer's name form the parent screen. This is only used to select an appropriate customer
    						required for creates

    	NewContact			must have a value of 'NEW'
	    					required for new records

    	hdnUpdateDateTime	the updateDateTime from the database
							required for updates and deletes

********************************************************************************************

        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       22-Jan-01	 DTy		Increase contact priority from 10 to 30.
       19-Feb-02	 DTy		Provide extra space for email-address which had increased
                                  from 50 to 60 characters.
********************************************************************************************
-->


<%
Const ASP_NAME = "Site Name and Code Maintenance.asp"
Const NO_ID    = "null"
 
dim intAccessLevel

intAccessLevel = CInt(CheckLogon(strConst_ServiceLocationContact))
if intAccessLevel < intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to Service Location Contacts. Please contact your system administrator."
end if

    dim sitecode 
  dim siteName 
Dim strRealUserID
strRealUserID = Session("username")

'Response.Write "USER=" & strRealUserID

Dim strSite_Name, strSiteID, strSite_code, serLocId
Dim strSql
'Dim objRS, objRSContactRole, objCmd
'Dim strContactInfo
    Dim objRS
    strSiteID = Request("SiteID")
    serLocId= Request("ServLocID")
    
if strSiteID = "" then
	strSiteID = NO_ID
end if


    dim action
    set action = "NEW"

if (Request("hdnFrmAction") = "NEW" or Request("action") = "NEW")  then
	set strSiteID = NO_ID
    end if 
    if (strSiteID <> NO_ID)  then
    Dim cmdUpdateObj
    set cmdUpdateObj = server.CreateObject("ADODB.Command")
			set cmdUpdateObj.ActiveConnection = objConn
			cmdUpdateObj.CommandType = adCmdText
			cmdUpdateObj.CommandText = "select SITE_NAME,SITE_CODE from CRP.SITE_NAME_CODE where Site_ID ="& strSiteID
          set objRS =  cmdUpdateObj.Execute
    set sitecode =objRs("SITE_CODE")
    set siteName =objRs("SITE_NAME")

end if
    
'dim aRole		'used to get the ID from the drop down list
select case Request("hdnFrmAction")
	case "SAVE"
    
		if (strSiteID <> NO_ID) then

			
			set cmdUpdateObj = server.CreateObject("ADODB.Command")
			set cmdUpdateObj.ActiveConnection = objConn
			cmdUpdateObj.CommandType = adCmdText
			cmdUpdateObj.CommandText = "update CRP.SITE_NAME_CODE set SITE_NAME = '" & Request("txtSiteName") & "', SITE_CODE = '" & Request("txtSiteCode") & "' where Site_ID="& strSiteID 
            cmdUpdateObj.Execute

			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT UPDATE OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			
			end if
			strWinMessage = "Record saved successfully. You can now see the changes you made."

		else
			'create a new record
			if (intAccessLevel and intConst_Access_Create) <> intConst_Access_Create then
				DisplayError "BACK", "", 0, "CREATE DENIED", "You don't have access to create Service Location Contacts. Please contact your system administrator."
			end if
			dim cmdInsertObj
    dim updatedSiteName
    set updatedSiteName = Request("txtSiteName")
    dim updatedSiteCode 
    set updatedSiteCode=Request("txtSiteCode")
			set cmdInsertObj = server.CreateObject("ADODB.Command")
			set cmdInsertObj.ActiveConnection = objConn
			cmdInsertObj.CommandType = adCmdText
			cmdInsertObj.CommandText = "insert into CRP.SITE_NAME_CODE(SITE_NAME,SITE_CODE,SERVICE_LOCATION_ID) values('" & updatedSiteName & "','" & updatedSiteCode & "','" & serLocId & "') "

			
			cmdInsertObj.Execute
			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT insert OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			
			end if



			strWinMessage = "Record created successfully. You can now see the new record."
		end if
    
	' set objRS =  cmdUpdateObj.Execute
   ' set sitecode =objRs("SITE_CODE")
   ' set siteName =objRs("SITE_NAME")

end select



%>

<html>
<head>
    <meta http-equiv="Pragma" content="no-cache">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <link rel="stylesheet" type="text/css" href="stylesheets/styles.css">
    <script type="text/javascript" src="GeneralJavaFunctions.js"></script>
    <script type="text/javascript" src="AccessLevels.js"></script>
    <title>Service Location Contact Detail</title>



    <script id="clientEventHandlersJS" language="javascript">
        var intAccessLevel = <%=intAccessLevel%>;
  
        function btnClose_onclick()
        {
            parent.opener.location.href = "/ServLocDetail.asp?ServLocID="+ "<%= serLocId %>";
          //  parent.opener.iFrame_display();
            window.close();
        }

        function frmSiteNameCode_onsubmit()
        {
            if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update || (intAccessLevel & intConst_Access_Create) != intConst_Access_Create)
            {
                alert('Access Denied. Please contact your system administrator.');
                return false;
            }

            document.frmServLocContact.hdnFrmAction.value = "SAVE";
            boolNeedToSave = false;
            document.forms[0].submit();
            return true;
        }

        function body_onUnload(){
            debugger;
            parent.opener.location.href = "/ServLocDetail.asp?ServLocID="+ "<%= serLocId %>";
          //  parent.opener.iFrame_display();
        }

   
    </script>
</head>
<body onunload="body_onUnload();">

    <form name="frmServLocContact" action="<%=ASP_NAME%>" language="javascript">

        <input name="SiteId" type="hidden" value="<%= strSiteID %>">
        <input name="ServLocID" type="hidden" value="<%= serLocId %>">
        <input name="hdnFrmAction" type="hidden" value="<%= Request("action") %>" />
        <table border="0" width="100%">
            <thead>
                <tr>
                    <td colspan="2">Site Name and Code Maintenance</td>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td align="RIGHT" width="20%" nowrap>Site Name<font color="red">*</font></td>
                    <td colspan="3" width="80%">
                        <input id="txtSiteName" name="txtSiteName" style="height: 21px; width: 500px"
                            value="<%= objRS("SITE_NAME").Value%>">
                    </td>
                </tr>
                <tr>
                    <td align="RIGHT" width="20%" nowrap>Site Code<font color="red">*</font></td>
                    <td colspan="3" width="80%">
                        <input id="txtSiteCode" name="txtSiteCode" style="height: 21px; width: 500px"
                            value="<%= objRS("SITE_CODE").Value%>">
                    </td>
                </tr>
            </tbody>
        </table>

        <table>
            <tr>
                <td align="right" colspan="5">
                    <input id="btnClose" name="btnClose" type="button" value="Close" style="width: 2cm" onclick="return btnClose_onclick();">&nbsp;&nbsp;
		
			<input id="btnReset" name="btnReset" type="reset" value="Reset" style="width: 2cm">&nbsp;&nbsp;
			
			<input id="btnSave" name="btnSave" type="button" value="Save" style="width: 2cm" onclick="return frmSiteNameCode_onsubmit();">&nbsp;&nbsp;
                </td>
            </tr>
        </table>



    </form>

</body>
</html>
