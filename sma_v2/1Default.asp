<%@ Language=VBScript %>
<% Option Explicit  %>
<% Response.Buffer = True%>
<!--#include file="smaProcs.inc"-->
<!--#include file="smaConstants.inc"-->
<!--#include file="SMA_Env.inc"-->

<!--
***************************************************************************************************
* Name:		default.asp
*
* Purpose:	his page reads users's search critiera and bring back a list of matching Customer
*			Service records.
*
* Created By:
* Edited by:
***************************************************************************************************
		 Date		Author			Changes/enhancements made
		20-Jul-01	 DTy		Display correct error message when error occurs while openning
								  the database instead of displaying as user id/password errors.
								Fix logon problem which allows users with no SMA roles to logon
								  to SMA2. Change error message to: 'User has no assigned SMA role'.
		28-Sep-12	ACheung		Changes code to adopt to LDAP
		Jan 2015    LC
					getUserRolePolicyList("SMA2",strUserName, userRoleNames,userPolicyNames) where
					   userRoleNames list the roles for the user and
					   userPolicyNames list the policies, which are roles in permit of format of "SMA2%" for the user
					getPolicyList("SMA2",appRoleNbr, PolicyIds, PolicyActions, PolicyRoles) where
					   PolicyIds(i), PolicyActions(i), Policyroles(i,0-n) saves the i-th
					   policyId, PolicyAction and policyroles for this i-th policyId/policyAction

***************************************************************************************************
-->
<%


'On Error Resume Next
Dim strUserIDweb
Dim strUserName


dim strError
' SSO workaround for testing
'if Request.Cookies("EIDSSOpro") = "" then
'      Response.AddHeader "Set-Cookie", "EIDSSOpro=" + Mid(Request.ServerVariables("QUERY_STRING"),11) + "; HttpOnly"
'end if


strUserName = Session("username")

If UCase(Request.QueryString("redir")) = "Y" Then
	strError = "<BR>Your Session has expired. Any changes made on your previous screen were not saved to the database."
	strError = strError & "<BR><BR> To fix this problem please completely shut down all open SMA windows and try logon again.  If this problem persists, please contact you system administrator."
	response.write strError
	response.end
Else
	strError = ""
End If


dim i,j
dim userRoleNames()		' a list of user role names
dim userPolicyNames()
dim Policies()
dim Role, strRoles, strSMARoles

If (strUserName <> "") Then

	'if getUserRolePolicyList("SMA2",strUserName, userRoleNames,userPolicyNames) then
	if getUserRolePolicyList(SMA2,strUserName, userRoleNames,userPolicyNames) then

		   'if UBound(userPolicyNames) > 0 then
			'   for i = 0 to UBound(userPolicyNames)-1
			'      response.write " User Policy  " & i & ":" & userPolicyNames(i) &"<BR>"
			'   next
		  ' else
		   '	   response.write "You are not authorized to access SMA2. Please submit your request to access!"
		   '	   response.end
		  ' end if
		  if UBound(userPolicyNames) = 0 then
				response.write "You are not authorized to access SMA2. Please submit your request to access!"
		   	    response.end
		  end if

			
	
	      ' now to sort user roles
	    
	     
	       if UBound(userRoleNames) > 0 then     ' User is a secured user
	            userRoleNames(UBound(userRoleNames))="NON-LCD"
				  for i = 0 to UBound(userRoleNames)
				     for j = i+1 to UBound(userRoleNames)

					     dim a, b
					     a = userRoleNames(i)
					     b = userRoleNames(j)
					     if StrComp(a,b,1) = 1 then
					          ' response.write "compared, result -1 <BR>"
					          '  tmpuserRole=userRoleNames(i)
					           ' userRoleNames(i) = userRoleNames(j)
					          '  userRoleNames(j) = tmpuserRole
					           userRoleNames(j)=a
							   userRoleNames(i)=b
					     end if
				     next
				     
				     
				     if i=0 then
				      strRoles=userRoleNames(i)
				     else
				       strRoles = strRoles & ";" & userRoleNames(i)
				     end if
				   next
			 else
			       'response.write "User has no secured access"
			       strRoles=""
			 end if
 		   


		  'response.write "<BR> After sort, roles are : <BR>"
		  'for i = 0 to UBound(userRoleNames)-1
		  '  response.write " User Role  " & i & "," & userRoleNames(i) &"<BR>"
		  'next
		  'response.end 
	else
	  response.write "There is no Roles defined for SMA2 application,  please contact SMA2 application support!"
	  response.end
	end if
 

 
	'if getPolicyList("SMA2", Policies) then
	if getPolicyList(SMA2, Policies) then

	 	'response.write "Now to generate USER access level matrix..."
	 	'response.write "UBound(Policies,2)=" &UBound(Policies,2)
	 	'response.end
		Dim UserAccessLevel, n
		Set UserAccessLevel = Server.CreateObject("Scripting.Dictionary")

		dim prevalue
		For Each Role In userPolicyNames
		      ' UserAccessLevel.add Policies(0,i), Policies(1,i)
		       strSMARoles = strSMARoles &";"&Role
		       For i = 0 to UBound(Policies, 2)-1
		             If Role = Policies(2,i) Then
		                ' response.write "Role = " &Role &" and Policy Role =" &Policies(2,i)
		                if UserAccessLevel.Exists(Policies(0,i)) then
		                     ' response.write "Duplicate " & Policies(0,i)
		                      prevalue=UserAccessLevel.Item(Policies(0,i))
		                     ' response.write "pre value is " & prevalue &" and current one is " &Policies(1,i)
		                     ' response.write " , which creates " & (Cint(Policies(1,i)) or Cint(prevalue))
							  UserAccessLevel.Item(Policies(0,i))= (Cint(Policies(1,i)) or Cint(prevalue))
							  UserAccessLevel.remove(Policies(0,i))
							  UserAccessLevel.add Policies(0,i), (Cint(Policies(1,i)) or Cint(prevalue))
						else
		                      UserAccessLevel.add Policies(0,i), Policies(1,i)

		                end if
		             end if
		       Next
		Next
	else
	    response.write "No policy is defined for SMA2 application, please contact SMA2 application support!"
	    response.end
	end if

	'a=UserAccessLevel.Keys
	'for i=0 to UserAccessLevel.Count-1
	'  Response.Write(a(i))
	'  response.write "  "
	'  response.write UserAccessLevel.Item(a(i))
	'  Response.Write("<br>")
	'next

   'response.write strSMARoles
  ' response.end


   	dim strConnect, strSQL, objConn, objPass
	if len(strRoles) > 0  then
		
		
	    strSQL = " Select AG.ORACLE_ID from crp_SEC.ACCESS_GROUP_ID ag where oracle_id like 'APP_SMA_%' group by AG.ORACLE_ID having LISTAGG(upper(AG.access_GROUP), ';') " &_
				" WITHIN GROUP (order by AG.access_GROUP) = upper('" & strRoles &"')"





		strConnect = Decrypt(strConstSConnectString)
		'response.write strConnect
		'response.end
		
		'strConnect = strConstSConnectString

		Err.Clear()
		Set objConn = Server.CreateObject("ADODB.Connection")
		objConn.ConnectionString = strConnect
		objConn.Open
		
		If err Then
			DisplayError "BACK", "", err.Number, "Cannot connect to database...  Re-try again.", err.Description
			Response.End
		End If

		Set objPass = Server.CreateObject("ADODB.Recordset")
		objPass.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
		If objPass.EOF Then
			response.write "user has no valid ID to db"
			response.end
		Else

			strConnect=ucase(objPass.Fields("ORACLE_ID"))
			strConnect=Left(strConnect, len(strConnect)-3)
			'response.write strConnect
			'response.end
			strConnect=Decrypt(getConnString(strConnect))
			'strConnect=getConnString(ucase(objPass.Fields("ORACLE_ID")))

		End If
		objPass.Close
		Set objPass = Nothing
    else
    	strConnect = Decrypt(strConstConnectString) 		'match the one in sma_env.inc
	end if

   ' response.write strConnect
   ' response.end


	Session(strConst_Logon)=now
	Set Session("UserAccessLevel") = UserAccessLevel
	Session("username") = strUserName
	Session("userRoles") = strRoles
	Session("ConnectString") = strConnect
	Session("SMARoles") = strSMARoles


End If
%>

<html>
<head>
<title>Service Management Application</title>

<!-- Dropdown menu stuff -->
<link href='menu.css' rel='stylesheet'>
<script src='jquery.min.js' type='text/javascript'></script>

<!-- Script to control the dropdown menu -->
<script type="text/javascript">
	
/*
function pause(millis) 
 {
 var date = new Date();
 var curDate = null;
 do { curDate = new Date(); } 
 while(curDate-date < millis);
 } 
*/
	
$(function(){
    $("ul.dropdown li.menu").hover(function(){
        $(this).addClass("hover");
        $('ul:first',this).css('visibility', 'visible');
    }, function(){
//    	  pause(100);
        $(this).removeClass("hover");
        $('ul:first',this).css('visibility', 'hidden'); 
    });

    $("ul.dropdown li ul li:has(ul)").find("a:first").append(" &raquo; ");

});

</script>
<SCRIPT language="javascript" type="text/javascript">
function fct_selApp() {
//***************************************************************************************************
// Function:	fct_selApp															                *
// Purpose:		To open a new browser window with the selected application							*
//																									*
// Created By:	Gilles Archer 09/19/2000															*
//																									*																				*
//***************************************************************************************************
	var strAppName = selApps.value;

	switch (strAppName){
		case 'SRTII':
			selApps.selectedIndex = 0;
			window.open('<% response.write(SRT2_URL) %>');
			break;
		case 'ESDReports':
			selApps.selectedIndex = 0;
			window.open('<% response.write(ESD_REP_URL) %>');
			break;
		case 'SMA2':
			//do nothing
	}
}
</SCRIPT>
</head>
<BODY bgcolor="#50197b" bottommargin="0" leftmargin="0" rightmargin="0" topmargin="0">

<table border="1" width="100%" height="98%"><tr><td>

<TABLE border="0" cellpadding="0" cellspacing="0" width="100%">
	<TR>
		<TD width="150"><IMG height="49" src="images/top1a.gif" width="165"></TD>
		<TD width="100%" colspan="2" align="right" nowrap>
		<INPUT id="PageTitle" name="PageTitle" value="Service Management Application" style="BACKGROUND-COLOR: #50197b; BORDER-BOTTOM: #50197b; BORDER-LEFT: #50197b; BORDER-RIGHT: #50197b; BORDER-TOP: #50197b; COLOR: white; CURSOR: default; FONT-FAMILY: ; FONT-SIZE: large; FONT-WEIGHT: bold; height: 32px; TEXT-ALIGN: right; width: 613px" align="right" border="0" readOnly size="100">&nbsp;
		</TD>
	</TR>
	<TR>
		<TD width="150" valign=top>
		<P align="center">
		<SELECT id="selApps" name="selApps" onchange="fct_selApp();" style="BACKGROUND-COLOR: #50197b; COLOR: white; CURSOR: hand; FONT-WEIGHT: bold">
			<OPTION selected value="SMA2">SMA 2</OPTION>
			<OPTION value="SRTII">SRT II</OPTION>
			<OPTION value="ESDReports">ESD Reports</OPTION>
		</SELECT></P></TD>
    <TD width="100%" valign="top" align="center">
		<TABLE align="center"><tr><td>
			<ul class="dropdown">&nbsp;
				<li class="menu"><a href="#">Customer</a>
			    	<ul>
			        	<li><a href="SearchFrame.asp?fraSrc=Cust" target="mainFrame">Search</a></li>
			          <li><a href="SearchFrame.asp?fraSrc=Address" target="mainFrame">Address</a></li>
			          <li><a href="SearchFrame.asp?fraSrc=ServLoc" target="mainFrame">Service Location</a></li>
			          <li><a href="SearchFrame.asp?fraSrc=Contact" target="mainFrame">Contact</a></li>
				       	<li><a href="SearchFrame.asp?fraSrc=ContactRole" target="mainFrame">Contact Role</a></li>
			  	     	<li><a href="SearchFrame.asp?fraSrc=Cust_CP" target="mainFrame">Customer Profile and VPN info</a></li>
			        </ul>
			   </li>
			   <li class="menu"><a href="#">Customer Service</a>
			    	<ul>
			        	<li><a href="SearchFrame.asp?fraSrc=CustServ" target="mainFrame">Search</a></li>
			            <li><a href="EmailSetUpList.asp" target="mainFrame">Email Setup</a></li>
			            <li><a href="SearchFrame.asp?fraSrc=CustServ_CP" target="mainFrame">Customer Profile and VPN info</a></li>
			            <li><a href="SearchFrame.asp?fraSrc=FCustServ" target="mainFrame">Services with Feature</a></li>
			        </ul>
			    </li>
			    <li class="menu"><a href="#">Facility</a>
			       <ul>
			        	<li><a href="SearchFrame.asp?fraSrc=Facility" target="mainFrame">Search</a></li>
			        </ul>
			    </li>
			    <li class="menu"><a href="#">Asset</a>
			       <ul>
			        	<li><a href="SearchFrame.asp?fraSrc=Asset" target="mainFrame">Search</a></li>
			          <li class="menu"><a href="#">Asset Catalogue</a>
						       <ul>
			      			  	<li><a href="SearchFrame.asp?fraSrc=AssetCatalogue" target="mainFrame">Asset Catalogue</a></li>
			       			  	<li><a href="SearchFrame.asp?fraSrc=Make" target="mainFrame">Make</a></li>
			     		  	  	<li><a href="SearchFrame.asp?fraSrc=Model" target="mainFrame">Model</a></li>
			     			    	<li><a href="SearchFrame.asp?fraSrc=PartNum" target="mainFrame">Part Number</a></li>
			        		 </ul>
			   				</li>
			         <li class="menu"><a href="#">Asset Classification</a>
						       <ul>
			      			  	<li><a href="SearchFrame.asp?fraSrc=AssetClass" target="mainFrame">Asset Class</a></li>
			       			  	<li><a href="SearchFrame.asp?fraSrc=AssetSubclass" target="mainFrame">Asset Subclass</a></li>
			     		  	  	<li><a href="SearchFrame.asp?fraSrc=AssetType" target="mainFrame">Asset Type</a></li>
			        		 </ul>
			   				</li>
			        </ul>
			    </li>
			    <li class="menu"><a href="#">Managed Objects</a>
			       <ul>
			        	<li><a href="SearchFrame.asp?fraSrc=ManagedObjects" target="mainFrame">Search</a></li>
			        </ul>
			    </li>
			    <li class="menu"><a href="#">PVC</a>
			       <ul>
			        	<li><a href="SearchFrame.asp?fraSrc=FacilityPVC" target="mainFrame">Search</a></li>
			        </ul>
			    </li>
			   <li class="menu"><a href="#">Correlation</a>
			       <ul>
			        	<li><a href="SearchFrame.asp?fraSrc=Correlation" target="mainFrame">Search</a></li>
			        </ul>
			    </li>
			        <li class="menu"><a href="#">Administration</a>
			       <ul>
			           <li class="menu"><a href="#">Service Catalogue</a>
						       <ul>
			      			  	<li><a href="SearchFrame.asp?fraSrc=ServiceType" target="mainFrame">Service Type</a></li>
			       			  	<li><a href="SearchFrame.asp?fraSrc=ServiceTypeA" target="mainFrame">Service Type Attributes</a></li>
			     		  	  	<li><a href="SearchFrame.asp?fraSrc=ServiceInstA" target="mainFrame">Service Instance Attributes</a></li>
			     			    	<li><a href="SearchFrame.asp?fraSrc=ServiceTypeKA" target="mainFrame">Kenan Package Component Search</a></li>
			    			    	<li><a href="SearchFrame.asp?fraSrc=SLA" target="mainFrame">Service Level Agreement</a></li>
			    			    	<li><a href="SearchFrame.asp?fraSrc=ServiceCategory" target="mainFrame">Service Category</a></li>
			    			    	<li><a href="SearchFrame.asp?fraSrc=LOB" target="mainFrame">Line of Business</a></li>
			    			    	<li><a href="SearchFrame.asp?fraSrc=Schedule" target="mainFrame">Schedule Definition</a></li>
			    			    	<li><a href="SearchFrame.asp?fraSrc=Holiday" target="mainFrame">Holiday Definition</a></li>
			        		 </ul>
			   				</li>
			         <li class="menu"><a href="#">Municipalities</a>
						       <ul>
			      			  	<li><a href="SearchFrame.asp?fraSrc=City" target="mainFrame">City</a></li>
			       			  	<li><a href="SearchFrame.asp?fraSrc=Province" target="mainFrame">Province</a></li>
			     		  	  	<li><a href="SearchFrame.asp?fraSrc=Country" target="mainFrame">Country</a></li>
			        		 </ul>
			   				</li>
			         <li class="menu"><a href="#">POS PLUS</a>
						       <ul>
			      			  	<li><a href="SearchFrame.asp?fraSrc=RSAST3Tier3" target="mainFrame">Tier 3</a></li>
			        		 </ul>
			   				</li>
			         <li class="menu"><a href="#">Cleanup Functions</a>
						       <ul>
			      			  	<li><a href="SearchFrame.asp?fraSrc=XLSEntry" target="mainFrame">Validation Spreadsheets</a></li>
			      			  	<li><a href="SearchFrame.asp?fraSrc=CustClean" target="mainFrame">Customer Cleanup</a></li>
			      			  	<li><a href="SearchFrame.asp?fraSrc=ContactClean" target="mainFrame">Contact Cleanup</a></li>
			      			  	<li><a href="SearchFrame.asp?fraSrc=SmartDataFix" target="mainFrame">Smart Data Fix</a></li>
			        		 </ul>
			   				</li>
			       </ul>
			    </li>
			</ul>
		</td></tr></table>
		</TD>
	<TD width="150" align="right" valign=top nowrap>
		<P align="right">
		<IMG src="images/back_002.gif"    alt="Go Back"        onclick="parent.mainFrame.history.go(-1);                          " width="31" height="31">&nbsp;
		<IMG src="images/forward_002.gif" alt="Go Forward"     onclick="parent.mainFrame.history.go(1);                           " width="31" height="31">&nbsp;
		<IMG src="images/help.gif"        alt="SMA User Guide" onclick="window.open('help/SMA2 User Guide v2-3.doc', 'Help');" width="31" height="31">
		</P></TD></TR>
</TABLE>

</td></tr>
<tr height="100%"><td bgcolor="white">

<iframe id="mainFrame" name="mainFrame" src="SMA%20Version.htm" frameborder="0" scrolling="yes" height="100%" width="100%">

</td></tr>
</table>
</html>





