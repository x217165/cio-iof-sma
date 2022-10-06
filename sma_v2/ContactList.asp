<%@ Language=VBScript %>
<% option explicit
 on error resume next %>
<!--#include file = "smaConstants.inc" -->
<!--#include file = "databaseconnect.asp"-->
<!--#include file = "smaProcs.inc" -->
<!--
***************************************************************************************************
* Name:			ContactList.asp
*
* Purpose:		To display the results of a contact search.
*				Search criteria are chosen via ContactCriteria.asp
* Created By:	Nancy Mooney 08/16/00
***************************************************************************************************

        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       02-07-01	     DTy		bolActiveOnly should be check for 'yes' and not 'on'.
                                Exclude constacts that are:
                                  Marked as soft deleted in CONTACT.,
                                     i.e., RECORD_STATUS_IND='D'.
                                  Staff who left their employer.
                                     i.e., CONTACT.STAFF_STATUS_LCODE='Departed'.
                                Set the apostrophe properly using routineORAstring & routineHTMLstring
                                  for txtLName, txtFName, txtWorksForName & txtEmail.
       03-29-01	     DTy		Enclose form objects value with quotes in order to prevent dropping
                                  of second or subsequent words of its content.
       07-20-01	     DTy		When 'Active Only' is selected, quotes in order to prevent dropping
                                  Exclude customers that are soft deleted.
                                  Exclude addresses that are soft deleted.
       18-Feb-02	 DTy		Active customers are those whose status is either
                                  'Prospect', 'OnHold' or 'Current'.
                                Align email of returned Contact Info.
       08-Mar-02     DTy        Fix alignment problem on Excel spreadsheet.
								Add 'Contact ID' as a displayable field.
       15-Mar-02     DTy        Fix improper translation of apostrophe in address.
       30-Mar-02     DTy	Facilitate 'Contact Cleanup' Contact ID and Name lookup.
				Add Middle Name column.
       16-Oct-07     ACheung	Add new columns "Responsibility" and "PIN"

       ***************************************************************************************************

-->
<!-- start vbscript portion -->
<%
 dim strLName, strFName, strMName, strWorksForName, strWPhone, strWPhoneArea,strWPhoneMid, strWPhoneEnd, strEmail,strRegion, strWinName
 dim intContactType, bolActiveOnly
 dim strSQL, strSelectClause, strWhereClause, strFromClause, strRecordStatus, strOrderBy
 Dim rsContactList, aList
 dim intPageNumber, intPageCount,strTelusOnly, strCase


	'get submitted values (search criteria; window name)
	strLName = UCase(trim(Request.Form("txtLName")))
	strFName = UCase(trim(Request.Form("txtFName")))
	strWinName = trim(Request("hdnWinName"))
	strWorksForName = UCase(trim(Request.Form("txtWorksForName")))
	strWPhoneArea = trim(Request.Form("txtWPhoneArea"))
	strWPhoneMid = trim(Request.Form("txtWPhoneMid"))
	strWPhoneEnd = trim(Request.Form("txtWPhoneEnd"))
	strWPhone = strWPhoneArea & strWPhoneMid & strWPhoneEnd
	strEmail = Ucase(trim(Request.Form("txtEmail")))
	strRegion = Request("selRegion")
	intContactType = Request.Form("radContactType")
	bolActiveOnly = Request.Form("chkActiveOnly")
    strTelusOnly = Request.Form("hdnTelusOnly")
    strCase = Request.Form("hdnCase")

	'build query

	'no criteria selected display all
	strSelectClause = "SELECT " & _
					"distinct(t1.contact_id), " & _
					"t1.last_name, " & _
					"t1.first_name, " & _
					"t1.contact_name, " & _
					"t1.work_for_customer_id, " & _
					"t1.position_title, " & _
					"t1.work_number, " & _
					"t1.work_number_ext, " & _
					"t2.customer_name, " & _
					"t2.noc_region_lcode, " & _
					"t1.cell_number, " & _
					"t1.pager_number, " & _
					"t1.fax_number, " & _
					"t1.email_address, " & _
					"t3.building_name, " & _
					"t3.street, " & _
					"t3.municipality_name, " & _
					"t3.province_state_lcode, " & _
					"t3.country_lcode, " & _
					"t3.postal_code_zip, " & _
					"t1.middle_name, " & _
					"t1.responsibility, " & _
					"t1.pin_access "

	strFromClause = " FROM " & _
				"crp.contact t1, " & _
				"crp.customer t2, " & _
				"crp.v_address_consolidated_street t3 "

	strWhereClause = " WHERE " & _
				"t1.work_for_customer_id = t2.customer_id and " & _
				"t1.address_id = t3.address_id (+) "

	'add other search parameters to the where clause
	If len(strLName) > 0 then
      strWhereClause = strWhereClause & " AND Upper(t1.last_name) LIKE '" & routineORAstring(strLName) &"%'"
	End If

	If len(strFName) > 0 then
      strWhereClause = strWhereClause & " AND Upper(t1.first_name) LIKE '" & routineORAstring(strFName) &"%'"
	End If

	If len(strWorksForName) > 0 then
		'include alias table
		strFromClause = strFromClause & ",crp.customer_name_alias t0 "
		'join alias table to customer table and specify customer search string
		strWhereClause = strWhereClause & " and t0.customer_id = t2.customer_id and " & _
		" t0.customer_name_alias_upper like '" & routineORAstring(strWorksForName) & "%'"
	End If

	If len(strWPhone) > 0 then
      strWhereClause = strWhereClause & " AND t1.work_number = '" & strWPhone & "'"
    End If

	Select case intContactType
		case "1" 'TELUS staff
			strWhereClause = strWhereClause & " AND t1.staff_flag = 'Y' "
		case "2" 'External contacts
			strWhereClause = strWhereClause & " AND t1.staff_flag = 'N' "
		case "3" 'both - no criteria to add
	End Select

	if bolActiveOnly = "yes" then
		strRecordStatus = " and t2.customer_status_lcode IN ('Prospect', 'Current', 'OnHold') " &_
		   " and t1.record_status_ind = 'A' and " & _
		   " t2.record_status_ind = 'A' and t3.record_status_ind (+) = 'A' and " & _
		   "(t1.staff_status_lcode is null or " & _
		   "(t1.staff_status_lcode is not null and t1.staff_status_lcode <> 'Departed')) "
	else
		'display all record
		strRecordStatus = " "
	End If

	if len(strEmail) > 0 then
		strWhereClause = strWhereClause &  " AND Upper(t1.email_address) LIKE '" & strEmail &"%'"
	End if

	'region picked
	If strRegion <> "All" then
		strWhereClause = strWhereClause & " and " & _
			"t2.noc_region_lcode = '" & strRegion & "'"
	End If

	strOrderBy = " order by upper(t2.customer_name), upper(t1.last_name), upper(t1.first_name)"

	'join all pieces to make a complete query
	strSQL = strSelectClause & strFromClause & strWhereClause & strRecordStatus & strOrderBy

	'Response.Write strSQL

    'get the recordset
    set rsContactList = server.CreateObject("ADODB.Recordset")
    rsContactList.Open strSQL, objConn
    If err then
		DisplayError "BACK", "", err.Number, "ContactList.asp - Cannot open database" , err.Description
	End if

	'put recordset into array
	if not rsContactList.EOF then
		aList = rsContactList.GetRows
	else
		Response.Write "0 Records Found"
		Response.End
	end if

	'release and kill the recordset and the connection objects
	rsContactList.Close
	set rsContactList = nothing
	objConn.Close
	set objConn = nothing

	'calculate page number
	intPageCount = Int(UBound(aList,2) / intConstDisplayPageSize) + 1
	select case Request("Action")
		case "<<"	intPageNumber = 1
		case "<"	intPageNumber = Request("txtPageNumber")-1
					if intPageNumber < 1 then intPageNumber = 1
		case ">"	intPageNumber = Request("txtPageNumber") + 1
					if intPageNumber > intPageCount then intPageNumber = intPageCount
		case ">>"	intPageNumber=intPageCount
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
						set objTxtStream = objFSO.CreateTextFile(strExportPath&strRealUserID&"-contact.xls", true, false)

						if err then
							DisplayError "CLOSE", "", err.Number, "ContactList.asp - Cannot create Excel spreadsheet file due to the following reasons. Please contact your website administrator.", err.Description
						end if

						with objTxtStream
							.WriteLine "<table border=1>"

							'export the header
							.WriteLine "<TR>"
							.WriteLine "<TH>Works For</TH>"
							.WriteLine "<TH>Contact ID</TH>"
							.WriteLine "<TH>Last Name</TH>"
							.WriteLine "<TH>First Name</TH>"
							.WriteLine "<TH>Middle Name</TH>"
							.WriteLine "<TH>Region</TH>"
							.WriteLine "<TH>Email</TH>"
							.WriteLine "<TH>Work Phone</TH>"
							.WriteLine "<TH>Ext</TH>"
							.WriteLine "<TH>Building</TH>"
							.WriteLine "<TH>Address</TH>"
							.WriteLine "<TH>City</TH>"
							.WriteLine "<TH>Prov/State</TH>"
							.WriteLine "<TH>Responsibility</TH>"
							.WriteLine "<TH>PIN</TH>"
							.WriteLine "</TR>"

							'export the body
							for k = 0 to UBound(aList, 2)

								'Parse out the work phone number
	 							Dim strWorkPhoneArea,strWorkPhoneMid,strWorkPhoneEnd,strWorkPhone
	 							strWorkPhoneArea = mid(alist(6,k),1,3)
	 							strWorkPhoneMid = mid(alist(6,k),4,3)
	 							strWorkPhoneEnd = mid(alist(6,k),7,4)
	 							strWorkPhone = "(" & strWorkPhoneArea & ") " & strWorkPhoneMid & "-" & strWorkPhoneEnd
	 							If strWorkPhone = "() -" then
	 								strWorkPhone = ""
	 							End If
								.WriteLine "<TR>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(8,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(0,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(1,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(2,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(20,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(9,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(13,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(strWorkPhone)&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(7,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(14,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(15,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(16,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(17,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(21,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(22,k))&"&nbsp;</TD>"
								.WriteLine "</TR>"
							next
							.WriteLine "</table>"
						end with
						objTxtStream.Close
						strsql = "<script type=""text/javascript"">document.location=""export/"&strRealUserID&"-contact.xls"";</script>"
						Response.Write strsql
						Response.End
'						Response.redirect "export/"&strRealUserID&"-contact.xls"
					elseif Request("txtGoToPageNo") <> "" then
						intPageNumber = CInt(Request("txtGoToPageNo"))
					else
						intPageNumber = 1
					end if
	end select

	if intPageNumber < 1 then intPageNumber = 1
	if intPageNumber > intPageCount then intPageNumber = intPageCount

	dim k,m,n
	m = (intPageNumber - 1) * intConstDisplayPageSize
	n = (intPageNumber) * intConstDisplayPageSize - 1
	if n > UBound(aList,2) then
		n=UBound(aList,2)
	end if

	'check if the client is still connected just before sending any html to the browser
	if Response.IsClientConnected = false then
		Response.End
	end if

	'catch any unexpected error
	if err then
		DisplayError "BACK", "", err.Number, "Unexpected error.", err.Description
	end if

%>
<!-- end vbscript portion -->

<HTML>
<HEAD>
	<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<LINK rel="stylesheet" type="text/css" href="stylesheets/styles.css">
	<TITLE>SMA - Contact</TITLE>
	<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
	<script ID=clientEventHandlersJS type="text/javascript">
	<!--

setPageTitle("SMA - Contacts");

	// navigation

function go_back(strCase, lngContactID, strContactName,strLName, strFName, strMName, strWorkFor, strPosition, strWP, strWExt, strCP, strFP, strPP, strEmail, strBuilding, strStreet, strCPC, strPC, strResponsibility, bolPIN){

var strContactInfo
strContactInfo = 'Works for:\t' + strWorkFor + '\nPosition:\t' + '\nWork # :\t' + strWP + ' Ext: ' + strWExt + '\nCell # :\t' + strCP + '\nPager # :\t' + strPP + '\nFax # :\t' + strFP + '\nEmail:\t' + strEmail + '\nBuilding:\t' + strBuilding + '\nAddress:\t' + strStreet + '\n\t' + strCPC + '\n\t' + strPC + '\n\t' + strResponsibility + '\n\t' + bolPIN;

  switch (strCase) {

	case 'A':
			parent.opener.document.forms[0].hdnContactID1.value = lngContactID;
			parent.opener.document.forms[0].txtLName1.value = strLName;
			parent.opener.document.forms[0].txtFName1.value = strFName;
			parent.opener.document.forms[0].txtContactName1.value = strContactName;
			break ;
	case 'B':
			parent.opener.document.forms[0].hdnContactID2.value = lngContactID;
			parent.opener.document.forms[0].txtLName2.value = strLName;
			parent.opener.document.forms[0].txtFName2.value = strFName;
			parent.opener.document.forms[0].txtContactName2.value = strContactName;
			break;
	case 'E':
			if (strEmail == '') {alert('The selected contact does not have a valid email address.');return;}
			//populate email fields
			switch (parent.opener.document.frmEmail.hdnDestination.value) {
				case 'to':
					if (parent.opener.document.frmEmail.txtTO.value != '') {parent.opener.document.frmEmail.txtTO.value += '; '}
					parent.opener.document.frmEmail.txtTO.value += strFName+" "+strLName+" <"+strEmail+">";
					break;
				case 'cc':
					if (parent.opener.document.frmEmail.txtCC.value != '') {parent.opener.document.frmEmail.txtCC.value += '; '}
					parent.opener.document.frmEmail.txtCC.value += strFName+" "+strLName+" <"+strEmail+">";
					break;
				case 'bcc':
					if (parent.opener.document.frmEmail.txtBCC.value != '') {parent.opener.document.frmEmail.txtBCC.value += '; '}
					parent.opener.document.frmEmail.txtBCC.value += strFName+" "+strLName+" <"+strEmail+">";
					break;
			}
			break;
	case 'M':
			parent.opener.document.forms[0].hdnManagerContactID.value = lngContactID;
			parent.opener.document.forms[0].txtManagerContactName.value = strContactName;
			break;

	case 'F': // this condition handles FR Contact in ContactCleanEntry.asp
			parent.opener.document.forms[0].txtFRContact.value = "(" + lngContactID + ") " + strLName + ", " + strFName + " " + strMName;
			parent.opener.document.forms[0].hdnFRContactID.value = lngContactID;
			parent.opener.document.forms[0].hdnFRContactLName.value = strLName;
			parent.opener.document.forms[0].hdnFRContactFName.value = strFName;
			parent.opener.document.forms[0].hdnFRContactMName.value = strMName;
			break;

	case 'T': // this condition handles TO Contact in ContactCleanEntry.asp
			parent.opener.document.forms[0].txtTOContact.value = "(" + lngContactID + ") " + strLName + ", " + strFName + " " + strMName;
			parent.opener.document.forms[0].hdnTOContactID.value = lngContactID;
			parent.opener.document.forms[0].hdnTOContactLName.value = strLName;
			parent.opener.document.forms[0].hdnTOContactFName.value = strFName;
			parent.opener.document.forms[0].hdnTOContactMName.value = strMName;
			break;

	default:
	  try {
			if (document.frmContactList.hdnTelusOnly.value=="yes"){
				parent.opener.document.forms[0].txtcustodian.value = strLName+", "+strFName;
				parent.opener.document.forms[0].hdnStaffID.value =lngContactID;
			}
			else{
				parent.opener.document.forms[0].hdnContactID.value = lngContactID;
				parent.opener.document.forms[0].txtLName.value = strLName;
				parent.opener.document.forms[0].txtFName.value = strFName;
				parent.opener.document.forms[0].txtContactName.value = strContactName;
				parent.opener.document.forms[0].txtContactInfo.value = strContactInfo;
			} //end if
		  } // try
	  catch (e){}
		break ;
	}
	parent.window.close ();
}

	// buttons

	function btnEdit_onclick(lngContactID){
	var url ;

	url = 'CustDetail.asp?ContactID=' + lngContactID;
	self.open(url,'Popup','top=50, left=100, height=600, width=800' );
	}

	//-->
	</SCRIPT>

</HEAD>

<BODY>
<FORM name=frmContactList method=post action="ContactList.asp" >

	<INPUT name=txtLName		type=hidden value="<%=routineHTMLstring(strLName)%>" >
	<INPUT name=txtFName		type=hidden value="<%=routineHTMLstring(strFName)%>" >
	<INPUT name=txtWorksForName type=hidden value="<%=routineHTMLstring(strWorksForName)%>" >
	<INPUT name=txtEmail		type=hidden value="<%=strEmail%>" >
	<INPUT name=txtWPhoneArea	type=hidden value="<%=strWPhoneArea%>" >
	<INPUT name=txtWPhoneMid	type=hidden value="<%=strWPhoneMid%>" >
	<INPUT name=txtWPhoneEnd	type=hidden value="<%=strWPhone%>" >
	<INPUT name=selRegion		type=hidden value="<%=strRegion%>" >
	<INPUT name=radContactType	type=hidden value="<%=intContactType%>" >
	<INPUT name=chkActiveOnly	type=hidden value="<%=bolActiveOnly%>" >
	<INPUT name=hdnWinName		type=hidden value="<%=strWinName%>" >
	<INPUT name=hdnTelusOnly	type=hidden value="<%=strTelusOnly%>" >
	<INPUT name=hdnExport		type=hidden value>

<TABLE border=1 cellpadding=2 cellspacing=0 width=100%>
	<thead align=left >
		<tr><td align=left colspan=15>Contact Results</td></tr>
	</thead>
	<TBODY>
	<TR>
        <TH align=left nowrap>Works For</TH>
        <TH align=left nowrap>Contact ID</TH>
        <TH align=left nowrap>Last Name</TH>
        <TH align=left nowrap>First Name</TH>
        <TH align=left nowrap>Middle Name</TH>
        <TH align=left nowrap>Region</TH>
        <TH align=left nowrap>Email</th>
        <TH align=left nowrap>Work Phone</TH>
        <TH align=left nowrap>Ext</TH>
        <TH align=left nowrap>Building</TH>
        <TH align=left nowrap>Address</TH>
        <TH align=left nowrap>City</TH>
        <TH align=left nowrap>Prov/State</TH>
        <TH align=left nowrap>Responsibility</TH>
        <TH align=left nowrap>PIN</TH>
    </TR>
	<%
	'Response.Write (strWinName & "<BR>")
	'display the table
	 for k = m to n

	 	'Parse out the work phone number
	 	Dim strWPArea,strWPMid,strWPEnd,strWP
	 	strWPArea = mid(alist(6,k),1,3)
	 	strWPMid = mid(alist(6,k),4,3)
	 	strWPEnd = mid(alist(6,k),7,4)
	 	strWP = "(" & strWPArea & ") " & strWPMid & "-" & strWPEnd
	 	If strWP = "() -" then
	 		strWP = ""
	 	End If

	 	if strWinName = "Popup" then
	 		'create the string for the go_back function

			'Parse out the cell phone number
	 		Dim strCPArea,strCPMid,strCPEnd,strCP
	 		strCPArea = mid(alist(10,k),1,3)
	 		strCPMid = mid(alist(10,k),4,3)
	 		strCPEnd = mid(alist(10,k),7,4)
	 		strCP = "(" & strCPArea & ") " & strCPMid & "-" & strCPEnd
	 		If strCP = "() -" then
	 			strCP = ""
	 		End If

	 		'Parse out the pager number
	 		Dim strPPArea,strPPMid,strPPEnd,strPP
	 		strPPArea = mid(alist(11,k),1,3)
	 		strPPMid = mid(alist(11,k),4,3)
	 		strPPEnd = mid(alist(11,k),7,4)
	 		strPP = "(" & strPPArea & ") " & strPPMid & "-" & strPPEnd
	 		If strPP = "() -" then
	 			strPP = ""
	 		End If

	 		'Parse out the fax number
	 		Dim strFPArea,strFPMid,strFPEnd,strFP
	 		strFPArea = mid(alist(12,k),1,3)
	 		strFPMid = mid(alist(12,k),4,3)
	 		strFPEnd = mid(alist(12,k),7,4)
	 		strFP = "(" & strFPArea & ") " & strFPMid & "-" & strFPEnd
	 		If strFP = "() -" then
	 			strFP = ""
	 		End If

	 		'parse out postal code
	 		dim strPCBegin, strPCEnd, intPClen, strPC
	 		strPC = aList(19,k)
	 		select case aList(18,k)
	 				case "CA"
	 				strPCBegin = mid(strPC,1,3)
	 				strPCEnd = mid(strPC,4,3)
	 				strPC = strPCBegin & " " & strPCEnd
	 			case "US"
	 				intPClen = len(strPC)
	 				strPCBegin = mid(strPC,1,5)
	 				strPCEnd = mid(strPC,6,intPCLen-5)
	 				strPC = strPCBegin & " " & strPCEnd
	 		end select

	 		'create the rest of the variables for go_back()
	 		dim lngContactID, strContactName, strWorkFor, strPosition, strWExt, strEmail2, strBuilding, strStreet, strCity, strProv, strCountry, strCPC, strResponsibility, bolPIN

	 		lngContactID = aList(0,k)
	 		strContactName = routineJavascriptString(aList(3,k))
	 		strLName = routineJavascriptString(aList(1,k))
	 		strFName = routineJavascriptString(aList(2,k))
	 		strMName = routineJavascriptString(aList(20,k))
	 		strWorkFor = routineJavascriptString(aList(8,k))
	 		strPosition = routineJavascriptString(aList(5,k))
	 		strWExt = routineJavascriptString(aList(7,k))
	 		strEmail2 = routineJavascriptString(aList(13,k))
	 		strBuilding = routineJavascriptString(aList(14,k))
	 		strStreet = routineJavascriptString(aList(15,k))
			strResponsibility = routineJavascriptString(aList(21,k))
			bolPIN = routineJavascriptString(aList(22,k))

	 		'create CPC (City/Province/Country)
			if  aList(16,k) <> "" then
				strCity = aList(16,k) & " "
			else
				strCity = ""
			end if
			if aList(17,k) <> "" then
				strProv = aList(17,k) & " "
			else
				strProv = ""
			end if
			if aList(18,k) <> "" then
				strCountry = aList(18,k)
			else
				strCountry = ""
			end if
			strCPC = routineJavascriptString(strCity & strProv & strCountry)
	 		'strCPC = routineJavascriptString(aList(16,k)) & " " & routineJavascriptString(aList(17,k)) & " " & routineJavascriptString(aList(18,k))
		end if

	 	'alternate row background color
	 	if Int(k/2) = k/2 then
	 		Response.Write "<tr bgcolor=White>"
	 	else
	 		Response.Write "<tr>"
	 	end if
 
	 	if strWinName= "Popup" then
			'create the rows to be displayed
	 		Response.Write "<td nowrap><a href=""#"" onClick=""return go_back('" & strCase & "', '" & lngContactID & "', '" & strContactName & "', '" & strLName & "', '" & strFName & "', '" & strMName & "', '" & strWorkFor & "', '" & strPosition & "', '" & strWP & "', '" & strWExt & "', '" & strCP & "', '" & strFP & "', '" & strPP & "', '" & strEmail2 & "', '" & strBuilding & "', '" & strStreet & "', '" & strCPC & "', '" & strPC & "', '" & strResponsibility & "', '" & bolPIN & "')"">" &aList(8,k) & "</a>&nbsp;</td>"&vbCrLf
	 		Response.Write "<td nowrap><a href=""#"" onClick=""return go_back('" & strCase & "', '" & lngContactID & "', '" & strContactName & "', '" & strLName & "', '" & strFName & "', '" & strMName & "', '" & strWorkFor & "', '" & strPosition & "', '" & strWP & "', '" & strWExt & "', '" & strCP & "', '" & strFP & "', '" & strPP & "', '" & strEmail2 & "', '" & strBuilding & "', '" & strStreet & "', '" & strCPC & "', '" & strPC & "', '" & strResponsibility & "', '" & bolPIN & "')"">" &aList(0,k) & "</a>&nbsp;</td>"&vbCrLf
	 		Response.Write "<td nowrap><a href=""#"" onClick=""return go_back('" & strCase & "', '" & lngContactID & "', '" & strContactName & "', '" & strLName & "', '" & strFName & "', '" & strMName & "', '" & strWorkFor & "', '" & strPosition & "', '" & strWP & "', '" & strWExt & "', '" & strCP & "', '" & strFP & "', '" & strPP & "', '" & strEmail2 & "', '" & strBuilding & "', '" & strStreet & "', '" & strCPC & "', '" & strPC & "', '" & strResponsibility & "', '" & bolPIN & "')"">" &aList(1,k) & "</a>&nbsp;</td>"&vbCrLf
	 		Response.Write "<td nowrap><a href=""#"" onClick=""return go_back('" & strCase & "', '" & lngContactID & "', '" & strContactName & "', '" & strLName & "', '" & strFName & "', '" & strMName & "', '" & strWorkFor & "', '" & strPosition & "', '" & strWP & "', '" & strWExt & "', '" & strCP & "', '" & strFP & "', '" & strPP & "', '" & strEmail2 & "', '" & strBuilding & "', '" & strStreet & "', '" & strCPC & "', '" & strPC & "', '" & strResponsibility & "', '" & bolPIN & "')"">" &aList(2,k) & "</a>&nbsp;</td>"&vbCrLf
	 		Response.Write "<td nowrap><a href=""#"" onClick=""return go_back('" & strCase & "', '" & lngContactID & "', '" & strContactName & "', '" & strLName & "', '" & strFName & "', '" & strMName & "', '" & strWorkFor & "', '" & strPosition & "', '" & strWP & "', '" & strWExt & "', '" & strCP & "', '" & strFP & "', '" & strPP & "', '" & strEmail2 & "', '" & strBuilding & "', '" & strStreet & "', '" & strCPC & "', '" & strPC & "', '" & strResponsibility & "', '" & bolPIN & "')"">" &aList(20,k) & "</a>&nbsp;</td>"&vbCrLf
	 		Response.Write "<td nowrap><a href=""#"" onClick=""return go_back('" & strCase & "', '" & lngContactID & "', '" & strContactName & "', '" & strLName & "', '" & strFName & "', '" & strMName & "', '" & strWorkFor & "', '" & strPosition & "', '" & strWP & "', '" & strWExt & "', '" & strCP & "', '" & strFP & "', '" & strPP & "', '" & strEmail2 & "', '" & strBuilding & "', '" & strStreet & "', '" & strCPC & "', '" & strPC & "', '" & strResponsibility & "', '" & bolPIN & "')"">" &aList(9,k) & "</a>&nbsp;</td>"&vbCrLf
	 		Response.Write "<td nowrap><a href=""#"" onClick=""return go_back('" & strCase & "', '" & lngContactID & "', '" & strContactName & "', '" & strLName & "', '" & strFName & "', '" & strMName & "', '" & strWorkFor & "', '" & strPosition & "', '" & strWP & "', '" & strWExt & "', '" & strCP & "', '" & strFP & "', '" & strPP & "', '" & strEmail2 & "', '" & strBuilding & "', '" & strStreet & "', '" & strCPC & "', '" & strPC & "', '" & strResponsibility & "', '" & bolPIN & "')"">" &aList(13,k)& "</a>&nbsp;</td>"&vbCrLf
	 		Response.Write "<td nowrap><a href=""#"" onClick=""return go_back('" & strCase & "', '" & lngContactID & "', '" & strContactName & "', '" & strLName & "', '" & strFName & "', '" & strMName & "', '" & strWorkFor & "', '" & strPosition & "', '" & strWP & "', '" & strWExt & "', '" & strCP & "', '" & strFP & "', '" & strPP & "', '" & strEmail2 & "', '" & strBuilding & "', '" & strStreet & "', '" & strCPC & "', '" & strPC & "', '" & strResponsibility & "', '" & bolPIN & "')"">" &strWP      & "</a>&nbsp;</td>"&vbCrLf
	 		Response.Write "<td nowrap><a href=""#"" onClick=""return go_back('" & strCase & "', '" & lngContactID & "', '" & strContactName & "', '" & strLName & "', '" & strFName & "', '" & strMName & "', '" & strWorkFor & "', '" & strPosition & "', '" & strWP & "', '" & strWExt & "', '" & strCP & "', '" & strFP & "', '" & strPP & "', '" & strEmail2 & "', '" & strBuilding & "', '" & strStreet & "', '" & strCPC & "', '" & strPC & "', '" & strResponsibility & "', '" & bolPIN & "')"">" &aList(7,k) & "</a>&nbsp;</td>"&vbCrLf
	 		Response.Write "<td nowrap><a href=""#"" onClick=""return go_back('" & strCase & "', '" & lngContactID & "', '" & strContactName & "', '" & strLName & "', '" & strFName & "', '" & strMName & "', '" & strWorkFor & "', '" & strPosition & "', '" & strWP & "', '" & strWExt & "', '" & strCP & "', '" & strFP & "', '" & strPP & "', '" & strEmail2 & "', '" & strBuilding & "', '" & strStreet & "', '" & strCPC & "', '" & strPC & "', '" & strResponsibility & "', '" & bolPIN & "')"">" &aList(14,k)& "</a>&nbsp;</td>"&vbCrLf
	 		Response.Write "<td nowrap><a href=""#"" onClick=""return go_back('" & strCase & "', '" & lngContactID & "', '" & strContactName & "', '" & strLName & "', '" & strFName & "', '" & strMName & "', '" & strWorkFor & "', '" & strPosition & "', '" & strWP & "', '" & strWExt & "', '" & strCP & "', '" & strFP & "', '" & strPP & "', '" & strEmail2 & "', '" & strBuilding & "', '" & strStreet & "', '" & strCPC & "', '" & strPC & "', '" & strResponsibility & "', '" & bolPIN & "')"">" &aList(15,k)& "</a>&nbsp;</td>"&vbCrLf
	 		Response.Write "<td nowrap><a href=""#"" onClick=""return go_back('" & strCase & "', '" & lngContactID & "', '" & strContactName & "', '" & strLName & "', '" & strFName & "', '" & strMName & "', '" & strWorkFor & "', '" & strPosition & "', '" & strWP & "', '" & strWExt & "', '" & strCP & "', '" & strFP & "', '" & strPP & "', '" & strEmail2 & "', '" & strBuilding & "', '" & strStreet & "', '" & strCPC & "', '" & strPC & "', '" & strResponsibility & "', '" & bolPIN & "')"">" &aList(16,k)& "</a>&nbsp;</td>"&vbCrLf
	 		Response.Write "<td nowrap><a href=""#"" onClick=""return go_back('" & strCase & "', '" & lngContactID & "', '" & strContactName & "', '" & strLName & "', '" & strFName & "', '" & strMName & "', '" & strWorkFor & "', '" & strPosition & "', '" & strWP & "', '" & strWExt & "', '" & strCP & "', '" & strFP & "', '" & strPP & "', '" & strEmail2 & "', '" & strBuilding & "', '" & strStreet & "', '" & strCPC & "', '" & strPC & "', '" & strResponsibility & "', '" & bolPIN & "')"">" &aList(17,k)& "</a>&nbsp;</td>"&vbCrLf
	 		Response.Write "<td nowrap><a href=""#"" onClick=""return go_back('" & strCase & "', '" & lngContactID & "', '" & strContactName & "', '" & strLName & "', '" & strFName & "', '" & strMName & "', '" & strWorkFor & "', '" & strPosition & "', '" & strWP & "', '" & strWExt & "', '" & strCP & "', '" & strFP & "', '" & strPP & "', '" & strEmail2 & "', '" & strBuilding & "', '" & strStreet & "', '" & strCPC & "', '" & strPC & "', '" & strResponsibility & "', '" & bolPIN & "')"">" &aList(21,k)& "</a>&nbsp;</td>"&vbCrLf
	 		Response.Write "<td nowrap><a href=""#"" onClick=""return go_back('" & strCase & "', '" & lngContactID & "', '" & strContactName & "', '" & strLName & "', '" & strFName & "', '" & strMName & "', '" & strWorkFor & "', '" & strPosition & "', '" & strWP & "', '" & strWExt & "', '" & strCP & "', '" & strFP & "', '" & strPP & "', '" & strEmail2 & "', '" & strBuilding & "', '" & strStreet & "', '" & strCPC & "', '" & strPC & "', '" & strResponsibility & "', '" & bolPIN & "')"">" &aList(22,k)& "</a>&nbsp;</td>"&vbCrLf
	 		Response.Write "</tr>"
	 	else
            Response.Write "<td nowrap><a target=""_parent"" href=""ContactDetail.asp?ContactID="&aList(0,k)&""">"&aList(8,k)&"</a>&nbsp;</td>"&vbCrLf
	 		Response.Write "<td nowrap><a target=""_parent"" href=""ContactDetail.asp?ContactID="&aList(0,k)&""">"&aList(0,k)&"</a>&nbsp;</td>"&vbCrLf
	 		Response.Write "<td nowrap><a target=""_parent"" href=""ContactDetail.asp?ContactID="&aList(0,k)&""">"&aList(1,k)&"</a>&nbsp;</td>"&vbCrLf
	 		Response.Write "<td nowrap><a target=""_parent"" href=""ContactDetail.asp?ContactID="&aList(0,k)&""">"&aList(2,k)&"</a>&nbsp;</td>"&vbCrLf
	 		Response.Write "<td nowrap><a target=""_parent"" href=""ContactDetail.asp?ContactID="&aList(0,k)&""">"&aList(20,k)&"</a>&nbsp;</td>"&vbCrLf
	 		Response.Write "<td nowrap><a target=""_parent"" href=""ContactDetail.asp?ContactID="&aList(0,k)&""">"&aList(9,k)&"</a>&nbsp;</td>"&vbCrLf
	 		Response.Write "<td nowrap><a target=""_parent"" href=""ContactDetail.asp?ContactID="&aList(0,k)&""">"&aList(13,k)&"</a>&nbsp;</td>"&vbCrLf
	 		Response.Write "<td nowrap><a target=""_parent"" href=""ContactDetail.asp?ContactID="&aList(0,k)&""">"&strWP&"</a>&nbsp;</td>"&vbCrLf
	 		Response.Write "<td nowrap><a target=""_parent"" href=""ContactDetail.asp?ContactID="&aList(0,k)&""">"&aList(7,k)&"</a>&nbsp;</td>"&vbCrLf
	 		Response.Write "<td nowrap><a target=""_parent"" href=""ContactDetail.asp?ContactID="&aList(0,k)&""">"&aList(14,k)&"</a>&nbsp;</td>"&vbCrLf
	 		Response.Write "<td nowrap><a target=""_parent"" href=""ContactDetail.asp?ContactID="&aList(0,k)&""">"&aList(15,k)&"</a>&nbsp;</td>"&vbCrLf
	 		Response.Write "<td nowrap><a target=""_parent"" href=""ContactDetail.asp?ContactID="&aList(0,k)&""">"&aList(16,k)&"</a>&nbsp;</td>"&vbCrLf
	 		Response.Write "<td nowrap><a target=""_parent"" href=""ContactDetail.asp?ContactID="&aList(0,k)&""">"&aList(17,k)&"</a>&nbsp;</td>"&vbCrLf
	 		Response.Write "<td nowrap><a target=""_parent"" href=""ContactDetail.asp?ContactID="&aList(0,k)&""">"&aList(21,k)&"</a>&nbsp;</td>"&vbCrLf
	 		Response.Write "<td nowrap><a target=""_parent"" href=""ContactDetail.asp?ContactID="&aList(0,k)&""">"&aList(22,k)&"</a>&nbsp;</td>"&vbCrLf
	 		Response.Write "</tr>"
	 	end if
	next
	 %>
	</TBODY>
	 <TFOOT>
		<TR>
			<TD align=left colSpan=15>
				<input type=hidden name=txtPageNumber value=<%=intPageNumber%>>
				<input type="submit" name="action" value="&lt;&lt;">
				<input type="submit" name="action" value="&lt;">
				<input type="text" name="txtGoToPageNo" onClick="document.frmContactList.txtGoToPageNo.value=''" title="You can jump to a specific page by typing the page number in this box" value="page <%=intPageNumber%> of <%=intPageCount%>" style="HEIGHT: 22px; WIDTH: 150px">
				<input type="submit" name="action" value="&gt;">
				<input type="submit" name="action" value="&gt;&gt;">&nbsp;&nbsp;
				<img SRC="images/excel.gif" onclick="document.frmContactList.target='new';document.frmContactList.hdnExport.value='xls';frmContactList.txtPageNumber.value='';document.frmContactList.submit();document.frmContactList.hdnExport.value='';document.frmContactList.target='_self';" WIDTH="32" HEIGHT="32">
			</TD>
		</TR>
	</TFOOT>
	<CAPTION>Records <%=m+1%> to <%=n+1%> of <%=UBound(aList, 2)+1 & " records"%></CAPTION>
</TABLE>
</FORM>
</BODY>
</HTML>
