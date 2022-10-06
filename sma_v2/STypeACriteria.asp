<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<!--#include file = "smaConstants.inc"-->
<!--#include file = "smaProcs.inc"-->
<!--#include file = "databaseconnect.asp" -->
<!--
******************************************************************************
*
*
* In Param:		This pages reads following cookies
*					STypeAttributeDescription
*
*
*******************************************************************************
-->
<%
Dim strWinName, strBusinessID, strServiceCategoryID, strSTypeID, strSTypeDesc, strServiceLevelID, lIndex, strLANG
Dim lRow, arrLOBList, pfRow, arrSCategoryList, arrSCategoryListEN, arrSTAttList, arrRevenueList, arrSTAttvList
Dim objRS, strSQL, strWhereClause
Dim intAccessLevel

	intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
	If (intAccessLevel And intConst_Access_ReadOnly) <> intConst_Access_ReadOnly Then
		DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to Service Category. Please contact your system administrator"
	End If

	strWinName = Request.Cookies("WinName")
	strBusinessID = Request.Cookies("BusinessID")
	strServiceCategoryID = Request.Cookies("ServiceCategoryID")
	strServiceLevelID = Request.Cookies("ServiceLevelID")
	strSTypeID = Request.Cookies("ServiceType")
	strSTypeDesc = Request.Cookies("STypeDesc")

'TQ_INOSS
	'strLANG = Request.Cookies("UserInformation")("language_preference")
	'if (Len(strLANG) = 0) then strLANG = "EN"
	strLANG = "EN"

	strSQL =" SELECT LOB_ID, LOB_CODE, LOB_DESC " &_
			" FROM CRP.V_LOB " &_
			" WHERE lob_id NOT IN " &_
		    " (SELECT lob_id " &_
		    " FROM crp.v_lob " &_
		    " WHERE language_preference_lcode = '" & strLANG & "' ) " &_
			" AND LANGUAGE_PREFERENCE_LCODE = 'EN' " &_
			" AND RECORD_STATUS_IND = 'A'" &_
			" UNION SELECT LOB_ID, LOB_CODE, LOB_DESC " &_
			" FROM crp.v_lob " &_
			" WHERE language_preference_lcode = '" & strLANG & "' " &_
			" AND RECORD_STATUS_IND = 'A' " &_
			" ORDER BY LOB_DESC "

	'Create Recordset object
	Set objRS = Server.CreateObject("ADODB.Recordset")
	objRS.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	arrLOBList = objRS.GetRows

	'Get SERVICE_CATEGORY
	'strLANG = Request.Cookies("UserInformation")("language_preference")
	'if (Len(strLANG) = 0) then strLANG = "EN"

	strSQL = "SELECT SERVICE_CATEGORY_ID, LOB_ID, SERVICE_CATEGORY_DESC, LANGUAGE_PREFERENCE_LCODE " &_
			"FROM CRP.V_SERVICE_CATEGORY "&_
			"WHERE SERVICE_CATEGORY_ID NOT IN" &_
			      	"(SELECT SERVICE_CATEGORY_ID " &_
		      		"FROM CRP.V_SERVICE_CATEGORY " &_
			      	"WHERE language_preference_lcode = '" & strLANG & "' ) " &_
			"AND LANGUAGE_PREFERENCE_LCODE = 'EN'" &_
			"AND RECORD_STATUS_IND = 'A'" &_
			"UNION SELECT SERVICE_CATEGORY_ID, LOB_ID, SERVICE_CATEGORY_DESC, LANGUAGE_PREFERENCE_LCODE " &_
			"FROM CRP.V_SERVICE_CATEGORY " &_
			"WHERE language_preference_lcode = '" & strLANG & "'" &_
			"AND RECORD_STATUS_IND = 'A' " &_
			"ORDER BY SERVICE_CATEGORY_DESC"

	'Create Recordset object
	Set objRS = Server.CreateObject("ADODB.Recordset")
	objRS.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	arrSCategoryList = objRS.GetRows

	' Get English versions
	strSQL = " SELECT service_category_id, service_category_desc " &_
			 " FROM crp.service_category " &_
			 " WHERE record_status_ind = 'A' "

	Set objRS = Server.CreateObject("ADODB.Recordset")
	objRS.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	arrSCategoryListEN = objRS.GetRows

'	Response.Write (arrSCategoryList(0,0) & " " & arrSCategoryListEN(0,0))
'	Response.End

	for lRow = 0 to UBound(arrSCategoryList, 2)
		if (strcomp(arrSCategoryList(3,lRow), "EN") <> 0 ) then
			' For non-English entries only, append the English
			for pfRow = 0 to UBound(arrSCategoryListEN, 2)
				if (strcomp(arrSCategoryList(0,lRow), arrSCategoryListEN(0,pfRow)) = 0) then
					if (strcomp(arrSCategoryList(2,lRow), arrSCategoryListEN(1,pfRow)) <> 0) then
						' Append English version to the end of any translated service categories
						arrSCategoryList(2,lRow) = arrSCategoryList(2,lRow) & " (" & arrSCategoryListEN(1,pfRow) & ")"
						pfRow = UBound(arrSCategoryListEN, 2) + 1 ' Done: Skip out of the loop early
					end if
				end if
			next
		end if
	next

	'Get the Service Level Agreements
	'StrSql = "SELECT  b.srvc_type_att_name ||'\'|| c.srvc_type_att_val_name, a.srvc_type_att_val_usage_id " &_
	'		 "FROM crp.srvc_type_att_val_usage a,  "&_
	'		 "crp.srvc_type_att b, crp.srvc_type_att_val c " &_
	'		 "WHERE b.srvc_type_att_id=a.srvc_type_att_id " &_
	'		 "AND c.srvc_type_att_val_id=a.srvc_type_att_val_id " &_
	'		 "AND a.srvc_type_att_val_usage_id IN " &_
	'		 "(SELECT DISTINCT srvc_type_att_val_usage_id FROM crp.srvc_type_att_val_xref)"

 	StrSql = "SELECT srvc_type_att_name, srvc_type_att_id from crp.srvc_type_att where record_status_ind='A' order by upper(srvc_type_att_name)"

	'Create Recordset object
	'Set objRS = Server.CreateObject("ADODB.Recordset")
	'objRS.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	set arrSTAttList = objConn.Execute(strSQL)
	'Response.Write (arrSTAttList(0))
	'Response.End

	StrSql = "SELECT srvc_type_att_val_name, srvc_type_att_val_id from crp.srvc_type_att_val where record_status_ind = 'A' order by upper(srvc_type_att_val_name)"
 	set arrSTAttvList = objConn.Execute(strSQL)




	StrSql = "SELECT REVENUE_REGION_LCODE, " &_
	         "       REVENUE_REGION_DESC " &_
			 "FROM SO.LCODE_REVENUE_REGION " &_
			 "WHERE RECORD_STATUS_IND = 'A' "

	'Create Recordset object
	Set objRS = Server.CreateObject("ADODB.Recordset")
	objRS.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	arrRevenueList = objRS.Getrows
%>
<HTML>
<HEAD>
<META name="Generator" content="Microsoft Visual Studio 6.0">
<META http-equiv="Pragma" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<LINK rel="stylesheet" type="text/css" href="stylesheets/styles.css">
<SCRIPT type="text/javascript" language="javascript" src="AccessLevels.js"></SCRIPT>
<SCRIPT type="text/javascript" language="javascript" src="GeneralJavaFunctions.js"></SCRIPT>
<SCRIPT type="text/javascript" language="javascript" id="clientEventHandlersJS">
<!--
var intAccessLevel = <%=intAccessLevel%>;
var strWinName = "<%=strWinName%>";

var arrLOBList = new Array();
var arrServiceCategoryList = new Array();

//set section title
setPageTitle("SMA - Service Type");

function window_onLoad() {
//****************************************************************************************
//
//
//
//
//****************************************************************************************
var intCounter;

	arrLOBList[0] = "";
	arrServiceCategoryList[0] = "";

	for (intCounter = 1; intCounter < document.frmSTypeSearch.selLOB.options.length; intCounter++) {
		var oOption = document.frmSTypeSearch.selLOB.options(intCounter);
		//Each array element holds LOB_ID|LOB_DESC
		arrLOBList[intCounter] = (oOption.value + "|" + oOption.text);
	}

	for (intCounter = 1; intCounter < document.frmSTypeSearch.selSCategory.options.length; intCounter++) {
		var oOption = document.frmSTypeSearch.selSCategory.options(intCounter);
		//Each array element holds SERVICE_CATEGORY_ID|LOB_ID|SERVICE_CATEGORY_DESC
		arrServiceCategoryList[intCounter] = (oOption.value + "|" + oOption.text);
		var strValue = oOption.value;
		var arrValue = strValue.split("|");
		oOption.value = arrValue[0];
	}

	if (document.frmSTypeSearch.selLOB.selectedIndex != 0
			|| document.frmSTypeSearch.selSCategory.selectedIndex != 0
			|| document.frmSTypeSearch.selSTTID.selectedIndex != 0
			|| document.frmSTypeSearch.hdnSTypeID.value != ""
			|| document.frmSTypeSearch.txtSTypeDescription.value != "") {
		document.frmSTypeSearch.hdnSTypeID.value = "";
		DeleteCookie("BusinessID");
		DeleteCookie("ServiceCategoryID");
		DeleteCookie("ServiceLevelID");
		DeleteCookie("ServiceType");
		DeleteCookie("STypeDesc");
		DeleteCookie("WinName");
		document.frmSTypeSearch.submit();
	}
}

function fct_onChangeLOB() {
var intCounter = 1;
var strBusinessID;

	if (document.frmSTypeSearch.selLOB.selectedIndex != 0) {
		strBusinessID = document.frmSTypeSearch.selLOB.value;

		//Remove all the OPTION tags from the Service Category
		for (intCounter = document.frmSTypeSearch.selSCategory.length - 1; intCounter > 0; intCounter--) {
			document.frmSTypeSearch.selSCategory.options.remove(intCounter);
		}
		//Add Service Categories that belong to the selected Line of Business
		for (intCounter = 1; intCounter < arrServiceCategoryList.length; intCounter++) {
			var strValue = arrServiceCategoryList[intCounter];
			var arrValue = strValue.split("|");
			if (arrValue[1] == strBusinessID) {
				var strElement = "<OPTION VALUE='" + arrValue[0] + "'>" + arrValue[2] + "</OPTION>";
				var oOption = document.createElement(strElement);
				document.frmSTypeSearch.selSCategory.options.add(oOption);
				oOption.innerText = arrValue[2];	//SERVICE_CATEGORY_DESC
//				oOption.Value = arrValue[1];		//LOB_ID
//				oOption.Value = arrValue[0];		//SERVICE_CATEGORY_ID
			}
		}
	}
	else {
		//Remove all the OPTION tags from the Service Category
		for (intCounter = document.frmSTypeSearch.selSCategory.length - 1; intCounter > 0; intCounter--) {
			document.frmSTypeSearch.selSCategory.options.remove(intCounter);
		}
		//Add all the Service Categories
		for (intCounter = 1; intCounter < arrServiceCategoryList.length; intCounter++) {
			var strValue = arrServiceCategoryList[intCounter];
			var arrValue = strValue.split("|");
			var strElement = "<OPTION VALUE='" + arrValue[0] + "'>" + arrValue[2] + "</OPTION>";
			var oOption = document.createElement(strElement);
			document.frmSTypeSearch.selSCategory.options.add(oOption);
			oOption.innerText = arrValue[2];	//SERVICE_CATEGORY_DESC
//			oOption.Value = arrValue[1];		//LOB_ID
//			oOption.Value = arrValue[0];		//SERVICE_CATEGORY_ID
		}
	}
}

function btnCalendar_onClick(intDateFieldNo) {
var NewWin;

	SetCookie("Field", intDateFieldNo);
	NewWin=window.open("TheCalendar.asp","NewWin","toolbar=no,status=no,width=260,height=225,menubar=no,resize=no");
	//NewWin.creator=self;
	NewWin.focus();
}


function fct_onChange() {
	document.frmSTypeSearch.hdnSTypeID.value = "";
}

function btnClear_onClick() {
	with (document.frmSTypeSearch) {
		selLOB.selectedIndex = 0;
		selSCategory.selectedIndex = 0;
		selSTTID.selectedIndex = 0;
		selSTTvID.selectedIndex = 0;
		hdnSTypeID.value = "";
		txtSTypeDescription.value = "";
		chkActiveOnly.checked = true;
	}
}

function btnNew_onClick() {
	parent.document.location.href ='SAttMRDetail.asp';
}



function fct_setDays(iIndex) {
var intDays = 31;
var strMonth = document.frmSTypeSearch.item("selmonth", iIndex).options[document.frmSTypeSearch.item("selmonth", iIndex).selectedIndex].value;
var strYear = document.frmSTypeSearch.item("selyear", iIndex).options[document.frmSTypeSearch.item("selyear", iIndex).selectedIndex].value;
var intCurrentDay = document.frmSTypeSearch.item("selday", iIndex).options[document.frmSTypeSearch.item("selday", iIndex).selectedIndex].value;
var intCounter = document.frmSTypeSearch.item("selday", iIndex).options.length;

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
			document.frmSTypeSearch.item("selday", iIndex).options[intCounter++] = oOption;
		}
	}
	else {
		while (intCounter > intDays) {
			document.frmSTypeSearch.item("selday", iIndex).options[intCounter--] = null;
		}
	}
	if (intCurrentDay > intDays) {
		document.frmSTypeSearch.item("selday", iIndex).selectedIndex = intDays;
	}
}
//-->
</SCRIPT>
</HEAD>
<BODY language="javascript" onLoad="window_onLoad();">
<FORM id="frmSTypeSearch" name="frmSTypeSearch" method="post" action="STypeAList.asp" target="fraResult">
	<INPUT type="hidden" id="hdnWinName" name="hdnWinName" value="<%=strWinName%>">
<TABLE cols="4" width=100%>
<THEAD>
	<TH colspan="4" align="left">Service Type Attribute Search</TH>
</THEAD>
<TBODY>
	<TR>
		<TD align="right" width="19%">Line of Business</TD>
		<TD align="left" width="42%"><SELECT id="selLOB" name="selLOB" style="width: 350px" onChange="fct_onChangeLOB();">
			<OPTION></OPTION>
			<%For lRow = LBound(arrLOBList, 2) To UBound(arrLOBList, 2)
				If StrComp(strBusinessID, arrLOBList(0, lRow), 0) = 0 Then%>
				<OPTION selected value="<%=arrLOBList(0, lRow)%>"><%=arrLOBList(1, lRow) & " - " & arrLOBList(2, lRow)%></OPTION>
				<%Else%>
				<OPTION value="<%=arrLOBList(0, lRow)%>"><%=arrLOBList(1, lRow) & " - " & arrLOBList(2, lRow)%></OPTION>
			<%	End If
			Next%>
			</SELECT>
		</TD>
	    <TD align="right" nowrap width="33%">&nbsp;</TD>
	    <TD align="left" width="28%">
	    &nbsp;&nbsp; </TD>
	</TR>
	<TR>
		<TD align="right" width="19%">Service Category</TD>
		<TD align="left" width="42%"><SELECT id="selSCategory" name="selSCategory" style="width: 350px">
			<OPTION></OPTION>
			<%For lRow = LBound(arrSCategoryList, 2) To UBound(arrSCategoryList, 2)
				If StrComp(strServiceCategoryID, arrSCategoryList(0, lRow), 0) = 0 Then%>
				<OPTION selected value="<%=arrSCategoryList(0, lRow) & "|" & arrSCategoryList(1, lRow)%>"><%=arrSCategoryList(2, lRow)%></OPTION>
				<%Else%>
				<OPTION value="<%=arrSCategoryList(0, lRow) & "|" & arrSCategoryList(1, lRow)%>"><%=arrSCategoryList(2, lRow)%></OPTION>
			<%	End If
			Next%>
			</SELECT>
		</TD>
		<TD align="right" width="33%">&nbsp;</TD>
		<TD align="left" width="28%">&nbsp;</TD>
	</TR>
	<TR>
		<TD align="right" width="19%">Service Type</TD>
		<TD align="left" width="42%">
			<INPUT type="hidden" id="hdnSTypeID" name="hdnSTypeID" value="<%=strSTypeID%>">
			<INPUT type="text" id="txtSTypeDescription" name="txtSTypeDescription" value="<%=strSTypeDesc%>" maxlength="80" style="width: 350px" onKeyPress="fct_onChange();">
		</TD>
		<TD width="33%">&nbsp;</TD>
		<TD align="left" width="28%">&nbsp;</TD>
	</TR>
	<TR>
		<TD align="right" width="19%">Service Type Attribute </TD>
		<TD align="left" width="42%"><SELECT id="selSTTID" name="selSTTID" style="width: 600px">
			<OPTION></OPTION>
						<%Do While Not arrSTAttList.EOF %>
				<OPTION value=<% =arrSTAttList(1)%>><% =arrSTAttList(0)%></OPTION>
			<%	arrSTAttList.MoveNext
			Loop%>
		</TD>
	</TR>
	<TR>
		<TD align="right" width="19%">Service Type Attribute Value</TD>
		<TD align="left" width="42%">
		<SELECT id="selSTTvID" name="selSTTvID" style="width: 600; height:22">
			<OPTION></OPTION>
						<%Do While Not arrSTAttvList.EOF %>
				<OPTION value=<% =arrSTAttvList(1)%>><% =arrSTAttvList(0)%></OPTION>
			<%	arrSTAttvList.MoveNext
			Loop%>
		</TD>
	</TR>

	<tr>
		<TD align="right" width="19%">Active Only</TD>
		<td>
		<INPUT id=chkActiveOnly name=chkActiveOnly1 tabindex=12 type=checkbox value=YES CHECKED style="HEIGHT: 24; WIDTH: 25"></TD>
		<td width="33%">
		</td>
	</tr>
</TBODY>
<TFOOT>
	<TR>
		<TD colspan="4" align="right">
			<%If UCase(strWinName) <> UCase("Popup") Then%>
				&nbsp;
			<%End If%>
			<INPUT id="btnNew"  name="btnNew"  type="button" value="STA Maintenance"  style="width: 3cm" language="javascript" onClick="btnNew_onClick()">
			<INPUT id="btnClear"  name="btnClear"  type="button" value="Clear"  style="width: 2cm" language="javascript" onClick="btnClear_onClick()">&nbsp;
			<INPUT id="btnSearch" name="btnSearch" type="submit" value="Search" style="width: 2cm" language="javascript">&nbsp;
		</TD>
	</TR>
</TFOOT>
</TABLE>
</FORM>
</BODY>
<%
	'Clean ADO Objects
	objRS.Close
	Set objRS = Nothing
	objConn.Close
	Set objConn = Nothing
%>
</HTML>
