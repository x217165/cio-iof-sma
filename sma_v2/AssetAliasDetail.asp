<%@ Language=VBScript %>
<% Option Explicit
'on error resume next
%>
<% Response.Buffer = true %>
<!--#include file="smaConstants.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--#include file="smaProcs.inc"-->
<%

dim intAccessLevel

intAccessLevel = CInt(CheckLogon(strConst_AssetAdditionalCosts))

Dim StrAliasID,StrAssetID,StrSql,strWhereClause,objRS,strNewFacility,strWinMessage,objRsAssetType,strUpdDate

 StrAliasID = Request("AliasID")
 StrAssetID = Request("AssetID")
 strNewFacility = Request("NewFacility")
 strUpdDate = Request("hdnUpdateDateTime")

 dim strRealUserID
 strRealUserID = Session("username")


  if strNewFacility = "NEW" then
  StrAliasID = 0
  strNewFacility =""
 end if


  select case Request("txtFrmAction")
	case "SAVE"
		if (Request("AliasID") <>"") then
		    if ((intAccessLevel and intConst_Access_Update) <> intConst_Access_Update) then
		      DisplayError "BACK", "", 0, "UPDATE DENIED", "You don't have access to update Asset Additional Cost. Please contact your system administrator"
		    end if
		    StrAliasID = Request("AliasID")
			'create command object for update stored proc
			dim cmdUpdateObj
			set cmdUpdateObj = server.CreateObject("ADODB.Command")
			set cmdUpdateObj.ActiveConnection = objConn
			cmdUpdateObj.CommandType = adCmdStoredProc
			cmdUpdateObj.CommandText = "sma_sp_userid.spk_sma_asset_inter.sp_aacost_update"
			'create parameters
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_user_id", adVarChar , adParamInput, 20, strRealUserID)
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_asset_addcost_id",adNumeric , adParamInput,, Clng(Request("AliasID")))
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_asset_id",adNumeric , adParamInput,, Clng(Request("AssetID")))
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_asset_cost_type_code", adVarChar,adParamInput, 10, Request("selcosttype"))
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_dollar_date", adVarChar,adParamInput, 20, Request("hdnDollarDt"))

			if Request("txtamount") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_amount", adVarChar,adParamInput,50, Request("txtamount"))
			else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_amount", adVarChar,adParamInput,50, null)
			end if

			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , Cdate(Request("hdnUpdateDateTime")))

            if Request("txtacomments") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_comments", adVarChar,adParamInput, 2000, Request("txtacomments"))
			else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_comments", adVarChar,adParamInput, 2000, null)
			end if

			'Response.Write "updating..." & StrAliasID

			'Response.Write "<b> count = " & cmdInsertObj.Parameters.count & "<br>"

  			'dim nx
  			 'for nx=0 to cmdUpdateObj.Parameters.count-1
  			  ' Response.Write " parm value= " & cmdUpdateObj.Parameters.Item(nx) & "<br>"
  			 ' next

  			'dim objparm
  		    ' for each objparm in cmdUpdateObj.Parameters
  			  'Response.Write "<b>" & objparm.name & "</b>"
  			 ' Response.Write " has size:  " & objparm.Size & " "
  			  'Response.Write " and value:  " & objparm.value & " "
  			  'Response.Write " and datatype:  " & objparm.Type & "<br> "
  		     'next

  	        'if objConn.Errors.Count <> 0 then
			   'DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT UPDATE FACILITY/PVC ALIAS - PARAMETER ERROR", objConn.Errors(0).Description
			   'objConn.Errors.Clear
		    'end if

			cmdUpdateObj.Execute


			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT UPDATE OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			end if

			strWinMessage = "Record saved successfully. You can now see the changes you made."
		else
		   '(Request("AliasID")="" ) and
		   
		    if ((intAccessLevel and intConst_Access_Create) <> intConst_Access_Create) then
		     DisplayError "BACK", "", 0, "INSERT DENIED", "You don't have access to create Asset Additional Cost. Please contact your system administrator"
		    end if
			'create a new record

			dim cmdInsertObj
			set cmdInsertObj = server.CreateObject("ADODB.Command")
			set cmdInsertObj.ActiveConnection = objConn
			cmdInsertObj.CommandType = adCmdStoredProc
			cmdInsertObj.CommandText = "sma_sp_userid.spk_sma_asset_inter.sp_aacost_insert"

			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_user_id", adVarChar , adParamInput, 20, strRealUserID)
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_asset_addcost_id",adNumeric , adParamOutput,, NULL)
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_asset_id",adNumeric , adParamInput,, Clng(Request("AssetID")))
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_asset_cost_type_code", adVarChar,adParamInput, 10, Request("selcosttype"))
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_dollar_date", adVarChar,adParamInput, 20, Request("hdnDollarDt"))

			if Request("txtamount") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_amount",  adVarChar,adParamInput,50, Request("txtamount"))
			else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_amount",  adVarChar,adParamInput,50, null)
			end if


            if Request("txtacomments") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_comments", adVarChar,adParamInput, 2000, Request("txtacomments"))
			else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_comments", adVarChar,adParamInput, 2000, null)
			end if


		    'if objConn.Errors.Count <> 0 then
			   'DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT CREATE ASSET ADDITIONAL COST - PARAMETER ERROR", objConn.Errors(0).Description
			   'objConn.Errors.Clear
		    'end if
	
			'dim objparm
	  		 ' for each objparm in cmdInsertObj.Parameters
	  		'	  Response.Write "<b>" & objparm.name & "</b>"
	  		'	  Response.Write " has size:  " & objparm.Size & " "
	  		'	  Response.Write " and value:  " & objparm.value & " "
	  		'	  Response.Write " and datatype:  " & objparm.Type & "<br> "
	  		' next
	
			 ' dim nx
	  		'	 for nx=0 to cmdInsertObj.Parameters.count-1
	  		'	   Response.Write " parm value= " & cmdInsertObj.Parameters.Item(nx) & "<br>"
	  		'	  next
'


			cmdInsertObj.Execute

			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT CREATE ASSET ADDITIONAL COST ", objConn.Errors(0).Description
				objConn.Errors.Clear
			else
				StrAliasID =  cmdInsertObj.Parameters("p_asset_addcost_id").Value
			end if
			strWinMessage = "Record created successfully. You can now see the new record."

		end if

	 	if err then
		  DisplayError "BACK", "", err.Number, "CANNOT CREATE ASSET ADDITIONAL COST - TRY AGAIN", err.Description
	 	end if

	case "DELETE"

		'delete record
		if ((intAccessLevel and intConst_Access_Delete) <> intConst_Access_Delete) then
		  DisplayError "BACK", "", 0, "DELETE DENIED", "You don't have access to delete Additional Cost. Please contact your system administrator"
		end if

			dim cmdDeleteObj
			set cmdDeleteObj = server.CreateObject("ADODB.Command")
			set cmdDeleteObj.ActiveConnection = objConn
			'Response.Write Request("hdnUpdateDateTime")
			cmdDeleteObj.CommandType = adCmdStoredProc
			cmdDeleteObj.CommandText = "sma_sp_userid.spk_sma_asset_inter.sp_aacost_delete"
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_asset_addcost_id", adNumeric, adParamInput,,Clng(StrAliasID))
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , CDate(Request("hdnUpdateDateTime")))

			'call the insert stored proc
  			'cmdDeleteObj.Parameters.Refresh

  			'dim objparm
  		   'for each objparm in cmdDeleteObj.Parameters
  			 ' Response.Write "<b>" & objparm.name & "</b>"
  			  'Response.Write " has size:  " & objparm.Size & " "
  			  'Response.Write " and value:  " & objparm.value & " "
  			  'Response.Write " and datatype:  " & objparm.Type & "<br> "
  		  'next

  			'Response.Write "<b> count = " & cmdInsertObj.Parameters.count & "<br>"
  			'dim nx
  			 'for nx=0 to cmdDeleteObj.Parameters.count-1
  			   'Response.Write " parm value= " & cmdDeleteObj.Parameters.Item(nx) & "<br>"
  			' next

			cmdDeleteObj.Execute
			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE ASSET ADDITIONAL COST", objConn.Errors(0).Description
				objConn.Errors.Clear
			end if
			StrAliasID = 0
			strWinMessage = "Record deleted successfully."
		  'else
		      'DisplayError "BACK", "", err.Number, "ACCESS DENIED. PLEASE CONTACT YOUR SYSTEM ADMINISTRATOR.", err.Description
	      'end if
       end select


 StrSql = "SELECT ASSET_COST_TYPE_CODE,ASSET_COST_TYPE_DESC FROM CRP.ASSET_COST_TYPE WHERE RECORD_STATUS_IND = 'A' ORDER BY ASSET_COST_TYPE_CODE"

 'Create Recordset object
 set objRsAssetType = objConn.Execute(StrSql)

 IF StrAliasID <> 0 THEN
 StrSql ="select "&_
         "ASSET_ADDITIONAL_COST_ID," &_
         "ASSET_ID," &_
         "ASSET_COST_TYPE_CODE," &_
         "TO_CHAR(DOLLAR_DATE,'MON-DD-YYYY') DOLLAR_DATE," &_
         "DOLLAR_AMOUNT,"&_
         "CHANGE_COMMENTS,"&_
         "TO_CHAR(CREATE_DATE_TIME,'MON-DD-YYYY HH:MI:SS') CREATE_DATE_CONV," &_
         "sma_sp_userid.spk_sma_library.sf_get_full_username(CREATE_REAL_USERID) CREATE_REAL_USERID," &_
         "UPDATE_DATE_TIME," &_
         "TO_CHAR(UPDATE_DATE_TIME,'MON-DD-YYYY HH:MI:SS') UPDATE_DATE_CONV," &_
         "RECORD_STATUS_IND," &_
         "sma_sp_userid.spk_sma_library.sf_get_full_username(UPDATE_REAL_USERID) UPDATE_REAL_USERID" &_
         " from crp.ASSET_ADDITIONAL_COST"



      strWhereClause =  "where ASSET_ADDITIONAL_COST_ID =" & StrAliasID

      StrSql =  StrSql & " "& strWhereClause

      set objRS = objConn.Execute(StrSql)

    if err then
	   DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET 32132", err.Description
    end if

  END IF
   'Response.Write "SQL STATEMENT WIH WHERE=" & StrSql & "<p>"

   'Create the command object

%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<LINK REL="stylesheet" TYPE="text/css" HREF="stylesheets/styles.css">
<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
<script type="text/javascript" SRC="AccessLevels.js"></script>
<TITLE>Alias Detail</TITLE>
<SCRIPT LANGUAGE=javascript>
<!--
var bolSaveRequired = false;
var intAccessLevel=<%=intAccessLevel%>;
var intConst_MessageDisplay=<%=intConst_MessageDisplay%>;


function fct_clearStatus() {
	window.status = "";
}

function fct_displayStatus(strMessage){
	window.status = strMessage;
	setTimeout('fct_clearStatus()',intConst_MessageDisplay);
}

function body_onLoad(strWinStatus){
	var strWinStatus='<%=routineJavascriptString(strWinMessage)%>';
	fct_displayStatus(strWinStatus);
}


function fct_onChange(){
if (intAccessLevel >= intConst_Access_Create){
 if (document.frmAddCost.AliasID.value != "")
     {
      bolSaveRequired = true;
     }

    }
}


function btnNew_click(){
if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {alert('Access denied. Please contact your system administrator.'); return;}
	self.document.location.href ="AssetAliasDetail.asp?NewFacility=NEW";
}


function fct_onDelete(){
if (document.frmAddCost.AliasID.value != '') {
 if (((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) || (document.frmAddCost.txtRecordStatusInd.value == "D"))
  {alert('Access denied. Please contact your system administrator.');
   return;}
	if (confirm('Do you really want to delete this object?')){
		document.location = "AssetAliasDetail.asp?txtFrmAction=DELETE&AliasID="+document.frmAddCost.AliasID.value+"&hdnUpdateDateTime="+document.frmAddCost.hdnUpdateDateTime.value;
	}
	} //null alias id
	else {
	fct_displayStatus('Unable to Delete record no Alias ID provided.');;
	return(false);
	}
}



function btnClose_onclick(){
window.close();
parent.opener.iFrame_display();
}

function body_onUnload(){

	opener.document.fmAssetDetail.btn_iFrameRefresh.click();
}

function body_onBeforeUnload(){
    document.frmAddCost.btnSave.focus();
	if (bolSaveRequired) {
		if	(((intAccessLevel & intConst_Access_Create == intConst_Access_Create) && (document.frmAddCost.AliasID.value == "")) || ((intAccessLevel & intConst_Access_Update == intConst_Access_Update) && (document.frmAddCost.AliasID.value != ""))) {
			event.returnValue = "There is unsaved data on the screen. To save changes, click CANCEL below then click SAVE on the main form.";
		}
	}
}

//-->
</SCRIPT>


<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function frmAddCost_onsubmit() {

document.frmAddCost.txtFrmAction.value = "SAVE";
 bolSaveRequired = false;
//Dollar Date
if	((((intAccessLevel & intConst_Access_Create) == intConst_Access_Create) && (document.frmAddCost.AliasID.value == "")) || ((intAccessLevel & intConst_Access_Update) == intConst_Access_Update) && (document.frmAddCost.AliasID.value != ""))


	if (isWhitespace(document.frmAddCost.selcosttype.item(document.frmAddCost.selcosttype.selectedIndex).value)) {
		alert('Please enter cost type code ');
		document.frmAddCost.selcosttype.focus();
		return(false);
	}


	if (document.frmAddCost.selmonth.item(document.frmAddCost.selmonth.selectedIndex).value != "")
	{
	strMonth = document.frmAddCost.selmonth.item(document.frmAddCost.selmonth.selectedIndex).value;
	strDay = document.frmAddCost.selday.item(document.frmAddCost.selday.selectedIndex).value;
	strYear = document.frmAddCost.selyear.item(document.frmAddCost.selyear.selectedIndex).value;

	strDate = strMonth + "/" + strDay + "/" + strYear;
	document.frmAddCost.hdnDollarDt.value = strDate;
	}
	else
	{
	//strDate = "";
	//document.frmAddCost.hdnDollarDt.value = strDate;
	alert('Please enter dollar date ');
	document.frmAddCost.selmonth.focus();
	return(false);
	}

	if (isWhitespace(document.frmAddCost.selmonth.item(document.frmAddCost.selmonth.selectedIndex).value)) {
		alert('Please enter Dollar Date ');
		document.frmAddCost.selmonth.focus();
		return(false);
	}

	if (isWhitespace(document.frmAddCost.txtamount.value)) {
		alert('Please enter Dollar Amount ');
		document.frmAddCost.txtamount.focus();
		return(false);
	}


	if (isNaN(document.frmAddCost.txtamount.value))
	  {
	   alert('Please enter a valid Dollar Amount');
	   document.frmAddCost.txtamount.focus();
	   return(false);
	  }

	  return(true);
}


function btnCalendar_onclick(intDateFieldNo) {
	var NewWin2;
	    SetCookie("Field",intDateFieldNo);
		NewWin2=window.open("calendar.asp","NewWin2","toolbar=no,status=no,width=260,height=225,menubar=no resize=no");
		//NewWin.creator=self;
	NewWin2.focus();
}

function fct_onReset(){
	bolSaveRequired = false;
	document.location = "AssetAliasDetail.asp?AliasID=<%=StrAliasID%>";
}

function round_value()
{
if (!(isNaN(document.frmAddCost.txtamount.value)))
{
document.frmAddCost.txtamount.value =Math.round(document.frmAddCost.txtamount.value*100)/100;
}
else
 {
 alert("Enter a Valid Amount!");
 document.frmAddCost.txtamount.value ="";
 document.frmAddCost.txtamount.focus();


 }
}

function btnSave_onclick()
{
var bolretval
bolretval=frmAddCost_onsubmit();
if(bolretval)
document.frmAddCost.submit();
}
//-->
</SCRIPT>
</HEAD>
<BODY onLoad="body_onLoad();" onUnload="body_onUnload();" onBeforeUnload="body_onBeforeUnload();">

<FORM name=frmAddCost LANGUAGE=javascript onsubmit="">
<INPUT type="hidden" name=txtFrmAction value="">
<INPUT type="hidden" name=hdnDollarDt value="">

<INPUT name=hdnUpdateDateTime type=hidden style="HEIGHT: 20px; WIDTH: 100px" value=<%if  StrAliasID <> 0 then  Response.Write """"&objRS("UPDATE_DATE_TIME")&"""" else Response.Write """""" end if%>  >&nbsp;&nbsp;
<INPUT  name=AliasID type=hidden style="HEIGHT: 21px; WIDTH: 200px" value= <%if StrAliasID <> 0 then  Response.Write """"&objRS("ASSET_ADDITIONAL_COST_ID")&"""" else Response.Write """""" end if%> >
<INPUT  type=hidden name=AssetID   style="HEIGHT: 21px; WIDTH: 200px" value= <%if StrAliasID <> 0 then  Response.Write """"&objRS("ASSET_ID")&"""" else Response.Write StrAssetID  end if%> onchange ="fct_onChange();">
<TABLE border=0 width=100%>
<thead>
	<TR ><TD colspan=2>Asset Additional Cost Detail</td></tr>
</thead>
<tbody>
<TR>
	<TD ALIGN=RIGHT width=20% NOWRAP>Cost Type<font color=red>*</font></TD>
	<TD width=80%>
		<SELECT  name=selcosttype style="HEIGHT: 20px; WIDTH: 120px" onchange ="fct_onChange();">
		<OPTION></OPTION>
		<%Do while Not objRsAssetType.EOF
		 Response.write "<OPTION "
		 if StrAliasID <> 0 then
		    if objRsAssetType("ASSET_COST_TYPE_CODE") = objRs("ASSET_COST_TYPE_CODE") then
				   Response.Write " selected "
			end if
		 end if
		  Response.Write " VALUE ="& objRsAssetType("ASSET_COST_TYPE_CODE") & ">" & routineHtmlString(objRsAssetType("ASSET_COST_TYPE_DESC")) & "</OPTION>"
		  objRsAssetType.MoveNext
		 Loop
		%>
		</SELECT>
	</TD>
</TR>
<TR>
<TD ALIGN=RIGHT NOWRAP>Dollar Date<font color=red>*</font></TD>
<TD width=25%><SELECT name=selmonth style="HEIGHT: 20px; WIDTH: 70px" onchange ="fct_onChange();">
 <OPTION></OPTION>
 <%
 dim k

 for k = 1 to 12
   Response.Write "<option "
 if StrAliasID <> 0 then
  if k = month(objRS("DOLLAR_DATE")) then
    Response.Write " selected "
  end if
 end if
  if k < 10 then
  k="0"&k
  end if
  Response.write " VALUE ="& k & ">" &ucase(monthName(k,true)) & "</OPTION>"
  next
 %>
 </SELECT>

 <SELECT  name=selday style="HEIGHT: 20px; WIDTH: 60px" onchange ="fct_onChange();">
 <OPTION></OPTION>
 <%

 for k = 1 to 31
   Response.Write "<option "
 if StrAliasID <> 0 then
  if k = day(objRS("DOLLAR_DATE")) then
    Response.Write " selected "
  end if
 end if
  if k < 10 then
  k="0"&k
  end if
  Response.write " VALUE ="& k & ">" &k & "</OPTION>"
  next
 %>
 </SELECT>
 <SELECT  name=selyear style="HEIGHT: 20px; WIDTH: 60px" onchange ="fct_onChange();">
 <OPTION></OPTION>
 <%
 dim i,baseYear
 baseYear = 1994
 for i = 0 to 30
   Response.Write "<option "
 if StrAliasID <> 0 then
  if (baseYear+i) = year(objRS("DOLLAR_DATE")) then
    Response.Write " selected "
  end if
 end if
  Response.write " VALUE ="& baseYear+i & ">" &baseYear+i & "</OPTION>"
  next
 %>
 </SELECT>
 <INPUT type="button" value="..." id=btnCalendar name=btnCalendar LANGUAGE=javascript onclick="return btnCalendar_onclick(1)">
 </TD>
 </tr>
 <TR>
	<TD ALIGN=RIGHT width=20% NOWRAP>Amount<font color=red>*</font></TD>
	<TD colspan=3 width=80%><INPUT  name=txtamount   style="HEIGHT: 21px; WIDTH: 500px" onchange ="fct_onChange(); round_value();" value= <%if StrAliasID <> 0 then  Response.Write FormatNumber(objRS("DOLLAR_AMOUNT"),-1,-2,-2,0) else Response.Write """""" end if%> >
</TR>


 <TR>
		<TD ALIGN=RIGHT NOWRAP ROWSPAN=2 VALIGN=TOP>Comments:</TD>
		<TD ALIGN=LEFT NOWRAP COLSPAN=3 ROWSPAN=2><TEXTAREA id=txtacomments name=txtacomments ROWS=3 style="WIDTH: 100%" onchange ="fct_onChange();"><%if StrAliasID <> 0 then  Response.Write routineHtmlString(objRS("CHANGE_COMMENTS")) else Response.Write null end if%></TEXTAREA></TD></TR>
	</TR>
</tbody>
</TABLE>

<TABLE>
	  <TR><TD align=right colspan=5>
			<INPUT id=btnClose name=btnClose  type=button value=Close LANGUAGE=javascript onclick="return btnClose_onclick()">&nbsp;&nbsp;
			<INPUT id=btnReset name=btnReset type=reset value=Reset onClick="fct_onReset();" style="HEIGHT: 24px; WIDTH: 51px">&nbsp;&nbsp;
			<INPUT id=btnAddNew  name=btnAddNew type=button value="New" LANGUAGE=javascript onclick="return btnNew_click()">&nbsp;&nbsp;
			<INPUT id=btnDelete  name=btnDelete type=button value=Delete LANGUAGE=javascript onclick="return fct_onDelete();">&nbsp;&nbsp;
			<INPUT  id=btnSave name=btnSave type=button value=Save onClick="btnSave_onclick();">&nbsp;&nbsp;
	  </TD></TR>
</table>

<FIELDSET >
	<LEGEND ALIGN=RIGHT><B>Audit Information</B></LEGEND>
	<Div SIZE=8pt ALIGN=RIGHT>
		Record Status Indicator:
		<INPUT align = left name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 18px"disabled value=<%if  StrAliasID <> 0 then  Response.Write """"&objRS("RECORD_STATUS_IND")&"""" else Response.Write """""" end if%>  >&nbsp;&nbsp;&nbsp;
		Create Date:&nbsp;&nbsp;
		<INPUT align = center name=txtcrdate type=text style="HEIGHT: 20px; WIDTH: 140px"disabled value=<%if  StrAliasID <> 0 then  Response.Write """"&objRS("CREATE_DATE_CONV")&"""" else Response.Write """""" end if%> >&nbsp;
		&nbsp;
		Created By:
		<INPUT align = right name=txtcrby type=text style="HEIGHT: 20px; WIDTH: 100px"disabled value=<%if  StrAliasID <> 0 then  Response.Write """"&routineHtmlString(objRS("CREATE_REAL_USERID"))&"""" else Response.Write """""" end if%> ><BR>
		Update Date:
		<INPUT align= center name=txtupdate type=text style="HEIGHT: 20px; WIDTH: 140px"disabled value=<%if  StrAliasID <> 0 then  Response.Write """"&objRS("UPDATE_DATE_CONV")&"""" else Response.Write """""" end if%>  >&nbsp;&nbsp;
		Updated By:
		<INPUT align=right name=txtupby type=text style="HEIGHT: 20px; WIDTH: 100px"disabled value=<%if  StrAliasID <> 0 then  Response.Write """"&routineHtmlString(objRS("UPDATE_REAL_USERID"))&"""" else Response.Write """""" end if%>  >
	</DIV>
</FIELDSET>

</FORM>
<%

 'Clean up our ADO objects
 IF StrAliasID <> 0 THEN
    objRS.close
    set objRS = Nothing
 END IF

    objRsAssetType.close
    set objRsAssetType = Nothing

    objConn.close
    set ObjConn = Nothing


%>


</BODY>
</HTML>
