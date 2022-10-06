<%@ Language=VBScript %>
<% Option Explicit 
 on error resume next
%>
<% Response.Buffer = true %>
<!--#include file="smaConstants.inc"-->
<!--#include file="databaseconnect.asp"-->
<!-- #include file="kenanconnect.asp" -->
<!--#include file="smaProcs.inc"-->
<%
'************************************************************************************************
'* Page name:	STypeDetail.asp																	*
'* Purpose:		To display the Service Type														*
'*				Chosen via STypeList.asp														*
'*																								*
'* Created by:					Date															*
'* Sara Sangha					02/15/2000														*
'*==============================================================================================*
'* Modifications By				Date				Modifcations								*
'*																								*		
'* 																								*
'************************************************************************************************

Dim intAccessLevel, strRealUserID
Dim strXRefID, strServiceTypeID, strKenCompID, strAttvID, strUsageID, strhelptext
Dim strSQL, objRS, strWinMessage, objRsSTKenan, objRsSTAvalue

intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
strRealUserID = Request.Cookies("UserInformation")("username")
strXRefID = Request("XRefID")  
strServiceTypeID = Request("ServiceTypeID")

if strXRefID <> 0 and strXRefID <> "" then
	
	'Response.Write "Service Type :" & strServiceTypeID & "<P>"	
	
		StrSql =" SELECT CREATE_DATE_TIME, " &_	
				    "CREATE_REAL_USERID, " &_	
				    "UPDATE_DATE_TIME, " &_	
				    "UPDATE_REAL_USERID, " &_
				    "RECORD_STATUS_IND, " &_	
				    "REP_HELP_TEXT " &_
			" FROM CRP.SERVICE_TYPE_KENAN_XREF " &_
			" WHERE service_type_id = " & strServiceTypeID  &_
			" and component_id = " & strXRefID
			

	set objRs = objConn.Execute(strSql)
	if err then
		DisplayError "BACK", "", err.Number, "ERROR IN SELECTING Kenan ATTRIBUTES", err.Description
	end if
	
end if

'if strXRefID <> 0 then  strhelptext = objRS(5) end if
'strhelptext = "test 999888111"

'response.write(objRs(5))
'response.end
if intAccessLevel and intConst_Access_ReadOnly <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access Service Type. Please contact your system administrator"
end if

Select case Request("txtFrmAction")
	
	case "SAVE" 

	 if (Request.Form("hdnXRefID") <> "") then
	
		'The XRefID is not null i.e. it is an existing record. So call the update procedure to update the record
		 if ((intAccessLevel and intConst_Access_Update) <> intConst_Access_Update) then
		   DisplayError "BACK", "", 0, "UPDATE DENIED", "You don't have access to update this record. Please contact your system administrator"
		 end if
		 
		    strXRefID = Request.Form("hdnXRefID")
		   
			dim cmdUpdateObj
			set cmdUpdateObj = server.CreateObject("ADODB.Command")
			set cmdUpdateObj.ActiveConnection = objConn
			cmdUpdateObj.CommandType = adCmdStoredProc
			cmdUpdateObj.CommandText = "sma_sp_userid.Sp_Srvtype_Kenan_Xref_Update" 
			
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_user_id", adVarChar , adParamInput, 20, strRealUserID) 
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_old_component_id",adNumeric , adParamInput,, Clng(Request("hdnXRefID")))					
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_new_component_id",adNumeric , adParamInput,, Clng(Request("selSTKenan")))					
'			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_component_id",adNumeric , adParamInput,, Clng(Request("selSTKenan")))					
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_service_type_id",adNumeric , adParamInput,, Clng(Request("hdnServiceTypeID")))					
			'IF	Request("txthelptext") <> "" THEN
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_help_text",adVarChar , adParamInput,255, Request("txthelptext"))
			'else
			'	cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_help_text",adVarChar , adParamInput,255, NULL)
			'end if
	
			'****************************
			'check parameter values		
  			'****************************
  			
  			'dim objparm
  			'for each objparm in cmdUpdateObj.Parameters
  			'	  Response.Write "<b>" & objparm.name & "</b>"
  			'	  Response.Write " has size:  " & objparm.Size & " "
  			'	  Response.Write " and value:  " & objparm.value & " "
  			'	  Response.Write " and datatype:  " & objparm.Type & "<br> "
  			'next

  			'Response.Write "<b> count = " & cmdUpdateObj.Parameters.count & "<br>"
  			'dim nx
  			'for nx=0 to cmdUpdateObj.Parameters.count-1
  			'   Response.Write nx+1 & " parm value= " & cmdUpdateObj.Parameters.Item(nx).Value  & "<br>"
  			'next 	

  								 
  			'response.write (cmdUpdateObj.CommandText)
  			'response.end

			
			cmdUpdateObj.Execute
		
			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT UPDATE RECORD", objConn.Errors(0).Description
				objConn.Errors.Clear
			else
				strWinMessage = "Record Updated successfully. You can now see the changes you made."
				strXRefID =  cmdUpdateObj.Parameters("p_new_component_id").Value
				strServiceTypeID = cmdUpdateObj.Parameters("p_service_type_id").Value
				txthelptext = cmdUpdateObj.Parameters("p_help_text").Value
			end if
			
	else 'create a new record 
      
		   if ((intAccessLevel and intConst_Access_Create) <> intConst_Access_Create) then
		     DisplayError "BACK", "", 0, "INSERT DENIED", "You don't have access to add Default Kenan Attribute. Please contact your system administrator"
		   end if		
			
			dim cmdInsertObj
			set cmdInsertObj = server.CreateObject("ADODB.Command")
			set cmdInsertObj.ActiveConnection = objConn
			cmdInsertObj.CommandType = adCmdStoredProc
			cmdInsertObj.CommandText = "sma_sp_userid.Sp_Srvtype_Kenan_Xref_Insert" 
			
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_user_id", adVarChar , adParamInput, 20, strRealUserID) 
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_service_type_Kenan_xref_id",adNumeric, adParamOutput,,null)
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_component_id",adNumeric , adParamInput,, Clng(Request("selSTKenan")))					
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_service_type_id",adNumeric , adParamInput,, Clng(Request("hdnServiceTypeID")))					
			'If Request("txthelptext") <> ""
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_help_text",adVarChar , adParamInput,255, Request("txthelptext"))					
			'else
			'	cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_help_text",adVarChar , adParamInput,255, NULL)					
			'end if
     		'****************************
			'check parameter values		
  			'****************************
  			
  			'dim objparm
  			'for each objparm in cmdInsertObj.Parameters
  			'	  Response.Write "<b>" & objparm.name & "</b>"
  			'	  Response.Write " has size:  " & objparm.Size & " "
  			'	  Response.Write " and value:  " & objparm.value & " "
  			'	  Response.Write " and datatype:  " & objparm.type & "<br> "
  			'next
		 
  			'Response.Write "<b> count = " & cmdInsertObj.Parameters.count & "<br>"
  			'dim nx
  			'for nx=0 to cmdInsertObj.Parameters.count-1
  			'   Response.Write nx+1 & " parm value= " & cmdInsertObj.Parameters.Item(nx).Value  & "<br>"
  			'next 	
		
  			'response.write (cmdInsertObj.CommandText)
			'response.end
					
				
				
			cmdInsertObj.Execute
			
			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT ADD NEW RECORD", objConn.Errors(0).Description
				objConn.Errors.Clear
			else
				 strXRefID =  cmdInsertObj.Parameters("p_component_id").Value
				 strServiceTypeID = cmdInsertObj.Parameters("p_service_type_id").Value
			end if
			strWinMessage = "Record created successfully. You can now see the new record."
			
	end if
	     
		
 end select  
		  
 strSQL = "SELECT component_id, COMPONENT_NAME, PACKAGE_NAME FROM ARBOR.V_PKG_COMPONENTS " &_
    		" order by component_id, COMPONENT_NAME, PACKAGE_NAME"
 'Create Recordset object  
 'response.write(strSQL)
 'response.end
'response.write(strhelptext)
'response.end
 set objRsSTKenan = objKenanConn.Execute(strSQL)
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<LINK REL="stylesheet" TYPE="text/css" HREF="stylesheets/styles.css">
<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
<script type="text/javascript" SRC="AccessLevels.js"></script>
<TITLE>Default Kenan Attribute</TITLE>
<SCRIPT LANGUAGE=javascript>
<!--

var bolSaveRequired = false;
var intAccessLevel=<%=intAccessLevel%>;
var intConst_MessageDisplay=<%=intConst_MessageDisplay%>;


function fct_clearStatus() {
//**********************************************************************************************
// Function:	btnClose_onclick()
//
// Purpose:		clear the message from window status bar.
//
// Creaded By:	Ian Harriott	
//**********************************************************************************************
	window.status = "" ;
}

function fct_displayStatus(strMessage){
//**********************************************************************************************
// Function:	btnClose_onclick()
//
// Purpose:		display a message in window status bar and then clear it after the set minutes.
//
// Creaded By:	Ian Harriott
//**********************************************************************************************
	window.status = strMessage;
	setTimeout('fct_clearStatus()',intConst_MessageDisplay);
}

function body_onLoad(strWinStatus){
//**********************************************************************************************
// Function:	btnClose_onclick()
//
// Purpose:		Whenever the page is loaded it displays a message in window status bar.
//
// Creaded By:	Sara Sangha		Feb. 15th, 2001
//**********************************************************************************************
	var strWinStatus='<%=strWinMessage%>';
	fct_displayStatus(strWinStatus);
}

function body_onUnload(){
//**********************************************************************************************
// Function:	btnClose_onclick()
//
// Purpose:		Refresh contents of iFrame in the base window.
//
// Creaded By:	Sara Sangha		Feb. 15th, 2001
//**********************************************************************************************

	opener.document.frmSTypeDetail.btn_KenanFrameRefresh.click();
}


function btnClose_onclick(){
//**********************************************************************************************
// Function:	btnClose_onclick()
//
// Purpose:		close the pop up window and Refresh the contents of iFrame in the base window.
//
// Creaded By:	Sara Sangha		Feb. 15th, 2001
//**********************************************************************************************
	
	window.close();
	parent.opener.iSKenanFrame_display()
//  iSTAFrame_display();

}

function frmSKenanDetail_onsubmit() {
//**********************************************************************************************
// Function:	frmSKenanDetail_onsubmit()
//
// Purpose:		set the frmAction to SAVE if the user has access to save the record.
//
// Creaded By:	Sara Sangha		Feb. 15th, 2001
//**********************************************************************************************

if	(((intAccessLevel & intConst_Access_Create == intConst_Access_Create) && (document.frmSKenanDetail.hdnXRefID.value == "")) 
		|| ((intAccessLevel & intConst_Access_Update == intConst_Access_Update) && (document.frmSKenanDetail.hdnXRefID.value != ""))) {
		
			document.frmSKenanDetail.txtFrmAction.value = "SAVE";
			bolSaveRequired = false;
			return(true);

			var strhelptext = document.frmSKenanDetail.txthelptext.value ;
			if (strhelptext.length > 255 ) {
				alert('The help text can be at most 255 characters.\n\nYou entered ' + strComments.length + ' character(s).');
				document.frmSKenanDetail.txthelptext.focus();
				return(false);  }

		}	
   else {
		alert('Access denied. Please contact your system administrator.'); 
		return(false);
	}
}


function btnSave_onclick() {
//**********************************************************************************************
// Function:	btnClose_onclick()
//
// Purpose:		close the pop up window and Refresh the iFrame in the base window.
//
// Creaded By:	Ian Harriott		Feb. 15th, 2001
//**********************************************************************************************
var bolretval
	
	bolretval= frmSKenanDetail_onsubmit();
	
	if(bolretval)
		document.frmSKenanDetail.submit();
	
	//window.close();
	parent.opener.iSKenanFrame_display();

}

function fct_onChange(){
//**********************************************************************************************
// Function:	btnClose_onclick()
//
// Purpose:		set the bolSaveRequired flag if anything changes on the screen.
//
// Creaded By:	Sara Sangha		Feb. 15th, 2001
//**********************************************************************************************
if (intAccessLevel >= intConst_Access_Create){
if (document.frmSKenanDetail.hdnXRefID.value != "") 
     {
     bolSaveRequired = true;
     }
     
    } 
}


function btnNew_click(){
//**********************************************************************************************
// Function:	btnClose_onclick()
//
// Purpose:		if the user has access to add new records then submit the page to itself with 
//				XRefID = 0 so that it will display a blank page.
//
// Creaded By:	Sara Sangha		Feb. 15th, 2001
//**********************************************************************************************
var strURL ;

	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
		alert('Access denied. Please contact your system administrator.');
	    return;
	}
	
	strURL = 'STypeKenanDetail.asp?XRefID=0&ServiceTypeID=' + document.frmSKenanDetail.hdnServiceTypeID.value ;  
	self.document.location.href = strURL ;
}


function fct_onDelete(){
//**********************************************************************************************
// Function:	btnClose_onclick()
//
// Purpose:		if the user has access to delete a record then set frmAction = 'DELETE' and pass in
//				in the required parameterst to delete a record
//
// Creaded By:	Sara Sangha		Feb. 15th, 2001
//**********************************************************************************************
var strURL ;

	if (document.frmSKenanDetail.hdnXRefID.value != "") {
		if (((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) || (document.frmSKenanDetail.txtRecordStatusInd.value == "D")){ 
			alert('Access denied. Please contact your system administrator.');
			return;
		}
		
		if (confirm('Do you really want to delete this object?')){
			
			strURL = 'STypeKenanDetail.asp?txtFrmAction=DELETE&XRefID=' 
					+ document.frmSKenanDetail.hdnXRefID.value + '&UpdateDateTime=' 
					+ document.frmSKenanDetail.hdnUpdateDateTime.value + '&ServiceTypeID=' 
					+ document.frmSKenanDetail.hdnServiceTypeID.value;
					
			document.location = strURL ;
		}
		 
	else {
		fct_displayStatus('Unable to Delete the record. No Record ID provided.');
		return(false);
	}	
  }	
}


function body_onBeforeUnload(){
//**********************************************************************************************
// Function:	btnClose_onclick()
//
// Purpose:		Give a warrening message is there is unsaved data on the screen.
//
// Creaded By:	Sara Sangha		Feb. 15th, 2001
//**********************************************************************************************

    document.frmSKenanDetail.btnSave.focus();
	if (bolSaveRequired) {
		if	(((intAccessLevel & intConst_Access_Create == intConst_Access_Create) && (document.frmSKenanDetail.txtcktalias.value == "")) || ((intAccessLevel & intConst_Access_Update == intConst_Access_Update) && (document.frmSKenanDetail.txtcktalias.value != ""))) {
			event.returnValue = "There is unsaved data on the screen. To save changes, click CANCEL below then click SAVE on the main form.";
		}
	}
}


function fct_onReset(){
//**********************************************************************************************
// Function:	btnClose_onclick()
//
// Purpose:		Refresh the contents on the screen from databaase.
//
// Creaded By:	Sara Sangha		Feb. 15th, 2001
//**********************************************************************************************
   	bolSaveRequired = false;
	//When reset screen for Update
    if (document.frmSKenanDetail.hdnXRefID.value != "")  {
	    document.frmSKenanDetail.selSTKenan.value = document.frmSKenanDetail.hdnXRefID.value;
	}
	//When reset screen for New
	else {
		
	document.frmSKenanDetail.selSTKenanv.value="";
	document.frmSKenanDetail.selSTKenan.value="";
	}
}


//-->
</SCRIPT>
</HEAD>

<BODY onLoad="body_onLoad();" onBeforeUnload="body_onBeforeUnload();" onUnload="body_onUnload();" >
<FORM  name=frmSKenanDetail action="STypeKenanDetail.asp" method="POST" onsubmit="return frmSKenanDetail_onsubmit()">
	<INPUT  name=txtFrmAction type=hidden value="" >
	<INPUT  name=hdnXRefID type=hidden  value= <%if strXRefID <> 0 then  Response.Write strXRefID else Response.Write """""" end if%> >
	<INPUT  name=hdnServiceTypeID type=hidden value= <% =strServiceTypeId %>>

<TABLE>
<thead>
	<TR ><TD colspan=2>Kenan Service Type Detail</td></tr>
</thead>

<tbody>
<TR>
	<TD ALIGN=RIGHT width=20% NOWRAP>Kenan Service Type Detail<font color=red>*</font></TD> 
	<TD width=80%>
		<SELECT id=selSTKenan name=selSTKenan style="HEIGHT: 22; WIDTH: 507" onchange ="fct_onChange();">
		<OPTION ></OPTION>
		<%Do while Not objRsSTKenan.EOF 
		   dim strken
		   strken = objRsSTKenan(0)& " | " & objRsSTKenan(1) & " | " & objRsSTKenan(2)%>
		   <option  <% if strXRefID <> 0 then
		   				if clng(strXRefID) = clng(objRsSTKenan(0)) then 
		              		response.write "selected "
		              	end if
		              end if %>
           value = <% =objRsSTKenan(0) %>
		   > <% =strken %> </option>
		<%  objRsSTKenan.MoveNext   
		Loop
		%>
	
		</SELECT>
	</TD>
</TR>
<TR>	
	<td valign="top" align="right" rowSpan="4">Help Text</td>
	<td valign="top" rowSpan="2"><textarea style="width=80%" name="txthelptext" onChange="fct_onChange();" rows=3><% if strXRefID <> 0 then Response.write routineHtmlString(objRS("REP_HELP_TEXT")) end if%></textarea></td>	
</TR>

</tbody>
</TABLE>

<TABLE>
	  <TR><TD align=right>
			<INPUT id=btnClose   name=btnClose  type=button style="width:2cm" value=Close  LANGUAGE=javascript onclick="return btnClose_onclick()"> &nbsp;&nbsp;
			<INPUT id=btnReset   name=btnReset  type=button style="width:2cm" value=Reset  LANGUAGE=javascript onClick="return fct_onReset();fct_onReset();" >           &nbsp;&nbsp;
			<INPUT id=btnSave    name=btnSave   type=button style="width:2cm" value=Save   LANGUAGE=javascript onclick="return btnSave_onclick();">        &nbsp;&nbsp;
	  </TD></TR>
</table>

<FIELDSET >
	<LEGEND ALIGN=RIGHT><B>Audit Information</B></LEGEND>
	<Div SIZE=8pt ALIGN=RIGHT>
		Record Status Indicator: 
		<INPUT align = left name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 18px"disabled value=<%if  strXRefID <> 0 then  Response.Write """"&objRS("RECORD_STATUS_IND")&"""" else Response.Write """""" end if%>  >&nbsp;&nbsp;&nbsp;
		Create Date:&nbsp;&nbsp;
		<INPUT align = center name=txtcrdate type=text style="HEIGHT: 20px; WIDTH: 140px"disabled value=<%if  strXRefID <> 0 then  Response.Write """"&objRS("CREATE_DATE_TIME")&"""" else Response.Write """""" end if%> >&nbsp;
		&nbsp;
		Created By:
		<INPUT align = right name=txtcrby type=text style="HEIGHT: 20px; WIDTH: 100px"disabled value=<%if  strXRefID <> 0 then  Response.Write """"&routineHtmlString(objRS("CREATE_REAL_USERID"))&"""" else Response.Write """""" end if%> ><BR>
		Update Date:
		<INPUT align= center name=txtupdate type=text style="HEIGHT: 20px; WIDTH: 140px"disabled value=<%if  strXRefID <> 0 then  Response.Write """"&objRS("UPDATE_DATE_TIME")&"""" else Response.Write """""" end if%>  >&nbsp;&nbsp;
		Updated By:
		<INPUT align=right name=txtupby type=text style="HEIGHT: 20px; WIDTH: 100px"disabled value=<%if  strXRefID <> 0 then  Response.Write """"&routineHtmlString(objRS("UPDATE_REAL_USERID"))&"""" else Response.Write """""" end if%>  >
	</DIV>
</FIELDSET>	
	
</FORM>
<%
   
 'Clean up our ADO objects
 if strXRefID <> 0 then
    objRS.close
    set objRS = Nothing
 end if

 '   objRsCktProv.close
  '  set objRsCktProv = Nothing

    objConn.close
    set ObjConn = Nothing    
       
      
%>


</BODY>
</HTML>
