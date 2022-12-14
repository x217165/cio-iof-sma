<%@ Language=VBScript %>
<%   OPTION EXPLICIT
on error resume next %>
<!-- #include file="smaConstants.inc" -->
<!-- #include file="smaProcs.inc" -->
<!-- #include file="databaseconnect.asp" -->
<!--
*************************************************************************************************
* Page name:	STypeAttrList.asp																*
* Purpose:		To display the service instance attributes and its values inside a frame 										*														*
*																								*
* Created by:					Date															*
* Sara Sangha					02/15/2000														*
* ====================================															*
* Modifications By				Date				Modifcations								*
*
*************************************************************************************************
-->
<%

Dim objRs, strSQL, strORDER
Dim strServiceTypeID, strAction,strAttID, strprevAttId, strnextAttId
Dim intAccessLevel

strServiceTypeID = Request("hdnServiceTypeID")


Dim strselIndex
strselIndex = Request("hdnSelIndex")
strAttID = Request("txtselAttID")
strprevAttId = Request("txtprevAttID")
strnextAttId = Request("txtnextAttId")


dim strRealUserID
strRealUserID = Session("username")

intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
if intAccessLevel and intConst_Access_ReadOnly <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access Service Type Attribute. Please contact your system administrator"
end if



if isnumeric(strServiceTypeID)  then


	'Response.Write "Service Type :" & strServiceTypeID & "<P>"

	StrSql ="Select distinct ATT.SRVC_INSTNC_ATT_NAME, "&_
			" XREF.SRVC_INSTNC_ATT_ID, "&_
			" decode (XREF.DISPLAY_ORDER,0,999,XREF.DISPLAY_ORDER) as display_order  " &_
			" from SO.SRVC_INSTNC_ATT_XREF xref, "&_
            " SO.SRVC_INSTNC_ATT att, " &_
        	" SO.SRVC_INSTNC_ATT_VAL_USAGE attu "&_
			" where XREF.SRVC_INSTNC_ATT_ID = ATT.SRVC_INSTNC_ATT_ID " &_
			" and     XREF.SRVC_INSTNC_ATT_XREF_ID = ATTU.SRVC_INSTNC_ATT_XREF_ID " &_
			" and     ATTU.RECORD_STATUS_IND='A' "  &_
			" and     XREF.SERVICE_TYPE_ID = " & strServiceTypeID


	strORDER = " order by display_order "

	StrSql = StrSql + strORDER

	'Response.Write "SQL :" & StrSql & "<P>"

	'response.write StrSql
	'response.end
 	set objRs = objConn.Execute(strSql)

	if err then

		DisplayError "BACK", "", err.Number, "ERROR IN SQL", StrSql
		objConn.Errors.Clear
	end if

	dim aList
	if not objRS.EOF then
		aList = objRS.GetRows
	else
		Response.Write "0 records found"
		Response.end
	end if
	objRs.Close
	set objRs = Nothing

	'release connection
	'set objRs.ActiveConnection = nothing

end if

if Request("txtFrmAction") = "moveup" or Request("txtFrmAction") = "movedown" then
	dim cmdUpdateObj
			set cmdUpdateObj = server.CreateObject("ADODB.Command")
			set cmdUpdateObj.ActiveConnection = objConn
			cmdUpdateObj.CommandType = adCmdStoredProc
			cmdUpdateObj.CommandText = "SMA_SP_USERID.Sp_Srvinst_Val_Xref_SetSeq"

			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_user_id", adVarChar , adParamInput, 20, strRealUserID)
			if Request("txtFrmAction") = "moveup" then
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_direction", adVarChar , adParamInput, 20, "up")
			else
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_direction", adVarChar , adParamInput, 20, "down")
			end if
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_service_type_id",adNumeric , adParamInput,, clng(strServiceTypeID))

			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_srvc_inst_att_id_prev",adNumeric , adParamInput,, clng(strprevAttId))
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_srvc_inst_att_id",adNumeric , adParamInput,, clng(strAttID))
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_srvc_inst_att_id_next",adNumeric , adParamInput,, clng(strnextAttId))


			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_seq",adNumeric , adParamInput,, Clng(strselIndex))


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
				response.redirect("STypeAttDetail.asp")
			else
				'strWinMessage = "Record Updated successfully. You can now see the changes you made."
        		'response.write("<script language=""javascript"">window.close();parent.opener.iSINSTFrame_display();</script>")
      response.redirect "StypeInstListSeq.asp?hdnServiceTypeID=" + strServiceTypeID





	end if


end if

%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK href="stylesheets/styles.css" rel="stylesheet" type="text/css">
<STYLE>
.regularItem {
	cursor: hand;
}
.whiteItem {
	cursor: hand;
	background-color: white; }

.Highlight {
	cursor: hand;
	background-color: #00974f;
	color: white;
}
</STYLE>

<script type="text/javascript">

var oldHighlightedElement;
var oldHighlightedElementClassName;

var txtselAttID,txtprevAttID,txtSelIndex,txtnextAttId
var txtsiatotal
//var txtsiaseq
var intServiceTypeID ;

function reloadPage()
{
  location.reload()
}

function cell_onClick(prevAttId, selAttID, intsiaseq, nextAttId, intServiceType, intsiatotal){

	//document.frmIFR.txtattID.value = intatID;
	//document.frmIFR.txtsiaseq.value = intsiaseq;
	//document.frmIFR.hdnServiceTypeID.value = intServiceType;
	//highlight current record
	txtprevAttID = prevAttId;
	txtselAttID = selAttID
    txtSelIndex = intsiaseq;
    txtnextAttId = nextAttId
    intServiceTypeID=intServiceType;
    txtsiatotal=intsiatotal;
   // intServiceTypeID =intServiceType

	if (oldHighlightedElement != null) {
		oldHighlightedElement.className = oldHighlightedElementClassName
	}
	oldHighlightedElement = window.event.srcElement.parentElement;
	oldHighlightedElementClassName = oldHighlightedElement.className;
	oldHighlightedElement.className = "Highlight";



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
// Purpose:		Refresh contents
//**********************************************************************************************
	opener.document.frmSTypeDetail.btn_iSINSSetSIASeq.click();
//  opener.document.frmSTypeDetail.btn_iSINSFrameRefresh.click();


}





function fct_onMoveUp(){
//lc	if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update) {alert('Access denied. Please contact your system administrator.'); return;}
    if (txtselAttID == undefined)
    {
    	alert('You need select a SIA to move up'); return;
    }
    if (txtSelIndex == 1)
    {
    	alert('You can not move up tis sia as it is at list top already'); return;
    }

	var strURL, strParams;
	strParam="";
	strParams = '&txtprevAttID='+ txtprevAttID +'&txtselAttID=' + txtselAttID +'&txtnextAttId='+ txtnextAttId;
    strParams = strParams +'&hdnSelIndex='+ txtSelIndex;


	document.location.href='StypeInstListSeq.asp?txtFrmAction=moveup&hdnServiceTypeID=' + intServiceTypeID +  strParams;

}  // ***************  end of fct_onMoveUp() ******************


function fct_onMoveDown(){
//lc	if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update) {alert('Access denied. Please contact your system administrator.'); return;}
    //alert('txtselAttId = ' + txtselAttID); return;
    if (txtselAttID == undefined)
    {
    	alert('You need select a SIA to move down'); return;
    }
    if (txtSelIndex == txtsiatotal)
    {
    	alert('You can not move down as this sia is at list bottom already'); return;
    }

	var strURL, strParams;
	strParam="";
	strParams = '&txtprevAttID='+ txtprevAttID +'&txtselAttID=' + txtselAttID +'&txtnextAttId='+ txtnextAttId;
    strParams = strParams +'&hdnSelIndex='+ txtSelIndex;

	document.location.href='StypeInstListSeq.asp?txtFrmAction=movedown&hdnServiceTypeID=' + intServiceTypeID +  strParams;
}

function btnClose_onclick(){
//**********************************************************************************************
// Function:	btnClose_onclick()
// Purpose:		close the pop up window and Refresh the contents of iFrame in the base window.
//**********************************************************************************************

	window.close();
	parent.opener.iSINSTFrame_display();

}


</script>

</HEAD>
<BODY onUnload="body_onUnload();>
<form name="frmInstSeq" action="STypeInstListSeq.asp" method="POST">

		<input type=hidden name=hdnServiceTypeID value="">
		<input type=hidden name=UpdateDateTime value="">
        <input type=hidden name=hdnSelIndex value="<%if selIndex="" then response.write"" else response.write selIndex end if%>">
        <input type=hidden name=txtprevAttID value="">
        <input type=hidden name=txtselAttID value="">
        <input type=hidden name=nextAttId value="">




<TABLE border=1 cellspacing=0 cellpadding=2 width="100%">
	<thead>
		<th>Service Instance Attribute</th>
		<th>Display Order</th>
	</thead>
	<tbody>

<%	if isnumeric(strServiceTypeID)  then

		dim k
		k = 0
		dim listsize
		listsize = UBound(aList, 2)
		while k <= listsize
			if Int(k/2) = k/2 then
				Response.Write "<tr class=""regularItem"">"
			else
				Response.Write "<tr class=""whiteItem"">"
			end if

			%>
		<td nowrap onClick="cell_onClick(<%if k=0 then response.write 0 else response.write aList(1,k-1) end if %>, <%=aList(1,k)%>, <%=k+1%>,
										 <%if k=listsize then response.write 0 else response.write aList(1,k+1) end if %>,
										 <%=strServiceTypeID%>,<%=UBound(aList, 2)+1%>);"><%=aList(0,k)%>&nbsp;</td>
		<td nowrap onClick="cell_onClick(<%if k=0 then response.write 0 else response.write aList(1,k-1) end if %>, <%=aList(1,k)%>, <%=k+1%>,
										 <%if k=listsize then response.write 0 else response.write aList(1,k+1) end if %>,
										 <%=strServiceTypeID%>,<%=UBound(aList, 2)+1%>);"><%if aList(2,k)="999" then response.write"" else response.write(aList(2,k))%>&nbsp;</td>
			</tr>
			<%
			k = k+1
		wend



 end if
 %>
 <tr>

</tr>
</tbody>

</table>
	<INPUT id=btnClose   name=btnClose  type=button style="width:2cm" value=Close  LANGUAGE=javascript onclick="return btnClose_onclick()">

	<img SRC="images/down.gif" title WIDTH="31" HEIGHT="31" onclick="fct_onMoveDown()" style="float: right" >
	<img SRC="images/up.gif" title WIDTH="34" HEIGHT="31" onclick="fct_onMoveUp();" style="float: right"></FORM>
</BODY>
</HTML>


