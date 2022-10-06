<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True
 on error resume next%>
<!--#include file = "smaConstants.inc"-->
<!--#include file = "smaProcs.inc"-->
<!--#include file = "databaseconnect.asp" -->
<!--
*************************************************************************************************
* Page name:	SAttMRDetail.asp																*
* Purpose:		To display the Service Type	Attribute Maintenace Page							*
* Created by:	Linda Chen 08/11/2009															*
* Modifications By				Date				Modifcations								*
*************************************************************************************************
-->
<%Dim intAccessLevel

	'Check user's rights
	intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
%>

<HTML>
<HEAD>
<META name="Generator" content="Microsoft Visual Studio 6.0">
<META http-equiv="Pragma" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<LINK REL="stylesheet" TYPE="text/css" HREF="stylesheets/styles.css">
<SCRIPT type="text/javascript" language="javascript" src="AccessLevels.js"></SCRIPT>
<SCRIPT type="text/javascript" language="javascript" src="GeneralJavaFunctions.js"></SCRIPT>
<SCRIPT type="text/javascript" language="javascript" id="clientEventHandlersJS">
<!-- //Hide Client-Side SCRIPT
var intAccessLevel = <%=intAccessLevel%>;

function iSTAttFrame_display() {
  var strAttrURL = 'StypeAttUsage1.asp?hdnAttID=0' ;
  document.frames("attfr").document.location.href = strAttrURL ;
}

function iSTAttvFrame_display() {
  var strAttrURL = 'StypeAttUsage2.asp?hdnAttvID==0';
  document.frames("attvfr").document.location.href = strAttrURL ;
}

function iSTAttuFrame_display() {
  var strAttrURL = 'STypeAttUsage3.asp?hdnseluSTAtt=0' ;
  document.frames("attufr").document.location.href = strAttrURL ;
}


function iSTAttFrm_Add(){
  var NewWin;
  var strSource ;
  if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
		alert('Access denied. Please contact your system administrator.');
		return;
   }
  strSource = 'SAttDetail.asp?hdnAttID=0';
  NewWin=window.open(strSource ,"NewWin","toolbar=no,status=no,width=700,height=250,menubar=no resize=no");

  NewWin.focus();
}

function iSTAttvFrm_Add(){
  var NewWin;
  var strSource ;
  if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
		alert('Access denied. Please contact your system administrator.');
		return;
   }
  strSource = 'SAttvDetail.asp?hdnAttvID=0';
  NewWin=window.open(strSource ,"NewWin","toolbar=no,status=no,width=700,height=250,menubar=no resize=no");

  NewWin.focus();
}


function iSTAttuFrm_Add(){
  var NewWin;
  var strSource ;
  if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
		alert('Access denied. Please contact your system administrator.');
		return;
   }
  strSource = 'SAttRuleDetail.asp?hdnAttID=0&hdnAttvID=0';
  NewWin=window.open(strSource ,"NewWin","toolbar=no,status=no,width=700,height=250,menubar=no resize=no");

  NewWin.focus();
}


function iSTAttFrm_Update(){
  var NewWin ;
   var strSource = 'SAttDetail.asp?hdnAttID=';

  strSource = strSource + document.frames("attfr").frmSAttRM.selmAtt.value;
  if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update) {
		alert('Access denied. Please contact your system administrator.');
		return ;
	}
  if (document.frames("attfr").frmSAttRM.selmAtt.value !=0){
	NewWin=window.open(strSource ,"NewWin","toolbar=no,status=no,width=700,height=250,menubar=no resize=no") ;
	//NewWin=window.open(strSource ,"NewWin") ;
	NewWin.focus();
   }
  else {
	alert('You must select a record to update!');
   }
} // ************* End of iSTAttFrm_Update() ************

function iSTAttvFrm_Update(){
  var NewWin ;
   var strSource = 'SAttvDetail.asp?hdnAttvID=';
  //changed txtAttID to hdnstrAttID in below line in July 6 2009
  strSource = strSource + document.frames("attvfr").frmSAttRM.selmAttv.value;
 if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update) {
		alert('Access denied. Please contact your system administrator.');
		return ;
	}
  if (document.frames("attvfr").frmSAttRM.selmAttv.value !=0){
	NewWin=window.open(strSource ,"NewWin","toolbar=no,status=no,width=700,height=250,menubar=no resize=no") ;

	NewWin.focus();
   }
  else {
	alert('You must select a record to update!');
   }
} // ************* End of iSTAttFrm_Update() ************


function iSTAttuFrm_Update(){
  var NewWin ;
   var strSource = 'SAttRuleDetail.asp?hdnAttID=';
  //changed txtAttID to hdnstrAttID in below line in July 6 2009
  strSource = strSource + document.frames("attufr").frmSAttRM.seluSTAtt.value +'&hdnAttvID=';
  strSource = strSource + document.frames("attufr").frmSAttRM.seluSTAttv.value;

  if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update) {
		alert('Access denied. Please contact your system administrator.');
		return ;
	}
  if (document.frames("attufr").frmSAttRM.seluSTAtt.value !=0){
	NewWin=window.open(strSource ,"NewWin","toolbar=no,status=no,width=700,height=250,menubar=no resize=no") ;

	NewWin.focus();
   }
  else {
	alert('You must select a record to update!');
   }
} // ************* End of iSTAttFrm_Update() ************


function iSTAttFrame_Delete(){
  var strURL ;

 if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) {
   	alert('Access denied. Please contact your system administrator.') ;
	    return ;
  }
	if (document.frames("attfr").frmSAttRM.selmAtt.value !=0) {
		if (confirm('Do you really want to delete this record?')){
			strURL = 'SAttDetail.asp?hdnFrmAction=DELETE&hdnAttID=' + document.frames("attfr").frmSAttRM.selmAtt.value;
			document.frames("attfr").document.location.href = strURL ;
		}
	  }
    else {
		alert('You must select a record to delete!') ;
    }

}  // ***************  end of btn_STattFrameDelete() ******************

function iSTAttvFrame_Delete(){
  var strURL ;

 if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) {
    	alert('Access denied. Please contact your system administrator.') ;
	    return ;
  }
	if (document.frames("attvfr").frmSAttRM.selmAttv.value !=0) {
		if (confirm('Do you really want to delete this record?')){
			strURL = 'SAttvDetail.asp?hdnFrmAction=DELETE&hdnAttvID=' + document.frames("attvfr").frmSAttRM.selmAttv.value;
			document.frames("attvfr").document.location.href = strURL ;
		}
	  }
    else {
		alert('You must select a record to delete!') ;
    }

}  // ***************  end of btn_STattFrameDelete() ******************


function iSTAttuFrame_Delete(){
  var strURL ;
 if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) {
   	alert('Access denied. Please contact your system administrator.') ;
	    return ;
  }
	if (document.frames("attufr").frmSAttRM.seluSTAtt.value !=0 )
	{
		if (document.frames("attufr").frmSAttRM.seluSTAttv.value !=0)
		{
			if (confirm('Do you really want to delete this record?'))
			{
				strURL = 'SAttRuleDetail.asp?hdnFrmAction=DELETE&hdnAttID='+ document.frames("attufr").frmSAttRM.seluSTAtt.value;
				strURL = strURL + '&hdnAttvID=' + document.frames("attufr").frmSAttRM.seluSTAttv.value;
				document.frames("attufr").document.location.href = strURL ;
			}
	  	}
	 }
    else
    {
		alert('You must select a record to delete!') ;
    }

}  // ***************  end of btn_STattFrameDelete() ******************



//function body_onLoad(){
//	iFrame_display();
//}



function window_onBeforeUnload() {
	//Ensure that fct_onChange() fires for any changed data.
//	document.frmSTypeDetail.btnSave.focus();

//	if (bolSaveRequired) {
//		event.returnValue = "There is unsaved data on the screen. To save changes, click CANCEL below then click SAVE on the main FORM.";
//	}
}

function window_onUnload() {
//
}

function ClearStatus() {
	window.status = "";
}

function DisplayStatus(strWinStatus) {
	window.status = strWinStatus;
	setTimeout('ClearStatus()', 5000);
	iSTAttFrame_display();
	iSTAttvFrame_display();
	iSTAttuFrame_display();

}

function btnReset_onClick() {
	if(confirm('All changes will be lost. Do you really want to reset the page?')){
		bolSaveRequired = false;
		document.location.href = "STypeDetail.asp?ServiceTypeID=<%=strServiceTypeID%>";
	}
}


function fct_onChange() {
// some comments
}

function iSTAttrFrame_Report(){
  var NewWin ;
  var strSource = 'STypeAttRpt.asp?hdnselSTAtt=';
  strSource = strSource + document.frames("attufr").frmSAttRM.seluSTAtt.value;
  strSource = strSource + "&hdnselSTAttv=0" + document.frames("attufr").frmSAttRM.seluSTAttv.value;

  NewWin=window.open(strSource ,"NewWin","scrollbars=1,toolbar=0,status=0,width=700,height=600,menubar=0 resizable=1");
	//NewWin=window.open(strSource ,"NewWin") ;
 // NewWin.focus();
//  document.location.href=strSource;
}



// Unhide Client-Side SCRIPT -->
</SCRIPT>
</HEAD>

<BODY language="javascript" onLoad="DisplayStatus('');"
	onLoad="iSTAttFrame_display();iSTAttvFrame_display();iSTAttuFrame_display();"
	onBeforeUnload="window_onBeforeUnload();" onUnload="window_onUnload();">
<FORM id="frmSTAttList" name="frmSTAttList" action="SAttMRDetail.asp" method="post">


<table>



<TABLE>
<thead>
	<tr><td  width="80%">Service Type Attribute</td></tr>
</thead>
<tbody>
	<td>
 		<iframe id=attfr width=100% height=70 src="" scrolling=yes marginheight=1 marginwidth=1></iframe>
		<input type="button" style= "width: 2cm" value="Delete" name="btn_STattFrameDelete"  onClick="iSTAttFrame_Delete();">&nbsp;&nbsp;
		<input type="button" style= "width: 2cm" value="Refresh" name="btn_STAttFrameRefresh"  onClick="iSTAttFrame_display();">               &nbsp;&nbsp;
		<input type="button" style= "width: 2cm" value="New"     name="btn_STAttFrameAdd"      onClick="iSTAttFrm_Add();">   &nbsp;&nbsp;
		<input type="button" style= "width: 2cm" value="Update"  name="btn_STAttFrameupdate"   onClick="iSTAttFrm_Update();">
	</td>
	<td></td>
</tbody>
</TABLE>





<table>
<thead>
	<td width="80%"> Service Type Attribute Value</td>
</thead>
<tbody>
		<td>
 		<iframe id=attvfr width=100% height=70 src="" scrolling=yes marginheight=1 marginwidth=1></iframe>
		<input type="button" style= "width: 2cm" value="Delete" name="btn_STattvFrameDelete"  onClick="iSTAttvFrame_Delete();">&nbsp;&nbsp;
		<input type="button" style= "width: 2cm" value="Refresh" name="btn_STAttvFrameRefresh"  onClick="iSTAttvFrame_display();">               &nbsp;&nbsp;
		<input type="button" style= "width: 2cm" value="New"     name="btn_iSTAttvFrameAdd"      onClick="iSTAttvFrm_Add();">   &nbsp;&nbsp;
		<input type="button" style= "width: 2cm" value="Update"  name="btn_iSTAttvFrameupdate"   onClick="iSTAttvFrm_Update();">
		</td>

	<td></td>
</tbody>
</table>


<table>
<thead>
	<td  width="80%">Service Type Attribute Rule and Report</td>
</thead>
<tbody>
		<td>
 		<iframe id=attufr width=100% height=100 src="" scrolling=yes marginheight=1 marginwidth=1></iframe>
		<input type="button" style= "width: 2cm" value="Delete" name="btn_STattuFrameDelete"  onClick="iSTAttuFrame_Delete();">&nbsp;&nbsp;
		<input type="button" style= "width: 2cm" value="Refresh" name="btn_STAttuFrameRefresh"  onClick="iSTAttuFrame_display();">               &nbsp;&nbsp;
		<input type="button" style= "width: 2cm" value="New"     name="btn_iSTAttuFrameAdd"      onClick="iSTAttuFrm_Add();">   &nbsp;&nbsp;
		<input type="hidden" style= "width: 2cm" value="Update"  name="btn_iSTAttuFrameupdate"   onClick="iSTAttuFrm_Update();">&nbsp;&nbsp;
		<input type="button" style= "width: 2cm" value="Report"  name="btn_iSTAttuFrameReport"   onClick="iSTAttrFrame_Report();">
	</td>
	<td></td>
</tbody>
</TABLE>

</table>
</FORM>
</BODY>
</HTML>