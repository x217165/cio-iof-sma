<%@ Language=VBScript %>
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

function iSInstFrame_display() { 
  var strAttrURL = 'SInstAttUsage1.asp?' ;
  document.frames("attfr").document.location.href = strAttrURL ;
} 

function iSInstvFrame_display() { 
  var strAttrURL = 'SInstAttUsage2.asp?';
  document.frames("attvfr").document.location.href = strAttrURL ;
} 

function iSInstuFrame_display() { 
  var strAttrURL = 'SInstAttUsage3.asp?hdnseluSIAtt=0' ;
  document.frames("attufr").document.location.href = strAttrURL ;
} 

function iSInstFrm_Add(){
  var NewWin;
  var strSource ;
  if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
		alert('Access denied. Please contact your system administrator.'); 
		return;
   }
  strSource = 'SInstAttDetail.asp?hdnInstAttID=0';  
  NewWin=window.open(strSource ,"NewWin","toolbar=no,status=no,width=700,height=250,menubar=no resize=no");
  //NewWin=window.open(strSource ,"NewWin") ;
  NewWin.focus();
}

function iSInstvFrm_Add(){
  var NewWin;
  var strSource ;
  if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
		alert('Access denied. Please contact your system administrator.'); 
		return;
   }
  strSource = 'SInstAttvDetail.asp?hdnInstAttvID=0';  
  NewWin=window.open(strSource ,"NewWin","toolbar=no,status=no,width=700,height=250,menubar=no resize=no");
  //NewWin=window.open(strSource ,"NewWin") ;
  NewWin.focus();
}


function iSInstuFrm_Add(){
  var NewWin;
  var strSource ;
  if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
		alert('Access denied. Please contact your system administrator.'); 
		return;
   }
  strSource = 'SInstRuleDetail.asp?hdnAttID=0&hdnAttvID=0';  
  NewWin=window.open(strSource ,"NewWin","toolbar=no,status=no,width=700,height=250,menubar=no resize=no");
  //NewWin=window.open(strSource ,"NewWin") ;
  NewWin.focus();
}


function iSInstFrm_Update(){
  var NewWin ;
   var strSource = 'SInstAttDetail.asp?hdnInstAttID=';
  //changed txtAttID to hdnstrAttID in below line in July 6 2009
  strSource = strSource + document.frames("attfr").frmInstAttRM.selmAtt.value;
  if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update) {
		alert('Access denied. Please contact your system administrator.'); 
		return ;
	}
  if (document.frames("attfr").frmInstAttRM.selmAtt.value !=0){
	NewWin=window.open(strSource ,"NewWin","toolbar=no,status=no,width=700,height=250,menubar=no resize=no") ;
	//NewWin=window.open(strSource ,"NewWin") ;
	NewWin.focus();
   }
  else {
	alert('You must select a record to update!');
   }
} // ************* End of iSInstFrm_Update() ************

function iSInstvFrm_Update(){
  var NewWin ;
   var strSource = 'SInstAttvDetail.asp?hdnInstAttvID=';
  //changed txtAttID to hdnstrAttID in below line in July 6 2009
  strSource = strSource + document.frames("attvfr").frmInstAttRM.selmAttv.value;
  if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update) {
		alert('Access denied. Please contact your system administrator.'); 
		return ;
	}
  if (document.frames("attvfr").frmInstAttRM.selmAttv.value !=0){
	NewWin=window.open(strSource ,"NewWin","toolbar=no,status=no,width=700,height=250,menubar=no resize=no") ;
	//NewWin=window.open(strSource ,"NewWin") ;
	NewWin.focus();
   }
  else {
	alert('You must select a record to update!');
   }
} 


function iSInstuFrm_Update(){
  var NewWin ;
   var strSource = 'SInstRuleDetail.asp?hdnAttID=';
  //changed txtAttID to hdnstrAttID in below line in July 6 2009
  strSource = strSource + document.frames("attufr").frmSIAttRM.seluSIAtt.value +'&hdnAttvID=';
  strSource = strSource + document.frames("attufr").frmSIAttRM.seluSIAttv.value;

  if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update) {
		alert('Access denied. Please contact your system administrator.'); 
		return ;
	}
  if (document.frames("attufr").frmSIAttRM.seluSIAtt.value !=0){
	NewWin=window.open(strSource ,"NewWin","toolbar=no,status=no,width=700,height=250,menubar=no resize=no") ;
	//NewWin=window.open(strSource ,"NewWin") ;
	NewWin.focus();
   }
  else {
	alert('You must select a record to update!');
   }
} // ************* End of iSInstFrm_Update() ************


function iSInstFrame_Delete(){
  var strURL ;
 /*  This section temp commented for my test  -- LC  */
 if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) {
   	alert('Access denied. Please contact your system administrator.') ;
	    return ;
  } 
	if (document.frames("attfr").frmInstAttRM.selmAtt.value !=0) {
		if (confirm('Do you really want to delete this record?')){	
			strURL = 'SInstAttDetail.asp?hdnFrmAction=DELETE&hdnInstAttID=' + document.frames("attfr").frmInstAttRM.selmAtt.value;
			document.frames("attfr").document.location.href = strURL ;
		}
	  }
    else {
		alert('You must select a record to delete!') ;
    }

 
}  // ***************  end of btn_SInstFrameDelete() ******************

function iSInstvFrame_Delete(){
  var strURL ;
 /*  This section temp commented for my test  -- LC  */
 if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) {
   	alert('Access denied. Please contact your system administrator.') ;
	    return ;
  } 
	if (document.frames("attvfr").frmInstAttRM.selmAttv.value !=0) {
		if (confirm('Do you really want to delete this record?')){	
			strURL = 'SInstAttvDetail.asp?hdnFrmAction=DELETE&hdnInstAttvID=' + document.frames("attvfr").frmInstAttRM.selmAttv.value;
			document.frames("attvfr").document.location.href = strURL ;
		}
	  }
    else {
		alert('You must select a record to delete!') ;
    }
 
}  // ***************  end of btn_SInstFrameDelete() ******************


function iSInstuFrame_Delete(){
  var strURL ;
 /*  This section temp commented for my test  -- LC  */
 if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) {
   	alert('Access denied. Please contact your system administrator.') ;
	    return ;
  } 
	if (document.frames("attufr").frmSIAttRM.seluSIAtt.value !=0 )
	{
		if (document.frames("attufr").frmSIAttRM.seluSIAttv.value !=0) 
		{
/*		  strURL = 'SInstRuleDetail.asp?hdnFrmAction=DELETE&hdnAttID='+ document.frames("attufr").frmSIAttRM.seluSIAtt.value;
		  strURL = strURL + '&hdnAttvID=' + document.frames("attufr").frmSIAttRM.seluSIAttv.value;
		  document.write(strURL);*/
			if (confirm('Do you really want to delete this record?'))
			{	
				strURL = 'SInstRuleDetail.asp?hdnFrmAction=DELETE&hdnAttID='+ document.frames("attufr").frmSIAttRM.seluSIAtt.value;
				strURL = strURL + '&hdnAttvID=' + document.frames("attufr").frmSIAttRM.seluSIAttv.value;
				document.frames("attufr").document.location.href = strURL ;
			}
	  	}
	 }
    else 
    {
		alert('You must select a record to delete!') ;
    }
 	 //  iSInstuFrame_display();
}  // ***************  end of btn_SInstFrameDelete() ******************



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
	iSInstFrame_display();
	iSInstvFrame_display();
	iSInstuFrame_display();
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

function iSInstrFrame_Report(){
  var NewWin ;
  var strSource = 'SInstAttRpt.asp?hdnselSInst=';
  strSource = strSource + document.frames("attufr").frmSIAttRM.seluSIAtt.value;
  strSource = strSource + "&hdnselSInstv=" + document.frames("attufr").frmSIAttRM.seluSIAttv.value;
  NewWin=window.open(strSource ,"NewWin","scrollbars=1 toolbar=0,status=0,width=700,height=600,menubar=0 resizable=1");
	//NewWin=window.open(strSource ,"NewWin") ;
 // NewWin.focus();
//  document.location.href=strSource;
} 



// Unhide Client-Side SCRIPT -->
</SCRIPT>
</HEAD>

<BODY language="javascript" onLoad="DisplayStatus('');" 
	onLoad="iSInstFrame_display();iSInstvFrame_display();iSInstuFrame_display();"  
	onBeforeUnload="window_onBeforeUnload();" onUnload="window_onUnload();">
<FORM id="frmSInstList" name="frmSInstList" action="SInstMRDetail.asp" method="post">


<table>



<TABLE>
<thead>
	<tr><td  width="80%">Service Instance Attribute</td></tr>
</thead>
<tbody>
	<td>
 		<iframe id=attfr height=70 src="" scrolling=yes marginheight=1 marginwidth=1 style="width: 99%"></iframe>
		<input type="button" style= "width: 2cm" value="Delete" name="btn_SInstFrameDelete"  onClick="iSInstFrame_Delete();">&nbsp;&nbsp;
		<input type="button" style= "width: 2cm" value="Refresh" name="btn_SInstFrameRefresh"  onClick="iSInstFrame_display();">               &nbsp;&nbsp;
		<input type="button" style= "width: 2cm" value="New"     name="btn_SInstFrameAdd"      onClick="iSInstFrm_Add();">   &nbsp;&nbsp;
		<input type="button" style= "width: 2cm" value="Update"  name="btn_SInstFrameupdate"   onClick="iSInstFrm_Update();">
	</td>
	<td></td>
</tbody>
</TABLE>





<table>
<thead>
	<td width="80%"> Service Instance Attribute Value</td>
</thead>
<tbody>
		<td>
 		<iframe id=attvfr width=100% height=70 src="" scrolling=yes marginheight=1 marginwidth=1></iframe>
		<input type="button" style= "width: 2cm" value="Delete" name="btn_SInstvFrameDelete"  onClick="iSInstvFrame_Delete();">&nbsp;&nbsp;
		<input type="button" style= "width: 2cm" value="Refresh" name="btn_SInstvFrameRefresh"  onClick="iSInstvFrame_display();">               &nbsp;&nbsp;
		<input type="button" style= "width: 2cm" value="New"     name="btn_iSInstvFrameAdd"      onClick="iSInstvFrm_Add();">   &nbsp;&nbsp;
		<input type="button" style= "width: 2cm" value="Update"  name="btn_iSInstvFrameupdate"   onClick="iSInstvFrm_Update();">
		</td>
		
	<td></td>
</tbody>
</table>


<table>
<thead>
	<td  width="80%">Service Instance Attribute Rule and Report</td>
</thead>
<tbody>
		<td>
 		<iframe id=attufr width=100% height=100 src="" scrolling=yes marginheight=1 marginwidth=1></iframe>
		<input type="button" style= "width: 2cm" value="Delete" name="btn_SInstuFrameDelete"  onClick="iSInstuFrame_Delete();">&nbsp;&nbsp;
		<input type="button" style= "width: 2cm" value="Refresh" name="btn_SInstuFrameRefresh"  onClick="iSInstuFrame_display();">               &nbsp;&nbsp;
		<input type="button" style= "width: 2cm" value="New"     name="btn_iSInstuFrameAdd"      onClick="iSInstuFrm_Add();">   &nbsp;&nbsp;
		<input type=hidden style= "width: 2cm" value="Update"  name="btn_iSInstuFrameupdate"   onClick="iSInstuFrm_Update();">
		<input type="button" style= "width: 2cm" value="Report" name="btn_SInstrReport"  onClick="iSInstrFrame_Report();">
	</td>
	<td></td>
</tbody>
</TABLE>

</table>
</FORM>
</BODY>
</HTML>