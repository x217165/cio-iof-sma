<%@ Language=VBScript %>
<% Option Explicit%>
<% Response.Buffer = true %>
<!--#include file = "SmaConstants.inc"-->
<!--#include file = "SmaProcs.inc"-->
<!--#include file = "databaseconnect.asp"-->
<!--
************************************************************************************
*	Page name:  AssetTypeCriteria.asp
*	Purpose:	To dynamically create the criteria to search crp.asset_type.
*	In param:	AssetClassTypeID,AssetClassID,AssetSubClassID,AssetTypeID,AssetTypeDesc
*	Out param:
*	Created by: Nancy Mooney Oct.16,2000
************************************************************************************
-->
<%
Dim strAssetClassTypeID,strAssetClassID,strAssetSubclassID,strAssetTypeID, strAssetTypeDesc
Dim lRow, arrAssetClassTypeList, arrAssetClassList, arrAssetSubclassList
Dim objRSSelect, strSQL, strWhereClause
Dim intAccessLevel

'security
intAccessLevel = CInt(CheckLogon(strConst_AssetTypeClassification))
If (intAccessLevel and intConst_Access_ReadOnly)<> intConst_Access_ReadOnly then
	DisplayError "BACK","",0,"ACCESS DENIED", "You don't have access to Asset Type. Please contact your system administrator."
end if

'strAssetClassTypeID = Request("AssetClassTypeID")
'strAssetClassID = Request("AssetClassID")
'strAssetSubclassID = Request("AssetSubclassID")
'strAssetTypeID = Request("AssetTypeID")
'strAssetTypeDesc = Request("AssetTypeDesc")

'Create Recordset object  
Set objRSSelect = Server.CreateObject("ADODB.Recordset")

'get the asset class type
strSQL = "SELECT ASSET_CLASS_TYPE_ID, ASSET_CLASS_TYPE_DESC " & _
	"FROM CRP.ASSET_CLASS_TYPE " & _
	"WHERE RECORD_STATUS_IND = 'A' " & _
	"ORDER BY ASSET_CLASS_TYPE_DESC"
On Error Resume Next
objRSSelect.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
If objConn.Errors.Count <> 0 Then
	DisplayError "BACK", "", objConn.Errors(0).NativeError, "ERROR LOADING DATA (Asset Class Type)", objConn.Errors(0).Description
	objConn.Errors.Clear
End If
arrAssetClassTypeList = objRSSelect.GetRows
objRSSelect.Close
	
'get the asset class 
strSQL = "SELECT ASSET_CLASS_ID, ASSET_CLASS_TYPE_ID, ASSET_CLASS_DESC " & _
	"FROM CRP.ASSET_CLASS " & _
	"WHERE RECORD_STATUS_IND = 'A' " & _
	"ORDER BY ASSET_CLASS_DESC"
On Error Resume Next
objRSSelect.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
If objConn.Errors.Count <> 0 Then
	DisplayError "BACK", "", objConn.Errors(0).NativeError, "ERROR LOADING DATA (Asset Class)", objConn.Errors(0).Description
	objConn.Errors.Clear
End If
arrAssetClassList = objRSSelect.GetRows
objRSSelect.Close

'get the asset sub class
strSQL = "SELECT ASSET_SUB_CLASS_ID, ASSET_CLASS_ID, ASSET_SUB_CLASS_DESC " & _
	"FROM CRP.ASSET_SUB_CLASS " & _
	"WHERE RECORD_STATUS_IND = 'A' " & _
	"ORDER BY ASSET_SUB_CLASS_DESC"
On Error Resume Next
objRSSelect.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
If objConn.Errors.Count <> 0 Then
	DisplayError "BACK", "", objConn.Errors(0).NativeError, "ERROR LOADING DATA (Asset Subclass)", objConn.Errors(0).Description
	objConn.Errors.Clear
End If
arrAssetSubclassList = objRSSelect.GetRows
objRSSelect.Close

Set objRSSelect = Nothing

%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<LINK REL="stylesheet" TYPE="text/css" HREF="stylesheets/styles.css">
<SCRIPT type="text/javascript" language="javascript" src="AccessLevels.js"></SCRIPT>
<SCRIPT type="text/javascript" language="javascript" src="GeneralJavaFunctions.js"></SCRIPT>
<SCRIPT type="text/javascript" language="javascript" id="clientEventHandlersJS">
<!-- hide script
	var intAccessLevel = <%=intAccessLevel%>;
	var strAssetClassTypeID = "<%=strAssetClassTypeID%>";
	var strAssetClassID = "<%=strAssetClassID%>";
	var strAssetSubclassID = "<%=strAssetSubclassID%>";
	var strAssetTypeID = "<%=strAssetTypeID%>";
	var strAssetTypeDesc = "<%=strAssetTypeDesc%>";
	
	var arrAssetClassTypeList = new Array();
	var arrAssetClassList = new Array();
	var arrAssetSubclassList = new Array();
	
	function confirm_search(theForm) {
		var bolConfirm;
		if ( (isWhitespace(theForm.selAssetClassType.value)) && (isWhitespace(theForm.selAssetClass.value)) && (isWhitespace(theForm.selAssetSubclass.value)) && (isWhitespace(theForm.txtAssetTypeDesc.value)) ){
			bolConfirm = window.confirm("No search criteria have been entered. This may take a long time...continue?");
			if (bolConfirm){
				true;
			}
			else{
				false;
			}
		}
		true;
	}
	
	function window_onLoad() {
		var intCnt;
		arrAssetClassTypeList[0] = "";
		arrAssetClassList[0] = "";
		arrAssetSubclassList[0] = "";
		
		
		for (intCnt=1; intCnt < document.frmAssetTypeCriteria.selAssetClassType.options.length; intCnt++){
			var oOption = document.frmAssetTypeCriteria.selAssetClassType.options(intCnt);
			//each array element holds Asset_class_type_id|Asset_class_type_desc
			arrAssetClassTypeList[intCnt]=(oOption.value + "|" + oOption.text);
		}
		
		for (intCnt=1; intCnt < document.frmAssetTypeCriteria.selAssetClass.options.length; intCnt++){
			var oOption = document.frmAssetTypeCriteria.selAssetClass.options(intCnt);
			//each array element holds Asset_class_id|Asset_class_desc
			arrAssetClassList[intCnt]=(oOption.value + "|" + oOption.text);
		}
		
		for(intCnt=1;intCnt<document.frmAssetTypeCriteria.selAssetSubclass.options.length;intCnt++){
			var oOption = document.frmAssetTypeCriteria.selAssetSubclass.options(intCnt);
			//each array element holds Asset_sub_class_id|Asset_sub_class_desc
			arrAssetSubclassList[intCnt]=(oOption.value + "|" + oOption.text);
		}
		
		/*if (strAssetClassTypeID != "" || strAssetClassID != "" || strAssetSubclassID != "" || strAssetTypeID != "" || strAssetTypeDesc != "") {
			DeleteCookie("AssetClassTypeID");
			DeleteCookie("AssetClassID");
			DeleteCookie("AssetSubclassID");
			DeleteCookie("AssetTypeDesc");
			DeleteCookie("AssetTypeID");
			document.frmAssetTypeCriteria.submit();
		}*/
	}
	
	function btnClear_onClick(){
		document.frmAssetTypeCriteria.selAssetClassType.selectedIndex = 0;
		//remove all the option tags from the asset class
		for (intCnt = document.frmAssetTypeCriteria.selAssetClass.length-1;intCnt > 0;intCnt--){
			document.frmAssetTypeCriteria.selAssetClass.options.remove(intCnt);
		}
		//add ALL asset classes
		for (intCnt = 1; intCnt < arrAssetClassList.length;intCnt++){
			var strValue = arrAssetClassList[intCnt];
			var arrValue = strValue.split("|");
			var strElement = "<OPTION VALUE='" + arrValue[0] + "'>" + arrValue[2] + "</option>";
			var oOption = document.createElement(strElement);
			document.frmAssetTypeCriteria.selAssetClass.options.add(oOption);
			oOption.innerText = arrValue[2]; //ASSET_CLASS_DESC
		}
		document.frmAssetTypeCriteria.selAssetClass.selectedIndex = 0;
		for (intCnt = document.frmAssetTypeCriteria.selAssetSubclass.length-1;intCnt > 0;intCnt--){
			document.frmAssetTypeCriteria.selAssetSubclass.options.remove(intCnt);
		}
		//add ALL asset sub classes
		for(intCnt = 1; intCnt < arrAssetSubclassList.length;intCnt++){
			var strValue = arrAssetSubclassList[intCnt];
			var arrValue = strValue.split("|");
			var strElement = "<OPTION VALUE='" + arrValue[0] + "'>" + arrValue[2] + "</option>";
			var oOption = document.createElement(strElement);
			document.frmAssetTypeCriteria.selAssetSubclass.options.add(oOption);
			oOption.innerText = arrValue[2]; //ASSET_SUB_CLASS_DESC
		}
		document.frmAssetTypeCriteria.selAssetSubclass.selectedIndex = 0;
		document.frmAssetTypeCriteria.txtAssetTypeDesc.value = "";
	}
	
	function btnNew_onclick(){
		if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {alert('Access denied. Please contact your system administrator.'); return;}
			parent.document.location.href ="AssetTypeDetail.asp?AssetTypeID=NEW";
	}
	
	function fct_onChangeAssetClassType(){
		var intCnt;
		var strAssetClassTypeID;
		
		if(document.frmAssetTypeCriteria.selAssetClassType.selectedIndex != 0){
			strAssetClassTypeID = document.frmAssetTypeCriteria.selAssetClassType.value;
			document.frmAssetTypeCriteria.hdnAssetClassTypeID.value = strAssetClassTypeID;
			
			//remove all the option tags from the asset class
			for (intCnt = document.frmAssetTypeCriteria.selAssetClass.length-1;intCnt > 0;intCnt--){
				document.frmAssetTypeCriteria.selAssetClass.options.remove(intCnt);
			}
			//add asset classes that belong to SELECTED asset class type
			for (intCnt = 1; intCnt < arrAssetClassList.length;intCnt++){
				var strValue = arrAssetClassList[intCnt];
				var arrValue = strValue.split("|");
				if (arrValue[1] == strAssetClassTypeID){
					var strElement = "<OPTION VALUE='" + arrValue[0] + "'>" + arrValue[2] + "</option>";
					var oOption = document.createElement(strElement);
					document.frmAssetTypeCriteria.selAssetClass.options.add(oOption);
					oOption.innerText = arrValue[2]; //ASSET_CLASS_DESC
				}
			}
			
			if((document.frmAssetTypeCriteria.selAssetClass.selectedIndex == 0)&&(document.frmAssetTypeCriteria.selAssetSubclass != 0)){
			//if(document.frmAssetTypeCriteria.selAssetClass.selectedIndex == 0){
				//remove the SELECTED option tags from the asset sub class
				for (intCnt = document.frmAssetTypeCriteria.selAssetSubclass.length-1;intCnt > 0;intCnt--){
					document.frmAssetTypeCriteria.selAssetSubclass.options.remove(intCnt);
				}
				//add ALL asset sub classes
				for(intCnt = 1; intCnt < arrAssetSubclassList.length;intCnt++){
					var strValue = arrAssetSubclassList[intCnt];
					var arrValue = strValue.split("|");
					var strElement = "<OPTION VALUE='" + arrValue[0] + "'>" + arrValue[2] + "</option>";
					var oOption = document.createElement(strElement);
					document.frmAssetTypeCriteria.selAssetSubclass.options.add(oOption);
					oOption.innerText = arrValue[2]; //ASSET_SUB_CLASS_DESC
				}
			}
		}//end if
		else{
			//remove all the option tags from the asset class
			for(intCnt = document.frmAssetTypeCriteria.selAssetClass.length-1;intCnt>0;intCnt--){
				document.frmAssetTypeCriteria.selAssetClass.options.remove(intCnt);
			}
			//add all asset classes
			for(intCnt = 1; intCnt < arrAssetClassList.length;intCnt++){
				var strValue = arrAssetClassList[intCnt];
				var arrValue = strValue.split("|");
				var strElement = "<OPTION VALUE='" + arrValue[0] + "'>" + arrValue[2] + "</option>";
				var oOption = document.createElement(strElement);
				document.frmAssetTypeCriteria.selAssetClass.options.add(oOption);
				oOption.innerText = arrValue[2]; //ASSET_CLASS_DESC
			}	
			//remove all the option tags from the asset sub class
			for(intCnt = document.frmAssetTypeCriteria.selAssetSubclass.length-1;intCnt>0;intCnt--){
				document.frmAssetTypeCriteria.selAssetSubclass.options.remove(intCnt);
			}
			//add all asset subclasses
			for(intCnt = 1; intCnt < arrAssetSubclassList.length;intCnt++){
				var strValue = arrAssetSubclassList[intCnt];
				var arrValue = strValue.split("|");
				var strElement = "<OPTION VALUE='" + arrValue[0] + "'>" + arrValue[2] + "</option>";
				var oOption = document.createElement(strElement);
				document.frmAssetTypeCriteria.selAssetSubclass.options.add(oOption);
				oOption.innerText = arrValue[2]; //ASSET_CLASS_DESC
			}
		}//end else
	}//end fct_onChangeAssetClassType		
	
	function fct_onChangeAssetClass(){
		var intCnt;
		var strAssetClassID
		
		if(document.frmAssetTypeCriteria.selAssetClass.selectedIndex != 0){
			
			strAssetClassID = document.frmAssetTypeCriteria.selAssetClass.value;
			var aValue= strAssetClassID.split("|");
			document.frmAssetTypeCriteria.hdnAssetClassID.value = aValue[0];
			
			//remove all the option tags from the asset sub class
			for (intCnt = document.frmAssetTypeCriteria.selAssetSubclass.length-1;intCnt > 0;intCnt--){
				document.frmAssetTypeCriteria.selAssetSubclass.options.remove(intCnt);
			}
			//add asset subclasses that belong to SELECTED asset sub class 
			for (intCnt = 1; intCnt < arrAssetSubclassList.length;intCnt++){
				var strValue = arrAssetSubclassList[intCnt];
				var arrValue = strValue.split("|");
				if (arrValue[1] == strAssetClassID){
					var strElement = "<OPTION VALUE='" + arrValue[0] + "'>" + arrValue[2] + "</option>";
					var oOption = document.createElement(strElement);
					document.frmAssetTypeCriteria.selAssetSubclass.options.add(oOption);
					oOption.innerText = arrValue[2]; //ASSET_SUB_CLASS_DESC
				}//end 2nd if
			}//end 2nd for
		}//end 1st if
		else{
			//remove all the option tags from the asset sub class
			for(intCnt = document.frmAssetTypeCriteria.selAssetSubclass.length-1;intCnt>0;intCnt--){
				document.frmAssetTypeCriteria.selAssetSubclass.options.remove(intCnt);
			}
			//add all asset subclasses
			for(intCnt = 1; intCnt < arrAssetSubclassList.length;intCnt++){
				var strValue = arrAssetSubclassList[intCnt];
				var arrValue = strValue.split("|");
				var strElement = "<OPTION VALUE='" + arrValue[0] + "'>" + arrValue[2] + "</option>";
				var oOption = document.createElement(strElement);
				document.frmAssetTypeCriteria.selAssetSubclass.options.add(oOption);
				oOption.innerText = arrValue[2]; //ASSET_SUB_CLASS_DESC
			}	
		}//end else
	}//end fct_onChangeAssetClass		
	
	function fct_onChangeAssetSubclass(){
		var intCnt;
		var strAssetSubclassID
		
		if(document.frmAssetTypeCriteria.selAssetSubclass.selectedIndex != 0){
			
			strAssetSubclassID = document.frmAssetTypeCriteria.selAssetSubclass.value;
			var aValue= strAssetSubclassID.split("|");
			document.frmAssetTypeCriteria.hdnAssetSubclassID.value = aValue[0];
			
			if(document.frmAssetTypeCriteria.selAssetClass.selectedIndex == 0){
				//remove all the option tags from the asset class type
				for (intCnt = document.frmAssetTypeCriteria.selAssetClassType.length-1;intCnt > 0;intCnt--){
					document.frmAssetTypeCriteria.selAssetClassType.options.remove(intCnt);
				}
				//add all asset class types
				for(intCnt = 1; intCnt < arrAssetClassTypeList.length;intCnt++){
					var strValue = arrAssetClassTypeList[intCnt];
					var arrValue = strValue.split("|");
					var strElement = "<OPTION VALUE='" + arrValue[0] + "'>" + arrValue[1] + "</option>";
					var oOption = document.createElement(strElement);
					document.frmAssetTypeCriteria.selAssetClassType.options.add(oOption);
					oOption.innerText = arrValue[1]; //ASSET_CLASS_TYPE_DESC
				}
			}
		}
	}//end fct_onChangeAssetClass		
//-->
</script>	
	
</HEAD>
<BODY language="javascript" onLoad="window_onLoad();">
<FORM id="frmAssetTypeCriteria" name="frmAssetTypeCriteria" method="post" action="AssetTypeList.asp" target="fraResult" onSubmit="confirm_search(this)">
<!--hidden variables-->
	<INPUT type=hidden name=hdnAssetClassTypeID value="<%=strAssetClassTypeID%>">
	<INPUT type=hidden name=hdnAssetClassID value="<%=strAssetClassID%>">
	<INPUT type=hidden name=hdnAssetSubclassID value="<%=strAssetSubclassID%>">
	<INPUT type=hidden name=hdnAssetTypeID value="<%=strAssetTypeID%>">
<TABLE cols=4 width=100% border=0>
<THEAD>
	<TR><TD colspan=4 align=left>Asset Type Search</td></tr>
</thead>
<tbody>
	<tr>
		<td align=right width=15%>Asset Class Type</td>
		<td align=left width=40%><select name=selAssetClassType style="width: 365px" onChange="fct_onChangeAssetClassType();">
			<option></option>
			<%for lRow = LBound(arrAssetClassTypeList, 2) to UBound(arrAssetClassTypeList, 2)
				if StrComp(strAssetClassTypeID, arrAssetClassTypeList(0,lRow),0)= 0 then %>
					<OPTION value="<%=arrAssetClassTypeList(0,lRow)%>" selected><%=arrAssetClassTypeList(1,lRow)%>&lt;&gt;OPTION&gt;
				<%else%>
					<OPTION value="<%=arrAssetClassTypeList(0,lRow)%>"><%=arrAssetClassTypeList(1,lRow)%></OPTION>
				<%end if
			next%>
			</select>
		</td>
		<td align=right width=10% >Active Only</td>
		<td align=left><INPUT name=chkActiveOnly type=checkbox checked></td>
	</tr>
	<tr>
		<td align=right  width=15%>Asset Class</td>
		<td align=left width=40%><select name=selAssetClass style="width: 365px" onChange="fct_onChangeAssetClass();">
			<option></option>
			<%for lRow = LBound(arrAssetClassList, 2) to UBound(arrAssetClassList, 2)
				if StrComp(strAssetClassID, arrAssetClassList(0,lRow),0)=0 then %>
					<OPTION value = "<%=arrAssetClassList(0,lRow) & "|" & arrAssetClassList(1,lRow)%>" selected><%=arrAssetClassList(2,lRow)%></option>
				<%else%>
					<OPTION value = "<%=arrAssetClassList(0,lRow) & "|" & arrAssetClassList(1,lRow)%>"><%=arrAssetClassList(2,lRow)%></option>
				<% end if
			next%>
			</select>
		</td>
	</tr>
	<tr>
		<td align=right width=15%>Asset Subclass</td>
		<td align=left width=40%><select name=selAssetSubclass style="width: 365px" onchange="fct_onChangeAssetSubclass();">
			<option></option>
			<%for lRow = LBound(arrAssetSubclassList,2) to UBound(arrAssetSubclassList,2)
				if StrComp(strAssetSubclassID, arrAssetSubclassList(0,lRow),0)=0 then %>
					<OPTION value = "<%=arrAssetSubclassList(0,lRow) & "|" & arrAssetSubclassList(1,lRow)%>" selected><%=arrAssetSubclassList(2,lRow)%></option>
				<%else%>
					<OPTION value = "<%=arrAssetSubclassList(0,lRow) & "|" & arrAssetSubclassList(1,lRow)%>"><%=arrAssetSubclassList(2,lRow)%></option> 
				<% end if
			next%>
			</select>
		</td>
	</tr>
	<tr>
		<td align=right width=15%>Asset Type</td>
		<td align=left width=40% ><INPUT name=txtAssetTypeDesc maxlength=50 size=50 value="<%=strAssetTypeDesc%>"></td>
	</tr>
</tbody>
<tfoot>
	<tr>
		<td colspan=4 align=right>
			<INPUT id=btnNew name=btnNew type=button style="width: 2cm" value=New onclick="return btnNew_onclick()">&nbsp;&nbsp;
			<INPUT name="btnClear" type="button" value="Clear" style="width: 2cm" language="javascript" onClick="btnClear_onClick();">&nbsp;&nbsp;
			<INPUT name="btnSearch" type="submit" value="Search" style="width: 2cm" >&nbsp;&nbsp;
		</td>
	</tr>	
</tfoot>	
</table>
</form>
</BODY>
<%
'clean up ado objects
objConn.Close
Set objConn = Nothing
%>
</HTML>
