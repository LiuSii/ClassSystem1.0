<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/room.asp" -->
<%
Dim MM_editAction
MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
Dim MM_abortEdit
MM_abortEdit = false
%>
<%
If (CStr(Request("MM_update")) = "get") Then
  If (Not MM_abortEdit) Then
    ' execute the update
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_room_STRING
    MM_editCmd.CommandText = "UPDATE dbo.Manager SET Mno = ?, Mname = ?, Mpassword = ?, Miphone = ? WHERE Mno = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 201, 1, 50, Request.Form("username")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 50, Request.Form("name")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 201, 1, 50, Request.Form("password")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 201, 1, 50, Request.Form("iphone")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 200, 1, 50, Request.Form("MM_recordId")) ' adVarChar
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "Mpage.asp"
    If (Request.QueryString <> "") Then
      If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0) Then
        MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
      Else
        MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
      End If
    End If
    Response.Redirect(MM_editRedirectUrl)
  End If
End If
%>
<%
Dim Mupdate__MMColParam
Mupdate__MMColParam = "1"
If (Session("MM_Username") <> "") Then 
  Mupdate__MMColParam = Session("MM_Username")
End If
%>
<%
Dim Mupdate
Dim Mupdate_cmd
Dim Mupdate_numRows

Set Mupdate_cmd = Server.CreateObject ("ADODB.Command")
Mupdate_cmd.ActiveConnection = MM_room_STRING
Mupdate_cmd.CommandText = "SELECT * FROM dbo.Manager WHERE Mno = ?" 
Mupdate_cmd.Prepared = true
Mupdate_cmd.Parameters.Append Mupdate_cmd.CreateParameter("param1", 200, 1, 50, Mupdate__MMColParam) ' adVarChar

Set Mupdate = Mupdate_cmd.Execute
Mupdate_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
		<title>修改信息</title>
        <style type="text/css">
    		body {background-image: url(../image/background/background.jpg);}<!--添加背景图片-->
    	</style>
<script type="text/javascript">
function MM_validateForm() { //v4.0
  if (document.getElementById){
    var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;
    for (i=0; i<(args.length-2); i+=3) { test=args[i+2]; val=document.getElementById(args[i]);
      if (val) { nm=val.name; if ((val=val.value)!="") {
        if (test.indexOf('isEmail')!=-1) { p=val.indexOf('@');
          if (p<1 || p==(val.length-1)) errors+='- '+nm+' must contain an e-mail address.\n';
        } else if (test!='R') { num = parseFloat(val);
          if (isNaN(val)) errors+='- '+nm+' must contain a number.\n';
          if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
            min=test.substring(8,p); max=test.substring(p+1);
            if (num<min || max<num) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
      } } } else if (test.charAt(0) == 'R') errors += '- '+nm+' is required.\n'; }
    } if (errors) alert('The following error(s) occurred:\n'+errors);
    document.MM_returnValue = (errors == '');
} }
        </script>
</head>

<body>
    
    <br>
    <br>
    <br>
    <br>
    <br>
    <br>
    <br>
    <br>
    <br>
    
   		<table width="400" border="0" align="center">
  			<form ACTION="<%=MM_editAction%>" METHOD="POST" name="get" id="get">
 			<tr>
    			<td style="color:white">用户名</td>
   			  <td><input name="username" type="text" id="username" value="<%=(Mupdate.Fields.Item("Mno").Value)%>" size="30" /></td>
 			</tr>
            <tr>
    			<td style="color:white" onfocus="MM_validateForm('username','','R','name','','R','password','','R','iphone','','R');return document.MM_returnValue">姓名</td>
   			  <td><input name="name" type="text" id="name" value="<%=(Mupdate.Fields.Item("Mname").Value)%>" size="30" /></td>
 			</tr>
            <tr>
    			<td style="color:white">密码</td>
    			<td><input name="password" type="text" id="password" value="<%=(Mupdate.Fields.Item("Mpassword").Value)%>" size="30" /></td>  
		  	</tr>
            <tr>
    			<td style="color:white">手机</td>
    			<td><input name="iphone" type="text" id="iphone" value="<%=(Mupdate.Fields.Item("Miphone").Value)%>" size="30" /></td>  
		  	</tr>
  			<tr>
              <td align="center"><input name="ok" type="submit" value="修改" /></td>
   			  <td align="center"><input name="cancel" type="button" value="取消"/></td>
  			</tr>
            <input type="hidden" name="MM_insert" value="get" />
            <input type="hidden" name="MM_update" value="get" />
            <input type="hidden" name="MM_recordId" value="<%= Mupdate.Fields.Item("Mno").Value %>" />
            </form>
		</table>
</body>
</html>
<%
Mupdate.Close()
Set Mupdate = Nothing
%>
