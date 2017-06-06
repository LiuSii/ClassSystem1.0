<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../../Connections/room.asp" -->
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
    MM_editCmd.CommandText = "UPDATE dbo.Student SET Sno = ?, Sname = ?, Spassword = ?, Siphone = ? WHERE Sno = ?" 
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
    MM_editRedirectUrl = "../Mpage.asp"
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
' *** Delete Record: construct a sql delete statement and execute it

If (CStr(Request("MM_delete")) = "del" And CStr(Request("MM_recordId")) <> "") Then

  If (Not MM_abortEdit) Then
    ' execute the delete
    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_room_STRING
    MM_editCmd.CommandText = "DELETE FROM dbo.Student WHERE Sno = ?"
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 200, 1, 50, Request.Form("MM_recordId")) ' adVarChar
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
	
    MM_editRedirectUrl = "../Mpage.asp"
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
Dim Mupdatestu__MMColParam
Mupdatestu__MMColParam = "1"
If (Request.Form("no") <> "") Then 
  Mupdatestu__MMColParam = Request.Form("no")
End If
%>
<%
Dim Mupdatestu
Dim Mupdatestu_cmd
Dim Mupdatestu_numRows

Set Mupdatestu_cmd = Server.CreateObject ("ADODB.Command")
Mupdatestu_cmd.ActiveConnection = MM_room_STRING
Mupdatestu_cmd.CommandText = "SELECT * FROM dbo.Student WHERE Sno = ?" 
Mupdatestu_cmd.Prepared = true
Mupdatestu_cmd.Parameters.Append Mupdatestu_cmd.CreateParameter("param1", 200, 1, 50, Mupdatestu__MMColParam) ' adVarChar

Set Mupdatestu = Mupdatestu_cmd.Execute
Mupdatestu_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
		<title>修改学生信息</title>
        <style type="text/css">
    		body {background-image: url(../../image/background/background.jpg);}<!--添加背景图片-->
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
  			<form ACTION="<%=MM_editAction%>" METHOD="POST" name="get" id="get" >
 			<tr>
    			<td style="color:white">用户名</td>
   			  <td><input name="username" type="text" id="username" value="<%=(Mupdatestu.Fields.Item("Sno").Value)%>" size="30" /></td>
 			</tr>
            <tr>
    			<td style="color:white">姓名</td>
    			<td onfocus="MM_validateForm('username','','R','name','','R','password','','R','iphone','','R');return document.MM_returnValue"><input name="name" type="text" id="name" value="<%=(Mupdatestu.Fields.Item("Sname").Value)%>" size="30" /></td>
 			</tr>
            <tr>
    			<td style="color:white">密码</td>
    			<td><input name="password" type="text" id="password" value="<%=(Mupdatestu.Fields.Item("Spassword").Value)%>" size="30" /></td>  
		  	</tr>
            <tr>
    			<td style="color:white">手机</td>
    			<td><input name="iphone" type="text" id="iphone" value="<%=(Mupdatestu.Fields.Item("Siphone").Value)%>" size="30" /></td>  
		  	</tr>
  			<tr>
              <td align="center"><input name="ok" type="submit" value="修改" /></td>
   			  <td align="center"><input name="cancel" type="button" value="取消"/></td>
  			</tr>
            <input type="hidden" name="MM_update" value="get" />
            <input type="hidden" name="MM_recordId" value="<%= Mupdatestu.Fields.Item("Sno").Value %>" />
            </form>
		</table>
        </form>
        
<div align="center">
<form name="del" action="<%=MM_editAction%>" method="POST" id="del">
        <input name="del" type="submit" value="删除" />
        <input type="hidden" name="MM_delete" value="del" />
        <input type="hidden" name="MM_recordId" value="<%= Mupdatestu.Fields.Item("Sno").Value %>" />
</form>
</div>
    
</body>
</html>
<%
Mupdatestu.Close()
Set Mupdatestu = Nothing
%>
