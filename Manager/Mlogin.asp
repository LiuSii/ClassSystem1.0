<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/room.asp" -->
<%
Dim managerlogin
Dim managerlogin_cmd
Dim managerlogin_numRows

Set managerlogin_cmd = Server.CreateObject ("ADODB.Command")
managerlogin_cmd.ActiveConnection = MM_room_STRING
managerlogin_cmd.CommandText = "SELECT * FROM dbo.Manager" 
managerlogin_cmd.Prepared = true

Set managerlogin = managerlogin_cmd.Execute
managerlogin_numRows = 0
%>
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString <> "" Then MM_LoginAction = MM_LoginAction + "?" + Server.HTMLEncode(Request.QueryString)
MM_valUsername = CStr(Request.Form("username"))
If MM_valUsername <> "" Then
  Dim MM_fldUserAuthorization
  Dim MM_redirectLoginSuccess
  Dim MM_redirectLoginFailed
  Dim MM_loginSQL
  Dim MM_rsUser
  Dim MM_rsUser_cmd
  
  MM_fldUserAuthorization = ""
  MM_redirectLoginSuccess = "Mpage.asp"
  MM_redirectLoginFailed = "Mlogin.asp"

  MM_loginSQL = "SELECT Mno, Mpassword"
  If MM_fldUserAuthorization <> "" Then MM_loginSQL = MM_loginSQL & "," & MM_fldUserAuthorization
  MM_loginSQL = MM_loginSQL & " FROM dbo.Manager WHERE Mno = ? AND Mpassword = ?"
  Set MM_rsUser_cmd = Server.CreateObject ("ADODB.Command")
  MM_rsUser_cmd.ActiveConnection = MM_room_STRING
  MM_rsUser_cmd.CommandText = MM_loginSQL
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param1", 200, 1, 50, MM_valUsername) ' adVarChar
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param2", 200, 1, 50, Request.Form("password")) ' adVarChar
  MM_rsUser_cmd.Prepared = true
  Set MM_rsUser = MM_rsUser_cmd.Execute

  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then 
    ' username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername
    If (MM_fldUserAuthorization <> "") Then
      Session("MM_UserAuthorization") = CStr(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value)
    Else
      Session("MM_UserAuthorization") = ""
    End If
    if CStr(Request.QueryString("accessdenied")) <> "" And false Then
      MM_redirectLoginSuccess = Request.QueryString("accessdenied")
    End If
    MM_rsUser.Close
    Response.Redirect(MM_redirectLoginSuccess)
  End If
  MM_rsUser.Close
  Response.Redirect(MM_redirectLoginFailed)
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
		<title>管理员登陆</title>
        <style type="text/css">
    		body {background-image: url(../image/background/background.jpg);}<!--添加背景图片-->
    	</style>
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
    	<form action="<%=MM_LoginAction%>" method="POST" id="getSlogin">
 			<tr>
    			<td style="color:white">用户名</td>
    			<td><input name="username" type="text" id="username" onfocus="MM_validateForm('username','','R','password','','R');return document.MM_returnValue" size="30" /></td>
 			</tr>
            <tr>
    			<td style="color:white">密码</td>
    			<td><input name="password" type="password" id="password" size="30" /></td>  
		  	</tr>
  			<tr>
              <td align="center"><input name="ok" type="submit" value="提交" /></td>
   			  <td align="center"><a href="../FirstPage.asp">取消</a></td>
  			</tr>
           </form>
	</table>  
    
	</body>
</html>
<%
managerlogin.Close()
Set managerlogin = Nothing
%>
