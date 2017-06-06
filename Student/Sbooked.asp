<%@LANGUAGE="VBSCRIPT"%>
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
If (CStr(Request("MM_insert")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_room_STRING
    MM_editCmd.CommandText = "INSERT INTO dbo.Wait (Wpers, Wroom, Wdate, Wtime, Wsuage, Wnote) VALUES (?, ?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 50, Request.Form("name")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 50, Request.Form("room")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 202, 1, 50, Request.Form("date")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 201, 1, 50, Request.Form("time")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 202, 1, 50, Request.Form("use")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 202, 1, 50, Request.Form("note")) ' adVarWChar
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If
End If
%>
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_room_STRING
Recordset1_cmd.CommandText = "SELECT * FROM dbo.Wait" 
Recordset1_cmd.Prepared = true

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>确认借用</title>
        <style type="text/css">
    		body {background-image: url(../image/background/background.jpg);}<!--添加背景图片-->
    	</style>
</head>

<body>

<h3 align="center" style="color:#FFF">请填写借用表</h3>
<div align="center" style="color:#FFF">
<form ACTION="<%=MM_editAction%>" name="form1" method="POST">
<table width="200" border="1">
  <tr>
    <td>学号</td>
    <td><input name="name" type="text" id="name" /></td>
  </tr>
  <tr>
    <td>借用教室</td>
    <td><input name="room" type="text" id="room" /></td>
  </tr>
  <tr>
    <td>日期</td>
    <td><input name="date" type="text" id="date" /></td>
  </tr>
  <tr>
    <td>时间</td>
    <td><input name="time" type="text" id="time" /></td>
  </tr>
  <tr>
    <td>事由</td>
    <td><input name="use" type="text" id="use" /></td>
  </tr>
  <tr>
    <td>备注</td>
    <td><input name="note" type="text" id="note" /></td>
  </tr>
  <tr>
    <td><input name="ok" type="submit" id="ok" value="提交" /></td>
    <td><input type="button" value="取消" /></td>
  </tr>
</table>
<input type="hidden" name="MM_insert" value="form1" />
</form>
</div>

<div align="center" style="color:#FFF">
<form action="Sbookeded.asp" method="post" id="2">
请继续输入教室编号<input name="no" type="text" />
<input name="" type="submit" value="继续" />
</form>

</div>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
