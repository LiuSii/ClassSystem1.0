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
If (CStr(Request("MM_insert")) = "i") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_room_STRING
    MM_editCmd.CommandText = "INSERT INTO dbo.Log (Lpers, Lroom, Ldate, Ltime, Lusage, Lnote) VALUES (?, ?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 50, Request.Form("name")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 50, Request.Form("cla")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 202, 1, 50, Request.Form("date")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 201, 1, 50, Request.Form("time")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 202, 1, 50, Request.Form("thing")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 202, 1, 50, Request.Form("note")) ' adVarWChar
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
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
    MM_editCmd.CommandText = "DELETE FROM dbo.Wait WHERE id = ?"
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, Request.Form("MM_recordId")) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "check.asp"
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
Dim wait_log__MMColParam
wait_log__MMColParam = "1"
If (Request.Form("no") <> "") Then 
  wait_log__MMColParam = Request.Form("no")
End If
%>
<%
Dim wait_log
Dim wait_log_cmd
Dim wait_log_numRows

Set wait_log_cmd = Server.CreateObject ("ADODB.Command")
wait_log_cmd.ActiveConnection = MM_room_STRING
wait_log_cmd.CommandText = "SELECT * FROM dbo.Wait WHERE id = ?" 
wait_log_cmd.Prepared = true
wait_log_cmd.Parameters.Append wait_log_cmd.CreateParameter("param1", 5, 1, -1, wait_log__MMColParam) ' adDouble

Set wait_log = wait_log_cmd.Execute
wait_log_numRows = 0
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>确认成功</title>
        <style type="text/css">
    		body {background-image: url(../image/background/background.jpg);}<!--添加背景图片-->
    	</style>
</head>

<body>

    <h3 style="color:white" align="center">确认事项</h3>
    <div style="width:1200;height:400; overflow:scroll; border:0 solid;overflow-x:auto;overflow-y:auto">
    <form ACTION="<%=MM_editAction%>" method="POST" name="i" id="i">
    <table style="color:#FFF" width="900" border="1" align="center">
  <tr>
    <td>序号</td>
    <td>号码</td>
    <td>教室</td>
    <td>日期</td>
    <td>时间</td>
    <td>事由</td>
    <td>备注</td>
  </tr>
  <tr>
    <td><input name="no" type="text" id="no" value="<%=(wait_log.Fields.Item("id").Value)%>" size="30" /></td>  
    <td><input name="name" type="text" id="name" value="<%=(wait_log.Fields.Item("Wpers").Value)%>" size="30"/></td>
    <td><input name="cla" type="text" id="cla" value="<%=(wait_log.Fields.Item("Wroom").Value)%>" size="30"/></td>
    <td><input name="date" type="text" id="date" value="<%=(wait_log.Fields.Item("Wdate").Value)%>" size="30"/></td>
    <td><input name="time" type="text" id="time" value="<%=(wait_log.Fields.Item("Wtime").Value)%>" size="30"/></td>
    <td><input name="thing" type="text" id="thing" value="<%=(wait_log.Fields.Item("Wsuage").Value)%>" size="30"/></td>
    <td><input name="note" type="text" id="note" value="<%=(wait_log.Fields.Item("Wnote").Value)%>" size="30"/></td>    
  </tr>
    </table>
     <input name="确认" type="submit" id="确认" value="提交" />
     <input type="hidden" name="MM_insert" value="i" />
    </form>
</div>

<div align="center" style="color:#FFF">
<form name="del" action="<%=MM_editAction%>" method="POST" id="del"><input name="" type="submit" value="删除" />
  <input type="hidden" name="MM_delete" value="del" />
  <input type="hidden" name="MM_recordId" value="<%= wait_log.Fields.Item("id").Value %>" />
</form>
</div>

</body>
</html>
<%
wait_log.Close()
Set wait_log = Nothing
%>