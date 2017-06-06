<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../../../Connections/room.asp" -->
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
    MM_editCmd.CommandText = "INSERT INTO dbo.Logo (Ltimeo, Ldateo, Lroomo) VALUES (?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 201, 1, 50, Request.Form("zero1")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 50, Request.Form("zero2")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 202, 1, 50, Request.Form("no")) ' adVarWChar
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "MUCoverify.asp"
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
Dim Recordset1__MMColParam
Recordset1__MMColParam = "1"
If (Request.Form("noo") <> "") Then 
  Recordset1__MMColParam = Request.Form("noo")
End If
%>
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_room_STRING
Recordset1_cmd.CommandText = "SELECT * FROM dbo.Conference WHERE Cono = ?" 
Recordset1_cmd.Prepared = true
Recordset1_cmd.Parameters.Append Recordset1_cmd.CreateParameter("param1", 200, 1, 50, Recordset1__MMColParam) ' adVarChar

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>会议室上线</title>
        <style type="text/css">
    		body {background-image: url(../../../image/background/background.jpg);}<!--添加背景图片-->
    	</style>
</head>

<body>
<br>
<br>
<br>
<br>
<br>

<div align="center" style="color:#FFF">
<form name="form1" action="<%=MM_editAction%>" method="POST">
<input name="zero1" type="hidden" value="0" />
<input name="zero2" type="hidden" value="0" />
<table width="200" border="1">
  <tr>
    <td><input name="no" type="text" value="<%=(Recordset1.Fields.Item("Cono").Value)%>" /></td>
    <td><input name="capa" type="text" value="<%=(Recordset1.Fields.Item("Cocapa").Value)%>" /></td>
    <td><input name="site" type="text" value="<%=(Recordset1.Fields.Item("Cosite").Value)%>" /></td>
  </tr>
</table>
<input type="submit" value="上线" />
<input type="hidden" name="MM_insert" value="form1" />
</form>
</div>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
