<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/room.asp" -->
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_room_STRING
Recordset1_cmd.CommandText = "SELECT * FROM dbo.Waito" 
Recordset1_cmd.Prepared = true

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
Recordset1_numRows = Recordset1_numRows + Repeat1__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>审核事项</title>
        <style type="text/css">
    		body {background-image: url(../image/background/background.jpg);}<!--添加背景图片-->
    	</style>
</head>

<body>
    <h3 style="color:white" align="center">待审核事项</h3>
    <div style="width:1200;height:400; overflow:scroll; border:0 solid;overflow-x:auto;overflow-y:auto">
    <form action="" method="post">
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
  <% 
While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) 
%>
  <tr>
    <td><%=(Recordset1.Fields.Item("id").Value)%></td>
    <td><%=(Recordset1.Fields.Item("Wperso").Value)%></td>
    <td><%=(Recordset1.Fields.Item("Wroomo").Value)%></td>
    <td><%=(Recordset1.Fields.Item("Wdateo").Value)%></td>
    <td><%=(Recordset1.Fields.Item("Wtimeo").Value)%></td>
    <td><%=(Recordset1.Fields.Item("Wsuageo").Value)%></td>
    <td><%=(Recordset1.Fields.Item("Wnoteo").Value)%></td>
  </tr>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%>
    </table>
</form>
</div>

<div style="color:#FFF" align="center">
    <form action="checkedcon.asp" method="post" id="dt">
    <br>
    <br>
请输入要确认的事项序号&nbsp;&nbsp;<input name="no" type="text" value="" id="no"/>
<br>
<br>
<input type="submit" value="提交" />
</form></div>

</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
