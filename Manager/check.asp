<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/room.asp" -->
<%
Dim Mcheckw
Dim Mcheckw_cmd
Dim Mcheckw_numRows

Set Mcheckw_cmd = Server.CreateObject ("ADODB.Command")
Mcheckw_cmd.ActiveConnection = MM_room_STRING
Mcheckw_cmd.CommandText = "SELECT * FROM dbo.Wait" 
Mcheckw_cmd.Prepared = true

Set Mcheckw = Mcheckw_cmd.Execute
Mcheckw_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
Mcheckw_numRows = Mcheckw_numRows + Repeat1__numRows
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
While ((Repeat1__numRows <> 0) AND (NOT Mcheckw.EOF)) 
%>
  <tr>
    <td><%=(Mcheckw.Fields.Item("id").Value)%></td>  
    <td><%=(Mcheckw.Fields.Item("Wpers").Value)%></td>
    <td><%=(Mcheckw.Fields.Item("Wroom").Value)%></td>
    <td><%=(Mcheckw.Fields.Item("Wdate").Value)%></td>
    <td><%=(Mcheckw.Fields.Item("Wtime").Value)%></td>
    <td><%=(Mcheckw.Fields.Item("Wsuage").Value)%></td>
    <td><%=(Mcheckw.Fields.Item("Wnote").Value)%></td>    
  </tr>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Mcheckw.MoveNext()
Wend
%>
    </table>
</form>
</div>

<div style="color:#FFF" align="center">
    <form action="checked.asp" method="post" id="dt">
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
Mcheckw.Close()
Set Mcheckw = Nothing
%>
