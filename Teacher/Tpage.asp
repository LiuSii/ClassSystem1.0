<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/room.asp" -->
<%
Dim tea__MMColParam
tea__MMColParam = "1"
If (Session("MM_Username") <> "") Then 
  tea__MMColParam = Session("MM_Username")
End If
%>
<%
Dim tea
Dim tea_cmd
Dim tea_numRows

Set tea_cmd = Server.CreateObject ("ADODB.Command")
tea_cmd.ActiveConnection = MM_room_STRING
tea_cmd.CommandText = "SELECT * FROM dbo.Teacher WHERE Tno = ?" 
tea_cmd.Prepared = true
tea_cmd.Parameters.Append tea_cmd.CreateParameter("param1", 200, 1, 50, tea__MMColParam) ' adVarChar

Set tea = tea_cmd.Execute
tea_numRows = 0
%>
<%
Dim tealogo__MMColParam
tealogo__MMColParam = "1"
If (Session("MM_Username") <> "") Then 
  tealogo__MMColParam = Session("MM_Username")
End If
%>
<%
Dim tealogo
Dim tealogo_cmd
Dim tealogo_numRows

Set tealogo_cmd = Server.CreateObject ("ADODB.Command")
tealogo_cmd.ActiveConnection = MM_room_STRING
tealogo_cmd.CommandText = "SELECT * FROM dbo.Logo WHERE Lperso = ?" 
tealogo_cmd.Prepared = true
tealogo_cmd.Parameters.Append tealogo_cmd.CreateParameter("param1", 200, 1, 50, tealogo__MMColParam) ' adVarChar

Set tealogo = tealogo_cmd.Execute
tealogo_numRows = 0
%>
<%
Dim teawaito__MMColParam
teawaito__MMColParam = "1"
If (Session("MM_Username") <> "") Then 
  teawaito__MMColParam = Session("MM_Username")
End If
%>
<%
Dim teawaito
Dim teawaito_cmd
Dim teawaito_numRows

Set teawaito_cmd = Server.CreateObject ("ADODB.Command")
teawaito_cmd.ActiveConnection = MM_room_STRING
teawaito_cmd.CommandText = "SELECT * FROM dbo.Waito WHERE Wperso = ?" 
teawaito_cmd.Prepared = true
teawaito_cmd.Parameters.Append teawaito_cmd.CreateParameter("param1", 200, 1, 50, teawaito__MMColParam) ' adVarChar

Set teawaito = teawaito_cmd.Execute
teawaito_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
tealogo_numRows = tealogo_numRows + Repeat1__numRows
%>
<%
Dim Repeat2__numRows
Dim Repeat2__index

Repeat2__numRows = -1
Repeat2__index = 0
teawaito_numRows = teawaito_numRows + Repeat2__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
		<title>教师主页</title>
        <style type="text/css">
    		body {background-image: url(../image/background/background.jpg);}<!--添加背景图片-->
    	</style>        
        <script src="KingCalendar/jquery_1_7.js" type="text/javascript"></script>
		<script type="text/javascript" src="KingCalendar/King-Calendar.js"></script>
	</head>

	<body>
    <h3 style="color:white" align="center">本人信息</h3>
    <div style="width:1200;height:400; overflow:scroll; border:0 solid;overflow-x:auto;overflow-y:auto">
    <form action="" method="get">
		<table width="900" border="1" align="center">
  			<tr>
    			<td style="color:white" width="200">工号</td>
    			<td style="color:white" width="300">姓名</td>
    			<td style="color:white" width="300">手机</td>
                <td style="color:white" width="100">修改</td>
  			</tr>
  			<tr>
    			<td style="color:white" width="200"><%=(tea.Fields.Item("Tno").Value)%></td>
    			<td style="color:white" width="300"><%=(tea.Fields.Item("Tname").Value)%></td>
    			<td style="color:white" width="300"><%=(tea.Fields.Item("Tiphone").Value)%></td>
                <td width="100"><a href="Tupdate.asp">修改</a></td>
  			</tr>
		</table>
      </form>
	</div>    
    
    <h3 style="color:white" align="center">已借会议室</h3>
    <div style="width:1200;height:400; overflow:scroll; border:0 solid;overflow-x:auto;overflow-y:auto">
    <form action="" method="get">
	  <table width="900" border="1" align="center">
	    <tr>
    			<td style="color:white" width="200">会议室</td>
    			<td style="color:white" width="300">日期</td>
    			<td style="color:white" width="300">时间</td>
                <td style="color:white" width="300">事由</td>
	    </tr>
        <% 
While ((Repeat1__numRows <> 0) AND (NOT tealogo.EOF)) 
%>
        <tr>
          <td style="color:white" width="200"><%=(tealogo.Fields.Item("Lroomo").Value)%></td>
          <td style="color:white" width="300"><%=(tealogo.Fields.Item("Ldateo").Value)%></td>
          <td style="color:white" width="300"><%=(tealogo.Fields.Item("Ltimeo").Value)%></td>
          <td style="color:white" width="300"><%=(tealogo.Fields.Item("Lsuageo").Value)%></td>
        </tr>
          <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  tealogo.MoveNext()
Wend
%>
      </table>
      </form>
	</div>    
    <br>
    <br>
        <h3 style="color:white" align="center">待审会议室</h3>
    <div style="width:1200;height:400; overflow:scroll; border:0 solid;overflow-x:auto;overflow-y:auto">
    <form action="" method="get">
	  <table width="900" border="1" align="center">
	    <tr>
    			<td style="color:white" width="200">会议室</td>
    			<td style="color:white" width="300">日期</td>
    			<td style="color:white" width="300">时间</td>
                <td style="color:white" width="300">事由</td>
	    </tr>
        <% 
While ((Repeat2__numRows <> 0) AND (NOT teawaito.EOF)) 
%>
        <tr>
          <td style="color:white" width="200"><%=(teawaito.Fields.Item("Wroomo").Value)%></td>
          <td style="color:white" width="300"><%=(teawaito.Fields.Item("Wdateo").Value)%></td>
          <td style="color:white" width="300"><%=(teawaito.Fields.Item("Wtimeo").Value)%></td>
          <td style="color:white" width="300"><%=(teawaito.Fields.Item("Wsuageo").Value)%></td>
        </tr>
          <% 
  Repeat2__index=Repeat2__index+1
  Repeat2__numRows=Repeat2__numRows-1
  teawaito.MoveNext()
Wend
%>
      </table>
      </form>
	</div>    
    
	<br>
    <br>
	<div style="color:#FFF" align="center">
    <form action="Tbook.asp" method="post" id="dt">请选择日期&nbsp;&nbsp;<input name="date" type="text" id="Calendar3" onFocus="StartCalendar({id:'Calendar3',lunarShow:true});" value="日期未选择" size="24">
    <br>
    <br>
请输入节数（1~6）&nbsp;&nbsp;<input name="time" type="text" value="请输入节数" id="time"/>
<br>
<br>
<input type="submit" value="提交" />
</form></div>
</body>
</html>
<%
tea.Close()
Set tea = Nothing
%>
<%
tealogo.Close()
Set tealogo = Nothing
%>
<%
teawaito.Close()
Set teawaito = Nothing
%>
