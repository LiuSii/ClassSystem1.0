<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/room.asp" -->
<%
Dim stu__MMColParam
stu__MMColParam = "1"
If (Session("MM_Username") <> "") Then 
  stu__MMColParam = Session("MM_Username")
End If
%>
<%
Dim stu
Dim stu_cmd
Dim stu_numRows

Set stu_cmd = Server.CreateObject ("ADODB.Command")
stu_cmd.ActiveConnection = MM_room_STRING
stu_cmd.CommandText = "SELECT * FROM dbo.Student WHERE Sno = ?" 
stu_cmd.Prepared = true
stu_cmd.Parameters.Append stu_cmd.CreateParameter("param1", 200, 1, 50, stu__MMColParam) ' adVarChar

Set stu = stu_cmd.Execute
stu_numRows = 0
%>
<%
Dim stulog__MMColParam
stulog__MMColParam = "1"
If (Session("MM_Username") <> "") Then 
  stulog__MMColParam = Session("MM_Username")
End If
%>
<%
Dim stulog
Dim stulog_cmd
Dim stulog_numRows

Set stulog_cmd = Server.CreateObject ("ADODB.Command")
stulog_cmd.ActiveConnection = MM_room_STRING
stulog_cmd.CommandText = "SELECT * FROM dbo.Log WHERE Lpers = ?" 
stulog_cmd.Prepared = true
stulog_cmd.Parameters.Append stulog_cmd.CreateParameter("param1", 200, 1, 50, stulog__MMColParam) ' adVarChar

Set stulog = stulog_cmd.Execute
stulog_numRows = 0
%>
<%
Dim Recordset1__MMColParam
Recordset1__MMColParam = "1"
If (Session("MM_Username") <> "") Then 
  Recordset1__MMColParam = Session("MM_Username")
End If
%>
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_room_STRING
Recordset1_cmd.CommandText = "SELECT * FROM dbo.Wait WHERE Wpers = ?" 
Recordset1_cmd.Prepared = true
Recordset1_cmd.Parameters.Append Recordset1_cmd.CreateParameter("param1", 200, 1, 50, Recordset1__MMColParam) ' adVarChar

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
stulog_numRows = stulog_numRows + Repeat1__numRows
%>
<%
Dim Repeat2__numRows
Dim Repeat2__index

Repeat2__numRows = -1
Repeat2__index = 0
Recordset1_numRows = Recordset1_numRows + Repeat2__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
		<title>学生主页</title>
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
    			<td style="color:white" width="200">学号</td>
    			<td style="color:white" width="300">姓名</td>
    			<td style="color:white" width="300">手机</td>
                <td style="color:white" width="100">修改</td>
  			</tr>
  			<tr>
    			<td style="color:white" width="200"><%=(stu.Fields.Item("Sno").Value)%></td>
    			<td style="color:white" width="300"><%=(stu.Fields.Item("Sname").Value)%></td>
    			<td style="color:white" width="300"><%=(stu.Fields.Item("Siphone").Value)%></td>
                <td width="100"><a href="Supdate.asp">修改</a></td>
		  </tr>
		</table>
      </form>
	</div>    
	<br>
    <br>

    <h3 style="color:white" align="center">已借教室</h3>
    <div style="width:1200;height:400; overflow:scroll; border:0 solid;overflow-x:auto;overflow-y:auto">
    <form action="" method="get">
	  <table width="900" border="1" align="center">
	    <tr>
    			<td style="color:white" width="200">教室</td>
    			<td style="color:white" width="300">日期</td>
    			<td style="color:white" width="300">时间</td>
                <td style="color:white" width="300">事由</td>
		  </tr>
        <% 
While ((Repeat1__numRows <> 0) AND (NOT stulog.EOF)) 
%>
  <tr>
    <td style="color:white" width="200"><%=(stulog.Fields.Item("Lroom").Value)%></td>
    <td style="color:white" width="300"><%=(stulog.Fields.Item("Ldate").Value)%></td>
    <td style="color:white" width="300"><%=(stulog.Fields.Item("Ltime").Value)%></td>
    <td style="color:white" width="300"><%=(stulog.Fields.Item("Lusage").Value)%></td>
  </tr>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  stulog.MoveNext()
Wend
%>
      </table>
      </form>
	</div>    
    <br>
    <br>
        <h3 style="color:white" align="center">待审核教室</h3>
    <div style="width:1200;height:400; overflow:scroll; border:0 solid;overflow-x:auto;overflow-y:auto">
    <form action="" method="get">
	  <table width="900" border="1" align="center">
	    <tr>
    			<td style="color:white" width="200">教室</td>
    			<td style="color:white" width="300">日期</td>
    			<td style="color:white" width="300">时间</td>
                <td style="color:white" width="300">事由</td>
		  </tr>
        <% 
While ((Repeat2__numRows <> 0) AND (NOT Recordset1.EOF)) 
%>
  <tr>
    <td style="color:white" width="200"><%=(Recordset1.Fields.Item("Wroom").Value)%></td>
    <td style="color:white" width="300"><%=(Recordset1.Fields.Item("Wdate").Value)%></td>
    <td style="color:white" width="300"><%=(Recordset1.Fields.Item("Wtime").Value)%></td>
    <td style="color:white" width="300"><%=(Recordset1.Fields.Item("Wsuage").Value)%></td>
  </tr>
  <% 
  Repeat2__index=Repeat2__index+1
  Repeat2__numRows=Repeat2__numRows-1
  Recordset1.MoveNext()
Wend
%>
      </table>
      </form>
	</div>    
    
	<br>
    <br>
	<div style="color:#FFF" align="center">
    <form action="Sbook.asp" method="post" id="dt">请选择日期&nbsp;&nbsp;<input name="date" type="text" id="Calendar3" onFocus="StartCalendar({id:'Calendar3',lunarShow:true});" value="日期未选择" size="24">
    <br>
    <br>
请输入节数（1~6）&nbsp;&nbsp;<input name="time" type="text" value="请输入节数" id="time"/>
<br>
<br>
<input type="submit" value="提交" />
</form></div>

	<br>
    
</body>
</html>
<%
stu.Close()
Set stu = Nothing
%>
<%
stulog.Close()
Set stulog = Nothing
%>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
