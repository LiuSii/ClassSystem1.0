<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/room.asp" -->
<%
Dim Mpage__MMColParam
Mpage__MMColParam = "1"
If (Session("MM_Username") <> "") Then 
  Mpage__MMColParam = Session("MM_Username")
End If
%>
<%
Dim Mpage
Dim Mpage_cmd
Dim Mpage_numRows

Set Mpage_cmd = Server.CreateObject ("ADODB.Command")
Mpage_cmd.ActiveConnection = MM_room_STRING
Mpage_cmd.CommandText = "SELECT * FROM dbo.Manager WHERE Mno = ?" 
Mpage_cmd.Prepared = true
Mpage_cmd.Parameters.Append Mpage_cmd.CreateParameter("param1", 200, 1, 50, Mpage__MMColParam) ' adVarChar

Set Mpage = Mpage_cmd.Execute
Mpage_numRows = 0
%>
<%
Dim Mstudent
Dim Mstudent_cmd
Dim Mstudent_numRows

Set Mstudent_cmd = Server.CreateObject ("ADODB.Command")
Mstudent_cmd.ActiveConnection = MM_room_STRING
Mstudent_cmd.CommandText = "SELECT * FROM dbo.Student" 
Mstudent_cmd.Prepared = true

Set Mstudent = Mstudent_cmd.Execute
Mstudent_numRows = 0
%>
<%
Dim Mteacher
Dim Mteacher_cmd
Dim Mteacher_numRows

Set Mteacher_cmd = Server.CreateObject ("ADODB.Command")
Mteacher_cmd.ActiveConnection = MM_room_STRING
Mteacher_cmd.CommandText = "SELECT * FROM dbo.Teacher" 
Mteacher_cmd.Prepared = true

Set Mteacher = Mteacher_cmd.Execute
Mteacher_numRows = 0
%>
<%
Dim Mclassroom
Dim Mclassroom_cmd
Dim Mclassroom_numRows

Set Mclassroom_cmd = Server.CreateObject ("ADODB.Command")
Mclassroom_cmd.ActiveConnection = MM_room_STRING
Mclassroom_cmd.CommandText = "SELECT * FROM dbo.Classroom" 
Mclassroom_cmd.Prepared = true

Set Mclassroom = Mclassroom_cmd.Execute
Mclassroom_numRows = 0
%>
<%
Dim Mconference
Dim Mconference_cmd
Dim Mconference_numRows

Set Mconference_cmd = Server.CreateObject ("ADODB.Command")
Mconference_cmd.ActiveConnection = MM_room_STRING
Mconference_cmd.CommandText = "SELECT * FROM dbo.Conference" 
Mconference_cmd.Prepared = true

Set Mconference = Mconference_cmd.Execute
Mconference_numRows = 0
%>
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_room_STRING
Recordset1_cmd.CommandText = "SELECT * FROM dbo.action" 
Recordset1_cmd.Prepared = true

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>

<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
Mstudent_numRows = Mstudent_numRows + Repeat1__numRows
%>
<%
Dim Repeat2__numRows
Dim Repeat2__index

Repeat2__numRows = -1
Repeat2__index = 0
Mteacher_numRows = Mteacher_numRows + Repeat2__numRows
%>
<%
Dim Repeat4__numRows
Dim Repeat4__index

Repeat4__numRows = -1
Repeat4__index = 0
Mclassroom_numRows = Mclassroom_numRows + Repeat4__numRows
%>
<%
Dim Repeat3__numRows
Dim Repeat3__index

Repeat3__numRows = -1
Repeat3__index = 0
Mconference_numRows = Mconference_numRows + Repeat3__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
		<title>管理员主页</title>
        <style type="text/css">
    		body {background-image: url(../image/background/background.jpg);}<!--添加背景图片-->
    	</style>
	</head>

	<body>
    
    <h3 style="color:white" align="center">个人信息</h3>
    <p style="color:white" align="center"><a href="Madd/Maddman.asp">添加管理员账户</a></p>
    <div style="width:1200;height:400; overflow:scroll; border:0 solid;overflow-x:auto;overflow-y:auto">
    <form action="" method="Sget">
		<table width="900" border="1" align="center">
  			<tr>
    			<td style="color:white" width="200">账号</td>
   			  <td style="color:white" width="300">姓名</td>
    			<td style="color:white" width="300">手机</td>
                <td style="color:white" width="100">管理</td>
  			</tr>
  			<tr>
   			  <td style="color:white" width="200"><%=(Mpage.Fields.Item("Mno").Value)%></td>
   			  <td style="color:white" width="300"><%=(Mpage.Fields.Item("Mname").Value)%></td>
   			  <td style="color:white" width="300"><%=(Mpage.Fields.Item("Miphone").Value)%></td>
                <td width="100"><a href="Mupdate.asp">修改</a></td>
  			</tr>
		</table>
      </form>
	</div>
    
    <h3 style="color:white" align="center">待审核教室</h3>
    <p align="center" style="color:#FFF"><a href="check.asp">进入查看</a></p>
    
    <h3 style="color:white" align="center">待审核会议室</h3>
    <p align="center" style="color:#FFF"><a href="checkcon.asp">进入查看</a></p>
    
    
    <h3 style="color:white" align="center">学生信息</h3>
    <p style="color:white" align="center"><a href="Madd/MaddStu.asp">新建</a>&nbsp;&nbsp;&nbsp;<a href="Melseupdate/Mupdateverify/MUStverify.asp">修改</a></p>
    
    <br>
        
    <h3 style="color:white" align="center">教室信息</h3>
    <p style="color:white" align="center"><a href="Madd/Maddcla.asp">新建</a>&nbsp;&nbsp;&nbsp;<a href="Melseupdate/Mupdateverify/MUClverify.asp">修改</a></p>
    
    <br>
        
    <h3 style="color:white" align="center">会议室信息</h3>
    <p style="color:white" align="center"><a href="Madd/Maddcon.asp">新建</a>&nbsp;&nbsp;&nbsp;<a href="Melseupdate/Mupdateverify/MUCoverify.asp">修改</a></p>
    
    <br>
    
    <h3 style="color:white" align="center">教师信息</h3>
    <p style="color:white" align="center"><a href="Madd/MaddTea.asp">新建</a>&nbsp;&nbsp;&nbsp;<a href="Melseupdate/Mupdateverify/MUTeverify.asp">修改</a></p>
    
    <div style="width:1200;height:400; overflow:scroll; border:0 solid;overflow-x:auto;overflow-y:auto">
        <h3 style="color:white" align="center">日志信息</h3>
   <table width="900" border="1" align="center" style="color:#FFF">
  <tr>
    <td>序号</td>
    <td>用户</td>
    <td>操作</td>
    <td>时间</td>
  </tr>
  <% 
While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) 
%>
    <tr>
      <td><%=(Recordset1.Fields.Item("id").Value)%></td>
      <td><%=(Recordset1.Fields.Item("per").Value)%></td>
      <td><%=(Recordset1.Fields.Item("action").Value)%></td>
      <td><%=(Recordset1.Fields.Item("time").Value)%></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%>
   </table>


	</div>
    
</body>
</html>
<%
Mpage.Close()
Set Mpage = Nothing
%>
<%
Mstudent.Close()
Set Mstudent = Nothing
%>
<%
Mteacher.Close()
Set Mteacher = Nothing
%>
<%
Mclassroom.Close()
Set Mclassroom = Nothing
%>
<%
Mconference.Close()
Set Mconference = Nothing
%>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
