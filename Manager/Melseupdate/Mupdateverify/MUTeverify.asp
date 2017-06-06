<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../../../Connections/room.asp" -->
<%
Dim Mshowtea
Dim Mshowtea_cmd
Dim Mshowtea_numRows

Set Mshowtea_cmd = Server.CreateObject ("ADODB.Command")
Mshowtea_cmd.ActiveConnection = MM_room_STRING
Mshowtea_cmd.CommandText = "SELECT * FROM dbo.Teacher" 
Mshowtea_cmd.Prepared = true

Set Mshowtea = Mshowtea_cmd.Execute
Mshowtea_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
Mshowtea_numRows = Mshowtea_numRows + Repeat1__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>确认教师</title>
        <style type="text/css">
    		body {background-image: url(../../../image/background/background.jpg);}<!--添加背景图片-->
    	</style>
</head>

<body>
    <br>
    <br>
    <br>

   <h3 style="color:white" align="center">教师信息</h3>
    <div style="width:1200;height:400; overflow:scroll; border:0 solid;overflow-x:auto;overflow-y:auto">
    <form action="" method="Tget">
	  <table width="900" border="1" align="center">
		<tr>
    			<td style="color:white" width="200">工号</td>
		  <td style="color:white" width="300">姓名</td>
    			<td style="color:white" width="300">手机</td>               
		  </tr>
        <% 
While ((Repeat1__numRows <> 0) AND (NOT Mshowtea.EOF)) 
%>
  <tr>
    <td style="color:white" width="200"><%=(Mshowtea.Fields.Item("Tno").Value)%></td>
    <td style="color:white" width="300"><%=(Mshowtea.Fields.Item("Tname").Value)%></td>
    <td style="color:white" width="300"><%=(Mshowtea.Fields.Item("Tiphone").Value)%></td>
  </tr>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Mshowtea.MoveNext()
Wend
%>
      </table>
      </form>
	</div>
    <br>
    <br>
    <br>
    <h3 align="center" style="color:#FFF">请确认要修改的教师工号<h3>
    <div align="center">
    	<form action="../Mteaupdate.asp" method="post"><input name="no" type="text" id="no" />
       		 <br>
             <br>
       		 <input name="" type="submit" value="提交" />
        </form>
    </div>
</body>
</html>
<%
Mshowtea.Close()
Set Mshowtea = Nothing
%>
