<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/room.asp" -->
<%
Dim Recordset1__MMColParam
Recordset1__MMColParam = "0"
If (Request.Form("time") <> "") Then 
  Recordset1__MMColParam = Request.Form("time")
End If
%>
<%
Dim Recordset1__MMColParam2
Recordset1__MMColParam2 = "0"
If (Request.Form("date") <> "") Then 
  Recordset1__MMColParam2 = Request.Form("date")
End If
%>
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_room_STRING
Recordset1_cmd.CommandText = "SELECT * FROM dbo.Log WHERE (Ltime <> ? or Ldate <> ?)" 
Recordset1_cmd.Prepared = true
Recordset1_cmd.Parameters.Append Recordset1_cmd.CreateParameter("param1", 200, 1, 50, Recordset1__MMColParam) ' adVarChar
Recordset1_cmd.Parameters.Append Recordset1_cmd.CreateParameter("param2", 200, 1, 255, Recordset1__MMColParam2) ' adVarChar

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
		<title>预定</title>
        <style type="text/css">
    		body {background-image: url(../image/background/background.jpg);}<!--添加背景图片-->
    	</style>   
	</head>

<body>
    <h3 style="color:white" align="center">可借教室</h3>
    <div style="width:1200;height:400; overflow:scroll; border:0 solid;overflow-x:auto;overflow-y:auto">
    <form action="" method="get">
		<table width="900" border="1" align="center">
  			<tr>
    			<td style="color:white" width="200">编号</td>              
  			</tr>
            <% 
While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) 
%>
  <tr>
    <td style="color:white" width="200"><%=(Recordset1.Fields.Item("Lroom").Value)%></td>
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
    
    <div align="center" style="color:#FFF">
    <form action="Sbooked.asp" method="post" id="su">
    请确认要借教室编号<input name="id" type="text" id="id" />
    <br>
    <input name="subm" type="submit" id="subm" value="确认" />
    </form>
</div>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
