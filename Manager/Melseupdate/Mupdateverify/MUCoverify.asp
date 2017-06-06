<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../../Connections/room.asp" -->
<%
Dim Mshowcon
Dim Mshowcon_cmd
Dim Mshowcon_numRows

Set Mshowcon_cmd = Server.CreateObject ("ADODB.Command")
Mshowcon_cmd.ActiveConnection = MM_room_STRING
Mshowcon_cmd.CommandText = "SELECT * FROM dbo.Conference" 
Mshowcon_cmd.Prepared = true

Set Mshowcon = Mshowcon_cmd.Execute
Mshowcon_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
Mshowcon_numRows = Mshowcon_numRows + Repeat1__numRows
%>
<%
Dim MM_paramName 
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

Dim MM_keepNone
Dim MM_keepURL
Dim MM_keepForm
Dim MM_keepBoth

Dim MM_removeList
Dim MM_item
Dim MM_nextItem

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then
  MM_removeList = MM_removeList & "&" & MM_paramName & "="
End If

MM_keepURL=""
MM_keepForm=""
MM_keepBoth=""
MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each MM_item In Request.QueryString
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & MM_nextItem & Server.URLencode(Request.QueryString(MM_item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each MM_item In Request.Form
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & MM_nextItem & Server.URLencode(Request.Form(MM_item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
If (MM_keepBoth <> "") Then 
  MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
End If
If (MM_keepURL <> "")  Then
  MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
End If
If (MM_keepForm <> "") Then
  MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)
End If

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>确认会议室</title>
        <style type="text/css">
    		body {background-image: url(../../../image/background/background.jpg);}<!--添加背景图片-->
    	</style>
</head>

<body>
    <br>
    <br>
    <br>
    <h3 style="color:white" align="center">会议室信息</h3>
    <div style="width:1200;height:400; overflow:scroll; border:0 solid;overflow-x:auto;overflow-y:auto">
    <form action="" method="Coget">
	  <table width="900" border="1" align="center">
		<tr>
    			<td style="color:white" width="200">编号</td>
    			<td style="color:white" width="300">容量</td>
    			<td style="color:white" width="300">位置</td>                                
	    </tr>
        <% 
While ((Repeat1__numRows <> 0) AND (NOT Mshowcon.EOF)) 
%>
  <tr>
    <td style="color:white" width="200"><%=(Mshowcon.Fields.Item("Cono").Value)%></td>
    <td style="color:white" width="300"><%=(Mshowcon.Fields.Item("Cocapa").Value)%></td>
    <td style="color:white" width="300"><%=(Mshowcon.Fields.Item("Cosite").Value)%></td>
  </tr>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Mshowcon.MoveNext()
Wend
%>
      </table>
      </form>
	</div>
    <br>
    <br>
    <br>
    <h3 align="center" style="color:#FFF">请确认要修改的会议室编号<h3>
    <div align="center">
    	<form action="../Mconupdate.asp" method="post"><input name="no" type="text" id="no" />
       		 <br>
             <br>
       		 <input name="" type="submit" value="提交" />
        </form>
    </div>
        <div align="center">
    	<form action="coup.asp" method="post"><input name="noo" type="text" id="no" />
       		 <br>
             <br>
       		 <input name="" type="submit" value="上线" />
        </form>
    </div>
</body>
</html>
<%
Mshowcon.Close()
Set Mshowcon = Nothing
%>
