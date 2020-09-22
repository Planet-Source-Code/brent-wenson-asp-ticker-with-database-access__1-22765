<%
' The following is the code for the scrolling news ticker.  This will only work on a IIS server and IE browsers.  Netscape users will still see the news, but it will not scroll.
	
	Dim sTxt, iSpeed, iTop, iLeft, iWidth, iHeight, sHtml1, sHtml2, sHtml4, strSQL,sMarquee
	Dim conCurrent
	Dim rstCurrent				
	Set conCurrent = CreateObject("ADODB.Connection")
	Set rstCurrent = Server.CreateObject("ADODB.Recordset")
	conCurrent.Open "Driver={Microsoft Access Driver (*.mdb)};" & _
	 "DBQ=c:\inetpub\wwwroot\news.mdb"
	strSQL = "SELECT * FROM news ORDER BY Entry_Date DESC"
	Set rstCurrent = conCurrent.Execute(strSQL)
	sHtml1 = "<P><CENTER><FONT SIZE='-1' COLOR='Black'>"
	sHtml2 = "<A HREF='newsdetail.asp?ID="
	sHtml3 = "'>"
	sHtml4 = "</A></FONT></CENTER></P>"
	sTxt = ""
	rstCurrent.movefirst
	do while not rstCurrent.eof	'I used variables here to try and reduce this long assignment
		sTxt = sTxt & sHtml1 & rstCurrent("Entry_Date") & " - " & sHtml2 & _
		 rstCurrent("Entry_ID") & sHtml3 & rstCurrent("Title") & sHtml4
		rstCurrent.movenext
	loop
	iSpeed = 40 	' Speed of Marquee (higher = slower)
   	iTop = 0		' Y Location Within Object
   	iLeft = 0		' X Location""""
   	iWidth = 125	' Width 
   	iHeight = 167	' Height
   	'Insert marquee into objects innerHtml Property (in this Case a table cell)
	sMarquee="<MARQUEE onmouseover='this.stop();' " & _
	 "onmouseout='this.start();'direction='up' scrollamount='1' " & _
	 "scrolldelay='" & iSpeed & "' top='" & iTop & "' left='" & iLeft & _
	 "' width='" & iWidth & "' height='" & iHeight & "'>" & sTxt & "</MARQUEE>"
 
   	conCurrent.close		'Don't forget to clean-up!
	set conCurrent = Nothing
    %>
<html>
<head>
<title>News Front Page</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="Authors" content="Brent Wenson">
</head>

<body bgcolor="tan" text="#000000">
<CENTER><H1>ASP News Ticker!</H1></CENTER>
<P>&nbsp;</P>
            <table style="border-left: 1 solid #000000;border-right: 1 solid #D0D5DF;border-																		top: 1 solid #000000; border-bottom: 1 solid #D0D5DF" bgcolor=white width=125 height=167 cellpadding="2" cellspacing="0" align="center">
              <tr> 
                <td width="iWidth" height="166" bgcolor="#CCCCCC" id="tDisplay" bordercolor="#3333FF"> 
                  <%=sMarquee%> </td>
              </tr>
            </table>

</body>
</html>
