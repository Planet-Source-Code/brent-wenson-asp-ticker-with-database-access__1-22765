<%
Dim ID		
ID = Request.QueryString("ID") 'ID matches the Entry_ID for the news item in the database
'The following lines of code connect to the database and create the recordset.
Dim conCurrent
Dim rstCurrent				
Set conCurrent = Server.CreateObject("ADODB.Connection")
Set rstCurrent = Server.CreateObject("ADODB.recordset")
conCurrent.connectionString = "driver={Microsoft Access Driver (*.mdb)};" & _
 "dbq=c:\inetpub\wwwroot\news.mdb"
conCurrent.Open
strSQL="SELECT * FROM news WHERE Entry_ID=" & ID  'We use the ID here to select the correct
Set rstCurrent = conCurrent.Execute(strSQL)	      'news item
			
%>
<html>
<head>
<title>News Detail Page</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="Authors" content="Brent Wenson">
</head>

<body bgcolor="tan" text="#000000">
<CENTER>
  <H1>ASP News Ticker!</H1>
  <hr>
</CENTER>
<%
' The next lines get the info from the current news record and write them to the HTML Page
%>
<P><i><%=rstCurrent("Entry_Date")%></i>&nbsp;-&nbsp;<B><%=rstCurrent("Title")%></B></P><BR>
<P><%=rstCurrent("Text")%></P>
<P>&nbsp;</P>
<P><a href="newsfront.asp">Return to the Ticker!</a></P>

</body>
</html>

<%
conCurrent.close
set conCurrent = Nothing
%>