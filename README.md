<div align="center">

## ASP Ticker with Database Access


</div>

### Description

I searched around trying to find a ticker that operates much like the one on this site to no avail. So I wrote one. It was actually quite easy.
 
### More Info
 
You will need to change the DSN name field names to access your database. Also, change the page names to whatever want.

Of course, this code requires IIS or similar running on the server.

The code should be pretty self-explanatory. The database needs to contain a number field that uniquely identifies the news item, this number is what is used to create the link and pass it to the next asp page. It moves through the recordset and adds a new item for each news item in the database.

Keep in mind, the link will take the user to whatever page you specify. So you need to grab the parameter from the querystring on the next page and create code to search your database and format the data as you choose.


<span>             |<span>
---                |---
**Submitted On**   |2001-05-07 23:02:06
**By**             |[Brent Wenson](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brent-wenson.md)
**Level**          |Intermediate
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB Script, ASP \(Active Server Pages\) 
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[ASP Ticker19377572001\.zip](https://github.com/Planet-Source-Code/brent-wenson-asp-ticker-with-database-access__1-22765/archive/master.zip)








