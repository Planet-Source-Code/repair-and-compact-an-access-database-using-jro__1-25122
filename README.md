<div align="center">

## Repair and Compact an Access Database using JRO


</div>

### Description

This code will let you repair and compact an Access 97 or 2000 (haven't tested on other versions yet) database. Using JRO, when you call the CompactDatabase function, it automatically repairs the database first. You must have a reference to Microsoft Jet And Replication Objects x.x Library in your project.

Also, where I have a DoEvents in the code, there really should be some routine that checks to see when the file is actually deleted. I've tested this on databases that are about 5 megs in size with no problem. If anyone has any ideas (maybe a routine that checks to see when the file is unlocked or deleted), let me now. Thanks and enjoy!!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[N/A](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/empty.md)
**Level**          |Intermediate
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0, VBA MS Access
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/repair-and-compact-an-access-database-using-jro__1-25122/archive/master.zip)





### Source Code

<p>Option Explicit<br>
<br>
</p>
<p>'Must have reference to Microsoft Jet And Replication Objects x.x Library <br>
</p>
<p>Public Sub CompactDB(DBName As String)<br>
<br>
Dim jr As jro.JetEngine<br>
Dim strOld As String, strNew As String<br>
Dim x As Integer<br>
<br>
Set jr = New jro.JetEngine<br>
<br>
strOld = DBName<br>
x = InStrRev(strOld, &quot;\&quot;)<br>
strNew = Left(strOld, x)<br>
strNew = strNew &amp; &quot;chngMe.mdb&quot;<br>
<br>
'Use Engine Type = 4 for Access 97, Engine Type = 5 for Access 2000<br>
jr.CompactDatabase &quot;Provider=Microsoft.Jet.OLEDB.4.0;Data Source=&quot; &amp; strOld,
_<br>
&quot;Provider=Microsoft.Jet.OLEDB.4.0;Data Source=&quot; &amp; strNew &amp; &quot;;Jet
OLEDB:Engine Type=4&quot;<br>
<br>
Kill strOld<br>
DoEvents<br>
Name strNew As strOld<br>
<br>
Set jr = Nothing<br>
<br>
End Sub<br>
</p>

