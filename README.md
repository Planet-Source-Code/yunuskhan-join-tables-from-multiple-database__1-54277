<div align="center">

## Join tables from multiple database


</div>

### Description

This is a very usefull code, which i want to share it with u all, as i am working on an accounting package, this package closes it files at year end, and places it in backup database files. But the running year may need reports based on old data, i.e. the date range from the close file. Here is the SQL which can be used to get data form multiply database file in the fastes possible manner.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Yunuskhan](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/yunuskhan.md)
**Level**          |Advanced
**User Rating**    |4.2 (21 globes from 5 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/yunuskhan-join-tables-from-multiple-database__1-54277/archive/master.zip)





### Source Code

```
SELECT Table1.*
FROM Table1
UNION ALL
SELECT Table1.*
FROM Table1 IN "C:\Test\OldDatabase1.mdb"
UNION ALL
SELECT Table1.*
FROM Table1 IN "C:\Test\OldDatabase2.mdb"
UNION ALL
SELECT Table1.*
FROM Table1 IN "C:\Test\OldDatabase3.mdb";
```

