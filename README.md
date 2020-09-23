<div align="center">

## Access ODBC DB using SQL and a Data Control


</div>

### Description

This code will allow you to populate DBGrids and other bound data controls using a Data Control, SQL, and an ODBC DataBase.
 
### More Info
 
Create a new form with a DBGrid and a Data Control. Set the datasource property on the grid to use the data control. Set the DB name property on the data control to whatever DB you set up in the ODBC applet on control panel. Set the DB type property to Use ODBC. Set Visible to false.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Scott Lewis](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/scott-lewis.md)
**Level**          |Unknown
**User Rating**    |3.2 (19 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/scott-lewis-access-odbc-db-using-sql-and-a-data-control__1-4114/archive/master.zip)





### Source Code

```
'This code assumes a DB with a table named "Appointments" and fields named '"AppName", "AppTime", "Appointment", and "Notes".
'put this into the Form_Load() area of the form the grid and data
'control are on.
  Data1.RecordSource = ""
  Data1.RecordSource = ReturnFieldsSQL
  Data1.Refresh
  DBGrid1.Refresh
'put this function in a module
Public Function ReturnFieldsSQL()
   Dim SQLS As String
   SQLS = "SELECT AppDate,"
   SQLS = SQLS + " " & "Apptime,"
   SQLS = SQLS + " " & "Appointment,"
   SQLS = SQLS + " " & "Notes"
   SQLS = SQLS + " " & "From [Appointments]"
   ReturnFieldsSQL = SQLS
End Function
'And thats all there is to it.
'This is a very simple function to use.
'You can alter the number of items to return.
'I'm still working on the syntax for the "Where" clause to go with this 'function.
'Once the form loads, if you do it right,
'the grid will be filled with the tables specified here.
```

