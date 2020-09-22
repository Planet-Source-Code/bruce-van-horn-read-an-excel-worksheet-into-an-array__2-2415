<div align="center">

## Read an Excel Worksheet Into An Array


</div>

### Description

This code allows you to read an excel spreadsheet into an array. Once its in an array, you can search, sort, and manipulate data, and save it back to the worksheet, a different worksheet, to a database, or output via JSP.

Please note this sample only shows you how to get the sheet into an array -- I leave the rest for future contributions or to your imagination.
 
### More Info
 
You need an excel spreadsheet, and therefore, MS Excel.

You need an Excel spreadsheet and the latest ODBC drivers on your system to access them. This only runs in Windows as far as I know. There is no need to create an ODBC link to the file in question, making it handy for dynamic analysis

When I wrote this, I forgot to fix it so the array ran at base 0. Consequently element 0 on each row will be null. I decided I like it better that way. I tend to think in terms of base 1 anyway, as does excel which starts its cell numbering at A1, not A0. If you want to go to the base 0 route, just modify the for loop to start at 0 and end at mycount-1.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Bruce Van Horn](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bruce-van-horn.md)
**Level**          |Intermediate
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |Java \(JDK 1\.2\)
**Category**       |[Databases/ JDBC](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-jdbc__2-61.md)
**World**          |[Java](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/java.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bruce-van-horn-read-an-excel-worksheet-into-an-array__2-2415/archive/master.zip)





### Source Code

```
package bean_ed;
import java.sql.*;
import java.util.*;
/**
 * Title: Excel Reader
 * Description: Read an MS Excel Spreadsheet into a two dimensional array
 * Copyright:  Copyright (c) 2001
 * Company:
 * @author    Bruce Van Horn
 * @version 1.0
 */
public class read_excel {
//excel worksheets act like tables, but you have to put
//funky braces and a $ at the end of the name.
public static void main(String [] args) {
String sql="select count(*) from [mismatch$]";
//load excel driver
  try {
     Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");
    }
  catch (java.lang.ClassNotFoundException ex) {
     System.err.print("No Such Database Driver");
     System.err.println(ex.getMessage());
    };
String myDB="jdbc:odbc:Driver={Microsoft Excel Driver (*.xls)};DBQ=d:/bom_process/text.xls;DriverID=22;READONLY=false}";
int mycount=0;
  try {
     java.sql.Connection db=java.sql.DriverManager.getConnection(myDB,"","");
     java.sql.Statement st=db.createStatement();
     ResultSet rs=st.executeQuery(sql);
     //count number of records so we can build an array
     int recordcount=0;
     while (rs.next()) {
      recordcount=rs.getInt(1);
     }
     //run the real query
     sql="select * from [mismatch$]";
     rs.close();
     rs=st.executeQuery(sql);
     ResultSetMetaData rsmd = rs.getMetaData();
     mycount=rsmd.getColumnCount();
     String xlsSheet[][] = new String[recordcount][mycount];
     int rowcounter=0;
     while (rs.next()) {
      //add elements to array
       for (int i=1; i<mycount; i++) {
       xlsSheet[rowcounter][i]=rs.getString(i);
       //System.out.println("col: "+i+" row: "+rowcounter);
      }
     rowcounter++;
     }
     System.out.println(recordcount);
     System.out.println(xlsSheet [5][0] + " " + xlsSheet [5][1] + " " + xlsSheet [5][2] );
     rs.close();
     st.close();
     db.close();
    }
  catch (java.sql.SQLException ex) {
     System.err.print("SQL Exception");
     System.err.println(ex.getMessage());
    };
}
}
```

