<div align="center">

## MySQL Timestamp Conversion


</div>

### Description

When I started to work with MySQL, I found that working with dates was a lot more complicated than with MS Access. I saw someone else's code to do the conversation between VBS and MySQL date types, but realized that it could be done better, shorter and cleaner. This code will allow a VBS timestamp to be converted to a timestamp that will be accepted by a MySQL database.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[MeGuido](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/meguido.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__4-5.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/meguido-mysql-timestamp-conversion__4-8659/archive/master.zip)





### Source Code

```
Function ConvertSQLTimeStamp(strDateTime)
'Depending on regional settings, VBS may display time with AM/PM, and date as MM-DD-YYYY
'MySQL accepts timestamps in the following format: 'YYYY-MM-DD HH:MM:SS' (military time)
'In reality MySQL will accept timestamps like the following: '1999-1-6 5:4:3' and store
'them as '1999-01-06 05:04:03' appropriately.
'Get the year
ConvertSQLTimeStamp = Year(strDateTime) & "-"
'Get the month
ConvertSQLTimeStamp = ConvertSQLTimeStamp & Month(strDateTime) & "-"
'Get the day
ConvertSQLTimeStamp = ConvertSQLTimeStamp & Day(strDateTime) & " "
'Get the time (HH:MM - military format)
ConvertSQLTimeStamp = ConvertSQLTimeStamp & FormatDateTime(strDateTime, vbShortTime)
'Get and add the second
ConvertSQLTimeStamp = ConvertSQLTimeStamp & ":" & DatePart("s", strDateTime)
End Function
Function ConvertVBSTimeStamp(strDateTime)
'This function is completely unnecessary, however it is here to show how to convert
'MySQL timestamps to a different format. VBS can convert SQL timestamp directly without
'any modifications necessary.
'Format strDateTime using the systems regional settings
ConvertVBSTimeStamp = FormatDateTime(strDateTime)
End Function
```

