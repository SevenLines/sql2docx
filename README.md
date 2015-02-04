# sql2docx
Simple utility which let's one create templates file in docx format and fill it with from database.

Create simple docx document at the executable dir (call it **template.docx**) and add table in it with next content:

<table>
<tr>
  <td>#</td>
  <td>name</td>
  <td>secondname</td>
</tr>
<tr>
  <td colspan=3>
  <b>[CONTENT][WITH_COUNTER]</b><br>
  SELECT name, secondname<br>
  FROM People<br>
  WHERE sex = 'm'<br>
  <br>
  SELECT name, secondname<br>
  FROM People<br>
  WHERE sex = 'f'<br>
  </td>
</tr>
</table>

And let's assume that our table People has next rows:

| name | second_name  |sex|
| ----      | ----    |---|
| John      | James   | m |
| Kate      | Sanders | f |
| Mick      | Jagger  | m |

Create .credentials file at the executable dir with next content
> UserID: YourUser <br>
> Password: YourPassword <br>
> DataSource: HostName <br>
> InitialCatalog: YourDatabaseName <br>

Run the next command:
>sql2docx.exe -i template.docx -o "C:\result.docx" -s

If everything went right you will find file result.docx on **C:** drive with next content

<table>
  <tr>
    <td>#</td><td>name</td><td>secondname</td>
  </tr>
  <tr>
    <td>1</td><td>John</td><td>James</td>
  </tr>
  <tr>
    <td>2</td><td>Mick</td><td>Jagger</td>
  </tr>
  <tr>
    <td>3</td><td>Kate</td><td>Sanders</td>
  </tr>
</table>

# Availabe options

| parameter | description |
| --- | --- | 
| -i | input file path |
| -o | output file path |
| -s | open file at the opertaion complete |
| -q | save scripts in table, for future document update |
| -u | update document generated with -q flag with. All tables content will be replaced with new data |
| -c | .credentials path |
| -p | parameters as list of tuples: key1=value1 key2=value2 ... |

# Parameters

You can use parameters as placeholder in any part of document like this
> {ParameterName}

Which will be replaced with 
> ParameterValue

Recommended way to pass parameters to script:
```sql
SELECT *
FROM People
WHERE name like '%/*$NamePart*//*$*/%'
```

So if you launch programm with arguments
>sql2docx.exe -i template.docx -o "C:\result.docx" -s -p NamePart="Mic"

Script will be replaced this way:
```sql
SELECT *
FROM People
WHERE name like '%/*$NamePart*/Mic/*$*/%'
```

# Table script flags
All script flags must be at the first cell at first paragraph of last row of table. Script must start at second paragraph of the cell.

**[CONTENT]** - if you want to script to be parsed with programm it must contain this flag.<br>
**[WITH_COUNTER]** - add first column with number of row<br>
**[USE_ZERO_CONUTER]** - starts counter from 0<br>
**[AUTO_EXPAND]** - by default programm never adds new columns to table. Add this flag to change this behavior and make programm add new columns if there are not enough.<br>
**[REPLACE_TABLE_ON_EMPTY]** - replace table with text if there is no rows at result
