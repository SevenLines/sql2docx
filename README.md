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
    <td>3</td><td>Mick</td><td>Jagger</td>
  </tr>
  <tr>
    <td>2</td><td>Kate</td><td>Sanders</td>
  </tr>
</table>

# Availabe options
---
