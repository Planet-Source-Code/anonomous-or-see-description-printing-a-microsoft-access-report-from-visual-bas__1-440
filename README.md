<div align="center">

## Printing a Microsoft Access Report from Visual Bas


</div>

### Description

How to print a Microsoft Access report from within VB. Also, VB 16-bit. (by Jose Garrick)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[anonomous \(or see description\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/anonomous-or-see-description.md)
**Level**          |Unknown
**User Rating**    |4.6 (37 globes from 8 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/anonomous-or-see-description-printing-a-microsoft-access-report-from-visual-bas__1-440/archive/master.zip)





### Source Code

```
Access 2.0 can be controlled using DDE, while Access 7.0 and later can be controlled using OLE Automation. In both cases, you are generally limited to what is available as a DoCmd statement/method. I'll assume for the moment that you'll be using one of the 32-bit versions of Access. You first setup a reference to Access in the VB References dialog box. Access 7.0 will show up as "Microsoft Access for Windows 95" and Access 8.0 will be listed as "Microsoft Access 8.0 Object Library".
Once that's done, you can create object variables in your application based on the Access application. This little snippet will open a database, run a report and close the database.
Dim ac As Access.Application
Set ac = New Access.Application
' put the path to your database in here
ac.OpenCurrentDatabase("c:\foo\foo.mdb")
' by default, the OpenReport method of the
' DoCmd object will send the report to the printer
ac.DoCmd.OpenReport "MyReport"
' close the database
ac.CloseCurrentDatabase
That's about all it takes. Just remember that you need to design the reports so that they can be run unattended. Watch for query prompts, message boxes, etc., in the report design or the code behind the report.
```

