<div align="center">

## DBGrid Sort


</div>

### Description

It is nice technique for dbgrid sorting.You

can sort Dbgrid Columns by clicking on the grid column header

in two ways ascending or descending.
 
### More Info
 
Sorted Records in dbgrid.

No Idea,


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Adam Abas](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/adam-abas.md)
**Level**          |Unknown
**User Rating**    |4.3 (39 globes from 9 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/adam-abas-dbgrid-sort__1-1442/archive/master.zip)





### Source Code

```
Option Explicit
Dim st As Boolean
***********************
Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
'Dbgrid Columns sort by clicking the grid header in two way ascending and descending
 If st = True Then
 DBGrid1.HoldFields
 Data1.RecordSource = " Select * from Authors Order By " & DBGrid1.Columns(ColIndex).DataField
 Data1.Refresh
 DBGrid1.ReBind
 Else
 DBGrid1.HoldFields
 Data1.RecordSource = " Select * from Authors Order By " & DBGrid1.Columns(ColIndex).DataField & " DESC "
 Data1.Refresh
 DBGrid1.ReBind
 End If
 st = Not st
End Sub
```

