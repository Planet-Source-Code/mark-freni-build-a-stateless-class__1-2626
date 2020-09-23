<div align="center">

## Build a Stateless Class


</div>

### Description

The example shows how to create a "Stateless" Class. By sending and receiving UDT's and disconnected recordsets you will find a significant increase in speed with your objects in a 3 tiered application. If you use collections, each time you get/set a property you make a network call. This method reduces network traffic.
 
### More Info
 
I declare a Public Type in the class module. In the user interface you need to declare the type in the procedure before you use it. Example:

' Dim udt as UdtTest

The user needs to understand UDT's, ADO, and Classes

Each of the functions return a value


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mark Freni](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mark-freni.md)
**Level**          |Unknown
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Data Structures](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/data-structures__1-33.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mark-freni-build-a-stateless-class__1-2626/archive/master.zip)





### Source Code

```
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'FORM cTest
'AUTHOR Mark Freni
'DESC Class to hold tblTest Functions,
' procedures, and variables
'FUNCTIONS GetList, Update, Add
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Option Explicit
 Public Type UdtTest
	' If the table has many fields this becomes
	' very convenient
 TestID As Long
 Field_1 As String
 Field_2 As Integer
 Active As Boolean
 End Type
Public Function GetList(Optional ByVal _
 ReturnAll As Boolean = False) As ADODB.Recordset
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Function : GetList
' Purpose : Provide a disconnected recordset of tblTest
' Author : Mark Freni
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
 On Error GoTo FUNCT_ERR
 Dim conn As New ADODB.Connection
 Dim strSql As String
 Dim rst As New ADODB.Recordset
 strSql = "SELECT * FROM tblTest"
 If ReturnAll Then
 strSql = strSql & " Where Active"
 End If
 With conn
 .CursorLocation = adUseClient
 .ConnectionString = strConnect
 End With
 conn.Open
 With rst
 .CursorLocation = adUseClient
 .LockType = adLockBatchOptimistic
 .CursorType = adOpenKeyset
 End With
 '~OPEN THE RECORDSET
 rst.Open strSql, conn
 Set rst.ActiveConnection = Nothing
 Set GetList = rst
FUNCT_EXIT:
 Set conn = Nothing
 Exit Function
FUNCT_ERR:
 Err.Raise Err.Number, Err.Source, Err.Description
 Resume FUNCT_EXIT
End Function
Public Function Add(udt As UdtTest) As Boolean
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Function : Add
' Purpose : Add a Record to tblTest
' Author : Mark Freni
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
 On Error GoTo FUNCT_ERR
 Dim conn As New ADODB.Connection
 Dim rst As New ADODB.Recordset
 Dim strSql As String
 conn.Open strConnect
 rst.CursorLocation = adUseClient
 rst.CursorType = adOpenKeyset
 rst.LockType = adLockBatchOptimistic
 rst.Open "tblTest", conn
 rst.AddNew
 With udt
	' I don't need to worry about setting quotes
	' using this method, the UDT tells the
	' recordset what datatypes the values are
 If Len(.Field_1) > 0 then rst("Field_1") = .Field_1
 If Len(.Field_2) > 0 then rst("Field_2") = .Field_2
 End With
 rst.UpdateBatch
 If rst.STATE = 1 Then rst.Close
 conn.Close
 Add = True
FUNCT_EXIT:
 Set conn = Nothing
 Set rst = Nothing
 Exit Function
FUNCT_ERR:
 Add = False
 Err.Raise Err.Number, Err.Source, Err.Description
 Resume FUNCT_EXIT
End Function
Public Function Update(udt As UdtTest) As Boolean
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Function : Update
' Purpose : Update a Record in tblTest
' Author : Mark Freni
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
 On Error GoTo FUNCT_ERR
 Dim conn As New ADODB.Connection
 Dim rst As New ADODB.Recordset
 Dim strSql As String
 conn.Open strConnect
 rst.CursorLocation = adUseServer
 rst.LockType = adLockBatchOptimistic
 strSql = "SELECT * FROM tblTest WHERE TestID =" & udt.TestID
 rst.Open strSql, conn
 If rst.EOF Then
 Update = False
 GoTo FUNCT_EXIT
 End If
 With udt
 If Len(.Field_1) > 0 Then rst("Field_1") = .Field_1
 If Len(.Field_2) > 0 Then rst("Field_2") = .Field_2
 If .Active Then rst("Active") = .Active
 End With
 rst.UpdateBatch
 If rst.STATE = 1 Then rst.Close
 conn.Close
 Update = True
FUNCT_EXIT:
 Set conn = Nothing
 Exit Function
FUNCT_ERR:
 Err.Raise Err.Number, Err.Source, Err.Description
 Update = False
 Resume FUNCT_EXIT
End Function
```

