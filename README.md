<div align="center">

## How to dynamically invoke DCOM objects


</div>

### Description

Ever wished you could say

set c = createobject(server,myobj.myclass)

and have DCOM invoke the object on the specified server dynamically?.

This avoids having to configure which server a DCOM object should run

on each time you want to invoke the object (with the DCOMCNFG

utility). Plus you can run the same object on different servers

without running DCOMCNFG each time. andrew@geac.co.nz (Andrew Mayo)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Newsgroup Posting](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/newsgroup-posting.md)
**Level**          |Unknown
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/newsgroup-posting-how-to-dynamically-invoke-dcom-objects__1-722/archive/master.zip)

### API Declarations

```
'DCOM structures
Type SERV_STRUC
  reserved1 As Long
  ptrserver As Long
  authinfo As Long
  reserved2 As Long
End Type
Type MULTI_QI
  piid As Long
  pitf As Object
  hresult As Long
End Type
Declare Function CLSIDFromProgID Lib "ole32.dll" (progid As Any, clsid
As Any) As Long
Declare Function OleInitialize Lib "ole32.dll" (ByVal nullptr As Long)
As Long
Declare Function CoCreateInstanceEx Lib "ole32.dll" (clsid As Any,
ByVal punkouter As Long, ByVal context As Long, server As SERV_STRUC,
ByVal nelems As Long, mqi As MULTI_QI) As Long
Declare Function VarPtr Lib "vb40032.dll" (x As Any) As Long
Declare Function GetComputerName Lib "kernel32" Alias
"GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
```


### Source Code

```
The function below is intended to be made a public function
in a class library. Just say
dim c as object
dim d as object
set c = createobject("whatever.yourobjectis")
set d = invokeDCOMObject("someserver","someobject.someclass")
and you will get back an object reference to the remote DCOM object,
(or d will still be Nothing if the invocation failed). Set d to
nothing when you are done with the DCOM object.
The trick is to call CoCreateInstanceEx to do the dirty work - and get
an iDispatch interface pointer in one step. This is very efficient,
too. You get the interface by passing the 'well-known' REFIID of
iDispatch. If there is a way to programmatically do this I don't know
how, so I hard-coded the REFIID into a little routine.
'class-level variable for storing last error. You might want to
provide a property get routine to retrieve it.
dim clsLastError as string
Public Function InvokeDCOMOBject(remserver As String, objectname As
String) As Object
' Function which given a server and a object, will instantiate this
object on
' the server specified [if remserver is "" then this means local
computer]
' We use CoCreateInstanceEx to do the dirty work.
Dim clsid(256) As Byte
Dim progid() As Byte
Dim server() As Byte
Dim qi As MULTI_QI
Dim st As SERV_STRUC
Dim refiid(16) As Byte
Dim lrc As Long
clsLastError = ""
'now, there is a special case. If remserver is null or is same as our
machine,
'we do a local invoke instead and pass that back.
If remserver = "" Or UCase$(remserver) = UCase$(GetCompName()) Then
  On Error Resume Next
  Err = 0
  Set InvokeDCOMOBject = CreateObject(objectname)
  If Err <> 0 Then
    'record last error
    clsLastError = Err.errdesc
  End If
  On Error GoTo 0
  Exit Function
End If
'otherwise, it is genuinely remote.
'set an IID for IDispatch
GetIIDforIDispatch refiid()
'point to the IID
qi.piid = VarPtr(refiid(0))
'specify the object to be launched
progid = objectname & Chr$(0)
'specify the server
server = remserver & Chr$(0)
'initialise OLE
lrc = OleInitialize(0)
'get the CLSID for the object
lrc = CLSIDFromProgID(progid(0), clsid(0))
If lrc <> 0 Then
  clsLastError = "Unable to obtain CLSID from progid " & objectname
& vbCrLf & "Possibly it is not registered on both this server and
server " & remserver
  Exit Function
End If
'point to server name
st.ptrserver = VarPtr(server(0))
'invoke a remote instance of the desired object
lrc = CoCreateInstanceEx(clsid(0), 0, 16, st, 1, qi)
If lrc <> 0 Then
  clsLastError = "CoCreateInstanceEx failed with error code " &
Hex$(lrc)
  Exit Function
End If
'pass back object ref.
Set InvokeDCOMOBject = qi.pitf
End Function
Public Sub GetIIDforIDispatch(p() As Byte)
'fills in the well-known IID for IDispatch into the byte array p.
p(1) = 4
p(2) = 2
p(8) = &HC0
p(15) = &H46
End Sub
Function GetCompName() As String
'return the computer name
Dim buf As String
Dim rc As Long
buf = String$(256, 0)
rc = GetComputerName(buf, Len(buf))
If InStr(buf, Chr$(0)) > 1 Then
  GetCompName = UCase$(Left$(buf, InStr(buf, Chr$(0)) - 1))
End If
End Function
```

