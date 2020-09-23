<div align="center">

## Get & set a file's attributes


</div>

### Description

Sets a file's attributes. With this you can create archive, hidden, normal,

read-only, and system files.
 
### More Info
 
FullFilePath is a string containing the path and filename of a file.

FileAttributes is a long integer that contains the value to set as the file's

attributes. Use the constants listed below to set this. (SetAttributes only)

The default for this function is to set the file's attributes to "archive",

which is the standard type of file.

To set more than one attribute for a file use the "or" statement.

I.E.- FILE_ATTRIBUTE_READONLY or FILE_ATTRIBUTE_HIDDEN.

Test for a certain attribute like this: If (GetAttribute And attributeconstant) <> 0 then GetAttribute = attribute defined by constant.

SetAttributes returns true (1) if successful, otherwise it returns false (0).

GetAttributes returns the attributes of FullFilePath

The FILE_ATTRIBUTE_NORMAL (&H80) attribute use to create attribute-less files

CAN NOT be combined with any other attribute.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Timothy Pew](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/timothy-pew.md)
**Level**          |Unknown
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/timothy-pew-get-set-a-file-s-attributes__1-1812/archive/master.zip)

### API Declarations

```
Declare Function SetFileAttributes Lib "kernel32.dll" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
```


### Source Code

```
'use these constants to set the attributes you want
FILE_ATTRIBUTE_ARCHIVE = &H20
FILE_ATTRIBUTE_COMPRESSED = &H800
FILE_ATTRIBUTE_DIRECTORY = &H10
FILE_ATTRIBUTE_HIDDEN = &H2
FILE_ATTRIBUTE_NORMAL = &H80
FILE_ATTRIBUTE_READONLY = &H1
FILE_ATTRIBUTE_SYSTEM = &H4
Public Function SetAttributes(ByVal FullFilePath As String, Optional ByVal FileAttributes As Long = &H20) As Long
 'makes sure that the file path is not too long
 FullFilePath = Left(FullFilePath, 255)
 SetAttributes = SetFileAttributes(FullFilePath, FileAttributes)
End Function
Public Function GetAttributes(ByVal FullFilePath as String) as Integer
 GetAttributes = GetFileAttributes(FullFilePath)
End Function
```

