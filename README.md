<div align="center">

## Get XP Key the \.NET way


</div>

### Description

To get the CD-Key used to install Windows XP, completely in VB .NET, instead of using the API.
 
### More Info
 
This is an update to an update. The original code done in VB6 is here:

http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=57164&amp;lngWId=1

It was converted to VB .NET here:

http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=4578&amp;lngWId=10

I have removed all API calls in favor of .NET functions.

There is no error catching, so you'll have to add it yourself if you decide to use the code in your app.

The CD-Key that was used to install Windows XP.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Terry Olsen](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/terry-olsen.md)
**Level**          |Intermediate
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB\.NET
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__10-1.md)
**World**          |[\.Net \(C\#, VB\.net\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/net-c-vb-net.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/terry-olsen-get-xp-key-the-net-way__10-4834/archive/master.zip)

### API Declarations

None that I am aware of.


### Source Code

```
Imports Microsoft.Win32
Module modXPKey
  Public Function sGetXPKey() As String
    'Open the Registry Key and then get the value (byte array) from the SubKey
    Dim RegKey As RegistryKey = _
      Registry.LocalMachine.OpenSubKey("Software\Microsoft\Windows NT\CurrentVersion", False)
    Dim bytDPID() As Byte = RegKey.GetValue("DigitalProductID")
    'Transfer only the needed bytes into our Key Array
    ' Key starts at byte 52 and is 15 bytes long.
    Dim bytKey(14) As Byte '0-14 = 15 bytes
    Array.Copy(bytDPID, 52, bytKey, 0, 15)
    'Our "Array" of valid CD-Key characters
    Dim strChar As String = "BCDFGHJKMPQRTVWXY2346789"
    'Finally, our decoded CD-Key to be returned
    Dim strKey As String = ""
    'How Microsoft encodes this to begin with, I'd love to know...
    'but here's how we decode the byte array into a string
    'containing the CD-KEY.
    For j As Integer = 0 To 24
      Dim nCur As Short = 0
      For i As Integer = 14 To 0 Step -1
        nCur = CShort(nCur * 256 Xor bytKey(i))
        bytKey(i) = CByte(Int(nCur / 24))
        nCur = CShort(nCur Mod 24)
      Next
      strKey = strChar.Substring(nCur, 1) & strKey
    Next
    'Finally, insert the dashes into the string.
    For i As Integer = 4 To 1 Step -1
      strKey = strKey.Insert(i * 5, "-")
    Next
    Return strKey
  End Function
End Module
```

