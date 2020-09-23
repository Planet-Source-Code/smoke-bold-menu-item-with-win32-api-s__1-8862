<div align="center">

## Bold Menu Item With Win32 API's


</div>

### Description

Makes a menu item bold (for defualt items) by using Windows API's.
 
### More Info
 
Call the function in the following manner:

Call SetBold(Me, 1, 2)

Arguments:

----

Me = The form

1 = Menu index for mnuEdit

2 = Item index for mnuEditPaste


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Smoke](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/smoke.md)
**Level**          |Intermediate
**User Rating**    |4.8 (24 globes from 5 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/smoke-bold-menu-item-with-win32-api-s__1-8862/archive/master.zip)

### API Declarations

```
Private Declare Function GetMenu _
Lib "user32" ( _
ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu _
Lib "user32" ( _
ByVal hMenu As Long, _
ByVal nPos As Long) As Long
Private Declare Function SetMenuDefaultItem _
Lib "user32" ( _
ByVal hMenu As Long, _
ByVal uItem As Long, _
ByVal fByPos As Long) As Long
```


### Source Code

```
Public Sub SetBold(frmBold As Form, iMenuIndex As Long, iItemIndex As Long)
Dim hMnu As Long, hSubMnu As Long
hMnu = GetMenu(frmBold.hwnd)
hSubMnu = GetSubMenu(hMnu, iMenuIndex)
Call SetMenuDefaultItem(hSubMnu, iItemIndex, 1&)
End Sub
```

