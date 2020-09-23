<div align="center">

## COwnMenu \- Make Office\-style picture menus


</div>

### Description

To create an "Office-style menu" (or owner-draw menu) you must register that menu item with Windows as MF_OWNERDRAW and then process the WM_MEASUREITEM and WM_DRAWITEM messages sent to the menu's parent window. The attached project file simplifies this process by encapsulating all menu drawing operations in a class called "COwnMenu" and hiding the details of working with Windows in a code module entitled "OMenu_h." With this mini-system in place, all you have to do to get owner-drawn menus in your program is call SetSubclass on the menu's owner form and RegisterMenu to set a menu item as owner drawn. The provided example project contains complete documentation.
 
### More Info
 
Support for this code is not provided, please read the documentation in the project, as you'll find the answers to most relevant questions there.

The files omenu_h.bas and cownmenu.cls are meant to be included in projects which intend to include owner-draw menus.


<span>             |<span>
---                |---
**Submitted On**   |1999-11-13 21:38:06
**By**             |[Kalani Thielen](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kalani-thielen.md)
**Level**          |Unknown
**User Rating**    |4.6 (41 globes from 9 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD1809\.zip](https://github.com/Planet-Source-Code/kalani-thielen-cownmenu-make-office-style-picture-menus__1-4450/archive/master.zip)








