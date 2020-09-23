Attribute VB_Name = "OMenu_h"
Option Explicit

'/////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////
'////                                                                         ////
'//// OMenu_h - This module is built in conjunction with the COwnMenu class.  ////
'////           This program demonstrates a popular object registration and   ////
'////           iteration process. This module maintains a list of COwnMenu   ////
'////           objects and pumps information and commands to them as the     ////
'////           Operating System dictates.                                    ////
'////                                                                         ////
'//// ----------------------------------------------------------------------- ////
'////                                                                         ////
'//// This program was created by Kalani Thielen on 04/14/98                  ////
'//// You may use the provided code module and object module if this text     ////
'//// appears within it.                                                      ////
'////                                                                         ////
'//// NOTE: If this code is used within a commercial (for profit) application ////
'////       please send US $20.00 in a self-addressed stamped envelope to:    ////
'////               Kalani Thielen                                            ////
'////               430 Quintana Road PMB 122                                 ////
'////               Morro Bay, CA 93442                                       ////
'////                                                                         ////
'//// For more programming information visit my website,                      ////
'//// the website is: http://www.calcoast.com/kalani/                         ////
'////                                                                         ////
'/////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////

'///////////////////////////////////////////////////
'// m_omList() is a dynamic array of COwnMenu
'// objects which represent individual menu entries
'///////////////////////////////////////////////////
Private m_omList() As COwnMenu
Private m_nOMCount As Long
Private m_bListInitialized As Boolean

'//////////////////////////////////////////////////////
'/// m_lPrevProc is the address of the procedure
'/// previously associated with the subclassed window
'//////////////////////////////////////////////////////
Private m_lPrevProc As Long

'////////////////////////////////////////////////////////////////
'//// Windows API functions
'////////////////////////////////////////////////////////////////
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long

'////////////////////////////////////////////////////////////////
'//// Windows API Constants
'////////////////////////////////////////////////////////////////
Private Const MF_OWNERDRAW = &H100&
Private Const MF_BYPOSITION = &H400&
Private Const GWL_WNDPROC = (-4)
Private Const WM_DRAWITEM = &H2B
Private Const WM_MEASUREITEM = &H2C
Private Const WM_COMMAND = &H111

'////////////////////////////////////////////////////////////////
'//// Structures used for Windows API functions
'////////////////////////////////////////////////////////////////
Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type MEASUREITEMSTRUCT
        CtlType As Long
        CtlID As Long
        itemID As Long
        itemWidth As Long
        itemHeight As Long
        itemData As Long
End Type

Public Type DRAWITEMSTRUCT
        CtlType As Long
        CtlID As Long
        itemID As Long
        itemAction As Long
        itemState As Long
        hwndItem As Long
        hdc As Long
        rcItem As RECT
        itemData As Long
End Type

'// text measurement functions/structures
Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZE) As Long
Type SIZE
        cx As Long
        cy As Long
End Type

'/////////////////////////////////////////////////////////
'////
'//// FreeMenus - Frees the memory allocated on the heap
'////             for our COwnMenu objects
'////
'/////////////////////////////////////////////////////////
Public Sub FreeMenus()
Dim nIndex As Long
For nIndex = 0 To m_nOMCount
    Set m_omList(nIndex) = Nothing
Next nIndex

m_nOMCount = 0
ReDim m_omList(0)
End Sub


'// Thiw procedure will tell Windows how big our items are.
Private Sub MeasureItem(ByRef mnu As COwnMenu, ByRef lpMeasureInfo As MEASUREITEMSTRUCT)
Dim hDrawDC As Long
Const MENU_HEIGHT = 20 '// average menu size, change if you want larger menu items
Const IMAGE_WIDTH = 16 '// the width of the image blt'ed into the menu dc

hDrawDC = GetDC(mnu.hwndOwner)

Dim lpSize As SIZE
GetTextExtentPoint32 hDrawDC, mnu.Caption, Len(mnu.Caption), lpSize

lpMeasureInfo.itemHeight = MENU_HEIGHT
lpMeasureInfo.itemWidth = lpSize.cx + IMAGE_WIDTH

ReleaseDC mnu.hwndOwner, hDrawDC
End Sub
Public Sub MakeOwnerDraw(hMenu As Long, nIndex As Long, nID As Long)
'// Modify the menu's attributes
ModifyMenu hMenu, nIndex, MF_OWNERDRAW Or MF_BYPOSITION, nID, vbNullString
End Sub



'/////////////////////////////////////////////////////////////////
'////
'//// IconProc - Your standard WndProc (Handles window messages)
'////
'/////////////////////////////////////////////////////////////////
Public Function IconProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim nRegisteredIndex As Long '// Used to iterate through all registered menu objects

'// We must make sure that the menu object array has been initialized
'// if it has not then we have no business processing any messages
If m_bListInitialized = False Then
    IconProc = CallWindowProc(m_lPrevProc, hwnd, uMsg, wParam, lParam)
    Exit Function
End If

'// The familiar window message select case block
Select Case uMsg
    Case WM_DRAWITEM
        '// The following code will copy a structure pointed to by lParam
        '// into our lpDrawInfo structure
        Dim lpDrawInfo As DRAWITEMSTRUCT
        CopyMem lpDrawInfo, ByVal lParam, Len(lpDrawInfo)
        
        '// We must draw an owner drawn menu
        '// loop through all currently created menu objects
        '// and see if we have correctly received this message
        For nRegisteredIndex = 0 To m_nOMCount
            If (m_omList(nRegisteredIndex).MenuID) = lpDrawInfo.itemID Then
                '// We have found our registered menu
                '// Let's tell the menu object to draw itself
                m_omList(nRegisteredIndex).InitStruct lpDrawInfo.hdc, lpDrawInfo.itemAction, lpDrawInfo.itemID, lpDrawInfo.itemState, lpDrawInfo.rcItem.Left, lpDrawInfo.rcItem.Top, lpDrawInfo.rcItem.Bottom, lpDrawInfo.rcItem.Right
                m_omList(nRegisteredIndex).DrawMenu
                Exit For
            End If
        Next nRegisteredIndex
    
    Case WM_MEASUREITEM
        Dim lpMeasureInfo As MEASUREITEMSTRUCT
        
        '// Get the MEASUREITEM struct from the pointer
        CopyMem lpMeasureInfo, ByVal lParam, Len(lpMeasureInfo)
        For nRegisteredIndex = 0 To m_nOMCount
            If (m_omList(nRegisteredIndex).MenuID) = lpMeasureInfo.itemID Then
                '// We have found our registered menu
                MeasureItem m_omList(nRegisteredIndex), lpMeasureInfo
                Exit For
            End If
        Next nRegisteredIndex
        CopyMem ByVal lParam, lpMeasureInfo, Len(lpMeasureInfo)
    
    Case Else
        '// Call previous WndProc
        IconProc = CallWindowProc(m_lPrevProc, hwnd, uMsg, wParam, lParam)
End Select
End Function

Public Sub RegisterMenu(hMenu As Long, nPosition As Long, hwndOwner As Long, sMessage As String, objPicture As Object)
'// Set this menu entry up as an owner drawn menu
MakeOwnerDraw hMenu, nPosition, GetMenuItemID(hMenu, nPosition)

'// Create a new owner drawn menu object on the heap
If (m_bListInitialized = False) Then
    ReDim m_omList(0)
    Set m_omList(0) = New COwnMenu
    
    m_omList(0).InitMenu GetMenuItemID(hMenu, nPosition), sMessage, objPicture
    
    m_bListInitialized = True
Else
    m_nOMCount = m_nOMCount + 1
    
    ReDim Preserve m_omList(m_nOMCount)
    Set m_omList(m_nOMCount) = New COwnMenu
    m_omList(m_nOMCount).hwndOwner = hwndOwner
    m_omList(m_nOMCount).InitMenu GetMenuItemID(hMenu, nPosition), sMessage, objPicture
End If
End Sub


Public Sub SetSubclass(frm As Form)
'// Store value of previous WndProc function
m_lPrevProc = GetWindowLong(frm.hwnd, GWL_WNDPROC)

'// Set new WndProc
SetWindowLong frm.hwnd, GWL_WNDPROC, AddressOf IconProc
End Sub


