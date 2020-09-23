VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "COwnMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'///////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////
'////                                                                               ////
'//// COwnMenu - This object demonstrates the process of both drawing an owner      ////
'////            drawn menu and encapsulating a complex process while still         ////
'////            allowing simple code in the actual implementation of the object.   ////
'////            While this class module may be a true work of art (No, I'm not     ////
'////            really *that* arrogant) it does leave room for a great deal of     ////
'////            improvement and customization. Hopefully you will find that the    ////
'////            framework set up in this demonstration will accomodate you in your ////
'////            mission to create any style of menu (like those funky MSN ones).   ////
'////            I only ask that you give me credit for the work I have done and    ////
'////            if you create new objects to accomodate for varying menu styles    ////
'////            that you keep this text in the object as well as your own notes    ////
'////                                                                               ////
'//// ----------------------------------------------------------------------------- ////
'////                                                                               ////
'//// This program was created by Kalani Thielen on 04/14/98                        ////
'//// You may use the provided code module and object module if this text           ////
'//// appears within it.                                                            ////
'////                                                                               ////
'//// NOTE: If this code is used within a commercial (for profit) application       ////
'////       please send US $20.00 in a self-addressed stamped envelope to:          ////
'////               Kalani Thielen                                                  ////
'////               430 Quintana Road PMB 122                                       ////
'////               Morro Bay, CA 93442                                             ////
'////                                                                               ////
'//// For more programming information visit my website,                            ////
'//// the website is: http://www.calcoast.com/kalani/                               ////
'////                                                                               ////
'///////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////

'//////////////////////////////////////////////////
'////// Object Data
'//////////////////////////////////////////////////

Private m_hMenu As Long '// The menu entry's handle
Private m_hMenuID As Long '// The menu entry's ID
Private m_sMessage As String '// The menu entry's text
Private m_objPicture As Object '// The menu entry's picture object
Private m_lpDrawStruct As DRAWITEMSTRUCT '// The menu entry's current drawing information
Public hwndOwner As Long     '// the window which owns this object

'////////////////////////////////////////////////////////////////////
'//// Windows API declarations - Used for drawing graphical data
'////                            into our menu entry's device context
'////////////////////////////////////////////////////////////////////
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

'//////////////////////////////////////////////////
'///// Constants used for Windows API functions
'//////////////////////////////////////////////////
Private Const SRCCOPY = &HCC0020

Private Const PS_SOLID = 0

Private Const COLOR_MENU = 4
Private Const COLOR_MENUTEXT = 7
Private Const COLOR_HIGHLIGHT = 13
Private Const COLOR_HIGHLIGHTTEXT = 14

Private Const ODS_SELECTED = &H1

Private Const NEWTRANSPARENT = 3

'///////////////////////////////////////////////////////////////////////
'//// Constants used by our object
'//// These values represent customizable aspects of this object
'//// which may be modified for compilation or customized as to provide
'//// dynamic modification of them.
'///////////////////////////////////////////////////////////////////////
Private Const DRAWWIDTH = 1
Private Const PicWidth = 20
Private Const TEXTBUFFER = 5

Public Property Let Caption(sMessage As String)
m_sMessage = sMessage
End Property

Public Property Get Caption() As String
Caption = m_sMessage
End Property

Public Sub InitMenu(hMenuID As Long, sMessage As String, objPicture As Object)
'// Set object state
m_hMenuID = hMenuID
m_sMessage = sMessage
Set m_objPicture = objPicture
End Sub

Public Sub InitStruct(hdc As Long, nAction As Long, nID As Long, nState As Long, nLeft As Long, nTop As Long, nBottom As Long, nRight As Long)
'// VB doesn't allow us to pass user defined structures to classes
'// therefore we have to go about it in a roundabout fashion
'// this leaves *you* as the programmer with room for future improvement
m_lpDrawStruct.hdc = hdc
m_lpDrawStruct.itemAction = nAction
m_lpDrawStruct.itemID = nID
m_lpDrawStruct.itemState = nState
m_lpDrawStruct.rcItem.Left = nLeft
m_lpDrawStruct.rcItem.Top = nTop
m_lpDrawStruct.rcItem.Bottom = nBottom
m_lpDrawStruct.rcItem.Right = nRight
End Sub

Public Property Get MenuID() As Long
MenuID = m_hMenuID
End Property


'///////////////////////////////////////////////////////////////////
'///////
'/////// PrintClear - Prints text with a clear background
'///////
'///////////////////////////////////////////////////////////////////
Private Sub PrintClear(crColor As Long)
'// Set DC background mode to clear
SetBkMode m_lpDrawStruct.hdc, NEWTRANSPARENT

'// Get old type color
Dim crOldType As Long
crOldType = GetTextColor(m_lpDrawStruct.hdc)

'// Set new type color
SetTextColor m_lpDrawStruct.hdc, crColor

'// Print text
TextOut m_lpDrawStruct.hdc, PicWidth + TEXTBUFFER, m_lpDrawStruct.rcItem.Top + 2, m_sMessage, Len(m_sMessage)

'// Reset old color
SetTextColor m_lpDrawStruct.hdc, crOldType
End Sub

'/////////////////////////////////////////////////////////////
'////// DrawMenu - Draws this menu entry
'/////////////////////////////////////////////////////////////
Public Sub DrawMenu()
'// Create a temporary copy of our member DRAWITEMSTRUCT
Dim lpDrawInfo As DRAWITEMSTRUCT
lpDrawInfo = m_lpDrawStruct

'// Create a brushes or get colors for specific menu attributes
'// These attributes are selected by the user in the Windows Display settings dialog
'// This ensures that menu customization will affect our menus too
Dim hSelectedItem As Long, crSelected As Long
Dim hMenuColor As Long

hSelectedItem = CreateSolidBrush(GetSysColor(COLOR_HIGHLIGHT))
crSelected = GetSysColor(COLOR_HIGHLIGHTTEXT)
hMenuColor = CreateSolidBrush(GetSysColor(COLOR_MENU))
        
'// Draw a highlight in the selection color if this element is selected
'// If this element is not selected we must clean up our previosu drawing

Dim DrawRect As RECT
DrawRect = m_lpDrawStruct.rcItem
DrawRect.Left = (PicWidth + TEXTBUFFER) - 4

If lpDrawInfo.itemState = ODS_SELECTED Then
    FillRect lpDrawInfo.hdc, DrawRect, hSelectedItem
Else
    FillRect lpDrawInfo.hdc, lpDrawInfo.rcItem, hMenuColor
End If

'// Print this menu entry's caption
PrintClear IIf(lpDrawInfo.itemState = ODS_SELECTED, crSelected, RGB(0, 0, 0))

'// Draw the bitmap for this menu entry
StretchBlt lpDrawInfo.hdc, lpDrawInfo.rcItem.Left + DRAWWIDTH, lpDrawInfo.rcItem.Top + DRAWWIDTH, PicWidth - DRAWWIDTH, (lpDrawInfo.rcItem.Bottom - lpDrawInfo.rcItem.Top) - DRAWWIDTH, m_objPicture.hdc, 0, 0, m_objPicture.ScaleWidth, m_objPicture.ScaleHeight, SRCCOPY

'// If our menu is selected we need to draw a 3D box around the picture
If lpDrawInfo.itemState = ODS_SELECTED Then Draw3D

'// Delete used GDI objects
DeleteObject hSelectedItem
DeleteObject hMenuColor
End Sub

'//////////////////////////////////////////////////////////////////////////////////
'////
'//// Draw3D - Draws a "3D" box around our picture
'////
'//////////////////////////////////////////////////////////////////////////////////
Private Sub Draw3D()
'// Create a drawing space in
Dim rctPicture As RECT
rctPicture.Top = m_lpDrawStruct.rcItem.Top
rctPicture.Left = m_lpDrawStruct.rcItem.Left
rctPicture.Right = PicWidth
rctPicture.Bottom = m_lpDrawStruct.rcItem.Bottom

'// Create pens for drawing the box border
Dim hpBlack As Long, hpWhite As Long

hpWhite = CreatePen(PS_SOLID, DRAWWIDTH, RGB(255, 255, 255)) '// White half of box
hpBlack = CreatePen(PS_SOLID, DRAWWIDTH, RGB(70, 70, 70))    '// Dark Grey half of box

'// Draw upper left corner of box
DeleteObject SelectObject(m_lpDrawStruct.hdc, hpWhite)

MoveToEx m_lpDrawStruct.hdc, rctPicture.Left + 1, (rctPicture.Bottom - 1), 0
LineTo m_lpDrawStruct.hdc, rctPicture.Left + 1, rctPicture.Top + 1
LineTo m_lpDrawStruct.hdc, (rctPicture.Right - 1), rctPicture.Top + 1

'// Draw lower right corner of box
DeleteObject SelectObject(m_lpDrawStruct.hdc, hpBlack)

LineTo m_lpDrawStruct.hdc, (rctPicture.Right - 1), rctPicture.Bottom - 1
LineTo m_lpDrawStruct.hdc, rctPicture.Left + 1, rctPicture.Bottom - 1

'// Clean up GDI objects
DeleteObject hpWhite
DeleteObject hpBlack
End Sub
