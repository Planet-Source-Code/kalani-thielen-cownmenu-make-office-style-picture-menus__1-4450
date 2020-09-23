VERSION 5.00
Begin VB.Form frmOwnMenu 
   Caption         =   "Owner Drawn Menu Example"
   ClientHeight    =   3375
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4185
   Icon            =   "frmOwnMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   225
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   279
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pctEntry4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   0
      Picture         =   "frmOwnMenu.frx":000C
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   3
      Top             =   1920
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox pctEntry2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   0
      Picture         =   "frmOwnMenu.frx":08D6
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox pctEntry3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   0
      Picture         =   "frmOwnMenu.frx":11A0
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox pctEntry1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   0
      Picture         =   "frmOwnMenu.frx":15E2
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lblKFiles 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The 'K' Files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3240
      MouseIcon       =   "frmOwnMenu.frx":1EAC
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label lblPlug 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "For programming tutorials and products visit:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   3090
   End
   Begin VB.Label lblKalInfo 
      BackStyle       =   0  'Transparent
      Height          =   2895
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3855
   End
   Begin VB.Menu mnuOwn 
      Caption         =   "Owner Drawn Menus"
      Begin VB.Menu mnuEntry1 
         Caption         =   "Entry #1"
      End
      Begin VB.Menu mnuEntry2 
         Caption         =   "Entry #2"
      End
      Begin VB.Menu mnuEntry3 
         Caption         =   "Entry #3"
      End
      Begin VB.Menu mnuEntry4 
         Caption         =   "Entry #4"
      End
   End
   Begin VB.Menu mnuReg 
      Caption         =   "Regular Menus"
      Begin VB.Menu mnuReg1 
         Caption         =   "Entry #1"
      End
      Begin VB.Menu mnuReg2 
         Caption         =   "Entry #2"
      End
      Begin VB.Menu mnuReg3 
         Caption         =   "Entry #3"
      End
      Begin VB.Menu mnuReg4 
         Caption         =   "Entry #4"
      End
   End
End
Attribute VB_Name = "frmOwnMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'/////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////
'////                                                                         ////
'//// frmOwnMenu - There isn't much to say here. Luckily the majority of      ////
'////              the work is taken care of by the OMenu_h code module,      ////
'////              which serves as an object manager and message handler,     ////
'////              and COwnMenu, which processes the actual commands and      ////
'////              draws each menu item to the screen. The only real work     ////
'////              that is done in this form module is in the InitMenus       ////
'////              procedure, which registers each menu entry with OMenu_h,   ////
'////              and in Form_Load, which initiates the subclass and calls   ////
'////              the InitMenus member function of this form. It is also     ////
'////              important to note that in Form_QueryUnload a procedure     ////
'////              in OMenu_h named "FreeMenus" is called. This procedure     ////
'////              frees the memory that is dynamically allocated by          ////
'////              OMenu_h in its registration process.                       ////
'////                                                                         ////
'//// ----------------------------------------------------------------------- ////
'////                                                                         ////
'//// If you've read this far at least it means you are making some           ////
'//// attempt to learn the code provided (Good luck to you!). If you have     ////
'//// any questions or comments please email them to KalaniCA@aol.com         ////
'//// If this example has been of use to you, you may want to visit           ////
'//// my website, at http://www.calcoast.com/kalani/                          ////
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

'// Function used to go to my web site
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_MAXIMIZE = 3

'/////////////////////////////////////////////////////////
'////
'//// InitMenus - Initializes our owner drawn menus
'////             this procedure simply registers each
'////             menu item with an appropriate
'////             COwnMenu object
'////
'/////////////////////////////////////////////////////////
Private Sub InitMenus()
'// Get top level menu handle
Dim hMainMenu As Long, hSubMenu As Long
hMainMenu = GetMenu(Me.hwnd)
hSubMenu = GetSubMenu(hMainMenu, 0)

'// Register each of our menus
RegisterMenu hSubMenu, 0, Me.hwnd, "Owner Drawn Entry #1", pctEntry1
RegisterMenu hSubMenu, 1, Me.hwnd, "Owner Drawn Entry #2", pctEntry2
RegisterMenu hSubMenu, 2, Me.hwnd, "Owner Drawn Entry #3", pctEntry3
RegisterMenu hSubMenu, 3, Me.hwnd, "Owner Drawn Entry #4", pctEntry4
End Sub
Private Sub ShowInfo()
Dim sMsg As String

sMsg = "Owner Drawn Menu Example by Kalani Thielen" & vbCrLf & vbCrLf
sMsg = sMsg & "This program demonstrates the process of subclassing your window"
sMsg = sMsg & " in order to catch commands which are passed by the Windows OS"
sMsg = sMsg & " to the windows which have created owner drawn menus." & vbCrLf & vbCrLf
sMsg = sMsg & "The COwnMenu object encapsulates the process of drawing each menu"
sMsg = sMsg & " entry and the OMenu_h code module manages a list of COwnMenu"
sMsg = sMsg & " objects which represent each menu entry that has been registered"
sMsg = sMsg & " as an owner drawn menu."

lblKalInfo.Caption = sMsg
End Sub

Private Sub Form_Load()
'// Initialize our menu objects
InitMenus

'// Set a subclass on this window so that we can process
'// requests to draw our owner drawn menus
SetSubclass Me

'// Show information about this program
ShowInfo
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'// Free the memory allocated by creating our owner drawn menus
FreeMenus
End Sub



Private Sub lblKFiles_Click()
'// Log on to the K Files web page at "http://members.aol.com/KalaniCOM"
ShellExecute 0, "open", "http://www.calcoast.com/kalani/", vbNullString, vbNullString, SW_MAXIMIZE
End Sub


Private Sub lblKFiles_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
lblKFiles.ForeColor = RGB(255, 0, 0)
End Sub


Private Sub lblKFiles_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
lblKFiles.ForeColor = RGB(0, 0, 255)
End Sub


