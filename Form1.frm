VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   660
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Begin Project1.ucNavigationBar ucNavigationBar1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   714
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Custom Background"
      Height          =   405
      Index           =   5
      Left            =   5040
      TabIndex        =   6
      Top             =   2700
      Width           =   1695
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Remove Selected"
      Height          =   405
      Index           =   4
      Left            =   5040
      TabIndex        =   5
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Toggle Enabled"
      Height          =   405
      Index           =   3
      Left            =   5040
      TabIndex        =   4
      Top             =   2250
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   3960
      Left            =   90
      TabIndex        =   3
      Top             =   480
      Width           =   4875
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Select Item 4"
      Height          =   375
      Index           =   1
      Left            =   5040
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Rename Item 1"
      Height          =   375
      Index           =   2
      Left            =   5040
      TabIndex        =   1
      Top             =   1380
      Width           =   1695
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Refresh"
      Height          =   375
      Index           =   0
      Left            =   5040
      TabIndex        =   0
      Top             =   540
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
    LoadItems
End Sub

Private Sub LoadItems()
    With ucNavigationBar1
        .Redraw = False
        .Height = 26 * Screen.TwipsPerPixelY
        'clear anything
        .ItemClear
        'root settings
        .RootSelection = False
        .ItemSetText "root", ""
        Set .Picture = LoadPicture(App.Path & "\sample.ico")  ' icon with mask
        'Set .Picture = LoadPicture(App.Path & "\explorer.bmp")  ' masked bitmap
        'list items
        .ItemAdd "level1", "", "Styles"
        .ItemAdd "level1.style1", "level1", "Office 97"
        .ItemAdd "level1.style2", "level1", "Office XP"
        .ItemAdd "level1.style3", "level1", "Office 2003"
        .ItemAdd "level1.style4", "level1", "Office 2007"
        .ItemAdd "item1", "", "Item 1", "Test 1"
        .ItemAdd "item2", "", "Item 2", "Test 2"
        .ItemAdd "item3", "item1", "Item 3", "Test 3"
        .ItemAdd "item4", "item3", "Item 4", "Test 4"
        'initial selection
        .ItemSelect "level1.style3"
        'update contents
        .Redraw = True
    End With
    List1.Clear
End Sub

Private Sub cmdAction_Click(Index As Integer)
    Select Case Index
        Case 0  'refresh
            LoadItems
        Case 1  'select item4
            ucNavigationBar1.ItemSelect "item4"
        Case 2  'rename item 1
            ucNavigationBar1.ItemSetText "item1", "Renamed Item " & Format$(Rnd * 100, "0.00000")
            ucNavigationBar1.ItemSetText "root", "", "Opa!"
        Case 3  'toggle enabled
            ucNavigationBar1.Enabled = Not ucNavigationBar1.Enabled
        Case 4  'remove item 1
            ucNavigationBar1.ItemRemove ucNavigationBar1.ItemSelected
        Case 5  'custom background
            If (ucNavigationBar1.CustomBackground) Then
                ucNavigationBar1.CustomBackground = False
                cmdAction(Index).Caption = "Custom Background"
            Else
                ucNavigationBar1.CustomBackground = True
                cmdAction(Index).Caption = "Themed Background"
            End If
    End Select
End Sub

Private Sub ucNavigationBar1_ButtonClick(ByVal ButtonKey As String, ByVal ButtonText As String, Cancel As Boolean)
    List1.AddItem "Button '" & ButtonText & "' [" & ButtonKey & "] selected"
    Select Case ButtonKey
        Case "level1.style1": ucNavigationBar1.Style = Office97
        Case "level1.style2": ucNavigationBar1.Style = OfficeXP
        Case "level1.style3": ucNavigationBar1.Style = Office2003
        Case "level1.style4": ucNavigationBar1.Style = Office2007
    End Select
End Sub
