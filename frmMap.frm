VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMap 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7350
   Icon            =   "frmMap.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraMapControls 
      BackColor       =   &H00000000&
      Caption         =   "Map Control"
      ForeColor       =   &H00E0E0E0&
      Height          =   1395
      Left            =   120
      TabIndex        =   63
      Top             =   420
      Visible         =   0   'False
      Width           =   1395
      Begin VB.CommandButton cmdMove 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   870
         MaskColor       =   &H80000016&
         TabIndex        =   32
         Top             =   990
         Width           =   375
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "U"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   150
         MaskColor       =   &H80000016&
         TabIndex        =   31
         Top             =   990
         Width           =   375
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "SE"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   870
         MaskColor       =   &H80000016&
         TabIndex        =   30
         Top             =   750
         Width           =   375
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   510
         MaskColor       =   &H80000016&
         TabIndex        =   29
         Top             =   750
         Width           =   375
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "SW"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   150
         MaskColor       =   &H80000016&
         TabIndex        =   28
         Top             =   750
         Width           =   375
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   870
         MaskColor       =   &H80000016&
         TabIndex        =   27
         Top             =   510
         Width           =   375
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "W"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   150
         MaskColor       =   &H80000016&
         TabIndex        =   26
         Top             =   510
         Width           =   375
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "NE"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   870
         MaskColor       =   &H80000016&
         TabIndex        =   25
         Top             =   270
         Width           =   375
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   510
         MaskColor       =   &H80000016&
         TabIndex        =   24
         Top             =   270
         Width           =   375
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "NW"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   150
         MaskColor       =   &H80000016&
         TabIndex        =   23
         Top             =   270
         Width           =   375
      End
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00000000&
      Caption         =   "Options"
      ForeColor       =   &H00E0E0E0&
      Height          =   4035
      Left            =   2280
      TabIndex        =   64
      Top             =   420
      Visible         =   0   'False
      Width           =   2595
      Begin VB.CommandButton cmdQ 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   2280
         TabIndex        =   17
         Top             =   2280
         Width           =   195
      End
      Begin VB.CheckBox chkMapOptions 
         BackColor       =   &H00000000&
         Caption         =   "Allow Main To Overlap"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   16
         Top             =   2340
         Width           =   2115
      End
      Begin VB.ComboBox cmbMapSize 
         Height          =   315
         ItemData        =   "frmMap.frx":0CCA
         Left            =   180
         List            =   "frmMap.frx":0CD7
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2580
         Width           =   2235
      End
      Begin VB.CommandButton cmdQ 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2280
         TabIndex        =   14
         Top             =   1860
         Width           =   195
      End
      Begin VB.CheckBox chkMapOptions 
         BackColor       =   &H00000000&
         Caption         =   "Show Map Controls"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   13
         Top             =   1860
         Width           =   1875
      End
      Begin VB.CheckBox chkMapOptions 
         BackColor       =   &H00000000&
         Caption         =   "Not ""Always on Top"""
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   15
         Top             =   2100
         Width           =   1935
      End
      Begin VB.CommandButton cmdQ 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   2280
         TabIndex        =   6
         Top             =   240
         Width           =   195
      End
      Begin VB.CommandButton cmdViewMapLegend 
         Caption         =   "View Help/&Legend"
         Height          =   315
         Left            =   180
         TabIndex        =   22
         Top             =   3600
         Width           =   2235
      End
      Begin VB.CommandButton cmdMapShowUnused 
         Caption         =   "S&how Unused Blocks"
         Height          =   315
         Left            =   180
         TabIndex        =   21
         Top             =   3300
         Width           =   2235
      End
      Begin VB.CommandButton cmdMapFindText 
         Caption         =   "Find &Next"
         Height          =   315
         Index           =   1
         Left            =   1380
         TabIndex        =   20
         Top             =   3000
         Width           =   1035
      End
      Begin VB.CommandButton cmdMapFindText 
         Caption         =   "&Find Room"
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   19
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CheckBox chkMapOptions 
         BackColor       =   &H00000000&
         Caption         =   "Don't Follow Hidden Exits"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   8
         Top             =   510
         Width           =   2235
      End
      Begin VB.CheckBox chkMapOptions 
         BackColor       =   &H00000000&
         Caption         =   "Follow Map Changes"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   240
         Width           =   1875
      End
      Begin VB.CheckBox chkMapOptions 
         BackColor       =   &H00000000&
         Caption         =   "Don't Mark Lairs"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   9
         Top             =   750
         Width           =   2235
      End
      Begin VB.CheckBox chkMapOptions 
         BackColor       =   &H00000000&
         Caption         =   "Don't Mark NPCs"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   10
         Top             =   1005
         Width           =   2235
      End
      Begin VB.CheckBox chkMapOptions 
         BackColor       =   &H00000000&
         Caption         =   "Don't Mark Commands"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   11
         Top             =   1260
         Width           =   2235
      End
      Begin VB.CheckBox chkMapOptions 
         BackColor       =   &H00000000&
         Caption         =   "Don't Show Tooltips"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   12
         Top             =   1500
         Width           =   2235
      End
   End
   Begin VB.Frame fraPresets 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Presets"
      ForeColor       =   &H00E0E0E0&
      Height          =   3855
      Left            =   2280
      TabIndex        =   60
      Top             =   420
      Visible         =   0   'False
      Width           =   2595
      Begin VB.CommandButton cmdMapPresetSelect 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   1320
         TabIndex        =   37
         Top             =   300
         Width           =   315
      End
      Begin VB.CommandButton cmdMapPresetSelect 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1020
         TabIndex        =   36
         Top             =   300
         Width           =   315
      End
      Begin VB.CommandButton cmdEditPreset 
         Caption         =   "!"
         Height          =   315
         Index           =   9
         Left            =   2280
         TabIndex        =   58
         Top             =   3420
         Width           =   195
      End
      Begin VB.CommandButton cmdEditPreset 
         Caption         =   "!"
         Height          =   315
         Index           =   8
         Left            =   2280
         TabIndex        =   56
         Top             =   3120
         Width           =   195
      End
      Begin VB.CommandButton cmdEditPreset 
         Caption         =   "!"
         Height          =   315
         Index           =   7
         Left            =   2280
         TabIndex        =   54
         Top             =   2820
         Width           =   195
      End
      Begin VB.CommandButton cmdEditPreset 
         Caption         =   "!"
         Height          =   315
         Index           =   6
         Left            =   2280
         TabIndex        =   52
         Top             =   2520
         Width           =   195
      End
      Begin VB.CommandButton cmdEditPreset 
         Caption         =   "!"
         Height          =   315
         Index           =   5
         Left            =   2280
         TabIndex        =   50
         Top             =   2220
         Width           =   195
      End
      Begin VB.CommandButton cmdEditPreset 
         Caption         =   "!"
         Height          =   315
         Index           =   4
         Left            =   2280
         TabIndex        =   48
         Top             =   1920
         Width           =   195
      End
      Begin VB.CommandButton cmdEditPreset 
         Caption         =   "!"
         Height          =   315
         Index           =   3
         Left            =   2280
         TabIndex        =   46
         Top             =   1620
         Width           =   195
      End
      Begin VB.CommandButton cmdEditPreset 
         Caption         =   "!"
         Height          =   315
         Index           =   2
         Left            =   2280
         TabIndex        =   44
         Top             =   1320
         Width           =   195
      End
      Begin VB.CommandButton cmdEditPreset 
         Caption         =   "!"
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   42
         Top             =   1020
         Width           =   195
      End
      Begin VB.CommandButton cmdMapPresetSelect 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   35
         Top             =   300
         Width           =   315
      End
      Begin VB.CommandButton cmdMapPresetSelect 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   420
         TabIndex        =   34
         Top             =   300
         Width           =   315
      End
      Begin VB.CommandButton cmdMapPreset 
         Caption         =   "Lava Fields"
         Height          =   315
         Index           =   9
         Left            =   120
         TabIndex        =   57
         Top             =   3420
         Width           =   2115
      End
      Begin VB.CommandButton cmdMapPreset 
         Caption         =   "Ancient Ruin"
         Height          =   315
         Index           =   8
         Left            =   120
         TabIndex        =   55
         Top             =   3120
         Width           =   2115
      End
      Begin VB.CommandButton cmdMapPreset 
         Caption         =   "Storm Fortress"
         Height          =   315
         Index           =   7
         Left            =   120
         TabIndex        =   53
         Top             =   2820
         Width           =   2115
      End
      Begin VB.CommandButton cmdMapPreset 
         Caption         =   "Black Fortress"
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   51
         Top             =   2520
         Width           =   2115
      End
      Begin VB.CommandButton cmdMapPreset 
         Caption         =   "Commander Markus"
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   49
         Top             =   2220
         Width           =   2115
      End
      Begin VB.CommandButton cmdMapPreset 
         Caption         =   "Rhudar"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   47
         Top             =   1920
         Width           =   2115
      End
      Begin VB.CommandButton cmdMapPreset 
         Caption         =   "Lost City"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   45
         Top             =   1620
         Width           =   2115
      End
      Begin VB.CommandButton cmdMapPreset 
         Caption         =   "Arlysia"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   43
         Top             =   1320
         Width           =   2115
      End
      Begin VB.CommandButton cmdMapPreset 
         Caption         =   "Khazarad"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   41
         Top             =   1020
         Width           =   2115
      End
      Begin VB.CommandButton cmdMapPresetSelect 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   33
         Top             =   300
         Width           =   315
      End
      Begin VB.CommandButton cmdResetPresets 
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   38
         Top             =   300
         Width           =   675
      End
      Begin VB.CommandButton cmdMapPreset 
         Caption         =   "Town Square"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   720
         Width           =   2115
      End
      Begin VB.CommandButton cmdEditPreset 
         Caption         =   "!"
         Height          =   315
         Index           =   0
         Left            =   2280
         TabIndex        =   40
         Top             =   720
         Width           =   195
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   60
      TabIndex        =   65
      Top             =   60
      Width           =   7245
      Begin VB.CommandButton cmdDrawMap 
         Caption         =   "&Draw"
         Default         =   -1  'True
         Height          =   315
         Index           =   0
         Left            =   1380
         TabIndex        =   2
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdPresets 
         Caption         =   "&Presets"
         Height          =   315
         Left            =   3900
         TabIndex        =   5
         ToolTipText     =   "Goes back one room"
         Top             =   0
         Width           =   915
      End
      Begin VB.CommandButton cmdOptions 
         Caption         =   "&Options"
         Height          =   315
         Left            =   2880
         TabIndex        =   4
         ToolTipText     =   "Goes back one room"
         Top             =   0
         Width           =   975
      End
      Begin VB.TextBox txtRoomRoom 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   600
         MaxLength       =   5
         TabIndex        =   1
         Text            =   "1"
         Top             =   0
         Width           =   735
      End
      Begin VB.TextBox txtRoomMap 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   0
         MaxLength       =   5
         TabIndex        =   0
         Text            =   "1"
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton cmdDrawMap 
         Caption         =   "&Last"
         Height          =   315
         Index           =   1
         Left            =   2220
         TabIndex        =   3
         ToolTipText     =   "Back to last room"
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.PictureBox picMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   7245
      Left            =   60
      ScaleHeight     =   481
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   481
      TabIndex        =   59
      Top             =   390
      Width           =   7245
      Begin VB.Label lblRoomCell 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1
         Left            =   60
         TabIndex        =   62
         Top             =   60
         Visible         =   0   'False
         Width           =   135
      End
   End
   Begin MSComctlLib.ListView lvMapLoc 
      Height          =   1035
      Left            =   60
      TabIndex        =   61
      Top             =   7680
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   1826
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   14737632
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Menu mnuMapPopUp 
      Caption         =   "MapMenuPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuMapPopUpItem 
         Caption         =   "Follow Up and Redraw"
         Index           =   0
      End
      Begin VB.Menu mnuMapPopUpItem 
         Caption         =   "Follow Down and Redraw"
         Index           =   1
      End
      Begin VB.Menu mnuMapPopUpItem 
         Caption         =   "Redraw From Here"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Enum EnumDrawRoom
    drSquare = 0
    drStar = 1
    drOpenCircle = 2
    drUp = 3
    drDown = 4
    drCircle = 5
    drLineN = 6
    drLineS = 7
    drLineE = 8
    drLineW = 9
    drLineNe = 10
    drLineNw = 11
    drLineSe = 12
    drLineSw = 13
End Enum

Dim nMapLastFind(0 To 2) As Long
Dim nMapLastCellIndex As Integer
Dim bMapStillMapping As Boolean
Dim sMapSECorner As Integer
Dim nMapRowLength As Integer
Public nMapStartRoom As Long
Public nMapStartMap As Long
Dim nMapCenterCell As Integer
Dim sMapSearch As String
Dim nMapLastRoom As Long
Dim nMapLastMap As Long
Dim nMapCurrentRecord As Variant
Public bMapSwapButtons As Boolean
Public bMapCancelFind As Boolean
Dim CellRoom() As Long
Dim UnchartedCells() As Integer
Dim StopBuild As Boolean

Dim TTlbl As clsToolTip

Private Sub cmbMapSize_Click()
Select Case cmbMapSize.ListIndex
    Case 1: '30x30
        nMapCenterCell = 436
    Case 2: '40x40
        nMapCenterCell = 779
    Case Else: '20x20
        nMapCenterCell = 219
        
End Select
End Sub

Private Sub Form_Activate()
If chkMapOptions(6).Value = 0 Then Call SetTopMostWindow(Me.hwnd, True)
End Sub

Private Sub Form_Load()
On Error GoTo Error:
Dim lR As Long

Set TTlbl = New clsToolTip

With TTlbl
    .DelayTime = 20
    .VisibleTime = 20000
    .BkColor = &HC0FFFF
    .TxtColor = &H0
    .Style = ttStyleStandard
    '.Style = ttStyleStandard
End With

If Not ReadINI("Settings", "MapExternalOnTop") = "1" Then
    lR = SetTopMostWindow(Me.hwnd, True)
Else
    chkMapOptions(6).Value = 1
End If

lvMapLoc.ColumnHeaders.clear
lvMapLoc.ColumnHeaders.Add 1, "References", "References", 4500

bMapSwapButtons = frmMain.bMapSwapButtons

Me.Top = ReadINI("Settings", "ExMapTop")
Me.Left = ReadINI("Settings", "ExMapLeft")

chkMapOptions(0).Value = ReadINI("Settings", "ExMapFollowMap")
chkMapOptions(1).Value = ReadINI("Settings", "ExMapNoHidden")
chkMapOptions(2).Value = ReadINI("Settings", "ExMapNoLairs")
chkMapOptions(3).Value = ReadINI("Settings", "ExMapNoNPC")
chkMapOptions(4).Value = ReadINI("Settings", "ExMapNoCMD")
chkMapOptions(5).Value = ReadINI("Settings", "ExMapNoTooltips")
chkMapOptions(8).Value = ReadINI("Settings", "ExMapMainOverlap")

Call LoadPresets

cmbMapSize.ListIndex = Val(ReadINI("Settings", "ExMapSize"))

Call ResizeMap

Exit Sub
Error:
Call HandleError("Form_Load")
Resume Next
End Sub

Private Sub chkMapOptions_Click(Index As Integer)
Dim lR As Long

If Index = 6 Then
    If chkMapOptions(6).Value = 1 Then
        lR = SetTopMostWindow(Me.hwnd, False)
    Else
        lR = SetTopMostWindow(Me.hwnd, True)
    End If
    If FormIsLoaded("frmResults") Then
        If frmResults.objFormOwner Is Me Then
            If chkMapOptions(6).Value = 1 Then
                lR = SetTopMostWindow(frmResults.hwnd, False)
            Else
                lR = SetTopMostWindow(frmResults.hwnd, True)
            End If
        End If
    End If
ElseIf Index = 7 Then
    If chkMapOptions(7).Value = 1 Then
        fraMapControls.Visible = True
    Else
        fraMapControls.Visible = False
    End If
End If

End Sub

Private Sub cmdMove_Click(Index As Integer)
On Error GoTo Error:
Dim sLook As String, RoomExit As RoomExitType
Dim nExitType As Integer, nRecNum As Long

tabRooms.Index = "idxRooms"
tabRooms.Seek "=", nMapStartMap, nMapStartRoom
If tabRooms.NoMatch Then GoTo out:

Select Case Index
    Case 0: sLook = "N"
    Case 1: sLook = "S"
    Case 2: sLook = "E"
    Case 3: sLook = "W"
    Case 4: sLook = "NE"
    Case 5: sLook = "NW"
    Case 6: sLook = "SE"
    Case 7: sLook = "SW"
    Case 8: sLook = "U"
    Case 9: sLook = "D"
End Select

If Left(tabRooms.Fields(sLook), 6) = "Action" Then
    GoTo out:
ElseIf Not Val(tabRooms.Fields(sLook)) = 0 Then
    RoomExit = ExtractMapRoom(tabRooms.Fields(sLook))
    
    tabRooms.Index = "idxRooms"
    tabRooms.Seek "=", RoomExit.Map, RoomExit.Room
    If tabRooms.NoMatch Then
        MsgBox "Error going in that direction."
        GoTo out:
    End If
Else
    GoTo out:
End If

Call MapStartMapping(RoomExit.Map, RoomExit.Room)

out:
Exit Sub
Error:
Call HandleError("cmdMove_Click")
Resume out:
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If Me.ActiveControl Is txtRoomMap Then
    Exit Sub
ElseIf Me.ActiveControl Is txtRoomRoom Then
    Exit Sub
End If

Select Case KeyAscii
    Case 46, 45: 'd
        Call cmdMove_Click(9)
    Case 48, 61: 'u
        Call cmdMove_Click(8)
    Case 49, 122: 'sw
        Call cmdMove_Click(7)
    Case 50, 120: 's
        Call cmdMove_Click(1)
    Case 51, 99: 'se
        Call cmdMove_Click(6)
    Case 52, 97: 'w
        Call cmdMove_Click(3)
    'Case 53:
    Case 54, 100: 'e
        Call cmdMove_Click(2)
    Case 55, 113: 'nw
        Call cmdMove_Click(5)
    Case 56, 119: 'n
        Call cmdMove_Click(0)
    Case 57, 101: 'ne
        Call cmdMove_Click(4)
End Select

End Sub

Private Sub MapGoDirection(ByVal nSourceMapNumber As Long, ByVal nSourceRoomNumber As Long, ByVal sDirection As String)
On Error GoTo Error:
Dim RoomExits As RoomExitType

tabRooms.Index = "idxRooms"
tabRooms.Seek "=", nSourceMapNumber, nSourceRoomNumber
If tabRooms.NoMatch Then
    MsgBox "Source room (" & nSourceMapNumber & "/" & nSourceRoomNumber & ") not found."
    Exit Sub
End If

RoomExits = ExtractMapRoom(tabRooms.Fields(sDirection))
If Not RoomExits.Map = 0 And Not RoomExits.Room = 0 Then
    Call MapStartMapping(RoomExits.Map, RoomExits.Room)
End If
Exit Sub
Error:
Call HandleError("MapGoDirection")
End Sub

Private Sub fraMapControls_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
fraMapControls.Top = y
fraMapControls.Left = x
End Sub

Private Sub mnuMapPopUpItem_Click(Index As Integer)
On Error GoTo Error:

Select Case Index
    Case 0: 'up
        Call MapGoDirection(CellRoom(nMapLastCellIndex, 1), CellRoom(nMapLastCellIndex, 2), "U")
    Case 1: 'down
        Call MapGoDirection(CellRoom(nMapLastCellIndex, 1), CellRoom(nMapLastCellIndex, 2), "D")
    Case 2: 'redraw
        Call MapStartMapping(CellRoom(nMapLastCellIndex, 1), CellRoom(nMapLastCellIndex, 2))
End Select

Exit Sub

Error:
Call HandleError("mnuMapPopUpItem_Click")
End Sub

Private Sub cmdMapPresetSelect_Click(Index As Integer)
Dim nStart As Integer, x As Integer, sSectionName As String
Dim cReg As clsRegistryRoutines

Set cReg = New clsRegistryRoutines

If InStr(1, frmMain.lblDatVer.Caption, "-") = 0 Then
    sSectionName = "Custom_Presets"
Else
    sSectionName = RemoveCharacter(frmMain.lblDatVer.Caption, " ") & "_Presets"
End If

cReg.hkey = HKEY_LOCAL_MACHINE
cReg.KeyRoot = "Software\MMUD Explorer\Presets"
cReg.Subkey = sSectionName

Select Case Index
    Case 0: nStart = 0
    Case 1: nStart = 10
    Case 2: nStart = 20
    Case 3: nStart = 30
    Case 4: nStart = 40
    Case Else: Exit Sub
End Select

For x = nStart To nStart + 9
    cmdMapPreset(x Mod 10).Caption = cReg.GetRegistryValue("Name" & x, "unset")
    cmdMapPreset(x Mod 10).Tag = x
Next x
End Sub

Private Sub cmdQ_Click(Index As Integer)

Select Case Index
    Case 0:
        MsgBox "Clicking ""Options"" again will hide the options window and refresh the map.", vbInformation
    Case 1:
        MsgBox "You can also use your keypad to move around on the map.", vbInformation
    Case 2:
        MsgBox "This will allow the 'Main' MMUD Explorer window to overlap the" & vbCrLf _
            & "map window (when set to 'Always on Top') when double clicking one" & vbCrLf _
            & "of the references below.  (Click the Map window again to re-activate" & vbCrLf _
            & "the 'Always on Top' functionality.)", vbInformation
End Select

End Sub

Private Sub cmdMapFindText_Click(Index As Integer)
On Error GoTo Error:
Dim sTemp As String

If tabRooms.RecordCount = 0 Then Exit Sub

tabRooms.Index = "idxRooms"
If Index = 0 Or nMapLastFind(0) = 0 Or nMapLastFind(1) = 0 Then
    sTemp = InputBox("Enter text to search for.", "Search for room name", sMapSearch)
    If sTemp = "" Then Exit Sub
    
    sMapSearch = sTemp
    nMapLastFind(2) = 0
    tabRooms.MoveFirst
Else
    tabRooms.Seek "=", nMapLastFind(0), nMapLastFind(1)
    If tabRooms.NoMatch Then
        MsgBox "Room " & nMapLastFind(0) & "/" & nMapLastFind(1) & " not found.", vbInformation
        Exit Sub
    End If
    tabRooms.MoveNext
End If
DoEvents

fraOptions.Visible = False
Me.Enabled = False
frmMain.Enabled = False

bMapCancelFind = False

If chkMapOptions(6).Value = 0 Then Call SetTopMostWindow(Me.hwnd, False)

Load frmProgressBar
Call frmProgressBar.SetRange(tabRooms.RecordCount)
frmProgressBar.ProgressBar.Value = nMapLastFind(2)
frmProgressBar.lblCaption.Caption = "Searching for Room Name ..."
Set frmProgressBar.objFormOwner = Me

DoEvents
frmProgressBar.Show vbModeless, Me
DoEvents

Do Until tabRooms.EOF Or bMapCancelFind
    If InStr(1, LCase(tabRooms.Fields("Name")), LCase(sMapSearch)) > 0 Then Exit Do
    Call frmProgressBar.IncreaseProgress
    tabRooms.MoveNext
    DoEvents
Loop
If tabRooms.EOF Then
    nMapLastFind(0) = 0
    nMapLastFind(1) = 0
    nMapLastFind(2) = 0
    MsgBox "Name not found.", vbInformation
    GoTo out:
End If

nMapLastFind(0) = tabRooms.Fields("Map Number")
nMapLastFind(1) = tabRooms.Fields("Room Number")
nMapLastFind(2) = frmProgressBar.ProgressBar.Value

If Not bMapCancelFind Then
    Call MapStartMapping(tabRooms.Fields("Map Number"), tabRooms.Fields("Room Number"))
End If

out:
On Error Resume Next
Unload frmProgressBar
Me.Enabled = True
If chkMapOptions(6).Value = 0 Then Call SetTopMostWindow(Me.hwnd, True)
frmMain.Enabled = True
Me.SetFocus
Exit Sub

Error:
Call HandleError("cmdMapFindText_Click")
Resume out:
End Sub

Private Sub cmdMapShowUnused_Click()
Dim x As Integer

If cmdMapShowUnused.Caption = "S&how Unused Blocks" Then
    For x = 1 To sMapSECorner
        lblRoomCell(x).Visible = True
    Next
    cmdMapShowUnused.Caption = "&Hide Unused Blocks"
Else
    For x = 1 To sMapSECorner
        If CellRoom(x, 1) = 0 Then lblRoomCell(x).Visible = False
    Next
    cmdMapShowUnused.Caption = "S&how Unused Blocks"
End If

fraOptions.Visible = False
End Sub

Private Sub cmdViewMapLegend_Click()
On Error GoTo Error:

If cmdViewMapLegend.Tag = "1" Then
    Unload frmMapLegend
    cmdViewMapLegend.Tag = "0"
Else
    cmdViewMapLegend.Tag = "1"
    frmMapLegend.Show vbModeless, Me
    Set frmMapLegend.objFormOwner = Me
    
    If chkMapOptions(6).Value = 0 Then Call SetTopMostWindow(Me.hwnd, True)
    
    'Call SetOwner(frmMapLegend.hwnd, Me.hwnd)
'    If chkMapOptions(6).Value = 1 Then
'        lR = SetTopMostWindow(frmMap.hwnd, False)
'    Else
'        lR = SetTopMostWindow(frmMap.hwnd, True)
'    End If
End If
'fraOptions.Visible = False

Exit Sub

Error:
Call HandleError("cmdViewMapLegend_Click")
End Sub


Private Sub cmdDrawMap_Click(Index As Integer)
fraOptions.Visible = False
If Index = 0 Then
    If Val(txtRoomMap.Text) > 32767 Then txtRoomMap.Text = 32767
    If Val(txtRoomRoom.Text) > 32767 Then txtRoomRoom.Text = 32767
    Call MapStartMapping(Val(txtRoomMap.Text), Val(txtRoomRoom.Text))
Else
    Call MapStartMapping(nMapLastMap, nMapLastRoom)
End If
End Sub

Private Sub ResizeMap()
On Error GoTo Error:

If Me.WindowState = vbMinimized Then Me.WindowState = vbNormal

Select Case cmbMapSize.ListIndex
    Case 1: '30x30
        sMapSECorner = 900
        nMapRowLength = 30
        If nMapCenterCell > sMapSECorner Then nMapCenterCell = 436
        If nMapCenterCell = 0 Then nMapCenterCell = 436
        Me.Height = 9135 + TITLEBAR_OFFSET
        Me.Width = 7440
        lvMapLoc.Top = 7680
        lvMapLoc.Width = 7245
        picMap.Width = 7245
        picMap.Height = 7245
        lvMapLoc.ColumnHeaders(1).Width = 6800
        
    Case 2: '40x40
        sMapSECorner = 1600
        nMapRowLength = 40
        If nMapCenterCell > sMapSECorner Then nMapCenterCell = 779
        If nMapCenterCell = 0 Then nMapCenterCell = 779
        Me.Height = 11535 + TITLEBAR_OFFSET
        Me.Width = 9825
        lvMapLoc.Top = 10080
        lvMapLoc.Width = 9645
        picMap.Width = 9645
        picMap.Height = 9645
        lvMapLoc.ColumnHeaders(1).Width = 9200
        
    Case Else: '20x20
        sMapSECorner = 400
        nMapRowLength = 20
        If nMapCenterCell > sMapSECorner Then nMapCenterCell = 219
        If nMapCenterCell = 0 Then nMapCenterCell = 219
        Me.Height = 6735 + TITLEBAR_OFFSET
        Me.Width = 5055
        lvMapLoc.Top = 5280
        lvMapLoc.Width = 4845
        picMap.Width = 4845
        picMap.Height = 4845
        lvMapLoc.ColumnHeaders(1).Width = 4400
        
End Select

ReDim CellRoom(1 To sMapSECorner, 1 To 2) As Long
ReDim UnchartedCells(1 To sMapSECorner) As Integer

Dim i As Integer

For i = lblRoomCell.UBound + 1 To sMapSECorner
    Load lblRoomCell(i)
Next

For i = lblRoomCell.LBound To lblRoomCell.UBound
    Dim row As Integer
    Dim col As Integer
    
    row = Int((i - 1) / nMapRowLength)
    col = (i - 1) Mod nMapRowLength
    
    lblRoomCell(i).Left = 4 + 16 * (col)
    lblRoomCell(i).Top = 4 + 16 * (row)
Next

out:
Exit Sub
Error:
Call HandleError("ResizeMap")
Resume out:

End Sub
Public Sub MapStartMapping(ByVal nStartMap As Long, ByVal nStartRoom As Long, Optional nCenterCell As Integer)
On Error GoTo Error:
Dim x As Integer, nMapSize As Integer, bCheckAgain As Boolean, y As Integer

If bMapStillMapping Then Exit Sub

tabRooms.Index = "idxRooms"
tabRooms.Seek "=", nStartMap, nStartRoom
If tabRooms.NoMatch Then
    MsgBox "Room " & nStartMap & "/" & nStartRoom & " was not found.", vbInformation
    'Me.Caption = "Rooms"
    Exit Sub
Else
    Me.Caption = "Map -- " & tabRooms.Fields("Name") & " (" & nStartMap & "/" & nStartRoom & ")  "
End If

'If chkMapOptions(6).Value = 0 Then Call SetTopMostWindow(Me.hwnd, True)

If Not nMapStartRoom = nStartRoom Then
    nMapLastRoom = nMapStartRoom
    nMapLastMap = nMapStartMap
End If

bMapStillMapping = True
Call LockWindowUpdate(Me.hwnd)

'picMap.Visible = False
picMap.Cls
Me.MousePointer = vbHourglass
DoEvents
'20x20
'sMapSECorner = 400
'nMapRowLength = 20
If Not nCenterCell = 0 Then nMapCenterCell = nCenterCell
'If nMapCenterCell > sMapSECorner Then nMapCenterCell = 210

For x = 1 To 900
    TTlbl.DelToolTip picMap.hwnd, 0
    lblRoomCell(x).BackColor = &HFFFFFF
    lblRoomCell(x).Visible = False
    lblRoomCell(x).Tag = 0
    UnchartedCells(x) = 0
    CellRoom(x, 1) = 0
    CellRoom(x, 2) = 0
Next x

Call ResizeMap

StopBuild = False

nMapStartRoom = nStartRoom
nMapStartMap = nStartMap

CellRoom(nMapCenterCell, 1) = nMapStartMap
CellRoom(nMapCenterCell, 2) = nMapStartRoom

fraPresets.Visible = False
Call MapMapExits(nMapCenterCell, nMapStartRoom, nMapStartMap)

DoEvents
again:
bCheckAgain = False
For x = 1 To sMapSECorner
    If StopBuild = True Then GoTo Cancel:
    If UnchartedCells(x) = 1 Then
        For y = 1 To sMapSECorner
            If Not CellRoom(x, 1) = 0 Then
                If Not x = y Then
                    If CellRoom(y, 2) = CellRoom(x, 2) Then
                        If CellRoom(y, 1) = CellRoom(x, 1) Then
                            CellRoom(x, 2) = 0
                            CellRoom(x, 1) = 0
                            UnchartedCells(x) = 0
                            GoTo skiproom:
                        End If
                    End If
                End If
            End If
        Next y
        Call MapMapExits(x, CellRoom(x, 2), CellRoom(x, 1))
        bCheckAgain = True
    End If
skiproom:
    'DoEvents
Next x

'For x = nMapCenterCell To 1 Step -1 '1 To sMapSECorner
'    If StopBuild = True Then GoTo Cancel:
'    If UnchartedCells(x) = 1 Then
'        For y = 1 To sMapSECorner
'            If Not CellRoom(x, 1) = 0 Then
'                If Not x = y Then
'                    If CellRoom(y, 2) = CellRoom(x, 2) Then
'                        If CellRoom(y, 1) = CellRoom(x, 1) Then
'                            CellRoom(x, 2) = 0
'                            CellRoom(x, 1) = 0
'                            UnchartedCells(x) = 0
'                            GoTo skiproom:
'                        End If
'                    End If
'                End If
'            End If
'        Next y
'        Call MapMapExits(x, CellRoom(x, 2), CellRoom(x, 1))
'        bCheckAgain = True
'    End If
'skiproom:
'    'DoEvents
'Next x
'For x = nMapCenterCell To sMapSECorner
'    If StopBuild = True Then GoTo Cancel:
'    If UnchartedCells(x) = 1 Then
'        For y = 1 To sMapSECorner
'            If Not CellRoom(x, 1) = 0 Then
'                If Not x = y Then
'                    If CellRoom(y, 2) = CellRoom(x, 2) Then
'                        If CellRoom(y, 1) = CellRoom(x, 1) Then
'                            CellRoom(x, 2) = 0
'                            CellRoom(x, 1) = 0
'                            UnchartedCells(x) = 0
'                            GoTo skiproom2:
'                        End If
'                    End If
'                End If
'            End If
'        Next y
'        Call MapMapExits(x, CellRoom(x, 2), CellRoom(x, 1))
'        bCheckAgain = True
'    End If
'skiproom2:
'    'DoEvents
'Next x

If bCheckAgain Then GoTo again:

Call MapDrawOnRoom(lblRoomCell(nMapCenterCell), drSquare, 4, BrightBlue)

DoEvents
cmdMapShowUnused.Caption = "S&how Unused Blocks"
For x = 1 To sMapSECorner
    If Not CellRoom(x, 1) = 0 Then lblRoomCell(x).Visible = True
Next x
DoEvents

Call lblRoomCell_MouseDown(nMapCenterCell, IIf(bMapSwapButtons, 2, 1), 0, 0, 0)
'picMap.Visible = True

Cancel:
On Error Resume Next
Me.MousePointer = vbDefault
bMapStillMapping = False
Call LockWindowUpdate(0&)

Exit Sub
Error:
Call HandleError("MapStartMapping")
Resume Cancel:
End Sub
Private Sub MapMapExits(Cell As Integer, Room As Long, Map As Long)
Dim ActivatedCell As Integer, x As Integer
Dim rc As RECT, ToolTipString As String, sText As String, y As Long
Dim sRemote As String, sMonsters As String, sArray() As String, sPlaced As String
Dim RoomExit As RoomExitType, sLook As String, nExitType As Integer, sRoomCMDs As String

On Error GoTo Error:

'=============================================================================
'
'                 NOTE: THIS ROUTINE IS ON BOTH frmMain AND frmMap
'
'=============================================================================

CellRoom(Cell, 1) = Map
CellRoom(Cell, 2) = Room

tabRooms.Index = "idxRooms"
tabRooms.Seek "=", Map, Room
If tabRooms.NoMatch Then
    UnchartedCells(Cell) = 2
    Call MapDrawOnRoom(lblRoomCell(Cell), drSquare, 8, BrightRed)
    ToolTipString = "Map " & Map & " Room " & Room
    rc.Left = lblRoomCell(Cell).Left
    rc.Top = lblRoomCell(Cell).Top
    rc.Bottom = (lblRoomCell(Cell).Top + lblRoomCell(Cell).Height)
    rc.Right = (lblRoomCell(Cell).Left + lblRoomCell(Cell).Width)
    TTlbl.SetToolTipItem picMap.hwnd, 0, rc.Left, rc.Top, rc.Right, rc.Bottom, ToolTipString, False
    Exit Sub
End If

ToolTipString = Map & "/" & Room & " - " & tabRooms.Fields("Name")

If chkMapOptions(4).Value = 0 And tabRooms.Fields("CMD") > 0 Then
    sRoomCMDs = vbCrLf & vbCrLf & "Room commands: " & GetTextblockCMDS(tabRooms.Fields("CMD"))
    Call MapDrawOnRoom(lblRoomCell(Cell), drSquare, 6, BrightGreen)
Else
    sRoomCMDs = ""
End If

If chkMapOptions(3).Value = 0 And tabRooms.Fields("NPC") > 0 Then
    ToolTipString = ToolTipString & vbCrLf & "NPC: " & GetMonsterName(tabRooms.Fields("NPC"), bHideRecordNumbers)
    Call MapDrawOnRoom(lblRoomCell(Cell), drOpenCircle, 2, BrightRed)
End If

If Len(tabRooms.Fields("Placed")) > 1 Then
    sArray() = Split(tabRooms.Fields("Placed"), ",")
    If UBound(sArray()) >= 0 Then
        For x = 0 To UBound(sArray())
            If Val(sArray(x)) > 0 Then
                If Not sPlaced = "" Then sPlaced = sPlaced & ", "
                sPlaced = sPlaced & GetItemName(Val(sArray(0)), bHideRecordNumbers)
            End If
        Next x
        ToolTipString = ToolTipString & vbCrLf & "Placed Items: " & sPlaced
        'Call MapDrawOnRoom(lblRoomCell(Cell), drOpenCircle, 2, BrightRed)
    End If
    Erase sArray()
End If

If chkMapOptions(2).Value = 0 And Not tabRooms.Fields("Lair") = Chr(0) Then
    sMonsters = GetMultiMonsterNames(Mid(tabRooms.Fields("Lair"), InStr(1, tabRooms.Fields("Lair"), ":") + 2), bHideRecordNumbers)
    sMonsters = "Also Here " & Left(tabRooms.Fields("Lair"), InStr(1, tabRooms.Fields("Lair"), ":") + 1) & sMonsters
    Call MapDrawOnRoom(lblRoomCell(Cell), drCircle, 5, BrightMagenta)
End If

If tabRooms.Fields("Shop") > 2 Then
    ToolTipString = ToolTipString & vbCrLf & "Shop: " & GetShopName(tabRooms.Fields("Shop"), bHideRecordNumbers) '& "(" & tabRooms.Fields("Shop") & ")"
End If

If tabRooms.Fields("Spell") > 0 Then
    ToolTipString = ToolTipString & vbCrLf & "Room Spell: " & GetSpellName(tabRooms.Fields("Spell"), bHideRecordNumbers)
End If

'map exits
For x = 0 To 9
    Select Case x
        Case 0: sLook = "N"
        Case 1: sLook = "S"
        Case 2: sLook = "E"
        Case 3: sLook = "W"
        Case 4: sLook = "NE"
        Case 5: sLook = "NW"
        Case 6: sLook = "SE"
        Case 7: sLook = "SW"
        Case 8: sLook = "U"
        Case 9: sLook = "D"
    End Select
    
    nExitType = 0
    If Left(tabRooms.Fields(sLook), 6) = "Action" Then
        sRemote = sRemote & vbCrLf & tabRooms.Fields(sLook)
        If chkMapOptions(4).Value = 0 Then Call MapDrawOnRoom(lblRoomCell(Cell), drSquare, 6, BrightGreen)
    
    ElseIf Not Val(tabRooms.Fields(sLook)) = 0 Then
        RoomExit = ExtractMapRoom(tabRooms.Fields(sLook))
        
        If Len(RoomExit.ExitType) > 2 Then
            Select Case Left(RoomExit.ExitType, 5)
                Case "(Key:": nExitType = 2
                Case "(Item": nExitType = 3
                Case "(Toll": nExitType = 4
                Case "(Hidd": nExitType = 6
                Case "(Door": nExitType = 7
                Case "(Trap": nExitType = 9
                Case "(Text": nExitType = 10
                Case "(Gate": nExitType = 11
                Case "Actio": nExitType = 12
                Case "(Clas": nExitType = 13
                Case "(Race": nExitType = 14
                Case "(Leve": nExitType = 15
                Case "(Time": nExitType = 16
                Case "(Tick": nExitType = 17
                Case "(Max ": nExitType = 18
                Case "(Bloc": nExitType = 19
                Case "(Alig": nExitType = 20
                Case "(Dela": nExitType = 21
                Case "(Cast": nExitType = 22
                Case "(Abil": nExitType = 23
                Case "(Spel": nExitType = 24
            End Select
        End If
        If Not RoomExit.Map = Map Then nExitType = 8 'map change
        
        'sText = sText & vbCrLf & sLook & ": " & RoomExit.Map & "/" & RoomExit.Room

        'note order of case'ings is important here
        Select Case nExitType
            Case 2: 'key
                y = ExtractValueFromString(RoomExit.ExitType, "Key: ")
                sText = sText & vbCrLf & sLook & " (Key: " _
                    & GetItemName(y, bHideRecordNumbers) _
                    & " " & Mid(RoomExit.ExitType, InStr(1, RoomExit.ExitType, y) + Len(CStr(y)) + 1)

                ActivatedCell = MapActivateCell(Cell, x, nExitType)
                If ActivatedCell = -1 Then GoTo skip:

                If chkMapOptions(1).Value = 1 And nExitType = 6 Then GoTo skip:

                CellRoom(ActivatedCell, 1) = Map
                CellRoom(ActivatedCell, 2) = RoomExit.Room
                If UnchartedCells(ActivatedCell) = 0 Then UnchartedCells(ActivatedCell) = 1

            Case 3: 'item
                y = ExtractValueFromString(RoomExit.ExitType, "Item: ")
                sText = sText & vbCrLf & sLook & " (Item): " _
                    & GetItemName(y, bHideRecordNumbers) _
                    & " " & Mid(RoomExit.ExitType, InStr(1, RoomExit.ExitType, y) + Len(CStr(y)) + 1)

                ActivatedCell = MapActivateCell(Cell, x, nExitType)
                If ActivatedCell = -1 Then GoTo skip:

                If chkMapOptions(1).Value = 1 And nExitType = 6 Then GoTo skip:

                CellRoom(ActivatedCell, 1) = Map
                CellRoom(ActivatedCell, 2) = RoomExit.Room
                If UnchartedCells(ActivatedCell) = 0 Then UnchartedCells(ActivatedCell) = 1
                
            Case 8: 'map change
                ActivatedCell = MapActivateCell(Cell, x, nExitType)
                If ActivatedCell = -1 Then GoTo skip:
                If chkMapOptions(0).Value = 1 Then
                    CellRoom(ActivatedCell, 1) = RoomExit.Map
                    CellRoom(ActivatedCell, 2) = RoomExit.Room
                    If UnchartedCells(ActivatedCell) = 0 Then UnchartedCells(ActivatedCell) = 1
                End If
            Case 12: 'action
                sRemote = sRemote & vbCrLf & tabRooms.Fields(sLook)
                If chkMapOptions(4).Value = 0 Then Call MapDrawOnRoom(lblRoomCell(Cell), drSquare, 6, BrightGreen)
            Case Is > 0:
                sText = sText & vbCrLf & sLook & ": " & RoomExit.ExitType
                ActivatedCell = MapActivateCell(Cell, x, nExitType)
                If ActivatedCell = -1 Then GoTo skip:
                
                If chkMapOptions(1).Value = 1 And nExitType = 6 Then GoTo skip:
                
                CellRoom(ActivatedCell, 1) = Map
                CellRoom(ActivatedCell, 2) = RoomExit.Room
                If UnchartedCells(ActivatedCell) = 0 Then UnchartedCells(ActivatedCell) = 1

            Case Else:
                ActivatedCell = MapActivateCell(Cell, x, nExitType) 'nExitType)
                If ActivatedCell = -1 Then GoTo skip:
                CellRoom(ActivatedCell, 1) = Map
                CellRoom(ActivatedCell, 2) = RoomExit.Room
                If UnchartedCells(ActivatedCell) = 0 Then UnchartedCells(ActivatedCell) = 1
        End Select
    End If
skip:
Next x

'set color of this room
If Val(tabRooms.Fields("U")) = 0 And Val(tabRooms.Fields("D")) = 0 Then
    lblRoomCell(Cell).BackColor = &HC0C0C0   '&H0& '-- nothing
ElseIf Val(tabRooms.Fields("U")) > 0 And Val(tabRooms.Fields("D")) = 0 Then
    lblRoomCell(Cell).BackColor = &HFF00& '-- up
ElseIf Val(tabRooms.Fields("U")) = 0 And Val(tabRooms.Fields("D")) > 0 Then
    lblRoomCell(Cell).BackColor = &HFFFF& '-- down
Else
    lblRoomCell(Cell).BackColor = &HFFFF00 '-- both
End If

If chkMapOptions(5).Value = 0 Then
    ToolTipString = ToolTipString & sText & IIf(sRemote = "", "", vbCrLf & sRemote) & sRoomCMDs _
        & IIf(sMonsters = "", "", vbCrLf & vbCrLf & sMonsters)
    
    rc.Left = lblRoomCell(Cell).Left
    rc.Top = lblRoomCell(Cell).Top
    rc.Bottom = (lblRoomCell(Cell).Top + lblRoomCell(Cell).Height)
    rc.Right = (lblRoomCell(Cell).Left + lblRoomCell(Cell).Width)
    TTlbl.SetToolTipItem picMap.hwnd, 0, rc.Left, rc.Top, rc.Right, rc.Bottom, ToolTipString, False
End If

UnchartedCells(Cell) = 2

Exit Sub

Error:
Call HandleError("MapMapExits")
End Sub

Private Function MapActivateCell(ByVal FromCell As Integer, ByVal direction As Integer, ByVal ExitType As Integer) As Integer

Dim temp As Integer, row As Integer, col As Integer, drLine As Integer
Dim LineColor As Long

'figure out which cell is to be activated
On Error GoTo Error:

row = Int((FromCell - 1) / nMapRowLength)
col = (FromCell - 1) Mod nMapRowLength

Select Case direction
    Case 0: 'north
        row = row - 1
        drLine = drLineN
    Case 1: 'south
        row = row + 1
        drLine = drLineS
    Case 2: 'east
        col = col + 1
        drLine = drLineE
    Case 3: 'west
        col = col - 1
        drLine = drLineW
    Case 4: 'ne
        row = row - 1
        col = col + 1
        drLine = drLineNe
    Case 5: 'nw
        row = row - 1
        col = col - 1
        drLine = drLineNw
    Case 6: 'se
        row = row + 1
        col = col + 1
        drLine = drLineSe
    Case 7: 'sw
        row = row + 1
        col = col - 1
        drLine = drLineSw
    Case Else:
        GoTo DontActivate
End Select

If (row < 0) Or (row >= nMapRowLength) Or (col < 0) Or (col >= nMapRowLength) Then
    Call MapDrawOnRoom(lblRoomCell(FromCell), drLine, 4, Grey)
    GoTo DontActivate
End If

MapActivateCell = 1 + (row * nMapRowLength) + col

'set line mode
'ScaleMode = vbPixels
DrawWidth = 4

'pick line color
Select Case ExitType
    Case 2: LineColor = 10    'l green - key
    Case 3: LineColor = 10    'l green - item
    Case 4: LineColor = 10    'l green - toll
    Case 5: LineColor = 11    'l cyan - action
    Case 6: LineColor = 5     'd magenta - hidden
    Case 7: LineColor = 9     'l blue - door/gate
    Case 8: LineColor = 13    'l magenta - map change
    Case 9: LineColor = 12    'l red - trap/spell trap
    Case 10: LineColor = 14   'l yellow - text
    Case 11: LineColor = 9    'l blue - door/gate
    Case 12: LineColor = 11   'l cyan - remote action
    Case 13: LineColor = 4    'd red - class
    Case 14: LineColor = 4    'd red - race
    Case 15: LineColor = 4    'd red - level
    Case 16: LineColor = 2    'gray - timed
    Case 20: LineColor = 4    'd red - alignment
    Case 23: LineColor = 4    'd red - ability
    Case 24: LineColor = 12   'l red - trap/spell trap
    Case Else: LineColor = 8 '0  'black - anything else
End Select
    
'If chkNoColors.value = 1 Then LineColor = 0
'If chkNoLineColors.value = 1 Then LineColor = 0

'draw the line
Call MapDrawOnRoom(lblRoomCell(FromCell), drLine, 4, LineColor)

'if the cell to be activated has already been mapped, dont map it again
If UnchartedCells(MapActivateCell) = 2 Then GoTo DontActivate:

Select Case ExitType
    Case 12: MapActivateCell = -1 'if it's a remote action, dont map it
    Case 8: 'if it's a map change, check to see if it should be mapped
        If chkMapOptions(0).Value = 1 Then
            lblRoomCell(MapActivateCell).BackColor = &H0
        Else
            MapActivateCell = -1
        End If
    Case Else: lblRoomCell(MapActivateCell).BackColor = &H0
End Select

Exit Function
DontActivate:
MapActivateCell = -1

Exit Function

Error:
Call HandleError("MapActivateCell")

End Function

Private Sub MapDrawOnRoom(ByRef oLabel As Label, ByVal drDrawType As EnumDrawRoom, ByVal nSize As Integer, ByVal nColor As QBColorCode)
Dim x1 As Integer, x2 As Integer, y1 As Integer, y2 As Integer
Dim nTemp As Integer

nTemp = picMap.DrawWidth

'If chkNoColors.value = 1 Then nColor = Black

Select Case drDrawType
    Case 0: 'square
        picMap.DrawWidth = nSize
        x1 = oLabel.Left
        y1 = oLabel.Top
        x2 = oLabel.Left + oLabel.Width
        y2 = oLabel.Top + oLabel.Height
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor), BF
        
    Case 1: 'star
        picMap.DrawWidth = nSize
        '/
        x1 = oLabel.Left - 4
        y1 = oLabel.Top + oLabel.Height + 4
        x2 = oLabel.Left + 4
        y2 = oLabel.Top - 4
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor)
        
        '\
        x1 = x2
        y1 = y2
        x2 = oLabel.Left + oLabel.Width + 4
        y2 = oLabel.Top + oLabel.Height + 4
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor)
        
        '\
        x1 = x2
        y1 = y2
        x2 = oLabel.Left - 4
        y2 = oLabel.Top
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor)
        
        '-
        x1 = x2
        y1 = y2
        x2 = oLabel.Left + oLabel.Width + 4
        y2 = y1
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor)
        
        '/
        x1 = x2
        y1 = y2
        x2 = oLabel.Left - 4
        y2 = oLabel.Top + oLabel.Height + 4
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor)
        
    Case 2: 'open circle
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 4
        y1 = oLabel.Top + 4
        picMap.Circle (x1, y1), 8, QBColor(nColor)
      
     Case 3: 'up
        picMap.DrawWidth = nSize
        x1 = oLabel.Left
        y1 = oLabel.Top
        x2 = oLabel.Left + oLabel.Width
        y2 = oLabel.Top + 2
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor), B
        
     Case 4: 'down
        picMap.DrawWidth = nSize
        x1 = oLabel.Left - 1
        y1 = oLabel.Top + oLabel.Height - 1
        x2 = oLabel.Left + oLabel.Width
        y2 = y1 + 2
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor), B
    
    Case 5: 'circle
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 4
        y1 = oLabel.Top + 4
        picMap.Circle (x1, y1), 5, QBColor(nColor)
    
    Case 6: 'LineN
        'If chkNoLineColors.value = 1 Then nColor = Black
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 4
        y1 = oLabel.Top + 4
        x2 = x1
        y2 = y1 - 8
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor), BF
        
    Case 7: 'LineS
        'If chkNoLineColors.value = 1 Then nColor = Black
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 4
        y1 = oLabel.Top + 4
        x2 = x1
        y2 = y1 + 9
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor), BF
        
    Case 8: 'LineE
        'If chkNoLineColors.value = 1 Then nColor = Black
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 4
        y1 = oLabel.Top + 4
        x2 = x1 + 9
        y2 = y1
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor), BF
        
    Case 9: 'LineW
        'If chkNoLineColors.value = 1 Then nColor = Black
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 4
        y1 = oLabel.Top + 4
        x2 = x1 - 8
        y2 = y1
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor), BF
        
    Case 10: 'LineNE
        'If chkNoLineColors.value = 1 Then nColor = Black
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 4
        y1 = oLabel.Top + 4
        x2 = x1 + 8
        y2 = y1 - 8
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor)
        
    Case 11: 'LineNW
        'If chkNoLineColors.value = 1 Then nColor = Black
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 5
        y1 = oLabel.Top + 5
        x2 = x1 - 8
        y2 = y1 - 8
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor)
        
    Case 12: 'LineSE
        'If chkNoLineColors.value = 1 Then nColor = Black
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 5
        y1 = oLabel.Top + 5
        x2 = x1 + 8
        y2 = y1 + 8
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor)
    
    Case 13: 'LineSW
        'If chkNoLineColors.value = 1 Then nColor = Black
        picMap.DrawWidth = nSize
        x1 = oLabel.Left + 4
        y1 = oLabel.Top + 4
        x2 = x1 - 8
        y2 = y1 + 8
        picMap.Line (x1, y1)-(x2, y2), QBColor(nColor)
        
End Select

picMap.DrawWidth = nTemp
End Sub

Private Sub MapGetRoomLoc(ByVal nMapNumber As Long, ByVal nRoomNumber As Long)
On Error GoTo Error:
Dim x As Long, sLook As String, nExitType As Integer, RoomExit As RoomExitType, oLI As ListItem, RoomExit2 As RoomExitType
Dim nRecNum As Long, y As Long, sNumbers As String, sCommand As String, nMap As Long, nRoom As Long, sChar As String
Dim sArray() As String

'=============================================================================
'
'                 NOTE: THIS ROUTINE IS ON BOTH frmMain AND frmMap
'
'=============================================================================

tabRooms.Index = "idxRooms"
tabRooms.Seek "=", nMapNumber, nRoomNumber
If tabRooms.NoMatch Then
    MsgBox "Room (" & nMapNumber & "/" & nRoomNumber & ") was not found."
    Exit Sub
End If

lvMapLoc.ColumnHeaders(1).Text = "References [" & tabRooms.Fields("Name") & " (" & nMapNumber & "/" & nRoomNumber & ")]"

If tabRooms.Fields("CMD") > 0 Then 'chkMapOptions(4).Value = 0 And
    tabTBInfo.Index = "pkTBInfo"
    tabTBInfo.Seek "=", tabRooms.Fields("CMD")
    If tabTBInfo.NoMatch = False Then
        sCommand = tabTBInfo.Fields("Action")
        x = InStr(1, sCommand, "teleport ")
        If x > 0 Then
            Do While x < Len(sCommand)
                x = x + Len("teleport ") 'position x just after the search text
                y = x
                Do While y < Len(sCommand) + 2
                    sChar = Mid(sCommand, y, 1)
                    Select Case sChar
                        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9":
                        Case " ":
                            If y > x And nRoom = 0 Then
                                nRoom = Val(Mid(sCommand, x, y - x))
                                x = y + 1
                            Else
                                nMap = Val(Mid(sCommand, x, y - x))
                                Exit Do
                            End If
                        Case Else:
                            If y > x And nRoom = 0 Then
                                nRoom = Val(Mid(sCommand, x, y - x))
                                Exit Do
                            Else
                                nMap = Val(Mid(sCommand, x, y - x))
                                Exit Do
                            End If
                            Exit Do
                    End Select
                    y = y + 1
                Loop
                
                If Not nRoom = 0 Then
                    If nMap = 0 Then nMap = nMapNumber
                    For Each oLI In lvMapLoc.ListItems
                        If oLI.Tag = nMap & "/" & nRoom Then GoTo skiptele:
                    Next
                    
                    Set oLI = lvMapLoc.ListItems.Add()
                    oLI.Text = "Teleport: " & GetTextblockCMDText("teleport " & nRoom & " " & nMap, sCommand) _
                        & " --> " & GetRoomName(, nMap, nRoom, False)
                    oLI.Tag = nMap & "/" & nRoom
                End If
skiptele:
                nRoom = 0
                nMap = 0
                x = InStr(y, sCommand, "teleport ")
                If x = 0 Then x = Len(sCommand)
            Loop
            tabRooms.Seek "=", nMapNumber, nRoomNumber
        End If
        
        Set oLI = lvMapLoc.ListItems.Add()
        oLI.Text = "Commands: Textblock " & tabRooms.Fields("CMD")
        oLI.Tag = tabRooms.Fields("CMD")
    End If
End If

If chkMapOptions(3).Value = 0 And tabRooms.Fields("NPC") > 0 Then
    Set oLI = lvMapLoc.ListItems.Add()
    oLI.Text = "NPC: " & GetMonsterName(tabRooms.Fields("NPC"), bHideRecordNumbers)
    oLI.Tag = tabRooms.Fields("NPC")
End If

If tabRooms.Fields("Shop") > 0 Then
    Set oLI = lvMapLoc.ListItems.Add()
    oLI.Text = "Shop: " & GetShopName(tabRooms.Fields("Shop"), bHideRecordNumbers) '& "(" & tabRooms.Fields("Shop") & ")"
    oLI.Tag = tabRooms.Fields("Shop")
End If

If tabRooms.Fields("Spell") > 0 Then
    Set oLI = lvMapLoc.ListItems.Add()
    oLI.Text = "Spell: " & GetSpellName(tabRooms.Fields("Spell"), bHideRecordNumbers)
    oLI.Tag = tabRooms.Fields("Spell")
End If

For x = 0 To 9
    Select Case x
        Case 0: sLook = "N"
        Case 1: sLook = "S"
        Case 2: sLook = "E"
        Case 3: sLook = "W"
        Case 4: sLook = "NE"
        Case 5: sLook = "NW"
        Case 6: sLook = "SE"
        Case 7: sLook = "SW"
        Case 8: sLook = "U"
        Case 9: sLook = "D"
    End Select
    
    nExitType = 0
    If Not Val(tabRooms.Fields(sLook)) = 0 Then
        RoomExit = ExtractMapRoom(tabRooms.Fields(sLook))
        
        If Len(RoomExit.ExitType) > 2 Then
            Select Case Left(RoomExit.ExitType, 5)
                Case "(Key:": nExitType = 2
                Case "(Item": nExitType = 3
                Case "(Toll": nExitType = 4
                Case "(Hidd": nExitType = 6
                Case "(Door": nExitType = 7
                Case "(Trap": nExitType = 9
                Case "(Text": nExitType = 10
                Case "(Gate": nExitType = 11
                Case "Actio": nExitType = 12
                Case "(Clas": nExitType = 13
                Case "(Race": nExitType = 14
                Case "(Leve": nExitType = 15
                Case "(Time": nExitType = 16
                Case "(Tick": nExitType = 17
                Case "(Max ": nExitType = 18
                Case "(Bloc": nExitType = 19
                Case "(Alig": nExitType = 20
                Case "(Dela": nExitType = 21
                Case "(Cast": nExitType = 22
                Case "(Abil": nExitType = 23
                Case "(Spel": nExitType = 24
            End Select
        End If
        
        Select Case nExitType
            Case 0:
            Case 2, 3, 17:
                nRecNum = ExtractNumbersFromString(RoomExit.ExitType)
                If nRecNum > 0 Then
                    Set oLI = lvMapLoc.ListItems.Add()
                    oLI.Text = "Item: " & GetItemName(nRecNum, bHideRecordNumbers) '& " (" & nRecNum & ")"
                    oLI.Tag = nRecNum
                End If
            Case 22, 24:
'                nRecNum = ExtractNumbersFromString(RoomExit.ExitType)
'                If nRecNum > 0 Then
'                    Set oLI = lvMapLoc.ListItems.Add()
'                    oLI.Text = "Spell: " & GetSpellName(nRecNum, bHideRecordNumbers) '& " (" & nRecNum & ")"
'                    oLI.Tag = nRecNum
'                End If
                nRecNum = ExtractValueFromString(RoomExit.ExitType, "pre-") ' ExtractNumbersFromString(RoomExit.ExitType)
                If nRecNum > 0 Then
                    Set oLI = lvMapLoc.ListItems.Add()
                    oLI.Text = "Spell: " & GetSpellName(nRecNum, bHideRecordNumbers) '& " (" & nRecNum & ")"
                    oLI.Tag = nRecNum
                End If
                nRecNum = ExtractValueFromString(RoomExit.ExitType, "post-") ' ExtractNumbersFromString(RoomExit.ExitType)
                If nRecNum > 0 Then
                    Set oLI = lvMapLoc.ListItems.Add()
                    oLI.Text = "Spell: " & GetSpellName(nRecNum, bHideRecordNumbers) '& " (" & nRecNum & ")"
                    oLI.Tag = nRecNum
                End If
            Case 12:
                RoomExit2 = ExtractMapRoom(RoomExit.ExitType)
                If RoomExit2.Map > 0 Then
                    sChar = "Action On: " & GetRoomName(, RoomExit2.Map, RoomExit2.Room, False) '& " (" & RoomExit2.Map & "/" & RoomExit2.Room & ")"
                    For Each oLI In lvMapLoc.ListItems
                        If oLI.Text = sChar Then GoTo nextexit:
                    Next
                    Set oLI = lvMapLoc.ListItems.Add()
                    oLI.Text = sChar
                    oLI.Tag = RoomExit2.Map & "/" & RoomExit2.Room
                    tabRooms.Seek "=", nMapNumber, nRoomNumber
                End If
        End Select
    ElseIf Left(tabRooms.Fields(sLook), 6) = "Action" Then
        RoomExit2 = ExtractMapRoom(tabRooms.Fields(sLook))
        If RoomExit2.Map > 0 Then
            sChar = "Action On: " & GetRoomName(, RoomExit2.Map, RoomExit2.Room, False) '& " (" & RoomExit2.Map & "/" & RoomExit2.Room & ")"
            For Each oLI In lvMapLoc.ListItems
                If oLI.Text = sChar Then GoTo nextexit:
            Next
            Set oLI = lvMapLoc.ListItems.Add()
            oLI.Text = sChar
            oLI.Tag = RoomExit2.Map & "/" & RoomExit2.Room
            tabRooms.Seek "=", nMapNumber, nRoomNumber
        End If
    End If
nextexit:
Next x

If chkMapOptions(2).Value = 0 And Len(tabRooms.Fields("Lair")) > 1 Then
    tabMonsters.Index = "pkMonsters"
    sNumbers = Mid(tabRooms.Fields("Lair"), InStr(1, tabRooms.Fields("Lair"), ":") + 2)
    x = 0
    Do While Not InStr(x + 1, sNumbers, ",") = 0
        y = InStr(x + 1, sNumbers, ",")
        
        tabMonsters.Seek "=", Val(Mid(sNumbers, x + 1, y - x - 1))
        If tabMonsters.NoMatch = False Then
            Set oLI = lvMapLoc.ListItems.Add()
            oLI.Text = "Lair: " & tabMonsters.Fields("Name") & IIf(bHideRecordNumbers, "", "(" & tabMonsters.Fields("Number") & ")")
            oLI.Tag = tabMonsters.Fields("Number")
        End If
        x = y
    Loop
End If

If Len(tabRooms.Fields("Placed")) > 1 Then
    sArray() = Split(tabRooms.Fields("Placed"), ",")
    If UBound(sArray()) >= 0 Then
        For x = 0 To UBound(sArray())
            If Val(sArray(x)) > 0 Then
                tabItems.Index = "pkItems"
                tabItems.Seek "=", Val(sArray(0))
                If tabItems.NoMatch = False Then
                    Set oLI = lvMapLoc.ListItems.Add()
                    oLI.Text = "Item: " & tabItems.Fields("Name") & IIf(bHideRecordNumbers, "", "(" & tabItems.Fields("Number") & ")")
                    oLI.Tag = tabItems.Fields("Number")
                End If
            End If
        Next x
    End If
    Erase sArray()
End If

'If lvMapLoc.ListItems.Count > 0 Then
'    Call SortListView(lvMapLoc, 1, ldtstring, True)
'End If

Set oLI = Nothing
Exit Sub
Error:
Call HandleError("MapGetRoomLoc")
Set oLI = Nothing
End Sub

Private Sub cmdOptions_Click()
If fraOptions.Visible = True Then
    fraOptions.Visible = False
    If nMapStartMap > 0 And nMapStartRoom > 0 Then
        Call MapStartMapping(nMapStartMap, nMapStartRoom)
    End If
    Exit Sub
End If

fraPresets.Visible = False
fraOptions.Visible = True
End Sub

Private Sub cmdPresets_Click()
If fraPresets.Visible = True Then
    fraPresets.Visible = False
    Exit Sub
End If

fraOptions.Visible = False
fraPresets.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

Call WriteINI("Settings", "ExMapFollowMap", chkMapOptions(0).Value)
Call WriteINI("Settings", "ExMapNoHidden", chkMapOptions(1).Value)
Call WriteINI("Settings", "ExMapNoLairs", chkMapOptions(2).Value)
Call WriteINI("Settings", "ExMapNoNPC", chkMapOptions(3).Value)
Call WriteINI("Settings", "ExMapNoCMD", chkMapOptions(4).Value)
Call WriteINI("Settings", "ExMapNoTooltips", chkMapOptions(5).Value)
Call WriteINI("Settings", "MapExternalOnTop", chkMapOptions(6).Value)
Call WriteINI("Settings", "ExMapSize", cmbMapSize.ListIndex)
Call WriteINI("Settings", "ExMapMainOverlap", chkMapOptions(8).Value)

Set TTlbl = Nothing

If Not Me.WindowState = vbMinimized And Not Me.WindowState = vbMaximized Then
    Call WriteINI("Settings", "ExMapTop", Me.Top)
    Call WriteINI("Settings", "ExMapLeft", Me.Left)
End If

If Not bAppTerminating Then
    If frmMain.WindowState = vbMinimized Then frmMain.WindowState = frmMain.nWindowState
    frmMain.SetFocus
End If

End Sub

Private Sub lblRoomCell_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo Error:

nMapLastCellIndex = Index
lvMapLoc.ListItems.clear

If CellRoom(Index, 1) = 0 Then
    If Button = 2 And Shift = 1 Then
        Call MapStartMapping(nMapStartMap, nMapStartRoom, Index)
        Exit Sub
    Else
        Exit Sub
    End If
End If

If bMapSwapButtons Then
    If Button = 2 Then
        Button = 1
    ElseIf Button = 1 Then
        Button = 2
    End If
End If

If Button = 1 Then
    Call MapGetRoomLoc(CellRoom(Index, 1), CellRoom(Index, 2))
ElseIf Button = 2 Then
    fraOptions.Visible = False
    If lblRoomCell(Index).BackColor = &HFF00& Then '-- up
        Call PopUpMapMenu(True, False)
    ElseIf lblRoomCell(Index).BackColor = &HFFFF& Then '-- down
        Call PopUpMapMenu(False, True)
    ElseIf lblRoomCell(Index).BackColor = &HFFFF00 Then '-- both
        Call PopUpMapMenu(True, True)
    Else
        If Shift = 1 Then
            Call MapStartMapping(nMapStartMap, nMapStartRoom, Index)
        Else
            Call MapStartMapping(CellRoom(Index, 1), CellRoom(Index, 2))
        End If
    End If
End If

Exit Sub
Error:
Call HandleError

End Sub

Public Sub PopUpMapMenu(ByVal bUp As Boolean, bDown As Boolean)
On Error GoTo Error:


If bUp Then mnuMapPopUpItem(0).Visible = True Else mnuMapPopUpItem(0).Visible = False
If bDown Then mnuMapPopUpItem(1).Visible = True Else mnuMapPopUpItem(1).Visible = False

DoEvents
PopupMenu mnuMapPopUp

Exit Sub

Error:
Call HandleError("PopUpMapMenu")

End Sub

Private Sub lvMapLoc_DblClick()
Dim lR As Long

On Error GoTo Error:

If lvMapLoc.ListItems.Count = 0 Then Exit Sub
Call frmMain.GotoLocation(lvMapLoc.SelectedItem, nMapStartMap, Me)
'If frmMain.WindowState = vbMinimized Then frmMain.WindowState = vbNormal
'frmMain.SetFocus

'If chkMapOptions(6).Value = 0 Then
'    If FormIsLoaded("frmResults") Then
'        If frmResults.objFormOwner Is Me Then
'            lR = SetTopMostWindow(frmResults.hwnd, True)
'        End If
'    End If
'End If
DoEvents
out:
Exit Sub
Error:
Call HandleError("lvMapLoc_DblClick")
Resume out:
End Sub

Private Sub txtRoomMap_GotFocus()
Call SelectAll(txtRoomMap)
End Sub

Private Sub txtRoomRoom_GotFocus()
Call SelectAll(txtRoomRoom)
End Sub

Public Sub LoadPresets(Optional ByVal bReset As Boolean)
Dim x As Integer, sSectionName As String, nMap As Long, nRoom As Long, sName As String
Dim cReg As clsRegistryRoutines, nError As Integer, bResult As Boolean

On Error GoTo Error:

Set cReg = New clsRegistryRoutines

Me.MousePointer = vbHourglass

If InStr(1, frmMain.lblDatVer.Caption, "-") = 0 Then
    sSectionName = "Custom_Presets"
Else
    sSectionName = RemoveCharacter(frmMain.lblDatVer.Caption, " ") & "_Presets"
End If

nError = RegCreateKeyPath(HKEY_LOCAL_MACHINE, "Software\MMUD Explorer\Presets\" & sSectionName)
If nError > 0 Then GoTo Error:

cReg.hkey = HKEY_LOCAL_MACHINE
cReg.KeyRoot = "Software\MMUD Explorer\Presets"
cReg.Subkey = sSectionName

If bReset Then
    For x = 0 To 49
        bResult = cReg.SetRegistryValue("Map" & x, "0", REG_SZ)
        If bResult = False Then Err.Raise 0, "LoadPresets", "Error Setting Registry Values"
    Next
End If

For x = 0 To 49
    nMap = Val(cReg.GetRegistryValue("Map" & x, 0))
    nRoom = Val(cReg.GetRegistryValue("Room" & x, 0))
    sName = cReg.GetRegistryValue("Name" & x, 0)
    
    If nMap = 0 Or nRoom = 0 Or sName = "" Then
        Select Case x
            Case 0: nMap = 10: nRoom = 271: sName = "Aged Titan"
            Case 1: nMap = 3: nRoom = 560: sName = "Ancient Ruin"
            Case 2: nMap = 17: nRoom = 2269: sName = "Arlysia"
            Case 3: nMap = 7: nRoom = 1176: sName = "Black Fortress"
            Case 4: nMap = 3: nRoom = 669: sName = "Black Wastelands"
            Case 5: nMap = 17: nRoom = 241: sName = "Blackwood Graveyard"
            Case 6: nMap = 8: nRoom = 461: sName = "Dark-Elf Castle"
            Case 7: nMap = 6: nRoom = 552: sName = "Gnome Village"
            Case 8: nMap = 12: nRoom = 1919: sName = "Great Pyramid"
            Case 9: nMap = 6: nRoom = 1255: sName = "Khazarad"
            Case 10: nMap = 7: nRoom = 884: sName = "Lava Fields"
            Case 11: nMap = 16: nRoom = 454: sName = "Lost City"
            Case 12: nMap = 12: nRoom = 5: sName = "Nekojin Village"
            Case 13: nMap = 2: nRoom = 2523: sName = "Rhudar"
            Case 14: nMap = 12: nRoom = 2099: sName = "Saracen Fort"
            Case 15: nMap = 12: nRoom = 1173: sName = "Small Pyramid"
            Case 16: nMap = 16: nRoom = 1179: sName = "Storm Fortress"
            Case 17: nMap = 16: nRoom = 1: sName = "Tasloi Village"
            Case 18: nMap = 1: nRoom = 224: sName = "Town Square"
            Case 19: nMap = 16: nRoom = 1990: sName = "Volcano"
            Case Else: nMap = 1: nRoom = 1: sName = "unset"
        End Select
        
        Call cReg.SetRegistryValue("Map" & x, nMap, REG_SZ)
        Call cReg.SetRegistryValue("Room" & x, nRoom, REG_SZ)
        Call cReg.SetRegistryValue("Name" & x, sName, REG_SZ)
    End If
    
Next x

For x = 0 To 9
    cmdMapPreset(x).Caption = cReg.GetRegistryValue("Name" & x, "unset")
    cmdMapPreset(x).Tag = x
Next x

Me.MousePointer = vbDefault

Exit Sub

Error:
Call HandleError("LoadPresets")

End Sub

Private Sub cmdResetPresets_Click()
Dim nYesNo As Integer

nYesNo = MsgBox("Are you sure you want to reset the presets to the default set?", vbYesNo + vbDefaultButton2 + vbQuestion, "Reset Presets?")

If nYesNo = vbYes Then Call LoadPresets(True)

Call frmMain.LoadPresets

End Sub

Private Sub cmdMapPreset_Click(Index As Integer)
Dim nMap As Long, nRoom As Long, sSectionName As String
Dim cReg As clsRegistryRoutines
On Error GoTo Error:

Set cReg = New clsRegistryRoutines

If InStr(1, frmMain.lblDatVer.Caption, "-") = 0 Then
    sSectionName = "Custom_Presets"
Else
    sSectionName = RemoveCharacter(frmMain.lblDatVer.Caption, " ") & "_Presets"
End If

cReg.hkey = HKEY_LOCAL_MACHINE
cReg.KeyRoot = "Software\MMUD Explorer\Presets"
cReg.Subkey = sSectionName

nMap = cReg.GetRegistryValue("Map" & cmdMapPreset(Index).Tag, 0) 'Val(ReadINI(sSectionName, "Map" & cmdMapPreset(index).Tag))
nRoom = cReg.GetRegistryValue("Room" & cmdMapPreset(Index).Tag, 0) 'Val(ReadINI(sSectionName, "Room" & cmdMapPreset(index).Tag))

Call MapStartMapping(nMap, nRoom)

out:
Exit Sub
Error:
Call HandleError("cmdMapPreset_Click")
Resume out:

End Sub

Private Sub cmdEditPreset_Click(Index As Integer)
Dim sSectionName As String, lR As Long
Dim cReg As clsRegistryRoutines
On Error GoTo Error:
Set cReg = New clsRegistryRoutines

If InStr(1, frmMain.lblDatVer.Caption, "-") = 0 Then
    sSectionName = "Custom_Presets"
Else
    sSectionName = RemoveCharacter(frmMain.lblDatVer.Caption, " ") & "_Presets"
End If
'sSectionName = RemoveCharacter(lblDatVer.Caption, " ") & "_Presets"

cReg.hkey = HKEY_LOCAL_MACHINE
cReg.KeyRoot = "Software\MMUD Explorer\Presets"
cReg.Subkey = sSectionName

Unload frmEditPreset
Load frmEditPreset
frmEditPreset.nPreset = Val(cmdMapPreset(Index).Tag)
frmEditPreset.lblCaption.Caption = "Editing Preset #" & (cmdMapPreset(Index).Tag + 1)
frmEditPreset.txtMap.Text = cReg.GetRegistryValue("Map" & cmdMapPreset(Index).Tag, 0) 'ReadINI(sSectionName, "Map" & cmdMapPreset(index).Tag)
frmEditPreset.txtRoom.Text = cReg.GetRegistryValue("Room" & cmdMapPreset(Index).Tag, 0) 'ReadINI(sSectionName, "Room" & cmdMapPreset(index).Tag)
frmEditPreset.txtCaption.Text = cReg.GetRegistryValue("Name" & cmdMapPreset(Index).Tag, "unset") 'ReadINI(sSectionName, "Name" & cmdMapPreset(index).Tag)
Set frmEditPreset.objFormOwner = Me
If chkMapOptions(6).Value = 0 Then lR = SetTopMostWindow(Me.hwnd, False)
DoEvents
frmEditPreset.Show vbModal, Me
If chkMapOptions(6).Value = 0 Then lR = SetTopMostWindow(Me.hwnd, True)

If Not frmEditPreset.nPreset < 0 Then
    Call cReg.SetRegistryValue("Map" & cmdMapPreset(Index).Tag, frmEditPreset.txtMap.Text, REG_SZ)
    Call cReg.SetRegistryValue("Room" & cmdMapPreset(Index).Tag, frmEditPreset.txtRoom.Text, REG_SZ)
    Call cReg.SetRegistryValue("Name" & cmdMapPreset(Index).Tag, frmEditPreset.txtCaption.Text, REG_SZ)
    cmdMapPreset(Index).Caption = frmEditPreset.txtCaption.Text
End If

Unload frmEditPreset

Call frmMain.LoadPresets

Exit Sub
Error:
Call HandleError("cmdEditPreset_Click")
Unload frmEditPreset
End Sub
