VERSION 5.00
Begin VB.Form CustomStartMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customize Start Menu"
   ClientHeight    =   4500
   ClientLeft      =   3300
   ClientTop       =   2790
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   6090
   Begin VB.CommandButton Command4 
      Caption         =   "S&hut Down"
      Height          =   375
      Left            =   4380
      TabIndex        =   44
      Top             =   2910
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Update"
      Height          =   375
      Left            =   3000
      TabIndex        =   41
      Top             =   3420
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "&Start Menu "
      Height          =   2595
      Left            =   3030
      TabIndex        =   37
      Top             =   120
      Width           =   2865
      Begin VB.CheckBox Check11 
         Caption         =   "Disable Recent Docs History"
         Height          =   225
         Left            =   210
         TabIndex        =   43
         Top             =   2190
         Width           =   2355
      End
      Begin VB.CheckBox Check10 
         Caption         =   "Hide Control Panel/Printers"
         Height          =   225
         Left            =   210
         TabIndex        =   40
         Top             =   1710
         Width           =   2235
      End
      Begin VB.CheckBox Check9 
         Caption         =   "Hide &Find Menu"
         Height          =   225
         Left            =   210
         TabIndex        =   39
         Top             =   1470
         Width           =   1755
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Hide &Logoff Menu"
         Height          =   225
         Left            =   210
         TabIndex        =   38
         Top             =   1230
         Width           =   1695
      End
      Begin VB.CheckBox Check7 
         Caption         =   "&Clear Recent Docs On Exit"
         Height          =   225
         Left            =   210
         TabIndex        =   36
         Top             =   1950
         Width           =   2325
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Hide &Run Menu"
         Height          =   225
         Left            =   210
         TabIndex        =   32
         Top             =   270
         Width           =   1935
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Hide Sh&ut Down Menu"
         Height          =   225
         Left            =   210
         TabIndex        =   33
         Top             =   510
         Width           =   2115
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Hide Recent &Docs Menu"
         Height          =   225
         Left            =   210
         TabIndex        =   35
         Top             =   990
         Width           =   2115
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Hide F&avorites Menu"
         Height          =   225
         Left            =   210
         TabIndex        =   34
         Top             =   750
         Width           =   1965
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "&Name On Systray "
      Height          =   765
      Left            =   180
      TabIndex        =   30
      Top             =   3030
      Width           =   2265
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   210
         MaxLength       =   8
         TabIndex        =   31
         Top             =   300
         Width           =   1785
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&CANCEL"
      Height          =   375
      Left            =   4380
      TabIndex        =   1
      Top             =   3420
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   2910
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Hide Drive(s) in EXPLORER: "
      Height          =   2640
      Left            =   180
      TabIndex        =   2
      Top             =   120
      Width           =   2595
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Caption         =   "All"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1815
         TabIndex        =   29
         Top             =   2160
         Width           =   480
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Z"
         Enabled         =   0   'False
         Height          =   255
         Index           =   25
         Left            =   1830
         TabIndex        =   28
         Top             =   1920
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Y"
         Enabled         =   0   'False
         Height          =   255
         Index           =   24
         Left            =   1830
         TabIndex        =   27
         Top             =   1680
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "X"
         Enabled         =   0   'False
         Height          =   255
         Index           =   23
         Left            =   1830
         TabIndex        =   26
         Top             =   1440
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "W"
         Enabled         =   0   'False
         Height          =   255
         Index           =   22
         Left            =   1830
         TabIndex        =   25
         Top             =   1200
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "V"
         Enabled         =   0   'False
         Height          =   255
         Index           =   21
         Left            =   1830
         TabIndex        =   24
         Top             =   960
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "U"
         Enabled         =   0   'False
         Height          =   255
         Index           =   20
         Left            =   1830
         TabIndex        =   23
         Top             =   720
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "T"
         Enabled         =   0   'False
         Height          =   255
         Index           =   19
         Left            =   1830
         TabIndex        =   22
         Top             =   480
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "S"
         Enabled         =   0   'False
         Height          =   255
         Index           =   18
         Left            =   1830
         TabIndex        =   21
         Top             =   240
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "R"
         Enabled         =   0   'False
         Height          =   255
         Index           =   17
         Left            =   990
         TabIndex        =   20
         Top             =   2160
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Q"
         Enabled         =   0   'False
         Height          =   255
         Index           =   16
         Left            =   990
         TabIndex        =   19
         Top             =   1920
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "P"
         Enabled         =   0   'False
         Height          =   255
         Index           =   15
         Left            =   990
         TabIndex        =   18
         Top             =   1680
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "O"
         Enabled         =   0   'False
         Height          =   255
         Index           =   14
         Left            =   990
         TabIndex        =   17
         Top             =   1440
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "N"
         Enabled         =   0   'False
         Height          =   255
         Index           =   13
         Left            =   990
         TabIndex        =   16
         Top             =   1200
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "M"
         Enabled         =   0   'False
         Height          =   255
         Index           =   12
         Left            =   990
         TabIndex        =   15
         Top             =   960
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "L"
         Enabled         =   0   'False
         Height          =   255
         Index           =   11
         Left            =   990
         TabIndex        =   14
         Top             =   720
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "K"
         Enabled         =   0   'False
         Height          =   255
         Index           =   10
         Left            =   990
         TabIndex        =   13
         Top             =   480
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "J"
         Enabled         =   0   'False
         Height          =   255
         Index           =   9
         Left            =   990
         TabIndex        =   12
         Top             =   240
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "I"
         Enabled         =   0   'False
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   11
         Top             =   2130
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "H"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   10
         Top             =   1890
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "G"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   9
         Top             =   1650
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "F"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "E"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "D"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "C"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "B"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "A"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   465
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "linda.69@mailcity.com"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   150
      TabIndex        =   42
      Top             =   4020
      Width           =   1575
   End
End
Attribute VB_Name = "CustomStartMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' linda.69@mailcity.com
'

Option Explicit
DefInt A-Z

Private Declare Function ExitWindowsEx% Lib "user32" (ByVal l As Long, ByVal i As Integer)

Dim strSysTray01 As String
Dim strSysTray02 As String
Dim strTemp As String
Dim iDrives As Byte

Dim bCRDOE As Byte         'Clear Recent Docs On Exit
Dim bC As Byte             'Shutdown Menu
Dim bFM As Byte            'Favorites Menu
Dim bFind As Byte          'Find Menu
Dim bLOM As Byte           'Logoff Menu
Dim bRDM As Byte           'Recent Documents Menu
Dim bRM As Byte            'Run Menu
Dim bSFM As Byte           'Control Panel/Printers Menu
Dim bRDH As Byte           'Disable Recent Docs History

Sub GetSettings()
  strTemp = GetBinaryValue("HKEY_CURRENT_USER\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\POLICIES\EXPLORER", _
        "ClearRecentDocsOnExit")
  If strTemp <> "Error" Then
    strTemp = Mid(strTemp, 1, 1)      'we only need the first
    bCRDOE = Asc(strTemp)
  Else
    bCRDOE = 255
  End If
  
  strTemp = GetBinaryValue("HKEY_CURRENT_USER\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\POLICIES\EXPLORER", _
        "NoClose")
  If strTemp <> "Error" Then
    strTemp = Mid(strTemp, 1, 1)
    bC = Asc(strTemp)
  Else
    bC = 255
  End If
  
  strTemp = GetBinaryValue("HKEY_CURRENT_USER\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\POLICIES\EXPLORER", _
        "NoFavoritesMenu")
  If strTemp <> "Error" Then
    strTemp = Mid(strTemp, 1, 1)
    bFM = Asc(strTemp)
  Else
    bFM = 255
  End If
  
  strTemp = GetBinaryValue("HKEY_CURRENT_USER\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\POLICIES\EXPLORER", _
        "NoFind")
  If strTemp <> "Error" Then
    strTemp = Mid(strTemp, 1, 1)
    bFind = Asc(strTemp)
  Else
    bFind = 255
  End If
  
  strTemp = GetBinaryValue("HKEY_CURRENT_USER\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\POLICIES\EXPLORER", _
        "NoLogOff")
  If strTemp <> "Error" Then
    strTemp = Mid(strTemp, 1, 1)
    bLOM = Asc(strTemp)
  Else
    bLOM = 255
  End If
  
  strTemp = GetBinaryValue("HKEY_CURRENT_USER\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\POLICIES\EXPLORER", _
        "NoRecentDocsMenu")
  If strTemp <> "Error" Then
    strTemp = Mid(strTemp, 1, 1)
    bRDM = Asc(strTemp)
  Else
    bRDM = 255
  End If
  
  strTemp = GetBinaryValue("HKEY_CURRENT_USER\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\POLICIES\EXPLORER", _
        "NoRun")
  If strTemp <> "Error" Then
    strTemp = Mid(strTemp, 1, 1)
    bRM = Asc(strTemp)
  Else
    bRM = 255
  End If
  
  strTemp = GetBinaryValue("HKEY_CURRENT_USER\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\POLICIES\EXPLORER", _
        "NoSetFolders")
  If strTemp <> "Error" Then
    strTemp = Mid(strTemp, 1, 1)
    bSFM = Asc(strTemp)
  Else
    bSFM = 255
  End If
  
  strTemp = GetBinaryValue("HKEY_CURRENT_USER\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\POLICIES\EXPLORER", _
        "NoRecentDocsHistory")
  If strTemp <> "Error" Then
    strTemp = Mid(strTemp, 1, 1)
    bRDH = Asc(strTemp)
  Else
    bRDH = 255
  End If
  
  strSysTray01 = GetStringValue("HKEY_USERS\.DEFAULT\CONTROL PANEL\INTERNATIONAL", _
               "s1159")
  strSysTray02 = GetStringValue("HKEY_USERS\.DEFAULT\CONTROL PANEL\INTERNATIONAL", _
               "s2359")
  
  strTemp = GetBinaryValue("HKEY_CURRENT_USER\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\POLICIES\EXPLORER", _
            "NoDrives")
  If strTemp <> "Error" Then
    strTemp = Mid(strTemp, 1, 1)
    iDrives = Asc(strTemp)
  Else
    iDrives = 0
  End If
  
  Do
    Select Case iDrives
      Case Is >= &H80
        Check1(7).Value = 1
        iDrives = iDrives - 128
      Case Is >= &H40
        Check1(6).Value = 1
        iDrives = iDrives - 64
      Case Is >= &H20
        Check1(5).Value = 1
        iDrives = iDrives - 32
      Case Is >= &H10
        Check1(4).Value = 1
        iDrives = iDrives - 16
      Case Is >= &H8
        Check1(3).Value = 1
        iDrives = iDrives - 8
      Case Is >= &H4
        Check1(2).Value = 1
        iDrives = iDrives - 4
      Case Is >= &H2
        Check1(1).Value = 1
        iDrives = iDrives - 2
      Case Is >= &H1
        Check1(0).Value = 1
        iDrives = iDrives - 1
    End Select
  Loop Until iDrives = 0
  
  Text1.Text = IIf(strSysTray01 = "Error", "", strSysTray01)
                  
  Check3.Value = IIf((bRM = 255 Or bRM = 0), 0, 1)     '255 error
  Check4.Value = IIf((bC = 255 Or bC = 0), 0, 1)       '0   false
  Check5.Value = IIf((bFM = 255 Or bFM = 0), 0, 1)     '1   true
  Check6.Value = IIf((bRDM = 255 Or bRDM = 0), 0, 1)
  Check8.Value = IIf((bLOM = 255 Or bLOM = 0), 0, 1)
  Check9.Value = IIf((bFind = 255 Or bFind = 0), 0, 1)
  Check10.Value = IIf((bSFM = 255 Or bSFM = 0), 0, 1)
  Check11.Value = IIf((bRDH = 255 Or bRDH = 0), 0, 1)
  Check7.Value = IIf((bCRDOE = 255 Or bCRDOE = 0), 0, 1)
End Sub

Sub SaveSettings()

  If Text1.Text <> "" Then strSysTray01 = Text1.Text
                   
  bRM = Check3.Value: bC = Check4.Value: bFM = Check5.Value
  bRDM = Check6.Value: bLOM = Check8.Value: bFind = Check9.Value
  bSFM = Check10.Value: bRDH = Check11.Value: bCRDOE = Check7.Value
  
  strTemp = Chr(bRM) + Chr(0) + Chr(0) + Chr(0)  'value needs a DWORD
  SetBinaryValue "HKEY_CURRENT_USER\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\POLICIES\EXPLORER", _
        "NoRun", strTemp
        
  strTemp = Chr(bC) + Chr(0) + Chr(0) + Chr(0)
  SetBinaryValue "HKEY_CURRENT_USER\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\POLICIES\EXPLORER", _
        "NoClose", strTemp
  
  strTemp = Chr(bFM) + Chr(0) + Chr(0) + Chr(0)
  SetBinaryValue "HKEY_CURRENT_USER\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\POLICIES\EXPLORER", _
        "NoFavoritesMenu", strTemp
  
  strTemp = Chr(bRDM) + Chr(0) + Chr(0) + Chr(0)
  SetBinaryValue "HKEY_CURRENT_USER\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\POLICIES\EXPLORER", _
        "NoRecentDocsMenu", strTemp
  
  strTemp = Chr(bLOM) + Chr(0) + Chr(0) + Chr(0)
  SetBinaryValue "HKEY_CURRENT_USER\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\POLICIES\EXPLORER", _
        "NoLogOff", strTemp
  
  strTemp = Chr(bFind) + Chr(0) + Chr(0) + Chr(0)
  SetBinaryValue "HKEY_CURRENT_USER\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\POLICIES\EXPLORER", _
        "NoFind", strTemp
  
  strTemp = Chr(bSFM) + Chr(0) + Chr(0) + Chr(0)
  SetBinaryValue "HKEY_CURRENT_USER\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\POLICIES\EXPLORER", _
        "NoSetFolders", strTemp
   
  strTemp = Chr(bCRDOE) + Chr(0) + Chr(0) + Chr(0)
  SetBinaryValue "HKEY_CURRENT_USER\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\POLICIES\EXPLORER", _
        "ClearRecentDocsOnExit", strTemp
  
  strTemp = Chr(bRDH) + Chr(0) + Chr(0) + Chr(0)
  SetBinaryValue "HKEY_CURRENT_USER\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\POLICIES\EXPLORER", _
        "NoRecentDocsHistory", strTemp
  
  SetStringValue "HKEY_USERS\.DEFAULT\CONTROL PANEL\INTERNATIONAL", _
               "s1159", strSysTray01
  SetStringValue "HKEY_USERS\.DEFAULT\CONTROL PANEL\INTERNATIONAL", _
               "s2359", strSysTray01

  iDrives = 0
 
  If Check1(7).Value = 1 Then iDrives = &H80
  If Check1(6).Value = 1 Then iDrives = iDrives + &H40
  If Check1(5).Value = 1 Then iDrives = iDrives + &H20
  If Check1(4).Value = 1 Then iDrives = iDrives + &H10
  If Check1(3).Value = 1 Then iDrives = iDrives + &H8
  If Check1(2).Value = 1 Then iDrives = iDrives + &H4
  If Check1(1).Value = 1 Then iDrives = iDrives + &H2
  If Check1(0).Value = 1 Then iDrives = iDrives + &H1

  strTemp = Chr(iDrives) + Chr(0) + Chr(0) + Chr(0)
  SetBinaryValue "HKEY_CURRENT_USER\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\POLICIES\EXPLORER", _
            "NoDrives", Chr(iDrives)
End Sub

Private Sub Form_Load()
  GetSettings
End Sub

Private Sub Command1_Click()
  SaveSettings
  MsgBox "You Need to Shutdown computer for settings to have an effect.", vbInformation, "ShutDown"
  Unload Me
End Sub

Private Sub Command2_Click()
  End
End Sub

Private Sub Command3_Click()
  SaveSettings
  GetSettings
End Sub

Private Sub Command4_Click()
Dim i As Integer
  SaveSettings
  ExitWindowsEx 2, i
End Sub

