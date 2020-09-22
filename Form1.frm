VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "C-TecK Email Finder"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   12015
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   7800
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   10920
      TabIndex        =   4
      Top             =   0
      Width           =   1095
   End
   Begin InetCtlsObjects.Inet web 
      Left            =   8640
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Default         =   -1  'True
      Height          =   375
      Left            =   9840
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "Monster.com"
      Top             =   120
      Width           =   9855
   End
   Begin MSComctlLib.ListView lvDetails 
      Height          =   7095
      Left            =   5160
      TabIndex        =   2
      Top             =   360
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   12515
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvURLs 
      Height          =   7095
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   12515
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      Height          =   7455
      Left            =   11880
      Top             =   360
      Width           =   135
   End
   Begin VB.Label lblEmails 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   8400
      TabIndex        =   10
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Emails Found :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   7080
      TabIndex        =   9
      Top             =   7560
      Width           =   1335
   End
   Begin VB.Label lblScanned 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   5280
      TabIndex        =   8
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Pages Scanned :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Label lblTotalURLs 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pages In Cache :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      Height          =   7095
      Left            =   0
      Top             =   360
      Width           =   135
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   0
      Top             =   7440
      Width           =   12015
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   0
      Top             =   240
      Width           =   10095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      Height          =   7095
      Left            =   5040
      Top             =   360
      Width           =   135
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu set 
         Caption         =   "Settings"
      End
      Begin VB.Menu clear 
         Caption         =   "Clear"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu about 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Times1 As Long
Public PageScan As Long
Public MaxURL As Long
Public End1 As Boolean
Public FindAny As Boolean
Public URLTime As Long
Dim AllEmail() As String
Dim AllURLs() As String

Private Declare Function GetTempFileName Lib "kernel32" _
    Alias "GetTempFileNameA" (ByVal lpszPath As String, _
    ByVal lpPrefixString As String, ByVal wUnique As Long, _
    ByVal lpTempFileName As String) As Long


Private Sub clear_Click()
lvDetails.ListItems.clear
lvURLs.ListItems.clear
Times1 = 1
End Sub

Private Sub Command1_Click()
On Error Resume Next
    
    Dim WebSite As String
    Dim CurrentSite As String
    Dim StartPoint As Long
    Dim EndPos As Long
    Dim TheLink As String
    Dim LinkLength As Long
    Dim StartPos As Long
    Dim MainLoop1 As Long
    Dim Times3 As Long
    Dim TheSiteCon As String
    Dim URLsFound As Long
    Dim EmailAddress As String
    Dim lineofsite As String
    Dim I As Long
    Dim H65 As Long
    Dim midpos As Long
    Dim temp9 As Long
    Dim Temp10 As Long
    Dim sFile As String
    
    If Text1.Text = "" Then 'Check to see if the user put if a Web Site Address
        MsgBox "Please put in a Web Address to be spidered."
        Exit Sub
    End If
    
    ' cleare the lists first.
    lvURLs.ListItems.clear
    lvDetails.ListItems.clear
    
    ReDim AllEmail(1000000) As String
    ReDim AllURLs(1000000) As String

    URLsFound = 0
    CurrentSite = Text1.Text
    End1 = False
    StartPoint = 1
    Command1.Enabled = False
    Command2.Enabled = True
    AllURLs(URLTime) = Text1.Text
    
    ' make the temp dir if it isnt there
    DirExists App.Path & "\temp"
    
Start:
    CurrentSite = AllURLs(URLTime)
    If CurrentSite = "" Then
        MsgBox "There are no more link to find"
        GoTo DoEmail
    End If
    
    'Get the WebSite Sourcecode
    WebSite = web.OpenURL(CurrentSite)
    
    ' generate a temp file name, with full path
    sFile = GenTempName(App.Path & "\temp", 0)
    Open sFile For Output As #1 'writes to a file
    Print #1, WebSite
    Close #1
    
    Open sFile For Input As #1 'Opens the temp file
    
    Do While EOF(1) = False                     'Starts the Main Loop
        Line Input #1, lineofsite
        StartPos = InStr(StartPoint, lineofsite, "<a href=", vbTextCompare)
        If StartPos > 0 Then     'Looks in it for a link.
            ' find the > at the end of the link
            EndPos = InStr(StartPos, lineofsite, Chr(62), vbTextCompare) 'Get the end part of the link.
            ' if there is no > then go to the next line
            If EndPos = 0 Then GoTo Restart_Loop:
            LinkLength = EndPos - StartPos
            ' get just the link
            TheLink = Trim$(Mid$(lineofsite, StartPos + 8, LinkLength - 8))
            If InStr(1, TheLink, " ") <> 0 Then
                TheLink = Trim$(Left$(TheLink, InStr(1, TheLink, " ")))
            End If
            If Left$(TheLink, 1) = Chr(34) Then
                TheLink = Mid$(TheLink, 2)
            End If
            If Right$(TheLink, 1) = Chr(34) Then
                TheLink = Left$(TheLink, Len(TheLink) - 1)
            End If
            ' for now we only get ones that start with http
            If Not Left$(TheLink, 4) = "http" Then GoTo Restart_Loop:
            If Left$(TheLink, 1) = "/" Then GoTo Restart_Loop
            
            ' skip a lot of formats that we dont want to play with
            If Right$(TheLink, 4) = ".gif" Then GoTo Restart_Loop:
            If Right$(TheLink, 4) = ".jpg" Then GoTo Restart_Loop:
            If Right$(TheLink, 4) = ".png" Then GoTo Restart_Loop:
            If Right$(TheLink, 4) = ".mpg" Then GoTo Restart_Loop:
            If Right$(TheLink, 4) = ".avi" Then GoTo Restart_Loop:
            If Right$(TheLink, 4) = ".swf" Then GoTo Restart_Loop:
            If Right$(TheLink, 4) = ".asf" Then GoTo Restart_Loop:
            If Right$(TheLink, 4) = ".rm" Then GoTo Restart_Loop:
            If Right$(TheLink, 4) = ".ra" Then GoTo Restart_Loop:
            If Right$(TheLink, 4) = ".wav" Then GoTo Restart_Loop:
            If Right$(TheLink, 4) = ".mp3" Then GoTo Restart_Loop:
                        
            I = 1
            Do Until Times1 = I
                If TheLink = AllURLs(I) Then
                    GoTo Restart_Loop:
                End If
                I = I + 1
                If End1 = True Then Exit Sub
            Loop
            
            If AddURL(Times1, TheLink) Then
                AllURLs(Times1) = TheLink
                ProgressBar1.Value = Times1
                Times1 = Times1 + 1
                lblTotalURLs.Caption = Times1
            End If
            If End1 = True Then Exit Sub
            If Times1 >= MaxURL Then GoTo DoEmail
        End If
Restart_Loop:
    Loop

    URLTime = URLTime + 1
    URLsFound = URLsFound + 1
    MainLoop1 = MainLoop1 + 1
    Close #1
    GoTo Start


DoEmail:
    ProgressBar1.Value = 0
    H65 = 1
    Do Until Times3 = MaxURL
        TheSiteCon = web.OpenURL(AllURLs(Times3))
        PageScan = PageScan + 1
        lblScanned.Caption = PageScan
        midpos = InStr(1, TheSiteCon, "@", vbTextCompare)
        temp9 = midpos
        If midpos > 0 Then
            Do Until Mid(TheSiteCon, midpos, 1) = ")" Or Mid(TheSiteCon, midpos, 1) = "(" Or Mid(TheSiteCon, midpos, 1) = Chr(13) Or Mid(TheSiteCon, midpos, 1) = "'" Or Mid(TheSiteCon, midpos, 1) = Chr(34) Or Mid(TheSiteCon, midpos, 1) = " " Or Mid(TheSiteCon, midpos, 1) = ">" Or Mid(TheSiteCon, midpos, 1) = "<" Or Mid(TheSiteCon, midpos, 1) = ":"
                midpos = midpos - 1
            Loop
            midpos = midpos + 1
            Do Until Mid(TheSiteCon, midpos, 1) = ")" Or Mid(TheSiteCon, midpos, 1) = "(" Or Mid(TheSiteCon, temp9, 1) = Chr(13) Or Mid(TheSiteCon, temp9, 1) = "'" Or Mid(TheSiteCon, temp9, 1) = Chr(34) Or Mid(TheSiteCon, temp9, 1) = " " Or Mid(TheSiteCon, temp9, 1) = ">" Or Mid(TheSiteCon, temp9, 1) = "<"
                temp9 = temp9 + 1
            Loop
            EndPos = temp9
            LinkLength = EndPos - midpos
            EmailAddress = Mid(TheSiteCon, midpos, LinkLength)
            EmailAddress = LCase(EmailAddress)
            If EmailAddress = "@" Then GoTo Restart_Email_Loop:
            Temp10 = 1
            If InStr(1, EmailAddress, "sale") Or InStr(1, EmailAddress, "info") Or InStr(1, EmailAddress, "yourname") Or InStr(1, EmailAddress, "webmaster") Or InStr(1, EmailAddress, "help") Then GoTo Restart_Email_Loop
            If Left(EmailAddress, 1) = "@" Then GoTo Restart_Email_Loop
            Do Until H65 = Temp10
                If AllEmail(Temp10) = EmailAddress Then GoTo Restart_Email_Loop
                Temp10 = Temp10 + 1
            Loop
            EmailAddress = CleanEmail(EmailAddress)
            If EmailAddress <> "" Then
                AllEmail(H65) = EmailAddress
                lvDetails.ListItems.Add , , EmailAddress
                lvDetails.ListItems(H65).ListSubItems.Add , , AllURLs(Times3)
                lblEmails.Caption = H65
                H65 = H65 + 1
            End If
        End If
Restart_Email_Loop:
        Times3 = Times3 + 1
        ProgressBar1.Value = Times3
        If End1 = True Then Exit Sub
    Loop

Command2_Click

End Sub

Private Sub Command2_Click()
    End1 = True
    Command2.Enabled = False
    Command1.Enabled = True
End Sub

Private Sub Form_Load()
    PageScan = 0
    MaxURL = 500
    ProgressBar1.Max = MaxURL
    Times1 = 1
    lvDetails.FullRowSelect = True
    lvDetails.View = lvwReport
    lvDetails.GridLines = True
    lvDetails.ColumnHeaders.clear
    lvDetails.ListItems.clear
    lvDetails.ColumnHeaders.Add , , "Email Address", 4000
    lvDetails.ColumnHeaders.Add , , "URL", 10000
    
    lvURLs.ColumnHeaders.Add , , "Number", 1000
    lvURLs.ColumnHeaders.Add , , "URLs", 4000
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ClearDirectory App.Path & "\temp"
End Sub

Private Sub set_Click()
    Form2.Show
    Unload Form2
End Sub

Public Function AddURL(ByVal Link As Long, ByVal URL As String) As String
        
    Dim oItem As ListItem
    
    Set oItem = lvURLs.ListItems.Add(, , Link)
    oItem.SubItems(1) = URL
    Set oItem = Nothing
End Function

Public Function CleanEmail(ByVal sEmail As String) As String
    On Error GoTo error_handler:
    
    If Len(sEmail) > 5 Then
        sEmail = Replace(sEmail, ")", "")
        sEmail = Replace(sEmail, "(", "")
        If IsBad(sEmail) Then
            CleanEmail = ""
        Else
            CleanEmail = sEmail
        End If
    Else
        CleanEmail = ""
    End If

    Exit Function
error_handler:
    CleanEmail = ""
    Err.clear
End Function

Public Function IsBad(ByVal sEmail As String) As Boolean
    On Error GoTo error_handler:

    Dim bToggle As Boolean
    Dim iLoop As Integer
    Dim sChar As String
    
    bToggle = False
    
    ' if it has more than 1 @ symbol its bad
    If CountInStr(sEmail, "@", True) > 1 Then
        bToggle = True
    End If
    
    ' now we check for the right characters
    For iLoop = 1 To Len(sEmail)
        sChar = Mid$(sEmail, iLoop, 1)
        If Asc(sChar) = 46 Or Between(Asc(sChar), 48, 57) Or Between(Asc(sChar), 64, 90) Or Between(Asc(sChar), 97, 122) Then
            ' its a good character
        Else
            ' its bad
            bToggle = True
            Exit For
        End If
    Next iLoop
    
    IsBad = bToggle
    
    Exit Function
error_handler:
    IsBad = True
    Err.clear
End Function

Public Function CountInStr(ByVal strString As String, ByVal strFind As String, _
    Optional ByVal boolIgnoreCase As Boolean) As Integer
    ' Return the number of times a string st
    '     rFind exists
    ' within another string strString
    Dim I As Integer, intTemp As Integer
    
    If boolIgnoreCase Then
        ' Set both parameter strings to lower ca
        '     se
        strString = LCase(strString)
        strFind = LCase(strFind)
    End If

    Do
        ' Loop through the string while still ma
        '     tching
        I = InStr(I + 1, strString, strFind)
        If I <> 0 Then intTemp = intTemp + 1
    Loop While I <> 0
    CountInStr = intTemp
End Function

Public Function Between(ByVal Number As Long, Min As Long, Max As Long) As Boolean
    If Number >= Min And Number <= Max Then Between = True Else Between = False
End Function

Public Function DirExists(ByVal sDir As String) As Boolean

    On Error GoTo Err_Handler
    Dim strDir As String

    strDir = Dir(sDir, vbDirectory)

    If (strDir = "") Then
         'If it doesn't exist, create it
        CreateDirectoryStruct sDir
    End If
    
    DirExists = True
    Exit Function

Err_Handler:
    DirExists = False
End Function

Public Function FileExists(ByVal sFile As String) As Boolean
  Dim lLength As Long

  If sFile <> vbNullString Then
    On Error Resume Next
    lLength = Len(Dir$(sFile))
    On Error GoTo err_routine
    FileExists = (Not Err And lLength > 0)
  Else
    FileExists = False
  End If

exit_routine:
  Exit Function

err_routine:
  FileExists = False
  Resume exit_routine

End Function

Public Sub CreateDirectoryStruct(ByVal CreateThisPath As String)

    On Error GoTo Err_Handler
    'do initial check
    Dim RET As Boolean
    Dim Temp As String
    Dim ComputerName As String
    Dim IntoItCount As Integer
    Dim x As Integer
    Dim WakeString As String
    Dim MadeIt As Integer

    If Dir$(CreateThisPath, vbDirectory) <> "" Then Exit Sub
    'is this a network path?

    If Left$(CreateThisPath, 2) = "\\" Then ' this is a UNC NetworkPath
        'must extract the machine name first, th
        '     en get to the first folder
        IntoItCount = 3
        ComputerName = Mid$(CreateThisPath, IntoItCount, InStr(IntoItCount, CreateThisPath, "\") - IntoItCount)
        IntoItCount = IntoItCount + Len(ComputerName) + 1
        IntoItCount = InStr(IntoItCount, CreateThisPath, "\") + 1
        'temp = Mid$(CreateThisPath, IntoItCount
        '     , x)
    Else ' this is a regular path
        IntoItCount = 4
    End If
    WakeString = Left$(CreateThisPath, IntoItCount - 1)
    'start a loop through the CreateThisPath
    '     string

    Do
        x = InStr(IntoItCount, CreateThisPath, "\")

        If x <> 0 Then
            x = x - IntoItCount
            Temp = Mid$(CreateThisPath, IntoItCount, x)
        Else
            Temp = Mid$(CreateThisPath, IntoItCount)
        End If
        IntoItCount = IntoItCount + Len(Temp) + 1
        Temp = WakeString + Temp
        'Create a directory if it doesn't alread
        '     y exist
        RET = (Dir$(Temp, vbDirectory) <> "")


        If Not RET Then
            'ret& = CreateDirectory(temp, Security)
            MkDir Temp
        End If
        IntoItCount = IntoItCount 'track where we are in the String
        WakeString = Left$(CreateThisPath, IntoItCount - 1)
    Loop While WakeString <> CreateThisPath

    Exit Sub

Err_Handler:
    Err.Raise Err.Number
End Sub

Private Function GenTempName(ByVal sPath As String, ByVal lUnique As Long) As String
    Dim sPrefix As String
    Dim sTempFileName As String
    If sPath = "" Then
        sPath = TempDir()
    End If
    sPrefix = "fVB"
    sTempFileName = Space$(100)
    GetTempFileName sPath, sPrefix, lUnique, sTempFileName
    sTempFileName = Mid$(sTempFileName, 1, InStr(sTempFileName, Chr$(0)) - 1)
    GenTempName = sTempFileName
End Function

Function TempDir() As String
    Dim sTemp As String
    sTemp = Environ$("temp")
    If Right(sTemp, 1) <> "\" Then
        sTemp = sTemp & "\"
    End If
    TempDir = sTemp
End Function

Private Sub ClearDirectory(ByVal psDirName As String)
    'This function attempts to delete all fi
    '     les
    'and subdirectories of the given
    'directory name, and leaves the given
    'directory intact, but completely empty.
    '
    '
    'If the Kill command generates an error
    '     (i.e.
    'file is in use by another process -
    'permission denied error), then that fil
    '     e and
    'subdirectory will be skipped, and the
    'program will continue (On Error Resume
'     Next).
'
'EXAMPLE CALL:
' ClearDirectory "C:\Temp\"
Dim sSubDir


If Len(psDirName) > 0 Then


    If Right(psDirName, 1) <> "\" Then
        psDirName = psDirName & "\"
    End If
    'Attempt to remove any files in director
    '     y
    'with one command (if error, we'll
    'attempt to delete the files one at a
    'time later in the loop):
    On Error Resume Next
    Kill psDirName & "*.*"


    DoEvents
        
        sSubDir = Dir(psDirName, vbDirectory)


        Do While Len(sSubDir) > 0
            'Ignore the current directory and the
            'encompassing directory:
            If sSubDir <> "." And _
            sSubDir <> ".." Then
            'Use bitwise comparison to make
            'sure MyName is a directory:
            If (GetAttr(psDirName & sSubDir) And _
            vbDirectory) = vbDirectory Then
            'Use recursion to clear files
            'from subdir:
            ClearDirectory psDirName & _
            sSubDir & "\"
            'Remove directory once files
            'have been cleared (deleted)
            'from it:
            RmDir psDirName & sSubDir


            DoEvents
                'ReInitialize Dir Command
                'after using recursion:
                sSubDir = Dir(psDirName, vbDirectory)
            Else
                'This file is remaining because
                'most likely, the Kill statement
                'before this loop errored out
                'when attempting to delete all
                'the files at once in this
                'directory. This attempt to
                'delete a single file by itself
                'may work because another
                '(locked) file within this same
                'directory may have prevented
                '(non-locked) files from being
                'deleted:
                Kill psDirName & sSubDir
                sSubDir = Dir
            End If
        Else
            sSubDir = Dir
        End If
    Loop
End If
End Sub
