VERSION 5.00
Begin VB.Form frmWake 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cattle Prod"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8190
   Icon            =   "frmWake.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdWake 
      Caption         =   "Wake"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6960
      TabIndex        =   13
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdWakeAll 
      Caption         =   "Wake All"
      Height          =   375
      Left            =   5760
      TabIndex        =   12
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "MAC List (Imported)"
      Height          =   3975
      Left            =   5640
      TabIndex        =   10
      Top             =   0
      Width           =   2415
      Begin VB.ListBox lstMAC 
         Height          =   3570
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enter a Single MAC Address or Open File"
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   5415
      Begin VB.TextBox txtMAC 
         Height          =   285
         Left            =   840
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Address Format is 6 two digit characters"
         Height          =   255
         Left            =   2400
         TabIndex        =   14
         Top             =   360
         Width           =   2895
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmWake.frx":0442
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Timer tmrLight 
      Interval        =   500
      Left            =   7800
      Top             =   4200
   End
   Begin VB.Frame fmeMultiple 
      Caption         =   "Import MAC File"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   5415
      Begin VB.CommandButton cmdImport 
         Caption         =   "Import MAC's"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3960
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtImport 
         Height          =   285
         Left            =   2040
         TabIndex        =   4
         Top             =   360
         Width           =   1815
      End
      Begin VB.FileListBox filImport 
         Height          =   2040
         Left            =   2040
         Pattern         =   "*.csv"
         TabIndex        =   3
         Top             =   840
         Width           =   3255
      End
      Begin VB.DirListBox dirImport 
         Height          =   2115
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.DriveListBox drvImport 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Image imgOff 
      Height          =   480
      Left            =   7200
      Picture         =   "frmWake.frx":2C74
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgOn 
      Height          =   480
      Left            =   6720
      Picture         =   "frmWake.frx":30B6
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   3840
      Width           =   4095
   End
   Begin VB.Label lblNoMAC 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmWake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdImport_Click()
    ' Import file may be any comma delimited file, but first column must be mac address
    lstMAC.Clear
    Dim varMAC As String
    Open txtImport For Input As #2
        Do Until EOF(2)
            Line Input #2, varData
            varPos = InStr(1, varData, ",")
            varPos2 = InStr(varPos + 1, varData, ",")
            varMAC = Mid(varData, varPos2 + 1, 25)
            If Len(varMAC) = 12 Then lstMAC.AddItem varMAC
        Loop
    Close 2
    lblNoMAC = lstMAC.ListCount & " PC's"
    lblNoMAC.Refresh
End Sub

Private Sub cmdWake_Click()
    Dim fMatch As Boolean, sRTT As String, sHost As String
        fMatch = True
        sHost = "255.255.255.255"
        If Ping(sHost, sRTT, fMatch) Then
            ' Not found
        Else
            ' Found it
        End If
End Sub

Private Sub cmdWakeAll_Click()
    For a = 0 To lstMAC.ListCount - 1
        txtMAC = lstMAC.List(a)
        txtMAC.Refresh
        cmdWake_Click
        
        Dt = Chr(255) & Chr(255) & Chr(255) & Chr(255) & Chr(255) & Chr(255)
        MAC1 = "&H" & Mid(frmWake.txtMAC, 1, 2)
        MAC1 = CDec(MAC1)
        MAC1 = Chr(MAC1)
        MAC2 = "&H" & Mid(frmWake.txtMAC, 3, 2)
        MAC2 = CDec(MAC2)
        MAC2 = Chr(MAC2)
        MAC3 = "&H" & Mid(frmWake.txtMAC, 5, 2)
        MAC3 = CDec(MAC3)
        MAC3 = Chr(MAC3)
        MAC4 = "&H" & Mid(frmWake.txtMAC, 7, 2)
        MAC4 = CDec(MAC4)
        MAC4 = Chr(MAC4)
        MAC5 = "&H" & Mid(frmWake.txtMAC, 9, 2)
        MAC5 = CDec(MAC5)
        MAC5 = Chr(MAC5)
        MAC6 = "&H" & Mid(frmWake.txtMAC, 11, 2)
        MAC6 = CDec(MAC6)
        MAC6 = Chr(MAC6)
        
        For lp = 1 To 16
            Dt = Dt & MAC1 & MAC2 & MAC3 & MAC4 & MAC5 & MAC6
        Next lp

        lblStatus = a + 1 & " of " & lstMAC.ListCount
        lblStatus.Refresh
    Next a
End Sub

Private Sub dirImport_Change()
    filImport.Path = dirImport.Path
End Sub

Private Sub drvImport_Change()
    dirImport.Path = drvImport.Drive
End Sub

Private Sub filImport_Click()
    If Right(filImport.Path, 1) = "\" Then
        txtImport = filImport.Path & filImport.List(filImport.ListIndex)
    Else
        txtImport = filImport.Path & "\" & filImport.List(filImport.ListIndex)
    End If
End Sub

Private Sub Form_Load()
    ' command line = "/f:c:\temp\macs.csv" , where /f: indicates the filename to import
    If Command = "" Then
        
    Else
        Me.Hide
        mnuMultiple_Click
        If InStr(1, Command, "/f:") > 0 Then
            lpos = InStr(1, Command, "/f:") + 3
            rpos = InStr(lpos, Command, " ")
            If rpos = 0 Then rpos = Len(Command) + 1
            'txtImport = "c:\temp\dhcp client info.csv"
            txtImport = Mid(Command, lpos, rpos - lpos)
            cmdImport_Click
            cmdWakeAll_Click
        End
        End If
    End If
End Sub

Private Sub lstMAC_Click()
    txtMAC = lstMAC.List(lstMAC.ListIndex)
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuMultiple_Click()
    frmWake.Height = 5370
    frmWake.Width = 8385
    mnuSingle.Checked = False
    mnuMultiple.Checked = True
    cmdWake.Visible = False
    cmdWakeAll.Visible = True
End Sub

Private Sub mnuSingle_Click()
    frmWake.Height = 1200
    frmWake.Width = 4155
    mnuSingle.Checked = True
    mnuMultiple.Checked = False
    cmdWake.Visible = True
    cmdWakeAll.Visible = False
End Sub

Private Sub tmrLight_Timer()
    Static varLight
    varLight = Not (varLight)
    If varLight = 0 Then frmWake.Icon = imgOff.Picture
    If varLight = -1 Then frmWake.Icon = imgOn.Picture
    tmrLight.Enabled = True
End Sub

Private Sub txtImport_Change()
    If txtImport <> "" Then
        cmdImport.Enabled = True
    Else
        cmdImport.Enabled = False
    End If
End Sub

Private Sub txtMAC_Change()
    If Len(txtMAC) <> 12 Then
        cmdWake.Enabled = False
    Else
        cmdWake.Enabled = True
    End If
End Sub
