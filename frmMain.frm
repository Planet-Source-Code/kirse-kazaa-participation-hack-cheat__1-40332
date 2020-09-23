VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KaZaA Participation Hacker"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrHackIt 
      Enabled         =   0   'False
      Interval        =   750
      Left            =   4110
      Top             =   900
   End
   Begin KazaaCheat.DownLoad dl 
      Left            =   4080
      Top             =   390
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin MSComctlLib.ProgressBar prog 
      Height          =   225
      Left            =   780
      TabIndex        =   6
      Top             =   2970
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   2910
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1235
            MinWidth        =   1235
            Text            =   "kc ¹·¹"
            TextSave        =   "kc ¹·¹"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab tab 
      Height          =   2925
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   5159
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Main"
      TabPicture(0)   =   "frmMain.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblHaxorTheNetwork"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "drv"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "dir"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "file"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Settings"
      TabPicture(1)   =   "frmMain.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtBuff"
      Tab(1).Control(1)=   "mp3Only"
      Tab(1).Control(2)=   "Label8"
      Tab(1).Control(3)=   "Label7"
      Tab(1).Control(4)=   "Label6"
      Tab(1).Control(5)=   "Label5"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Help/About"
      TabPicture(2)   =   "frmMain.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label4"
      Tab(2).Control(1)=   "Label3"
      Tab(2).Control(2)=   "Label2"
      Tab(2).Control(3)=   "Label1"
      Tab(2).ControlCount=   4
      Begin VB.TextBox txtBuff 
         Height          =   315
         Left            =   -72960
         TabIndex        =   14
         Text            =   "2048"
         ToolTipText     =   "Higher is faster... But not too high!"
         Top             =   1110
         Width           =   765
      End
      Begin VB.CheckBox mp3Only 
         Caption         =   "Fake Transfer only Mp3s"
         Height          =   315
         Left            =   -74760
         TabIndex        =   11
         Top             =   510
         Width           =   2205
      End
      Begin VB.FileListBox file 
         Height          =   1845
         Left            =   2310
         TabIndex        =   4
         Top             =   480
         Width           =   2205
      End
      Begin VB.DirListBox dir 
         Height          =   1890
         Left            =   180
         TabIndex        =   3
         Top             =   840
         Width           =   2025
      End
      Begin VB.DriveListBox drv 
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Top             =   480
         Width           =   2025
      End
      Begin VB.Label Label8 
         Caption         =   "THIS PROGRAM ONLY WORKS FOR KAZAA 2.0 (NOT >2.0) - THEY DISABLED THIS HACK!!!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   -74760
         TabIndex        =   16
         Top             =   1590
         Width           =   4215
      End
      Begin VB.Label Label7 
         Caption         =   "k/Bytes"
         Height          =   225
         Left            =   -72090
         TabIndex        =   15
         Top             =   1140
         Width           =   765
      End
      Begin VB.Label Label6 
         Caption         =   "Download Buffer size is:"
         Height          =   285
         Left            =   -74760
         TabIndex        =   13
         Top             =   1140
         Width           =   1875
      End
      Begin VB.Label Label5 
         Caption         =   "Avoids transferring those BIG mpegs, EXEs..."
         Height          =   255
         Left            =   -74490
         TabIndex        =   12
         Top             =   780
         Width           =   3375
      End
      Begin VB.Label Label4 
         Caption         =   $"frmMain.frx":0054
         Height          =   585
         Left            =   -74790
         TabIndex        =   10
         Top             =   2040
         Width           =   4215
      End
      Begin VB.Label Label3 
         Caption         =   "1) Navigate to your Kazaa Shared Files directory, then press the H4XOR the kNetwork button. (Make sure Kazaa is running)."
         Height          =   615
         Left            =   -74340
         TabIndex        =   9
         Top             =   1080
         Width           =   3645
      End
      Begin VB.Label Label2 
         Caption         =   "How to Use:"
         Height          =   225
         Left            =   -74670
         TabIndex        =   8
         Top             =   810
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "By: Eric Kirse"
         Height          =   195
         Left            =   -74820
         TabIndex        =   7
         Top             =   510
         Width           =   930
      End
      Begin VB.Label lblHaxorTheNetwork 
         Caption         =   "H4XOR the kNetwork!"
         Height          =   225
         Left            =   2580
         TabIndex        =   5
         Top             =   2490
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' KaZaA Participation Hacker
'   by: Eric Kirse
'   http dl control: its on PSC, and not by me
'
'   pretty much everything is documented...
'
'   this was hacked out in like 20 min, cause i got the idea
'       and wanted to see if it was possible
'
'   turns out, stuff like this already exists, and kazaa has
'   already blocked them in versions >= 2.1
'
'   anyhow, its fun being a supreme being on kazaa network
'       since most people prolly havent dl-ed the latest version

Dim fileNum As Integer
Dim transAll As Boolean
Dim isComplete As Boolean
Dim mp3sOnly As Boolean
Dim showVote As Boolean
' some module level variables i use

Private Sub dir_Change()
    ' update the file listing...
    file.Path = dir.Path
    pt 3, file.ListCount & " files in directory."
End Sub

Private Sub dl_DLComplete()
    ' download is complete, so change the variable
    '   and move on to the next file
    ' also delete our temp piece of trash
    
    isComplete = True
    fileNum = fileNum + 1
    file.ListIndex = fileNum
    Kill dir.Path & "\tmp.kcf"
End Sub

Private Sub dl_Rate(lpRate As String)
    pt 3, "File " & (fileNum + 1) & " Xfer Rate: " & lpRate
End Sub

Private Sub dl_RecievedBytes(lnumBYTES As Long)
    prog.Value = lnumBYTES
End Sub

Private Sub pt(ByVal panel As Integer, ByVal ptext As String)
    ' just a quicker way for me to change panel text...
    status.Panels.Item(panel).Text = ptext
End Sub

Private Sub drv_Change()
    ' ummm.
    dir.Path = drv.Drive
End Sub

Private Sub Form_Load()
    fileNum = 0
    isComplete = True
    transAll = True
    MsgBox "Sometimes the program stops abnormally..."
    MsgBox "This 90% of the time comes from setting the buffer size TOO HIGH... so keep it low!"
End Sub

Private Sub lblHaxorTheNetwork_Click()
    ' begin hacking the kazaa network =)
    isComplete = True
    tmrHackIt.Enabled = True
    dl.CHUNK = Val(txtBuff.Text)
    If Not (showVote) Then
        frmVote.Show 1
        showVote = True
    End If
End Sub

Private Sub mp3Only_Click()
    If mp3Only.Value = vbChecked Then
        mp3sOnly = True
    Else
        mp3sOnly = False
    End If
End Sub

Private Sub tmrHackIt_Timer()
    prog.Value = 0
    
    If (isComplete) Then
        ' is the file thats being downloaded complete??
        '   > basically checks can we start a new dl
        If (fileNum < file.ListCount) Then
            ' are we at the end of the filelist
            If (mp3sOnly) Then
                ' fake transfer only mp3s?
                If (LCase(Right(file.FileName, 3))) = "mp3" Then
                    isComplete = False
                        ' were downloading a file now
                    dl.Url = "http://127.0.0.1:1214/" & file.FileName
                        ' set the local url of the file to fake x-fer
                    dl.GetFileInformation
                        ' get its information
                    If dl.FileSize = 0 Then
                        ' just making sure weve got a good file
                        isComplete = True
                        fileNum = fileNum + 1
                        file.ListIndex = fileNum
                        Exit Sub
                    End If
                    prog.Min = 0    ' set up our progress bar
                    prog.Max = dl.FileSize
                    dl.SaveLocation = dir.Path & "\tmp.kcf"
                        ' set the save location of our file to be fake x-fered
                    dl.DownLoad
                        ' download it!
                Else
                    fileNum = fileNum + 1
                        ' next file...
                    file.ListIndex = fileNum
                End If
            Else
                ' same as documented above, 'cept this code just handles
                '   when we dl all files (not just mp3s)
                isComplete = False
                dl.Url = "http://127.0.0.1:1214/" & file.FileName
                dl.GetFileInformation
                If dl.FileSize = 0 Then
                    isComplete = True
                    fileNum = fileNum + 1
                    file.ListIndex = fileNum
                    Exit Sub
                End If
                prog.Min = 0
                prog.Max = dl.FileSize
                dl.SaveLocation = dir.Path & "\tmp.kcf"
                dl.DownLoad
            End If
        Else
            ' were done downloading the files...
            isComplete = True
            fileNum = 0
            file.ListIndex = fileNum
            pt 3, "check your rating..."
            tmrHackIt.Enabled = False
        End If
    End If
End Sub
