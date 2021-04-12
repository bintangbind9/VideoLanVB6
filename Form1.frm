VERSION 5.00
Object = "{DF2BBE39-40A8-433B-A279-073F48DA94B6}#1.0#0"; "axvlc.dll"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   3600
      Left            =   360
      TabIndex        =   4
      Top             =   4680
      Width           =   2895
   End
   Begin VB.CommandButton cmdCapture 
      Caption         =   "Capture"
      Height          =   495
      Left            =   8280
      TabIndex        =   3
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdOn 
      Caption         =   "ON"
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Text            =   "http://192.168.1.18:4747/video"
      Top             =   4080
      Width           =   4455
   End
   Begin AXVLCCtl.VLCPlugin2 cam1 
      Height          =   3495
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   4455
      AutoLoop        =   0   'False
      AutoPlay        =   -1  'True
      Toolbar         =   -1  'True
      ExtentWidth     =   7858
      ExtentHeight    =   6165
      MRL             =   ""
      Object.Visible         =   -1  'True
      Volume          =   100
      StartTime       =   0
      BaseURL         =   ""
      BackColor       =   0
      FullscreenEnabled=   -1  'True
      Branding        =   -1  'True
   End
   Begin VB.Image img2 
      BorderStyle     =   1  'Fixed Single
      Height          =   3495
      Left            =   5160
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   4455
   End
   Begin VB.Image img1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3495
      Left            =   5160
      Stretch         =   -1  'True
      Top             =   360
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCapture_Click()
    Dim resp As Variant
    Dim ls_prev_path As String
    Dim ls_temp_path As String
    Dim pic As IPictureDisp
    
    ls_prev_path = CurDir
    ls_temp_path = App.Path + IIf(Right(App.Path, l) = "\", "", "\")
    ls_temp_path = ls_temp_path + "Temp\"
    
    If Dir(ls_temp_path, vbDirectory) <> "" Then
        If Dir(ls_temp_path + "*.bmp") <> "" Then Kill ls_temp_path + "*.bmp"
    Else
        MkDir ls_temp_path
    End If
    
    ChDir ls_temp_path
    
    'Capture
    cam1.playlist.pause
    cam1.video.takeSnapshot
    cam1.playlist.play
    
    DoEvents
    
    If Dir(ls_temp_path + "*.*") <> "" Then
        ls_pic_file = ls_temp_path + Dir(ls_temp_path + "*.*")
    End If
    
    Filename = ls_pic_file
    filejpg = App.Path & "\gbr\" & Format(Now, "YYYYMMdd HHmmss") & ".jpg"
    
    'Convert bpm to jpg
    dib = FreeImage_LoadEx(Filename)
    If (dib) Then
        Call FreeImage_SaveEx(dib, filejpg)
        Call FreeImage_Unload(dib)
    End If
    
    'Delete File Temp
    Kill Filename
    
    'Tampil hasil capture
    If Dir$(filejpg) <> "" Then
        img1.Picture = LoadPicture(filejpg)
    End If
    
    ChDir ls_prev_path
    
    File1.Refresh
End Sub

Private Sub cmdOn_Click()
    If Me.cmdOn.Caption = "ON" Then
        Me.cmdOn.Caption = "OFF"
        Me.cam1.playlist.Add (Me.Text1.Text)
        Me.cam1.playlist.play
    Else
        Me.cmdOn.Caption = "ON"
        Me.cam1.playlist.stop
    End If
End Sub

Private Sub File1_Click()
    img2.Picture = LoadPicture(App.Path & "\gbr\" & File1.Filename)
End Sub

Private Sub Form_Load()
    File1 = App.Path & "\gbr"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.cmdOn.Caption = "ON" Then
        cam1.playlist.stop
    End If
End Sub
