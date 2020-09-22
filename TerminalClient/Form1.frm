VERSION 5.00
Object = "{FC7C887E-70BD-4ADB-8BED-8681D74F36D1}#1.0#0"; "msrdp.ocx"
Begin VB.Form frmMain 
   Caption         =   "Terminal Server Client"
   ClientHeight    =   1260
   ClientLeft      =   7005
   ClientTop       =   5835
   ClientWidth     =   3765
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   84
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   251
   Begin VB.ComboBox cboRes 
      CausesValidation=   0   'False
      Height          =   315
      ItemData        =   "Form1.frx":0442
      Left            =   840
      List            =   "Form1.frx":0458
      TabIndex        =   3
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&GO"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin MSTSCLibCtl.MsTscAx msts 
      Height          =   735
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   975
      Server          =   ""
      Domain          =   ""
      UserName        =   ""
      FullScreen      =   ""
      StartConnected  =   0
   End
   Begin VB.Label Label3 
      Caption         =   "Resolution"
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "UserName"
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Server"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Using the Microsoft Terminal Server Control
'
'Coded by Chris Peneguy
'
'Date Dec. 7 2001
'
'http://www.secureinsights.com
'
'chris@secureinsights.com
'
'Use this code any which way you like
'
Option Explicit
Dim Server As String    'Server Address
Dim UserName As String  'User Login Name
Dim resWidth As String  'Resolution Size - Width
Dim resHeight As String 'Resolution Size - Height
Dim Reso As String
Const FullScreenWarnTxt1 = "Your current security settings do not allow automatically switching to fullscreen mode."
Const FullScreenWarnTxt2 = "You can use ctrl-alt-pause to toggle your terminal services session to fullscreen mode"
Const FullScreenTitleTxt = "Terminal Services Connection "
Const ErrMsgText = "Error connecting to terminal server: "

Private Sub Form_Load()

msts.Visible = False

End Sub

Private Sub cmdGo_Click()


Server = txtServer.Text       'Server Address
UserName = txtUserName.Text   'User Login Name
Reso = cboRes.Text            'Resolution Temp

If Server = "" Then
  MsgBox "Please enter a Server address"
  txtServer.SetFocus
 Exit Sub
 
  ElseIf UserName = "" Then
    MsgBox "Please enter a UserName"
    txtUserName.SetFocus
  Exit Sub
  
   ElseIf Reso = "" Then
      MsgBox "Please choose a Resolution"
      cboRes.SetFocus
   Exit Sub

Else
 Call Res
 Call Connect
End If

End Sub

Sub Res()
 'Sets the Resolution size for the terminal based
 'on the resolution chosen in the combo box
 
 If Reso = "Full-Screen " Then
      resWidth = Screen.Width \ Screen.TwipsPerPixelX    'Converts from Twips to Pixels
      resHeight = Screen.Height \ Screen.TwipsPerPixelY  'Convers from Twis to Pixels
 Exit Sub
 
    ElseIf Reso = "800 x 600" Then
        resWidth = "800"
        resHeight = "600"
    Exit Sub
   
      ElseIf Reso = "1024 x 768" Then
          resWidth = "1024"
          resHeight = "768"
      Exit Sub
      
        ElseIf Reso = "1152 x 864" Then
            resWidth = "1152"
            resHeight = "864"
        Exit Sub
        
          ElseIf Reso = "1280 x 1024" Then
              resWidth = "1280"
              resHeight = "1024"
          Exit Sub
          
            ElseIf Reso = "1600 x 1200" Then
                resWidth = "1600"
                resHeight = "1200"
            Exit Sub
            
              Else
              MsgBox "Please Choose a Screen Resolution"
              
End If
        
End Sub


Sub Connect()
'Connecting to the Terminal Server

msts.Server = Server
msts.UserName = UserName

   If msts.SecuredSettingsEnabled Then
         msts.SecuredSettings.FullScreen = 1
         msts.DesktopHeight = resHeight
         msts.DesktopWidth = resWidth
         
       Else
          MsgBox (FullScreenWarnTxt1 & vbCrLf & FullScreenWarnTxt2)
          msts.DesktopHeight = resHeight
          msts.DesktopWidth = resWidth
   End If

msts.FullScreenTitle = FullScreenTitleTxt & vbCrLf & Server
msts.Connect


End Sub

