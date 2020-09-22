VERSION 5.00
Begin VB.Form frmTray 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tray Icon"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   4740
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNotification 
      Height          =   1485
      Left            =   150
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   2565
      Width           =   4425
   End
   Begin VB.CommandButton cmdUpdateTip 
      Caption         =   "Update Tip Text"
      Height          =   315
      Left            =   2370
      TabIndex        =   7
      Top             =   615
      Width           =   1635
   End
   Begin VB.TextBox txtTipText 
      Height          =   315
      Left            =   2370
      MaxLength       =   63
      TabIndex        =   6
      Text            =   "This is a Tray Icon!"
      Top             =   180
      Width           =   2205
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Icon"
      Height          =   315
      Left            =   150
      TabIndex        =   5
      Top             =   615
      Width           =   1635
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show Icon"
      Height          =   315
      Left            =   150
      TabIndex        =   4
      Top             =   180
      Width           =   1635
   End
   Begin VB.PictureBox picChange 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   2
      Left            =   1560
      Picture         =   "frmTray.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   1560
      Width           =   480
   End
   Begin VB.PictureBox picChange 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   855
      Picture         =   "frmTray.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   1560
      Width           =   480
   End
   Begin VB.PictureBox picChange 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   150
      Picture         =   "frmTray.frx":0884
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   1560
      Width           =   480
   End
   Begin VB.PictureBox picNotify 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   3885
      Picture         =   "frmTray.frx":0CC6
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Notification:"
      Height          =   195
      Index           =   1
      Left            =   150
      TabIndex        =   10
      Top             =   2280
      Width           =   840
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Choose the icon:"
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   8
      Top             =   1230
      Width           =   1200
   End
End
Attribute VB_Name = "frmTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents Tray As CTray
Attribute Tray.VB_VarHelpID = -1

Private Sub Form_Load()

    Set Tray = New CTray
    
    With Tray
        .TipText = Trim(txtTipText.Text)
        .PicBox = picNotify
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set Tray = Nothing

End Sub

Private Sub cmdShow_Click()

    Tray.ShowIcon
    
End Sub

Private Sub cmdDelete_Click()

    Tray.DeleteIcon
    
End Sub

Private Sub cmdUpdateTip_Click()

    Tray.TipText = Trim(txtTipText.Text)
    
End Sub

Private Sub picChange_Click(Index As Integer)

    picNotify.Picture = picChange(Index).Picture
    
End Sub

Private Sub Tray_LButtonDblClick()

    With txtNotification
        .Text = "Left Button Double Click" & vbCrLf & .Text
    End With
    
End Sub

Private Sub Tray_LButtonDown()

    With txtNotification
        .Text = "Left Button Down" & vbCrLf & .Text
    End With
    
End Sub

Private Sub Tray_LButtonUp()

    With txtNotification
        .Text = "Left Button Up" & vbCrLf & .Text
    End With
    
End Sub

Private Sub Tray_RButtonDblClick()

    With txtNotification
        .Text = "Right Button Double Click" & vbCrLf & .Text
    End With
    
End Sub

Private Sub Tray_RButtonDown()

    With txtNotification
        .Text = "Right Button Down" & vbCrLf & .Text
    End With
    
End Sub

Private Sub Tray_RButtonUp()

    With txtNotification
        .Text = "Right Button Up" & vbCrLf & .Text
    End With
    
End Sub
