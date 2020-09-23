VERSION 5.00
Begin VB.Form frmBlue 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blue Form"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   $"frmBlue.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
   Begin VB.Menu mnuBluePopUp 
      Caption         =   "Blue PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuClose 
         Caption         =   "Close this form"
      End
      Begin VB.Menu mnuExplain 
         Caption         =   "Tell Me Why This Happens..."
      End
   End
End
Attribute VB_Name = "frmBlue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
    
        '** The user Right-Clicked.
        '** Try to show the popup menu.
        '** If the user selected 'See Wrong Example', this will not work.
        PopupMenu mnuBluePopUp
        
    End If

End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '** call the form's MouseDown event incase the user clicked the label
    Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub mnuClose_Click()
    '** close the blue form
    Unload Me
End Sub

Private Sub mnuExplain_Click()

    '** tell the user why this happens
    MsgBox "Windows only provides for one popup menu at a time, " _
        & "even in seperate programs.  Since this form was shown " _
        & "modaly, the the red forms's popup menu's click event has " _
        & "not finished executing, and therefore the red form still has " _
        & "control over the one allowed popup menu." & Chr(13) & Chr(13) _
        & "This is a particalarly hard bug to track down, because " _
        & "if this is encountered, no error is thrown.  The call just " _
        & "simply doesn't do anything."
        
End Sub
