VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Form"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label2 
      Caption         =   $"frmMain.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   2400
      Width           =   6735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Right-Click this form to see examples"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   3495
   End
   Begin VB.Menu mnuMainPopUp 
      Caption         =   "Main PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "Show..."
         Begin VB.Menu mnuRed 
            Caption         =   "Red Form"
         End
         Begin VB.Menu mnuBlue 
            Caption         =   "Blue Form"
         End
         Begin VB.Menu mnuYellow 
            Caption         =   "Yellow Form"
         End
      End
      Begin VB.Menu mnuExplain 
         Caption         =   "Explain this to me..."
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum PopupResult
    NONE = 0
    RED_FORM = 1
    BLUE_FORM = 2
    YELLOW_FORM = 3
    EXPLAIN = 4
    CLOSE_WINDOW = 5
End Enum


Dim intPopUpResult As PopupResult


'** throw the arguments in case user clicks the labels instead of the form
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseDown Button, Shift, X, Y
End Sub
Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseDown Button, Shift, X, Y
End Sub



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
        '** the user right clicked
        
        intPopUpResult = NONE
        
        
        PopupMenu mnuMainPopUp
        
        '** figure out what the user clicked
        Select Case intPopUpResult
        
            Case PopupResult.RED_FORM
                frmRed.Show 1
                
            Case PopupResult.BLUE_FORM
                frmBlue.Show 1
            
            Case PopupResult.YELLOW_FORM
                frmYellow.Show 1
            
            Case PopupResult.EXPLAIN
                ExplainToUser
            
            Case PopupResult.CLOSE_WINDOW
                Unload Me
        
        End Select
    
    End If
End Sub



Private Sub mnuRed_Click()
    intPopUpResult = RED_FORM
End Sub

Private Sub mnuBlue_Click()
    intPopUpResult = BLUE_FORM
End Sub

Private Sub mnuYellow_Click()
    intPopUpResult = YELLOW_FORM
End Sub

Private Sub mnuExplain_Click()
    intPopUpResult = EXPLAIN
End Sub

Private Sub mnuExit_Click()
    intPopUpResult = CLOSE_WINDOW
End Sub


Private Sub ExplainToUser()

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


