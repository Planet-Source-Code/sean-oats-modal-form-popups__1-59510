VERSION 5.00
Begin VB.Form frmYellow 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Yellow Form"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuYellowPopUp 
      Caption         =   "Yellow PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuWhite 
         Caption         =   "Show the White form"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close the yellow form"
      End
   End
End
Attribute VB_Name = "frmYellow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum PopupResult
    NONE = 0
    SHOW_WHITE = 1
    CLOSE_WINDOW = 2
End Enum

Dim intPopUpResult As PopupResult

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
        intPopUpResult = NONE
    
        PopupMenu mnuYellowPopUp
        
        Select Case intPopUpResult
        
            Case PopupResult.SHOW_WHITE
                frmWhite.Show 1
            
            Case PopupResult.CLOSE_WINDOW
                Unload Me
        End Select
    
    End If
    
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub mnuClose_Click()
    intPopUpResult = CLOSE_WINDOW
End Sub


Private Sub mnuWhite_Click()
    intPopUpResult = SHOW_WHITE
End Sub


