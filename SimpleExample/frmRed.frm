VERSION 5.00
Begin VB.Form frmRed 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Red Form"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   7950
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label2 
      Caption         =   "Once you understand what is happening, and how to handle it, see the Fancy Example."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   3240
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Right-Click the form to see examples"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   2160
      TabIndex        =   0
      Top             =   600
      Width           =   3255
   End
   Begin VB.Menu mnuMainPopUp 
      Caption         =   "Main Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuRight 
         Caption         =   "See Right Example"
      End
      Begin VB.Menu mnuWrong 
         Caption         =   "See Wrong Example"
      End
   End
End
Attribute VB_Name = "frmRed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnShowRightWay As Boolean


Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '** call the form's MouseDown event in case the user Right-Clicks on
    '**   the label instead of the form
    Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '** call the form's MouseDown event in case the user Right-Clicks on
    '**   the label instead of the form
    Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
        '** the user Right-Clicked
        
        '** set the boolean to false
        blnShowRightWay = False
        
        '** chow the popup menu
        PopupMenu mnuMainPopUp
        
        '** check to see if the user selected 'See Right Example'
        If blnShowRightWay Then
        
            '** now that the menu is unloaded, launch the blue form
            frmBlue.Show 1
        
        Else
            '** Either the user selected 'See Wrong Example' or clicked off the menu
        End If
        
    End If

End Sub


Private Sub mnuRight_Click()
    '** set the boolean to True, indicating that the blue form should
    '**   be launched after the menu is unloaded
    blnShowRightWay = True
End Sub

Private Sub mnuWrong_Click()
    '** launch the blue form
    frmBlue.Show 1
End Sub
