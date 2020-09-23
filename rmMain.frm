VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Draw Percent"
   ClientHeight    =   1020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   3045
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDraw 
      Caption         =   "Draw Percent"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.PictureBox picPercent 
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
   Begin VB.Line lneSep 
      BorderColor     =   &H8000000C&
      Index           =   1
      X1              =   0
      X2              =   3000
      Y1              =   470
      Y2              =   470
   End
   Begin VB.Line lneSep 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   0
      X2              =   3000
      Y1              =   480
      Y2              =   480
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************
'*            Draw Percent by Joseph Huntley              *
'*               joseph_huntley@email.com                 *
'*                http://joseph.vr9.com                   *
'*                                                        *
'*  Made:  October 1, 1999                                *
'*  Level: Beginner                                       *
'**********************************************************
'*   The form here are only used to demonstrate how to    *
'* use the function 'DrawPercent'. You may copy the       *
'* functions into your project for use. If you need any   *
'* help please e-mail me.                                 *
'**********************************************************
'* Notes: The subroutine 'Pause' is not required for      *
'*        use of the function 'DrawPercent'. It is only   *
'*        used to slow down the progress so you can see   *
'*        what this example does in detail.               *
'**********************************************************

Sub Pause(dblInterval As Double)
 'NOTE: You do not need this function! It is only to
 '      slow down the progress bar for the form.
 
 Dim dblCurrent As Double
 
 dblCurrent# = Timer
 
   Do While Timer - dblCurrent# < dblInterval#
     DoEvents
   Loop
 
End Sub
Sub DrawPercent(picPic As Object, lngPercent As Long)
'**********************************************************
'*            Draw Percent by Joseph Huntley              *
'*               joseph_huntley@email.com                 *
'*                http://joseph.vr9.com                   *
'**********************************************************
'*   You may use this code freely as long as credit is    *
'* given to the author, and the header remains intact.    *
'**********************************************************


'--------------------- The Arguments -----------------------
'picPic     - The object to draw on.
'lngPercent - The percentage to print out.
'-----------------------------------------------------------

'Description: Draws a percentage bar on an object.

  With picPic
    .Cls
    .ScaleMode = vbpixel
    .DrawMode = vbNotXorPen
    .BackColor = vbWhite 'Change for different background color
    .ForeColor = vbBlue  'Change for different foreground color
    .AutoRedraw = True
    .CurrentX = .ScaleWidth / 2 - .TextWidth(CStr(lngPercent&) & "%") / 2
    .CurrentY = .ScaleHeight / 2 - .TextHeight(CStr(lngPercent&) & "%") / 2
    picPic.Print CStr(lngPercent&) & "%"
    picPic.Line (1, 1)-((.Width / 100) * lngPercent&, .Height), vbBlue, BF
    .Refresh
  End With
  
End Sub

Private Sub cmdDraw_Click()
   Dim lngCurPercent As Long
   
     For lngCurPercent& = 1 To 100
        Call DrawPercent(picPercent, lngCurPercent&)
        Call Pause(0.001)
     Next lngCurPercent&
   
End Sub

Private Sub Form_Load()

End Sub
