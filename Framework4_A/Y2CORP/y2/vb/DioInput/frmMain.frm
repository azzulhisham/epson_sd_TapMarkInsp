VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  '�Œ�(����)
   Caption         =   "DIO Input Sample for Visual Basic"
   ClientHeight    =   345
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   345
   ScaleWidth      =   4830
   StartUpPosition =   3  'Windows �̊���l
   Begin VB.Menu mnuFile 
      Caption         =   "�t�@�C��"
      Begin VB.Menu mnuExit 
         Caption         =   "�I��"
      End
   End
   Begin VB.Menu mnuBoard 
      Caption         =   "�{�[�h"
      Begin VB.Menu mnuOpen 
         Caption         =   "�I�[�v��"
      End
   End
   Begin VB.Menu mnuInput 
      Caption         =   "����"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "�w���v"
      Begin VB.Menu mnuAbout 
         Caption         =   "�o�[�W�������"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Unload(Cancel As Integer)
    If g_intOpenFlag = DEVICE_OPENED Then
        If YdciClose(g_intID) = False Then
            MsgBox "�N���[�Y���s", vbExclamation
        End If
    End If
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuInput_Click()
    frmInput.Show vbModal, Me
End Sub

Private Sub mnuOpen_Click()
    frmOpen.Show vbModal, Me
End Sub
