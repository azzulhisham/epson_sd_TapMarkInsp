VERSION 5.00
Begin VB.Form frmOutput 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "出力確認"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   2790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtOutputData 
      Height          =   300
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   1
      Text            =   "0"
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox txtOutputNo 
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Text            =   "0"
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblOutputData 
      AutoSize        =   -1  'True
      Caption         =   "(0:OFF, 1:ON, 2:変更なし)"
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   1890
   End
   Begin VB.Label lblOutputData 
      AutoSize        =   -1  'True
      Caption         =   "出力データ"
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   780
      Width           =   855
   End
   Begin VB.Label lblOutputNo 
      AutoSize        =   -1  'True
      Caption         =   "出力No."
      Height          =   180
      Left            =   240
      TabIndex        =   4
      Top             =   300
      Width           =   600
   End
End
Attribute VB_Name = "frmOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngResult As Long
    Dim bytStart As Byte
    Dim bytData As Byte
    
    bytStart = Val(txtOutputNo.Text)
    bytData = Val(txtOutputData.Text)
    
    lngResult = YdciDioOutput(g_intID, bytData, bytStart, 1)
    If lngResult = YDCI_RESULT_SUCCESS Then
        MsgBox "出力正常終了", vbInformation
    Else
        MsgBox "YdciDioOutput ERROR : 0x" & Hex(lngResult), vbExclamation
    End If
End Sub
