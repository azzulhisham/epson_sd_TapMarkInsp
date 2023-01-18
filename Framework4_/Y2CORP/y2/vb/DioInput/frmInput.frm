VERSION 5.00
Begin VB.Form frmInput 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "入力確認"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   2820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   Begin VB.TextBox txtInputNo 
      Height          =   300
      Left            =   1080
      TabIndex        =   0
      Text            =   "0"
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblInputNo 
      AutoSize        =   -1  'True
      Caption         =   "入力No."
      Height          =   180
      Left            =   240
      TabIndex        =   3
      Top             =   300
      Width           =   600
   End
End
Attribute VB_Name = "frmInput"
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
    Dim bytData As Byte
    Dim bytInputNo As Byte
    
    bytInputNo = Val(txtInputNo.Text)
    
    lngResult = YdciDioInput(g_intID, bytData, bytInputNo, 1)
    If lngResult = YDCI_RESULT_SUCCESS Then
        MsgBox "入力データ: " & bytData, vbInformation
    Else
        MsgBox "YdciDioInput ERROR : 0x" & Hex(lngResult), vbExclamation
    End If
End Sub
