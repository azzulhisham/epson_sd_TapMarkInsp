Attribute VB_Name = "YdciApi"
'************************************ ��` ************************************

Public Const YDCI_OPEN_NORMAL = 0
Public Const YDCI_OPEN_OUT_NOT_INIT = &H1


'************************************ API *************************************

Declare Function YdciOpen Lib "Ydci.dll" (ByVal wBoardSw As Integer, _
                                          ByVal lpszModelName As String, _
                                          ByRef pwID As Integer, _
                                          Optional ByVal wMode As Integer = YDCI_OPEN_NORMAL) _
                                          As Long
Declare Function YdciClose Lib "Ydci.dll" (ByVal wID As Integer) _
                                           As Boolean
Declare Function YdciDioInput Lib "Ydci.dll" (ByVal wID As Integer, _
                                              ByRef pbyData As Byte, _
                                              ByVal wStart As Integer, _
                                              ByVal wCount As Integer) _
                                              As Long
Declare Function YdciDioOutput Lib "Ydci.dll" (ByVal wID As Integer, _
                                               ByRef pbyData As Byte, _
                                               ByVal wStart As Integer, _
                                               ByVal wCount As Integer) _
                                               As Long
Declare Function YdciDioOutputStatus Lib "Ydci.dll" (ByVal wID As Integer, _
                                               ByRef pbyData As Byte, _
                                               ByVal wStart As Integer, _
                                               ByVal wCount As Integer) _
                                               As Long
Declare Function YdciRlyOutput Lib "Ydci.dll" (ByVal wID As Integer, _
                                               ByRef pbyData As Byte, _
                                               ByVal wStart As Integer, _
                                               ByVal wCount As Integer) _
                                               As Long
Declare Function YdciRlyOutputStatus Lib "Ydci.dll" (ByVal wID As Integer, _
                                               ByRef pbyData As Byte, _
                                               ByVal wStart As Integer, _
                                               ByVal wCount As Integer) _
                                               As Long


'******************************** �G���[�R�[�h ********************************
Public Const YDCI_RESULT_SUCCESS = 0                    ' ����I��
Public Const YDCI_RESULT_ERROR = &HCE000001             ' �G���[
Public Const YDCI_RESULT_NOT_OPEN = &HCE000002          ' �I�[�v������Ă��Ȃ�
Public Const YDCI_RESULT_ALREADY_OPEN = &HCE000003      ' �I�[�v���ς�
Public Const YDCI_RESULT_INVALID_ID = &HCE000004        ' ID���s��
Public Const YDCI_RESULT_INVALID_EPNO = &HCE000005      ' EP�w�肪�s��(�����G���[)
Public Const YDCI_RESULT_CANNOT_OPEN = &HCE000006       ' �f�o�C�X���I�[�v���ł��Ȃ�����
Public Const YDCI_RESULT_PARAMETER_ERROR = &HCE000008   ' �����s��
Public Const YDCI_RESULT_MEM_ALLOC_ERROR = &HCE000009   ' �������m�ۃG���[
Public Const YDCI_RESULT_MODELNAME_ERROR = &HCE00000A   ' �^���w�肪�s��
Public Const YDCI_RESULT_NOT_SUPPORTED = &HCE00000C     ' �T�|�[�g����Ă��Ȃ�
Public Const YDCI_RESULT_INVALID_BOARD_SW = &HCE000013  ' �{�[�h���ʃX�C�b�`���s��

Public Const YDCI_RESULT_FATAL_ERROR = &HCEFFFFFF       ' �v���I�G���[

