Attribute VB_Name = "ErrorNumber"
Option Explicit

Private Const UserErrorStartNumber As Long = vbObjectError + 2000

''' �ė��p����ÏW�x�A�����x�̊ϓ_����A�N���X���o���G���[�̃R�[�h�̓N���X�Ɏ�������ׂ��B

Enum ErrorType
    
    ErrorTypeGeneral = UserErrorStartNumber
    ErrorTypeFolderNotFound

End Enum
