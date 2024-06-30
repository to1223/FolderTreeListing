Attribute VB_Name = "ErrorNumber"
Option Explicit

Private Const UserErrorStartNumber As Long = vbObjectError + 2000

''' 再利用性や凝集度、結合度の観点から、クラスが出すエラーのコードはクラスに持たせるべき。

Enum ErrorType
    
    ErrorTypeGeneral = UserErrorStartNumber
    ErrorTypeFolderNotFound

End Enum
