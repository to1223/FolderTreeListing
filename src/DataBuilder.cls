VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private buf_() As String
Private dataCount_ As Long

Private rootPath_ As String
Private triggerFolderName_ As String
Private triggerSearchLevel_ As Long
Private folderListLevel_ As Long
Private ignoreCase_ As Boolean


Public Enum DataBuilderErrorType

    DataBuilderErrorGeneral = vbObjectError + 2000
    DataBuilderErrorFolderNotFound

End Enum


''' �f�[�^�̏�����
Public Sub Init( _
    ByVal root_path As String, _
    ByVal trigger_folder_name As String, _
    ByVal trigger_search_level As Long, _
    ByVal folder_list_level As Long, _
    ByVal ignore_case As Boolean _
)

    Const InitialBufferSize As Long = 1
    Debug.Assert InitialBufferSize >= 1
    
    ' �z��̍s���� Full path (�n�C�p�[�����N�p) + Level��
    ReDim buf_(folder_list_level, InitialBufferSize - 1)
    dataCount_ = 0
    
    rootPath_ = root_path
    triggerFolderName_ = trigger_folder_name
    triggerSearchLevel_ = trigger_search_level
    folderListLevel_ = folder_list_level
    ignoreCase_ = ignore_case
    
End Sub


''' �f�[�^���쐬����
Public Sub Build()

    ' �t�H���_�m�F

    If triggerFolderName_ = "" Then
        AddFolderRecursive rootPath_, 1
    Else
        RecursiveProcess rootPath_, 1
    End If

    ' �z��T�C�Y���]���ȏꍇ�͏k��
    

End Sub


''' �g���K�[�ƂȂ�t�H���_���w�肵���K�w�܂ŒT���A���݂���΂��̔z���̃t�H���_���w�肵���K�w�܂Ń��X�g�ɒǉ�����B
Private Sub RecursiveProcess( _
    ByVal folder_path As String, _
    ByVal this_level As Long _
)

    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    
    Dim sub_folder As Variant
    Dim sub_folder_name As String
    Dim trigger_folder_name As String
    
    For Each sub_folder In fso.GetFolder(folder_path).SubFolders
        sub_folder_name = sub_folder.Name
        trigger_folder_name = triggerFolderName_
        
        If ignoreCase_ Then
            sub_folder_name = UCase(sub_folder_name)
            trigger_folder_name = UCase(trigger_folder_name)
        End If
    
        If sub_folder_name = trigger_folder_name Then
            AddFolderRecursive sub_folder.Path, 1
            GoTo Continue
        End If
        
        If this_level + 1 <= triggerSearchLevel_ Then
            RecursiveProcess sub_folder.Path, this_level + 1
        End If
        
Continue:
    Next

Closing:
    Set fso = Nothing

End Sub


''' �w�肵���t�H���_�̔z���̃t�H���_���w�肵���K�w�܂Ń��X�g�ɒǉ�����B
Private Sub AddFolderRecursive( _
    ByVal folder_path As String, _
    ByVal this_level As Long _
)

    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    
    Dim sub_folder As Variant
    
    For Each sub_folder In fso.GetFolder(folder_path).SubFolders
        AddData sub_folder.Path, this_level
        
        ' ���̊K�w�� max_level �𒴂��Ȃ���΁A�ċA�Ăяo��
        If this_level + 1 <= folderListLevel_ Then
            AddFolderRecursive sub_folder.Path, this_level + 1
        End If
        
    Next

Closing:
    Set fso = Nothing

End Sub


''' 1���ڂ̃f�[�^��ǉ�����
Private Sub AddData(ByVal folder_path As String, ByVal this_level As Long)

    Dim fso As FileSystemObject
    Set fso = New FileSystemObject

    ' �z��T�C�Y������Ȃ���� ReDim
    If UBound(buf_, 2) = dataCount_ - 1 Then
        ReDim Preserve buf_(folderListLevel_, UBound(buf_, 2) + 1)
    End If
    
    ' path ����Athis_level ���̃t�H���_�����o���A�z��ɒǉ�
    Dim i As Long
    Dim target_path As String
    
    target_path = folder_path
    
    ' �n�C�p�[�����N�p�̃t���p�X���i�[
    buf_(0, dataCount_) = target_path
    
    ' �K�w�ɉ������t�H���_�����i�[
    For i = this_level To 1 Step -1
        buf_(i, dataCount_) = fso.GetFileName(target_path)
        target_path = fso.GetParentFolderName(target_path)
    Next
            
    dataCount_ = dataCount_ + 1
    
'    Debug.Print "--- Count = " & dataCount_ & " ---"
'    For i = 0 To UBound(buf_, 1)
'        Debug.Print buf_(i, dataCount_ - 1)
'    Next
    
End Sub


Public Property Get Data() As Variant

    Data = buf_

End Property

