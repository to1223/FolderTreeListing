Attribute VB_Name = "Test"
Option Explicit

Public Sub Test1()

    Const RootPath As String = "C:\Users\Takuya\Projects\FolderTreeListing\TestData\Root"
    Const TriggerFolderName As String = "sec_B"
    Const TriggerSearchLevel As Long = 2
    Const FolderListLevel As Long = 2
    Const IgnoreCase As Boolean = True

    Dim builder As DataBuilder
    Set builder = New DataBuilder
    
    builder.Init RootPath, TriggerFolderName, TriggerSearchLevel, FolderListLevel, IgnoreCase
    builder.Build

End Sub
