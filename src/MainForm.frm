VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Конвертер свойств"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    ExitApp
End Sub

Private Sub btnRun_Click()
    RunRenaming chkOverwrite.value, chkRemoveOld.value, chkRecursive.value
    ExitApp
End Sub

Private Sub chkOverwrite_Click()
    SaveBoolSetting "overwrite", chkOverwrite.value
End Sub

Private Sub chkRemoveOld_Click()
    SaveBoolSetting "removeold", chkRemoveOld.value
End Sub

Private Sub UserForm_Initialize()
    chkOverwrite.value = GetBoolSetting("overwrite")
    chkRemoveOld.value = GetBoolSetting("removeold")
End Sub
