Attribute VB_Name = "Main"
Option Explicit

Public Const macroName As String = "ConvertProp"
Public Const macroSection As String = "Main"

Dim swApp As Object
Dim patternOfChanging As Dictionary

Sub Main()
    Set swApp = Application.SldWorks
    Dim doc As ModelDoc2
    
    Set doc = swApp.ActiveDoc
    If doc Is Nothing Then Exit Sub
    If doc.GetType = swDocDRAWING Then
        MsgBox "Макрос не работает с чертежами!", vbCritical
    Else
        If doc.GetType = swDocPART Then
            MainForm.chkRecursive.value = False
            MainForm.chkRecursive.Visible = False
        End If
        MainForm.Show
    End If
End Sub

Sub RunRenaming(needOverwriteExisting As Boolean, needRemoveOld As Boolean, needRenameRecursively As Boolean)
    Dim currentDoc As ModelDoc2
    Dim component As ModelDoc2
    Dim i As Integer
       
    Init
    Set currentDoc = swApp.ActiveDoc
    RenameAllProperties currentDoc, needOverwriteExisting, needRemoveOld
    If currentDoc.GetType = swDocASSEMBLY And needRenameRecursively Then
        Dim deps As Variant  'String() or empty
        deps = swApp.GetDocumentDependencies2(currentDoc.GetPathName, True, False, False)
        If Not IsArrayEmpty(deps) Then
            For i = LBound(deps) + 1 To UBound(deps) Step 2  'full names
                Set component = swApp.GetOpenDocumentByName(deps(i))
                RenameAllProperties component, needOverwriteExisting, needRemoveOld
            Next
        End If
    End If
End Sub

Function Init() 'for mask button
    Set patternOfChanging = New Dictionary
    patternOfChanging.Add "Weight", "Масса"
    patternOfChanging.Add "Material", "Материал"
    patternOfChanging.Add "Konstruowa?", "Разработал"
    patternOfChanging.Add "Kre?li?", "Начертил"
    patternOfChanging.Add "Sprawdzi?", "Проверил"
    patternOfChanging.Add "Zatwierdzi?", "Утвердил"
    patternOfChanging.Add "Nazwa przedmiotu", "Наименование"
    patternOfChanging.Add "Numer rysunku", "Обозначение"
    patternOfChanging.Add "Asortyment", "Заготовка"
    patternOfChanging.Add "Gabaryt S", "Типоразмер"
    patternOfChanging.Add "Gabaryt L", "Длина"
    patternOfChanging.Add "Gabaryt B", "Ширина"
    
    'patternOfChanging.Add "Detal / handlowe"
End Function

Sub RenameProperty(oldName As String, newName As String, propertyManager As CustomPropertyManager, _
                   needOverwriteExisting As Boolean, needRemoveOld As Boolean, ByRef wasRenaming As Boolean)
    Dim rawValue As String
    Dim resolvedValue As String
    Dim wasResolved As Boolean
    Dim overwriteExisting As swCustomPropertyAddOption_e
    Dim propertyType As swCustomInfoType_e
    Dim isRenamingSuccess As Boolean
        
    If needOverwriteExisting Then
        overwriteExisting = swCustomPropertyDeleteAndAdd
    Else
        overwriteExisting = swCustomPropertyOnlyIfNew
    End If
    propertyManager.Get5 oldName, False, rawValue, resolvedValue, wasResolved
    propertyType = propertyManager.GetType2(oldName)
    isRenamingSuccess = (propertyManager.Add3(newName, propertyType, rawValue, overwriteExisting) = swCustomInfoAddResult_AddedOrChanged)
    If isRenamingSuccess Then
        wasRenaming = True
        If needRemoveOld Then
            propertyManager.Delete2 oldName
        End If
    End If
End Sub

Sub RenameAllProperties(model As ModelDoc2, needOverwriteExisting As Boolean, _
                       needRemoveOld As Boolean)
    Dim x As Variant
    Dim confName As String
    Dim modelExtension As ModelDocExtension
    Dim wasRenaming As Boolean
    
    Set modelExtension = model.extension
    RenamePropertiesInConf "", modelExtension, needOverwriteExisting, needRemoveOld, wasRenaming
    For Each x In model.GetConfigurationNames
        confName = x
        RenamePropertiesInConf confName, modelExtension, needOverwriteExisting, needRemoveOld, wasRenaming
    Next
    If wasRenaming Then
        model.SetSaveFlag
    End If
End Sub

Sub RenamePropertiesInConf(conf As String, modelExtension As ModelDocExtension, _
                           needOverwriteExisting As Boolean, needRemoveOld As Boolean, _
                           ByRef wasRenaming As Boolean)
    Dim propertyManager As CustomPropertyManager
    Dim x As Variant
    Dim propName As String
    Dim allConfNames As Variant
    Dim newPropName As String
    
    Set propertyManager = modelExtension.CustomPropertyManager(conf)
    allConfNames = propertyManager.GetNames
    If Not IsArrayEmpty(allConfNames) Then
        For Each x In allConfNames
            propName = x
            If IsPolishProperty(propName, newPropName) Then
                RenameProperty propName, newPropName, propertyManager, needOverwriteExisting, needRemoveOld, wasRenaming
            End If
        Next
    End If
End Sub

Function IsPolishProperty(propName As String, ByRef newPropName As String) As Boolean
    Dim key As Variant
    
    For Each key In patternOfChanging.Keys
        If propName Like key Then
            IsPolishProperty = True
            newPropName = patternOfChanging(key)
            Exit Function
        End If
    Next
    IsPolishProperty = False
End Function

Function ExitApp() 'mask for button
    Unload MainForm
    End
End Function
