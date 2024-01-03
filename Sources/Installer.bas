Attribute VB_Name = "Installer"
' The code is taken from:
' https://jkp-ads.com/articles/distributemacro10.asp
' Corrections: Application.LibraryPath -> Application.UserLibraryPath

Option Explicit

Dim vReply As Variant
Dim AddInLibPath As String
Dim CurAddInPath As String
Const sAppName As String = "SVGlib"
Const sFilename As String = sAppName & ".xlam"
Const sRegKey As String = "SVGlib"    ''' RegKey for settings
Const Version As String = "0.0.6"

Sub Install()
    vReply = MsgBox("This will install " & sAppName & vbNewLine & _
    "in your default Add-in directory." & vbNewLine & vbNewLine & _
    "Proceed?", vbYesNo, sAppName & " Setup")
    If vReply = vbYes Then
        On Error Resume Next
        Workbooks(sFilename).Close False
        If Application.OperatingSystem Like "*Win*" Then
            CurAddInPath = ThisWorkbook.Path & "\" & sFilename
            AddInLibPath = Application.UserLibraryPath & "\" & sFilename
            'User librarypath does not have a trailing path separator
            'AddInLibPath = Application.UserLibraryPath & sFilename
        Else
            'MAC syntax differs from Win
            CurAddInPath = ThisWorkbook.Path & ":" & sFilename
            AddInLibPath = Application.UserLibraryPath & sFilename
        End If
        On Error Resume Next
        FileCopy CurAddInPath, AddInLibPath
        If Err.Number <> 0 Then
            SomeThingWrong
            Exit Sub
        End If
        With AddIns.Add(FileName:=AddInLibPath)
            .Installed = True
        End With
        
        
        Dim config$
        config = "<?xml version=""1.0"" encoding=""utf-8""?>" & vbCrLf
        config = "<AddIn Version=""" & Version & """ Name=""" & sRegKey & """ />"
        
        Dim iFile&
        iFile = FreeFile
        Open Application.UserLibraryPath & "\" & sRegKey & ".xml" For Binary As #iFile
        Put #iFile, , config
        Close #iFile
        
        Application.CommandBars(1).FindControl(ID:=943, recursive:=True).Execute
    Else
        vReply = MsgBox(prompt:="Install Cancelled", Buttons:=vbOKOnly, Title:=sAppName & " Setup")
    End If
End Sub

Sub SomeThingWrong()
    If Application.OperatingSystem Like "*Win*" Then
        vReply = MsgBox(prompt:="Something went wrong during copying" _
        & vbNewLine & "of the add-in to your add-in directory:" _
        & vbNewLine & vbNewLine & Application.LibraryPath & "\" _
        & vbNewLine & vbNewLine & "You can install " & sAppName _
        & " manually by copying the file" & vbNewLine & _
        sFilename & " to this directory yourself and installing the addin" _
        & vbNewLine & "using Tools, Addins from the menu of Excel." _
        & vbNewLine & vbNewLine & "Don't press OK yet, first do" _
        & " the copying from Windows Explorer." & vbNewLine _
        & "It gives you the opportunity to ALT-TAB back to Excel" _
        & vbNewLine & "to read this text." _
        , Buttons:=vbOKOnly, Title:=sAppName & " Setup")
    Else
        vReply = MsgBox(prompt:="Something went wrong during copying" _
        & vbNewLine & "of the add-in to your add-in directory:" _
        & vbNewLine & vbNewLine & Application.LibraryPath _
        & vbNewLine & vbNewLine & "You can install " & sAppName & _
        " manually by copying the file" & vbNewLine & sFilename & _
        " to this directory yourself and installing the addin" _
        & vbNewLine & "using Tools, Addins from the menu of Excel." _
        & vbNewLine & vbNewLine & "Don't press OK yet," _
        & " first do the copying in the Finder." _
        & vbNewLine & "It gives you the opportunity to Command-TAB back to Excel" _
        & vbNewLine & "to read this text." _
        , Buttons:=vbOKOnly, Title:=sAppName & " Setup")
    End If
End Sub

Sub Uninstall()
    vReply = MsgBox("This will remove the " & sAppName & vbNewLine & _
    "from your system." & vbNewLine & _
    vbNewLine & "Proceed?", vbYesNo, sAppName & " Setup")
    If vReply = vbYes Then
        If Application.OperatingSystem Like "*Win*" Then
            CurAddInPath = ThisWorkbook.Path & "\" & sFilename
            AddInLibPath = Application.UserLibraryPath & "\" & sFilename
        Else
            'MAC syntax differs from Win
            CurAddInPath = ThisWorkbook.Path & ":" & sFilename
            AddInLibPath = Application.UserLibraryPath & sFilename
        End If
        On Error Resume Next
        Workbooks(sFilename).Close False
        Kill AddInLibPath
        DeleteSetting sRegKey
        Kill Application.UserLibraryPath & "\" & sRegKey & ".xml"
        
        MsgBox " The " & sAppName & " has been removed from your computer." _
        & vbNewLine & "To complete the removal, please select the " & sAppName _
        & vbNewLine & "in the following dialog and acknowledge the removal" _
        , vbInformation + vbOKOnly
        Application.CommandBars(1).FindControl(ID:=943, recursive:=True).Execute
    End If
End Sub

Sub ExportVba()
    Dim qwe As Object
    Dim VBComp As VBIDE.VBComponent
    Dim ExportFile$, Directory$, FileName$, Extension$, i&
    Dim allowedModules As New Collection
    
    Directory = Application.ActiveWorkbook.Path & "\Sources\"
    Call allowedModules.Add("Installer")
    
    For Each qwe In ActiveWorkbook.VBProject.VBComponents
        For i = 1 To allowedModules.Count
            If InStr(1, qwe.Name, allowedModules(i)) > 0 Then
                Set VBComp = ThisWorkbook.VBProject.VBComponents(qwe.Name)
                Select Case VBComp.Type
                    Case vbext_ct_ClassModule
                        Extension = ".cls"
                    Case vbext_ct_Document
                        Extension = ".cls"
                    Case vbext_ct_MSForm
                        Extension = ".frm"
                    Case vbext_ct_StdModule
                        Extension = ".bas"
                    Case Else
                        Extension = ".bas"
                End Select
                FileName = VBComp.Name
                ExportFile = Directory & FileName & Extension
                VBComp.Export ExportFile
            End If
        Next i
    Next
End Sub
