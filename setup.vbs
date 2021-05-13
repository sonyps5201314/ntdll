Option Explicit

Const adTypeBinary = 1
Const adSaveCreateOverWrite = 2

Function IIf(expr, truepart, falsepart)
    IIf = falsepart
    If expr Then IIf = truepart
End Function

'https://stackoverflow.com/questions/6405236/forcing-msxml-to-format-xml-output-with-indents-and-newlines
Private Sub FormatDocToFile(ByVal Doc, ByVal FileName)
    'Reformats the DOMDocument "Doc" into an ADODB.Stream
    'and writes it to the specified file.
    '
    'Note the UTF-8 output never gets a BOM.  If we want one we
    'have to write it here explicitly after opening the Stream.
    Dim rdrDom 'As MSXML2.SAXXMLReader
    Dim stmFormatted 'As ADODB.Stream
    Dim wtrFormatted 'As MSXML2.MXXMLWriter

    Set stmFormatted = CreateObject("ADODB.Stream")
    With stmFormatted
        .Open
        .Type = adTypeBinary
        Set wtrFormatted = CreateObject("MSXML2.MXXMLWriter")
        With wtrFormatted
            .omitXMLDeclaration = False
            .standalone = True
            .byteOrderMark = False 'If not set (even to False) then
                                   '.encoding is ignored.
            .encoding = "utf-8"    'Even if .byteOrderMark = True
                                   'UTF-8 never gets a BOM.
            .indent = True
            .output = stmFormatted
            Set rdrDom = CreateObject("MSXML2.SAXXMLReader")
            With rdrDom
                Set .contentHandler = wtrFormatted
                Set .dtdHandler = wtrFormatted
                Set .errorHandler = wtrFormatted
                .putProperty "http://xml.org/sax/properties/lexical-handler", wtrFormatted
                .putProperty "http://xml.org/sax/properties/declaration-handler", wtrFormatted
                .parse Doc
            End With
        End With
        .SaveToFile FileName, adSaveCreateOverWrite
        .Close
    End With
End Sub

'本脚本修改自https://github.com/KindDragon/vld/blob/master/setup/vld-setup.iss
Private Function VarIsNull(V)
    VarIsNull = IsEmpty(V)
End Function

Public Function FileExists(File)
    Dim fs 'As FileSystemObject
    Set fs = CreateObject("Scripting.FileSystemObject")
    FileExists = fs.FileExists(File)
End Function
Function DirExists(Folder)
    Dim fs 'As FileSystemObject
    Set fs = CreateObject("Scripting.FileSystemObject")
    DirExists = fs.FolderExists(Folder)
End Function

'本脚本修改自https://www.sqlservercentral.com/articles/creating-folders-using-vb-and-recursion
Function CreateDir(FolderName)
    Dim fs 'As Scripting.FileSystemObject
    Dim iBreak
    On Error Resume Next
    
    'search from right to find path
    iBreak = InStrRev(FolderName, "\")
    If iBreak > 0 Then
        Call CreateDir(Left(FolderName, iBreak - 1))
    End If
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FolderExists(FolderName) = False Then
        CreateDir = VarIsNull(fs.CreateFolder(FolderName))
    Else
        CreateDir = True
    End If
End Function

Function CreateDefaultUserProps(FileName, XmlText)
    Dim fs 'As FileSystemObject
    Set fs = CreateObject("Scripting.FileSystemObject")
    Dim ts 'As TextStream
    Set ts = fs.CreateTextFile(FileName)
    ts.Write(XmlText)
    ts.Close
End Function

Function Pos(SubStr, S)
   Pos = InStr(S, SubStr)
End Function

Sub StringChangeEx(S, FromStr, ToStr, SupportDBCS)
    S = Replace(S, FromStr, ToStr, 1, -1, IIf(SupportDBCS, vbTextCompare, vbBinaryCompare))
End Sub

Function ExpandConstant(S)
    Dim scriptdir
    scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
    'scriptdir = "F:\MyCppProjects\ntdll"
    ExpandConstant = Replace(S, "{app}", scriptdir)
End Function

Function EncodeString(S)
    Dim result
    result = S
    Call StringChangeEx(result, "(", "%28", True)
    Call StringChangeEx(result, ")", "%29", True)
    EncodeString = result
End Function

Sub UpdateString(dirList, Path, suffix)
    If (Len(dirList) = 0) Then
      dirList = Path + suffix
    Else
      If ((Pos(Path, dirList) = 0) And (Pos(EncodeString(Path), dirList) = 0)) Then
        dirList = Path + dirList
      End If
    End If
    'Debug.Print (dirList)
End Sub

Private Sub ModifyProps(FileName, includefolder, dlibfolder, slibfolder)
  Dim XMLDocument
  Dim XMLParent, IdgNode, XMLNode, XMLNodes
  Dim IncludeDirectoriesNode
  Dim AdditionalIncludeDirectories
  Dim DynamicLibraryDirectoriesNode
  Dim AdditionalDynamicLibraryDirectories
  Dim StaticLibraryDirectoriesNode
  Dim AdditionalStaticLibraryDirectories
  If (False = FileExists(FileName)) Then
    Exit Sub
  End If
    Set XMLDocument = CreateObject("Msxml2.DOMDocument.3.0")
    XMLDocument.async = False
    XMLDocument.Load (FileName)
    If (XMLDocument.parseError.errorCode = 0) Then
      Call XMLDocument.setProperty("SelectionLanguage", "XPath")
      Call XMLDocument.setProperty("SelectionNamespaces", "xmlns:b='http://schemas.microsoft.com/developer/msbuild/2003'")
      Set XMLNodes = XMLDocument.selectNodes("//b:Project")
      If (XMLNodes.length = 0) Then
        Exit Sub
      End If
      Set IdgNode = XMLNodes.Item(0)
      Set XMLNodes = IdgNode.selectNodes("//b:PropertyGroup")
      Dim PropertyGroupNode
      If (XMLNodes.length > 0) Then
        For Each XMLNode In XMLNodes
            If XMLNode.Attributes.length = 0 Then
                Set PropertyGroupNode = XMLNode
                Exit For
            End If
        Next
      End If
      If VarIsNull(PropertyGroupNode) Then
        Set XMLNode = XMLDocument.createNode(1, "PropertyGroup", "http://schemas.microsoft.com/developer/msbuild/2003")
        Set PropertyGroupNode = IdgNode.appendChild(XMLNode)
      End If
      
      If Len(includefolder) > 0 Then
        Set XMLNodes = PropertyGroupNode.selectNodes("//b:IncludePath")
        If (XMLNodes.length > 0) Then
           For Each XMLNode In XMLNodes
              If XMLNode.Attributes.length = 0 Then
                Set XMLParent = XMLNode
                Exit For
              End If
           Next
        End If
        If VarIsNull(XMLParent) Then
          Set XMLNode = XMLDocument.createNode(1, "IncludePath", "http://schemas.microsoft.com/developer/msbuild/2003")
          Set XMLParent = PropertyGroupNode.appendChild(XMLNode)
        End If
        Set IncludeDirectoriesNode = XMLParent
      End If
      If Len(dlibfolder) > 0 Then
        Set XMLNodes = IdgNode.selectNodes("//b:Link")
        If (XMLNodes.length > 0) Then
          Set XMLParent = XMLNodes.Item(0)
        Else
          Set XMLNode = XMLDocument.createNode(1, "Link", "http://schemas.microsoft.com/developer/msbuild/2003")
          Set XMLParent = IdgNode.appendChild(XMLNode)
        End If
        Dim s_value
        s_value = "'$(PlatformToolset)' == 'v142' or '$(PlatformToolset)' == 'v141' or '$(PlatformToolset)' == 'v141_xp' or '$(PlatformToolset)' == 'v140' or '$(PlatformToolset)' == 'v140_xp'"
        Set XMLNodes = XMLParent.selectNodes("//b:Link/b:AdditionalLibraryDirectories[@Condition=""" + s_value + """]")
        If (XMLNodes.length > 0) Then
          Set DynamicLibraryDirectoriesNode = XMLNodes.Item(0)
        Else
          Set XMLNode = XMLDocument.createNode(1, "AdditionalLibraryDirectories", "http://schemas.microsoft.com/developer/msbuild/2003")
          Set DynamicLibraryDirectoriesNode = XMLParent.appendChild(XMLNode)
          Call DynamicLibraryDirectoriesNode.setAttribute("Condition", s_value)
        End If
     End If
     If Len(slibfolder) > 0 Then
        Set XMLNodes = IdgNode.selectNodes("//b:Lib")
        If (XMLNodes.length > 0) Then
          Set XMLParent = XMLNodes.Item(0)
        Else
          Set XMLNode = XMLDocument.createNode(1, "Lib", "http://schemas.microsoft.com/developer/msbuild/2003")
          Set XMLParent = IdgNode.appendChild(XMLNode)
        End If
        Set XMLNodes = XMLParent.selectNodes("//b:Lib/b:AdditionalLibraryDirectories[@Condition=""" + s_value + """]")
        If (XMLNodes.length > 0) Then
          Set StaticLibraryDirectoriesNode = XMLNodes.Item(0)
        Else
          Set XMLNode = XMLDocument.createNode(1, "AdditionalLibraryDirectories", "http://schemas.microsoft.com/developer/msbuild/2003")
          Set StaticLibraryDirectoriesNode = XMLParent.appendChild(XMLNode)
          Call StaticLibraryDirectoriesNode.setAttribute("Condition", s_value)
        End If
      End If
      
      If Len(includefolder) > 0 Then
        AdditionalIncludeDirectories = ""
        If (False = VarIsNull(IncludeDirectoriesNode)) Then
          AdditionalIncludeDirectories = IncludeDirectoriesNode.Text
        End If
      End If
      If Len(dlibfolder) > 0 Then
        AdditionalDynamicLibraryDirectories = ""
        If (False = VarIsNull(DynamicLibraryDirectoriesNode)) Then
          AdditionalDynamicLibraryDirectories = DynamicLibraryDirectoriesNode.Text
        End If
      End If
      If Len(slibfolder) > 0 Then
        AdditionalStaticLibraryDirectories = ""
        If (False = VarIsNull(StaticLibraryDirectoriesNode)) Then
          AdditionalStaticLibraryDirectories = StaticLibraryDirectoriesNode.Text
        End If
      End If
      
      If Len(includefolder) > 0 Then
        Call UpdateString(AdditionalIncludeDirectories, ExpandConstant(includefolder + ";"), "$(IncludePath)")
      End If
      If Len(dlibfolder) > 0 Then
        Call UpdateString(AdditionalDynamicLibraryDirectories, ExpandConstant(dlibfolder + ";"), "%(AdditionalLibraryDirectories)")
      End If
      If Len(slibfolder) > 0 Then
        Call UpdateString(AdditionalStaticLibraryDirectories, ExpandConstant(slibfolder + ";"), "%(AdditionalLibraryDirectories)")
      End If
      
      If Len(includefolder) > 0 Then
        IncludeDirectoriesNode.Text = AdditionalIncludeDirectories
      End If
      If Len(dlibfolder) > 0 Then
        DynamicLibraryDirectoriesNode.Text = AdditionalDynamicLibraryDirectories
      End If
      If Len(slibfolder) > 0 Then
        StaticLibraryDirectoriesNode.Text = AdditionalStaticLibraryDirectories
      End If
      'XMLDocument.save (FileName)
      FormatDocToFile XMLDocument, FileName
    End If
End Sub

Sub ModifyAllProps()
    Dim objShell
    Dim Path
    Set objShell = CreateObject("WScript.Shell")
    Path = objShell.Environment("Process").Item("LOCALAPPDATA") + "\Microsoft\MSBuild\v4.0\"
    If Not DirExists(Path) Then
        If CreateDir(Path) Then
            Dim XmlText
            XmlText = XmlText & "<?xml version=""1.0"" encoding=""utf-8""?> " & vbCrLf
            XmlText = XmlText & "<Project DefaultTargets=""Build"" ToolsVersion=""12.0"" xmlns=""http://schemas.microsoft.com/developer/msbuild/2003"">" & vbCrLf
            XmlText = XmlText & "  <ImportGroup Label=""PropertySheets"">" & vbCrLf
            XmlText = XmlText & "  </ImportGroup>" & vbCrLf
            XmlText = XmlText & "  <PropertyGroup Label=""UserMacros"" />" & vbCrLf
            XmlText = XmlText & "  <PropertyGroup />" & vbCrLf
            XmlText = XmlText & "  <ItemDefinitionGroup />" & vbCrLf
            XmlText = XmlText & "  <ItemGroup />" & vbCrLf
            XmlText = XmlText & "</Project>" & vbCrLf

            Call CreateDefaultUserProps(Path + "Microsoft.Cpp.Win32.user.props", XmlText)
            Call CreateDefaultUserProps(Path + "Microsoft.Cpp.x64.user.props", XmlText)
        End If
    End If
    If (DirExists(Path)) Then
        Call ModifyProps(Path + "Microsoft.Cpp.Win32.user.props", "{app}\include", vbNullString, vbNullString)
        Call ModifyProps(Path + "Microsoft.Cpp.x64.user.props", "{app}\include", vbNullString, vbNullString)
    End If
End Sub

ModifyAllProps