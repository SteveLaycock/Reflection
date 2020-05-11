Attribute VB_Name = "HelperForVBIde"
Attribute VB_Description = "Interaction with the VBA IDE"
Option Explicit
'@Folder("VBASupport")
'@ModuleDescription("Interaction with the VBA IDE")

Public Const DefaultProject                     As String = "RegulatoryCMC"
Private Const HasPredeclaredId                  As String = "'@PredeclaredId"
Private Const HasMake                           As String = "Function Make"
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
' Code line limit should be 120 characters.
' Comment line limit should be 80 characters
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Private Declare PtrSafe Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer
    
Private Type State

    ProjectTree                           As Kvp
    
End Type

Private s                               As State

'@Description("Returns the name of the first project found which contains the definition on ipObject)
Public Function GetObjectProject(ByRef ipClassName As String) As String

    Dim myProject                       As VBIDE.VBProject
    Dim myComponent                     As VBIDE.VBComponent
    
    For Each myProject In Application.VBE.VBProjects
    
        For Each myComponent In myProject.VBComponents
        
            If myComponent.Name = ipClassName Then
            
                GetObjectProject = myProject.Name
                Exit Function
                
            End If
        
        Next
    
    Next
    
    ' It shouldn't be possible to get to here as we have to pass an
    ' existing object to get the project name
    Err.Raise _
        KvpEnumCommonMessages.AsEnum(ObjectHasNoDefinition), _
        ipClassName, _
        ipClassName & KvpEnumCommonMessages.AsEnum(ObjectHasNoDefinition)

End Function

'@Description("True if function exists in any project").
Public Function MethodExists _
       ( _
       ByVal ipMethodName As String, _
       Optional ByVal ipProcKind As ProcKindEnum = 0, _
       Optional ByVal ipProjectName As String = vbNullString, _
       Optional ByVal ipmodulename As String = vbNullString _
       ) As Boolean

    Dim myKey                              As String
    
    EnsureProjectTree
    myKey = Fmt("{0}:{1}:{2}:{3}", ipMethodName, KvpEnumProcKinds.Item(ipProcKind), ipProjectName, ipmodulename)
    MethodExists = s.ProjectTree.HoldsKey(myKey)

End Function

'@Description"Populates the ProjectTree is it is nothing")
Public Sub EnsureProjectTree()

    If s.ProjectTree Is Nothing Then

        Set s.ProjectTree = GetPopulatedMethodTree

    End If
    
End Sub

'@Description("Creates a list of the form MethodName:ProcKind:ModuleName:ProjectName")
Public Function GetPopulatedMethodTree() As Kvp
Attribute GetPopulatedMethodTree.VB_Description = "Creates a list of the form MethodName:ProcKind:ModuleName:ProjectName"
' Creates a list of the form MethodName:ProcKind:ModuleName:ProjectName vs index
' ProcKind, ModuleName and ProjectName are included on a combinatorial basis
' i.e. one or more of the entries may be a vbnull string
' see method AddToProjectTree

Const ONE As Long = 1

Dim myProject                           As VBIDE.VBProject
Dim myComponent                         As VBIDE.VBComponent
Dim myProjectTree                       As Kvp
Dim myTempArrayList                     As ArrayList

    Set myTempArrayList = New Kvp
    
    For Each myProject In Application.VBE.VBProjects
    
        For Each myComponent In myProject.VBComponents
        
            AddMethodsInCodeModule myProject, myComponent.CodeModule, myTempArrayList
            
        Next
        
    Next
    
    '@Ignore MemberNotOnInterface
    myTempArrayList.Sort
    Set myProjectTree = New Kvp
    '@Ignore MemberNotOnInterface
    myProjectTree.AddByIndexFromArray myTempArrayList.ToArray
    Set GetPopulatedMethodTree = myProjectTree.Mirror.GetItem(ONE)

End Function

'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C


'@Description("Adds method information to s.ProjectTree")
Private Sub AddMethodsInCodeModule _
        ( _
        ByRef ipProject As VBIDE.VBProject, _
        ByRef ipCodeModule As VBIDE.CodeModule, _
        ByRef iopProjectTreeArray As ArrayList _
        )
Attribute AddMethodsInCodeModule.VB_Description = "Adds method information to s.ProjectTree"

    Dim myLineInMethod                              As Long
    Dim myMethodName                                As String
    Dim myVBAProcKind                             As VBIDE.vbext_ProcKind
    Dim myUProcKind                               As String
    
    If IgnoreProject(ipProject.Name) Then Exit Sub
    If IgnoreCodeModule(ipCodeModule.Name) Then Exit Sub
    myLineInMethod = ipCodeModule.CountOfDeclarationLines + 1
    myVBAProcKind = 0
    
    Do Until myLineInMethod >= ipCodeModule.CountOfLines
        
        With ipCodeModule
        
            ' myProcKind is set to the correct kind as byref output value
            myMethodName = .ProcOfLine(myLineInMethod, myVBAProcKind)
            myUProcKind = KvpEnumProcKinds.Item(GetUnambigousProcKind(ipCodeModule, myMethodName, myVBAProcKind))
            'If InStr(myMethodName, "MethodExists") > 0 Then Stop
            AddToProjectTree iopProjectTreeArray, myMethodName, myUProcKind, .Name, ipProject.Name
            myLineInMethod = _
                           .ProcStartLine(myMethodName, myVBAProcKind) _
                         + .ProcCountLines(myMethodName, myVBAProcKind) _
                         + 1
                
        End With
        
    Loop
            
End Sub

Public Function IgnoreProject(ByVal ipProjectName As String) As Boolean

    Dim myReturn                                As Boolean

    'list of projects to ignore
    myReturn = InStr(ipProjectName, "Normal")
    'myreturn = myreturn or instr(ipprojectname, "Test")
    IgnoreProject = myReturn
  
End Function

Public Function IgnoreCodeModule(ByVal ipCodeModuleName As String) As Boolean

    Dim myReturn                                As Boolean
        
    ' List of Modules to ignore
    myReturn = _
             InStr(ipCodeModuleName, "Test") _
             Or InStr(ipCodeModuleName, "This") _
             Or InStr(ipCodeModuleName, "xx_") _
             Or InStr(ipCodeModuleName, "Module") _
             Or InStr(ipCodeModuleName, "Class") _
             Or InStr(ipCodeModuleName, "Template") _
             Or InStr(ipCodeModuleName, "ErrEx")
    IgnoreCodeModule = myReturn
    
End Function

Public Sub AddToProjectTree _
       ( _
       ByRef iopProjectTree As ArrayList, _
       ByVal ipMethodName As String, _
       ByVal ipProcKind As String, _
       ByVal ipmodulename As String, _
       ByVal ipProjectName As String)
    
    Dim myItem                              As String
    
    With iopProjectTree
        
        myItem = GetMethodKey(ipMethodName, vbNullString, vbNullString, vbNullString)
        If Not .Contains(myItem) Then .Add myItem
            
        myItem = GetMethodKey(ipMethodName, ipProcKind, vbNullString, vbNullString)
        If Not .Contains(myItem) Then .Add myItem

        myItem = GetMethodKey(ipMethodName, ipProcKind, ipmodulename, vbNullString)
        If Not .Contains(myItem) Then .Add myItem

        myItem = GetMethodKey(ipMethodName, ipProcKind, vbNullString, ipProjectName)
        If Not .Contains(myItem) Then .Add myItem

        myItem = GetMethodKey(ipMethodName, ipProcKind, ipmodulename, ipProjectName)
        If Not .Contains(myItem) Then .Add myItem

        myItem = GetMethodKey(ipMethodName, vbNullString, ipmodulename, vbNullString)
        If Not .Contains(myItem) Then .Add myItem

        myItem = GetMethodKey(ipMethodName, vbNullString, vbNullString, ipProjectName)
        If Not .Contains(myItem) Then .Add myItem
            
        myItem = GetMethodKey(ipMethodName, vbNullString, ipmodulename, ipProjectName)
        If Not .Contains(myItem) Then .Add myItem

    End With
    
End Sub

Public Function GetUnambigousProcKind _
       ( _
       ByRef ipCodeModule As VBIDE.CodeModule, _
       ByVal ipMethodName As String, _
       ByVal ipVBExtProcKind As vbext_ProcKind _
       ) As ProcKindEnum

    Dim myLineText                              As String

    If ipVBExtProcKind > 1 Then
        
        GetUnambigousProcKind = ipVBExtProcKind  '1,2, or 3
    
    Else
    
        myLineText = ipCodeModule.Lines(ipCodeModule.ProcBodyLine(ipMethodName, ipVBExtProcKind), 1)
        GetUnambigousProcKind = IIf(InStr(myLineText, "Function"), IsFunction, IsSub)
        
    End If
    
End Function

'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C

Public Sub ListMethods(ByRef ipProjectMethodTree As Kvp)

    Dim myMethod As Variant

    Debug.Print ipProjectMethodTree.GetItem(1&).Count
    For Each myMethod In ipProjectMethodTree.GetItem(1&)
    
        Debug.Print ipProjectMethodTree.GetItem(1&).GetKey(myMethod)
        
    Next
        
End Sub

Public Function GetMethodKey _
       ( _
       ByVal ipMethodName As String, _
       Optional ByVal ipProcKind As String = vbNullString, _
       Optional ByVal ipmodulename As String = vbNullString, _
       Optional ByVal ipProjectName As String = vbNullString _
       ) As String

    GetMethodKey = _
        Fmt _
        ( _
           "{0}:{1}:{2}:{3}", _
           ipMethodName, _
           ipProcKind, _
           ipmodulename, _
           ipProjectName _
        )
    
        
End Function

Public Function IsFactory(ByVal ipClassName As String, Optional ipProjectName As String = DefaultProject) As Boolean

    Dim myCodeModule As CodeModule
    Set myCodeModule = Application.VBE.VBProjects.Item(ipProjectName).VBComponents.Item(ipClassName).CodeModule
    
    IsFactory = myCodeModule.Find(HasPredeclaredId, 1, 1, -1, -1) And myCodeModule.Find(HasMake, 1, 1, -1, -1)
        
End Function

Public Function IsNotFactory(ByVal ipClassName As String, Optional ipProjectName As String = DefaultProject) As Boolean

    IsNotFactory = Not IsFactory(ipClassName, ipProjectName)
    
End Function


Public Function IsStatic(ByVal ipClassName As String, Optional ipProjectName As String = DefaultProject) As Boolean
    
    Dim myCodeModule                        As CodeModule
    Set myCodeModule = Application.VBE.VBProjects.Item(ipProjectName).VBComponents.Item(ipClassName).CodeModule
    IsStatic = myCodeModule.Find(HasPredeclaredId, 1, 1, -1, -1) And Not myCodeModule.Find(HasMake, 1, 1, -1, -1)
    
End Function

Public Function IsNotStatic(ByVal ipClassName As String, Optional ipProjectName As String = DefaultProject) As Boolean

    IsNotStatic = Not IsStatic(ipClassName, ipProjectName)
    
End Function


Public Function ControlKeyIsPressed() As Boolean

    Const VK_LCONTROL                               As Long = &HA2
    Const VK_RCONTROL                               As Long = &HA3
    'Const VK_LMENU                                  As Long = &HA4
    'Const VK_RMENU                                  As Long = &HA5

    Dim my_result                                   As Long

    my_result = False
    
    If _
        (GetKeyState(VK_LCONTROL) And &H8000) _
        Or (GetKeyState(VK_RCONTROL) And &H8000) _
        Then
    
        my_result = True
        
    End If
    
    ControlKeyIsPressed = my_result
    
End Function

Public Function EnumCount(ByVal ipEnumName As String) As Long

Dim myProject                           As VBIDE.VBProject
Dim myComponent                         As VBIDE.VBComponent
Dim mySearchText                        As String
Dim myStartLine                         As Long
Dim myEndLine                           As Long
Dim myCounter                           As Long
Dim myIndex                             As Long
Dim myLine                              As String

    mySearchText = "Enum " & ipEnumName
    
    For Each myProject In Application.VBE.VBProjects
    
        For Each myComponent In myProject.VBComponents
        
            With myComponent.CodeModule
                Debug.Print myComponent.CodeModule.Name
                'If myComponent.CodeModule = "KvpEnumAdminNames" Then Stop
                myStartLine = 1
                
                If myComponent.CodeModule.Find(mySearchText, myStartLine, 1, -1, -1) Then
                
                    myEndLine = myStartLine
                    myComponent.CodeModule.Find "End Enum", myEndLine, 1, -1, -1
                    myCounter = 0
                    
                    For myIndex = myStartLine + 1 To myEndLine - 1
                    
                        myLine = .Lines(myIndex, 1)
                        myLine = Replace(myLine, KvpEnumCharacters.Item(Space), vbNullString)
                        myLine = Replace(myLine, vbTab, vbNullString)
                        If Not ((Len(myLine) = 0) Or (Left$(myLine, 1) = "'")) Then myCounter = myCounter + 1
                        
                    Next
                     
                    EnumCount = myCounter
                    Exit Function
                End If
                
            End With
            
        Next
                
    Next

    EnumCount = 0

End Function

