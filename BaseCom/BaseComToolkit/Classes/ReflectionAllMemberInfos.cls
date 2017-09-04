VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReflectionAllMemberInfos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Benötig Verweis auf TypeLib Information

Option Explicit

Private mAllMemberInfos As Collection

Private Sub Class_Initialize()
  Set mAllMemberInfos = New Collection
End Sub

Private Sub Class_Terminate()
  Set mAllMemberInfos = Nothing
End Sub

Sub GetMembers(obj As Object)
    Const INVOKE_PROPERTYGET = 2

    Set mAllMemberInfos = New Collection
    Dim TLI       As Object: Set TLI = CreateObject("tli.tliApplication")
    Dim interface As Object: Set interface = TLI.InterfaceInfoFromObject(obj) 'interfaceInfo

    Dim member    As Object 'MemberInfo

    For Each member In interface.Members

        Dim mi As ReflectionMemberInfos: Set mi = New ReflectionMemberInfos

        'get property value; using memberID is faster than member.Name
        If member.InvokeKind = INVOKE_PROPERTYGET Then
            Dim result    As Variant: result = TLI.InvokeHook(obj, member.MemberID, INVOKE_PROPERTYGET)
    
            'adjust long values; convert to hex (else shows neg value)
            Dim valueType As String
    
            If TypeName(result) = "Long" Then
                result = "&H" & Hex(result) & "&"
                valueType = "long"
            Else
                valueType = TypeName(result)
            End If
    
            mi.MemberName = member.name
            mi.MemberID = Hex$(member.MemberID)
            mi.HelpString = member.HelpString
            mi.value = result
            mi.valueType = valueType
            mi.ParameterCount = member.parameters.count
            mi.MemberType = member.InvokeKind
    
            mAllMemberInfos.Add mi
        End If

    Next member

    Set TLI = Nothing

End Sub

Public Property Get AllMemberInfos() As Variant
  Set AllMemberInfos = mAllMemberInfos
End Property

