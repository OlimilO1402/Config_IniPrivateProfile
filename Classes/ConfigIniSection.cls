VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfigIniSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit 'Z:141(08.01.2008)
#If VBA7 Then
    Private Declare PtrSafe Function GetPrivateProfileSectionW Lib "kernel32" (ByVal pSection As LongPtr, ByVal pRetBuff As LongPtr, ByVal nSize As Long, ByVal lpFileName As LongPtr) As Long
#Else
    Private Declare Function GetPrivateProfileSectionW Lib "kernel32" (ByVal pSection As LongPtr, ByVal pRetBuff As LongPtr, ByVal nSize As Long, ByVal lpFileName As LongPtr) As Long
#End If
Private m_Name      As String
Private m_IniFile   As ConfigIniDocument
Private m_KeyValues As List 'Of ConfigIniKeyValue

Private Sub Class_Initialize()
    Set m_KeyValues = MNew.List(vbObject)
End Sub

Friend Sub New_(IniFile As ConfigIniDocument, ByVal SectionName As String)
    Set m_IniFile = IniFile: m_Name = SectionName
End Sub
'Copy to MNew:
'Public Function ConfigIniSection(IniFile As ConfigIniDocument, SectionName As String) As ConfigIniSection
'    Set ConfigIniSection = New ConfigIniSection: ConfigIniSection.New_ IniFile, SectionName
'End Function

Friend Sub NewC(other As ConfigIniSection)
    With other
        m_Name = .Name
        Set m_IniFile = .IniFile
        'make a deep copy
        Set m_KeyValues = .KeyValues.Clone
    End With
End Sub

Friend Function Clone() As ConfigIniSection
    Set Clone = New ConfigIniSection: Clone.NewC Me
End Function

Public Sub Clear()
    'clear only the keyvalue-list
    Class_Initialize
End Sub

Friend Property Set KeyValues(aKeyValueList As Collection)
    Set m_KeyValues = aKeyValueList
End Property
Public Property Get KeyValues() As List
    Set KeyValues = m_KeyValues
End Property

Public Function Contains(ByVal Key As String) As Boolean
    Contains = m_KeyValues.ContainsKey(Key)
End Function

Public Property Let Name(ByVal Value As String)
    m_Name = Value
End Property
Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Get IniFile() As ConfigIniDocument
    Set IniFile = m_IniFile
End Property

Public Function Filter(ByVal AllKeyName As String) As ConfigIniSection
    Set Filter = MNew.ConfigIniSection(m_IniFile, AllKeyName)
    Dim i As Long, cikv As ConfigIniKeyValue
    For i = 0 To m_KeyValues.Count - 1
        Set cikv = m_KeyValues.Item(i)
        If cikv.Name = AllKeyName Then
            Filter.KeyValues.Add cikv
        End If
    Next
End Function

Public Property Get KeyValue(IndexKey As Variant) As ConfigIniKeyValue
    Set KeyValue = m_KeyValues.ItemByKey(CStr(IndexKey))
    If KeyValue Is Nothing Then
        'return a KeyValue in every case, if it's not existing, then we create a new one
        Set KeyValue = MNew.ConfigIniKeyValue(m_IniFile, Me, CStr(IndexKey), vbNullString)
        m_KeyValues.Add KeyValue
    End If
End Property

Public Function KeyValueExists(KeyName As String) As Boolean
    KeyValueExists = Not GetKeyValue(KeyName) Is Nothing
End Function
Private Function GetKeyValue(KeyName As String) As ConfigIniKeyValue
Try: On Error Resume Next
    Set GetKeyValue = m_KeyValues.ItemByKey(KeyName)
Catch: On Error GoTo 0
End Function

Public Function AddKeyValue(ByVal KeyName As String, Optional ByVal VarDefault) As ConfigIniKeyValue
    'Returns an existing, or if there is none, a new ConfigIniKeyValue-object
    Set AddKeyValue = GetKeyValue(KeyName)
    If AddKeyValue Is Nothing Then
        Set AddKeyValue = MNew.ConfigIniKeyValue(m_IniFile, Me, KeyName, VarDefault)
        m_KeyValues.Add AddKeyValue
    End If
End Function

Public Function DeleteKeyValue(KeyName As String) As Boolean
    'Deletes the ConfigIniKeyValue-object with the name KeyName from this section
    If KeyValueExists(KeyName) Then
        Dim i As Long: i = m_KeyValues.IndexOf(KeyName)
        m_KeyValues.Remove i
        DeleteKeyValue = True
    End If
End Function

Public Sub Load()
    If Len(m_Name) Then LoadPrivProf Else LoadRoot
End Sub

Private Sub LoadPrivProf()
Try: On Error GoTo Catch
    Dim nkey As ConfigIniKeyValue
    Dim KeyNames() As String
    Dim sKey As String, aKey As String
    Dim i As Long, epos As Long
    m_IniFile.GetIniArr KeyNames, vbNullString, m_Name
    Call Clear
    For i = 0 To UBound(KeyNames)
        sKey = KeyNames(i)
        If Len(sKey) > 0 Then
            epos = InStr(1, sKey, "=")
            If epos > 1 Then aKey = Left$(sKey, epos - 1)
            Set nkey = MNew.ConfigIniKeyValue(m_IniFile, Me, aKey, vbNullString)
            Dim kk As String
            If epos > 1 Then
                'the key could get the value itself, but should it do so,
                'nope since the value is already there
                nkey.Value = Mid$(sKey, epos + 1)
                kk = aKey
            Else
                nkey.HasValue = False
                kk = CStr(m_KeyValues.Count)
            End If
            m_KeyValues.Add nkey ', kk
        End If
    Next
    GoTo Finally
Catch:
    ErrHandler "LoadAsRoot"
Finally:
End Sub

Private Sub LoadRoot()
Try: On Error GoTo Catch
    Dim pfn As PathFileName: Set pfn = m_IniFile.pfn
    Dim line As String
    Dim sKeyName As String, sValue As String, bEqualsign As Boolean, bLineIsSectionHeader As Boolean
    Do
        line = Trim(pfn.ReadLine)
        If Len(line) Then
            sKeyName = vbNullString: sValue = vbNullString: bEqualsign = False: bLineIsSectionHeader = False
            If Not Line_TryParse(line, sKeyName, bEqualsign, sValue, bLineIsSectionHeader) Then
                'there must be something wrong unexpectedly so exit if every case
                GoTo Finally
            Else
                'exit if a section-header-name is detected "["
                If bLineIsSectionHeader Then GoTo Finally
                'don't use AddKeyValue here, because there can be more keys with the same name
                Dim nkey As ConfigIniKeyValue: Set nkey = MNew.ConfigIniKeyValue(m_IniFile, Me, sKeyName)
                If Len(sValue) Then
                    nkey.Value = sValue
                Else
                    nkey.HasEqual = bEqualsign
                End If
                m_KeyValues.Add nkey
            End If
        End If
    Loop Until pfn.IsEOF
    GoTo Finally
Catch:
    ErrHandler "LoadAsRoot"
Finally:
    pfn.CloseFile
End Sub

Private Function Line_TryParse(line As String, keyname_out As String, hasEqualSign_out As Boolean, value_out As String, isSectionHeader_out As Boolean) As Boolean
Try: On Error GoTo Catch
    line = Trim(line)
    If Left(line, 1) = "[" Then
        isSectionHeader_out = True
    Else
        If InStr(1, line, "=") > 0 Then
            hasEqualSign_out = True
            Dim sa() As String: sa = Split(line, "=")
            Dim u As Long: u = UBound(sa)
            keyname_out = sa(0)
            If u > 0 Then value_out = sa(1)
        Else
            keyname_out = line
        End If
    End If
    Line_TryParse = True
    Exit Function
Catch:
    ErrHandler "Line_TryParse", m_Name, keyname_out, value_out, "isSectionHeader: " & isSectionHeader_out
End Function

Public Function CountEntries() As Long
    CountEntries = m_IniFile.CountEntriesInSection(m_Name)
End Function

Public Function StrKeyValsToCol() As Collection
    Set StrKeyValsToCol = m_IniFile.SectionToCol(m_Name)
End Function

'##############################'  ToStr  '##############################'
Public Function NameToStr() As String
    If Len(m_Name) Then NameToStr = "[" & m_Name & "]" '& vbCrLf
End Function

Public Function ToStr() As String
    Dim s As String: s = Me.NameToStr
    If Len(s) Then s = s & vbCrLf
    Dim KeyValue As ConfigIniKeyValue
    Dim i As Long
    For i = 0 To m_KeyValues.Count - 1
        Set KeyValue = m_KeyValues.Item(i)
        s = s & KeyValue.ToStr & vbCrLf
    Next
    ToStr = s
End Function

Private Function GetIniSection(ByVal aSection As String, ByRef aRetBuff As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    GetIniSection = GetPrivateProfileSectionW(StrPtr(aSection), StrPtr(aRetBuff), nSize, StrPtr(lpFileName))
End Function

' v #################### v '     Local ErrHandler      ' v #################### v '
Private Function ErrHandler(ByVal FuncName As String, _
                            Optional aSection As String, _
                            Optional aKey As String, _
                            Optional AddInfo1 As String, _
                            Optional AddInfo2 As String, _
                            Optional bLoud As Boolean = True, _
                            Optional decor As VbMsgBoxStyle = vbOKCancel) As VbMsgBoxResult
'So wenn ein Fehler hier in dieser Klasse auftritt,
'dann wird er hoffentlich hier ankommen
    Dim Info As String
    If Len(aSection) Then _
        Info = Info & "Section: " & aSection & vbCrLf
    If Len(aKey) Then _
        Info = Info & "Key:     " & aKey & vbCrLf
    If Len(AddInfo1) Then _
        Info = Info & "Info1:     " & AddInfo1 & vbCrLf
    If Len(AddInfo2) Then _
        Info = Info & "Info2:     " & AddInfo2 & vbCrLf
    
    If bLoud Then ErrHandler = ErrHandler2(FuncName, Info, , True, False, decor)

End Function


''copy this same function to every class, form or module
''the name of the class or form will be added automatically
''in standard-modules the function "TypeName(Me)" will not work, so simply replace it with the name of the Module
'' v ############################## v '   Local ErrHandler   ' v ############################## v '
Private Function ErrHandler2(ByVal FuncName As String, _
                             Optional ByVal AddInfo As String, _
                             Optional WinApiError, _
                             Optional bLoud As Boolean = True, _
                             Optional bErrLog As Boolean = True, _
                             Optional vbDecor As VbMsgBoxStyle = vbOKCancel, _
                             Optional bRetry As Boolean) As VbMsgBoxResult
    
    If bRetry Then
        
        ErrHandler2 = MessErrorRetry(TypeName(Me), FuncName, AddInfo, WinApiError, bErrLog)
        
    Else
        
        ErrHandler2 = MessError(TypeName(Me), FuncName, AddInfo, WinApiError, bLoud, bErrLog, vbDecor)
        
    End If
    
End Function

