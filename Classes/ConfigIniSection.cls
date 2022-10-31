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
        Set m_KeyValues = .KeyValues.Clone
    End With
End Sub

Friend Function Clone() As ConfigIniSection
    Set Clone = New ConfigIniSection: Clone.NewC Me
End Function

Public Sub Clear()
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
'Public Function CloneKeys() As Collection
'    Set CloneKeys = New Collection
'    'jetzt auch alle ConfigIniKeyValues Clonen
'    Dim aKey As ConfigIniKeyValue
'    For Each aKey In mKeyCol
'        Call CloneKeys.Add(aKey.Clone)
'    Next
'End Function
'Public Property Set Key(IndexKey As Variant, aKeyVal As ConfigIniKeyValue)
'  Set mKeyCol.Item(IndexKey) = aKeyVal
'  Blödsinn das geht doch so überhaupt nicht
'End Property
Public Property Get keyvalue(IndexKey As Variant) As ConfigIniKeyValue
    Set keyvalue = m_KeyValues.ItemByKey(CStr(IndexKey)) 'mKeyCol.Item(IndexKey)
    If keyvalue Is Nothing Then
        'in jedem Fall etwas zurückliefern, wenn nicht vorhanden dann anlegen
        'ist das praktikabel? we will see, hehe
        Set keyvalue = MNew.ConfigIniKeyValue(m_IniFile, Me, CStr(IndexKey), vbNullString)
        m_KeyValues.Add keyvalue ', CStr(IndexKey))
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
                'soll sich der key den Value selber nochmal holen?
                'aber das doch doof, wenn der Value doch schon daliegt
                'nkey.GetValue
                nkey.Value = Mid$(sKey, epos + 1)
                kk = aKey
            Else
                nkey.HasValue = False
                kk = CStr(m_KeyValues.Count)
            End If
            m_KeyValues.Add nkey ', kk
        End If
    Next
End Sub

Public Function CountEntries() As Long
    CountEntries = m_IniFile.CountEntriesInSection(m_Name)
End Function

Public Function StrKeyValsToCol() As Collection
    Set StrKeyValsToCol = m_IniFile.SectionToCol(m_Name)
End Function

'##############################'  ToStr  '##############################'
Public Function NameToStr() As String
    NameToStr = "[" & m_Name & "]"
End Function

Public Function ToStr() As String
    Dim s As String: s = Me.NameToStr
    Dim keyvalue As ConfigIniKeyValue
    Dim i As Long
    For i = 0 To m_KeyValues.Count - 1
        Set keyvalue = m_KeyValues.Item(i)
        s = s & vbCrLf & keyvalue.ToStr
    Next
    ToStr = s
End Function

Private Function GetIniSection(ByVal aSection As String, ByRef aRetBuff As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    GetIniSection = GetPrivateProfileSectionW(StrPtr(aSection), StrPtr(aRetBuff), nSize, StrPtr(lpFileName))
End Function
