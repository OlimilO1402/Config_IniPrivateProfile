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
Private m_IniFile   As ConfigIniDocument
Private m_Section   As String
Private m_KeyValues As List 'Of ConfigIniKeyValue

Private Sub Class_Initialize()
    Set m_KeyValues = MNew.List(vbObject)
End Sub

Friend Sub New_(IniFile As ConfigIniDocument, ByVal SectionName As String)
    Set m_IniFile = IniFile
    m_Section = SectionName
End Sub
'Copy to MNew:
'Public Function ConfigIniSection(IniFile As ConfigIniDocument, SectionName As String) As ConfigIniSection
'    Set ConfigIniSection = New ConfigIniSection: ConfigIniSection.New_ IniFile, SectionName
'End Function

Friend Sub NewC(other As ConfigIniSection)
    With other
        Set m_IniFile = .IniFile
        m_Section = .Name
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

Public Property Let Name(aName As String)
    m_Section = aName
End Property
Public Property Get Name() As String
    Name = m_Section
End Property

Public Property Get IniFile() As ConfigIniDocument
    Set IniFile = m_IniFile
End Property

Public Function Filter(ByVal AllKeyName As String) As List
    Dim cikv As ConfigIniKeyValue
    Dim i As Long: Set Filter = New List
    For i = 0 To m_KeyValues.Count - 1
        Set cikv = m_KeyValues.Item(i)
        If cikv.Name = AllKeyName Then
            Filter.Add cikv
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
Public Property Get KeyValue(IndexKey As Variant) As ConfigIniKeyValue
    Set KeyValue = m_KeyValues.ItemByKey(CStr(IndexKey)) 'mKeyCol.Item(IndexKey)
    If KeyValue Is Nothing Then
        'in jedem Fall etwas zurückliefern, wenn nicht vorhanden dann anlegen
        'ist das praktikabel? we will see, hehe
        Set KeyValue = MNew.ConfigIniKeyValue(m_IniFile, Me, CStr(IndexKey), vbNullString)
        m_KeyValues.Add KeyValue ', CStr(IndexKey))
    End If
End Property

Public Function KeyValueExists(aKeyName As String) As Boolean
    KeyValueExists = Not GetKeyValue(aKeyName) Is Nothing
End Function
Private Function GetKeyValue(aKeyName As String) As ConfigIniKeyValue
Try: On Error Resume Next
    Set GetKeyValue = m_KeyValues.ItemByKey(aKeyName)
Catch: On Error GoTo 0
End Function

Public Function AddKeyValue(ByVal aKeyName As String, Optional ByVal VarDefault) As ConfigIniKeyValue
    'wenn der Schlüssel bereits existiert, wird dieser zurückgegeben
    Set AddKeyValue = GetKeyValue(aKeyName)
    If AddKeyValue Is Nothing Then
        Set AddKeyValue = MNew.ConfigIniKeyValue(m_IniFile, Me, aKeyName, VarDefault)
        m_KeyValues.Add AddKeyValue
    End If
End Function

Public Function DeleteKeyValue(aKeyName As String) As Boolean
    If KeyValueExists(aKeyName) Then
        m_KeyValues.Remove m_KeyValues.IndexOf(aKeyName)
        DeleteKeyValue = True
    End If
End Function

Public Sub Load()
    Dim nkey As ConfigIniKeyValue
    Dim StrKeyArr() As String
    Dim sKey As String, aKey As String
    Dim i As Long, epos As Long
    m_IniFile.GetIniArr StrKeyArr, vbNullString, m_Section
    Call Clear
    For i = 0 To UBound(StrKeyArr)
        sKey = StrKeyArr(i)
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
    CountEntries = m_IniFile.CountEntriesInSection(m_Section)
End Function

Public Function StrKeyValsToCol() As Collection
    Set StrKeyValsToCol = m_IniFile.SectionToCol(m_Section)
End Function

'##############################'  ToStr  '##############################'
Public Function NameToStr() As String
    NameToStr = "[" & m_Section & "]"
End Function

Public Function ToStr() As String
    Dim s As String: s = Me.NameToStr
    Dim nkey As ConfigIniKeyValue
    Dim i As Long
    For i = 0 To m_KeyValues.Count - 1
        Set nkey = m_KeyValues.Item(i)
        s = s & vbCrLf & nkey.ToStr
    Next
    ToStr = s
End Function

Private Function GetIniSection(ByVal aSection As String, ByRef aRetBuff As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    GetIniSection = GetPrivateProfileSectionW(StrPtr(aSection), StrPtr(aRetBuff), nSize, StrPtr(lpFileName))
End Function
