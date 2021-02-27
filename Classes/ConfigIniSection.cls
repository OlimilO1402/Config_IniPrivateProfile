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
Private Declare Function GetIniSectionA Lib "kernel32" Alias "GetPrivateProfileSectionA" ( _
  ByVal section As String, _
  ByVal RetBuff As String, _
  ByVal nSize As Long, _
  ByVal lpFileName As String) As Long
Private mIniFile As ConfigIniDocument
Private mSection As String
Private mKeyCol  As Collection 'stores all Keys as ConfigIniKeyValue

Private Sub Class_Initialize()
    Call Clear
End Sub

Friend Sub New_(aIniFile As ConfigIniDocument, StrSectionName As String)
    Set mIniFile = aIniFile
    mSection = StrSectionName
End Sub

Public Function Contains(ByRef key As String) As Boolean
  On Error Resume Next
  '"Extras->Optionen->Bei nicht verarbeiteten Fehlern"
    If IsEmpty(mKeyCol(key)) Then: 'DoNothing
    Contains = (Err.Number = 0)
  On Error GoTo 0
End Function

Public Property Let Name(aName As String)
    mSection = aName
End Property
Public Property Get Name() As String
    Name = mSection
End Property

Public Property Set Keys(aCol As Collection)
    Set mKeyCol = aCol
End Property
Public Property Get Keys() As Collection
    Set Keys = mKeyCol
End Property
Public Function CloneKeys() As Collection
    Set CloneKeys = New Collection
    'jetzt auch alle ConfigIniKeyValues Clonen
    Dim aKey As ConfigIniKeyValue
    For Each aKey In mKeyCol
        Call CloneKeys.Add(aKey.Clone)
    Next
End Function
Public Function Clone() As ConfigIniSection
    Set Clone = MNew.ConfigIniSection(mIniFile, mSection)
    Set Clone.Keys = CloneKeys
End Function
'Public Property Set Key(IndexKey As Variant, aKeyVal As ConfigIniKeyValue)
'  Set mKeyCol.Item(IndexKey) = aKeyVal
'  Blödsinn das geht doch so überhaupt nicht
'End Property
Public Property Get key(IndexKey As Variant) As ConfigIniKeyValue
    Set key = GetKey(CStr(IndexKey)) 'mKeyCol.Item(IndexKey)
    If key Is Nothing Then
        'in jedem Fall etwas zurückliefern, wenn nicht vorhanden dann anlegen
        'ist das praktikabel? we will see, hehe
        Set key = MNew.ConfigIniKeyValue(mIniFile, Me, CStr(IndexKey), vbNullString)
        Call mKeyCol.Add(key, CStr(IndexKey))
    End If
End Property

Public Function KeyExists(aKeyName As String) As Boolean
    KeyExists = Not GetKey(aKeyName) Is Nothing
End Function
Private Function GetKey(aKeyName As String) As ConfigIniKeyValue
Try: On Error Resume Next
    Set GetKey = mKeyCol.Item(aKeyName)
Catch: On Error GoTo 0
End Function
Public Function AddKey(ByVal aKeyName As String, _
                       Optional ByVal VarDefault) As ConfigIniKeyValue
    'wenn der Schlüssel bereits existiert, wird dieser zurückgegeben
    Set AddKey = GetKey(aKeyName)
    If AddKey Is Nothing Then
        Set AddKey = MNew.ConfigIniKeyValue(mIniFile, Me, aKeyName, VarDefault)
        Call mKeyCol.Add(AddKey, aKeyName)
    End If
End Function
Public Function DeleteKey(aKeyName As String) As Boolean
    If KeyExists(aKeyName) Then
        Call mKeyCol.Remove(aKeyName)
        DeleteKey = True
    End If
End Function
Public Sub Clear()
    Set mKeyCol = New Collection
End Sub
Public Sub Load()
    Dim nkey As ConfigIniKeyValue
    Dim StrKeyArr() As String
    Dim sKey As String, aKey As String
    Dim i As Long, epos As Long
    Call mIniFile.GetIniArr(StrKeyArr, vbNullString, mSection)
    Call Clear
    For i = 0 To UBound(StrKeyArr)
        sKey = StrKeyArr(i)
        If Len(sKey) > 0 Then
            epos = InStr(1, sKey, "=")
            If epos > 1 Then aKey = Left$(sKey, epos - 1)
            Set nkey = MNew.ConfigIniKeyValue(mIniFile, Me, aKey, vbNullString)
            Dim kk As String
            If epos > 1 Then
                'soll sich der key den Value selber nochmal holen?
                'aber das doch doof, wenn der Value doch schon daliegt
                'nkey.GetValue
                nkey.Value = Mid$(sKey, epos + 1)
                kk = aKey
            Else
                nkey.HasValue = False
                kk = CStr(mKeyCol.Count)
            End If
            Call mKeyCol.Add(nkey, kk)
        End If
    Next
End Sub

Public Function CountEntries() As Long
    CountEntries = mIniFile.CountEntriesInSection(mSection)
End Function

Public Function StrKeyValsToCol() As Collection
    Set StrKeyValsToCol = mIniFile.SectionToCol(mSection)
End Function

'##############################'  ToStr  '##############################'
Public Function NameToStr() As String
    NameToStr = "[" & mSection & "]"
End Function
Public Function ToStr() As String
    Dim nkey As ConfigIniKeyValue
    Dim i As Long
    ToStr = ToStr & NameToStr
    For i = 1 To mKeyCol.Count
        ToStr = ToStr & vbCrLf
        Set nkey = mKeyCol.Item(i)
        ToStr = ToStr & nkey.ToStr
    Next
End Function

