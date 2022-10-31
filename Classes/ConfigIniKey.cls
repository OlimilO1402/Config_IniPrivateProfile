VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfigIniKeyValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit 'Z:127(08.01.2008)
Private m_IniFile  As ConfigIniDocument
Private m_Section  As ConfigIniSection
Private m_KeyName  As String
Private m_HasValue As Boolean
Private m_Value    As String
Private m_Default  As Variant

Private Sub Class_Initialize()
    m_HasValue = True
End Sub

Friend Sub New_(aIniFile As ConfigIniDocument, _
                aSection As ConfigIniSection, _
                aKeyName As String, _
                Optional VarDefault As Variant)
    Set m_IniFile = aIniFile: Set m_Section = aSection
    m_KeyName = aKeyName: m_Default = VarDefault
End Sub

'Copy to MNew:
'Public Function ConfigIniKeyValue(aIniFile As ConfigIniDocument, _
'                                  aSection As ConfigIniSection, _
'                                  aKeyName As String, _
'                                  Optional VarDefault As Variant) As ConfigIniKeyValue
'    Set ConfigIniKeyValue = New ConfigIniKeyValue: ConfigIniKeyValue.New_ aIniFile, aSection, aKeyName, VarDefault
'End Function
'Copy to MNew:
'Public Function ConfigIniKeyValueS(aIniFile As ConfigIniDocument, _
'                                   aSection As String, _
'                                   aKeyName As String, _
'                                   Optional VarDefault As Variant) As ConfigIniKeyValue
'    Set ConfigIniKeyValueS = New ConfigIniKeyValue: ConfigIniKeyValueS.NewS aIniFile, aSection, aKeyName, VarDefault
'End Function

Friend Sub NewS(aIniFile As ConfigIniDocument, _
                aSection As String, _
                aKeyName As String, _
                Optional VarDefault As Variant)
    Set m_IniFile = aIniFile: Set m_Section = MNew.ConfigIniSection(aIniFile, aSection)
    m_KeyName = aKeyName: m_Default = VarDefault: m_Value = ValueStr
End Sub

Friend Sub NewC(other As ConfigIniKeyValue)
    With other
        Set m_IniFile = .IniFile: Set m_Section = .Section
            m_KeyName = .Name:       m_HasValue = .HasValue
              m_Value = .Value:       m_Default = .Default
    End With
End Sub

Friend Function Clone() As ConfigIniKeyValue
    Set Clone = New ConfigIniKeyValue: Clone.NewC Me
End Function

Public Sub GetValue()
    m_Value = ValueStr
End Sub

Public Property Set IniFile(aIni As ConfigIniDocument)
    Set m_IniFile = aIni
End Property
Public Property Get IniFile() As ConfigIniDocument
    Set IniFile = m_IniFile
End Property

Public Property Set Section(asec As ConfigIniSection)
    Set m_Section = asec
End Property
Public Property Get Section() As ConfigIniSection
    Set Section = m_Section
End Property

Public Property Let Name(ByVal aValue As String)
    m_KeyName = aValue
End Property
Public Property Get Name() As String
    Name = m_KeyName
End Property

Public Property Let HasValue(ByVal aValue As Boolean)
    m_HasValue = aValue
End Property
Public Property Get HasValue() As Boolean
    HasValue = m_HasValue
End Property

Public Property Let Value(StrVal As String)
    If StrComp(m_Value, StrVal, vbTextCompare) <> 0 Then
        m_IniFile.IsDirty = True
    End If
    m_Value = StrVal
    m_HasValue = True
End Property
Public Property Get Value() As String
    Value = m_Value
End Property

'Public Property Set Default(VarVal As Variant)
'    Set m_Default = VarVal
'End Property
Public Property Let Default(VarVal As Variant)
    m_Default = VarVal
End Property
Public Property Get Default() As Variant
    Default = m_Default
End Property

Public Sub Save()
    If m_HasValue Then ValueStr = m_Value
End Sub

'##############################'      Values      '##############################'
'einzelne Key-Werte lesen/schreiben
'einen Boolean Lesen/Schreiben
Public Property Let ValueBol(ByVal BolVal As Boolean)
    m_IniFile.ValueBol(m_Section.Name, m_KeyName, CBool(m_Default)) = BolVal
End Property
Public Property Get ValueBol() As Boolean
    ValueBol = m_IniFile.ValueBol(m_Section.Name, m_KeyName, CBool(m_Default))
End Property

'einen Integer bzw. Long Lesen/Schreiben als String
Public Property Let ValueInt(LngVal As Long)
    If Not IsNumeric(m_Default) Then m_Default = 0
    m_IniFile.ValueInt(m_Section.Name, m_KeyName, CLng(m_Default)) = LngVal
End Property
Public Property Get ValueInt() As Long
    If Not IsNumeric(m_Default) Then m_Default = 0
    ValueInt = m_IniFile.ValueInt(m_Section.Name, m_KeyName, CLng(m_Default))
End Property

'einen String Lesen/Schreiben
Public Property Let ValueStr(StrVal As String)
    m_IniFile.ValueStr(m_Section.Name, m_KeyName, CStr(m_Default)) = StrVal
End Property
Public Property Get ValueStr() As String
    ValueStr = m_IniFile.ValueStr(m_Section.Name, m_KeyName, CStr(m_Default))
End Property

'einen UDType Lesen/Schreiben
'der UDType muß jedesmal mit VarPtr(UdTypeVar) übergeben werden,
'gesetzt und zurückgegeben wird nur ein RetVal
Public Property Let ValueStructP(StructLen As Long, pStruct As Long, rv As Long)
    m_IniFile.ValueStructP(m_Section.Name, m_KeyName, StructLen, pStruct) = rv
End Property
Public Property Get ValueStructP(StructLen As Long, pStruct As Long) As Long
    ValueStructP = m_IniFile.ValueStructP(m_Section.Name, m_KeyName, StructLen, pStruct)
End Property
'
'##############################'  ToStr  '##############################'
Public Function ToStr() As String
    ToStr = m_KeyName
    If m_HasValue Then ToStr = ToStr & "=" & m_Value
End Function

