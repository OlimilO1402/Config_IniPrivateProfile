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
Private mIniFile  As ConfigIniDocument
Private mSection  As ConfigIniSection
Private mKeyName  As String
Private mHasValue As Boolean
Private mValue    As String
Private mDefault  As Variant

Private Sub Class_Initialize()
    mHasValue = True 'songma hot vonhausaus schoamoi an valju
End Sub
Private Sub Class_Terminate()
    '
End Sub

Friend Sub New_(aIniFile As ConfigIniDocument, _
                aSection As ConfigIniSection, _
                aKeyName As String, _
                Optional VarDefault As Variant)
    Set mIniFile = aIniFile: Set mSection = aSection
    mKeyName = aKeyName: mDefault = VarDefault
End Sub
Public Sub GetValue()
    mValue = ValueStr
End Sub

Public Sub NewS(aIniFile As ConfigIniDocument, _
                aSection As String, _
                aKeyName As String, _
                Optional VarDefault As Variant)
    Set mIniFile = aIniFile: Set mSection = MNew.ConfigIniSection(aIniFile, aSection)
    mKeyName = aKeyName: mDefault = VarDefault: mValue = ValueStr
End Sub
Public Property Set IniFile(aIni As ConfigIniDocument)
    Set mIniFile = aIni
End Property
Public Property Get IniFile() As ConfigIniDocument
    Set IniFile = mIniFile
End Property
Public Property Set section(asec As ConfigIniSection)
    Set mSection = asec
End Property
Public Property Get section() As ConfigIniSection
    Set section = mSection
End Property
Public Property Let Name(StrVal As String)
    mKeyName = StrVal
End Property
Public Property Get Name() As String
    Name = mKeyName
End Property
Public Property Let HasValue(BolVal As Boolean)
    mHasValue = BolVal
End Property
Public Property Get HasValue() As Boolean
    HasValue = mHasValue
End Property
Public Function Clone() As ConfigIniKeyValue
    Set Clone = MNew.ConfigIniKeyValue(mIniFile, mSection, mKeyName, mDefault)
    Clone.Value = mValue
End Function
Public Property Let Value(StrVal As String)
    If StrComp(mValue, StrVal, vbTextCompare) <> 0 Then
        mIniFile.IsDirty = True
    End If
    mValue = StrVal
    'sofort schreiben?
    'ValueStr = mValue
End Property
Public Property Get Value() As String
  'If Len(mValue) = 0 Then
  '  Value = CStr(mDefault)
  'Else
    Value = mValue
  'End If
End Property
Public Property Set Default(VarVal As Variant)
    Set mDefault = VarVal
End Property
Public Property Let Default(VarVal As Variant)
    mDefault = VarVal
End Property
Public Property Get Default() As Variant
    Default = mDefault
End Property
Public Sub Save()
    If mHasValue Then ValueStr = mValue
End Sub
'##############################'      Values      '##############################'
'einzelne Key-Werte lesen/schreiben
'einen Boolean Lesen/Schreiben
Public Property Let ValueBol(BolVal As Boolean)
    mIniFile.ValueBol(mSection.Name, mKeyName, CBool(mDefault)) = BolVal
End Property
Public Property Get ValueBol() As Boolean
    ValueBol = mIniFile.ValueBol(mSection.Name, mKeyName, CBool(mDefault))
End Property

'einen Integer bzw. Long Lesen/Schreiben als String
Public Property Let ValueInt(LngVal As Long)
    If Not IsNumeric(mDefault) Then mDefault = 0
    mIniFile.ValueInt(mSection.Name, mKeyName, CLng(mDefault)) = LngVal
End Property
Public Property Get ValueInt() As Long
    If Not IsNumeric(mDefault) Then mDefault = 0
    ValueInt = mIniFile.ValueInt(mSection.Name, mKeyName, CLng(mDefault))
End Property

'einen String Lesen/Schreiben
Public Property Let ValueStr(StrVal As String)
    mIniFile.ValueStr(mSection.Name, mKeyName, CStr(mDefault)) = StrVal
End Property
Public Property Get ValueStr() As String
    ValueStr = mIniFile.ValueStr(mSection.Name, mKeyName, CStr(mDefault))
End Property

'einen UDType Lesen/Schreiben
'der UDType muß jedesmal mit VarPtr(UdTypeVar) übergeben werden,
'gesetzt und zurückgegeben wird nur ein RetVal
Public Property Let ValueStructP(StructLen As Long, pStruct As Long, rv As Long)
    mIniFile.ValueStructP(mSection.Name, mKeyName, StructLen, pStruct) = rv
End Property
Public Property Get ValueStructP(StructLen As Long, pStruct As Long) As Long
    ValueStructP = mIniFile.ValueStructP(mSection.Name, mKeyName, StructLen, pStruct)
End Property
'
'##############################'  ToStr  '##############################'
Public Function ToStr() As String
    ToStr = mKeyName
    If mHasValue Then ToStr = ToStr & "=" & mValue
End Function

