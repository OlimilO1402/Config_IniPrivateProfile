VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfigIniDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit 'Z:468(11.06.2008)
'##############################'   PrivateProfile-API   '##############################'
#If VBA7 Then
    
#Else
    'Ini Unicode reading
    Private Declare Function GetPrivateProfileIntW Lib "kernel32" (ByVal pSection As LongPtr, ByVal pKey As LongPtr, ByVal nDefault As Long, ByVal lpFileName As LongPtr) As Long
    Private Declare Function GetPrivateProfileStringW Lib "kernel32" (ByVal pSection As LongPtr, ByVal pKey As LongPtr, ByVal lpDefault As LongPtr, ByVal pRetBuff As LongPtr, ByVal nSize As Long, ByVal lpFileName As LongPtr) As Long
    Private Declare Function GetPrivateProfileStructW Lib "kernel32" (ByVal pSection As LongPtr, ByVal pKey As LongPtr, ByVal lpStruct As LongPtr, ByVal uSizeStruct As Long, ByVal lpFileName As LongPtr) As Long
    Private Declare Function GetPrivateProfileSectionW Lib "kernel32" (ByVal pSection As LongPtr, ByVal pRetBuff As LongPtr, ByVal nSize As Long, ByVal lpFileName As LongPtr) As Long
    Private Declare Function GetPrivateProfileSectionNamesW Lib "kernel32" (ByVal pRetBuff As LongPtr, ByVal nSize As Long, ByVal lpFileName As LongPtr) As Long
    'Ini Unicode writing
    Private Declare Function WritePrivateProfileStringW Lib "kernel32" (ByVal pSection As LongPtr, ByVal pKey As LongPtr, ByVal lpString As LongPtr, ByVal lpFileName As LongPtr) As Long
    Private Declare Function WritePrivateProfileStructW Lib "kernel32" (ByVal pSection As LongPtr, ByVal pKey As LongPtr, ByVal lpStruct As LongPtr, ByVal uSizeStruct As Long, ByVal lpFileName As LongPtr) As Long
    Private Declare Function WritePrivateProfileSectionW Lib "kernel32" (ByVal pSection As LongPtr, ByVal lpString As LongPtr, ByVal lpFileName As LongPtr) As Long
    
    'Ini Ansi reading
    Private Declare Function GetPrivateProfileIntA Lib "kernel32" (ByVal pSection As String, ByVal pKey As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
    Private Declare Function GetPrivateProfileStringA Lib "kernel32" (ByVal pSection As String, ByVal pKey As String, ByVal lpDefault As String, ByVal pRetBuff As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Private Declare Function GetPrivateProfileStructA Lib "kernel32" (ByVal pSection As String, ByVal pKey As String, ByVal lpStruct As LongPtr, ByVal uSizeStruct As Long, ByVal lpFileName As String) As Long
    Private Declare Function GetPrivateProfileSectionA Lib "kernel32" (ByVal pSection As String, ByVal pRetBuff As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Private Declare Function GetPrivateProfileSectionNamesA Lib "kernel32" (ByVal pRetBuff As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    'Ini Ansi writing
    Private Declare Function WritePrivateProfileStringA Lib "kernel32" (ByVal pSection As String, ByVal pKey As String, ByVal lpString As String, ByVal lpFileName As String) As Long
    Private Declare Function WritePrivateProfileStructA Lib "kernel32" (ByVal pSection As String, ByVal pKey As String, ByVal lpStruct As LongPtr, ByVal uSizeStruct As Long, ByVal lpFileName As String) As Long
    Private Declare Function WritePrivateProfileSectionA Lib "kernel32" (ByVal pSection As String, ByVal lpString As String, ByVal lpFileName As String) As Long
#End If
'private members
Private m_PFN      As PathFileName
Private m_PFName   As String 'Short for all writing-functions
Private m_IsLoaded As Boolean
Private m_IsDirty  As Boolean
Private m_Unicode  As Boolean
Private m_Root     As ConfigIniSection 'non-hashed
Private m_Sections As List 'Of ConfigIniSection  'hashed
Private m_StrSave  As String 'Zwischenspeicher für den letzten Stand der IniDatei
Private m_TempKey  As ConfigIniKeyValue

Private Sub Class_Initialize()
    'Set m_Root = MNew.ConfigIniSection(Me, vbNullString)
    Set m_Sections = MNew.List(vbObject, , True)
End Sub

Friend Sub New_(aPFN As PathFileName, Optional ByVal bUnicode As Boolean = False)
    Set m_PFN = aPFN: m_PFName = m_PFN.Value: m_Unicode = bUnicode
End Sub
'Copy to MNew:
'Public Function ConfigIniDocument(aPFN As PathFileName, Optional ByVal bUnicode As Boolean = False) As ConfigIniDocument
'    Set ConfigIniDocument = New ConfigIniDocument: ConfigIniDocument.New_ aPFN, bUnicode
'End Function

Friend Sub NewC(other As ConfigIniDocument)
    With other
        Me.FileName = .FileName: m_IsLoaded = .IsLoaded
        m_IsDirty = .IsDirty:    Set m_Root = .Root.Clone
        Set m_Sections = .Sections.Clone
    End With
End Sub

Friend Function Clone() As ConfigIniDocument
    Set Clone = New ConfigIniDocument: Clone.NewC Me
End Function

Public Sub Clear()
    Set m_Sections = MNew.List(vbObject, , True)
    m_IsLoaded = False 'set this here
End Sub

Public Property Get Root() As ConfigIniSection
    Set Root = m_Root
End Property

Public Function Contains(Section As String) As Boolean
    Contains = m_Sections.ContainsKey(Section)
End Function

Friend Property Let FileName(RHS As String)
    m_PFName = RHS
    m_PFN.Value = m_PFName
End Property
Public Property Get FileName() As String
    FileName = m_PFName
End Property

Public Property Get PFN() As PathFileName
    Set PFN = m_PFN
End Property

Public Property Get IsLoaded() As Boolean
    IsLoaded = m_IsLoaded
End Property

Friend Property Let IsDirty(BolVal As Boolean)
    m_IsDirty = BolVal
End Property
Public Property Get IsDirty() As Boolean
    IsDirty = m_IsDirty
End Property

Friend Property Let IsUnicode(ByVal Value As Boolean)
    m_Unicode = Value
End Property
Public Property Get IsUnicode() As Boolean
    IsUnicode = m_Unicode
End Property

Public Sub Load()
    'first load the root-section, it is the 0-section, it has no name, this section is not like a normal section
    Dim line As String
    line = Trim(m_PFN.ReadLine)
    If Left(line, 1) = "[" Then
        m_PFN.CloseFile
        LoadNormalIni
        Exit Sub
    End If
    Dim sa() As String
    sa = Split(line, "=")
    Dim cikv As ConfigIniKeyValue
    Set m_Root = MNew.ConfigIniSection(Me, vbNullString)
    Set cikv = m_Root.KeyValues.Add(MNew.ConfigIniKeyValue(Me, m_Root, sa(0)))
    cikv.Value = sa(1)
    Do While Not m_PFN.IsEOF
        line = Trim(m_PFN.ReadLine)
        If Len(line) Then
            If Left(line, 1) = "[" Then
                m_PFN.CloseFile
                LoadNormalIni
                Exit Sub
            Else
                sa = Split(line, "=")
                Set cikv = m_Root.KeyValues.Add(MNew.ConfigIniKeyValue(Me, m_Root, sa(0)))
                cikv.Value = sa(1)
            End If
        End If
    Loop
End Sub

Private Sub LoadNormalIni()
    Dim nSec As ConfigIniSection
    Dim StrSecArr() As String
    Dim StrSecName As Variant
Try: On Error GoTo Catch
    GetIniArr StrSecArr
    Call Clear
    Dim i As Long, u As Long: u = UBound(StrSecArr)
    For i = 0 To u
        'sSecName = StrSecArr(i)
        Set nSec = MNew.ConfigIniSection(Me, StrSecArr(i))
        Call nSec.Load
        Call m_Sections.Add(nSec, StrSecArr(i))
    Next
    m_IsLoaded = True
    Exit Sub
Catch:
    Call ErrHandler("Load")
End Sub

Public Sub Save()
Try: On Error GoTo Catch
    Dim asec As ConfigIniSection
    Dim aKey As ConfigIniKeyValue
    Dim i As Long, j As Long
    If m_IsDirty Then
        If Not m_Root.KeyValues.IsEmpty Then
            m_PFN.Delete
            For i = 0 To m_Root.KeyValues.Count - 1
                Set aKey = m_Root.KeyValues.Item(i)
                m_PFN.WriteLine aKey.ToStr
            Next
            m_PFN.CloseFile
        End If
        'For Each asec In m_Sections
        For i = 0 To m_Sections.Count - 1
            Set asec = m_Sections.Item(i)
            'For Each aKey In asec.KeyValues
            For j = 0 To asec.KeyValues.Count - 1
                Set aKey = asec.KeyValues.Item(j)
                aKey.Save
            Next
        Next
    End If
    GoTo Finally
Catch:
    ErrHandler "Save"
Finally:
    m_PFN.CloseFile
End Sub

'##############################'  KeyValues  '##############################'
'Returns an unknown ConfigIniKeyValue-Objekt:
'For direct writing of ini-entries use either all Value-X-functios below, or this
'Key-property, with the corresponding functions in the ConfigIniKeyValue-class
Public Property Get Key(ByVal aSection As String, ByVal aKey As String) As ConfigIniKeyValue
    If m_TempKey Is Nothing Then
        Set m_TempKey = MNew.ConfigIniKeyValueS(Me, aSection, aKey, vbNullString)
    End If
    If Not (m_TempKey.Section.Name = aSection) Then
        m_TempKey.Section.Name = aSection
    End If
    m_TempKey.Name = aKey
    Set Key = m_TempKey
End Property

'The following functions are for reading/writing single Key-values directly to the ini file
'Read/write a Boolean
Public Property Let ValueBol(ByVal aSection As String, ByVal aKey As String, Default As Boolean, BolVal As Boolean)
Try: On Error GoTo Catch
    ValueStr(aSection, aKey, CStr(Default)) = BolToStr(BolVal)
    Exit Property
Catch:
    Call ErrHandler("Let_ValueBol", aSection, aKey, CStr(Default))
End Property
Public Property Get ValueBol(ByVal aSection As String, ByVal aKey As String, Default As Boolean) As Boolean
Try: On Error GoTo Catch
    ValueBol = StrToBol(ValueStr(aSection, aKey, CStr(Default)))
    Exit Property
Catch:
    Call ErrHandler("Get_ValueBol", aSection, aKey, CStr(Default))
End Property

'Private Function BolToStr(ByVal b As Boolean) As String
'    If b Then BolToStr = "True" Else BolToStr = "False"
'End Function

'Read/write an Integer/Long
Public Property Let ValueInt(ByVal aSection As String, ByVal aKey As String, ByVal Default As Long, LngVal As Long)
Try: On Error GoTo Catch
    ValueStr(aSection, aKey, CStr(Default)) = CStr(LngVal)
    Exit Property
Catch:
    Call ErrHandler("Let_ValueInt", aSection, aKey, CStr(Default))
End Property
Public Property Get ValueInt(ByVal aSection As String, ByVal aKey As String, ByVal Default As Long) As Long
Try: On Error GoTo Catch
    ValueInt = GetIniIntA(aSection, aKey, Default, m_PFName)
    Exit Property
Catch:
    Call ErrHandler("Get_ValueInt", aSection, aKey, CStr(Default))
End Property

'Read/write a String
Public Property Let ValueStr(ByVal aSection As String, ByVal aKey As String, ByVal Default As String, StrVal As String)
Try: On Error GoTo Catch
    Dim rv As Long: rv = PutIniStr(aSection, aKey, StrVal, m_PFName)
    Exit Property
Catch:
    Call ErrHandler("Let ValueStr", aSection, aKey, Default, StrVal)
End Property
Public Property Get ValueStr(ByVal aSection As String, ByVal aKey As String, ByVal Default As String) As String
    Const ValLen As Long = 512: Dim StrVal As String * ValLen
Try: On Error GoTo Catch
    ValueStr = Left$(StrVal, GetIniStr(aSection, aKey, Default, StrVal, ValLen, m_PFName))
    Exit Property
Catch:
    Call ErrHandler("Get ValueStr", aSection, aKey, Default)
End Property

'Read/write an Array
'Pass the array ByRef, the return-value is the numer of items in the array
Public Property Let ValueStrArr(ByVal aSection As String, strArr() As String, RetVal As Long)
Try: On Error GoTo Catch
    Dim Buffer As String ', i As Long
    Buffer = Join(strArr, vbNullChar)
    'Hey, warum nicht join verwenden?
    'Dim L As Long: L = LBound(strArr)
    'Dim u As Long: u = UBound(strArr)
    'For i = L To u
    '    Buffer = Buffer & strArr(i)
    '    If i < u Then Buffer = Buffer & vbNullChar
    'Next
    'Zuerst die Section komplett löschen, denn falls sie vorher länger gewesen wäre würde sonst nur ein Teil überschrieben werden.
    Dim rv As Long
    rv = PutIniStr(aSection, vbNullString, vbNullString, m_PFName)
    rv = PutIniSection(aSection, Buffer, m_PFName)
    RetVal = u + 1
    Exit Property
Catch:
    Call ErrHandler("Let_ValueStrArr", aSection)
End Property
Public Property Get ValueStrArr(ByVal aSection As String, strArr() As String) As Long
Try: On Error GoTo Catch
    Dim StrBuff As String: Call GetSectionBuffer(aSection, StrBuff)
    Dim i As Long, n As Long
    Dim pos As Long: pos = 1
    Dim oPos As Long: oPos = pos
    If Len(StrBuff) > 0 Then
        'String mit Trennzeichen Chr$(0) in ein Feld umwandeln
        n = CountNullCharsInBuffer(StrBuff)
        ReDim strArr(0 To n - 1)
        For i = 0 To n - 1
            pos = InStr(oPos, StrBuff, vbNullChar)
            If pos > 0 Then
                strArr(i) = Mid$(StrBuff, oPos, pos - oPos)
            End If
            oPos = pos + 1
        Next
    End If
    ValueStrArr = n
    Exit Property
Catch:
    Call ErrHandler("Get_ValueStrArr", aSection)
End Property

'einen UDType Lesen/Schreiben
'der UDType muß bei Let/Get mit VarPtr(UdTypeVar) übergeben werden,
'gesetzt und zurückgegeben wird nur ein RetVal
'Bsp.:
'* eine UDT-Variable myudtval aus der ini lesen:
'    rv = mIni.ValueStructP(sn, kk, 16, VarPtr(myudtval))
'* eine UDT-Variable myudtval in die ini schreiben:
'    mIni.ValuestructP(sn, kk, 16, VarPtr(myudtval)) = rv
Public Property Let ValueStructP(ByVal aSection As String, ByVal aKey As String, ByVal StructLen As Long, ByVal pStruct As LongPtr, rv_out As Long)
Try: On Error GoTo Catch
    rv_out = PutIniStruct(aSection, aKey, ByVal pStruct, StructLen, m_PFName)
    Exit Property
Catch:
    Call ErrHandler("Let_ValueStructP", "Section: " & aSection & "; Key: " & aKey & "; " & CStr(StructLen) & "; ")
End Property
Public Property Get ValueStructP(ByVal aSection As String, ByVal aKey As String, ByVal StructLen As Long, ByVal pStruct As LongPtr) As Long
Try: On Error GoTo Catch
    ValueStructP = GetIniStruct(aSection, aKey, pStruct, StructLen, m_PFName)
    Exit Property
Catch:
    Call ErrHandler("Get_ValueStructP", aSection, aKey, CStr(StructLen))
End Property
'

'##############################'  Sections  '##############################'
'Public Property Set Sections(Value As List)
'    Set m_Sections = Value
'End Property
Public Property Get Sections() As List
    Set Sections = m_Sections
End Property

Public Property Get Section(IndexKey As Variant) As ConfigIniSection
Try: On Error GoTo Catch
    If IsNumeric(IndexKey) Then
        Set Section = m_Sections.Item(CLng(IndexKey))
    Else
        Set Section = m_Sections.ItemByKey(CStr(IndexKey))
    End If
    Exit Property
Catch:
    ErrHandler "Get_Section", CStr(IndexKey)
End Property

Public Function SectionExists(ByVal aSectionName As String) As Boolean
    SectionExists = Not GetSection(aSectionName) Is Nothing
End Function
Private Function GetSection(ByVal aSectionName As String) As ConfigIniSection
Try: On Error Resume Next
    Set GetSection = m_Sections.Item(aSectionName)
Catch: On Error GoTo 0
End Function
Public Function AddSection(ByVal aSectionName As String) As ConfigIniSection
    'wenn die Section bereits existiert, wird diese zurückgegeben
    If Me.Contains(aSectionName) Then
        Set AddSection = GetSection(aSectionName) 'ohne IIf!!
    Else
        Set AddSection = MNew.ConfigIniSection(Me, aSectionName)
        m_Sections.Add AddSection, aSectionName
    End If
End Function
Public Sub DeleteSection(ByVal aSectionName As String)
Try: On Error GoTo Catch
    PutIniStr aSectionName, vbNullString, vbNullString, m_PFName
    If SectionExists(aSectionName) Then
        Call m_Sections.Remove(aSectionName)
    End If
    Exit Sub
Catch:
    Call ErrHandler("DeleteSection", aSectionName)
End Sub
Public Sub DeleteKey(ByVal aSectionName As String, ByVal aKeyName As String)
Try: On Error GoTo Catch
    PutIniStr aSectionName, aKeyName, vbNullString, m_PFName
    Dim asec As ConfigIniSection
    If SectionExists(aSectionName) Then
        Set asec = m_Sections.Item(aSectionName)
        asec.DeleteKey aKeyName
    End If
    Exit Sub
Catch:
    ErrHandler "DeleteKey", aSectionName, aKeyName
End Sub
Public Property Get CountEntriesInSection(ByVal aSection As String) As Long
    Dim StrBuff As String:  Call GetSectionBuffer(aSection, StrBuff)
Try: On Error GoTo Catch
    CountEntriesInSection = CountNullCharsInBuffer(StrBuff)
    Exit Property
Catch:
    Call ErrHandler("CountEntriesInSection", aSection)
End Property
Public Property Get SectionNamesToCol() As Collection
    Set SectionNamesToCol = New Collection
    Dim StrBuff As String
    Dim oPos As Long, pos As Long: pos = 1
    Dim rv As Long, rv1 As Long:   rv1 = -1
Try: On Error GoTo Catch
    Dim BuffLen As Long: BuffLen = 256
    Dim sl As Long
    Do While (sl = rv) Or (rv = 0)
        StrBuff = String$(BuffLen, vbNullChar)
        rv = GetIniSectNames(StrBuff, BuffLen, m_PFName)
        sl = Len(StrBuff) - 2
        BuffLen = BuffLen + BuffLen
    Loop
    'zwei aufeinanderfolgende vbNullChar finden:
    Dim n2 As String: n2 = vbNullChar & vbNullChar
    pos = InStr(pos, StrBuff, n2)
    If pos > 0 Then 'sollte immer etwas finden
        StrBuff = Left$(StrBuff, pos - 1)
    End If
    Dim SNamArr() As String
    SNamArr = Split(StrBuff, vbNullChar)
    Dim i As Long
    Dim s As String
    For i = 0 To UBound(SNamArr)
        s = SNamArr(i)
        Call SectionNamesToCol.Add(s)
    Next
    Exit Function
Catch:
    Call ErrHandler("SectionNamesToCol")
End Property

Public Function SectionToCol(aSection As String) As Collection
'Die Funktion gibt in einer Collection alle Elemente der Section zurück (mit Keys)
    Set SectionToCol = New Collection
Try: On Error GoTo Catch
    Dim SNamArr() As String: Call GetIniArr(SNamArr, m_PFName, aSection)
    Dim i As Long
    For i = 0 To UBound(SNamArr)
        Call SectionToCol.Add(SNamArr(i))
    Next
    Exit Function
Catch:
    Call ErrHandler("SectionToCol")
End Function
Public Sub GetIniArr(aStrArr() As String, _
                     Optional ByVal aIniPFN As String, _
                     Optional ByVal aSection As String)
    Dim rv As Long, sl As Long, pos As Long
    Dim StrBuff As String
    Dim IniPFN As String
Try: On Error GoTo Catch
    If Len(aIniPFN) Then IniPFN = aIniPFN Else IniPFN = m_PFName
    Dim BuffLen As Long: BuffLen = 256 '1
    Do While (sl = rv) Or (rv = 0)
        StrBuff = String$(BuffLen, vbNullChar)
        If Len(aSection) Then
            rv = GetIniSection(aSection, StrBuff, BuffLen, IniPFN)
        Else
            rv = GetIniSectNames(StrBuff, BuffLen, IniPFN)
        End If
        sl = Len(StrBuff) - 2
        BuffLen = BuffLen + BuffLen
    Loop
    'zwei aufeinanderfolgende vbNullChar finden:
    pos = InStr(1, StrBuff, vbNullChar & vbNullChar)
    'sollte immer etwas finden
    '-1, weil ohne das letzte vbNullChar
    If pos > 0 Then StrBuff = Left$(StrBuff, pos - 1)
    aStrArr = Split(StrBuff, vbNullChar)
    Exit Sub
Catch:
    Call ErrHandler("GetIniArr")
End Sub

Private Sub GetSectionBuffer(ByVal aSection As String, StrBuff As String)
    Dim rv As Long, sl As Long, pos As Long: pos = 1
Try: On Error GoTo Catch
    Dim BuffLen As Long: BuffLen = 256 '1
    Do While (sl = rv) Or (rv = 0)
        StrBuff = String$(BuffLen, vbNullChar)
        rv = GetIniSection(aSection, StrBuff, BuffLen, m_PFName)
        sl = Len(StrBuff) - 2
        BuffLen = BuffLen + BuffLen
    Loop
    pos = InStr(pos, StrBuff, vbNullChar & vbNullChar)
    'sollte immer etwas finden
    '+1, damit immer zwei vbNullChar am Ende sind
    If pos > 0 Then StrBuff = Left$(StrBuff, pos + 1)
    Exit Sub
Catch:
    Call ErrHandler("GetSectionBuffer")
End Sub

'##############################'  ToStr  '##############################'
Public Function ToStr() As String
    Dim Section As ConfigIniSection
    For Each Section In m_Sections
        ToStr = ToStr & Section.ToStr & vbCrLf
    Next
End Function
  
'##############################'  Privates  '##############################'
'Private Function StrToBol(StrVal As String) As Boolean
'    If (StrComp(StrVal, "0", vbTextCompare) = 0) Or _
'       (StrComp(StrVal, "false", vbTextCompare) = 0) Or _
'       (StrComp(StrVal, "falsch", vbTextCompare) = 0) Or _
'       (StrComp(StrVal, "nein", vbTextCompare) = 0) Then
'        StrToBol = False
'    'ElseIf (StrComp(StrVal, vbNullString) = 0) Or _
'           (StrComp(StrVal, "1") = 0) Or _
'           (StrComp(StrVal, "-1") = 0) Or _
'           (StrComp(StrVal, "true") = 0) Or _
'           (StrComp(StrVal, "wahr") = 0) Or _
'           (StrComp(StrVal, "ja") = 0) Then
'    Else
'        StrToBol = True
'    End If
'End Function
'Private Function BolToStr(BolVal As Boolean) As String
'    If BolVal Then BolToStr = "True" Else BolToStr = "False"
'End Function
'
Private Function CountNullCharsInBuffer(aStrBuff As String) As Long
    Dim oPos As Long, pos As Long: pos = 1
    Dim n As Long
Try: On Error GoTo Catch
    Do
        oPos = pos
        pos = InStr(pos, aStrBuff, vbNullChar)
        If pos > oPos + 1 Then
            pos = pos + 1
            n = n + 1
        Else
            Exit Do
        End If
    Loop While pos > 0
    CountNullCharsInBuffer = n
    Exit Function
Catch:
    Call ErrHandler("CountNullCharsInBuffer")
End Function

Private Function GetIniInt(ByVal aSection As String, ByVal aKey As String, ByVal nDefault As Long, ByVal aFileName As String) As Long
    If m_Unicode Then
        GetIniInt = GetPrivateProfileIntW(StrPtr(aSection), StrPtr(aKey), nDefault, StrPtr(aFileName))
    Else
        GetIniInt = GetPrivateProfileIntA(aSection, aKey, nDefault, aFileName)
    End If
End Function

Private Function GetIniStr(ByVal aSection As String, ByVal aKey As String, ByVal Default As String, ByVal RetBuff As LongPtr, ByVal nSize As Long, ByVal aFileName As String) As Long
    If m_Unicode Then
        GetIniStr = GetPrivateProfileStringW(StrPtr(aSection), StrPtr(aKey), StrPtr(Default), RetBuff, nSize, StrPtr(aFileName))
    Else
        GetIniStr = GetPrivateProfileStringA(aSection, aKey, Default, RetBuff, nSize, aFileName)
    End If
End Function

Private Function GetIniStruct(ByVal aSection As String, ByVal aKey As String, ByVal lpStruct As LongPtr, ByVal uSizeStruct As Long, ByVal aFileName As String) As Long
    If m_Unicode Then
        GetIniStruct = GetPrivateProfileStructW(StrPtr(aSection), StrPtr(aKey), ByVal lpStruct, uSizeStruct, StrPtr(aFileName))
    Else
        GetIniStruct = GetPrivateProfileStructA(aSection, aKey, ByVal lpStruct, uSizeStruct, aFileName)
    End If
End Function

Private Function GetIniSection(ByVal aSection As String, ByVal aRetBuff As LongPtr, ByVal nSize As Long, ByVal aFileName As String) As Long
    If m_Unicode Then
        GetIniSection = GetPrivateProfileSectionW(StrPtr(aSection), aRetBuff, nSize, StrPtr(aFileName))
    Else
        GetIniSection = GetPrivateProfileSectionA(aSection, aRetBuff, nSize, aFileName)
    End If
End Function

Private Function GetIniSectNames(ByRef aRetBuff As String, ByVal nSize As Long, ByVal aFileName As String) As Long
    If m_Unicode Then
        GetIniSectNames = GetPrivateProfileSectionNamesW(StrPtr(aRetBuff), nSize, StrPtr(aFileName))
    Else
        GetIniSectNames = GetPrivateProfileSectionNamesA(aRetBuff, nSize, aFileName)
    End If
End Function

Private Function PutIniStr(ByVal aSection As String, ByVal aKey As String, ByVal lpString As String, ByVal aFileName As String) As Long
    If m_Unicode Then
        PutIniStr = WritePrivateProfileStringW(StrPtr(aSection), StrPtr(aKey), StrPtr(lpString), StrPtr(aFileName))
    Else
        PutIniStr = WritePrivateProfileStringA(aSection, aKey, lpString, aFileName)
    End If
End Function
Private Function PutIniStruct(ByVal aSection As String, ByVal aKey As String, ByVal lpStruct As LongPtr, ByVal uSizeStruct As Long, ByVal aFileName As String) As Long
    If m_Unicode Then
        PutIniStruct = WritePrivateProfileStructW(StrPtr(aSection), StrPtr(aKey), lpStruct, uSizeStruct, StrPtr(aFileName))
    Else
        PutIniStruct = WritePrivateProfileStructA(aSection, aKey, lpStruct, uSizeStruct, aFileName)
    End If
End Function
Private Function PutIniSection(ByVal aSection As String, ByVal aBuffer As String, ByVal lpFileName As String) As Long
    If m_Unicode Then
        PutIniSection = WritePrivateProfileSectionW(StrPtr(aRetBuff), StrPtr(aBuffer), StrPtr(lpFileName))
    Else
        PutIniSection = WritePrivateProfileSectionA(aRetBuff, aBuffer, lpFileName)
    End If
End Function

'####################' v     Local ErrHandler      v '####################'
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


