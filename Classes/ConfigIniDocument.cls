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
'Ini Lesen
Private Declare Function GetIniIntA Lib "kernel32" Alias "GetPrivateProfileIntA" ( _
  ByVal section As String, ByVal key As String, ByVal nDefault As Long, _
  ByVal lpFileName As String) As Long
Private Declare Function GetIniStrA Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
  ByVal section As String, ByVal key As String, ByVal lpDefault As String, _
  ByVal RetBuff As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetIniStructA Lib "kernel32" Alias "GetPrivateProfileStructA" ( _
  ByVal section As String, ByVal key As String, ByRef lpStruct As Any, _
  ByVal uSizeStruct As Long, ByVal lpFileName As String) As Long
Private Declare Function GetIniSectionA Lib "kernel32" Alias "GetPrivateProfileSectionA" ( _
  ByVal section As String, ByVal RetBuff As String, _
  ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetIniSectNamesA Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" ( _
  ByVal RetBuff As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'Ini Schreiben
Private Declare Function PutIniStrA Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
  ByVal section As String, ByVal key As String, ByVal lpString As Any, _
  ByVal lpFileName As String) As Long
Private Declare Function PutIniStructA Lib "kernel32" Alias "WritePrivateProfileStructA" ( _
  ByVal section As String, ByVal key As String, ByRef lpStruct As Any, _
  ByVal uSizeStruct As Long, ByVal lpFileName As String) As Long
Private Declare Function PutIniSectionA Lib "kernel32" Alias "WritePrivateProfileSectionA" ( _
  ByVal section As String, ByVal key As String, _
  ByVal lpFileName As String) As Long
  
'private members
'der Dateiname:
Private mPFNname    As String
Private mIsLoaded   As Boolean
Private mIsDirty    As Boolean
Private mSectionCol As Collection 'speichert alle Sections als ConfigIniSection
Private mStrSave    As String 'Zwischenspeicher für den letzten Stand der IniDatei
Private mTempKey    As ConfigIniKeyValue

'man sollte auch entscheiden können, ob eingelesen wurde, oder nicht
Private Sub Class_Initialize()
    Call Clear
End Sub
Private Sub Class_Terminate()
    Set mSectionCol = Nothing
End Sub

Friend Sub New_(Optional aPFN As String)
    mPFNname = aPFN
End Sub

Public Function Contains(section As String) As Boolean
  On Error Resume Next
  '"Extras->Optionen->Bei nicht verarbeiteten Fehlern"
    If IsEmpty(mSectionCol(section)) Then: 'DoNothing
    Contains = (Err.Number = 0)
  On Error GoTo 0
End Function

Public Property Let FileName(RHS As String)
    mPFNname = RHS
End Property
Public Property Get FileName() As String
    FileName = mPFNname
End Property

Public Sub Clear()
    Set mSectionCol = New Collection
    mIsLoaded = False
End Sub

Public Sub Load()
    Dim nSec As ConfigIniSection
    Dim StrSecArr() As String
    Dim StrSecName As Variant
Try: On Error GoTo Catch
    Call GetIniArr(StrSecArr)
    Call Clear
    Dim i As Long, u As Long: u = UBound(StrSecArr)
    For i = 0 To u
        'sSecName = StrSecArr(i)
        Set nSec = MNew.ConfigIniSection(Me, StrSecArr(i))
        Call nSec.Load
        Call mSectionCol.Add(nSec, StrSecArr(i))
    Next
    mIsLoaded = True
    Exit Sub
Catch:
    Call ErrHandler("Load")
End Sub
Public Property Get IsLoaded() As Boolean
    IsLoaded = mIsLoaded
End Property

Friend Property Let IsDirty(BolVal As Boolean)
    mIsDirty = BolVal
End Property
Public Property Get IsDirty() As Boolean
    IsDirty = mIsDirty
End Property

Public Sub Save()
    Dim asec As ConfigIniSection
    Dim aKey As ConfigIniKeyValue
    If mIsDirty Then
        For Each asec In mSectionCol
            For Each aKey In asec.Keys
                aKey.Save
            Next
        Next
    End If
End Sub
'##############################'  KeyValues  '##############################'
'gibt ein ConfigIniKeyValue-Objekt zurück:
'd.h. man kann entweder alle Value... hernehmen, oder dieses Key-Property
Public Property Get key(aSection As String, aKey As String) As ConfigIniKeyValue
    If mTempKey Is Nothing Then
        Set mTempKey = MNew.ConfigIniKeyValueS(Me, aSection, aKey, vbNullString)
    End If
    If Not (mTempKey.section.Name = aSection) Then
        mTempKey.section.Name = aSection
    End If
    mTempKey.Name = aKey
    Set key = mTempKey
End Property
'einzelne Key-Werte lesen/schreiben
'einen Boolean Lesen/Schreiben
Public Property Let ValueBol(aSection As String, aKey As String, Default As Boolean, BolVal As Boolean)
Try: On Error GoTo Catch
    ValueStr(aSection, aKey, CStr(Default)) = BolToStr(BolVal)
    Exit Property
Catch:
    Call ErrHandler("Let_ValueBol", aSection, aKey, CStr(Default))
End Property
Public Property Get ValueBol(aSection As String, aKey As String, Default As Boolean) As Boolean
Try: On Error GoTo Catch
    ValueBol = StrToBol(ValueStr(aSection, aKey, CStr(Default)))
    Exit Property
Catch:
    Call ErrHandler("Get_ValueBol", aSection, aKey, CStr(Default))
End Property

'einen Integer bzw. Long Lesen/Schreiben als String
Public Property Let ValueInt(aSection As String, aKey As String, Default As Long, LngVal As Long)
Try: On Error GoTo Catch
    ValueStr(aSection, aKey, CStr(Default)) = CStr(LngVal)
    Exit Property
Catch:
    Call ErrHandler("Let_ValueInt", aSection, aKey, CStr(Default))
End Property
Public Property Get ValueInt(aSection As String, aKey As String, Default As Long) As Long
Try: On Error GoTo Catch
    ValueInt = GetIniIntA(aSection, aKey, Default, mPFNname)
    Exit Property
Catch:
    Call ErrHandler("Get_ValueInt", aSection, aKey, CStr(Default))
End Property

'einen String Lesen/Schreiben
Public Property Let ValueStr(aSection As String, aKey As String, Default As String, StrVal As String)
Try: On Error GoTo Catch
    Dim rv As Long: rv = PutIniStrA(aSection, aKey, StrVal, mPFNname)
    Exit Property
Catch:
    Call ErrHandler("Let_ValueStr", aSection, aKey, Default, StrVal)
End Property
Public Property Get ValueStr(aSection As String, aKey As String, Default As String) As String
    Const ValLen As Long = 512: Dim StrVal As String * ValLen
Try: On Error GoTo Catch
    ValueStr = Left$(StrVal, GetIniStrA(aSection, aKey, Default, StrVal, ValLen, mPFNname))
    Exit Property
Catch:
    Call ErrHandler("Get_ValueStr", aSection, aKey, Default)
End Property

'ein Array Lesen/Schreiben
'das Array muß bei Let/Get ByRef übergeben werden,
'gesetzt und zurückgegeben wird nur ein RetVal
Public Property Let ValueStrArr(aSection As String, StrArr() As String, RetVal As Long)
Try: On Error GoTo Catch
    Dim Buffer As String, i As Long
    Dim l As Long: l = LBound(StrArr)
    Dim u As Long: u = UBound(StrArr)
    For i = l To u
        Buffer = Buffer & StrArr(i)
        If i < u Then Buffer = Buffer & vbNullChar
    Next
    'zuerst die section komplett löschen, da sie sonst
    Call PutIniStrA(aSection, vbNullString, vbNullString, mPFNname)
    Call PutIniSectionA(aSection, Buffer, mPFNname)
    Exit Property
Catch:
    Call ErrHandler("Let_ValueStrArr", aSection)
End Property
Public Property Get ValueStrArr(aSection As String, StrArr() As String) As Long
Try: On Error GoTo Catch
    Dim StrBuff As String: Call GetSectionBuffer(aSection, StrBuff)
    Dim i As Long, n As Long
    Dim pos As Long: pos = 1
    Dim oPos As Long: oPos = pos
    If Len(StrBuff) > 0 Then
        'String mit Trennzeichen Chr$(0) in ein Feld umwandeln
        n = CountNullCharsInBuffer(StrBuff)
        ReDim StrArr(0 To n - 1)
        For i = 0 To n - 1
            pos = InStr(oPos, StrBuff, vbNullChar)
            If pos > 0 Then
                StrArr(i) = Mid$(StrBuff, oPos, pos - oPos)
            End If
            oPos = pos + 1
        Next
    End If
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
Public Property Let ValueStructP(aSection As String, _
                                 aKey As String, _
                                 StructLen As Long, _
                                 pStruct As Long, _
                                 rv_out As Long)
Try: On Error GoTo Catch
    rv_out = PutIniStructA(aSection, aKey, ByVal pStruct, StructLen, mPFNname)
    Exit Property
Catch:
    Call ErrHandler("Let_ValueStructP", "Section: " & aSection & _
                    "; Key: " & aKey & "; " & CStr(StructLen) & "; ")
End Property
Public Property Get ValueStructP(aSection As String, _
                                 aKey As String, _
                                 StructLen As Long, _
                                 pStruct As Long) As Long
Try: On Error GoTo Catch
    ValueStructP = GetIniStructA(aSection, aKey, ByVal pStruct, StructLen, mPFNname)
    Exit Property
Catch:
    Call ErrHandler("Get_ValueStructP", aSection, aKey, CStr(StructLen))
End Property
'

'##############################'  Sections  '##############################'
Public Property Set Sections(aSecCol As Collection)
    Set mSectionCol = aSecCol
End Property
Public Property Get Sections() As Collection
    Set Sections = mSectionCol
End Property

Public Property Get section(IndexKey As Variant) As ConfigIniSection
Try: On Error GoTo Catch
    Set section = mSectionCol.Item(IndexKey)
    Exit Property
Catch:
    Call ErrHandler("Section", CStr(IndexKey))
End Property

Public Function SectionExists(aSectionName As String) As Boolean
    SectionExists = Not GetSection(aSectionName) Is Nothing
End Function
Private Function GetSection(aSectionName As String) As ConfigIniSection
Try: On Error Resume Next
    Set GetSection = mSectionCol.Item(aSectionName)
Catch: On Error GoTo 0
End Function
Public Function AddSection(aSectionName As String) As ConfigIniSection
'wenn die Section bereits existiert, wird diese zurückgegeben
    If Me.Contains(aSectionName) Then
        Set AddSection = GetSection(aSectionName) 'ohne IIf!!
    Else
        Set AddSection = MNew.ConfigIniSection(Me, aSectionName)
        Call mSectionCol.Add(AddSection, aSectionName)
    End If
End Function
Public Sub DeleteSection(aSectionName As String)
Try: On Error GoTo Catch
    Call PutIniStrA(aSectionName, vbNullString, vbNullString, mPFNname)
    If SectionExists(aSectionName) Then
        Call mSectionCol.Remove(aSectionName)
    End If
    Exit Sub
Catch:
    Call ErrHandler("DeleteSection", aSectionName)
End Sub
Public Sub DeleteKey(aSectionName As String, aKeyName As String)
Try: On Error GoTo Catch
    Call PutIniStrA(aSectionName, aKeyName, vbNullString, mPFNname)
    Dim asec As ConfigIniSection
    If SectionExists(aSectionName) Then
        Set asec = mSectionCol.Item(aSectionName)
        Call asec.DeleteKey(aKeyName)
    End If
    Exit Sub
Catch:
    Call ErrHandler("DeleteKey", aSectionName, aKeyName)
End Sub
Public Property Get CountEntriesInSection(aSection As String) As Long
    Dim StrBuff As String:  Call GetSectionBuffer(aSection, StrBuff)
Try: On Error GoTo Catch
    CountEntriesInSection = CountNullCharsInBuffer(StrBuff)
    Exit Function
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
        rv = GetIniSectNamesA(StrBuff, BuffLen, mPFNname)
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
    Dim SNamArr() As String: Call GetIniArr(SNamArr, mPFNname, aSection)
    Dim i As Long
    For i = 0 To UBound(SNamArr)
        Call SectionToCol.Add(SNamArr(i))
    Next
    Exit Function
Catch:
    Call ErrHandler("SectionToCol")
End Function
Public Sub GetIniArr(aStrArr() As String, _
                     Optional aIniPFN As String, _
                     Optional ByVal aSection As String)
    Dim rv As Long, sl As Long, pos As Long
    Dim StrBuff As String
    Dim iniPFN As String
Try: On Error GoTo Catch
    If Len(aIniPFN) Then iniPFN = aIniPFN Else iniPFN = mPFNname
    Dim BuffLen As Long: BuffLen = 256 '1
    Do While (sl = rv) Or (rv = 0)
        StrBuff = String$(BuffLen, vbNullChar)
        If Len(aSection) Then
            rv = GetIniSectionA(aSection, StrBuff, BuffLen, iniPFN)
        Else
            rv = GetIniSectNamesA(StrBuff, BuffLen, iniPFN)
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

Private Sub GetSectionBuffer(aSection As String, StrBuff As String)
    Dim rv As Long, sl As Long, pos As Long: pos = 1
Try: On Error GoTo Catch
    Dim BuffLen As Long: BuffLen = 256 '1
    Do While (sl = rv) Or (rv = 0)
      StrBuff = String$(BuffLen, vbNullChar)
      rv = GetIniSectionA(aSection, StrBuff, BuffLen, mPFNname)
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
    Dim section As ConfigIniSection
    For Each section In mSectionCol
        ToStr = ToStr & section.ToStr & vbCrLf
    Next
End Function
  
'##############################'  Privates  '##############################'
Private Function StrToBol(StrVal As String) As Boolean
    If (StrComp(StrVal, "0", vbTextCompare) = 0) Or _
       (StrComp(StrVal, "false", vbTextCompare) = 0) Or _
       (StrComp(StrVal, "falsch", vbTextCompare) = 0) Or _
       (StrComp(StrVal, "nein", vbTextCompare) = 0) Then
        StrToBol = False
    'ElseIf (StrComp(StrVal, vbNullString) = 0) Or _
           (StrComp(StrVal, "1") = 0) Or _
           (StrComp(StrVal, "-1") = 0) Or _
           (StrComp(StrVal, "true") = 0) Or _
           (StrComp(StrVal, "wahr") = 0) Or _
           (StrComp(StrVal, "ja") = 0) Then
    Else
        StrToBol = True
    End If
End Function
Private Function BolToStr(BolVal As Boolean) As String
    If BolVal Then BolToStr = "True" Else BolToStr = "False"
End Function

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

'####################' v     Local ErrHandler      v '####################'
Private Function ErrHandler(ByVal aProcName As String, _
                            Optional aSection As String, _
                            Optional aKey As String, _
                            Optional AddInfo1 As String, _
                            Optional AddInfo2 As String, _
                            Optional bolmess As Boolean = True, _
                            Optional decor As VbMsgBoxStyle) As VbMsgBoxResult
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
    
    If bolmess Then ErrHandler = GlobalErrHandler(Me, aProcName, Info, decor)

End Function
