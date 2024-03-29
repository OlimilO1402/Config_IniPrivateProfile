Attribute VB_Name = "MNew"
Option Explicit
Public Type Point2DF 'F f�r float
    X As Single
    Y As Single
End Type
Public Type SizeF
    Width  As Single
    Height As Single
End Type
Public Type PosSizeF
    Position As Point2DF
    Size     As SizeF
End Type
Public Type TestTyp1
    BolVal As Boolean      ' 2
    IntVal As Integer      ' 2
    LngVal As Long         ' 4
    SngVal As Single       ' 4
    DblVal As Double       ' 8
    StrVal As String * 10  '20 'fixe L�nge!!
                      'Sum: 40
    'in einem Type gehen nur Strings mit fester L�nge.
    'Bei Types mit Strings mit variabler L�nge, muss man so
    'vorgehen, dass jede Variable f�r sich geschrieben wird.
End Type
Public Type TestTyp2
    BolVal As Boolean
    IntVal As Integer
    LngVal As Long
    SngVal As Single
    DblVal As Double
    StrVal As String 'variable L�nge
End Type              'Sum: 40

Public Function Point2DF(aCtrl As Form) As Point2DF 'aCtrl As VBControlExtender
    With Point2DF
        .X = aCtrl.Left
        .Y = aCtrl.Top
    End With
End Function

Public Function SizeF(aCtrl As Form) As SizeF 'aCtrl As VBControlExtender
    With SizeF
        .Width = aCtrl.Width
        .Height = aCtrl.Height
    End With
End Function

Public Function PosSizeF(aCtrl As Form) As PosSizeF 'aCtrl As VBControlExtender
    With PosSizeF
        .Position = MNew.Point2DF(aCtrl)
        .Size = SizeF(aCtrl)
    End With
End Function
Public Function PosSizeF_ToStr(this As PosSizeF) As String
    Dim s As String
    With this
        With .Position
            s = "PosSizeF{X=" & .X & "; Y=" & .Y & "; "
        End With
        With .Size
            s = s & "Height=" & .Height & "; Width=" & .Width
        End With
    End With
    PosSizeF_ToStr = s & "}"
End Function

Public Sub PosSizeToControl(aCtrl As Form, ps As PosSizeF)
    aCtrl.Move ps.Position.X, ps.Position.Y, ps.Size.Width, ps.Size.Height
End Sub

Public Function ConfigIniDocument(aPFN As PathFileName) As ConfigIniDocument
    Set ConfigIniDocument = New ConfigIniDocument: ConfigIniDocument.New_ aPFN
End Function

Public Function ConfigIniSection(IniFile As ConfigIniDocument, ByVal SectionName As String) As ConfigIniSection
    Set ConfigIniSection = New ConfigIniSection: ConfigIniSection.New_ IniFile, SectionName
End Function

Public Function ConfigIniKeyValue(IniFile As ConfigIniDocument, Section As ConfigIniSection, ByVal KeyName As String, Optional VarDefault As Variant) As ConfigIniKeyValue
    Set ConfigIniKeyValue = New ConfigIniKeyValue: ConfigIniKeyValue.New_ IniFile, Section, KeyName, VarDefault
End Function

Public Function ConfigIniKeyValueS(IniFile As ConfigIniDocument, ByVal Section As String, ByVal KeyName As String, Optional VarDefault As Variant) As ConfigIniKeyValue
    Set ConfigIniKeyValueS = New ConfigIniKeyValue: ConfigIniKeyValueS.NewS IniFile, Section, KeyName, VarDefault
End Function

Public Function PathFileName(ByVal aPathFileName As String, _
                    Optional ByVal aFileName As String, _
                    Optional ByVal aExt As String) As PathFileName
    Set PathFileName = New PathFileName: PathFileName.New_ aPathFileName, aFileName, aExt
End Function

Public Function List(Of_T As EDataType, _
                     Optional ArrColStrTypList, _
                     Optional ByVal IsHashed As Boolean = False, _
                     Optional ByVal Capacity As Long = 32, _
                     Optional ByVal GrowRate As Single = 2, _
                     Optional ByVal GrowSize As Long = 0) As List
    Set List = New List: List.New_ Of_T, ArrColStrTypList, IsHashed, Capacity, GrowRate, GrowSize
End Function

