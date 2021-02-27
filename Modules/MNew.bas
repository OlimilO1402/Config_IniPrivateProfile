Attribute VB_Name = "MNew"
Option Explicit
Public Type Point2DF 'F für float
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
    StrVal As String * 10  '20 'fixe Länge!!
                      'Sum: 40
    'in einem Type gehen nur Strings mit fester Länge.
    'Bei Types mit Strings mit variabler Länge, muss man so
    'vorgehen, dass jede Variable für sich geschrieben wird.
End Type
Public Type TestTyp2
    BolVal As Boolean
    IntVal As Integer
    LngVal As Long
    SngVal As Single
    DblVal As Double
    StrVal As String 'variable Länge
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
Public Sub PosSizeToControl(aCtrl As Form, ps As PosSizeF)
    aCtrl.Move ps.Position.X, ps.Position.Y, ps.Size.Width, ps.Size.Height
'    With aCtrl
'        .Left
'        .Top
'        .Width
'        .Height
'    End With
End Sub
Public Function ConfigIniDocument(aPFN As String) As ConfigIniDocument
    Set ConfigIniDocument = New ConfigIniDocument: ConfigIniDocument.New_ aPFN
End Function

Public Function ConfigIniSection(aIniFile As ConfigIniDocument, StrSectionName As String) As ConfigIniSection
    Set ConfigIniSection = New ConfigIniSection: ConfigIniSection.New_ aIniFile, StrSectionName
End Function

Public Function ConfigIniKeyValue(aIniFile As ConfigIniDocument, _
                                  aSection As ConfigIniSection, _
                                  aKeyName As String, _
                                  Optional VarDefault As Variant) As ConfigIniKeyValue
    Set ConfigIniKeyValue = New ConfigIniKeyValue: ConfigIniKeyValue.New_ aIniFile, aSection, aKeyName, VarDefault
End Function

Public Function ConfigIniKeyValueS(aIniFile As ConfigIniDocument, _
                                   aSection As String, _
                                   aKeyName As String, _
                                   Optional VarDefault As Variant) As ConfigIniKeyValue
    Set ConfigIniKeyValueS = New ConfigIniKeyValue: ConfigIniKeyValueS.NewS aIniFile, aSection, aKeyName, VarDefault
End Function
