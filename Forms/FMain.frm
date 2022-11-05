VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Ini-Configfile"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9855
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   9855
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnTestVBP 
      Caption         =   "Test VBP"
      Height          =   375
      Left            =   8160
      TabIndex        =   8
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton BtnDeleteIniFile 
      Caption         =   "Delete Ini-file"
      Height          =   375
      Left            =   6600
      TabIndex        =   7
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton BtnReadRawIniData 
      Caption         =   "ReadRawIniData"
      Height          =   375
      Left            =   6600
      TabIndex        =   6
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton BtnSetWindowPosSize 
      Caption         =   "Read and set PosAndSize of window"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   720
      Width           =   3135
   End
   Begin VB.CommandButton BtnWriteWindowPosSize 
      Caption         =   "Write PosAndSize of window"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   360
      Width           =   3135
   End
   Begin VB.CommandButton BtnReadIniFile 
      Caption         =   "Read Ini-file"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton BtnWriteIniFile 
      Caption         =   "Write Ini-file"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton BtnTestReadAtOnce 
      Caption         =   "Test ReadeAtOnce"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton BtnTestWriteAtOnce 
      Caption         =   "Test WriteAtOnce"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   0
      MultiLine       =   -1  'True
      OLEDragMode     =   1  'Automatisch
      OLEDropMode     =   1  'Manuell
      ScrollBars      =   3  'Beides
      TabIndex        =   10
      ToolTipText     =   "Drag'n'drop ini files here"
      Top             =   1200
      Width           =   9735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "        "
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private IniFile As ConfigIniDocument

Private Sub Form_Load()
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_Resize()
    Dim L As Single
    Dim T As Single: T = Text1.Top
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then Text1.Move L, T, W, H
End Sub

Private Sub UpdateView()
    
    Label1.Caption = IniFile.pfn.Value
    Text1.Text = IniFile.pfn.ReadAllText
    IniFile.pfn.CloseFile
    
End Sub

Private Sub BtnTestWriteAtOnce_Click()
    
    Set IniFile = MNew.ConfigIniDocument(MNew.PathFileName(App.Path & "\mynewini.ini"))
    
    Dim SectionName As String, KeyValueName As String
    SectionName = "SectionOne"
    KeyValueName = "FirstValue":  IniFile.ValueBol(SectionName, KeyValueName, False) = True
    KeyValueName = "SecondValue": IniFile.ValueInt(SectionName, KeyValueName, 0) = 123456
    
    SectionName = "SectionTwo"
    KeyValueName = "ThirdValue":  IniFile.ValueStr(SectionName, KeyValueName, "Null") = "Eins"
    Dim rv As Long, p As PosSizeF: p = MNew.PosSizeF(Me)
    KeyValueName = "FourthValue":  IniFile.ValueStructP(SectionName, KeyValueName, LenB(p), VarPtr(p)) = rv
    
    UpdateView
    
End Sub

Private Sub BtnTestReadAtOnce_Click()
    
    Set IniFile = MNew.ConfigIniDocument(MNew.PathFileName(App.Path & "\mynewini.ini"))
    If Not IniFile.pfn.Exists Then
        MsgBox "File not found!" & vbCrLf & IniFile.pfn.Value
        Exit Sub
    End If
    Dim p As PosSizeF, rv As Long, txt As String
    Dim SectionName As String, KeyValueName As String
    SectionName = "SectionOne"
    txt = txt & "[" & SectionName & "]" & vbCrLf
    KeyValueName = "FirstValue":  Dim b As Boolean: b = IniFile.ValueBol(SectionName, KeyValueName, False):    txt = txt & KeyValueName & " = " & b & vbCrLf
    KeyValueName = "SecondValue": Dim i As Long:    i = IniFile.ValueInt(SectionName, KeyValueName, 1):        txt = txt & KeyValueName & " = " & i & vbCrLf
    
    SectionName = "SectionTwo"
    txt = txt & "[" & SectionName & "]" & vbCrLf
    KeyValueName = "ThirdValue":  Dim s As String:  s = IniFile.ValueStr(SectionName, KeyValueName, "Null"):   txt = txt & KeyValueName & " = " & s & vbCrLf
    KeyValueName = "FourthValue":  rv = IniFile.ValueStructP(SectionName, KeyValueName, LenB(p), VarPtr(p)):   txt = txt & KeyValueName & " = " & MNew.PosSizeF_ToStr(p) & vbCrLf
    
    Dim SectionNames As Collection: Set SectionNames = IniFile.SectionNamesToCol
    Dim v
    For Each v In SectionNames
        txt = txt & v & vbCrLf
    Next
    
    Dim iarr() As String: IniFile.GetIniArr iarr, , SectionName
    
    For i = 0 To UBound(iarr)
        v = iarr(i)
        txt = txt & v & vbCrLf
    Next
    
    UpdateView
    Text1.Text = Text1.Text & txt
    
End Sub

Private Sub BtnWriteIniFile_Click()
    
    'directly write some values to the Ini-file
    'by using the functions ValueStr, ValueBol & ValueInt you can
    'immediately write to the Ini-file
    'these function you will find in the class ConfigIniDocument
    'as well as in the class ConfigIniKeyValue
    Set IniFile = MNew.ConfigIniDocument(MNew.PathFileName(Environ("Temp") & "\Test.ini"))
'    If Not IniFile.pfn.Exists Then
'        MsgBox "File not found, write it first!" & vbCrLf & IniFile.pfn.Value
'        Exit Sub
'    End If
    
    Dim SectionName  As String
    Dim KeyValueName As String
    Dim Value        As String
    Dim Section  As ConfigIniSection
    Dim keyvalue As ConfigIniKeyValue
    
    SectionName = "TestReadWriteAtOnce"
    KeyValueName = "FirstEntry"
    
    IniFile.ValueStr(SectionName, KeyValueName, "") = "NewValueOfFirstEntry"
    
    'read from ini file what we have written:
    Value = IniFile.ValueStr(SectionName, KeyValueName, "")
    MsgBox "The read value is: " & Value
    
    SectionName = "TestSection1"
    Set Section = IniFile.AddSection(SectionName)
    
    KeyValueName = "FirstEntry"
    Set keyvalue = Section.AddKeyValue(KeyValueName)
    keyvalue.ValueInt = 123456
    
    KeyValueName = "SecondEntry"
    Set keyvalue = Section.AddKeyValue(KeyValueName)
    keyvalue.ValueInt = 456789
    
    'it's also possible to write UD-Type-variables at once:
    KeyValueName = "Form1PositionAndSize"
    Set keyvalue = Section.AddKeyValue(KeyValueName)
    
    Dim cs As PosSizeF: cs = MNew.PosSizeF(Me)
    Dim rv As Long
    keyvalue.ValueStructP(LenB(cs), VarPtr(cs)) = VarPtr(cs)
    
    Dim tt As TestTyp1
    With tt
        .BolVal = True
        .IntVal = 12345
        .LngVal = 123456789
        .SngVal = 0.123456
        .DblVal = 0.123456789
        .StrVal = "Test Entry"
    End With
    
    KeyValueName = "tt_As_TestTyp"
    Set keyvalue = Section.AddKeyValue(KeyValueName)
    
    keyvalue.ValueStructP(LenB(tt), VarPtr(tt)) = VarPtr(tt)
    
    'write a value yourself
    KeyValueName = "MyEntry"
    Set keyvalue = Section.AddKeyValue(KeyValueName)
    Value = InputBox("Write a value yourself: ", "Me too", "hoho")
    If Not (Len(Value) = 0) Then
        keyvalue.ValueStr = Value
    End If
    UpdateView
End Sub

Private Sub BtnReadIniFile_Click()
    'read Ini-file and display it
    'Dim IniFile As ConfigIniDocument:
    'Call IniFile.Load
    Set IniFile = MNew.ConfigIniDocument(MNew.PathFileName(Environ("Temp") & "\Test.ini"))
    
    If Not IniFile.pfn.Exists Then
        If MsgBox("Inifile does not exist, write it first?" & vbCrLf & IniFile.pfn.Value, vbOKCancel) = vbCancel Then Exit Sub
        BtnWriteIniFile_Click
    End If
    If IniFile Is Nothing Then
        Set IniFile = MNew.ConfigIniDocument(MNew.PathFileName(Environ("Temp") & "\Test.ini")):
    End If
    IniFile.Load
    'Text1.Text = IniFile.ToStr
    UpdateView
End Sub

Private Sub BtnWriteWindowPosSize_Click()
    
    Set IniFile = MNew.ConfigIniDocument(MNew.PathFileName(Environ("Temp") & "\Test.ini")): IniFile.Load
    If Not IniFile.pfn.Exists Then
        MsgBox "File not found, write it first!" & vbCrLf & IniFile.pfn.Value
        Exit Sub
    End If
    Dim Section As ConfigIniSection:  Set Section = IniFile.Section("TestSection1")
    
    Dim Key As String: Key = "Form1PositionAndSize"
    Dim KyValue As ConfigIniKeyValue: Set KyValue = Section.AddKeyValue(Key)
    
    Dim cs As PosSizeF: cs = MNew.PosSizeF(Me)
    
    Dim rv As Long: KyValue.ValueStructP(LenB(cs), VarPtr(cs)) = rv
End Sub

Private Sub BtnSetWindowPosSize_Click()
    
    Set IniFile = MNew.ConfigIniDocument(MNew.PathFileName(Environ("Temp") & "\Test.ini")): IniFile.Load
    If Not IniFile.pfn.Exists Then
        MsgBox "File not found, write it first!" & vbCrLf & IniFile.pfn.Value
        Exit Sub
    End If
    
    Dim Section As ConfigIniSection:  Set Section = IniFile.Section("TestSection1")
    
    Dim KeyName  As String: KeyName = "Form1PositionAndSize"
    Dim keyvalue As ConfigIniKeyValue: Set keyvalue = Section.AddKeyValue(KeyName)
    
    Dim cs As PosSizeF
    
    Dim rv As Long: rv = keyvalue.ValueStructP(LenB(cs), VarPtr(cs))
    
    With cs
        Me.WindowState = FormWindowStateConstants.vbNormal
        Me.Move .Position.X, .Position.Y, .Size.Width, .Size.Height
    End With
End Sub

Private Sub BtnReadRawIniData_Click()
    Set IniFile = MNew.ConfigIniDocument(MNew.PathFileName(Environ("Temp") & "\Test.ini"))
    If Not IniFile.pfn.Exists Then
        MsgBox "File not found, write it first!" & vbCrLf & IniFile.pfn.Value
        Exit Sub
    End If
    UpdateView
End Sub

Private Sub BtnDeleteIniFile_Click()
    If Not IniFile.pfn.Exists Then
        MsgBox "File not found, nothing to delete here" & vbCrLf & IniFile.pfn.Value
        Exit Sub
    End If
Try: On Error GoTo Catch
    IniFile.pfn.Delete
Catch:
End Sub

Private Sub BtnTestVBP_Click()
    
    Set IniFile = MNew.ConfigIniDocument(MNew.PathFileName(App.Path & "\PConfigIni.vbp"))
    If Not IniFile.pfn.Exists Then
        MsgBox "File not found:" & vbCrLf & IniFile.pfn.Value
        Exit Sub
    End If
    IniFile.Load
    Dim i As Long, u As Long
    Dim s As String, cikv As ConfigIniKeyValue
    
    Dim classes As ConfigIniSection: Set classes = IniFile.Root.Filter("Class")
    u = classes.KeyValues.Count - 1
    For i = 0 To u
        Set cikv = classes.KeyValues.Item(i)
        s = s & cikv.Value & vbCrLf
    Next
    
    s = s & vbCrLf
    
    Dim modules As ConfigIniSection: Set modules = IniFile.Root.Filter("Module")
    u = modules.KeyValues.Count - 1
    For i = 0 To u
        Set cikv = modules.KeyValues.Item(i)
        s = s & cikv.Value & vbCrLf
    Next
    
    s = s & vbCrLf
    
    Dim forms As ConfigIniSection: Set forms = IniFile.Root.Filter("Form")
    u = forms.KeyValues.Count - 1
    For i = 0 To u
        Set cikv = forms.KeyValues.Item(i)
        s = s & cikv.Value & vbCrLf
    Next
    UpdateView
    Text1.Text = Text1.Text & vbCrLf & "##############################" & vbCrLf & s

End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not Data.GetFormat(ClipBoardConstants.vbCFFiles) Then Exit Sub
    Dim pfn As PathFileName
    If Data.Files.Count = 0 Then Exit Sub
    Set pfn = MNew.PathFileName(Data.Files.Item(1))
    If Not pfn.Exists Then Exit Sub
    Set IniFile = MNew.ConfigIniDocument(pfn)
    IniFile.Load
    Text1.Text = IniFile.ToStr
End Sub
