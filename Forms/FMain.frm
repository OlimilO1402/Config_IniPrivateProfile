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
   Begin VB.TextBox Text2 
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
      Left            =   4800
      MultiLine       =   -1  'True
      OLEDragMode     =   1  'Automatisch
      OLEDropMode     =   1  'Manuell
      ScrollBars      =   3  'Beides
      TabIndex        =   12
      ToolTipText     =   "Drag'n'drop *.ini- or *.vbp/vbg-files here"
      Top             =   1200
      Width           =   4935
   End
   Begin VB.CommandButton BtnOpenIni 
      Caption         =   "Open Ini"
      Height          =   375
      Left            =   8160
      TabIndex        =   11
      Top             =   720
      Width           =   1575
   End
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
      ToolTipText     =   "Drag'n'drop *.ini- or *.vbp-files here"
      Top             =   1200
      Width           =   4815
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
    If W > 0 And H > 0 Then
        Text1.Move L, T, W / 2, H
        Text2.Move W / 2, T, W / 2, H
    End If
End Sub

Private Sub UpdateView()
Try: On Error GoTo Catch
    Text1.Text = vbNullString
    Text2.Text = vbNullString
    Label1.Caption = IniFile.PFN.Value
    Text1.Text = IniFile.PFN.ReadAllText
    IniFile.Load
    Text2.Text = IniFile.ToStr
    IniFile.PFN.CloseFile
    Exit Sub
Catch:
    ErrHandler "UpdateView"
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
    If Not IniFile.PFN.Exists Then
        MsgBox "File not found!" & vbCrLf & IniFile.PFN.Value
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
    Dim KeyValue As ConfigIniKeyValue
    
    SectionName = "TestReadWriteAtOnce"
    KeyValueName = "FirstEntry"
    
    IniFile.ValueStr(SectionName, KeyValueName, "") = "NewValueOfFirstEntry"
    
    'read from ini file what we have written:
    Value = IniFile.ValueStr(SectionName, KeyValueName, "")
    MsgBox "The read value is: " & Value
    
    SectionName = "TestSection1"
    Set Section = IniFile.AddSection(SectionName)
    
    KeyValueName = "FirstEntry"
    Set KeyValue = Section.AddKeyValue(KeyValueName)
    KeyValue.ValueInt = 123456
    
    KeyValueName = "SecondEntry"
    Set KeyValue = Section.AddKeyValue(KeyValueName)
    KeyValue.ValueInt = 456789
    
    'it's also possible to write UD-Type-variables at once:
    KeyValueName = "Form1PositionAndSize"
    Set KeyValue = Section.AddKeyValue(KeyValueName)
    
    Dim cs As PosSizeF: cs = MNew.PosSizeF(Me)
    Dim rv As Long
    KeyValue.ValueStructP(LenB(cs), VarPtr(cs)) = VarPtr(cs)
    
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
    Set KeyValue = Section.AddKeyValue(KeyValueName)
    
    KeyValue.ValueStructP(LenB(tt), VarPtr(tt)) = VarPtr(tt)
    
    'write a value yourself
    KeyValueName = "MyEntry"
    Set KeyValue = Section.AddKeyValue(KeyValueName)
    Value = InputBox("Write a value yourself: ", "Me too", "hoho")
    If Not (Len(Value) = 0) Then
        KeyValue.ValueStr = Value
    End If
    UpdateView
End Sub

Private Sub BtnReadIniFile_Click()
    'read Ini-file and display it
    'Dim IniFile As ConfigIniDocument:
    'Call IniFile.Load
    Set IniFile = MNew.ConfigIniDocument(MNew.PathFileName(Environ("Temp") & "\Test.ini"))
    
    If Not IniFile.PFN.Exists Then
        If MsgBox("Inifile does not exist, write it first?" & vbCrLf & IniFile.PFN.Value, vbOKCancel) = vbCancel Then Exit Sub
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
    If Not IniFile.PFN.Exists Then
        MsgBox "File not found, write it first!" & vbCrLf & IniFile.PFN.Value
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
    If Not IniFile.PFN.Exists Then
        MsgBox "File not found, write it first!" & vbCrLf & IniFile.PFN.Value
        Exit Sub
    End If
    
    Dim Section As ConfigIniSection:  Set Section = IniFile.Section("TestSection1")
    
    Dim KeyName  As String: KeyName = "Form1PositionAndSize"
    Dim KeyValue As ConfigIniKeyValue: Set KeyValue = Section.AddKeyValue(KeyName)
    
    Dim cs As PosSizeF
    
    Dim rv As Long: rv = KeyValue.ValueStructP(LenB(cs), VarPtr(cs))
    
    With cs
        Me.WindowState = FormWindowStateConstants.vbNormal
        Me.Move .Position.X, .Position.Y, .Size.Width, .Size.Height
    End With
End Sub

Private Sub BtnReadRawIniData_Click()
    Set IniFile = MNew.ConfigIniDocument(MNew.PathFileName(Environ("Temp") & "\Test.ini"))
    If Not IniFile.PFN.Exists Then
        MsgBox "File not found, write it first!" & vbCrLf & IniFile.PFN.Value
        Exit Sub
    End If
    UpdateView
End Sub

Private Sub BtnDeleteIniFile_Click()
    If Not IniFile.PFN.Exists Then
        MsgBox "File not found, nothing to delete here" & vbCrLf & IniFile.PFN.Value
        Exit Sub
    End If
Try: On Error GoTo Catch
    IniFile.PFN.Delete
Catch:
End Sub

Private Sub BtnTestVBP_Click()
    
    TestVBP App.Path & "\PConfigIni.vbp"
    
End Sub

Sub TestVBP(sPfnVBP As String)
    Set IniFile = MNew.ConfigIniDocument(MNew.PathFileName(sPfnVBP))
    If Not IniFile.PFN.Exists Then
        MsgBox "File not found:" & vbCrLf & IniFile.PFN.Value
        Exit Sub
    End If
    
    UpdateView
    
    Dim i As Long, u As Long
    Dim s As String, cikv As ConfigIniKeyValue
    
    Dim startupprojects As ConfigIniSection: Set startupprojects = IniFile.Root.Filter("StartupProject")
    u = startupprojects.KeyValues.Count - 1
    If 0 <= u Then
        s = s & "StartupProject:" & vbCrLf & "===============" & vbCrLf
        For i = 0 To u
            Set cikv = startupprojects.KeyValues.Item(i)
            s = s & cikv.Value & vbCrLf
        Next
        s = s & vbCrLf
    End If
    
    Dim projects As ConfigIniSection: Set projects = IniFile.Root.Filter("Project")
    u = projects.KeyValues.Count - 1
    If 0 <= u Then
        s = s & "Projects:" & vbCrLf & "=========" & vbCrLf
        For i = 0 To u
            Set cikv = projects.KeyValues.Item(i)
            s = s & cikv.Value & vbCrLf
        Next
        s = s & vbCrLf
    End If
    
    Dim references As ConfigIniSection: Set references = IniFile.Root.Filter("Reference")
    u = references.KeyValues.Count - 1
    If 0 <= u Then
        s = s & "References:" & vbCrLf & "===========" & vbCrLf
        For i = 0 To u
            Set cikv = references.KeyValues.Item(i)
            s = s & cikv.Value & vbCrLf
        Next
        s = s & vbCrLf
    End If
    
    Dim objects As ConfigIniSection: Set objects = IniFile.Root.Filter("Object")
    u = objects.KeyValues.Count - 1
    If 0 <= u Then
        s = s & "Objects:" & vbCrLf & "========" & vbCrLf
        For i = 0 To u
            Set cikv = objects.KeyValues.Item(i)
            s = s & cikv.Value & vbCrLf
        Next
        s = s & vbCrLf
    End If
    
    Dim forms As ConfigIniSection: Set forms = IniFile.Root.Filter("Form")
    u = forms.KeyValues.Count - 1
    If 0 <= u Then
        s = s & "Forms:" & vbCrLf & "======" & vbCrLf
        For i = 0 To u
            Set cikv = forms.KeyValues.Item(i)
            s = s & cikv.Value & vbCrLf
        Next
        s = s & vbCrLf
    End If
    
    Dim classes As ConfigIniSection: Set classes = IniFile.Root.Filter("Class")
    u = classes.KeyValues.Count - 1
    If 0 <= u Then
        s = s & "Classes:" & vbCrLf & "========" & vbCrLf
        For i = 0 To u
            Set cikv = classes.KeyValues.Item(i)
            s = s & cikv.Value & vbCrLf
        Next
        s = s & vbCrLf
    End If
    
    Dim modules As ConfigIniSection: Set modules = IniFile.Root.Filter("Module")
    u = modules.KeyValues.Count - 1
    If 0 <= u Then
        s = s & "Modules:" & vbCrLf & "========" & vbCrLf
        For i = 0 To u
            Set cikv = modules.KeyValues.Item(i)
            s = s & cikv.Value & vbCrLf
        Next
        s = s & vbCrLf
    End If
    
    Dim userctrls As ConfigIniSection: Set userctrls = IniFile.Root.Filter("UserControl")
    u = userctrls.KeyValues.Count - 1
    If 0 <= u Then
        s = s & "UserControl:" & vbCrLf & "============" & vbCrLf
        For i = 0 To u
            Set cikv = userctrls.KeyValues.Item(i)
            s = s & cikv.Value & vbCrLf
        Next
        s = s & vbCrLf
    End If
    
    Dim designers As ConfigIniSection: Set designers = IniFile.Root.Filter("Designer")
    u = designers.KeyValues.Count - 1
    If 0 <= u Then
        s = s & "Designer:" & vbCrLf & "=========" & vbCrLf
        For i = 0 To u
            Set cikv = designers.KeyValues.Item(i)
            s = s & cikv.Value & vbCrLf
        Next
    End If
    
    Text2.Text = Text2.Text & vbCrLf & "##############################" & vbCrLf & s
End Sub
Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    TBOLEDragDrop Text1, Data, Effect, Button, Shift, X, Y
End Sub

Private Sub Text2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    TBOLEDragDrop Text2, Data, Effect, Button, Shift, X, Y
End Sub

Sub TBOLEDragDrop(TB As TextBox, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not Data.GetFormat(ClipBoardConstants.vbCFFiles) Then Exit Sub
    If Data.Files.Count = 0 Then Exit Sub
    Dim s As String: s = Data.Files.Item(1)
    If Len(s) = 0 Then
        MsgBox "String is not a valid filename"
        Exit Sub
    End If
    Dim PFN As PathFileName: Set PFN = MNew.PathFileName(s)
    Dim ext As String: ext = LCase(PFN.Extension)
    If ext = ".ini" Then
        CreateIni s
    ElseIf ext = ".vbp" Or ext = ".vbg" Then
        TestVBP s
    Else
        MsgBox "Sorry can't read file format: " & ext & "Just email me if you like me to implement this"
    End If
End Sub

Private Sub BtnOpenIni_Click()
    Dim OFD As OpenFileDialog: Set OFD = New OpenFileDialog
    OFD.Filter = "Ini-files |*.ini"
    If OFD.ShowDialog(Me) = vbCancel Then Exit Sub
    CreateIni OFD.FileName
End Sub

Private Sub CreateIni(FNm As String)
Try: On Error GoTo Catch
    Dim PFN As PathFileName
    Set PFN = MNew.PathFileName(FNm)
    If Not PFN.Exists Then Exit Sub
    Set IniFile = MNew.ConfigIniDocument(PFN)
    UpdateView
    Exit Sub
Catch:
    ErrHandler "CreateIni"
End Sub

'copy this same function to every class, form or module
'the name of the class or form will be added automatically
'in standard-modules the function "TypeName(Me)" will not work, so simply replace it with the name of the Module
' v ############################## v '   Local ErrHandler   ' v ############################## v '
Private Function ErrHandler(ByVal FuncName As String, _
                            Optional ByVal AddInfo As String, _
                            Optional WinApiError, _
                            Optional bLoud As Boolean = True, _
                            Optional bErrLog As Boolean = True, _
                            Optional vbDecor As VbMsgBoxStyle = vbOKCancel, _
                            Optional bRetry As Boolean) As VbMsgBoxResult

    If bRetry Then

        ErrHandler = MessErrorRetry(TypeName(Me), FuncName, AddInfo, WinApiError, bErrLog)

    Else

        ErrHandler = MessError(TypeName(Me), FuncName, AddInfo, WinApiError, bLoud, bErrLog, vbDecor)

    End If

End Function


