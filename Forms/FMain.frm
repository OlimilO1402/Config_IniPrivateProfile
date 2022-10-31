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
      TabIndex        =   6
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton BtnDeleteIniFile 
      Caption         =   "Delete Ini-file"
      Height          =   375
      Left            =   6600
      TabIndex        =   7
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton BtnReadRawIniData 
      Caption         =   "ReadRawIniData"
      Height          =   375
      Left            =   6600
      TabIndex        =   8
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton BtnSetWindowPosSize 
      Caption         =   "Read and set PosAndSize of window"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   840
      Width           =   3135
   End
   Begin VB.CommandButton BtnWriteWindowPosSize 
      Caption         =   "Write PosAndSize of window"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   480
      Width           =   3135
   End
   Begin VB.CommandButton BtnReadIniFile 
      Caption         =   "Read Ini-file"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton BtnWriteIniFile 
      Caption         =   "Write Ini-file"
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton BtnTestReadAtOnce 
      Caption         =   "Test ReadeAtOnce"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton BtnTestWriteAtOnce 
      Caption         =   "Test WriteAtOnce"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   480
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
      ScrollBars      =   3  'Beides
      TabIndex        =   4
      Top             =   1320
      Width           =   9735
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private IniPFN  As PathFileName 'String
Private IniFile As ConfigIniDocument

Private Sub Form_Load()
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    Set IniPFN = MNew.PathFileName(Environ("Temp") & "\Test.ini")
    Set IniFile = MNew.ConfigIniDocument(IniPFN)
    Label1.Caption = IniPFN.Value
End Sub

Private Sub BtnTestWriteAtOnce_Click()
    Dim FNm As String: FNm = App.Path & "\mynewini.ini"
    Label1.Caption = FNm
    Dim ini As ConfigIniDocument: Set ini = MNew.ConfigIniDocument(MNew.PathFileName(FNm))
    Dim aSection As String, aKey As String
    aSection = "SectionOne"
    aKey = "FirstValue":  ini.ValueBol(aSection, aKey, False) = True
    aKey = "SecondValue": ini.ValueInt(aSection, aKey, 0) = 123456
    aSection = "SectionTwo"
    aKey = "ThirdValue":  ini.ValueStr(aSection, aKey, "Null") = "Eins"
    Dim p As PosSizeF: p = MNew.PosSizeF(Me)
    Dim rv As Long
    aKey = "FourthValue":  ini.ValueStructP(aSection, aKey, LenB(p), VarPtr(p)) = rv
    'Debug.Print rv
    Dim s As String: s = ini.pfn.ReadAllStr: ini.pfn.CloseFile
    Text1.Text = s 'IIf(ini.FileIsUnicode, StrConv(s, vbFromUnicode), s)
    'ini.pfn.OpenFileExplorer
End Sub
Private Sub BtnTestReadAtOnce_Click()
    Text1.Text = ""
    Dim T As String
    Dim ini As ConfigIniDocument: Set ini = MNew.ConfigIniDocument(MNew.PathFileName(App.Path & "\mynewini.ini"))
    Dim p As PosSizeF: Dim rv As Long
    Dim aSection As String, aKey As String
    aSection = "SectionOne"
    T = T & "[" & aSection & "]" & vbCrLf
    aKey = "FirstValue":  Dim b As Boolean: b = ini.ValueBol(aSection, aKey, False):    T = T & aKey & " = " & b & vbCrLf
    aKey = "SecondValue": Dim i As Long:    i = ini.ValueInt(aSection, aKey, 1):        T = T & aKey & " = " & i & vbCrLf
    
    aSection = "SectionTwo"
    T = T & "[" & aSection & "]" & vbCrLf
    aKey = "ThirdValue":  Dim s As String:  s = ini.ValueStr(aSection, aKey, "Null"):   T = T & aKey & " = " & s & vbCrLf
    aKey = "FourthValue": rv = ini.ValueStructP(aSection, aKey, LenB(p), VarPtr(p)):    T = T & aKey & " = " & MNew.PosSizeF_ToStr(p) & vbCrLf

    Dim secnms As Collection: Set secnms = ini.SectionNamesToCol
    Dim v
    For Each v In secnms
        T = T & v & vbCrLf
    Next
    Dim iarr() As String
    
    ini.GetIniArr iarr, , aSection
    
    'Dim i As Long
    For i = 0 To UBound(iarr)
        v = iarr(i)
        T = T & v & vbCrLf
    Next
    
'    Dim iniSec As ConfigIniSection: Set iniSec = MNew.ConfigIniSection(ini, aSection)
'    Dim kvs As Collection: Set kvs = iniSec.StrKeyValsToCol
'    For Each v In kvs
'        T = T & v & vbCrLf
'    Next
    Text1.Text = T
End Sub

Private Sub BtnTestVBP_Click()
    Dim pfn As PathFileName: Set pfn = MNew.PathFileName(App.Path & "\PConfigIni_vbp.ini")
    If Not pfn.Exists Then
        MsgBox "File does not exist:" & vbCrLf & pfn.Value
    End If
    Dim cid As ConfigIniDocument: Set cid = MNew.ConfigIniDocument(pfn)
    cid.Load
    Dim i As Long, u As Long
    Dim s As String, cikv As ConfigIniKeyValue
    
    Dim classes As ConfigIniSection: Set classes = cid.Root.Filter("Class")
    u = classes.KeyValues.Count - 1
    For i = 0 To u
        Set cikv = classes.KeyValues.Item(i)
        s = s & cikv.Value & vbCrLf
    Next
    
    s = s & vbCrLf
    
    Dim modules As ConfigIniSection: Set modules = cid.Root.Filter("Module")
    u = modules.KeyValues.Count - 1
    For i = 0 To u
        Set cikv = modules.KeyValues.Item(i)
        s = s & cikv.Value & vbCrLf
    Next
    
    s = s & vbCrLf
    
    Dim forms As ConfigIniSection: Set forms = cid.Root.Filter("Form")
    u = forms.KeyValues.Count - 1
    For i = 0 To u
        Set cikv = forms.KeyValues.Item(i)
        s = s & cikv.Value & vbCrLf
    Next
    
    Text1.Text = s

End Sub

Private Sub BtnReadIniFile_Click()
    'read Ini-file and display it
    'Dim IniFile As ConfigIniDocument: Set IniFile = MNew.ConfigIniDocument(IniFileName)
    'Call IniFile.Load
    If Not IniPFN.Exists Then
        If MsgBox("Inifile does not exist, write it first?" & vbCrLf & IniPFN.Value, vbOKCancel) = vbCancel Then Exit Sub
        BtnWriteIniFile_Click
    End If
    If IniFile Is Nothing Then
        Set IniFile = MNew.ConfigIniDocument(IniPFN):
    End If
    IniFile.Load
    Text1.Text = IniFile.ToStr
End Sub

Private Sub BtnWriteIniFile_Click()
    
    Dim Section As ConfigIniSection
    Dim KyValue As ConfigIniKeyValue
    Dim sec     As String
    Dim Key     As String
    Dim val     As String
    
    'directly write some values to the Ini-file
    'by using the functions ValueStr, ValueBol & ValueInt you can
    'immediately write to the Ini-file
    'these function you will find in the class ConfigIniDocument
    'as well as in the class ConfigIniKeyValue
    
    sec = "TestReadWriteAtOnce"
    Key = "FirstEntry"
    IniFile.ValueStr(sec, Key, "") = "NewValueOfFirstEntry"
    
    'read from ini file what we have written:
    val = IniFile.ValueStr(sec, Key, "")
    MsgBox "The read value is: " & val
    
    sec = "TestSection1"
    Set Section = IniFile.AddSection(sec)
    
    Key = "FirstEntry"
    Set KyValue = Section.AddKeyValue(Key)
    KyValue.ValueInt = 123456
    
    Key = "SecondEntry"
    Set KyValue = Section.AddKeyValue(Key)
    KyValue.ValueInt = 456789
    
    'it's also possible to write UD-Type-variables at once:
    Key = "Form1PositionAndSize"
    Set KyValue = Section.AddKeyValue(Key)
    
    Dim cs As PosSizeF: cs = MNew.PosSizeF(Me)
    Dim rv As Long
    KyValue.ValueStructP(LenB(cs), VarPtr(cs)) = VarPtr(cs)
    
    Dim tt As TestTyp1
    With tt
        .BolVal = True
        .IntVal = 12345
        .LngVal = 123456789
        .SngVal = 0.123456
        .DblVal = 0.123456789
        .StrVal = "Test Entry"
    End With
    
    Key = "tt_As_TestTyp"
    Set KyValue = Section.AddKeyValue(Key)
    
    KyValue.ValueStructP(LenB(tt), VarPtr(tt)) = VarPtr(tt)
    
    'write a value yourself
    Key = "MyEntry"
    Set KyValue = Section.AddKeyValue(Key)
    val = InputBox("Write a value yourself: ", "Me too", "hoho")
    If Not (Len(val) = 0) Then
        KyValue.ValueStr = val
    End If
    
End Sub

Private Sub BtnReadRawIniData_Click()
    If IniPFN Is Nothing Then Set IniPFN = MNew.PathFileName(Environ("Temp") & "\Test.ini")
    If Not IniPFN.Exists Then
        MsgBox "File does not exist, write inifile first."
        Exit Sub
    End If
    Text1.Text = IniPFN.ReadAllStr
    IniPFN.CloseFile
End Sub

Private Sub BtnDeleteIniFile_Click()
Try: On Error GoTo Catch
    IniPFN.Delete
Catch:
End Sub

Private Sub BtnWriteWindowPosSize_Click()
    'Dim IniFile As ConfigIniDocument: Set IniFile = MNew.ConfigIniDocument(IniFileName): IniFile.Load
    Dim Section As ConfigIniSection:  Set Section = IniFile.Section("TestSection1")
    
    Dim Key As String: Key = "Form1PositionAndSize"
    Dim KyValue As ConfigIniKeyValue: Set KyValue = Section.AddKeyValue(Key)
    
    Dim cs As PosSizeF: cs = MNew.PosSizeF(Me)
    
    Dim rv As Long: KyValue.ValueStructP(LenB(cs), VarPtr(cs)) = rv
End Sub

Private Sub BtnSetWindowPosSize_Click()
    
    'Dim IniFile As ConfigIniDocument: Set IniFile = MNew.ConfigIniDocument(IniFileName): IniFile.Load
    Dim Section As ConfigIniSection:  Set Section = IniFile.Section("TestSection1")
    
    Dim Key As String: Key = "Form1PositionAndSize"
    Dim KyValue As ConfigIniKeyValue: Set KyValue = Section.AddKeyValue(Key)
    
    Dim cs As PosSizeF
    
    Dim rv As Long: rv = KyValue.ValueStructP(LenB(cs), VarPtr(cs))
    
    With cs
        Me.WindowState = FormWindowStateConstants.vbNormal
        Me.Move .Position.X, .Position.Y, .Size.Width, .Size.Height
    End With
End Sub

Private Sub Form_Resize()
    Dim L As Single
    Dim T As Single: T = Text1.Top
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then Text1.Move L, T, W, H
End Sub
