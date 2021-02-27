VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnWriteWindowPosSize 
      Caption         =   "Write PosAndSize of window"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   3135
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
      Height          =   4335
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   4
      Top             =   1560
      Width           =   6975
   End
   Begin VB.CommandButton BtnSetWindowPosSize 
      Caption         =   "Read and set PosAndSize of window"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   1080
      Width           =   3135
   End
   Begin VB.CommandButton BtnReadIniFile 
      Caption         =   "Read Ini-file"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   600
      Width           =   3135
   End
   Begin VB.CommandButton BtnWriteIniFile 
      Caption         =   "Write Ini-file"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public IniFileName As String
Private IniFile As ConfigIniDocument

Private Sub Form_Load()
    IniFileName = Environ("Temp") & "\Test.ini"
    Set IniFile = MNew.ConfigIniDocument(IniFileName):
    IniFile.Load
    Label1.Caption = IniFileName
End Sub

Private Sub BtnWriteIniFile_Click()
    
    Dim section As ConfigIniSection
    Dim KyValue As ConfigIniKeyValue
    Dim sec     As String
    Dim key     As String
    Dim val     As String
    
    'directly write some values to the Ini-file
    'by using the functions ValueStr, ValueBol & ValueInt you can
    'immediately write to the Ini-file
    'these function you will find in the class ConfigIniDocument
    'as well as in the class ConfigIniKeyValue
    
    sec = "TestReadWriteAtOnce"
    key = "FirstEntry"
    IniFile.ValueStr(sec, key, "") = "NewValueOfFirstEntry"
    
    'read from ini file what we have written:
    val = IniFile.ValueStr(sec, key, "")
    MsgBox "The read value is: " & val
    
    sec = "TestSection1"
    Set section = IniFile.AddSection(sec)
    
    key = "FirstEntry"
    Set KyValue = section.AddKey(key)
    KyValue.ValueInt = 123456
    
    key = "SecondEntry"
    Set KyValue = section.AddKey(key)
    KyValue.ValueInt = 456789
    
    'it's also possible to write UD-Type-variables at once:
    key = "Form1PositionAndSize"
    Set KyValue = section.AddKey(key)
    
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
    
    key = "tt_As_TestTyp"
    Set KyValue = section.AddKey(key)
    
    KyValue.ValueStructP(LenB(tt), VarPtr(tt)) = VarPtr(tt)
    
    'write a value yourself
    key = "MyEntry"
    Set KyValue = section.AddKey(key)
    val = InputBox("Write a value yourself: ", "Me too", "hoho")
    If Not (Len(val) = 0) Then
        KyValue.ValueStr = val
    End If
    
End Sub

Private Sub BtnReadIniFile_Click()
    'read Ini-file and show it
    'Dim IniFile As ConfigIniDocument: Set IniFile = MNew.ConfigIniDocument(IniFileName)
    'Call IniFile.Load
    Text1.Text = IniFile.ToStr
End Sub

Private Sub BtnWriteWindowPosSize_Click()
    'Dim IniFile As ConfigIniDocument: Set IniFile = MNew.ConfigIniDocument(IniFileName): IniFile.Load
    Dim section As ConfigIniSection:  Set section = IniFile.section("TestSection1")
    
    Dim key As String: key = "Form1PositionAndSize"
    Dim KyValue As ConfigIniKeyValue: Set KyValue = section.AddKey(key)
    
    Dim cs As PosSizeF: cs = MNew.PosSizeF(Me)
    
    Dim rv As Long: KyValue.ValueStructP(LenB(cs), VarPtr(cs)) = rv
End Sub

Private Sub BtnSetWindowPosSize_Click()
    
    'Dim IniFile As ConfigIniDocument: Set IniFile = MNew.ConfigIniDocument(IniFileName): IniFile.Load
    Dim section As ConfigIniSection:  Set section = IniFile.section("TestSection1")
    
    Dim key As String: key = "Form1PositionAndSize"
    Dim KyValue As ConfigIniKeyValue: Set KyValue = section.AddKey(key)
    
    Dim cs As PosSizeF
    
    Dim rv As Long: rv = KyValue.ValueStructP(LenB(cs), VarPtr(cs))
    
    With cs
        Me.Move .Position.X, .Position.Y, .Size.Width, .Size.Height
    End With
End Sub

Private Sub Form_Resize()
    Dim l As Single, T As Single, W As Single, H As Single
    T = Text1.Top
    W = Me.ScaleWidth: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then Text1.Move l, T, W, H
End Sub
