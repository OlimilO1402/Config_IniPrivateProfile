VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   2280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4335
   LinkTopic       =   "Form2"
   ScaleHeight     =   2280
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows-Standard
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    'die Inidatei einlesen und anzeigen
    Dim IniFile As ConfigIniDocument
    Set IniFile = MNew.ConfigIniDocument(Form1.IniFileName)
    Call IniFile.Load
    MsgBox IniFile.ToString
End Sub

Private Sub Command2_Click()

End Sub
