Attribute VB_Name = "MGlobalErrHandler"
Option Explicit

Public Function GlobalErrHandler(aObj As Object, _
                                 aProcName As String, _
                                 Info As String, _
                                 Optional decor As VbMsgBoxStyle = vbCritical _
                                 ) As VbMsgBoxResult
    Dim mess As String
    mess = "In " & TypeName(aObj) & "::" & aProcName & " ist ein Fehler aufgetreten."
    GlobalErrHandler = MsgBox(mess, decor)
End Function
