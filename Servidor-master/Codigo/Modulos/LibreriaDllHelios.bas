Attribute VB_Name = "LibreriaDllHelios"
Option Explicit

Public Declare Sub Helios Lib "Helios.dll" (ByVal Archivo As String, ByVal Texto As String)
Public Declare Sub Escribe Lib "Helios.dll" (ByVal Archivo As String, ByVal Texto As String)

Public Sub CrearLog(File As String, Data As String, Optional CrLf As Boolean = True)

'Llamo a la dll Helios 18/08/2021
        If FileExist(CarpetaLogs & "\CantidaddeUsuarios.log", vbNormal) Then Kill CarpetaLogs & "\CantidaddeUsuarios.log"
        Call Helios(File, Data)

End Sub
