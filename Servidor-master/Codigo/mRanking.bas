Attribute VB_Name = "mRanking"
Option Explicit

Public Const MAX_TOP As Byte = 10
Public Const MAX_RANKINGS As Byte = 7

Public Type tRanking
    value(1 To MAX_TOP) As Long
    Nombre(1 To MAX_TOP) As String
End Type

Public Ranking(1 To MAX_RANKINGS) As tRanking

Public Enum eRanking
    TopFrags = 1
    TopTorneos = 2
    TopLevel = 3
    TopOro = 4
    TopRetos = 5
    TopClanes = 6
    TopMuertesP = 7
End Enum



Public Function RenameRanking(ByVal Ranking As eRanking) As String


'@ Devolvemos el nombre del TAG [] del archivo .DAT
    Select Case Ranking
    Case eRanking.TopClanes
        RenameRanking = "Criminales Matados"
    Case eRanking.TopFrags
        RenameRanking = "Usuarios Matados"
    Case eRanking.TopLevel
        RenameRanking = "Nivel"
    Case eRanking.TopOro
        RenameRanking = "Oro"
    Case eRanking.TopRetos
        RenameRanking = "Retos"
    Case eRanking.TopTorneos
        RenameRanking = "Torneos"
    Case eRanking.TopMuertesP
        RenameRanking = "Muertes Propias"
    Case Else
        RenameRanking = vbNullString
    End Select
End Function
Public Function RenameValue(ByVal UserIndex As Integer, ByVal Ranking As eRanking) As Long
' @ Devolvemos a que hace referencia el ranking
    With UserList(UserIndex)
        Select Case Ranking
        Case eRanking.TopClanes
            RenameValue = .Faccion.CriminalesMatados
            'RenameValue = guilds(.GuildIndex).Puntos
        Case eRanking.TopFrags
            RenameValue = .Stats.UsuariosMatados
        Case eRanking.TopLevel
            RenameValue = .Stats.ELV
        Case eRanking.TopOro
            RenameValue = .Stats.GLD
        Case eRanking.TopMuertesP
            RenameValue = .Stats.MuertesPropias

        Case eRanking.TopRetos
            RenameValue = .RetosGanados
            ' Case eRanking.TopTorneos
            ' RenameValue = .Stats.TorneosGanados
        End Select
    End With
End Function

Public Sub LoadRanking()
    ' @ Cargamos los rankings
   
    Dim LoopI As Integer
    Dim loopX As Integer
    Dim ln As String
   
    For loopX = 1 To MAX_RANKINGS
        For LoopI = 1 To MAX_TOP
            ln = GetVar(DatPath & "Ranking.Dat", RenameRanking(loopX), "Top" & LoopI)
            Ranking(loopX).Nombre(LoopI) = ReadField(1, ln, 45)
            Ranking(loopX).value(LoopI) = val(ReadField(2, ln, 45))
        Next LoopI
    Next loopX
   
End Sub
   
Public Sub SaveRanking(ByVal rank As eRanking)
' @ Guardamos el ranking

    Dim LoopI As Integer
   
        For LoopI = 1 To MAX_TOP
            Call WriteVar(DatPath & "Ranking.Dat", RenameRanking(rank), _
                "Top" & LoopI, Ranking(rank).Nombre(LoopI) & "-" & Ranking(rank).value(LoopI))
        Next LoopI
End Sub

Public Sub CheckRankingUser(ByVal UserIndex As Integer, ByVal rank As eRanking)
    ' @ Desde aca nos hacemos la siguientes preguntas
    ' @ El personaje está en el ranking?
    ' @ El personaje puede ingresar al ranking?
   
    Dim loopX As Integer
    Dim LoopY As Integer
    Dim loopZ As Integer
    Dim i As Integer
    Dim value As Long
    Dim Actualizacion As Byte
    Dim Auxiliar As String
    Dim PosRanking As Byte
   
    With UserList(UserIndex)
       
        ' @ Not gms
        If EsGM(UserIndex) Then Exit Sub
       
        value = RenameValue(UserIndex, rank)
       
        ' @ Buscamos al personaje en el ranking
        For i = 1 To MAX_TOP
            If Ranking(rank).Nombre(i) = UCase$(.Name) Then
                PosRanking = i
                Exit For
            End If
        Next i
       
        ' @ Si el personaje esta en el ranking actualizamos los valores.
        If PosRanking <> 0 Then
            ' ¿Si está actualizado pa que?
            If value <> Ranking(rank).value(PosRanking) Then
                Call ActualizarPosRanking(PosRanking, rank, value)
               
               
                ' ¿Es la pos 1? No hace falta ordenarlos
                If Not PosRanking = 1 Then
                    ' @ Chequeamos los datos para actualizar el ranking
                    For LoopY = 1 To MAX_TOP
                        For loopZ = 1 To MAX_TOP - LoopY
                               
                            If Ranking(rank).value(loopZ) < Ranking(rank).value(loopZ + 1) Then
                               
                                ' Actualizamos el valor
                                Auxiliar = Ranking(rank).value(loopZ)
                                Ranking(rank).value(loopZ) = Ranking(rank).value(loopZ + 1)
                                Ranking(rank).value(loopZ + 1) = Auxiliar
                               
                                ' Actualizamos el nombre
                                Auxiliar = Ranking(rank).Nombre(loopZ)
                                Ranking(rank).Nombre(loopZ) = Ranking(rank).Nombre(loopZ + 1)
                                Ranking(rank).Nombre(loopZ + 1) = Auxiliar
                                Actualizacion = 1
                            End If
                        Next loopZ
                    Next LoopY
                End If
                   
                If Actualizacion <> 0 Then
                    Call SaveRanking(rank)
                End If
            End If
           
            Exit Sub
        End If
       
        ' @ Nos fijamos si podemos ingresar al ranking
        For loopX = 1 To MAX_TOP
            If value > Ranking(rank).value(loopX) Then
                Call ActualizarRanking(loopX, rank, .Name, value)
                Exit For
            End If
        Next loopX
       
    End With
End Sub

Public Sub ActualizarPosRanking(ByVal Top As Byte, ByVal rank As eRanking, ByVal value As Long)
    ' @ Actualizamos la pos indicada en caso de que el personaje esté en el ranking
    Dim loopX As Integer

    With Ranking(rank)
       
        .value(Top) = value
    End With
End Sub
Public Sub ActualizarRanking(ByVal Top As Byte, ByVal rank As eRanking, ByVal UserName As String, ByVal value As Long)
   
    '@ Actualizamos la lista de ranking
   
    Dim LoopC As Integer
    Dim i As Integer
    Dim j As Integer
    Dim Valor(1 To MAX_TOP) As Long
    Dim Nombre(1 To MAX_TOP) As String
   
    ' @ Copia necesaria para evitar que se dupliquen repetidamente
    For LoopC = 1 To MAX_TOP
        Valor(LoopC) = Ranking(rank).value(LoopC)
        Nombre(LoopC) = Ranking(rank).Nombre(LoopC)
    Next LoopC
   
    ' @ Corremos las pos, desde el "Top" que es la primera
    For LoopC = Top To MAX_TOP - 1
        Ranking(rank).value(LoopC + 1) = Valor(LoopC)
        Ranking(rank).Nombre(LoopC + 1) = Nombre(LoopC)
    Next LoopC


   
    Ranking(rank).Nombre(Top) = UCase$(UserName)
    Ranking(rank).value(Top) = value
    Call SaveRanking(rank)
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Ranking de " & RenameRanking(rank) & "»" & UserName & " ha subido al TOP " & Top & ".", FontTypeNames.FONTTYPE_GUILD))
End Sub

