Attribute VB_Name = "modVarios"
Option Explicit

Public Sub Chiquitolin(ByVal UserIndex As Integer, ByVal Chiquito As Boolean)
    UserList(UserIndex).flags.Chiquito = Chiquito
    UserList(UserIndex).Counters.Chiquito = 1
    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChiquito(UserList(UserIndex).Char.CharIndex, Chiquito))
End Sub

Public Function CanUse(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByRef motivo As String)

    With UserList(UserIndex)

        Select Case ObjData(ObjIndex).CanUse
            Case eCanUse.WithGaleon
                If .flags.Navegando = 0 Then
                    motivo = "Debes estar navegando para usar el objeto."
                    CanUse = False
                    Exit Function
                End If
                If .Char.Body <> iGaleon And .Char.Body <> iGaleonCiuda And .Char.Body <> iGaleonCaos Then
                    motivo = "Solo puedes usar este objeto encontrandote navegando en un Galeón."
                    CanUse = False
                    Exit Function
                End If
        End Select
        
        CanUse = True

    End With

End Function

Public Sub HitArea(ByVal UserIndex As Integer, ByVal AreaTiles As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)

    Dim HitArea As Integer

    Dim loopX As Integer
    Dim loopY As Integer
    Dim SourceX As Integer
    Dim SourceY As Integer
    Dim DestX As Integer
    Dim DestY As Integer
    
    SourceX = X - AreaTiles
    SourceY = Y - AreaTiles
    DestX = X + AreaTiles
    DestY = Y + AreaTiles
    
    
     For loopX = SourceX To DestX
        For loopY = SourceY To DestY
            If MapData(map).Tile(loopX, loopY).UserIndex > 0 Then
                Call HitAreaUser(UserIndex, MapData(map).Tile(loopX, loopY).UserIndex, GetHitArea(X, Y, loopX, loopY))
            End If
            If MapData(map).Tile(loopX, loopY).NpcIndex > 0 Then
                Call HitAreaNpc(UserIndex, MapData(map).Tile(loopX, loopY).NpcIndex, GetHitArea(X, Y, loopX, loopY))
            End If
        Next loopY
    Next loopX

End Sub

Public Sub HitAreaUser(ByVal UserIndex As Integer, ByVal TargetUserIndex As Integer, ByVal HitArea As Byte)
    Call UsuarioAtacaUsuario(UserIndex, TargetUserIndex, HitArea)
End Sub

Public Sub HitAreaNpc(ByVal UserIndex As Integer, ByVal TargetNpcIndex As Integer, ByVal HitArea As Byte)
    Call UsuarioAtacaNpc(UserIndex, TargetNpcIndex, HitArea)
End Sub

Public Function GetHitArea(ByVal targetX As Integer, ByVal targetY As Integer, ByVal X As Integer, ByVal Y As Integer) As Byte
    GetHitArea = (Abs(targetX - X) + Abs(targetY - Y))
End Function

