Attribute VB_Name = "modInvocaciones"
Private mapX_1 As Integer
Private mapY_1 As Integer

Private mapX_2 As Integer
Private mapY_2 As Integer

Private mapX_3 As Integer
Private mapY_3 As Integer

Private mapX_4 As Integer
Private mapY_4 As Integer

Private Const ZONA_INVOCACION_CLAN As Byte = 114
Private Const ZONA_INVOCACION_CLAN_NPC As Integer = 909

Public Sub InvocarCriaturaMisteriosa(ByVal UserIndex As Integer)

Dim IDZona As Integer
IDZona = UserList(UserIndex).zona

'Si el Npc ya esta invocado salimos
'If Zonas(IDZona).NpcInvocadoIndex > 0 Then
        'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Invocaciones> Alguien ha intentado invocar una criatura misteriosa en " & Zonas(IDZona).Nombre & ".", FontTypeNames.FONTTYPE_SERVER))
        'Call WriteConsoleMsg(UserIndex, "La criatura ya ha sido invocada", FontTypeNames.FONTTYPE_INFO)
'    Exit Sub
'End If

Dim NpcInvocacion As Integer
Dim PosNpcInvocacion As WorldPos

Select Case IDZona
    Case ZONA_INVOCACION_CLAN
        mapX_1 = 532 'X del Usuario 1
        mapY_1 = 481 'Y del Usuario 1
        
        mapX_2 = 524 'X del Usuario 2
        mapY_2 = 486 'Y del Usuario 2
        
        mapX_3 = 540 'X del Usuario 3
        mapY_3 = 486 'Y del Usuario 3
        
        mapX_4 = 532 'X del Usuario 4
        mapY_4 = 490 'Y del Usuario 4
        
        NpcInvocacion = ZONA_INVOCACION_CLAN_NPC
        
        'Coordenadas donde hace Spawn el NPC_Invocado
        PosNpcInvocacion.map = Zonas(IDZona).Mapa
        PosNpcInvocacion.X = 532
        PosNpcInvocacion.Y = 486
    Case Else
        Exit Sub
End Select
'
'If MapData(Zonas(IDZona).Mapa).Tile(mapX_1, mapY_1).UserIndex > 0 Then
'     If UserList(MapData(Zonas(IDZona).Mapa).Tile(mapX_1, mapY_1).UserIndex).flags.Muerto = 0 Then

 'Comprobamos que todas las ubicaciones esten presionadas
 If MapData(Zonas(IDZona).Mapa).Tile(mapX_1, mapY_1).UserIndex > 0 And _
    MapData(Zonas(IDZona).Mapa).Tile(mapX_2, mapY_2).UserIndex > 0 And _
    MapData(Zonas(IDZona).Mapa).Tile(mapX_3, mapY_3).UserIndex > 0 And _
    MapData(Zonas(IDZona).Mapa).Tile(mapX_4, mapY_4).UserIndex > 0 Then

    'Verificamos que no sean caspers
    If UserList(MapData(Zonas(IDZona).Mapa).Tile(mapX_1, mapY_1).UserIndex).flags.Muerto = 0 And _
        UserList(MapData(Zonas(IDZona).Mapa).Tile(mapX_2, mapY_2).UserIndex).flags.Muerto = 0 And _
        UserList(MapData(Zonas(IDZona).Mapa).Tile(mapX_3, mapY_3).UserIndex).flags.Muerto = 0 And _
        UserList(MapData(Zonas(IDZona).Mapa).Tile(mapX_4, mapY_4).UserIndex).flags.Muerto = 0 Then
        
        If Zonas(IDZona).NpcInvocadoIndex > 0 Then
            Call QuitarNPC(Zonas(IDZona).NpcInvocadoIndex)
        End If
            
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Invocaciones> Se ha invocado una criatura misteriosa en " & Zonas(IDZona).Nombre & ".", FontTypeNames.FONTTYPE_SERVER))
        
        Zonas(IDZona).NpcInvocadoIndex = SpawnNpc(NpcInvocacion, PosNpcInvocacion, True, False, UserList(UserIndex).zona)
        Npclist(Zonas(IDZona).NpcInvocadoIndex).Orig = PosNpcInvocacion
        Npclist(Zonas(IDZona).NpcInvocadoIndex).flags.VuelveOrigPos = 1
    End If
    
 End If

End Sub

Public Sub VerificarInvocacion(ByVal NpcIndex As Integer)
    If Zonas(Npclist(NpcIndex).zona).NpcInvocadoIndex = NpcIndex Then
        Zonas(Npclist(NpcIndex).zona).NpcInvocadoIndex = 0 'Ya no tiene NPC Invocado
    End If
End Sub




