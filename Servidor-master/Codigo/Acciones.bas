Attribute VB_Name = "Acciones"
'AoYind 3.0.0
'Copyright (C) 2002 M?rquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n?mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C?digo Postal 1900
'Pablo Ignacio M?rquez

Option Explicit

''
' Modulo para manejar las acciones (doble click) de los carteles, puerta, ramitas
'

''
' Ejecuta la accion del doble click
'
' @param UserIndex UserIndex
' @param Map Numero de mapa
' @param X X
' @param Y Y

Sub Accion(ByVal UserIndex As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)
    Dim tempIndex As Integer

    On Error Resume Next
    '?Rango Visi?n? (ToxicWaste)
    If (Abs(UserList(UserIndex).Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(UserList(UserIndex).Pos.X - X) > RANGO_VISION_X) Then
        Exit Sub
    End If

    '?Posicion valida?
    If InMapBounds(map, X, Y) Then
        With UserList(UserIndex)
            If MapData(map).Tile(X, Y).NpcIndex > 0 Then     'Acciones NPCs
                tempIndex = MapData(map).Tile(X, Y).NpcIndex

                'Set the target NPC
                .flags.TargetNPC = tempIndex

                If Npclist(tempIndex).Comercia = 1 Then
                    '?Esta el user muerto? Si es asi no puede comerciar
                    If .flags.Muerto = 1 Then
                        Call WriteConsoleMsg(UserIndex, "??Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If

                    'Is it already in commerce mode??
                    If .flags.Comerciando Then
                        Exit Sub
                    End If

                    If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                        Call WriteConsoleMsg(UserIndex, "Est?s demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If

                    'Iniciamos la rutina pa' comerciar.
                    Call IniciarComercioNPC(UserIndex)

                ElseIf Npclist(tempIndex).NpcType = eNPCType.Banquero Then
                    '?Esta el user muerto? Si es asi no puede comerciar
                    If .flags.Muerto = 1 Then
                        Call WriteConsoleMsg(UserIndex, "??Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If

                    'Is it already in commerce mode??
                    If .flags.Comerciando Then
                        Exit Sub
                    End If

                    If Distancia(Npclist(tempIndex).Pos, .Pos) > 4 Then
                        Call WriteConsoleMsg(UserIndex, "Est?s demasiado lejos del banquero.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If

                    'A depositar de una
                    Call IniciarDeposito(UserIndex)

                ElseIf Npclist(tempIndex).NpcType = eNPCType.Revividor Or Npclist(tempIndex).NpcType = eNPCType.ResucitadorNewbie Then
                    If Distancia(.Pos, Npclist(tempIndex).Pos) > 10 Then
                        Call WriteConsoleMsg(UserIndex, "El sacerdote no puede curarte debido a que estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If

                    'Revivimos si es necesario
                    If .flags.Muerto = 1 And (Npclist(tempIndex).NpcType = eNPCType.Revividor Or EsNewbie(UserIndex)) Then
                        Call RevivirUsuario(UserIndex)
                    End If

                    If Npclist(tempIndex).NpcType = eNPCType.Revividor Or EsNewbie(UserIndex) Then
                        'curamos totalmente
                        .Stats.MinHP = .Stats.MaxHP
                        Call WriteUpdateUserStats(UserIndex)
                    End If
                ElseIf Npclist(tempIndex).NpcType = eNPCType.Marinero Then
                    Call HablaMarinero(UserIndex, tempIndex, True)
                ElseIf Npclist(tempIndex).NpcType = eNPCType.Entrenador Then

                    'Dead users can't use pets
                    If .flags.Muerto = 1 Then
                        Call WriteConsoleMsg(UserIndex, "??Est?s muerto!!", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If

                    'Validate target NPC
                    If .flags.TargetNPC = 0 Then
                        Call WriteConsoleMsg(UserIndex, "Primero ten?s que seleccionar un personaje, hace click izquierdo sobre ?l.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If

                    'Make sure it's close enough
                    If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
                        Call WriteConsoleMsg(UserIndex, "Est?s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If

                    Call WriteTrainerCreatureList(UserIndex, .flags.TargetNPC)
                    
                    
                ElseIf Npclist(MapData(map).Tile(X, Y).NpcIndex).NpcType = eNPCType.Quest Then

                If UserList(UserIndex).flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "Estas Muerto....", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

                Call EnviarQuest(UserIndex)
                End If

         

                '?Es un obj?
            ElseIf MapData(map).Tile(X, Y).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(map).Tile(X, Y).ObjInfo.ObjIndex

                .flags.TargetObj = tempIndex

                Select Case ObjData(tempIndex).OBJType
                Case eOBJType.otPuertas    'Es una puerta
                    Call AccionParaPuerta(map, X, Y, UserIndex)
                Case eOBJType.otCarteles    'Es un cartel
                    Call AccionParaCartel(map, X, Y, UserIndex)
                Case eOBJType.otLe?a    'Le?a
                    If tempIndex = FOGATA_APAG And .flags.Muerto = 0 Then
                        Call AccionParaRamita(map, X, Y, UserIndex)
                    End If
                End Select
                '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
            ElseIf MapData(map).Tile(X + 1, Y).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(map).Tile(X + 1, Y).ObjInfo.ObjIndex
                .flags.TargetObj = tempIndex

                Select Case ObjData(tempIndex).OBJType

                Case eOBJType.otPuertas    'Es una puerta
                    Call AccionParaPuerta(map, X + 1, Y, UserIndex)

                End Select

            ElseIf MapData(map).Tile(X + 1, Y + 1).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(map).Tile(X + 1, Y + 1).ObjInfo.ObjIndex
                .flags.TargetObj = tempIndex

                Select Case ObjData(tempIndex).OBJType
                Case eOBJType.otPuertas    'Es una puerta
                    Call AccionParaPuerta(map, X + 1, Y + 1, UserIndex)
                End Select

            ElseIf MapData(map).Tile(X, Y + 1).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(map).Tile(X, Y + 1).ObjInfo.ObjIndex
                .flags.TargetObj = tempIndex

                Select Case ObjData(tempIndex).OBJType
                Case eOBJType.otPuertas    'Es una puerta
                    Call AccionParaPuerta(map, X, Y + 1, UserIndex)
                End Select
            End If
        End With
    End If
End Sub

Sub AccionParaPuerta(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
On Error Resume Next

If Not (Distance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, X, Y) > 2) Then
    If ObjData(MapData(map).Tile(X, Y).ObjInfo.ObjIndex).Llave = 0 Then
        If ObjData(MapData(map).Tile(X, Y).ObjInfo.ObjIndex).Cerrada = 1 Then
                'Abre la puerta
                If ObjData(MapData(map).Tile(X, Y).ObjInfo.ObjIndex).Llave = 0 Then
                    
                    MapData(map).Tile(X, Y).ObjInfo.ObjIndex = ObjData(MapData(map).Tile(X, Y).ObjInfo.ObjIndex).IndexAbierta
                    
                    Call modSendData.SendToAreaByPos(map, X, Y, PrepareMessageObjectCreate(ObjData(MapData(map).Tile(X, Y).ObjInfo.ObjIndex).GrhIndex, X, Y))
                    
                    'Desbloquea
                    MapData(map).Tile(X, Y).Blocked = 0
                    MapData(map).Tile(X - 1, Y).Blocked = 0
                    
                    'Bloquea todos los mapas
                    Call Bloquear(True, UserIndex, X, Y, 0)
                    Call Bloquear(True, UserIndex, X - 1, Y, 0)
                    
                      
                    'Sonido
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PUERTA, X, Y))
                    
                Else
                     Call WriteConsoleMsg(UserIndex, "La puerta esta cerrada con llave.", FontTypeNames.FONTTYPE_INFO)
                End If
        Else
                'Cierra puerta
                MapData(map).Tile(X, Y).ObjInfo.ObjIndex = ObjData(MapData(map).Tile(X, Y).ObjInfo.ObjIndex).IndexCerrada
                
                Call modSendData.SendToAreaByPos(map, X, Y, PrepareMessageObjectCreate(ObjData(MapData(map).Tile(X, Y).ObjInfo.ObjIndex).GrhIndex, X, Y))
                                
                MapData(map).Tile(X, Y).Blocked = 1
                MapData(map).Tile(X - 1, Y).Blocked = 1
                
                
                Call Bloquear(True, map, X - 1, Y, 1)
                Call Bloquear(True, map, X, Y, 1)
                
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PUERTA, X, Y))
        End If
        
        UserList(UserIndex).flags.TargetObj = MapData(map).Tile(X, Y).ObjInfo.ObjIndex
    Else
        Call WriteConsoleMsg(UserIndex, "La puerta esta cerrada con llave.", FontTypeNames.FONTTYPE_INFO)
    End If
Else
    Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
End If

End Sub

Sub AccionParaCartel(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
On Error Resume Next

If ObjData(MapData(map).Tile(X, Y).ObjInfo.ObjIndex).OBJType = 8 Then
  
  If Len(ObjData(MapData(map).Tile(X, Y).ObjInfo.ObjIndex).texto) > 0 Then
    Call WriteShowSignal(UserIndex, MapData(map).Tile(X, Y).ObjInfo.ObjIndex)
  End If
  
End If

End Sub
Public Sub CheckFogatasLluvia()
If Lloviendo Then
    LimpiarMundo
End If
End Sub
Sub AccionParaRamita(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
On Error Resume Next

Dim Suerte As Byte
Dim exito As Byte
Dim Obj As Obj

Dim Pos As WorldPos
Pos.map = map
Pos.X = X
Pos.Y = Y

If Distancia(Pos, UserList(UserIndex).Pos) > 2 Then
    Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If Lloviendo And Intemperie(UserIndex) Then
    Call WriteConsoleMsg(UserIndex, "No puedes hacer una fogata mientras este lloviendo.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If MapData(map).Tile(X, Y).Trigger = eTrigger.ZONASEGURA Or Zonas(UserList(UserIndex).zona).Segura = 1 Then
    Call WriteConsoleMsg(UserIndex, "En puedes hacer fogatas en una zona segura.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If UserList(UserIndex).Stats.UserSkills(Supervivencia) > 1 And UserList(UserIndex).Stats.UserSkills(Supervivencia) < 6 Then
            Suerte = 5
ElseIf UserList(UserIndex).Stats.UserSkills(Supervivencia) >= 6 And UserList(UserIndex).Stats.UserSkills(Supervivencia) <= 10 Then
            Suerte = 4
ElseIf UserList(UserIndex).Stats.UserSkills(Supervivencia) >= 10 And UserList(UserIndex).Stats.UserSkills(Supervivencia) Then
            Suerte = 3
End If

exito = RandomNumber(1, Suerte)

If exito = 1 Then
        Obj.ObjIndex = FOGATA
        Obj.Amount = 1
        
        Call WriteConsoleMsg(UserIndex, "Has prendido la fogata.", FontTypeNames.FONTTYPE_INFO)
        
        Call MakeObj(Obj, map, X, Y)
        
        'Las fogatas prendidas se deben eliminar
        Dim Fogatita As New cGarbage
        Fogatita.map = map
        Fogatita.X = X
        Fogatita.Y = Y
        Call TrashCollector.Add(Fogatita)
Else
    Call WriteConsoleMsg(UserIndex, "No has podido hacer fuego.", FontTypeNames.FONTTYPE_INFO)
End If

Call SubirSkill(UserIndex, Supervivencia)

End Sub
