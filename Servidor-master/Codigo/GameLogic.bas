Attribute VB_Name = "Extra"
'AoYind 3.0.0
'Copyright (C) 2002 Márquez Pablo Ignacio
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
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Public Function EsNewbie(ByVal UserIndex As Integer) As Boolean
    EsNewbie = UserList(UserIndex).Stats.ELV <= LimiteNewbie
End Function

Public Function esArmada(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Autor: Pablo (ToxicWaste)
'Last Modification: 23/01/2007
'***************************************************
    esArmada = (UserList(UserIndex).Faccion.ArmadaReal = 1)
End Function

Public Function esCaos(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Autor: Pablo (ToxicWaste)
'Last Modification: 23/01/2007
'***************************************************
    esCaos = (UserList(UserIndex).Faccion.FuerzasCaos = 1)
End Function

Public Function EsGM(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Autor: Pablo (ToxicWaste)
'Last Modification: 23/01/2007
'***************************************************
    EsGM = (UserList(UserIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero))
End Function
Public Sub ResucitarOCurar(ByVal UserIndex As Integer)
If UserList(UserIndex).flags.Muerto = 1 Then
    Call RevivirUsuario(UserIndex)
    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
        
    Call WriteUpdateHP(UserIndex)
    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(20, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, 35, 1))

    
    Call WriteConsoleMsg(UserIndex, "¡¡Hás sido resucitado!!", FontTypeNames.FONTTYPE_INFO)
ElseIf UserList(UserIndex).Stats.MinHP < UserList(UserIndex).Stats.MaxHP Then
    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
        
    Call WriteUpdateHP(UserIndex)
        
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, 9, 1))
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(18, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

        
    Call WriteConsoleMsg(UserIndex, "¡¡Hás sido curado!!", FontTypeNames.FONTTYPE_INFO)
End If
End Sub
Public Sub DoTileEvents(ByVal UserIndex As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)
'***************************************************
'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 23/01/2007
'Handles the Map passage of Users. Allows the existance
'of exclusive maps for Newbies, Royal Army and Caos Legion members
'and enables GMs to enter every map without restriction.
'Uses: Mapinfo(map).Restringir = "NEWBIE" (newbies), "ARMADA", "CAOS", "FACCION" or "NO".
'***************************************************
    Dim nPos As WorldPos
    Dim mapa As Integer
    Dim xA As Integer
    Dim yA As Integer
    On Error GoTo errhandler

    mapa = map
    xA = X
    yA = Y + 1
    'Controla las salidas
    If InMapBounds(map, X, Y) Then
        With MapData(map).Tile(X, Y)

            If .Trigger = AUTORESU Then
                Call ResucitarOCurar(UserIndex)
            End If

            If .Trigger = TCHIQUITO Then
                If UserList(UserIndex).flags.Chiquito = False Then
                    If Not UserList(UserIndex).flags.UltimoMensaje = 23 Then
                        Call WriteConsoleMsg(UserIndex, "Tu contextura no te permite ingresar por la grieta, al parecer una persona pequeña tal vez podría hacerlo..", FontTypeNames.FONTTYPE_INFO)
                        UserList(UserIndex).flags.UltimoMensaje = 23
                    End If
                    Exit Sub
                End If
            End If

            If .TileExit.map > 0 And .TileExit.map <= NumMaps Then
                If Not EsGM(UserIndex) Then
                    If Zonas(BuscarZona(.TileExit.map, .TileExit.X, .TileExit.Y)).Item <> 0 Then

                        If Zonas(BuscarZona(.TileExit.map, .TileExit.X, .TileExit.Y)).itemDes = 0 Then
                            Call RestringirMapaItem(UserIndex, map, X, Y, Zonas(BuscarZona(.TileExit.map, .TileExit.X, .TileExit.Y)).Nombre, Zonas(BuscarZona(.TileExit.map, .TileExit.X, .TileExit.Y)).Item, mapa, xA, yA, nPos, False)
                        Else
                            Call RestringirMapaItem(UserIndex, map, X, Y, Zonas(BuscarZona(.TileExit.map, .TileExit.X, .TileExit.Y)).Nombre, Zonas(BuscarZona(.TileExit.map, .TileExit.X, .TileExit.Y)).Item, mapa, xA, yA, nPos, True)
                        End If
                        Exit Sub
                    End If
                    If Zonas(BuscarZona(.TileExit.map, .TileExit.X, .TileExit.Y)).RestringirM <> 0 And Zonas(BuscarZona(.TileExit.map, .TileExit.X, .TileExit.Y)).Restringir <> 13 Then
                        Call RestringirMapaMinMax(UserIndex, map, X, Y, Zonas(BuscarZona(.TileExit.map, .TileExit.X, .TileExit.Y)).Nombre, Zonas(BuscarZona(.TileExit.map, .TileExit.X, .TileExit.Y)).Restringir, Zonas(BuscarZona(.TileExit.map, .TileExit.X, .TileExit.Y)).RestringirM, mapa, xA, yA, nPos)
                        Exit Sub
                    End If

                    If Zonas(BuscarZona(.TileExit.map, .TileExit.X, .TileExit.Y)).Restringir <> 0 And Zonas(BuscarZona(.TileExit.map, .TileExit.X, .TileExit.Y)).Restringir <> 13 Then
                        Call RestringirMapa(UserIndex, map, X, Y, Zonas(BuscarZona(.TileExit.map, .TileExit.X, .TileExit.Y)).Nombre, Zonas(BuscarZona(.TileExit.map, .TileExit.X, .TileExit.Y)).Restringir, mapa, xA, yA, nPos)
                        Exit Sub
                    End If
                End If

                '¿Es mapa de newbies?
                If Zonas(BuscarZona(.TileExit.map, .TileExit.X, .TileExit.Y)).Restringir = 13 Then
                    '¿El usuario es un newbie?
                    If EsNewbie(UserIndex) Or EsGM(UserIndex) Then
                        If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(UserIndex)) Then
                            Call WarpUserChar(UserIndex, .TileExit.map, .TileExit.X, .TileExit.Y, True)
                        Else
                            Call ClosestLegalPos(.TileExit, nPos)
                            If nPos.X <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(UserIndex, nPos.map, nPos.X, nPos.Y, True)
                            End If
                        End If
                    Else    'No es newbie
                        Call WriteConsoleMsg(UserIndex, "Mapa exclusivo para newbies.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)

                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.map, nPos.X, nPos.Y, False)
                        End If
                    End If
                Else    'No es un mapa de newbies, ni Armadas, ni Caos, ni faccionario.
                    If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(UserIndex)) Then
                        Call WarpUserChar(UserIndex, .TileExit.map, .TileExit.X, .TileExit.Y, False)
                    Else
                        Call ClosestLegalPos(.TileExit, nPos)
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.map, nPos.X, nPos.Y, True)
                        End If
                    End If
                End If




                'Call RestringirMapaMinMax(UserIndex, map, X, Y, "Dungeon Infierno", 30, 35, mapa, xA, yA, nPos)

                'Te fusite del mapa. La criatura ya no es más tuya ni te reconoce como que vos la atacaste.
                Dim aN As Integer

                aN = UserList(UserIndex).flags.AtacadoPorNpc
                If aN > 0 Then
                    Npclist(aN).Movement = Npclist(aN).flags.OldMovement
                    Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
                    Npclist(aN).flags.AttackedBy = vbNullString
                End If

                aN = UserList(UserIndex).flags.NPCAtacado
                If aN > 0 Then
                    If Npclist(aN).flags.AttackedFirstBy = UserList(UserIndex).Name Then
                        Npclist(aN).flags.AttackedFirstBy = vbNullString
                    End If
                End If
                UserList(UserIndex).flags.AtacadoPorNpc = 0
                UserList(UserIndex).flags.NPCAtacado = 0
            End If

            If Zonas(UserList(UserIndex).zona).TieneNpcInvocacion > 0 Then
                Call InvocarCriaturaMisteriosa(UserIndex)
            End If

        End With
    End If
    Exit Sub

errhandler:
    Call LogError("Error en DotileEvents. Error: " & Err.Number & " - Desc: " & Err.Description)
End Sub
Sub CambiarOrigHeading(ByVal NpcIndex As Integer, ByVal Trigger As Byte)
Select Case Trigger
    Case eTrigger.MIRARLEFT
        Npclist(NpcIndex).origHeading = eHeading.WEST
        Npclist(NpcIndex).Char.heading = eHeading.WEST
    Case eTrigger.MIRARRIGHT
        Npclist(NpcIndex).origHeading = eHeading.EAST
        Npclist(NpcIndex).Char.heading = eHeading.EAST
    Case eTrigger.MIRARUP
        Npclist(NpcIndex).origHeading = eHeading.NORTH
        Npclist(NpcIndex).Char.heading = eHeading.NORTH
End Select
End Sub
Function InRangoVision(ByVal UserIndex As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean

If X > UserList(UserIndex).Pos.X - MargenX And X < UserList(UserIndex).Pos.X + MargenX Then
    If Y > UserList(UserIndex).Pos.Y - MargenY And Y < UserList(UserIndex).Pos.Y + MargenY Then
        InRangoVision = True
        Exit Function
    End If
End If
InRangoVision = False

End Function

Function InRangoVisionNPC(ByVal NpcIndex As Integer, X As Integer, Y As Integer) As Boolean

If X > Npclist(NpcIndex).Pos.X - MargenX And X < Npclist(NpcIndex).Pos.X + MargenX Then
    If Y > Npclist(NpcIndex).Pos.Y - MargenY And Y < Npclist(NpcIndex).Pos.Y + MargenY Then
        InRangoVisionNPC = True
        Exit Function
    End If
End If
InRangoVisionNPC = False

End Function


Function InMapBounds(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
            
If (map <= 0 Or map > NumMaps) Or X < 1 Or X > MapInfo(map).Width Or Y < 1 Or Y > MapInfo(map).Height Then
    InMapBounds = False
Else
    InMapBounds = True
End If

End Function

Sub ClosestLegalPos(Pos As WorldPos, ByRef nPos As WorldPos, Optional PuedeAgua As Boolean = False, Optional PuedeTierra As Boolean = True)
'*****************************************************************
'Author: Unknown (original version)
'Last Modification: 24/01/2007 (ToxicWaste)
'Encuentra la posicion legal mas cercana y la guarda en nPos
'*****************************************************************

Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Long
Dim tY As Long

nPos.map = Pos.map

Do While Not LegalPos(Pos.map, nPos.X, nPos.Y, PuedeAgua, PuedeTierra)
    If LoopC > 12 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = Pos.Y - LoopC To Pos.Y + LoopC
        For tX = Pos.X - LoopC To Pos.X + LoopC
            
            If LegalPos(nPos.map, tX, tY, PuedeAgua, PuedeTierra) Then
                nPos.X = tX
                nPos.Y = tY
                '¿Hay objeto?
                
                tX = Pos.X + LoopC
                tY = Pos.Y + LoopC
  
            End If
        
        Next tX
    Next tY
    
    LoopC = LoopC + 1
    
Loop

If Notfound = True Then
    nPos.X = 0
    nPos.Y = 0
End If

End Sub

Public Sub ClosestStablePos(Pos As WorldPos, ByRef nPos As WorldPos)
'*****************************************************************
'Encuentra la posicion legal mas cercana que no sea un portal y la guarda en nPos
'*****************************************************************

Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Long
Dim tY As Long

nPos.map = Pos.map

Do While Not LegalPos(Pos.map, nPos.X, nPos.Y)
    If LoopC > 12 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = Pos.Y - LoopC To Pos.Y + LoopC
        For tX = Pos.X - LoopC To Pos.X + LoopC
            
            If LegalPos(nPos.map, tX, tY) And MapData(nPos.map).Tile(tX, tY).TileExit.map = 0 Then
                nPos.X = tX
                nPos.Y = tY
                '¿Hay objeto?
                
                tX = Pos.X + LoopC
                tY = Pos.Y + LoopC
  
            End If
        
        Next tX
    Next tY
    
    LoopC = LoopC + 1
    
Loop

If Notfound = True Then
    nPos.X = 0
    nPos.Y = 0
End If

End Sub

Function NameIndex(ByVal Name As String) As Integer
    Dim UserIndex As Long
    
    '¿Nombre valido?
    If LenB(Name) = 0 Then
        NameIndex = 0
        Exit Function
    End If
    
    If InStrB(Name, "+") <> 0 Then
        Name = UCase$(Replace(Name, "+", " "))
    End If
    
    UserIndex = 1
    Do Until UCase$(UserList(UserIndex).Name) = UCase$(Name)
        
        UserIndex = UserIndex + 1
        
        If UserIndex > MaxUsers Then
            NameIndex = 0
            Exit Function
        End If
    Loop
     
    NameIndex = UserIndex
End Function

Function CheckForSameIP(ByVal UserIndex As Integer, ByVal UserIP As String) As Boolean
    Dim LoopC As Long
    
    For LoopC = 1 To MaxUsers
        If UserList(LoopC).flags.UserLogged = True Then
            If UserList(LoopC).ip = UserIP And UserIndex <> LoopC Then
                CheckForSameIP = True
                Exit Function
            End If
        End If
    Next LoopC
    
    CheckForSameIP = False
End Function

Function CheckForSameName(ByVal Name As String) As Boolean
'Controlo que no existan usuarios con el mismo nombre
    Dim LoopC As Long
    
    For LoopC = 1 To LastUser
        If UserList(LoopC).flags.UserLogged Then
            
            'If UCase$(UserList(LoopC).Name) = UCase$(Name) And UserList(LoopC).ConnID <> -1 Then
            'OJO PREGUNTAR POR EL CONNID <> -1 PRODUCE QUE UN PJ EN DETERMINADO
            'MOMENTO PUEDA ESTAR LOGUEADO 2 VECES (IE: CIERRA EL SOCKET DESDE ALLA)
            'ESE EVENTO NO DISPARA UN SAVE USER, LO QUE PUEDE SER UTILIZADO PARA DUPLICAR ITEMS
            'ESTE BUG EN ALKON PRODUJO QUE EL SERVIDOR ESTE CAIDO DURANTE 3 DIAS. ATENTOS.
            
            If UCase$(UserList(LoopC).Name) = UCase$(Name) Then
                CheckForSameName = True
                Exit Function
            End If
        End If
    Next LoopC
    
    CheckForSameName = False
End Function

Sub HeadtoPos(ByVal Head As eHeading, ByRef Pos As WorldPos)
'*****************************************************************
'Toma una posicion y se mueve hacia donde esta perfilado
'*****************************************************************
    Select Case Head
        Case eHeading.NORTH
            Pos.Y = Pos.Y - 1
        
        Case eHeading.SOUTH
            Pos.Y = Pos.Y + 1
        
        Case eHeading.EAST
            Pos.X = Pos.X + 1
        
        Case eHeading.WEST
            Pos.X = Pos.X - 1
    End Select
End Sub



Function LegalPos(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True) As Boolean
'***************************************************
'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 23/01/2007
'Checks if the position is Legal.
'***************************************************
'¿Es un mapa valido?
If (map <= 0 Or map > NumMaps) Or (map = 1 And (X < 1 Or X > MapInfo(map).Width Or Y < 1 Or Y > MapInfo(map).Height)) Then
            LegalPos = False
ElseIf (X < 1 Or X > MapInfo(map).Width Or Y < 1 Or Y > MapInfo(map).Height) Then
    LegalPos = False
Else
    If PuedeAgua And PuedeTierra Then
        LegalPos = (MapData(map).Tile(X, Y).Blocked <> 1) And _
                   (MapData(map).Tile(X, Y).UserIndex = 0) And _
                   (MapData(map).Tile(X, Y).NpcIndex = 0)
    ElseIf PuedeTierra And Not PuedeAgua Then
        LegalPos = (MapData(map).Tile(X, Y).Blocked <> 1) And _
                   (MapData(map).Tile(X, Y).UserIndex = 0) And _
                   (MapData(map).Tile(X, Y).NpcIndex = 0) And _
                   (Not HayAgua(map, X, Y))
    ElseIf PuedeAgua And Not PuedeTierra Then
        LegalPos = (MapData(map).Tile(X, Y).Blocked <> 1) And _
                   (MapData(map).Tile(X, Y).UserIndex = 0) And _
                   (MapData(map).Tile(X, Y).NpcIndex = 0) And _
                   (HayAgua(map, X, Y))
    Else
        LegalPos = False
    End If
   
End If

End Function

Function MoveToLegalPos(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True) As Boolean
'***************************************************
'Autor: ZaMa
'Last Modification: 26/03/2009
'Checks if the position is Legal, but considers that if there's a casper, it's a legal movement.
'***************************************************

Dim UserIndex As Integer
Dim IsDeadChar As Boolean


'¿Es un mapa valido?
If (map <= 0 Or map > NumMaps) Or (map = 1 And (X < 1 Or X > MapInfo(map).Width Or Y < 1 Or Y > MapInfo(map).Height)) Then
    MoveToLegalPos = False
ElseIf (X < 1 Or X > MapInfo(map).Width Or Y < 1 Or Y > MapInfo(map).Height) Then
    MoveToLegalPos = False
Else
        UserIndex = MapData(map).Tile(X, Y).UserIndex
        If UserIndex > 0 Then
            IsDeadChar = UserList(UserIndex).flags.Muerto = 1
        Else
            IsDeadChar = False
        End If
    
    If PuedeAgua And PuedeTierra Then
        MoveToLegalPos = (MapData(map).Tile(X, Y).Blocked <> 1) And _
                   (UserIndex = 0 Or IsDeadChar) And _
                   (MapData(map).Tile(X, Y).NpcIndex = 0)
    ElseIf PuedeTierra And Not PuedeAgua Then
        MoveToLegalPos = (MapData(map).Tile(X, Y).Blocked <> 1) And _
                   (UserIndex = 0 Or IsDeadChar) And _
                   (MapData(map).Tile(X, Y).NpcIndex = 0) And _
                   (Not HayAgua(map, X, Y))
    ElseIf PuedeAgua And Not PuedeTierra Then
        MoveToLegalPos = (MapData(map).Tile(X, Y).Blocked <> 1) And _
                   (UserIndex = 0 Or IsDeadChar) And _
                   (MapData(map).Tile(X, Y).NpcIndex = 0) And _
                   (HayAgua(map, X, Y))
    Else
        MoveToLegalPos = False
    End If
  
End If


End Function
Public Sub FindLegalPos(ByVal UserIndex As Integer, ByVal map As Integer, ByRef X As Integer, ByRef Y As Integer)
'***************************************************
'Autor: ZaMa
'Last Modification: 26/03/2009
'Search for a Legal pos for the user who is being teleported.
'***************************************************
If InMapBounds(map, X, Y) Then
    If MapData(map).Tile(X, Y).UserIndex <> 0 Or _
        MapData(map).Tile(X, Y).NpcIndex <> 0 Then
                    
        ' Se teletransporta a la misma pos a la que estaba
        If MapData(map).Tile(X, Y).UserIndex = UserIndex Then Exit Sub
                            
        Dim FoundPlace As Boolean
        Dim tX As Long
        Dim tY As Long
        Dim Rango As Long
        Dim OtherUserIndex As Integer
    
        For Rango = 1 To 5
            For tY = Y - Rango To Y + Rango
                For tX = X - Rango To X + Rango
                    'Reviso que no haya User ni NPC
                    If InMapBounds(map, tX, tY) Then
                        If MapData(map).Tile(tX, tY).UserIndex = 0 And _
                            MapData(map).Tile(tX, tY).NpcIndex = 0 Then
                            
                             FoundPlace = True
                            
                            Exit For
                        End If
                    End If

                Next tX
        
                If FoundPlace Then _
                    Exit For
            Next tY
            
            If FoundPlace Then _
                    Exit For
        Next Rango

    
        If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
            X = tX
            Y = tY
        Else
            'Muy poco probable, pero..
            'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
            OtherUserIndex = MapData(map).Tile(X, Y).UserIndex
            If OtherUserIndex <> 0 Then
                'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
                If UserList(OtherUserIndex).ComUsu.DestUsu > 0 Then
                    'Le avisamos al que estaba comerciando que se tuvo que ir.
                    If UserList(UserList(OtherUserIndex).ComUsu.DestUsu).flags.UserLogged Then
                        Call FinComerciarUsu(UserList(OtherUserIndex).ComUsu.DestUsu)
                        Call WriteConsoleMsg(UserList(OtherUserIndex).ComUsu.DestUsu, "Comercio cancelado. El otro usuario se ha desconectado.", FontTypeNames.FONTTYPE_TALK)
                        Call FlushBuffer(UserList(OtherUserIndex).ComUsu.DestUsu)
                    End If
                    'Lo sacamos.
                    If UserList(OtherUserIndex).flags.UserLogged Then
                        Call FinComerciarUsu(OtherUserIndex)
                        Call WriteErrorMsg(OtherUserIndex, "Alguien se ha conectado donde te encontrabas, por favor reconéctate...")
                        Call FlushBuffer(OtherUserIndex)
                    End If
                End If
            
                Call CloseSocket(OtherUserIndex)
            End If
        End If
    End If
Else
    FoundPlace = False
End If
End Sub

Public Sub FindLegalPosComplete(ByVal UserIndex As Integer, ByVal map As Integer, ByRef X As Integer, ByRef Y As Integer, Optional ByVal PuedeTierra As Boolean = True, Optional ByVal PuedeAgua As Boolean = False)
'***************************************************
'Autor: ZaMa
'Last Modification: 26/03/2009
'Search for a Legal pos for the user who is being teleported.
'***************************************************

    If MapData(map).Tile(X, Y).UserIndex <> 0 Or _
        MapData(map).Tile(X, Y).NpcIndex <> 0 Then
                    
        ' Se teletransporta a la misma pos a la que estaba
        If MapData(map).Tile(X, Y).UserIndex = UserIndex Then Exit Sub
                            
        Dim FoundPlace As Boolean
        Dim tX As Long
        Dim tY As Long
        Dim Rango As Long
        Dim OtherUserIndex As Integer
    
        For Rango = 1 To 5
            For tY = Y - Rango To Y + Rango
                For tX = X - Rango To X + Rango
                    'Reviso que no haya User ni NPC
                    If InMapBounds(map, tX, tY) Then
                        If MapData(map).Tile(tX, tY).UserIndex = 0 And _
                            MapData(map).Tile(tX, tY).NpcIndex = 0 Then
                            
                            If LegalPos(map, tX, tY, PuedeAgua, PuedeTierra) Then FoundPlace = True
                            
                            Exit For
                        End If
                    End If

                Next tX
        
                If FoundPlace Then _
                    Exit For
            Next tY
            
            If FoundPlace Then _
                    Exit For
        Next Rango

    
        If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
            X = tX
            Y = tY
        Else
            'Muy poco probable, pero..
            'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
            OtherUserIndex = MapData(map).Tile(X, Y).UserIndex
            If OtherUserIndex <> 0 Then
                'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
                If UserList(OtherUserIndex).ComUsu.DestUsu > 0 Then
                    'Le avisamos al que estaba comerciando que se tuvo que ir.
                    If UserList(UserList(OtherUserIndex).ComUsu.DestUsu).flags.UserLogged Then
                        Call FinComerciarUsu(UserList(OtherUserIndex).ComUsu.DestUsu)
                        Call WriteConsoleMsg(UserList(OtherUserIndex).ComUsu.DestUsu, "Comercio cancelado. El otro usuario se ha desconectado.", FontTypeNames.FONTTYPE_TALK)
                        Call FlushBuffer(UserList(OtherUserIndex).ComUsu.DestUsu)
                    End If
                    'Lo sacamos.
                    If UserList(OtherUserIndex).flags.UserLogged Then
                        Call FinComerciarUsu(OtherUserIndex)
                        Call WriteErrorMsg(OtherUserIndex, "Alguien se ha conectado donde te encontrabas, por favor reconéctate...")
                        Call FlushBuffer(OtherUserIndex)
                    End If
                End If
            
                Call CloseSocket(OtherUserIndex)
            End If
        End If
    End If

End Sub

Function LegalPosNPC(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal AguaValida As Byte, Optional ByVal IsPet As Boolean = False) As Boolean
'***************************************************
'Autor: Unkwnown
'Last Modification: 27/04/2009
'Checks if it's a Legal pos for the npc to move to.
'***************************************************
Dim IsDeadChar As Boolean
Dim UserIndex As Integer
    If (map <= 0 Or map > NumMaps) Or (map = 1 And (X < 1 Or X > MapInfo(map).Width Or Y < 1 Or Y > MapInfo(map).Height)) Then
        LegalPosNPC = False
        Exit Function
    ElseIf (X < 1 Or X > MapInfo(map).Width Or Y < 1 Or Y > MapInfo(map).Height) Then
        LegalPosNPC = False
        Exit Function
    End If

    UserIndex = MapData(map).Tile(X, Y).UserIndex
    If UserIndex > 0 Then
        IsDeadChar = UserList(UserIndex).flags.Muerto = 1
    Else
        IsDeadChar = False
    End If

    If AguaValida = 0 Then
        LegalPosNPC = (MapData(map).Tile(X, Y).Blocked <> 1) And _
        (MapData(map).Tile(X, Y).UserIndex = 0 Or IsDeadChar) And _
        (MapData(map).Tile(X, Y).NpcIndex = 0) And _
        (MapData(map).Tile(X, Y).Trigger <> eTrigger.POSINVALIDA Or IsPet) _
        And Not HayAgua(map, X, Y)
    Else
        LegalPosNPC = (MapData(map).Tile(X, Y).Blocked <> 1) And _
        (MapData(map).Tile(X, Y).UserIndex = 0 Or IsDeadChar) And _
        (MapData(map).Tile(X, Y).NpcIndex = 0) And _
        (MapData(map).Tile(X, Y).Trigger <> eTrigger.POSINVALIDA Or IsPet)
    End If
End Function

Sub SendHelp(ByVal index As Integer)
Dim NumHelpLines As Integer
Dim LoopC As Integer

NumHelpLines = val(GetVar(DatPath & "Help.dat", "INIT", "NumLines"))

For LoopC = 1 To NumHelpLines
    Call WriteConsoleMsg(index, GetVar(DatPath & "Help.dat", "Help", "Line" & LoopC), FontTypeNames.FONTTYPE_INFO)
Next LoopC

End Sub

Public Sub Expresar(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
    If Npclist(NpcIndex).NroExpresiones > 0 Then
        Dim randomi
        randomi = RandomNumber(1, Npclist(NpcIndex).NroExpresiones)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Npclist(NpcIndex).Expresiones(randomi), Npclist(NpcIndex).Char.CharIndex, vbWhite))
    End If
End Sub

Sub LookatTile(ByVal UserIndex As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 26/03/2009
'13/02/2009: ZaMa - EL nombre del gm que aparece por consola al clickearlo, tiene el color correspondiente a su rango
'***************************************************

    On Error GoTo errhandler

    'Responde al click del usuario sobre el mapa
    Dim FoundChar As Byte
    Dim FoundSomething As Byte
    Dim TempCharIndex As Integer
    Dim Stat As String
    Dim ft As FontTypeNames

    UserList(UserIndex).flags.TargetObj = 0

    '¿Rango Visión? (ToxicWaste)
    If (Abs(UserList(UserIndex).Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(UserList(UserIndex).Pos.X - X) > RANGO_VISION_X) Then
        Exit Sub
    End If

    '¿Posicion valida?
    If InMapBounds(map, X, Y) Then
        UserList(UserIndex).flags.TargetMap = map
        UserList(UserIndex).flags.targetX = X
        UserList(UserIndex).flags.targetY = Y
        '¿Es un obj?
        If MapData(map).Tile(X, Y).ObjInfo.ObjIndex > 0 Then
            'Informa el nombre
            UserList(UserIndex).flags.TargetObjMap = map
            UserList(UserIndex).flags.TargetObjX = X
            UserList(UserIndex).flags.TargetObjY = Y
            FoundSomething = 1
        ElseIf MapData(map).Tile(X + 1, Y).ObjInfo.ObjIndex > 0 Then
            'Informa el nombre
            If ObjData(MapData(map).Tile(X + 1, Y).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
                UserList(UserIndex).flags.TargetObjMap = map
                UserList(UserIndex).flags.TargetObjX = X + 1
                UserList(UserIndex).flags.TargetObjY = Y
                FoundSomething = 1
            End If
        ElseIf MapData(map).Tile(X + 1, Y + 1).ObjInfo.ObjIndex > 0 Then
            If ObjData(MapData(map).Tile(X + 1, Y + 1).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
                'Informa el nombre
                UserList(UserIndex).flags.TargetObjMap = map
                UserList(UserIndex).flags.TargetObjX = X + 1
                UserList(UserIndex).flags.TargetObjY = Y + 1
                FoundSomething = 1
            End If
        ElseIf MapData(map).Tile(X, Y + 1).ObjInfo.ObjIndex > 0 Then
            If ObjData(MapData(map).Tile(X, Y + 1).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
                'Informa el nombre
                UserList(UserIndex).flags.TargetObjMap = map
                UserList(UserIndex).flags.TargetObjX = X
                UserList(UserIndex).flags.TargetObjY = Y + 1
                FoundSomething = 1
            End If
        End If

        If FoundSomething = 1 Then
            UserList(UserIndex).flags.TargetObj = MapData(map).Tile(UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex
            If MostrarCantidad(UserList(UserIndex).flags.TargetObj) Then
                Call WriteConsoleMsg(UserIndex, ObjData(UserList(UserIndex).flags.TargetObj).Name & " - " & MapData(UserList(UserIndex).flags.TargetObjMap).Tile(UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.Amount & "", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, ObjData(UserList(UserIndex).flags.TargetObj).Name, FontTypeNames.FONTTYPE_INFO)
            End If

        End If

        'bots
        If Y + 1 <= YMaxMapSize Then
            UserList(UserIndex).flags.targetBot = MapData(map).Tile(X, Y).BotIndex
            If Not UserList(UserIndex).flags.targetBot <> 0 Then UserList(UserIndex).flags.targetBot = MapData(map).Tile(X, Y + 1).BotIndex

            'Target the botName : D
            If UserList(UserIndex).flags.targetBot <> 0 Then
                'If ia_Bot(.TargetBOT).GrupoID = UserList(UserIndex).Group_User.Grupo_ID Then
                If ia_Bot(UserList(UserIndex).flags.targetBot).Invocado Then
                    Dim tmp_Font As FontTypeNames

                    If ia_Bot(UserList(UserIndex).flags.targetBot).esCriminal Then
                        tmp_Font = FontTypeNames.FONTTYPE_FIGHT
                    Else
                        tmp_Font = FontTypeNames.FONTTYPE_CITIZEN
                    End If

                    Call WriteConsoleMsg(UserIndex, "Ves a " & ia_Bot(UserList(UserIndex).flags.targetBot).Tag, tmp_Font)
                End If
                'Else
                UserList(UserIndex).flags.targetBot = 0
                'End If
            End If
        End If
        'bots







        '¿Es un personaje?
        If MapData(map).Tile(X, Y).UserIndex > 0 Then
            TempCharIndex = MapData(map).Tile(X, Y).UserIndex
            FoundChar = 1
        End If
        If MapData(map).Tile(X, Y).NpcIndex > 0 Then
            TempCharIndex = MapData(map).Tile(X, Y).NpcIndex
            FoundChar = 2
        End If
        '¿Es un personaje?
        If FoundChar = 0 Then
            If Y + 1 <= MapInfo(map).Height Then
                If MapData(map).Tile(X, Y + 1).UserIndex > 0 Then
                    TempCharIndex = MapData(map).Tile(X, Y + 1).UserIndex
                    FoundChar = 1
                End If
                If MapData(map).Tile(X, Y + 1).NpcIndex > 0 Then
                    TempCharIndex = MapData(map).Tile(X, Y + 1).NpcIndex
                    FoundChar = 2
                End If
            End If
        End If


        'Reaccion al personaje
        If FoundChar = 1 Then    '  ¿Encontro un Usuario?

            If UserList(TempCharIndex).flags.AdminInvisible = 0 Or UserList(UserIndex).flags.Privilegios And PlayerType.Dios Then

                If LenB(UserList(TempCharIndex).DescRM) = 0 And UserList(TempCharIndex).showName Then    'No tiene descRM y quiere que se vea su nombre.
                    If EsNewbie(TempCharIndex) Then
                        Stat = " <NEWBIE>"
                    End If

                    If UserList(TempCharIndex).Faccion.ArmadaReal = 1 Then
                        Stat = Stat & " <Ejército Real> " & "<" & TituloReal(TempCharIndex) & ">"
                    ElseIf UserList(TempCharIndex).Faccion.FuerzasCaos = 1 Then
                        Stat = Stat & " <Legión Oscura> " & "<" & TituloCaos(TempCharIndex) & ">"
                    End If

                    If UserList(TempCharIndex).GuildIndex > 0 Then
                        Stat = Stat & " <" & modGuilds.GuildName(UserList(TempCharIndex).GuildIndex) & ">"
                    End If

                    If Len(UserList(TempCharIndex).desc) > 0 Then
                        Stat = "Ves a " & UserList(TempCharIndex).Name & Stat & " - " & UserList(TempCharIndex).desc
                    Else
                        Stat = "Ves a " & UserList(TempCharIndex).Name & Stat
                    End If


                    If UserList(TempCharIndex).flags.Privilegios And PlayerType.RoyalCouncil Then
                        Stat = Stat & " [CONSEJO DE BANDERBILL]"
                        ft = FontTypeNames.FONTTYPE_CONSEJOVesA
                    ElseIf UserList(TempCharIndex).flags.Privilegios And PlayerType.ChaosCouncil Then
                        Stat = Stat & " [CONCILIO DE LAS SOMBRAS]"
                        ft = FontTypeNames.FONTTYPE_CONSEJOCAOSVesA
                    Else
                        If Not UserList(TempCharIndex).flags.Privilegios And PlayerType.User Then
                            Stat = Stat & " <GAME MASTER>"

                            ' Elijo el color segun el rango del GM:
                            ' Dios
                            If UserList(TempCharIndex).flags.Privilegios = PlayerType.Dios Then
                                ft = FontTypeNames.FONTTYPE_DIOS
                                ' Gm
                            ElseIf UserList(TempCharIndex).flags.Privilegios = PlayerType.SemiDios Then
                                ft = FontTypeNames.FONTTYPE_GM
                                ' Conse
                            ElseIf UserList(TempCharIndex).flags.Privilegios = PlayerType.Consejero Then
                                ft = FontTypeNames.FONTTYPE_CONSE
                                ' Rm o Dsrm
                            ElseIf UserList(TempCharIndex).flags.Privilegios = (PlayerType.RoleMaster Or PlayerType.Consejero) Or UserList(TempCharIndex).flags.Privilegios = (PlayerType.RoleMaster Or PlayerType.Dios) Then
                                ft = FontTypeNames.FONTTYPE_EJECUCION
                            End If

                        ElseIf Criminal(TempCharIndex) Then
                            Stat = Stat & " <CRIMINAL>"
                            ft = FontTypeNames.FONTTYPE_FIGHT
                        Else
                            Stat = Stat & " <CIUDADANO>"
                            ft = FontTypeNames.FONTTYPE_CITIZEN
                        End If
                    End If
                Else  'Si tiene descRM la muestro siempre.
                    Stat = UserList(TempCharIndex).DescRM
                    ft = FontTypeNames.FONTTYPE_INFOBOLD
                End If

                If LenB(Stat) > 0 Then
                    Call WriteConsoleMsg(UserIndex, Stat, ft)
                End If

                FoundSomething = 1
                UserList(UserIndex).flags.TargetUser = TempCharIndex
                UserList(UserIndex).flags.TargetNPC = 0
                UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
            End If

        End If
        If FoundChar = 2 Then    '¿Encontro un NPC?
            Dim estatus As String
            If Npclist(TempCharIndex).Stats.MaxHP = 0 Then
                estatus = ""
            ElseIf UserList(UserIndex).flags.Privilegios And (PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin) Then
                estatus = "(" & Npclist(TempCharIndex).Stats.MinHP & "/" & Npclist(TempCharIndex).Stats.MaxHP & ") "
            Else
                If UserList(UserIndex).flags.Muerto = 0 Then
                    If UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 0 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) <= 10 Then
                        estatus = "(Dudoso) "
                    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) > 10 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) <= 20 Then
                        If Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP / 2) Then
                            estatus = "(Herido) "
                        Else
                            estatus = "(Sano) "
                        End If
                    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) > 20 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) <= 30 Then
                        If Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.5) Then
                            estatus = "(Malherido) "
                        ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.75) Then
                            estatus = "(Herido) "
                        Else
                            estatus = "(Sano) "
                        End If
                    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) > 30 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) <= 40 Then
                        If Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.25) Then
                            estatus = "(Muy malherido) "
                        ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.5) Then
                            estatus = "(Herido) "
                        ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.75) Then
                            estatus = "(Levemente herido) "
                        Else
                            estatus = "(Sano) "
                        End If
                    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) > 40 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) < 60 Then
                        If Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.05) Then
                            estatus = "(Agonizando) "
                        ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.1) Then
                            estatus = "(Casi muerto) "
                        ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.25) Then
                            estatus = "(Muy Malherido) "
                        ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.5) Then
                            estatus = "(Herido) "
                        ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.75) Then
                            estatus = "(Levemente herido) "
                        ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP) Then
                            estatus = "(Sano) "
                        Else
                            estatus = "(Intacto) "
                        End If
                    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 60 Then
                        estatus = "(" & Npclist(TempCharIndex).Stats.MinHP & "/" & Npclist(TempCharIndex).Stats.MaxHP & ") "
                    Else
                        estatus = "!error!"
                    End If
                End If
            End If

            If Npclist(TempCharIndex).NpcType = Marinero Then
                Call HablaMarinero(UserIndex, TempCharIndex)
            ElseIf Len(Npclist(TempCharIndex).desc) > 1 Then
               If UserList(UserIndex).genero = Hombre Then
                Call WriteChatOverHead(UserIndex, Npclist(TempCharIndex).desc, Npclist(TempCharIndex).Char.CharIndex, vbWhite)
                Else
                Call WriteChatOverHead(UserIndex, Npclist(TempCharIndex).desc2, Npclist(TempCharIndex).Char.CharIndex, vbWhite)
                End If
            ElseIf TempCharIndex = CentinelaNPCIndex Then
                'Enviamos nuevamente el texto del centinela según quien pregunta
                Call modCentinela.CentinelaSendClave(UserIndex)
            Else
                If Npclist(TempCharIndex).MaestroUser > 0 Then
                    Call WriteConsoleMsg(UserIndex, estatus & Npclist(TempCharIndex).Name & " es mascota de " & UserList(Npclist(TempCharIndex).MaestroUser).Name, FontTypeNames.FONTTYPE_INFO)
                ElseIf Npclist(TempCharIndex).NpcType = eNPCType.Mercader Then
                    Call MercaderClicked(TempCharIndex, UserIndex)
                Else
                    Call WriteConsoleMsg(UserIndex, estatus & Npclist(TempCharIndex).Name & ".", FontTypeNames.FONTTYPE_INFO)
                    'If UserList(UserIndex).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                    '    Call WriteConsoleMsg(UserIndex, "Le pegó primero: " & Npclist(TempCharIndex).flags.AttackedFirstBy & ".", FontTypeNames.FONTTYPE_INFO)
                    'End If
                End If

            End If
            FoundSomething = 1
            UserList(UserIndex).flags.TargetNpcTipo = Npclist(TempCharIndex).NpcType
            UserList(UserIndex).flags.TargetNPC = TempCharIndex
            UserList(UserIndex).flags.TargetUser = 0
            UserList(UserIndex).flags.TargetObj = 0
            'quest
            Dim i As Long, j As Long

            For i = 1 To MAXUSERQUESTS

                With UserList(UserIndex).QuestStats.Quests(i)

                    If .QuestIndex Then
                        If QuestList(.QuestIndex).RequiredTargetNPCs Then

                            For j = 1 To QuestList(.QuestIndex).RequiredTargetNPCs

                                If QuestList(.QuestIndex).RequiredTargetNPC(j).NpcIndex = Npclist(TempCharIndex).Numero Then
                                    If QuestList(.QuestIndex).RequiredTargetNPC(j).Amount > .NPCsTarget(j) Then
                                        .NPCsTarget(j) = .NPCsTarget(j) + 1

                                    End If

                                    If QuestList(.QuestIndex).RequiredTargetNPC(j).Amount = .NPCsTarget(j) Then
                                        Call FinishQuest(UserIndex, .QuestIndex, i)
                                        Call WriteUpdateNPCSimbolo(UserIndex, TempCharIndex, 3)
                                        Call WriteChatOverHead(UserIndex, "¡Quest Finalizada!", Npclist(TempCharIndex).Char.CharIndex, vbYellow)
                                        Call WriteConsoleMsg(UserIndex, "Quest Finalizada!", FontTypeNames.FONTTYPE_INFO)
                                    End If

                                End If

                            Next j

                        End If

                    End If

                End With

            Next i

        End If



        If FoundChar = 0 Then
            UserList(UserIndex).flags.TargetNPC = 0
            UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
            UserList(UserIndex).flags.TargetUser = 0
        End If

        '*** NO ENCOTRO NADA ***
        If FoundSomething = 0 Then
            UserList(UserIndex).flags.TargetNPC = 0
            UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
            UserList(UserIndex).flags.TargetUser = 0
            UserList(UserIndex).flags.TargetObj = 0
            UserList(UserIndex).flags.TargetObjMap = 0
            UserList(UserIndex).flags.TargetObjX = 0
            UserList(UserIndex).flags.TargetObjY = 0
            'Call WriteConsoleMsg(UserIndex, "No ves nada interesante.", FontTypeNames.FONTTYPE_INFO)
        End If

    Else
        If FoundSomething = 0 Then
            UserList(UserIndex).flags.TargetNPC = 0
            UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
            UserList(UserIndex).flags.TargetUser = 0
            UserList(UserIndex).flags.TargetObj = 0
            UserList(UserIndex).flags.TargetObjMap = 0
            UserList(UserIndex).flags.TargetObjX = 0
            UserList(UserIndex).flags.TargetObjY = 0
            'Call WriteConsoleMsg(UserIndex, "No ves nada interesante.", FontTypeNames.FONTTYPE_INFO)
        End If
    End If

    Exit Sub

errhandler:
    Call LogError("Error en LookAtTile. Error " & Err.Number & " : " & Err.Description)

End Sub

Function FindDirection(Pos As WorldPos, Target As WorldPos) As eHeading
'*****************************************************************
'Devuelve la direccion en la cual el target se encuentra
'desde pos, 0 si la direc es igual
'*****************************************************************
Dim X As Integer
Dim Y As Integer

X = Pos.X - Target.X
Y = Pos.Y - Target.Y

'NE
If Sgn(X) = -1 And Sgn(Y) = 1 Then
    FindDirection = IIf(RandomNumber(0, 1), eHeading.NORTH, eHeading.EAST)
    Exit Function
End If

'NW
If Sgn(X) = 1 And Sgn(Y) = 1 Then
    FindDirection = IIf(RandomNumber(0, 1), eHeading.WEST, eHeading.NORTH)
    Exit Function
End If

'SW
If Sgn(X) = 1 And Sgn(Y) = -1 Then
    FindDirection = IIf(RandomNumber(0, 1), eHeading.WEST, eHeading.SOUTH)
    Exit Function
End If

'SE
If Sgn(X) = -1 And Sgn(Y) = -1 Then
    FindDirection = IIf(RandomNumber(0, 1), eHeading.SOUTH, eHeading.EAST)
    Exit Function
End If

'Sur
If Sgn(X) = 0 And Sgn(Y) = -1 Then
    FindDirection = eHeading.SOUTH
    Exit Function
End If

'norte
If Sgn(X) = 0 And Sgn(Y) = 1 Then
    FindDirection = eHeading.NORTH
    Exit Function
End If

'oeste
If Sgn(X) = 1 And Sgn(Y) = 0 Then
    FindDirection = eHeading.WEST
    Exit Function
End If

'este
If Sgn(X) = -1 And Sgn(Y) = 0 Then
    FindDirection = eHeading.EAST
    Exit Function
End If

'misma
If Sgn(X) = 0 And Sgn(Y) = 0 Then
    FindDirection = 0
    Exit Function
End If

End Function

'[Barrin 30-11-03]
Public Function ItemNoEsDeMapa(ByVal index As Integer) As Boolean

ItemNoEsDeMapa = ObjData(index).OBJType <> eOBJType.otPuertas And _
            ObjData(index).OBJType <> eOBJType.otForos And _
            ObjData(index).OBJType <> eOBJType.otCarteles And _
            ObjData(index).OBJType <> eOBJType.otArboles And _
            ObjData(index).OBJType <> eOBJType.otArbolElfico And _
            ObjData(index).OBJType <> eOBJType.otFlores And _
            ObjData(index).OBJType <> eOBJType.otYacimiento And _
            ObjData(index).OBJType <> eOBJType.otTeleport
            
End Function
'[/Barrin 30-11-03]

Public Function MostrarCantidad(ByVal index As Integer) As Boolean
MostrarCantidad = ObjData(index).OBJType <> eOBJType.otPuertas And _
            ObjData(index).OBJType <> eOBJType.otForos And _
            ObjData(index).OBJType <> eOBJType.otCarteles And _
            ObjData(index).OBJType <> eOBJType.otArboles And _
            ObjData(index).OBJType <> eOBJType.otArbolElfico Or _
            ObjData(index).OBJType <> eOBJType.otFlores Or _
            ObjData(index).OBJType <> eOBJType.otYacimiento And _
            ObjData(index).OBJType <> eOBJType.otTeleport
End Function

Public Function EsObjetoFijo(ByVal OBJType As eOBJType) As Boolean

EsObjetoFijo = OBJType = eOBJType.otForos Or _
               OBJType = eOBJType.otCarteles Or _
               OBJType = eOBJType.otArboles Or _
               OBJType = eOBJType.otArbolElfico Or _
               OBJType = eOBJType.otFlores Or _
               OBJType = eOBJType.otYacimiento

End Function

Function min(val1 As Long, val2 As Long) As Long
If val1 < val2 Then
    min = val1
Else
    min = val2
End If
End Function


Sub RestringirMapa(ByVal UserIndex As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, Nombre As String, Lvl As Byte, mm As Integer, mx As Integer, my As Integer, nPos As WorldPos)
   With MapData(map).Tile(X, Y)
   If Zonas(BuscarZona(.TileExit.map, .TileExit.X, .TileExit.Y)).Nombre = Nombre Then
                      If UserList(UserIndex).Stats.ELV >= Lvl Then
                            If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(UserIndex)) Then
                                Call WarpUserChar(UserIndex, .TileExit.map, .TileExit.X, .TileExit.Y, True)
                            Else
                                Call ClosestLegalPos(.TileExit, nPos)
                                If nPos.X <> 0 And nPos.Y <> 0 Then
                                    Call WarpUserChar(UserIndex, nPos.map, nPos.X, nPos.Y, True)
                                End If
                            End If
                       Else
                           Call WriteConsoleMsg(UserIndex, "Este mapa es demasiado peligroso para tu nivel.", FontTypeNames.FONTTYPE_INFO)
                           Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
 
                           If nPos.X <> 0 And nPos.Y <> 0 Then
                               Call WarpUserChar(UserIndex, mm, mx, my, False)
                           End If
                       End If
                End If
                End With
End Sub


Sub RestringirMapaItem(ByVal UserIndex As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal zona As String, ByVal Item As Integer, mm As Integer, mx As Integer, my As Integer, nPos As WorldPos, QuitarItem As Boolean)
   
   
   
  With MapData(map).Tile(X, Y)
   
   If Zonas(BuscarZona(.TileExit.map, .TileExit.X, .TileExit.Y)).Nombre = zona Then
                 
                    If UsuarioTineItem(UserIndex, Item) = 1 Then
                        If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(UserIndex)) Then
                            Call WarpUserChar(UserIndex, .TileExit.map, .TileExit.X, .TileExit.Y, True)
                        Else
                            Call ClosestLegalPos(.TileExit, nPos)
                            If nPos.X <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(UserIndex, nPos.map, nPos.X, nPos.Y, True)
                            End If
                        End If
                    Else
                        Call WriteConsoleMsg(UserIndex, "Si no tienes " & ObjData(Item).Name & " no puedes entrar", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)

                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, mm, mx, my, False)
                        End If
                    End If

                End If
   
   
 End With
 If QuitarItem = True Then Call QuitarItemInv(UserIndex, Item)
  


 
End Sub


Sub RestringirMapaMinMax(ByVal UserIndex As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, Nombre As String, LvlMin As Byte, LvlMax As Byte, mm As Integer, mx As Integer, my As Integer, nPos As WorldPos)
    With MapData(map).Tile(X, Y)
        If Zonas(BuscarZona(.TileExit.map, .TileExit.X, .TileExit.Y)).Nombre = Nombre Then
            If UserList(UserIndex).Stats.ELV >= LvlMin And UserList(UserIndex).Stats.ELV <= LvlMax Then
                If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(UserIndex)) Then
                    Call WarpUserChar(UserIndex, .TileExit.map, .TileExit.X, .TileExit.Y, True)
                Else
                    Call ClosestLegalPos(.TileExit, nPos)
                    If nPos.X <> 0 And nPos.Y <> 0 Then
                        Call WarpUserChar(UserIndex, nPos.map, nPos.X, nPos.Y, True)
                    End If
                End If
            Else
                If UserList(UserIndex).Stats.ELV < LvlMin Then
                    Call WriteConsoleMsg(UserIndex, "Este mapa es demasiado peligroso para tu nivel.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "Eres demasiado poderoso para este mapa.. Déjaselo a los mas débiles.", FontTypeNames.FONTTYPE_INFO)
                End If
                Call ClosestStablePos(UserList(UserIndex).Pos, nPos)

                If nPos.X <> 0 And nPos.Y <> 0 Then
                    Call WarpUserChar(UserIndex, mm, mx, my, False)
                End If
            End If
        End If
    End With
End Sub

