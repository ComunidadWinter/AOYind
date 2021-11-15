Attribute VB_Name = "modMonturas"
Option Explicit

Public Sub Montar(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

On Error GoTo errhandle

With UserList(UserIndex)

If .flags.Navegando = 1 Then Exit Sub
If .flags.Muerto > 0 Then Exit Sub

NpcIndex = .flags.NpcMonturaIndex

If NpcIndex = 0 Then
    Call WriteMultiMessage(UserIndex, eMessages.DomarAnimalParaMontarlo)
    'Call WriteConsoleMsg(UserIndex, "¡Debes domesticar primero algún animal para montarlo!", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If Npclist(NpcIndex).EquitandoBody = 0 Then Exit Sub

If .flags.Equitando = False Then
    
    'If UserList(UserIndex).flags.TargetNPC = 0 Then
    '    Call WriteConsoleMsg(UserIndex, "¡Debes hacer click sobre una animal salvaje que se pueda montar!", FontTypeNames.FONTTYPE_INFO)
    'Exit Sub
    
    If Npclist(NpcIndex).MaestroUser <> UserIndex Then
        'Call WriteConsoleMsg(UserIndex, "¡Este animal no te acepta como su amo!", FontTypeNames.FONTTYPE_INFO)
        Call WriteMultiMessage(UserIndex, eMessages.MonturaNoTeAceptaComoSuAmo)
        Call LiberarMontura(UserIndex)
        Exit Sub
    End If
   
     If Distancia(Npclist(.flags.NpcMonturaIndex).Pos, .Pos) > 1 Then
        'Call WriteConsoleMsg(UserIndex, "¡Estás muy lejos del animal para montarlo!", FontTypeNames.FONTTYPE_INFO)
        Call WriteMultiMessage(UserIndex, eMessages.UserEstaLejos)
        Exit Sub
    End If
    
    Dim tHeading As Byte
    tHeading = .Char.heading
        
    'NPC Montura
    If Npclist(NpcIndex).EquitandoBody > 0 Then
      
        .Char.Body = Npclist(NpcIndex).EquitandoBody
        
        .Char.ShieldAnim = NingunEscudo
        .Char.WeaponAnim = NingunArma
        .flags.Desnudo = 0
        .flags.Equitando = True
        .flags.TargetNPC = 0 'Previene un error
        
        If Npclist(NpcIndex).flags.Snd1 > 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Npclist(NpcIndex).flags.Snd1, .Pos.X, .Pos.Y))
        End If
        
        'Call WriteConsoleMsg(UserIndex, "Hás montado un " & Npclist(NpcIndex).Name & "!", FontTypeNames.FONTTYPE_INFO)
        Call WriteMultiMessage(UserIndex, eMessages.UserHaMontado)
        
        Call ChangeUserChar(UserIndex, .Char.Body, UserList(UserIndex).Char.Head, tHeading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, 0)
    
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEquitando(.Char.CharIndex, True))
    
        .flags.NpcMonturaIndex = 0
        
        Call QuitarNPC(NpcIndex)
            
    End If
    
End If

End With

Exit Sub

errhandle:
Call LogError("ERROR En Equitacion: " & Err.Number & " - " & Err.Description)

End Sub


Public Sub QuitarMontura(ByVal UserIndex As Integer, Optional ByVal PerteneceUserIndex As Boolean = True)

With UserList(UserIndex)

If .flags.NpcMonturaNumero > 0 And .flags.Equitando = True Then

    Dim ind As Integer
    Dim Pos As WorldPos
    Dim tHeading As Byte
        
    Pos = .Pos
        
    tHeading = .Char.heading
    
    If .Char.heading = EAST Or .Char.heading = WEST Then
        If LegalPos(.Pos.map, .Pos.X, .Pos.Y - 1, False, True) Then
           Pos.X = .Pos.X
           Pos.Y = .Pos.Y - 1
        ElseIf LegalPos(.Pos.map, .Pos.X, .Pos.Y + 1, False, True) Then
           Pos.X = .Pos.X
           Pos.Y = .Pos.Y + 1
        Else
            Call FindLegalPos(UserIndex, Pos.map, Pos.X, Pos.Y)
            'Call WriteConsoleMsg(UserIndex, "No hay lugar para desmontar el animal!.", FontTypeNames.FONTTYPE_INFO)
            'Call WriteMultiMessage(UserIndex, eMessages.NoHayLugarParaDesmontar)
            'Exit Sub
        End If
    ElseIf .Char.heading = SOUTH Or .Char.heading = NORTH Then
        If LegalPos(.Pos.map, .Pos.X + 1, .Pos.Y, False, True) Then
           Pos.X = .Pos.X + 1
           Pos.Y = .Pos.Y
        ElseIf LegalPos(.Pos.map, .Pos.X - 1, .Pos.Y, False, True) Then
           Pos.X = .Pos.X - 1
           Pos.Y = .Pos.Y
        Else
            Call FindLegalPos(UserIndex, Pos.map, Pos.X, Pos.Y)
            'Call WriteConsoleMsg(UserIndex, "No hay lugar para desmontar el animal!.", FontTypeNames.FONTTYPE_INFO)
            'Call WriteMultiMessage(UserIndex, eMessages.NoHayLugarParaDesmontar)
            'Exit Sub
        End If
    End If

    .flags.Equitando = False
    
    If .flags.Muerto = 0 Then
        .Char.Head = .OrigChar.Head
        
        If .Invent.ArmourEqpObjIndex > 0 Then
            .Char.Body = ObjData(.Invent.ArmourEqpObjIndex).Ropaje
        Else
            Call DarCuerpoDesnudo(UserIndex)
        End If
            
        If .Invent.EscudoEqpObjIndex > 0 Then _
            .Char.ShieldAnim = ObjData(.Invent.EscudoEqpObjIndex).ShieldAnim
        If .Invent.WeaponEqpObjIndex > 0 Then _
            .Char.WeaponAnim = ObjData(.Invent.WeaponEqpObjIndex).WeaponAnim
        If .Invent.CascoEqpObjIndex > 0 Then _
            .Char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
             If .Invent.AlasEqpObjIndex > 0 Then _
            .Char.alaIndex = ObjData(.Invent.AlasEqpObjIndex).alaIndex
    Else
    
        .Char.Body = iCuerpoMuerto
        .Char.Head = iCabezaMuerto
        .Char.ShieldAnim = NingunEscudo
        .Char.WeaponAnim = NingunArma
        .Char.CascoAnim = NingunCasco
        .Char.alaIndex = 0
        
    End If
           
        ind = SpawnNpc(.flags.NpcMonturaNumero, Pos, False, False, 1)
        
        If ind = 0 Then Exit Sub
        
        Npclist(ind).Char.heading = tHeading
        
        If PerteneceUserIndex = True Then
            Npclist(ind).MaestroUser = UserIndex
            .flags.NpcMonturaIndex = ind
    '        Npclist(ind).flags.OldMovement = Npclist(ind).Movement
    '        Npclist(ind).flags.OldHostil = Npclist(ind).Hostile
            Call FollowAmo(ind)
        Else
            .flags.NpcMonturaIndex = 0
            .flags.NpcMonturaNumero = 0
        End If
    
    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_DESMONTAR, .Pos.X, .Pos.Y))
    'Call WriteConsoleMsg(UserIndex, "Te hás desmontado del animal!.", FontTypeNames.FONTTYPE_INFO)
    Call WriteMultiMessage(UserIndex, eMessages.UserHaDesmontado)
    
    Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, UserList(UserIndex).Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.alaIndex)
    Call ChangeNPCChar(ind, Npclist(ind).Char.Body, Npclist(ind).Char.Head, eHeading.WEST)
    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEquitando(.Char.CharIndex, False))
    
End If

End With

End Sub
Public Function UserMontado(ByVal UserIndex As Integer) As Boolean
    UserMontado = UserList(UserIndex).flags.Equitando
End Function


Public Function LiberarMontura(ByVal UserIndex As Integer)

With UserList(UserIndex)
    If .flags.NpcMonturaNumero > 0 Then
        Npclist(.flags.NpcMonturaIndex).Movement = Npclist(.flags.NpcMonturaIndex).flags.OldMovement
        Npclist(.flags.NpcMonturaIndex).Hostile = Npclist(.flags.NpcMonturaIndex).flags.OldHostil
        Npclist(.flags.NpcMonturaIndex).flags.AttackedBy = vbNullString
        Npclist(.flags.NpcMonturaIndex).MaestroUser = 0
        
        'Call WriteConsoleMsg(UserIndex, "El animal ya no te acepta como su dueño.", FontTypeNames.FONTTYPE_INFOBOLD)
        Call WriteMultiMessage(UserIndex, eMessages.MonturaNoTeAceptaComoSuAmo)
                
        'Previene un BUG
        If .flags.NpcMonturaIndex > 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Npclist(.flags.NpcMonturaIndex).flags.Snd1, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        End If
        
        .flags.NpcMonturaIndex = 0
        .flags.NpcMonturaNumero = 0
        
    End If
End With

End Function

Public Sub DomarMascotaMontura(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal puntosDomar As Integer, ByVal puntosRequeridos As Integer)

Dim agonizando As Boolean
Dim porcAgonizando As Byte
    agonizando = Npclist(NpcIndex).Stats.MinHP < Porcentaje(Npclist(NpcIndex).Stats.MaxHP, 10)
    
    If agonizando = True Then
        porcAgonizando = 1 'Si agoniza, calcula sobre el 1%
    Else
        porcAgonizando = 5 'Si no, calcula sobre el 5%
    End If
    
    'If puntosRequeridos <= puntosDomar And RandomNumber(1, Porcentaje(Npclist(NpcIndex).Stats.MaxHP, porcAgonizando)) = 1 Then
     If RandomNumber(1, Porcentaje(Npclist(NpcIndex).Stats.MaxHP, porcAgonizando)) = 1 Then
        UserList(UserIndex).flags.NpcMonturaIndex = NpcIndex
        UserList(UserIndex).flags.NpcMonturaNumero = Npclist(NpcIndex).Numero
        Npclist(NpcIndex).MaestroUser = UserIndex
        Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MaxHP
        Call FollowAmo(NpcIndex)
        'Call WriteConsoleMsg(UserIndex, "Hás logrado domar al animal salvaje.", FontTypeNames.FONTTYPE_INFOBOLD)
        Call WriteMultiMessage(UserIndex, eMessages.UserHaDomado)
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_LIBERAR_MONTURA, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
    Else
        'Call WriteConsoleMsg(UserIndex, "No hás logrado domar al animal salvaje.", FontTypeNames.FONTTYPE_INFO)
        Call WriteMultiMessage(UserIndex, eMessages.NoHasLogradoDomarlo)
    End If
    
End Sub





