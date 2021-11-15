Attribute VB_Name = "SistemaCombate"
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
'
'Diseño y corrección del modulo de combate por
'Gerardo Saiz, gerardosaiz@yahoo.com
'

'9/01/2008 Pablo (ToxicWaste) - Ahora TODOS los modificadores de Clase se controlan desde Balance.dat


Option Explicit

Public Const MAXDISTANCIAARCO As Byte = 21
Public Const MAXDISTANCIAMAGIA As Byte = 21

Public Function MinimoInt(ByVal a As Integer, ByVal B As Integer) As Integer
    If a > B Then
        MinimoInt = B
    Else
        MinimoInt = a
    End If
End Function

Public Function MaximoInt(ByVal a As Integer, ByVal B As Integer) As Integer
    If a > B Then
        MaximoInt = a
    Else
        MaximoInt = B
    End If
End Function

Public Function PoderEvasionEscudo(ByVal UserIndex As Integer) As Long
    PoderEvasionEscudo = (UserList(UserIndex).Stats.UserSkills(eSkill.Defensa) * ModClase(UserList(UserIndex).clase).Evasion) / 2
End Function

Public Function PoderEvasion(ByVal UserIndex As Integer) As Long
    Dim lTemp As Long
    With UserList(UserIndex)
        lTemp = (.Stats.UserSkills(eSkill.Tacticas) + _
          .Stats.UserSkills(eSkill.Tacticas) / 33 * .Stats.UserAtributos(eAtributos.agilidad)) * ModClase(.clase).Evasion
       
        PoderEvasion = (lTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))
    End With
End Function

Private Function PoderAtaqueArma(ByVal UserIndex As Integer) As Long
    Dim PoderAtaqueTemp As Long
    
    With UserList(UserIndex)
        If .Stats.UserSkills(eSkill.Armas) < 31 Then
            PoderAtaqueTemp = .Stats.UserSkills(eSkill.Armas) * ModClase(.clase).AtaqueArmas
        ElseIf .Stats.UserSkills(eSkill.Armas) < 61 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Armas) + .Stats.UserAtributos(eAtributos.agilidad)) * ModClase(.clase).AtaqueArmas
        ElseIf .Stats.UserSkills(eSkill.Armas) < 91 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Armas) + 2 * .Stats.UserAtributos(eAtributos.agilidad)) * ModClase(.clase).AtaqueArmas
        Else
           PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Armas) + 3 * .Stats.UserAtributos(eAtributos.agilidad)) * ModClase(.clase).AtaqueArmas
        End If
        
        PoderAtaqueArma = (PoderAtaqueTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))
    End With
End Function

Private Function PoderAtaqueProyectil(ByVal UserIndex As Integer) As Long
    Dim PoderAtaqueTemp As Long
    
    With UserList(UserIndex)
        If .Stats.UserSkills(eSkill.Proyectiles) < 31 Then
            PoderAtaqueTemp = .Stats.UserSkills(eSkill.Proyectiles) * ModClase(.clase).AtaqueProyectiles
        ElseIf .Stats.UserSkills(eSkill.Proyectiles) < 61 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Proyectiles) + .Stats.UserAtributos(eAtributos.agilidad)) * ModClase(.clase).AtaqueProyectiles
        ElseIf .Stats.UserSkills(eSkill.Proyectiles) < 91 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Proyectiles) + 2 * .Stats.UserAtributos(eAtributos.agilidad)) * ModClase(.clase).AtaqueProyectiles
        Else
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Proyectiles) + 3 * .Stats.UserAtributos(eAtributos.agilidad)) * ModClase(.clase).AtaqueProyectiles
        End If
        
        PoderAtaqueProyectil = (PoderAtaqueTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))
    End With
End Function

Private Function PoderAtaqueWrestling(ByVal UserIndex As Integer) As Long
    Dim PoderAtaqueTemp As Long
    
    With UserList(UserIndex)
        If .Stats.UserSkills(eSkill.Wrestling) < 31 Then
            PoderAtaqueTemp = .Stats.UserSkills(eSkill.Wrestling) * ModClase(.clase).AtaqueArmas
        ElseIf .Stats.UserSkills(eSkill.Wrestling) < 61 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Wrestling) + .Stats.UserAtributos(eAtributos.agilidad)) * ModClase(.clase).AtaqueArmas
        ElseIf .Stats.UserSkills(eSkill.Wrestling) < 91 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Wrestling) + 2 * .Stats.UserAtributos(eAtributos.agilidad)) * ModClase(.clase).AtaqueArmas
        Else
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Wrestling) + 3 * .Stats.UserAtributos(eAtributos.agilidad)) * ModClase(.clase).AtaqueArmas
        End If
        
        PoderAtaqueWrestling = (PoderAtaqueTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))
    End With
End Function

Public Function UserImpactoNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Boolean
    Dim PoderAtaque As Long
    Dim Arma As Integer
    Dim Skill As eSkill
    Dim ProbExito As Long
    
    Arma = UserList(UserIndex).Invent.WeaponEqpObjIndex
    
    If Arma > 0 Then 'Usando un arma
        If ObjData(Arma).proyectil = 1 Then
            PoderAtaque = PoderAtaqueProyectil(UserIndex)
            Skill = eSkill.Proyectiles

        Else
            PoderAtaque = PoderAtaqueArma(UserIndex)
            Skill = eSkill.Armas
        End If
    Else 'Peleando con puños
        PoderAtaque = PoderAtaqueWrestling(UserIndex)
        Skill = eSkill.Wrestling
    End If
    
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + ((PoderAtaque - Npclist(NpcIndex).PoderEvasion) * 0.4)))
    
    UserImpactoNpc = (RandomNumber(1, 80) <= ProbExito)
    
    'If UserImpactoNpc Then
        Call SubirSkill(UserIndex, Skill)
    'End If
End Function

Public Function NpcImpacto(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Revisa si un NPC logra impactar a un user o no
'03/15/2006 Maraxus - Evité una división por cero que eliminaba NPCs
'*************************************************
    Dim Rechazo As Boolean
    Dim ProbRechazo As Long
    Dim ProbExito As Long
    Dim UserEvasion As Long
    Dim NpcPoderAtaque As Long
    Dim PoderEvasioEscudo As Long
    Dim SkillTacticas As Long
    Dim SkillDefensa As Long
    
    UserEvasion = PoderEvasion(UserIndex)
    NpcPoderAtaque = Npclist(NpcIndex).PoderAtaque
    PoderEvasioEscudo = PoderEvasionEscudo(UserIndex)
    
    SkillTacticas = UserList(UserIndex).Stats.UserSkills(eSkill.Tacticas)
    SkillDefensa = UserList(UserIndex).Stats.UserSkills(eSkill.Defensa)
    
    'Esta usando un escudo ???
    If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then UserEvasion = UserEvasion + PoderEvasioEscudo
    
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + ((NpcPoderAtaque - UserEvasion) * 0.4)))
    
    NpcImpacto = (RandomNumber(1, 100) <= ProbExito)
    
    'Si tiene bonificacion de evasión por raza veo si le pega o no
    If NpcImpacto Then
        If ModRaza(UserList(UserIndex).raza).EvitarGolpe > 0 Then
            NpcImpacto = (RandomNumber(1, 1 / ModRaza(UserList(UserIndex).raza).EvitarGolpe) = 1)
        End If
    End If
    
    ' el usuario esta usando un escudo ???
    If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
        If Not NpcImpacto Then
            If SkillDefensa + SkillTacticas > 0 Then  'Evitamos división por cero
                ' Chances are rounded
                ProbRechazo = MaximoInt(10, MinimoInt(90, 100 * SkillDefensa / (SkillDefensa + SkillTacticas)))
                Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
                
                If Rechazo Then
                    'Se rechazo el ataque con el escudo
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAtaca(UserIndex, 2))
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_ESCUDO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                    Call WriteBlockedWithShieldUser(UserIndex)
                    Call SubirSkill(UserIndex, Defensa)
                Else
                    If RandomNumber(1, 3) = 1 Then Call SubirSkill(UserIndex, Defensa)
                End If
            End If
        End If
    End If
End Function

Public Function CalcularDaño(ByVal UserIndex As Integer, Optional ByVal NpcIndex As Integer = 0) As Long
    Dim DañoArma As Long
    Dim DañoUsuario As Long
    Dim Arma As ObjData
    Dim ModifClase As Single
    Dim proyectil As ObjData
    Dim DañoMaxArma As Long
    
    ''sacar esto si no queremos q la matadracos mate el Dragon si o si
    Dim matoDragon As Boolean
    matoDragon = False
    
    With UserList(UserIndex)
        If .Invent.WeaponEqpObjIndex > 0 Then
            Arma = ObjData(.Invent.WeaponEqpObjIndex)
            
            ' Ataca a un npc?
            If NpcIndex > 0 Then
                If Arma.proyectil = 1 Then
                    ModifClase = ModClase(.clase).DañoProyectiles
                    DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                    DañoMaxArma = Arma.MaxHIT
                    
                    If Arma.Municion = 1 Then
                        proyectil = ObjData(.Invent.MunicionEqpObjIndex)
                        DañoArma = DañoArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHIT)
                        ' For some reason this isn't done...
                        'DañoMaxArma = DañoMaxArma + proyectil.MaxHIT
                    End If
                Else
                    ModifClase = ModClase(.clase).DañoArmas
                    
                    If .Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then ' Usa la mata Dragones?
                        If Npclist(NpcIndex).NpcType = DRAGON Then 'Ataca Dragon?
                            DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                            DañoMaxArma = Arma.MaxHIT
                            matoDragon = True ''sacar esto si no queremos q la matadracos mate el Dragon si o si
                        Else ' Sino es Dragon daño es 1
                            DañoArma = 1
                            DañoMaxArma = 1
                        End If
                    Else
                        DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                        DañoMaxArma = Arma.MaxHIT
                    End If
                End If
            Else ' Ataca usuario
                If Arma.proyectil = 1 Then
                    ModifClase = ModClase(.clase).DañoProyectiles
                    DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                    DañoMaxArma = Arma.MaxHIT
                     
                    If Arma.Municion = 1 Then
                        proyectil = ObjData(.Invent.MunicionEqpObjIndex)
                        DañoArma = DañoArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHIT)
                        ' For some reason this isn't done...
                        'DañoMaxArma = DañoMaxArma + proyectil.MaxHIT
                    End If
                Else
                    ModifClase = ModClase(.clase).DañoArmas
                    
                    If .Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then
                        ModifClase = ModClase(.clase).DañoArmas
                        DañoArma = 1 ' Si usa la espada mataDragones daño es 1
                        DañoMaxArma = 1
                    Else
                        DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                        DañoMaxArma = Arma.MaxHIT
                    End If
                End If
            End If
        Else
            ModifClase = ModClase(.clase).DañoWrestling
            DañoArma = RandomNumber(1, 3) 'Hacemos que sea "tipo" una daga el ataque de Wrestling
            DañoMaxArma = 3
        End If
        
        DañoUsuario = RandomNumber(.Stats.MinHIT, .Stats.MaxHIT)
        
        ''sacar esto si no queremos q la matadracos mate el Dragon si o si
        If matoDragon Then
            CalcularDaño = Npclist(NpcIndex).Stats.MinHP + Npclist(NpcIndex).Stats.def
        Else
            CalcularDaño = (3 * DañoArma + ((DañoMaxArma / 5) * MaximoInt(0, .Stats.UserAtributos(eAtributos.fuerza) - 15)) + DañoUsuario) * ModifClase
        End If
    End With
End Function

Public Sub UserDañoNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, Optional ByVal HitArea As Byte = 0)
    Dim daño As Long
    
    daño = CalcularDaño(UserIndex, NpcIndex)
    
    'esta navegando? si es asi le sumamos el daño del barco
    If UserList(UserIndex).flags.Navegando = 1 And UserList(UserIndex).Invent.BarcoObjIndex > 0 Then
        daño = daño + RandomNumber(ObjData(UserList(UserIndex).Invent.BarcoObjIndex).MinHIT, ObjData(UserList(UserIndex).Invent.BarcoObjIndex).MaxHIT)
    End If
    
    With Npclist(NpcIndex)
        daño = daño - .Stats.def
        
        If daño < 0 Then daño = 0
        
        If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
            If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).proyectil Then
                daño = daño * (1 + ModRaza(UserList(UserIndex).raza).Proyectiles)
            Else
                'Si tiene bonificacion por raza
                daño = daño * (1 + ModRaza(UserList(UserIndex).raza).Armas)
            End If
        End If
        
          'HitArea
        If HitArea > 0 Then
            daño = daño - Porcentaje(daño, 15)
        Else
            'If RandomNumber(1, 75) = 1 Then
            '    daño = daño * 1.2
            '
            '    If UserList(UserIndex).genero = Hombre Then
            '        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_CRITICOH, .Pos.X, .Pos.Y))
            '    Else
            '        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_CRITICOM, .Pos.X, .Pos.Y))
            '    End If
            '
            '    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, 41, 0))
            'End If
        End If
        
        If daño < 32000 Then
            Call WriteUserHitNPC(UserIndex, daño, .Char.CharIndex, .NpcType, HitArea, .flags.Colorsangre)
        End If
        
        Call CalcularDarExp(UserIndex, NpcIndex, daño)
        
        .Stats.MinHP = .Stats.MinHP - daño
        
        If .Stats.MinHP > 0 Then
            'Trata de apuñalar por la espalda al enemigo
            If PuedeApuñalar(UserIndex) Then
               Call DoApuñalar(UserIndex, NpcIndex, 0, daño)
               Call SubirSkill(UserIndex, Apuñalar)
            End If
            
            'trata de dar golpe crítico
            'Call DoGolpeCritico(UserIndex, NpcIndex, 0, daño)
        End If
        
        If .EquitandoBody > 0 Then
            If .MaestroUser > 0 Then
                If RandomNumber(1, 5) Then
                    'Call WriteConsoleMsg(.MaestroUser, "¡Están atacando al animal que hás domado!", FontTypeNames.FONTTYPE_FIGHT)
                    Call WriteMultiMessage(UserIndex, eMessages.UserMonturaSiendoAtacada)
                End If
            End If
        
            If .Stats.MinHP < (.Stats.MinHP * 100 / .Stats.MaxHP) Then
                If Npclist(NpcIndex).MaestroUser > 0 Then
                    Call LiberarMontura(Npclist(NpcIndex).MaestroUser)
                End If
                'Call WriteConsoleMsg(UserIndex, "¡El animal está agonizando, habrá mas posibilidades de dormarlo!", FontTypeNames.FONTTYPE_INFO)
                Call WriteMultiMessage(UserIndex, eMessages.CriaturaAgonizando)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateAreaFX(Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y - 2, 40, 1))
            End If
        End If
        
        If .Stats.MinHP <= 0 Then
            ' Si era un Dragon perdemos la espada mataDragones
            If .NpcType = DRAGON Then
                'Si tiene equipada la matadracos se la sacamos
                If UserList(UserIndex).Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then
                    Call QuitarObjetos(EspadaMataDragonesIndex, 1, UserIndex)
                End If
                If .Stats.MaxHP > 100000 Then Call LogDesarrollo(UserList(UserIndex).Name & " mató un dragón")
            End If
            
            ' Para que las mascotas no sigan intentando luchar y
            ' comiencen a seguir al amo
            Dim j As Integer
            For j = 1 To MAXMASCOTAS
                If UserList(UserIndex).MascotasIndex(j) > 0 Then
                    If Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNPC = NpcIndex Then
                        Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNPC = 0
                        Npclist(UserList(UserIndex).MascotasIndex(j)).Movement = TipoAI.SigueAmo
                    End If
                End If
            Next j
            
            Call MuereNpc(NpcIndex, UserIndex)
        End If
    End With
End Sub

Public Sub NpcDaño(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
    Dim daño As Integer
    Dim Lugar As Integer
    Dim absorbido As Integer
    Dim defbarco As Integer
    Dim Obj As ObjData
    
    daño = RandomNumber(Npclist(NpcIndex).Stats.MinHIT, Npclist(NpcIndex).Stats.MaxHIT)
    
    With UserList(UserIndex)
        If .flags.Navegando = 1 And .Invent.BarcoObjIndex > 0 Then
            Obj = ObjData(.Invent.BarcoObjIndex)
            defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)
        End If
        
        Lugar = RandomNumber(PartesCuerpo.bCabeza, PartesCuerpo.bTorso)
        
        Select Case Lugar
            Case PartesCuerpo.bCabeza
                'Si tiene casco absorbe el golpe
                If .Invent.CascoEqpObjIndex > 0 Then
                   Obj = ObjData(.Invent.CascoEqpObjIndex)
                   absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
                End If
          Case Else
                'Si tiene armadura absorbe el golpe
                If .Invent.ArmourEqpObjIndex > 0 Then
                    Dim Obj2 As ObjData
                    Obj = ObjData(.Invent.ArmourEqpObjIndex)
                    If .Invent.EscudoEqpObjIndex Then
                        Obj2 = ObjData(.Invent.EscudoEqpObjIndex)
                        absorbido = RandomNumber(Obj.MinDef + Obj2.MinDef, Obj.MaxDef + Obj2.MaxDef)
                    Else
                        absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
                   End If
                End If
        End Select
        
        absorbido = absorbido + defbarco
        daño = daño - absorbido
        If daño < 1 Then daño = 1
        
        Call WriteNPCHitUser(UserIndex, Lugar, daño, Npclist(NpcIndex).NpcType)
  
        If .flags.Privilegios And PlayerType.User Then .Stats.MinHP = .Stats.MinHP - daño
        
        If .flags.Meditando Then
            If daño > Fix(.Stats.MinHP / 100 * .Stats.UserAtributos(eAtributos.Inteligencia) * .Stats.UserSkills(eSkill.Meditar) / 100 * 12 / (RandomNumber(0, 5) + 7)) Then
                .flags.Meditando = False
                Call WriteMeditateToggle(UserIndex)
                Call WriteConsoleMsg(UserIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
                .Char.FX = 0
                .Char.Loops = 0
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
            End If
        End If
        
        'Muere el usuario
        If .Stats.MinHP <= 0 Then
            
            Call WriteNPCKillUser(UserIndex, Npclist(NpcIndex).NpcType) ' Le informamos que ha muerto ;)
      
            
            'Si lo mato un guardia
            If Criminal(UserIndex) And EsGuardiaReal(NpcIndex) Then
                Call RestarCriminalidad(UserIndex)
                If Not Criminal(UserIndex) And .Faccion.FuerzasCaos = 1 Then Call ExpulsarFaccionCaos(UserIndex)
            End If
            
            If Npclist(NpcIndex).MaestroUser > 0 Then
                Call AllFollowAmo(Npclist(NpcIndex).MaestroUser)
            Else
                'Al matarlo no lo sigue mas
                If Npclist(NpcIndex).Stats.Alineacion = 0 Then
                    Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
                    Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
                    Npclist(NpcIndex).flags.AttackedBy = vbNullString
                End If
            End If
            
            Call UserDie(UserIndex)
        End If
    End With
End Sub

Public Sub RestarCriminalidad(ByVal UserIndex As Integer)
    Dim EraCriminal As Boolean
    EraCriminal = Criminal(UserIndex)
    
    With UserList(UserIndex).Reputacion
        If .BandidoRep > 0 Then
             .BandidoRep = .BandidoRep - vlASALTO
             If .BandidoRep < 0 Then .BandidoRep = 0
        ElseIf .LadronesRep > 0 Then
             .LadronesRep = .LadronesRep - (vlCAZADOR * 10)
             If .LadronesRep < 0 Then .LadronesRep = 0
        End If
    End With
    
    If EraCriminal And Not Criminal(UserIndex) Then
        Call RefreshCharStatus(UserIndex)
    End If
End Sub

Public Sub CheckPets(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, Optional ByVal CheckElementales As Boolean = True)
    Dim j As Integer
    
    For j = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasIndex(j) > 0 Then
           If UserList(UserIndex).MascotasIndex(j) <> NpcIndex Then
            If CheckElementales Or (Npclist(UserList(UserIndex).MascotasIndex(j)).Numero <> ELEMENTALFUEGO And Npclist(UserList(UserIndex).MascotasIndex(j)).Numero <> ELEMENTALTIERRA) Then
                If Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNPC = 0 Then Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNPC = NpcIndex
                Npclist(UserList(UserIndex).MascotasIndex(j)).Movement = TipoAI.NpcAtacaNpc
            End If
           End If
        End If
    Next j
End Sub

Public Sub AllFollowAmo(ByVal UserIndex As Integer)
    Dim j As Integer
    
    For j = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasIndex(j) > 0 Then
            Call FollowAmo(UserList(UserIndex).MascotasIndex(j))
        End If
    Next j
End Sub

Public Function NpcAtacaUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean

    If UserList(UserIndex).flags.AdminInvisible = 1 Then Exit Function
    If (Not UserList(UserIndex).flags.Privilegios And PlayerType.User) <> 0 And Not UserList(UserIndex).flags.AdminPerseguible Then Exit Function
    
    With Npclist(NpcIndex)
        ' El npc puede atacar ???
        If .CanAttack = 1 Then
            NpcAtacaUser = True
            Call CheckPets(NpcIndex, UserIndex, False)
            
            If .Target = 0 Then .Target = UserIndex
            
            If UserList(UserIndex).flags.AtacadoPorNpc = 0 And UserList(UserIndex).flags.AtacadoPorUser = 0 Then
                UserList(UserIndex).flags.AtacadoPorNpc = NpcIndex
            End If
        Else
            NpcAtacaUser = False
            Exit Function
        End If
        
        .CanAttack = 0
        .AttackTimer = TIMER_ATTACK
        
        If .flags.Snd1 > 0 Then
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(.flags.Snd1, .Pos.X, .Pos.Y))
        End If
    End With
    
    '[ATAK ANIM]
    If Npclist(NpcIndex).Char.WeaponAnim > 0 Then Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageNpcAtaca(NpcIndex, 1))
    
    If NpcImpacto(NpcIndex, UserIndex) Then
        With UserList(UserIndex)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y))
            
            If .flags.Meditando = False Then
                If .flags.Navegando = 0 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FXSANGRE, 0))
                End If
            End If
            
            Call NpcDaño(NpcIndex, UserIndex)
            Call WriteUpdateHP(UserIndex)
            
            '¿Puede envenenar?
            If Npclist(NpcIndex).Veneno = 1 Then Call NpcEnvenenarUser(UserIndex)
        End With
    Else
        Call WriteNPCSwing(UserIndex, Npclist(NpcIndex).NpcType)
    End If
    
    '-----Tal vez suba los skills------
    Call SubirSkill(UserIndex, Tacticas)
    
    'Controla el nivel del usuario
    Call CheckUserLevel(UserIndex)
End Function

Private Function NpcImpactoNpc(ByVal Atacante As Integer, ByVal Victima As Integer) As Boolean
    Dim PoderAtt As Long
    Dim PoderEva As Long
    Dim ProbExito As Long
    
    PoderAtt = Npclist(Atacante).PoderAtaque
    PoderEva = Npclist(Victima).PoderEvasion
    
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + (PoderAtt - PoderEva) * 0.4))
    NpcImpactoNpc = (RandomNumber(1, 100) <= ProbExito)
    
End Function

Public Sub NpcDañoNpc(ByVal Atacante As Integer, ByVal Victima As Integer)
    Dim daño As Integer
    
    With Npclist(Atacante)
        daño = RandomNumber(.Stats.MinHIT, .Stats.MaxHIT)
        Npclist(Victima).Stats.MinHP = Npclist(Victima).Stats.MinHP - daño
        
        If .NpcType = Guardia Then Npclist(Victima).flags.WasAttackByGuards = True
        
        If Npclist(Victima).Stats.MinHP < 1 Then
            .Movement = .flags.OldMovement
            
            If LenB(.flags.AttackedBy) <> 0 Then
                .Hostile = .flags.OldHostil
            End If
            
            If .MaestroUser > 0 Then
                Call FollowAmo(Atacante)
            End If
            
            If .NpcType = Guardia Then
                Npclist(Victima).flags.WasAttackByGuards = True
            End If
            
            Call MuereNpc(Victima, .MaestroUser)
        End If
    End With
End Sub

Public Sub NpcAtacaNpc(ByVal Atacante As Integer, ByVal Victima As Integer, Optional ByVal cambiarMOvimiento As Boolean = True)
'*************************************************
'Author: Unknown
'Last modified: 01/03/2009
'01/03/2009: ZaMa - Las mascotas no pueden atacar al rey si quedan pretorianos vivos.
'*************************************************
    
    With Npclist(Atacante)
        
        'Es el Rey Preatoriano?
        If Npclist(Victima).Numero = PRKING_NPC Then
            If pretorianosVivos > 0 Then
                Call WriteConsoleMsg(.MaestroUser, "Debes matar al resto del ejército antes de atacar al rey!", FontTypeNames.FONTTYPE_FIGHT)
                .TargetNPC = 0
                Exit Sub
            End If
        End If
        
        ' El npc puede atacar ???
        If .CanAttack = 1 Then
            .CanAttack = 0
            .AttackTimer = TIMER_ATTACK
            If cambiarMOvimiento Then
                Npclist(Victima).TargetNPC = Atacante
                Npclist(Victima).Movement = TipoAI.NpcAtacaNpc
            End If
        Else
            Exit Sub
        End If
        
        If .flags.Snd1 > 0 Then
            Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(.flags.Snd1, .Pos.X, .Pos.Y))
        End If
        
        '[ATAK ANIM]
        If Npclist(Atacante).Char.WeaponAnim > 0 Then Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessageNpcAtaca(Atacante, 1))
        
        If NpcImpactoNpc(Atacante, Victima) Then
            If Npclist(Victima).flags.Snd2 > 0 Then
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(Npclist(Victima).flags.Snd2, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
            Else
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO2, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
            End If
        
            If .MaestroUser > 0 Then
                Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y))
            Else
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
            End If
            
            Call NpcDañoNpc(Atacante, Victima)
        Else
            If .MaestroUser > 0 Then
                Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
            Else
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_SWING, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
            End If
        End If
    End With
End Sub

Public Function UsuarioAtacaNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, Optional ByVal HitArea As Byte = 0) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 13/02/2011 (Amraphen)
'12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados por npcs cuando los atacan.
'14/01/2010: ZaMa - Lo transformo en función, para que no se pierdan municiones al atacar targets inválidos.
'13/02/2011: Amraphen - Ahora la stamina es quitada cuando efectivamente se ataca al NPC.
'***************************************************

On Error GoTo errhandler

    If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then Exit Function
    
    Call NPCAtacado(NpcIndex, UserIndex)
    
    If UserImpactoNpc(UserIndex, NpcIndex) Then
        If Npclist(NpcIndex).flags.Snd2 > 0 Then
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(Npclist(NpcIndex).flags.Snd2, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y))
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO2, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y))
        End If
        
        Call UserDañoNpc(UserIndex, NpcIndex, HitArea)
    Else
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SWING, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        Call WriteUserSwing(UserIndex)
    End If
    
    
    '[ATAK ANIM]
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAtaca(UserIndex, 1))
    
    'Quitamos stamina
    Call QuitarSta(UserIndex, RandomNumber(1, 10))
    
    ' Reveló su condición de usuario al atacar, los npcs lo van a atacar
    'UserList(UserIndex).flags.Ignorado = False
    
    UsuarioAtacaNpc = True
    
    Exit Function
    
errhandler:
    Dim UserName As String
    
    If UserIndex > 0 Then UserName = UserList(UserIndex).Name
    
    Call LogError("Error en UsuarioAtacaNpc. Error " & Err.Number & " : " & Err.Description & ". User: " & _
                   UserIndex & "-> " & UserName & ". NpcIndex: " & NpcIndex & ".")
    
End Function

Public Sub UsuarioAtaca(ByVal UserIndex As Integer)
    Dim index As Integer
    Dim AttackPos As WorldPos

    'Check bow's interval
    'If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub

    'Check Spell-Magic interval
    'If Not IntervaloPermiteMagiaGolpe(UserIndex) Then
    'Check Attack interval
    If Not IntervaloPermiteAtacar(UserIndex) Then Exit Sub
    '   End If
    'End If

    With UserList(UserIndex)
        'Quitamos stamina
        If .Stats.MinSta >= 10 Then
            Call QuitarSta(UserIndex, RandomNumber(1, 10))
        Else
            If .genero = eGenero.Hombre Then
                'Call WriteConsoleMsg(UserIndex, "Estas muy cansado para luchar.", FontTypeNames.FONTTYPE_INFO)
                Call WriteMultiMessage(UserIndex, eMessages.EstasMuyCansadoParaLuchar)
            Else
                'Call WriteConsoleMsg(UserIndex, "Estas muy cansada para luchar.", FontTypeNames.FONTTYPE_INFO)
                Call WriteMultiMessage(UserIndex, eMessages.EstasMuyCansadaParaLuchar)
            End If
            Exit Sub
        End If

        AttackPos = .Pos
        Call HeadtoPos(.Char.heading, AttackPos)

        'Exit if not legal
        If AttackPos.X < 1 Or AttackPos.X > MapInfo(UserList(UserIndex).Pos.map).Width Or AttackPos.Y <= 1 Or AttackPos.Y > MapInfo(UserList(UserIndex).Pos.map).Height Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
            Exit Sub
        End If
        'bots
        'Está el bot?
        Dim bot_Index As Byte
   
        bot_Index = MapData(AttackPos.map).Tile(AttackPos.X, AttackPos.Y).BotIndex

        If bot_Index <> 0 Then
            'Checkeo que esté invocado.
            If ia_Bot(bot_Index).Invocado Then
                'compruebo que este en mi grupo
                'If ia_Bot(bot_Index).GrupoID = UserList(UserIndex).Group_User.Grupo_ID Then
                ia_DamageHit bot_Index, UserIndex
                'End If
            End If
        End If
        'bots
        index = MapData(AttackPos.map).Tile(AttackPos.X, AttackPos.Y).UserIndex

        'Look for user
        If index > 0 Then
            Call UsuarioAtacaUsuario(UserIndex, index)
            Call WriteUpdateUserStats(UserIndex)
            Call WriteUpdateUserStats(index)
            Exit Sub
        End If

        index = MapData(AttackPos.map).Tile(AttackPos.X, AttackPos.Y).NpcIndex

        'Look for NPC
        If index > 0 Then
            If Npclist(index).Attackable Then
                If Npclist(index).MaestroUser > 0 And Zonas(Npclist(index).zona).Segura = 1 Then
                    Call WriteConsoleMsg(UserIndex, "No podés atacar mascotas en zonas seguras", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub
                End If

                Call UsuarioAtacaNpc(UserIndex, index)
            Else
                'Call WriteConsoleMsg(UserIndex, "No podés atacar a este NPC", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteMultiMessage(UserIndex, eMessages.NoPodesAtacarEsteNPC)
            End If

            Call WriteUpdateUserStats(UserIndex)

            Exit Sub
        End If
        '[ATAK ANIM]
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAtaca(UserIndex, 1))

        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
        Call WriteUpdateUserStats(UserIndex)

        If .Counters.trabajando Then .Counters.trabajando = .Counters.trabajando - 1

        If .Counters.Ocultando Then .Counters.Ocultando = .Counters.Ocultando - 1
    End With
End Sub

Public Function UsuarioImpacto(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer) As Boolean
    Dim ProbRechazo As Long
    Dim Rechazo As Boolean
    Dim ProbExito As Long
    Dim PoderAtaque As Long
    Dim UserPoderEvasion As Long
    Dim UserPoderEvasionEscudo As Long
    Dim Arma As Integer
    Dim SkillTacticas As Long
    Dim SkillDefensa As Long
    
    SkillTacticas = UserList(VictimaIndex).Stats.UserSkills(eSkill.Tacticas)
    SkillDefensa = UserList(VictimaIndex).Stats.UserSkills(eSkill.Defensa)
    
    Arma = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
    
    'Calculamos el poder de evasion...
    UserPoderEvasion = PoderEvasion(VictimaIndex)
    
    If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then
       UserPoderEvasionEscudo = PoderEvasionEscudo(VictimaIndex)
       UserPoderEvasion = UserPoderEvasion + UserPoderEvasionEscudo
    Else
        UserPoderEvasionEscudo = 0
    End If
    
    'Esta usando un arma ???
    If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
        If ObjData(Arma).proyectil = 1 Then
            PoderAtaque = PoderAtaqueProyectil(AtacanteIndex)
            
        Else
            PoderAtaque = PoderAtaqueArma(AtacanteIndex)
        End If
    Else
        PoderAtaque = PoderAtaqueWrestling(AtacanteIndex)
    End If
    
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + (PoderAtaque - UserPoderEvasion) * 0.4))
    
    UsuarioImpacto = (RandomNumber(1, 100) <= ProbExito)
    
    'Si tiene bonificacion de evasión por raza veo si le pega o no
    If UsuarioImpacto Then
        If ModRaza(UserList(VictimaIndex).raza).EvitarGolpe > 0 Then
            UsuarioImpacto = (RandomNumber(1, 1 / ModRaza(UserList(VictimaIndex).raza).EvitarGolpe) = 1)
        End If
    End If
    ' el usuario esta usando un escudo ???
    If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then
        'Fallo ???
        If Not UsuarioImpacto Then
            ' Chances are rounded
            ProbRechazo = MaximoInt(10, MinimoInt(90, 100 * SkillDefensa / (SkillDefensa + SkillTacticas)))
            Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
            If Rechazo = True Then
                'Se rechazo el ataque con el escudo
                Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageAtaca(VictimaIndex, 2))
                Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessagePlayWave(SND_ESCUDO, UserList(VictimaIndex).Pos.X, UserList(VictimaIndex).Pos.Y))
                  
                Call WriteBlockedWithShieldOther(AtacanteIndex)
                Call WriteBlockedWithShieldUser(VictimaIndex)
                
                Call SubirSkill(VictimaIndex, Defensa)
            End If
        End If
    End If
    
    Call FlushBuffer(VictimaIndex)
End Function

Public Function UsuarioAtacaUsuario(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer, Optional ByVal HitArea As Byte = 0) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 14/01/2010 (ZaMa)
'14/01/2010: ZaMa - Lo transformo en función, para que no se pierdan municiones al atacar targets
'                    inválidos, y evitar un doble chequeo innecesario
'***************************************************

On Error GoTo errhandler

    If Not PuedeAtacar(AtacanteIndex, VictimaIndex) Then Exit Function
    
    With UserList(AtacanteIndex)
        If Distancia(.Pos, UserList(VictimaIndex).Pos) > MAXDISTANCIAARCO Then
           'Call WriteConsoleMsg(AtacanteIndex, "Estás muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
           Call WriteMultiMessage(AtacanteIndex, eMessages.LejosParaDisparar)
           Exit Function
        End If
        
        Call UsuarioAtacadoPorUsuario(AtacanteIndex, VictimaIndex)
        
        Dim UserImpacto As Boolean
        
        If HitArea = 0 Then
            UserImpacto = UsuarioImpacto(AtacanteIndex, VictimaIndex)
        Else
            UserImpacto = True
        End If
        
        If UserImpacto Then
            If HitArea = 0 Then
                Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y))
            Else
                If UserList(VictimaIndex).genero = Hombre Then
                    Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(SND_ALCANZO_EXPLOSION_HOMBRE, .Pos.X, .Pos.Y))
                Else
                    Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(SND_ALCANZO_EXPLOSION_MUJER, .Pos.X, .Pos.Y))
                End If
            End If
            
            If UserList(VictimaIndex).flags.Navegando = 0 Then
                Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.CharIndex, FXSANGRE, 0))
            End If
            
            'Pablo (ToxicWaste): Guantes de Hurto del Bandido en acción
            If .clase = eClass.Bandit Then
                Call DoDesequipar(AtacanteIndex, VictimaIndex)
                
            'y ahora, el ladrón puede llegar a paralizar con el golpe.
            ElseIf .clase = eClass.Thief Then
                Call DoHandInmo(AtacanteIndex, VictimaIndex)
            End If
            
            Call SubirSkill(VictimaIndex, eSkill.Tacticas)
            Call UserDañoUser(AtacanteIndex, VictimaIndex, HitArea)
        Else
            ' Invisible admins doesn't make sound to other clients except itself
            If .flags.AdminInvisible = 1 Then
                Call EnviarDatosASlot(AtacanteIndex, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
            Else
               
                Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
            End If
            
            Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
            Call WriteUserSwing(AtacanteIndex)
            Call WriteUserAttackedSwing(VictimaIndex, AtacanteIndex)
            
            Call SubirSkill(VictimaIndex, eSkill.Tacticas)
        End If
        
        Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessageAtaca(AtacanteIndex, 1))
        
        If .clase = eClass.Thief Then Call Desarmar(AtacanteIndex, VictimaIndex)
    End With
    
    UsuarioAtacaUsuario = True
    
    Exit Function
    
errhandler:
    Call LogError("Error en UsuarioAtacaUsuario. Error " & Err.Number & " : " & Err.Description)
End Function

Public Sub UserDañoUser(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer, Optional ByVal HitArea As Byte = 0)
    Dim daño As Long
    Dim Lugar As Integer
    Dim absorbido As Long
    Dim defbarco As Integer
    Dim Obj As ObjData
    Dim Resist As Byte
    
    daño = CalcularDaño(AtacanteIndex)
    
    Call UserEnvenena(AtacanteIndex, VictimaIndex)
    
    With UserList(AtacanteIndex)
        If .flags.Navegando = 1 And .Invent.BarcoObjIndex > 0 Then
             Obj = ObjData(.Invent.BarcoObjIndex)
             daño = daño + RandomNumber(Obj.MinHIT, Obj.MaxHIT)
        End If
        
        If UserList(VictimaIndex).flags.Navegando = 1 And UserList(VictimaIndex).Invent.BarcoObjIndex > 0 Then
             Obj = ObjData(UserList(VictimaIndex).Invent.BarcoObjIndex)
             defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)
        End If
        
        If .Invent.WeaponEqpObjIndex > 0 Then
            Resist = ObjData(.Invent.WeaponEqpObjIndex).Refuerzo
        End If
        
        Lugar = RandomNumber(PartesCuerpo.bCabeza, PartesCuerpo.bTorso)
        
        Select Case Lugar
            Case PartesCuerpo.bCabeza
                'Si tiene casco absorbe el golpe
                If UserList(VictimaIndex).Invent.CascoEqpObjIndex > 0 Then
                    Obj = ObjData(UserList(VictimaIndex).Invent.CascoEqpObjIndex)
                    absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
                    absorbido = absorbido + defbarco - Resist
                    daño = daño - absorbido
                    If daño < 0 Then daño = 1
                End If
            
            Case Else
                'Si tiene armadura absorbe el golpe
                If UserList(VictimaIndex).Invent.ArmourEqpObjIndex > 0 Then
                    Obj = ObjData(UserList(VictimaIndex).Invent.ArmourEqpObjIndex)
                    Dim Obj2 As ObjData
                    If UserList(VictimaIndex).Invent.EscudoEqpObjIndex Then
                        Obj2 = ObjData(UserList(VictimaIndex).Invent.EscudoEqpObjIndex)
                        absorbido = RandomNumber(Obj.MinDef + Obj2.MinDef, Obj.MaxDef + Obj2.MaxDef)
                    Else
                        absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
                    End If
                    absorbido = absorbido + defbarco - Resist
                    daño = daño - absorbido
                    If daño < 0 Then daño = 1
                End If
        End Select
        
        
        If .flags.Hambre = 0 And .flags.Sed = 0 Then
            'Si usa un arma quizas suba "Combate con armas"
            If .Invent.WeaponEqpObjIndex > 0 Then
                If ObjData(.Invent.WeaponEqpObjIndex).proyectil Then
                    'es un Arco. Sube Armas a Distancia
                    daño = daño * (1 + ModRaza(.raza).Proyectiles)
                    Call SubirSkill(AtacanteIndex, Proyectiles)
                Else
                    'Si tiene bonificacion por raza
                    daño = daño * (1 + ModRaza(.raza).Armas)
                    'Sube combate con armas.
                    Call SubirSkill(AtacanteIndex, Armas)
                End If
            Else
            'sino tal vez lucha libre
                Call SubirSkill(AtacanteIndex, Wrestling)
            End If
                    
            'Trata de apuñalar por la espalda al enemigo
            If PuedeApuñalar(AtacanteIndex) Then
                Call DoApuñalar(AtacanteIndex, 0, VictimaIndex, daño)
                Call SubirSkill(AtacanteIndex, Apuñalar)
            End If
            'e intenta dar un golpe crítico [Pablo (ToxicWaste)]
            'Call DoGolpeCritico(AtacanteIndex, 0, VictimaIndex, daño)
        End If
        
        'HitArea
        If HitArea > 0 Then
            daño = daño - Porcentaje(daño, 15)
        Else
            'If RandomNumber(1, 100) = 1 Then
            '    daño = daño * 1.2
            '
            '    If .genero = Hombre Then
            '        Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(SND_CRITICOH, .Pos.X, .Pos.Y))
            '    Else
            '        Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(SND_CRITICOM, .Pos.X, .Pos.Y))
            '    End If
            'End If
        End If
        
        Call WriteUserHittedUser(AtacanteIndex, Lugar, UserList(VictimaIndex).Char.CharIndex, daño, HitArea)
        Call WriteUserHittedByUser(VictimaIndex, Lugar, .Char.CharIndex, daño, HitArea)
        
        UserList(VictimaIndex).Stats.MinHP = UserList(VictimaIndex).Stats.MinHP - daño
        
        Call SubirSkill(VictimaIndex, Tacticas)
        
        If UserList(VictimaIndex).Stats.MinHP <= 0 Then
            'Store it!
            UserList(VictimaIndex).Stats.MinHP = 0
            Call Statistics.StoreFrag(AtacanteIndex, VictimaIndex)
            
            Call ContarMuerte(VictimaIndex, AtacanteIndex)
            
            ' Para que las mascotas no sigan intentando luchar y
            ' comiencen a seguir al amo
            Dim j As Integer
            For j = 1 To MAXMASCOTAS
                If .MascotasIndex(j) > 0 Then
                    If Npclist(.MascotasIndex(j)).Target = VictimaIndex Then
                        Npclist(.MascotasIndex(j)).Target = 0
                        Call FollowAmo(.MascotasIndex(j))
                    End If
                End If
            Next j
            
            Call ActStats(VictimaIndex, AtacanteIndex)
            Call UserDie(VictimaIndex, AtacanteIndex)
        Else
            'Está vivo - Actualizamos el HP
            Call WriteUpdateHP(VictimaIndex)
        End If
    End With
    
    'Controla el nivel del usuario
    Call CheckUserLevel(AtacanteIndex)
    
    Call FlushBuffer(VictimaIndex)
End Sub

Sub UsuarioAtacadoPorUsuario(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer)
'***************************************************
'Autor: Unknown
'Last Modification: 10/01/08
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
' 10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
'***************************************************
    If UserList(VictimIndex).flags.Meditando Then
        UserList(VictimIndex).flags.Meditando = False
        Call WriteMeditateToggle(VictimIndex)
        Call WriteConsoleMsg(VictimIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
        UserList(VictimIndex).Char.FX = 0
        UserList(VictimIndex).Char.Loops = 0
        Call SendData(SendTarget.ToPCArea, VictimIndex, PrepareMessageCreateFX(UserList(VictimIndex).Char.CharIndex, 0, 0))
    End If
    
    If TriggerZonaPelea(AttackerIndex, VictimIndex) = TRIGGER6_PERMITE Then Exit Sub
    If Zonas(UserList(AttackerIndex).zona).Terreno = eTerreno.Fortaleza And Zonas(UserList(VictimIndex).zona).Terreno = eTerreno.Fortaleza Then Exit Sub
    Dim EraCriminal As Boolean
    
    If Not Criminal(AttackerIndex) And Not Criminal(VictimIndex) Then
        Call VolverCriminal(AttackerIndex)
    End If
    

    
    EraCriminal = Criminal(AttackerIndex)
    
    With UserList(AttackerIndex).Reputacion
        If Not Criminal(VictimIndex) Then
            .BandidoRep = .BandidoRep + vlASALTO
            If .BandidoRep > MAXREP Then .BandidoRep = MAXREP
            
            .NobleRep = .NobleRep / 2
            If .NobleRep < 0 Then .NobleRep = 0
        Else
            .NobleRep = .NobleRep + vlNoble
            If .NobleRep > MAXREP Then .NobleRep = MAXREP
        End If
    End With
    
    If Criminal(AttackerIndex) Then
        If UserList(AttackerIndex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(AttackerIndex)
        
        If Not EraCriminal Then Call RefreshCharStatus(AttackerIndex)
    ElseIf EraCriminal Then
        Call RefreshCharStatus(AttackerIndex)
    End If
    
    Call AllMascotasAtacanUser(AttackerIndex, VictimIndex)
    Call AllMascotasAtacanUser(VictimIndex, AttackerIndex)
    
    'Si la victima esta saliendo se cancela la salida
    Call CancelExit(VictimIndex)
    Call FlushBuffer(VictimIndex)
End Sub

Sub AllMascotasAtacanUser(ByVal victim As Integer, ByVal Maestro As Integer)
    'Reaccion de las mascotas
    Dim iCount As Integer
    
    For iCount = 1 To MAXMASCOTAS
        If UserList(Maestro).MascotasIndex(iCount) > 0 Then
            Npclist(UserList(Maestro).MascotasIndex(iCount)).flags.AttackedBy = UserList(victim).Name
            Npclist(UserList(Maestro).MascotasIndex(iCount)).Movement = TipoAI.NPCDEFENSA
            Npclist(UserList(Maestro).MascotasIndex(iCount)).Hostile = 1
        End If
    Next iCount
End Sub

Public Function PuedeAtacar(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer) As Boolean
'***************************************************
'Autor: Unknown
'Last Modification: 24/02/2009
'Returns true if the AttackerIndex is allowed to attack the VictimIndex.
'24/01/2007 Pablo (ToxicWaste) - Ordeno todo y agrego situacion de Defensa en ciudad Armada y Caos.
'24/02/2009: ZaMa - Los usuarios pueden atacarse entre si.
'***************************************************
    On Error GoTo errhandler

    'MUY importante el orden de estos "IF"...

    'Estas muerto no podes atacar
    'eventos
    If UserList(AttackerIndex).Counters.TimeFight > 0 Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes atacar a nadie antes de la cuenta regresiva.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If

    'Eventos

    If UserList(AttackerIndex).flags.Muerto = 1 Then
        'Call WriteConsoleMsg(attackerIndex, "No podés atacar porque estás muerto.", FontTypeNames.FONTTYPE_INFO)
        Call WriteMultiMessage(AttackerIndex, eMessages.mUserMuerto)
        PuedeAtacar = False
        Exit Function
    End If

    'No podes atacar a alguien muerto
    If UserList(VictimIndex).flags.Muerto = 1 Then
        'Call WriteConsoleMsg(attackerIndex, "No podés atacar a un espíritu.", FontTypeNames.FONTTYPE_INFO)
        Call WriteMultiMessage(AttackerIndex, eMessages.VictimaMuerto)
        PuedeAtacar = False
        Exit Function
    End If

    'eventos
    If UserList(AttackerIndex).flags.SlotEvent > 0 Then
        If Events(UserList(AttackerIndex).flags.SlotEvent).TimeCount > 0 Then
            WriteConsoleMsg AttackerIndex, "No puedes atacar hasta que no termine la cuenta regresiva.", FontTypeNames.FONTTYPE_INFO
            PuedeAtacar = False
            Exit Function
        End If

        If Events(UserList(AttackerIndex).flags.SlotEvent).Users(UserList(AttackerIndex).flags.SlotUserEvent).Team > 0 Then

            If Not CanAttackUserEvent(AttackerIndex, VictimIndex) Then
                WriteConsoleMsg AttackerIndex, "No puedes atacar a tu compañero", FontTypeNames.FONTTYPE_INFO
                PuedeAtacar = False
                Exit Function
            End If

        End If
    End If
    'eventos

    'Estamos en una Arena? o un trigger zona segura?
    Select Case TriggerZonaPelea(AttackerIndex, VictimIndex)
    Case eTrigger6.TRIGGER6_PERMITE
        PuedeAtacar = True
        Exit Function

    Case eTrigger6.TRIGGER6_PROHIBE
        PuedeAtacar = False
        Exit Function

    Case eTrigger6.TRIGGER6_AUSENTE
        'Si no estamos en el Trigger 6 entonces es imposible atacar un gm
        If (UserList(VictimIndex).flags.Privilegios And PlayerType.User) = 0 And AttackerIndex <> VictimIndex Then
            If UserList(VictimIndex).flags.AdminInvisible = 0 Then Call WriteConsoleMsg(AttackerIndex, "El ser es demasiado poderoso", FontTypeNames.FONTTYPE_WARNING)
            PuedeAtacar = False
            Exit Function
        End If
    End Select

    If Zonas(UserList(AttackerIndex).zona).Terreno = eTerreno.Fortaleza And Zonas(UserList(VictimIndex).zona).Terreno = eTerreno.Fortaleza Then
        If UserList(AttackerIndex).GuildIndex > 0 And UserList(AttackerIndex).GuildIndex = UserList(VictimIndex).GuildIndex Then
            Call WriteConsoleMsg(AttackerIndex, "No puedes atacar a los miembros de tu clan dentro de los dominios de la fortaleza.", FontTypeNames.FONTTYPE_WARNING)
            PuedeAtacar = False
        Else
            PuedeAtacar = True
        End If
        Exit Function
    End If

    'Sos un Armada atacando un ciudadano?
    If (Not Criminal(VictimIndex)) And (esArmada(AttackerIndex)) Then
        Call WriteConsoleMsg(AttackerIndex, "Los soldados del Ejército Real tienen prohibido atacar ciudadanos.", FontTypeNames.FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function
    End If

    'Sos un Caos atacando otro caos?
    If esCaos(VictimIndex) And esCaos(AttackerIndex) Then
        Call WriteConsoleMsg(AttackerIndex, "Los miembros de la legión oscura tienen prohibido atacarse entre sí.", FontTypeNames.FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function
    End If

    'Tenes puesto el seguro?
    If UserList(AttackerIndex).flags.Seguro Then
        If Not Criminal(VictimIndex) Then
            Call WriteConsoleMsg(AttackerIndex, "No podes atacar ciudadanos, para hacerlo debes desactivar el seguro ingresando /SEG", FontTypeNames.FONTTYPE_WARNING)
            PuedeAtacar = False
            Exit Function
        End If
    End If

    'Estas en un Mapa Seguro?
    If Zonas(UserList(VictimIndex).zona).Segura = 1 Then
        If esArmada(AttackerIndex) Then
            If UserList(AttackerIndex).Faccion.RecompensasReal > 11 Then
                If UserList(VictimIndex).Pos.map = 58 Or UserList(VictimIndex).Pos.map = 59 Or UserList(VictimIndex).Pos.map = 60 Then    'TODO JAVIER
                    Call WriteConsoleMsg(VictimIndex, "¡Huye de la ciudad! estas siendo atacado y no podrás defenderte.", FontTypeNames.FONTTYPE_WARNING)
                    PuedeAtacar = True    'Beneficio de Armadas que atacan en su ciudad.
                    Exit Function
                End If
            End If
        End If
        If esCaos(AttackerIndex) Then
            If UserList(AttackerIndex).Faccion.RecompensasCaos > 11 Then
                If UserList(VictimIndex).Pos.map = 151 Or UserList(VictimIndex).Pos.map = 156 Then    'TODO JAVIER
                    Call WriteConsoleMsg(VictimIndex, "¡Huye de la ciudad! estas siendo atacado y no podrás defenderte.", FontTypeNames.FONTTYPE_WARNING)
                    PuedeAtacar = True    'Beneficio de Caos que atacan en su ciudad.
                    Exit Function
                End If
            End If
        End If
        Call WriteConsoleMsg(AttackerIndex, "Esta es una zona segura, aquí no se puede atacar otros usuarios.", FontTypeNames.FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function
    ElseIf Zonas(UserList(AttackerIndex).zona).Segura = 1 Then
        Call WriteConsoleMsg(AttackerIndex, "¡Estás es una zona segura!", FontTypeNames.FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function
    End If

    'Estas atacando desde un trigger seguro? o tu victima esta en uno asi?
    If MapData(UserList(VictimIndex).Pos.map).Tile(UserList(VictimIndex).Pos.X, UserList(VictimIndex).Pos.Y).Trigger = eTrigger.ZONASEGURA Or _
       MapData(UserList(AttackerIndex).Pos.map).Tile(UserList(AttackerIndex).Pos.X, UserList(AttackerIndex).Pos.Y).Trigger = eTrigger.ZONASEGURA Then
        'Call WriteConsoleMsg(attackerIndex, "No podés pelear aquí.", FontTypeNames.FONTTYPE_WARNING)
        Call WriteMultiMessage(AttackerIndex, eMessages.NoPodesPelearAqui)
        PuedeAtacar = False
        Exit Function
    End If

    PuedeAtacar = True
    Exit Function

errhandler:
    Call LogError("Error en PuedeAtacar. Error " & Err.Number & " : " & Err.Description)
End Function

Public Function PuedeAtacarNPC(ByVal AttackerIndex As Integer, ByVal NpcIndex As Integer) As Boolean
'***************************************************
'Autor: Unknown Author (Original version)
'Returns True if AttackerIndex can attack the NpcIndex
'Last Modification: 24/01/2007
'24/01/2007 Pablo (ToxicWaste) - Orden y corrección de ataque sobre una mascota y guardias
'14/08/2007 Pablo (ToxicWaste) - Reescribo y agrego TODOS los casos posibles cosa de usar
'esta función para todo lo referente a ataque a un NPC. Ya sea Magia, Físico o a Distancia.
'***************************************************
'Estas muerto?
    If UserList(AttackerIndex).flags.Muerto = 1 Then
        'Call WriteConsoleMsg(attackerIndex, "No podés atacar porque estás muerto.", FontTypeNames.FONTTYPE_INFO)
        Call WriteMultiMessage(AttackerIndex, eMessages.mUserMuerto)
        PuedeAtacarNPC = False
        Exit Function
    End If

    'Sos consejero?
    If UserList(AttackerIndex).flags.Privilegios And PlayerType.Consejero Then
        'No pueden atacar NPC los Consejeros.
        PuedeAtacarNPC = False
        Exit Function
    End If

    'Es una criatura atacable?
    If Npclist(NpcIndex).Attackable = 0 Then
        'Call WriteConsoleMsg(attackerIndex, "No puedes atacar esta criatura.", FontTypeNames.FONTTYPE_INFO)
        Call WriteMultiMessage(AttackerIndex, eMessages.NoPodesAtacarEsteNPC)
        PuedeAtacarNPC = False
        Exit Function
    End If

    'Es valida la distancia a la cual estamos atacando?
    If Distancia(UserList(AttackerIndex).Pos, Npclist(NpcIndex).Pos) >= MAXDISTANCIAARCO Then
        'Call WriteConsoleMsg(attackerIndex, "Estás muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
        Call WriteMultiMessage(AttackerIndex, eMessages.LejosParaDisparar)
        PuedeAtacarNPC = False
        Exit Function
    End If

    If Zonas(UserList(AttackerIndex).zona).Segura <> Zonas(Npclist(NpcIndex).zona).Segura Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes atacar a esta criatura fuera de la zona segura.", FontTypeNames.FONTTYPE_FIGHT)
        PuedeAtacarNPC = False
        Exit Function
    End If

    If Zonas(Npclist(NpcIndex).zona).Segura And Npclist(NpcIndex).NpcType = eNPCType.Mercader Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes atacar al personaje dentro de una zona segura.", FontTypeNames.FONTTYPE_FIGHT)
        PuedeAtacarNPC = False
        Exit Function
    End If

    'Es una criatura No-Hostil?
    If Npclist(NpcIndex).Hostile = 0 Then
        'Es Guardia del Caos?
        If EsGuardiaCaos(NpcIndex) Or EsMercader(NpcIndex, False) Then
            'Lo quiere atacar un caos?
            If esCaos(AttackerIndex) Then
                'Neo ataque de npcs prohibidos
                If EsMercader(NpcIndex, False) Then
                    Call WriteConsoleMsg(AttackerIndex, "No puedes atacar a un aliado del Caos siendo Legionario", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(AttackerIndex, "No puedes atacar guardias del Caos siendo Legionario", FontTypeNames.FONTTYPE_INFO)
                End If
                
                PuedeAtacarNPC = False
                Exit Function
            End If
            'Es guardia Real?
        ElseIf EsGuardiaReal(NpcIndex) Or EsMercader(NpcIndex, True) Then
            'Lo quiere atacar un Armada?
            If esArmada(AttackerIndex) Then
                'Neo
                If EsMercader(NpcIndex, True) Then
                    Call WriteConsoleMsg(AttackerIndex, "No puedes atacar a un aliado real siendo de la Armada Real", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(AttackerIndex, "No puedes atacar guardias reales siendo de la armada real", FontTypeNames.FONTTYPE_INFO)
                End If
                
                PuedeAtacarNPC = False
                Exit Function
            End If
            'Tienes el seguro puesto?
            If UserList(AttackerIndex).flags.Seguro Then
                'Neo
                If EsMercader(NpcIndex, True) Then
                    Call WriteConsoleMsg(AttackerIndex, "Debes quitar el seguro para poder atacar, utilizando /SEG", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(AttackerIndex, "Debes quitar el seguro para poder atacar guardias reales utilizando /SEG", FontTypeNames.FONTTYPE_INFO)
                End If
                
                PuedeAtacarNPC = False
                Exit Function
            Else
                'Neo
                If EsMercader(NpcIndex, True) Then
                    Call WriteConsoleMsg(AttackerIndex, "¡Atacaste a un aliado de la Armada Real! Ahora eres un Criminal.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(AttackerIndex, "¡Atacaste un Guardia Real! Eres un Criminal.", FontTypeNames.FONTTYPE_INFO)
                End If
                
                Call VolverCriminal(AttackerIndex)
                PuedeAtacarNPC = True
                Exit Function
            End If
        ElseIf EsGuardiaClan(NpcIndex) Then
            If FortalezaDelClan(UserList(AttackerIndex).GuildIndex, NpcIndex) Then
                Call WriteConsoleMsg(AttackerIndex, "No puedes atacar a los protectores de la fortaleza de tu clan.", FontTypeNames.FONTTYPE_FIGHT)
                PuedeAtacarNPC = False
                Exit Function
            Else
                Call AvisarAtaqueFortaleza(NpcIndex, AttackerIndex)
            End If
            'No era un Guardia, asi que es una criatura No-Hostil común.
            'Para asegurarnos que no sea una Mascota:
        ElseIf Npclist(NpcIndex).MaestroUser = 0 Then
            'Si sos ciudadano tenes que quitar el seguro para atacarla.
            If Not Criminal(AttackerIndex) Then
                'Sos ciudadano, tenes el seguro puesto?
                If UserList(AttackerIndex).flags.Seguro Then
                    Call WriteConsoleMsg(AttackerIndex, "Para atacar a este NPC debés quitar el seguro", FontTypeNames.FONTTYPE_INFO)
                    PuedeAtacarNPC = False
                    Exit Function
                Else
                    'No tiene seguro puesto. Puede atacar pero es penalizado.
                    Call WriteConsoleMsg(AttackerIndex, "Atacaste un NPC No-Hostil. Continúa haciendolo y serás Criminal.", FontTypeNames.FONTTYPE_INFO)
                    'NicoNZ: Cambio para que al atacar npcs no hostiles no bajen puntos de nobleza
                    'Call DisNobAuBan(attackerIndex, 1000, 1000)
                    Call DisNobAuBan(AttackerIndex, 0, 1000)
                    PuedeAtacarNPC = True
                    Exit Function
                End If
            End If
        End If
    End If

    'Es el NPC mascota de alguien?
    If Npclist(NpcIndex).MaestroUser > 0 Then
        If Not Criminal(Npclist(NpcIndex).MaestroUser) Then
            'Es mascota de un Ciudadano.
            If esArmada(AttackerIndex) Then
                'El atacante es Armada y esta intentando atacar mascota de un Ciudadano
                Call WriteConsoleMsg(AttackerIndex, "Los miembros de la armada real no pueden atacar mascotas de ciudadanos.", FontTypeNames.FONTTYPE_INFO)
                PuedeAtacarNPC = False
                Exit Function
            End If
            If Not Criminal(AttackerIndex) Then
                'El atacante es Ciudadano y esta intentando atacar mascota de un Ciudadano.
                If UserList(AttackerIndex).flags.Seguro Then
                    'El atacante tiene el seguro puesto. No puede atacar.
                    Call WriteConsoleMsg(AttackerIndex, "Para atacar mascotas de ciudadanos debes quitar el seguro utilizando /SEG", FontTypeNames.FONTTYPE_INFO)
                    PuedeAtacarNPC = False
                    Exit Function
                Else
                    'El atacante no tiene el seguro puesto. Recibe penalización.
                    Call WriteConsoleMsg(AttackerIndex, "Has atacado la mascota de un Ciudadano. Eres un Criminal.", FontTypeNames.FONTTYPE_INFO)
                    Call VolverCriminal(AttackerIndex)
                    PuedeAtacarNPC = True
                    Exit Function
                End If
            Else
                'El atacante es criminal y quiere atacar un elemental ciuda, pero tiene el seguro puesto (NicoNZ)
                If UserList(AttackerIndex).flags.Seguro Then
                    Call WriteConsoleMsg(AttackerIndex, "Para atacar mascotas de Ciudadanos debes quitar el seguro utilizando /SEG", FontTypeNames.FONTTYPE_INFO)
                    PuedeAtacarNPC = False
                    Exit Function
                End If
            End If
        Else
            'Es mascota de un Criminal.
            If esCaos(Npclist(NpcIndex).MaestroUser) Then
                'Es Caos el Dueño.
                If esCaos(AttackerIndex) Then
                    'Un Caos intenta atacar una criatura de un Caos. No puede atacar.
                    Call WriteConsoleMsg(AttackerIndex, "Los miembros de la Legión Oscura no pueden atacar mascotas de otros legionarios. ", FontTypeNames.FONTTYPE_INFO)
                    PuedeAtacarNPC = False
                    Exit Function
                End If
            End If
        End If
    End If
    'eventos
    If (UserList(AttackerIndex).flags.SlotEvent) > 0 And (Npclist(NpcIndex).flags.TeamEvent > 0) Then
        If Events(UserList(AttackerIndex).flags.SlotEvent).Modality = CastleMode Then
            If Not EventosAOyin.CanAttackReyCastle(AttackerIndex, NpcIndex) Then
                WriteConsoleMsg AttackerIndex, "No puedes atacar a tu rey", FontTypeNames.FONTTYPE_FIGHT
                Exit Function
            End If

        End If

    End If
    'eventos

    'Es el Rey Preatoriano?
    If esPretoriano(NpcIndex) = 4 Then
        If pretorianosVivos > 0 Then
            Call WriteConsoleMsg(AttackerIndex, "Debes matar al resto del ejército antes de atacar al rey.", FontTypeNames.FONTTYPE_FIGHT)
            PuedeAtacarNPC = False
            Exit Function
        End If
    End If

    PuedeAtacarNPC = True
End Function

Sub CalcularDarExp(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal ElDaño As Long)
'***************************************************
'Autor: Nacho (Integer)
'Last Modification: 03/09/06 Nacho
'Reescribi gran parte del Sub
'Ahora, da toda la experiencia del npc mientras este vivo.
'***************************************************
    Dim ExpaDar As Long
    
    '[Nacho] Chekeamos que las variables sean validas para las operaciones
    If ElDaño <= 0 Then ElDaño = 0
    If Npclist(NpcIndex).Stats.MaxHP <= 0 Then Exit Sub
    If ElDaño > Npclist(NpcIndex).Stats.MinHP Then ElDaño = Npclist(NpcIndex).Stats.MinHP
    
    '[Nacho] La experiencia a dar es la porcion de vida quitada * toda la experiencia
    ExpaDar = CLng(ElDaño * (Npclist(NpcIndex).GiveEXP * 0.85 / Npclist(NpcIndex).Stats.MaxHP))
    If ExpaDar <= 0 Then Exit Sub
    
    '[Nacho] Vamos contando cuanta experiencia sacamos, porque se da toda la que no se dio al user que mata al NPC
            'Esto es porque cuando un elemental ataca, no se da exp, y tambien porque la cuenta que hicimos antes
            'Podria dar un numero fraccionario, esas fracciones se acumulan hasta formar enteros ;P
    If ExpaDar > Npclist(NpcIndex).flags.ExpCount Then
        ExpaDar = Npclist(NpcIndex).flags.ExpCount
        Npclist(NpcIndex).flags.ExpCount = 0
    Else
        Npclist(NpcIndex).flags.ExpCount = Npclist(NpcIndex).flags.ExpCount - ExpaDar
    End If
    
    '[Nacho] Le damos la exp al user
    If ExpaDar > 0 Then
        If UserList(UserIndex).PartyIndex > 0 Then
            Call mdParty.ObtenerExito(UserIndex, ExpaDar, Npclist(NpcIndex).Pos.map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y)
        Else
            UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpaDar
            If UserList(UserIndex).Stats.Exp > MAXEXP Then _
                UserList(UserIndex).Stats.Exp = MAXEXP
            Call WriteConsoleMsg(UserIndex, "Has ganado " & ExpaDar & " puntos de experiencia.", FontTypeNames.FONTTYPE_EXP)
        End If
        
        Call CheckUserLevel(UserIndex)
    End If
End Sub

Public Function TriggerZonaPelea(ByVal Origen As Integer, ByVal Destino As Integer) As eTrigger6
'TODO: Pero que rebuscado!!
'Nigo:  Te lo rediseñe, pero no te borro el TODO para que lo revises.
On Error GoTo errhandler
    Dim tOrg As eTrigger
    Dim tDst As eTrigger
    
    tOrg = MapData(UserList(Origen).Pos.map).Tile(UserList(Origen).Pos.X, UserList(Origen).Pos.Y).Trigger
    tDst = MapData(UserList(Destino).Pos.map).Tile(UserList(Destino).Pos.X, UserList(Destino).Pos.Y).Trigger
    
    If tOrg = eTrigger.ZONAPELEA Or tDst = eTrigger.ZONAPELEA Then
        If tOrg = tDst Then
            TriggerZonaPelea = TRIGGER6_PERMITE
        Else
            TriggerZonaPelea = TRIGGER6_PROHIBE
        End If
    Else
        TriggerZonaPelea = TRIGGER6_AUSENTE
    End If

Exit Function
errhandler:
    TriggerZonaPelea = TRIGGER6_AUSENTE
    LogError ("Error en TriggerZonaPelea - " & Err.Description)
End Function

Sub UserEnvenena(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
    Dim ObjInd As Integer
    
    ObjInd = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
    
    If ObjInd > 0 Then
        If ObjData(ObjInd).proyectil = 1 Then
            ObjInd = UserList(AtacanteIndex).Invent.MunicionEqpObjIndex
        End If
        
        If ObjInd > 0 Then
            If ObjData(ObjInd).Envenena = 1 Then
                
                If RandomNumber(1, 100) < 60 Then
                    UserList(VictimaIndex).flags.Envenenado = 1
                    Call WriteConsoleMsg(VictimaIndex, UserList(AtacanteIndex).Name & " te ha envenenado!!", FontTypeNames.FONTTYPE_FIGHT)
                    Call WriteConsoleMsg(AtacanteIndex, "Has envenenado a " & UserList(VictimaIndex).Name & "!!", FontTypeNames.FONTTYPE_FIGHT)
                End If
            End If
        End If
    End If
    
    Call FlushBuffer(VictimaIndex)
End Sub


Public Sub LanzarProyectil(ByVal UserIndex As Integer, ByVal X As Integer, ByVal Y As Integer)
'***************************************************
'Autor: ZaMa
'Last Modification: 10/07/2010
'Throws an arrow or knive to target user/npc.
'***************************************************
On Error GoTo errhandler

    Dim MunicionSlot As Byte
    Dim MunicionIndex As Integer
    Dim WeaponSlot As Byte
    Dim WeaponIndex As Integer

    Dim TargetUserIndex As Integer
    Dim TargetNpcIndex As Integer

    Dim DummyInt As Integer
    
    Dim Threw As Boolean
    Threw = True
    
    'Make sure the item is valid and there is ammo equipped.
    With UserList(UserIndex)
        
        With .Invent
            MunicionSlot = .MunicionEqpSlot
            MunicionIndex = .MunicionEqpObjIndex
            WeaponSlot = .WeaponEqpSlot
            WeaponIndex = .WeaponEqpObjIndex
        End With
        
        ' Tiene arma equipada?
        If WeaponIndex = 0 Then
            DummyInt = 1
            Call WriteConsoleMsg(UserIndex, "No tienes un arma que lanza proyectiles equipada.", FontTypeNames.FONTTYPE_INFO)
            
        ' En un slot válido?
        ElseIf WeaponSlot < 1 Or WeaponSlot > .CurrentInventorySlots Then
            DummyInt = 1
            Call WriteConsoleMsg(UserIndex, "No tienes un arma que lanza proyectiles equipada.", FontTypeNames.FONTTYPE_INFO)
            
        ' Usa munición? (Si no la usa, puede ser un arma arrojadiza)
        ElseIf ObjData(WeaponIndex).Municion = 1 Then
        
            ' La municion esta equipada en un slot valido?
            If MunicionSlot < 1 Or MunicionSlot > .CurrentInventorySlots Then
                DummyInt = 1
                Call WriteConsoleMsg(UserIndex, "No tienes municiones equipadas.", FontTypeNames.FONTTYPE_INFO)
                
            ' Tiene munición?
            ElseIf MunicionIndex = 0 Then
                DummyInt = 1
                Call WriteConsoleMsg(UserIndex, "No tienes municiones equipadas.", FontTypeNames.FONTTYPE_INFO)
                
            ' Son flechas?
            ElseIf ObjData(MunicionIndex).OBJType <> eOBJType.otFlechas Then
                DummyInt = 1
                Call WriteConsoleMsg(UserIndex, "No tienes municiones.", FontTypeNames.FONTTYPE_INFO)
                
            ' Tiene suficientes?
            ElseIf .Invent.Object(MunicionSlot).Amount < 1 Then
                DummyInt = 1
                Call WriteConsoleMsg(UserIndex, "No tienes municiones.", FontTypeNames.FONTTYPE_INFO)
            End If
            
        ' Es un arma de proyectiles?
        ElseIf ObjData(WeaponIndex).proyectil <> 1 Then
            DummyInt = 2
        End If
        
        If DummyInt <> 0 Then
            If DummyInt = 1 Then
                Call Desequipar(UserIndex, WeaponSlot, False)
            End If
            
            Call Desequipar(UserIndex, MunicionSlot, True)
            Exit Sub
        End If
    
        'Quitamos stamina
        If .Stats.MinSta >= 10 Then
            Call QuitarSta(UserIndex, RandomNumber(1, 10))
        Else
            If .genero = eGenero.Hombre Then
                'Call WriteConsoleMsg(UserIndex, "Estás muy cansado para luchar.", FontTypeNames.FONTTYPE_INFO)
                Call WriteMultiMessage(UserIndex, eMessages.EstasMuyCansadoParaLuchar)
            Else
                'Call WriteConsoleMsg(UserIndex, "Estás muy cansada para luchar.", FontTypeNames.FONTTYPE_INFO)
                Call WriteMultiMessage(UserIndex, eMessages.EstasMuyCansadaParaLuchar)
            End If
            Exit Sub
        End If
        
        Dim motivo As String
        If Not CanUse(UserIndex, WeaponIndex, motivo) Then
            Call WriteConsoleMsg(UserIndex, motivo, FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        
        Call LookatTile(UserIndex, .Pos.map, X, Y)
        
        TargetUserIndex = .flags.TargetUser
        TargetNpcIndex = .flags.TargetNPC
        
        If MunicionIndex > 0 Then
            If .Invent.WeaponEqpObjIndex > 0 And .Invent.WeaponEqpObjIndex <> BALA_CANON Then  'BALA
                If .flags.Oculto = 0 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateEfecto(UserList(UserIndex).Char.CharIndex, 0, 0, 1, 0, ObjData(.Invent.MunicionEqpObjIndex).Efecto, .Pos.X, .Pos.Y, X, Y))
                End If
            End If
        End If
        
        Dim Slot As Byte
        If ObjData(WeaponIndex).HitArea > 0 Then
                Call HitArea(UserIndex, ObjData(WeaponIndex).HitArea, .Pos.map, X, Y)
                'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateAreaFX(x - 1, Y - 1, 7, 0))
                'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(27, x, Y))
                
                ' Tiene equipado arco y flecha?
                If ObjData(WeaponIndex).Municion = 1 Then
                    Slot = MunicionSlot
                ' Tiene equipado un arma arrojadiza
                Else
                    Slot = WeaponSlot
                End If
                
                'Take 1 knife/arrow away
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call UpdateUserInv(False, UserIndex, Slot)
                
            Exit Sub
        End If
        
        
        'Validate target
        If TargetUserIndex > 0 Then
            'Only allow to atack if the other one can retaliate (can see us)
            If Abs(UserList(TargetUserIndex).Pos.Y - .Pos.Y) > RANGO_VISION_Y Then
                'Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para atacar.", FontTypeNames.FONTTYPE_WARNING)
                Call WriteMultiMessage(UserIndex, eMessages.LejosParaAtacar)
                Exit Sub
            End If
            
            'Prevent from hitting self
            If TargetUserIndex = UserIndex Then
                'Call WriteConsoleMsg(UserIndex, "¡No puedes atacarte a vos mismo!", FontTypeNames.FONTTYPE_INFO)
                Call WriteMultiMessage(UserIndex, eMessages.NoPodesAtacarteAVosMismo)
                Exit Sub
            End If
            
            'Attack!
            Threw = UsuarioAtacaUsuario(UserIndex, TargetUserIndex)
            
        ElseIf TargetNpcIndex > 0 Then
            'Only allow to atack if the other one can retaliate (can see us)
            If Abs(Npclist(TargetNpcIndex).Pos.Y - .Pos.Y) > RANGO_VISION_Y And Abs(Npclist(TargetNpcIndex).Pos.X - .Pos.X) > RANGO_VISION_X Then
                'Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para atacar.", FontTypeNames.FONTTYPE_WARNING)
                Call WriteMultiMessage(UserIndex, eMessages.LejosParaAtacar)
                Exit Sub
            End If
            
            'Is it attackable???
            If Npclist(TargetNpcIndex).Attackable <> 0 Then
                'Attack!
                Threw = UsuarioAtacaNpc(UserIndex, TargetNpcIndex)
            End If
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
        End If
        
        ' Solo pierde la munición si pudo atacar al target, o tiro al aire
        If Threw Then
            
            ' Tiene equipado arco y flecha?
            If ObjData(WeaponIndex).Municion = 1 Then
                Slot = MunicionSlot
            ' Tiene equipado un arma arrojadiza
            Else
                Slot = WeaponSlot
            End If
            
            'Take 1 knife/arrow away
            Call QuitarUserInvItem(UserIndex, Slot, 1)
            Call UpdateUserInv(False, UserIndex, Slot)
        End If
        
    End With
    
    Exit Sub

errhandler:

    Dim UserName As String
    If UserIndex > 0 Then UserName = UserList(UserIndex).Name

    Call LogError("Error en LanzarProyectil " & Err.Number & ": " & Err.Description & _
                  ". User: " & UserName & "(" & UserIndex & ")")

End Sub



