Attribute VB_Name = "EventosAOyin"
Option Explicit

Public Const MAX_EVENT_SIMULTANEO As Byte = 5
Public Const MAX_USERS_EVENT As Byte = 64
Public Const MAX_MAP_FIGHT As Byte = 4
Public Const MAP_TILE_VS As Byte = 15

Public Enum eModalityEvent
    CastleMode = 1
    DagaRusa = 2
    DeathMatch = 3
    Enfrentamientos = 4
End Enum

Private Type tUserEvent
    Id As Integer
    Team As Byte
    Value As Integer
    Selected As Byte
    MapFight As Integer
End Type


Private Type tEvents
    Enabled As Boolean
    Run As Boolean
    Modality As eModalityEvent
    TeamCant As Byte
    
    Quotas As Byte
    Inscribed As Byte
    
    LvlMax As Byte
    LvlMin As Byte
    
    GldInscription As Long
    DspInscription As Long
    CanjeInscription As Long
    
    AllowedClasses() As Byte
    TimeInit As Long
    TimeCancel As Long
    TimeCount As Long
    TimeFinish As Long
    
    Users() As tUserEvent
    
    ' Por si alguno es con NPC
    NpcIndex As Integer
    
    ' Por si cambia el body del personaje y saca todo lo otro.
    CharBody As Integer
    CharHp As Integer
    
    npcUserIndex As Integer
End Type

Public Events(1 To MAX_EVENT_SIMULTANEO) As tEvents

Private Type tMap
    Run As Boolean
    map As Integer
    X As Byte
    Y As Byte
End Type

Private Type tMapEvent
    Fight(1 To MAX_MAP_FIGHT) As tMap
End Type

Private MapEvent As tMapEvent

Public Sub LoadMapEvent()
10        With MapEvent
20            .Fight(1).Run = False
30            .Fight(1).map = 24
40            .Fight(1).X = 81 '+7
50            .Fight(1).Y = 8 '+24
              
60            .Fight(2).Run = False
70            .Fight(2).map = 24
80            .Fight(2).X = 81 '+7
90            .Fight(2).Y = 34 '+7

100           .Fight(3).Run = False
110           .Fight(3).map = 24
120           .Fight(3).X = 81 '+24
130           .Fight(3).Y = 55 '+50
              
140           .Fight(4).Run = False
150           .Fight(4).map = 24
160           .Fight(4).X = 107 '+33
170           .Fight(4).Y = 8 '+50
          
          
          
180       End With
End Sub
          
          

'/MANEJO DE LOS TIEMPOS '/
Public Sub LoopEvent()
    Dim LoopC As Long
    Dim loopY As Integer
    
    For LoopC = 1 To MAX_EVENT_SIMULTANEO
        With Events(LoopC)
            If .Enabled Then
                If .TimeInit > 0 Then
                    .TimeInit = .TimeInit - 1
                        
                    Select Case .TimeInit
                        Case 0
                            SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(LoopC, .Modality) & "> Las inscripciones están abiertas", FontTypeNames.FONTTYPE_GM)
                          
                        Case 60
                            SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(LoopC, .Modality) & "> Las inscripciones abren en " & Int(.TimeInit / 60) & " minutos.", FontTypeNames.FONTTYPE_GUILD)
                        Case 120
                            SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(LoopC, .Modality) & "> Las inscripciones abren en " & Int(.TimeInit / 60) & " minutos.", FontTypeNames.FONTTYPE_GUILD)
                        Case 180
                            SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(LoopC, .Modality) & "> Las inscripciones abren en " & Int(.TimeInit / 60) & " minutos.", FontTypeNames.FONTTYPE_GUILD)
                        
                    End Select
                    
                    If .TimeInit = 0 Then
                        SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(LoopC, .Modality) & "» Inscripciones abiertas. /PARTICIPAR " & strModality(LoopC, .Modality) & " para ingresar al evento.", FontTypeNames.FONTTYPE_GUILD)
                        '.TimeCancel = 0
                    End If
                    
                
                End If
                
                
                
                If .TimeCancel > 0 Then
                    .TimeCancel = .TimeCancel - 1
                    
                    If .TimeCancel <= 0 Then
                        'SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(.Modality) & "» Ha sido cancelado ya que no se completaron los cupos.", FontTypeNames.FONTTYPE_WARNING)
                        EventosAOyin.CloseEvent LoopC, "Evento " & strModality(LoopC, .Modality) & " cancelado."
                    End If
                End If
                
                If .TimeCount > 0 Then
                    .TimeCount = .TimeCount - 1
                    
                    For loopY = LBound(.Users()) To UBound(.Users())
                        If .Users(loopY).Id > 0 Then
                            If .TimeCount = 0 Then
                                WriteConsoleMsg .Users(loopY).Id, "El juego ha comenzado!", FontTypeNames.FONTTYPE_GUILD
                            Else
                                WriteConsoleMsg .Users(loopY).Id, "CONTEO> " & .TimeCount, FontTypeNames.FONTTYPE_GM
                            End If
                        End If
                    Next loopY
                End If
                
                If .NpcIndex > 0 Then
                   If Events(Npclist(.NpcIndex).flags.SlotEvent).TimeCount > 0 Then Exit Sub
                   Call DagaRusa_MoveNpc(.NpcIndex)
                End If
                
                If .TimeFinish > 0 Then
                    .TimeFinish = .TimeFinish - 1
                    
                    If .TimeFinish = 0 Then
                        Call FinishEvent(LoopC)
                    End If
                End If
            End If
    
    
        End With
    Next LoopC
End Sub

'/ FIN MANEJO DE LOS TIEMPOS


'// Funciones generales '//
Private Function FreeSlotEvent() As Byte
    Dim LoopC As Integer
    
    For LoopC = 1 To MAX_EVENT_SIMULTANEO
        If Not Events(LoopC).Enabled Then
            FreeSlotEvent = LoopC
            Exit For
        End If
    Next LoopC
End Function

Private Function FreeSlotUser(ByVal SlotEvent As Byte) As Byte
    Dim LoopC As Integer
    
    With Events(SlotEvent)
        For LoopC = 1 To MAX_USERS_EVENT
            If .Users(LoopC).Id = 0 Then
                FreeSlotUser = LoopC
                Exit For
            End If
        Next LoopC
    End With
    
End Function
Public Function strUsersEvent(ByVal SlotEvent As Byte) As String

    ' Texto que marca los personajes que están en el evento.
    Dim LoopC As Integer
    
    With Events(SlotEvent)
        For LoopC = LBound(.Users()) To UBound(.Users())
            If .Users(LoopC).Id > 0 Then
                strUsersEvent = strUsersEvent & UserList(.Users(LoopC).Id).Name & "-"
            End If
        Next LoopC
    End With
End Function
Private Function CheckAllowedClasses(ByRef AllowedClasses() As Byte) As String
    Dim LoopC As Integer
    
    For LoopC = 1 To NUMCLASES
        If AllowedClasses(LoopC) = 1 Then
            If CheckAllowedClasses = vbNullString Then
                CheckAllowedClasses = ListaClases(LoopC)
            Else
                CheckAllowedClasses = CheckAllowedClasses & ", " & ListaClases(LoopC)
            End If
        End If
    Next LoopC
    
End Function

Private Function SearchLastUserEvent(ByVal SlotEvent As Byte) As Integer

    ' Busca el último usuario que está en el torneo. En todos los eventos será el ganador.
    
    Dim LoopC As Integer
    
    With Events(SlotEvent)
        For LoopC = LBound(.Users()) To UBound(.Users())
            If .Users(LoopC).Id > 0 Then
                SearchLastUserEvent = .Users(LoopC).Id
                Exit For
            End If
        Next LoopC
    End With
End Function

Private Function SearchSlotEvent(ByVal Modality As String) As Byte
          Dim LoopC As Integer
          
        SearchSlotEvent = 0
          
        For LoopC = 1 To MAX_EVENT_SIMULTANEO
            With Events(LoopC)
                If StrComp(UCase$(strModality(LoopC, .Modality)), UCase$(Modality)) = 0 Then
                    SearchSlotEvent = LoopC
                    Exit For
                End If
            End With
        Next LoopC

End Function
Private Sub ResetEvent(ByVal Slot As Byte)
    Dim LoopC As Integer
    Dim UserIndex As Integer

    With Events(Slot)
        For LoopC = LBound(.Users()) To UBound(.Users())
            If .Users(LoopC).Id > 0 Then
                AbandonateEvent .Users(LoopC).Id, False

            End If
        Next LoopC
        
        If .NpcIndex > 0 Then Call QuitarNPC(.NpcIndex)
        
        .Enabled = False
        .Quotas = 0
        .Inscribed = 0
        .DspInscription = 0
        .GldInscription = 0
        .CanjeInscription = 0
        .LvlMax = 0
        .LvlMin = 0
        .TimeCancel = 0
        .NpcIndex = 0
        .TimeInit = 0
        .TimeCount = 0
        .CharBody = 0
        .CharHp = 0
        .Modality = 0
        .Run = False
        
        For LoopC = LBound(.AllowedClasses()) To UBound(.AllowedClasses())
            .AllowedClasses(LoopC) = 0
        Next LoopC
        
    End With
End Sub

Private Function CheckUserEvent(ByVal UserIndex As Integer, ByVal SlotEvent As Byte, ByRef ErrorMsg As String) As Boolean
    CheckUserEvent = False
        
    With UserList(UserIndex)
        If .flags.Muerto Then
            ErrorMsg = "No puedes participar en eventos estando muerto."
            Exit Function
        End If
        
        If .flags.Oculto Then
            ErrorMsg = "No puedes participar en eventos estando oculto."
            Exit Function
        End If
        
        'If .flags.Angel Then
         '   ErrorMsg = "No puedes participar en eventos estando en modo ANGEL."
          '  Exit Function
       ' End If
        
        'If .flags.Demonio Then
         '   ErrorMsg = "No puedes participar en eventos estando en modo DEMON."
          '  Exit Function
        'End If
        
        If .flags.Navegando Then
            ErrorMsg = "No puedes participar en eventos navegando."
            Exit Function
        End If
        
        'If .flags.EnConsulta Then
         '   ErrorMsg = "No puedes participar si estas en consulta."
          '  Exit Function
        'End If
        
        If .flags.Mimetizado Then
            ErrorMsg = "No puedes entrar mimetizado."
            Exit Function
        End If
        
        'If .flags.Montando Then
         '   ErrorMsg = "No puedes entrar montando."
          '  Exit Function
        'End If
        
        If .flags.invisible Then
            ErrorMsg = "No puedes entrar invisible."
            Exit Function
        End If
        
        If .flags.SlotEvent > 0 Then
            ErrorMsg = "Ya te encuentras en un evento. Tipea /SALIREVENTO para salir del mismo."
            Exit Function
        End If
        
        If .Counters.Pena > 0 Then
            ErrorMsg = "No puedes participar de los eventos en la cárcel. Maldito prisionero!"
            Exit Function
        End If
        
        If Zonas(UserList(UserIndex).zona).Segura = 0 Then
            ErrorMsg = "No puedes participar de los eventos estando en zona insegura. Vé a la ciudad mas cercana"
            Exit Function
        End If
        
        If .flags.Comerciando Then
            ErrorMsg = "No puedes participar de los eventos si estás comerciando."
            Exit Function
        End If
        
        If Not Events(SlotEvent).Enabled Or Events(SlotEvent).TimeInit > 0 Then
            ErrorMsg = "No hay ningun torneo disponible con ese nombre o bien las inscripciones no están disponibles aún."
            Exit Function
        End If
        
        If Events(SlotEvent).Run Then
            ErrorMsg = "El torneo ya ha comenzado. Mejor suerte para la próxima."
            Exit Function
        End If
        
        
        If Events(SlotEvent).LvlMin <> 0 Then
            If Events(SlotEvent).LvlMin > .Stats.ELV Then
                ErrorMsg = "Tu nivel no te permite ingresar a este evento."
                Exit Function
            End If
        End If
        
        If Events(SlotEvent).LvlMin <> 0 Then
            If Events(SlotEvent).LvlMax < .Stats.ELV Then
                ErrorMsg = "Tu nivel no te permite ingresar al evento."
                Exit Function
            End If
        End If
        
        If Events(SlotEvent).AllowedClasses(.clase) = 0 Then
            ErrorMsg = "Tu clase no está permitida en el evento."
            Exit Function
        End If
        
       ' If Events(SlotEvent).CanjeInscription > .Stats.puntos Then
           ' ErrorMsg = "No tienes suficiente canjes para pagar el torneo."
           'Exit Function
        'End If
        
        If Events(SlotEvent).GldInscription > .Stats.GLD Then
            ErrorMsg = "No tienes suficiente oro para pagar el torneo. Pide prestado a un compañero."
            Exit Function
        End If
        
        If Events(SlotEvent).DspInscription > 0 Then
            If Not TieneObjetos(880, Events(SlotEvent).DspInscription, UserIndex) Then
                ErrorMsg = "No tienes suficientes monedas DSP para participar del evento."
                Exit Function
            End If
        End If
        
        If Events(SlotEvent).Inscribed = Events(SlotEvent).Quotas Then
            ErrorMsg = "Los cupos del evento al que deseas participar ya fueron alcanzados."
            Exit Function
        End If
        
    
    End With
    CheckUserEvent = True
End Function

' EDICIÓN GENERAL
Public Function strModality(ByVal SlotEvent As Byte, ByVal Modality As eModalityEvent) As String

          ' Modalidad de cada evento
          
10        Select Case Modality
              Case eModalityEvent.CastleMode
20                strModality = "REYvsREY"
                  
30            Case eModalityEvent.DagaRusa
40                strModality = "DAGARUSA"
                  
50            Case eModalityEvent.DeathMatch
60                strModality = "DEATHMATCH"

190           Case eModalityEvent.Enfrentamientos
200               strModality = Events(SlotEvent).TeamCant & "vs" & Events(SlotEvent).TeamCant
210       End Select
End Function
Private Function strDescEvent(ByVal Modality As eModalityEvent) As String

    ' Descripción del evento en curso.
    Select Case Modality
        Case eModalityEvent.CastleMode
            strDescEvent = "> Los usuarios entrarán de forma aleatorea para formar dos equipos. Ambos equipos deberán defender a su rey y a su vez atacar al del equipo contrario."
        Case eModalityEvent.DagaRusa
            strDescEvent = "> Los usuarios se teletransportarán a una posición donde estará un asesino dispuesto a apuñalarlos y acabar con su vida. El último que quede en pie es el ganador del evento."
        Case eModalityEvent.DeathMatch
            strDescEvent = "> Los usuarios ingresan y luchan en una arena donde se toparan con todos los demás concursantes. El que logre quedar en pie, será el ganador."
        Case eModalityEvent.Enfrentamientos
            strDescEvent = "> Los usuarios se enfretan en una sala de Duelos para mostrar sus habilidades."
    End Select
End Function
Private Sub InitEvent(ByVal SlotEvent As Byte)
    
    Select Case Events(SlotEvent).Modality
        Case eModalityEvent.CastleMode
            Call InitCastleMode(SlotEvent)
            
        Case eModalityEvent.DagaRusa
            Call InitDagaRusa(SlotEvent)
            
        Case eModalityEvent.DeathMatch
            Call InitDeathMatch(SlotEvent)
            
        Case eModalityEvent.Enfrentamientos
            Call InitFights(SlotEvent)
        Case Else
            Exit Sub
        
    End Select
End Sub
Public Function CanAttackUserEvent(ByVal UserIndex As Integer, ByVal Victima As Integer) As Boolean
    
    ' Si el personaje es del mismo team, no se puede atacar al usuario.
    Dim VictimaSlotUserEvent As Byte
    
    VictimaSlotUserEvent = UserList(Victima).flags.SlotUserEvent
    
    With UserList(UserIndex)
        If .flags.SlotEvent > 0 Then
            If Events(.flags.SlotEvent).Users(VictimaSlotUserEvent).Team = Events(.flags.SlotEvent).Users(.flags.SlotUserEvent).Team Then
                CanAttackUserEvent = False
                Exit Function
            End If
        End If
        CanAttackUserEvent = True
    End With
End Function

Private Function ChangeBodyEvent(ByVal SlotEvent As Byte, ByVal UserIndex As Integer, ByVal ChangeHead As Boolean)
    
    ' En caso de que el evento cambie el body, de lo cambiamos.
    With UserList(UserIndex)
        .CharMimetizado.Body = .Char.Body
        .CharMimetizado.Head = .Char.Head
        .CharMimetizado.CascoAnim = .Char.CascoAnim
        .CharMimetizado.ShieldAnim = .Char.ShieldAnim
        .CharMimetizado.WeaponAnim = .Char.WeaponAnim

        .Char.Body = Events(SlotEvent).CharBody
        .Char.Head = IIf(ChangeHead = False, .Char.Head, 0)
        .Char.CascoAnim = 0
        .Char.ShieldAnim = 0
        .Char.WeaponAnim = 0
                
        ChangeUserChar UserIndex, .Char.Body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, True
        RefreshCharStatus UserIndex
    
    End With
End Function

Private Function ResetBodyEvent(ByVal SlotEvent As Byte, ByVal UserIndex As Integer)

    ' En caso de que el evento cambie el body del personaje, se lo restauramos.
    
    With UserList(UserIndex)
        If .flags.Muerto Then Exit Function
        'If Events(SlotEvent).Users(.flags.SlotUserEvent).Selected = 0 Then Exit Function
        
        If .CharMimetizado.Body > 0 Then
            .Char.Body = .CharMimetizado.Body
            .Char.Head = .CharMimetizado.Head
            .Char.CascoAnim = .CharMimetizado.CascoAnim
            .Char.ShieldAnim = .CharMimetizado.ShieldAnim
            .Char.WeaponAnim = .CharMimetizado.WeaponAnim
            
            
            .CharMimetizado.Body = 0
            .CharMimetizado.Head = 0
            .CharMimetizado.CascoAnim = 0
            .CharMimetizado.ShieldAnim = 0
            .CharMimetizado.WeaponAnim = 0
            
            .showName = True
            
            'ChangeUserChar UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, True
            ''RefreshCharStatus UserIndex
        End If
    
    End With
End Function

Private Sub ChangeHpEvent(ByVal UserIndex As Integer)

    ' En caso de que el evento edite la vida del personaje, se la editamos.
    
    Dim SlotEvent As Byte
    
    With UserList(UserIndex)
        SlotEvent = .flags.SlotEvent
        
        .Stats.OldHp = .Stats.MaxHP
        
        .Stats.MaxHP = Events(SlotEvent).CharHp
        .Stats.MinHP = .Stats.MaxHP
        
        WriteUpdateUserStats UserIndex
    
    End With
End Sub

Private Sub ResetHpEvent(ByVal UserIndex As Integer)

    ' En caso de que el evento haya editado la vida de un personaje, se la volvemos a restaurar.
    
    With UserList(UserIndex)
        
        .Stats.MaxHP = .Stats.OldHp
        .Stats.MinHP = .Stats.MaxHP
        .Stats.OldHp = 0
        WriteUpdateUserStats UserIndex
        
    End With
End Sub

'// Fin Funciones generales '//

Public Sub NewEvent(ByVal UserIndex As Integer, _
                    ByVal Modality As eModalityEvent, _
                    ByVal Quotas As Byte, _
                    ByVal LvlMin As Byte, _
                    ByVal LvlMax As Byte, _
                    ByVal GldInscription As Long, _
                    ByVal DspInscription As Long, _
                    ByVal CanjeInscription As Long, _
                    ByVal TimeInit As Long, _
                    ByVal TimeCancel As Long, _
                    ByVal TeamCant As Byte, _
                    ByRef AllowedClasses() As Byte)
                    
    Dim Slot As Integer
    Dim strTemp As String

    Slot = FreeSlotEvent()
    
    If Slot = 0 Then
        WriteConsoleMsg UserIndex, "No hay más lugar disponible para crear un evento simultaneo. Espera a que termine alguno o bien cancela alguno.", FontTypeNames.FONTTYPE_INFO
        Exit Sub
    Else
        With Events(Slot)
            .Enabled = True
            .Modality = Modality
            .Quotas = Quotas
            .LvlMin = LvlMin
            .LvlMax = LvlMax
            .GldInscription = GldInscription
            .DspInscription = DspInscription
            .CanjeInscription = CanjeInscription
            .AllowedClasses = AllowedClasses
            .TimeInit = TimeInit
            .TimeCancel = TimeCancel
            .TeamCant = TeamCant
            
            ReDim .Users(1 To .Quotas) As tUserEvent
            
            ' strModality devuelve: "Evento '1vs1' : Descripción"
            strTemp = strModality(Slot, .Modality) & strDescEvent(.Modality) & vbCrLf
            strTemp = strTemp & "Cupos máximos: " & .Quotas & vbCrLf

            strTemp = strTemp & IIf((.LvlMin > 0), "Nivel mínimo: " & .LvlMin & vbCrLf, vbNullString)
            strTemp = strTemp & IIf((.LvlMax > 0), "Nivel máximo: " & .LvlMax & vbCrLf, vbNullString)
            
            If .GldInscription > 0 And .DspInscription > 0 And .CanjeInscription > 0 Then
                strTemp = strTemp & "Inscripción requerida: " & .GldInscription & " monedas de oro, " & .DspInscription & " monedas DSP y " & .CanjeInscription & " Canjes."
            ElseIf .GldInscription > 0 Then
                strTemp = strTemp & "Inscripción requerida: " & .GldInscription & " monedas de oro."
            ElseIf .DspInscription > 0 Then
                strTemp = strTemp & "Inscripción requerida: " & .DspInscription & " monedas DSP."
            ElseIf .CanjeInscription > 0 Then
                strTemp = strTemp & "Inscripción requerida: " & .CanjeInscription & " Canjes."
            Else
                strTemp = strTemp & "Inscripción GRATIS"
            End If
            
            strTemp = strTemp & vbCrLf
            
            strTemp = strTemp & "Clases permitidas: " & CheckAllowedClasses(AllowedClasses) & ". Comando para ingresar /PARTICIPAR " & strModality(Slot, .Modality) & vbCrLf
            strTemp = strTemp & "Las inscripciones abren en " & Int(.TimeInit / 60) & " minutos."
        End With
        
        SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strTemp, FontTypeNames.FONTTYPE_INFOBOLD)
    End If
    
End Sub

Public Sub CloseEvent(ByVal Slot As Byte, Optional ByVal MsgConsole As String = vbNullString)
    
    With Events(Slot)
        If MsgConsole <> vbNullString Then SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(MsgConsole, FontTypeNames.FONTTYPE_DIOS)
        

        
        Call ResetEvent(Slot)
    End With
End Sub




Public Sub ParticipeEvent(ByVal UserIndex As Integer, ByVal Modality As String)
    
    Dim ErrorMsg As String
    Dim SlotUser As Byte
    Dim Pos As WorldPos
    Dim SlotEvent As Integer
    
    SlotEvent = SearchSlotEvent(Modality)
    
    If SlotEvent = 0 Then
        'SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Error Fatal TESTEO", FontTypeNames.FONTTYPE_ADMIN)
        Exit Sub
    End If
    
    With UserList(UserIndex)
        If CheckUserEvent(UserIndex, SlotEvent, ErrorMsg) Then
            SlotUser = FreeSlotUser(SlotEvent)
            
            .flags.SlotEvent = SlotEvent
            .flags.SlotUserEvent = SlotUser
            
            .PosAnt.map = .Pos.map
            .PosAnt.X = .Pos.X
            .PosAnt.Y = .Pos.Y
            
            .Stats.GLD = .Stats.GLD - Events(SlotEvent).GldInscription
            Call WriteUpdateGold(UserIndex)
            
            '.Stats.puntos = .Stats.puntos - Events(SlotEvent).CanjeInscription
            'Call WriteUpdateUserStats(UserIndex)
            
            'Call QuitarObjetos(880, Events(SlotEvent).DspInscription, UserIndex)
            
            With Events(SlotEvent)
                Pos.map = 25
                Pos.X = 39
                Pos.Y = 29
                
                Call FindLegalPos(UserIndex, Pos.map, Pos.X, Pos.Y)
                Call WarpUserChar(UserIndex, Pos.map, Pos.X, Pos.Y, False)
            
                .Users(SlotUser).Id = UserIndex
                .Inscribed = .Inscribed + 1
                
                
                WriteConsoleMsg UserIndex, "Has ingresado al evento " & strModality(SlotEvent, .Modality) & ". Espera a que se completen los cupos para que comience.", FontTypeNames.FONTTYPE_INFO
                
                If .Inscribed = .Quotas Then
                    SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strModality(SlotEvent, .Modality) & "Los cupos han sido alcanzados. Les deseamos mucha suerte a cada uno de los participantes y que gane el mejor!", FontTypeNames.FONTTYPE_GUILD)
                    .Run = True
                    InitEvent SlotEvent
                    Exit Sub
                End If
            End With
        
        Else
            WriteConsoleMsg UserIndex, ErrorMsg, FontTypeNames.FONTTYPE_WARNING
        
        End If
    End With
End Sub



Public Sub AbandonateEvent(ByVal UserIndex As Integer, _
                            Optional ByVal MsgAbandonate As Boolean = False, _
                            Optional ByVal Forzado As Boolean = False)
          
10    On Error GoTo error

          Dim Pos As WorldPos
          Dim SlotEvent As Byte
          Dim SlotUserEvent As Byte
          Dim UserTeam As Byte
          Dim UserMapFight As Byte
          
20        With UserList(UserIndex)
30            SlotEvent = .flags.SlotEvent
40            SlotUserEvent = .flags.SlotUserEvent
              
50            If SlotEvent > 0 And SlotUserEvent > 0 Then
60                With Events(SlotEvent)
                      'LogEventos "El personaje " & UserList(UserIndex).Name & " abandonó el evento de modalidad " & strModality(SlotEvent, .Modality)
70
                        If .Inscribed > 0 Then .Inscribed = .Inscribed - 1

80                        UserTeam = .Users(SlotUserEvent).Team
90                        UserMapFight = .Users(SlotUserEvent).MapFight
                          
100                       .Users(SlotUserEvent).Id = 0
110                       .Users(SlotUserEvent).Team = 0
                          .Users(SlotUserEvent).Value = 0
130                       .Users(SlotUserEvent).Selected = 0
140                       .Users(SlotUserEvent).MapFight = 0
                          
150                       UserList(UserIndex).flags.SlotEvent = 0
160                       UserList(UserIndex).flags.SlotUserEvent = 0
                          
180                       Select Case .Modality
                                  
                           Case eModalityEvent.DagaRusa
                               If Forzado And .Run Then
                                   Call WriteUserInEvent(UserIndex)
                                      
                                   If .Users(SlotUserEvent).Value = 0 Then
                                       Npclist(.NpcIndex).flags.InscribedPrevio = Npclist(.NpcIndex).flags.InscribedPrevio - 1
                                   End If
                                 Call WriteParalizeOK(UserIndex)
                               End If
                                  
310                           Case eModalityEvent.Enfrentamientos
320                               If Forzado Then
330                                   If UserMapFight > 0 Then
340                                       If Not Fight_CheckContinue(UserIndex, SlotEvent, UserTeam) Then
350                                           Fight_WinForzado UserIndex, SlotEvent, UserMapFight
360                                       End If
370                                   End If
380                               End If
                                  
390                               If UserList(UserIndex).Counters.TimeFight > 0 Then
400                                   UserList(UserIndex).Counters.TimeFight = 0
410                                   Call WriteUserInEvent(UserIndex)
420                               End If
                                  
430                       End Select
                                  
440                       Pos.map = UserList(UserIndex).PosAnt.map
450                       Pos.X = UserList(UserIndex).PosAnt.X
460                       Pos.Y = UserList(UserIndex).PosAnt.Y
                          
470                       Call FindLegalPos(UserIndex, Pos.map, Pos.X, Pos.Y)
480                       Call WarpUserChar(UserIndex, Pos.map, Pos.X, Pos.Y, False)
                          
490                       If Events(SlotEvent).CharBody <> 0 Then
500                           Call ResetBodyEvent(SlotEvent, UserIndex)
510                       End If
                  
520                       If UserList(UserIndex).Stats.OldHp <> 0 Then
530                           ResetHpEvent UserIndex
540                       End If
                  
550                       UserList(UserIndex).showName = True
560                       RefreshCharStatus UserIndex
                          
                          If MsgAbandonate Then WriteConsoleMsg UserIndex, "Has abandonado el evento. Podrás recibir una pena por hacer esto.", FontTypeNames.FONTTYPE_WARNING
  
                          ' Abandono general del evento
580                       If .Inscribed = 1 And Forzado Then
590                           Call FinishEvent(SlotEvent)
                          
600                           CloseEvent SlotEvent
610                           Exit Sub
620                       End If
                          
                          
630               End With
640           End If
              
              
650       End With
          
660   Exit Sub

error:
670      ' LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : AbandonateEvent()"
End Sub

Private Sub FinishEvent(ByVal SlotEvent As Byte)
    
    Dim UserIndex As Integer
    Dim IsSelected As Boolean
    
    With Events(SlotEvent)
        Select Case .Modality
            Case eModalityEvent.CastleMode
                UserIndex = SearchLastUserEvent(SlotEvent)
                CastleMode_Premio UserIndex, False
                
            Case eModalityEvent.DagaRusa
                UserIndex = SearchLastUserEvent(SlotEvent)
                Call WriteParalizeOK(UserIndex)
                DagaRusa_Premio UserIndex
                
            Case eModalityEvent.DeathMatch
                UserIndex = SearchLastUserEvent(SlotEvent)
                DeathMatch_Premio UserIndex
        End Select
    End With
    
    
    
End Sub


'#################EVENTO CASTLE MODE##########################
Public Function CanAttackReyCastle(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Boolean
    With UserList(UserIndex)
        If .flags.SlotEvent > 0 Then
            If Npclist(NpcIndex).flags.TeamEvent = Events(.flags.SlotEvent).Users(.flags.SlotUserEvent).Team Then
                CanAttackReyCastle = False
                Exit Function
            End If
        End If
    
    
        CanAttackReyCastle = True
    End With
End Function
Private Sub CastleMode_InitRey()
    Dim NpcIndex As Integer
    Const NumRey As Integer = 931
    Dim Pos As WorldPos
    Dim loopX As Integer, loopY As Integer
    Const Rango As Byte = 5
    
    For loopX = 20 - Rango To 20 + Rango
        For loopY = 82 - Rango To 82 + Rango
            If MapData(23).Tile(loopX, loopY).NpcIndex > 0 Then
                Call QuitarNPC(MapData(23).Tile(loopX, loopY).NpcIndex)
            End If
        Next loopY
    Next loopX
    
    Pos.map = 23
        
    Pos.X = 20
    Pos.Y = 82
    NpcIndex = SpawnNpc(NumRey, Pos, False, False, 132)
    Npclist(NpcIndex).flags.TeamEvent = 1
        
        
    For loopX = 76 - Rango To 53 + Rango
        For loopY = 53 - Rango To 53 + Rango
            If MapData(23).Tile(loopX, loopY).NpcIndex > 0 Then
                Call QuitarNPC(MapData(23).Tile(loopX, loopY).NpcIndex)
            End If
        Next loopY
    Next loopX
    Pos.map = 23
    Pos.X = 76
    Pos.Y = 53
    NpcIndex = SpawnNpc(NumRey, Pos, False, False, 132)
    Npclist(NpcIndex).flags.TeamEvent = 2
    
End Sub

Public Sub InitCastleMode(ByVal SlotEvent As Byte)
    Dim LoopC As Integer
    
    Const NumRey As Integer = 931
    Dim NpcIndex As Integer
    Dim Pos As WorldPos
    
    ' Spawn the npc castle mode
    CastleMode_InitRey
    
    With Events(SlotEvent)
        For LoopC = LBound(.Users()) To UBound(.Users())
            If .Users(LoopC).Id > 0 Then
                If LoopC > (UBound(.Users()) / 2) Then
                    .Users(LoopC).Team = 2
                    Pos.map = 23
                    Pos.X = 76
                    Pos.Y = 53
                    
                    Call FindLegalPos(.Users(LoopC).Id, Pos.map, Pos.X, Pos.Y)
                    Call WarpUserChar(.Users(LoopC).Id, Pos.map, Pos.X, Pos.Y, False)
                Else
                    .Users(LoopC).Team = 1
                    Pos.map = 23
                    Pos.X = 20
                    Pos.Y = 82
                    
                    Call FindLegalPos(.Users(LoopC).Id, Pos.map, Pos.X, Pos.Y)
                    Call WarpUserChar(.Users(LoopC).Id, Pos.map, Pos.X, Pos.Y, False)
                    
                End If
            End If
        Next LoopC
    End With
    
End Sub
Public Sub CastleMode_UserRevive(ByVal UserIndex As Integer)

    Dim LoopC As Integer
    Dim Pos As WorldPos
    
    With UserList(UserIndex)
        If .flags.SlotEvent > 0 Then
            Call RevivirUsuario(UserIndex)
            
            
            Pos.map = 23
            Pos.X = RandomNumber(20, 80)
            Pos.Y = RandomNumber(20, 80)
            
            Call ClosestLegalPos(Pos, Pos)
            StatsEvent UserIndex
            'benjakpoCall FindLegalPos(Userindex, Pos.Map, Pos.X, Pos.Y)
            Call WarpUserChar(UserIndex, Pos.map, Pos.X, Pos.Y, True)
        
        End If
    End With
End Sub

Public Sub FinishCastleMode(ByVal SlotEvent As Byte, ByVal UserEventSlot As Integer)
    Dim LoopC As Integer
    Dim strTemp As String
    Dim NpcIndex As Integer
    Dim MiObj As Obj
    
    With Events(SlotEvent)
        For LoopC = LBound(.Users()) To UBound(.Users())
            If .Users(LoopC).Id > 0 Then
                If .Users(LoopC).Team = .Users(UserEventSlot).Team Then
                    If LoopC = UserEventSlot Then
                        CastleMode_Premio .Users(LoopC).Id, True
                    Else
                        CastleMode_Premio .Users(LoopC).Id, False
                    End If
                    
                    If strTemp = vbNullString Then
                        strTemp = UserList(.Users(LoopC).Id).Name
                    Else
                        strTemp = strTemp & ", " & UserList(.Users(LoopC).Id).Name
                    End If
                End If
            End If
        Next LoopC
        
        
        CloseEvent SlotEvent, "CastleMode» Ha finalizado. Ha ganado el equipo de " & UCase$(strTemp)
    End With
    
End Sub

Private Sub CastleMode_Premio(ByVal UserIndex As Integer, ByVal KillRey As Boolean)

    ' Entregamos el premio del CastleMode
    Dim MiObj As Obj
    
    With UserList(UserIndex)
        '.Stats.puntos = .Stats.puntos + 10
       ' WriteConsoleMsg UserIndex, "Felicitaciones, has recibido 5 Canjes por haber ganado el evento!", FontTypeNames.FONTTYPE_INFO
        
        If KillRey Then
            WriteConsoleMsg UserIndex, "Hemos notado que has aniquilado con la vida del rey oponente. ¡FELICITACIONES! Aquí tienes tu recompensa! 5 Canjes EXTRAS!!!", FontTypeNames.FONTTYPE_INFO
            '.Stats.puntos = .Stats.puntos + 5
        End If
        
        MiObj.ObjIndex = 899
        MiObj.Amount = 1
                        
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(.Pos, MiObj)
        End If
                        
        MiObj.ObjIndex = 900
        MiObj.Amount = 1
                        
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(.Pos, MiObj)
        End If
        
        '.Stats.TorneosGanados = .Stats.TorneosGanados + 1
        
        WriteUpdateUserStats UserIndex
    End With
End Sub
' FIN EVENTO CASTLE MODE #####################################

' ###################### EVENTO DAGA RUSA ###########################
Public Sub InitDagaRusa(ByVal SlotEvent As Byte)

    Dim LoopC As Integer
    Dim NpcIndex As Integer
    Dim Pos As WorldPos
    
    Dim Num As Integer
    
    With Events(SlotEvent)
        For LoopC = LBound(.Users()) To UBound(.Users())
            If .Users(LoopC).Id > 0 Then
                Call WarpUserChar(.Users(LoopC).Id, 25, 17 + Num, 66, False)
                Num = Num + 1
                Call WriteUserInEvent(.Users(LoopC).Id)
                UserList(.Users(LoopC).Id).Counters.Congelado = 1
                Call WriteParalizeOK(.Users(LoopC).Id)
            End If
        Next LoopC
        
        Pos.map = 25
        Pos.X = 17
        Pos.Y = 65
        NpcIndex = SpawnNpc(932, Pos, False, False, 132)
    
        If NpcIndex <> 0 Then
            Npclist(NpcIndex).Movement = NpcDagaRusa
            Npclist(NpcIndex).flags.SlotEvent = SlotEvent
            Npclist(NpcIndex).flags.InscribedPrevio = .Inscribed
            .NpcIndex = NpcIndex
            
            DagaRusa_MoveNpc NpcIndex, True
        End If
        
        
        .TimeCount = 10
    End With


End Sub
Public Function DagaRusa_NextUser(ByVal SlotEvent As Byte) As Byte
    Dim LoopC As Integer
    
    DagaRusa_NextUser = 0
    
    With Events(SlotEvent)
        For LoopC = LBound(.Users()) To UBound(.Users())
            If (.Users(LoopC).Id > 0) And (.Users(LoopC).Value = 0) Then
                DagaRusa_NextUser = .Users(LoopC).Id
                '.Users(LoopC).value = 1
                Exit For
            End If
        Next LoopC
    End With
        
End Function
Public Sub DagaRusa_ResetRonda(ByVal SlotEvent As Byte)

    Dim LoopC As Integer
    
    With Events(SlotEvent)
        For LoopC = LBound(.Users()) To UBound(.Users())
            
            .Users(LoopC).Value = 0
        Next LoopC
    
    End With
End Sub
Private Sub DagaRusa_CheckWin(ByVal SlotEvent As Byte)

    Dim UserIndex As Integer
    Dim MiObj As Obj
    
    With Events(SlotEvent)
        If .Inscribed = 1 Then
            UserIndex = SearchLastUserEvent(SlotEvent)
            DagaRusa_Premio UserIndex
            
            Call WriteParalizeOK(UserIndex)
            Call QuitarNPC(.NpcIndex)
            CloseEvent SlotEvent
            
        End If
    End With
End Sub

Private Sub DagaRusa_Premio(ByVal UserIndex As Integer)

    Dim MiObj As Obj
    
    With UserList(UserIndex)
         MiObj.Amount = 1
         MiObj.ObjIndex = 402
        
        SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Daga Rusa> El ganador es " & UserList(UserIndex).Name & ". Felicitaciones para el personaje, quien se ha ganado 5 Canjes y una MD! (Espada mata dragones)", FontTypeNames.FONTTYPE_GUILD)
        
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        End If
        
        '.Stats.puntos = .Stats.puntos + 5
        
        '.Stats.TorneosGanados = .Stats.TorneosGanados + 1
        
        WriteUpdateUserStats UserIndex
        
    End With
End Sub
Public Sub DagaRusa_AttackUser(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
    
    Dim N As Integer
    Dim Slot As Byte
    
    With UserList(UserIndex)
        
        N = 10
        
        If RandomNumber(1, 100) <= N Then
        
            ' Sound
            SendData SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y)
            ' Fx
            SendData SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FXSANGRE, 0)
            ' Cambio de Heading
            ChangeNPCChar NpcIndex, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, SOUTH
            'Apuñalada en el piso
            'SendData SendTarget.ToPCArea, UserIndex, PrepareMessageCreateDamage(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, 1000, DAMAGE_PUÑAL)
            
            WriteConsoleMsg UserIndex, "¡Has sido apuñalado!", FontTypeNames.FONTTYPE_FIGHT
            
            Slot = .flags.SlotEvent
            
            
            Call UserDie(UserIndex)
            EventosAOyin.AbandonateEvent (UserIndex)
            Call DagaRusa_CheckWin(Slot)
           
            
        Else
            ' Sound
            SendData SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y)
            ' Fx
            SendData SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FXSANGRE, 0)
            ' Cambio de Heading
            ChangeNPCChar NpcIndex, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, SOUTH

            WriteConsoleMsg UserIndex, "¡Parece que no te he apuñalado, ya verás!", FontTypeNames.FONTTYPE_FIGHT
           ' SendData SendTarget.ToPCArea, Userindex, PrepareMessageCreateDamage(UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y, 1000, DAMAGE_PUÑAL)
        End If
        
        
        
    End With
End Sub

' FIN EVENTO DAGA RUSA ###########################################
Private Function SelectModalityDeathMatch(ByVal SlotEvent As Byte) As Integer
    Dim Random As Integer
    
    Randomize
    Random = RandomNumber(1, 8)
    
    With Events(SlotEvent)
        Select Case Random
            Case 1 ' Zombie
                .CharBody = 11
            Case 2 ' Golem
                .CharBody = 11
            Case 3 ' Araña
                .CharBody = 42
            Case 4 ' Asesino
                .CharBody = 11 '48
            Case 5 'Medusa suprema
                .CharBody = 151
            Case 6 'Dragón azul
                .CharBody = 42 '247
            Case 7 'Viuda negra 185
                .CharBody = 185
            Case 8 'Tigre salvaje
                .CharBody = 147
        End Select
    End With
End Function

' DEATHMATCH ####################################################
Private Sub InitDeathMatch(ByVal SlotEvent As Byte)

    Dim LoopC As Integer
    Dim Pos As WorldPos
    
    'Call SelectModalityDeathMatch(SlotEvent)
    
    With Events(SlotEvent)
        For LoopC = LBound(.Users()) To UBound(.Users())
            If .Users(LoopC).Id > 0 Then
                .Users(LoopC).Team = LoopC
                .Users(LoopC).Selected = 1
                
                'ChangeBodyEvent SlotEvent, .Users(LoopC).Id, True
                'UserList(.Users(LoopC).Id).showName = False
                'RefreshCharStatus .Users(LoopC).Id
                
                
                Pos.map = 24
                Pos.X = RandomNumber(85, 128)
                Pos.Y = RandomNumber(82, 104)
            
                Call ClosestLegalPos(Pos, Pos)
                Call WarpUserChar(.Users(LoopC).Id, Pos.map, Pos.X, Pos.Y, True)
            End If
        
        Next LoopC
    
        .TimeCount = 20
    End With
    
End Sub

Public Sub DeathMatch_UserDie(ByVal SlotEvent As Byte, ByVal UserIndex As Integer)
    AbandonateEvent (UserIndex)
        
    If Events(SlotEvent).Inscribed = 1 Then
        UserIndex = SearchLastUserEvent(SlotEvent)
        DeathMatch_Premio UserIndex
        CloseEvent SlotEvent
    End If
End Sub
Private Sub DeathMatch_Premio(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("DeathMatch> El ganador es " & .Name & " quien se lleva 1 punto de torneo  y 10 Canjes!!!.", FontTypeNames.FONTTYPE_GUILD)
            
        '.Stats.puntos = .Stats.puntos + 10
        
        '.Stats.TorneosGanados = .Stats.TorneosGanados + 1
        
        WriteUpdateUserStats UserIndex
        
    End With
End Sub

' ENFRENTAMIENTOS ###############################################

Private Sub InitFights(ByVal SlotEvent As Byte)
10    On Error GoTo error
          
20        With Events(SlotEvent)
30            Fight_SelectedTeam SlotEvent
40            Fight_Combate SlotEvent
50        End With
60    Exit Sub

error:
70        'LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : InitFights()"
End Sub
Private Sub Fight_SelectedTeam(ByVal SlotEvent As Byte)
          
10    On Error GoTo error

          ' En los enfrentamientos utilizamos este procedimiento para seleccionar los grupos o bien el usuario queda solo por 1vs1.
          Dim loopX As Integer
          Dim loopY As Integer
          Dim Team As Byte
          
20        Team = 1
          
30        With Events(SlotEvent)
40            For loopX = LBound(.Users()) To UBound(.Users()) Step .TeamCant
50                For loopY = 0 To (.TeamCant - 1)
60                    .Users(loopX + loopY).Team = Team
70                Next loopY
                  
80                Team = Team + 1
90            Next loopX
          
100       End With
          
110   Exit Sub

error:
120       'LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : Fight_SelectedTeam()"
End Sub

Private Sub Fight_WarpTeam(ByVal SlotEvent As Byte, _
                                        ByVal ArenaSlot As Byte, _
                                        ByVal TeamEvent As Byte, _
                                        ByVal IsContrincante As Boolean, _
                                        ByRef StrTeam As String)

10    On Error GoTo error

          Dim LoopC As Integer
          Dim strTemp As String, strTemp1 As String, strTemp2 As String
          
20        With Events(SlotEvent)
30            For LoopC = LBound(.Users()) To UBound(.Users())
40                If .Users(LoopC).Id > 0 And .Users(LoopC).Team = TeamEvent Then
50                    If IsContrincante Then
60                        Call EventWarpUser(.Users(LoopC).Id, MapEvent.Fight(ArenaSlot).map, MapEvent.Fight(ArenaSlot).X + MAP_TILE_VS, MapEvent.Fight(ArenaSlot).Y + MAP_TILE_VS)
                          
90                    Else
100                       Call EventWarpUser(.Users(LoopC).Id, MapEvent.Fight(ArenaSlot).map, MapEvent.Fight(ArenaSlot).X, MapEvent.Fight(ArenaSlot).Y)

130                   End If
                      
140                   If StrTeam = vbNullString Then
150                       StrTeam = UserList(.Users(LoopC).Id).Name
160                   Else
170                       StrTeam = StrTeam & "-" & UserList(.Users(LoopC).Id).Name
180                   End If
                      
190                   .Users(LoopC).Value = 1
200                   .Users(LoopC).MapFight = ArenaSlot
                      
210                   UserList(.Users(LoopC).Id).Counters.TimeFight = 10
220                   Call WriteUserInEvent(.Users(LoopC).Id)
230               End If
240           Next LoopC
250       End With
          
260   Exit Sub

error:
270       'LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : Fight_WarpTeam()"
End Sub

Private Function Fight_Search_Enfrentamiento(ByVal UserIndex As Integer, ByVal UserTeam As Byte, ByVal SlotEvent As Byte) As Byte
10    On Error GoTo error

          ' Chequeamos que tengamos contrincante para luchar.
          Dim LoopC As Integer
          
20        Fight_Search_Enfrentamiento = 0
          
30        With Events(SlotEvent)
40            For LoopC = LBound(.Users()) To UBound(.Users())
50                If .Users(LoopC).Id > 0 And .Users(LoopC).Value = 0 Then
60                    If .Users(LoopC).Id <> UserIndex And .Users(LoopC).Team <> UserTeam Then
70                        Fight_Search_Enfrentamiento = .Users(LoopC).Team
80                        Exit For
90                    End If
100               End If
110           Next LoopC
          
120       End With
          
130   Exit Function

error:
140       'LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : Fight_Search_Enfrentamiento()"
End Function

Private Sub NewRound(ByVal SlotEvent As Byte)
          Dim LoopC As Long
          Dim Count As Long
          
10        With Events(SlotEvent)
20            Count = 0
              
30            For LoopC = LBound(.Users()) To UBound(.Users())
40                If .Users(LoopC).Id > 0 Then
                      ' Hay esperando
50                    If .Users(LoopC).Value = 0 Then
60                        Exit Sub
70                    End If
                      
                      ' Hay luchando
80                    If .Users(LoopC).MapFight > 0 Then
90                        Exit Sub
100                   End If
110               End If
120           Next LoopC
              
130           For LoopC = LBound(.Users()) To UBound(.Users())
140               .Users(LoopC).Value = 0
150           Next LoopC

            'LogEventos "Se reinicio la informacion de los fights()"
              
160       End With
End Sub
Private Function FreeSlotArena() As Byte
          Dim LoopC As Integer
          
10        FreeSlotArena = 0
          
20        For LoopC = 1 To MAX_MAP_FIGHT
30            If MapEvent.Fight(LoopC).Run = False Then
40                FreeSlotArena = LoopC
50                Exit For
60            End If
70        Next LoopC
End Function
Private Sub Fight_Combate(ByVal SlotEvent As Byte)
10    On Error GoTo error

          ' Buscamos una arena disponible y mandamos la mayor cantidad de usuarios disponibles.
          Dim LoopC As Integer
          Dim FreeArena As Byte
          Dim OponentTeam As Byte
          Dim strTemp As String
          Dim strTeam1 As String
          Dim strTeam2 As String
          
20        With Events(SlotEvent)
cheking:
30            For LoopC = LBound(.Users()) To UBound(.Users())
40                If .Users(LoopC).Id > 0 And .Users(LoopC).Value = 0 Then
50                    FreeArena = FreeSlotArena()
                      
60                    If FreeArena > 0 Then
70                        OponentTeam = Fight_Search_Enfrentamiento(.Users(LoopC).Id, .Users(LoopC).Team, SlotEvent)
                          
80                        If OponentTeam > 0 Then
90                            StatsEvent .Users(LoopC).Id
100                           Fight_WarpTeam SlotEvent, FreeArena, .Users(LoopC).Team, False, strTeam1
110                           Fight_WarpTeam SlotEvent, FreeArena, OponentTeam, True, strTeam2
120                           MapEvent.Fight(FreeArena).Run = True
                              
130                           strTemp = "Duelos " & Events(SlotEvent).TeamCant & "vs" & Events(SlotEvent).TeamCant & "» "
140                           strTemp = strTemp & strTeam1 & " vs " & strTeam2
150                           SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(strTemp, FontTypeNames.FONTTYPE_GUILD)
                              
160                           strTemp = vbNullString
170                           strTeam1 = vbNullString
180                           strTeam2 = vbNullString
                              
190                       Else
                              ' Pasa de ronda automaticamente
200                           .Users(LoopC).Value = 1
210                           WriteConsoleMsg .Users(LoopC).Id, "Hemos notado que no tienes un adversario. Pasaste a la siguiente ronda.", FontTypeNames.FONTTYPE_INFO
220                           NewRound SlotEvent
                              GoTo cheking:
230                       End If
240                   End If
250               End If
260           Next LoopC
              
270       End With
          
280   Exit Sub

error:
290       'LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : Fight_Combate()"
End Sub
Private Sub ResetValue(ByVal SlotEvent As Byte)
          Dim LoopC As Integer
          
10        With Events(SlotEvent)
20            For LoopC = LBound(.Users()) To UBound(.Users())
30                .Users(LoopC).Value = 0
40            Next LoopC
50        End With
End Sub
Private Function CheckTeam_UserDie(ByVal SlotEvent As Integer, ByVal TeamUser As Byte) As Boolean

10    On Error GoTo error

          Dim LoopC As Integer
          ' Encontramos a uno del Team vivo, significa que no hay terminación del duelo.
          
          
20        With Events(SlotEvent)
30            For LoopC = LBound(.Users()) To UBound(.Users())
40                If .Users(LoopC).Id > 0 Then
50                    If .Users(LoopC).Team = TeamUser Then
60                        If UserList(.Users(LoopC).Id).flags.Muerto = 0 Then
70                            CheckTeam_UserDie = False
80                            Exit Function
90                        End If
100                   End If
110               End If
120           Next LoopC
              
130           CheckTeam_UserDie = True
          
140       End With
          
150   Exit Function

error:
160       'LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : CheckTeam_UserDie()"
End Function
Private Sub Team_UserDie(ByVal SlotEvent As Byte, ByVal TeamSlot As Byte)
10    On Error GoTo error

          Dim LoopC As Integer
20        With Events(SlotEvent)
              
              
30            For LoopC = LBound(.Users()) To UBound(.Users())
40                If .Users(LoopC).Id > 0 Then
50                    If .Users(LoopC).Team = TeamSlot Then
60                        AbandonateEvent .Users(LoopC).Id
70                    End If
80                End If
90            Next LoopC
          
100       End With
          
110   Exit Sub

error:
120       'LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : Team_UserDie()"
End Sub
Public Function Fight_CheckContinue(ByVal UserIndex As Integer, ByVal SlotEvent As Byte, ByVal TeamSlot As Byte) As Boolean
          ' Esta función devuelve un TRUE cuando el enfrentamiento puede seguir.
          
          Dim LoopC As Integer, Cant As Integer
          
10        With Events(SlotEvent)
              
20            Fight_CheckContinue = False
              
30            For LoopC = LBound(.Users()) To UBound(.Users())
                  ' User válido
40                If .Users(LoopC).Id > 0 And .Users(LoopC).Id <> UserIndex Then
50                    If .Users(LoopC).Team = TeamSlot Then
60                        If UserList(.Users(LoopC).Id).flags.Muerto = 0 Then
70                            Fight_CheckContinue = True
80                            Exit For
90                        End If
100                   End If
110               End If
120           Next LoopC

130       End With
          
140   Exit Function

error:
150       'LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : Team_CheckContinue()"
End Function
Public Sub Fight_WinForzado(ByVal UserIndex As Integer, ByVal SlotEvent As Byte, ByVal MapFight As Byte)
10        On Error GoTo error
          
          Dim LoopC As Integer
          Dim strTempWin As String
          Dim TeamWin As Byte
          
20        With Events(SlotEvent)

              'LogEventos "El personaje " & UserList(UserIndex).Name & " deslogeó en lucha."
              
30            For LoopC = LBound(.Users()) To UBound(.Users())
40                With .Users(LoopC)
50                    If .Id > 0 And UserIndex <> .Id Then
60                        If .MapFight = MapFight Then
70                            If strTempWin = vbNullString Then
80                                strTempWin = UserList(.Id).Name
90                            Else
100                               strTempWin = strTempWin & "-" & UserList(.Id).Name
110                           End If
                              
                              '.value = 0
130                           .MapFight = 0
                              
140                           EventWarpUser .Id, 1, 174, 130
                              WriteConsoleMsg .Id, "Felicitaciones. Has ganado el enfrentamiento", FontTypeNames.FONTTYPE_INFO

180                           TeamWin = .Team
190                       End If
200                   End If
210               End With
220           Next LoopC

              MapEvent.Fight(MapFight).Run = False
              
              
230           If strTempWin <> vbNullString Then SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Duelos " & Events(SlotEvent).TeamCant & "vs" & Events(SlotEvent).TeamCant & "» Duelo ganado por " & strTempWin & ".", FontTypeNames.FONTTYPE_GUILD)
              
              ' Nos fijamos si resetea el Value
240           Call NewRound(SlotEvent)
              
              ' Nos fijamos si eran los últimos o si podemos mandar otro combate..
250           If TeamCant(SlotEvent, TeamWin) = .Inscribed Then
260               Fight_SearchTeamWin SlotEvent, TeamWin
270               CloseEvent SlotEvent
280           Else
290               Fight_Combate SlotEvent
300           End If
          
310       End With
          
320   Exit Sub

error:
330       'LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : Fight_WinForzado()"
End Sub
Private Sub StatsEvent(ByVal UserIndex As Integer)
10    On Error GoTo error

20        With UserList(UserIndex)
30            If .flags.Muerto Then
40                Call RevivirUsuario(UserIndex)
50                Exit Sub
60            End If
              
70            .Stats.MinHP = .Stats.MaxHP
80            .Stats.MinMAN = .Stats.MaxMAN
90            .Stats.MinAGU = 100
100           .Stats.MinHam = 100
                .Stats.MinSta = .Stats.MaxSta
              
110           WriteUpdateUserStats UserIndex
          
120       End With
          
130   Exit Sub

error:
140       'LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : StatsEvent()"
End Sub

Private Function SearchTeamAttacker(ByVal TeamUser As Byte)

End Function
Public Sub Fight_UserDie(ByVal SlotEvent As Byte, ByVal SlotUserEvent As Byte, ByVal AttackerIndex As Integer)
10    On Error GoTo error
    Dim TeamSlot As Byte
    Dim LoopC As Integer
    Dim strTempWin As String
    Dim TeamWin As Byte
    Dim MapFight As Byte
    
    ' Aca se hace que el que gané no siga luchando sino que espere.
    
20    With Events(SlotEvent)
30        TeamSlot = .Users(SlotUserEvent).Team
40        TeamWin = .Users(UserList(AttackerIndex).flags.SlotUserEvent).Team
        
50        If CheckTeam_UserDie(SlotEvent, TeamSlot) = False Then Exit Sub
        
60        For LoopC = LBound(.Users()) To UBound(.Users())
70            If .Users(LoopC).Id > 0 Then
80                    With .Users(LoopC)
90                        If .Team = TeamWin Then
100                           StatsEvent .Id
110
120                            If strTempWin = vbNullString Then
130                                strTempWin = UserList(.Id).Name
140                            Else
150                               strTempWin = strTempWin & "-" & UserList(.Id).Name
160                         End If
                            
                            
                            MapFight = .MapFight
170
                            
                               '.Value = 0
180                            .MapFight = 0
190                            EventWarpUser .Id, 1, 16, 24
                               WriteConsoleMsg .Id, "Felicitaciones. Has ganado el enfrentamiento", FontTypeNames.FONTTYPE_INFO
                           
                            ' / Update color char team
220                           RefreshCharStatus (.Id)
230                     End If
240                 End With
250             End If
260     Next LoopC
        
        MapEvent.Fight(MapFight).Run = False
        
        ' Abandono del user/team
270     Team_UserDie SlotEvent, TeamSlot
        
280     If strTempWin <> vbNullString Then SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Duelos " & Events(SlotEvent).TeamCant & "vs" & Events(SlotEvent).TeamCant & "» Enfrentamiento ganado por " & strTempWin & ".", FontTypeNames.FONTTYPE_GUILD)
        
        ' // Se fija de poder pasar a la siguiente ronda o esperar a los combates que faltan.
290     Call NewRound(SlotEvent)
        
        ' Si la cantidad es igual al inscripto quedó final.
300     If TeamCant(SlotEvent, TeamWin) = .Inscribed Then
310            Fight_SearchTeamWin SlotEvent, TeamWin
320            CloseEvent SlotEvent
330     Else
340            Fight_Combate SlotEvent
350     End If
        
360       End With
    
370   Exit Sub

error:
380       'LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : Fight_UserDie()" & " AT LINE: " & Erl
End Sub
Private Function TeamCant(ByVal SlotEvent As Byte, ByVal TeamSlot As Byte) As Byte

10    On Error GoTo error
          ' Devuelve la cantidad de miembros que tiene un clan
          Dim LoopC As Integer
          
20        TeamCant = 0
          
30        With Events(SlotEvent)
40            For LoopC = LBound(.Users()) To UBound(.Users())
50                If .Users(LoopC).Team = TeamSlot Then
60                    TeamCant = TeamCant + 1
70                End If
80            Next LoopC
90        End With
          
100   Exit Function

error:
110       'LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : TeamCant()"
End Function
Private Sub Fight_SearchTeamWin(ByVal SlotEvent As Byte, ByVal TeamWin As Byte)

10    On Error GoTo error

          Dim LoopC As Integer
          Dim strTemp As String
          Dim strReWard As String
          
          
20        With Events(SlotEvent)
30            For LoopC = LBound(.Users()) To UBound(.Users())
40                If .Users(LoopC).Id > 0 And .Users(LoopC).Team = TeamWin Then
                      WriteConsoleMsg .Users(LoopC).Id, "Has ganado el evento. ¡Felicitaciones!", FontTypeNames.FONTTYPE_INFO
                      'UserList(.Users(LoopC).Id).Stats.puntos = UserList(.Users(LoopC).Id).Stats.puntos + 10
60                    'Prizeuser .Users(LoopC).Id, False
                      
70                    If strTemp = vbNullString Then
80                        strTemp = UserList(.Users(LoopC).Id).Name
90                    Else
100                       strTemp = strTemp & ", " & UserList(.Users(LoopC).Id).Name
110                   End If
120               End If
130           Next LoopC
          
          
140       If .TeamCant > 1 Then
150           'If .GldInscription > 0 Or .DspInscription > 0 Then strReWard = "Los participantes han recibido 10 CANJES y "
160           If .GldInscription > 0 Then strReWard = strReWard & .GldInscription * .Quotas & " Monedas de oro. "
170           'If .DspInscription > 0 Then strReWard = strReWard & .DspInscription * .Quotas & " Monedas DSP. "
              'If .CanjeInscription > 0 Then strReWard = strReWard & .CanjeInscription * .Quotas & " Canjes."
              For LoopC = LBound(.Users()) To UBound(.Users())
              If .Users(LoopC).Id > 0 And .Users(LoopC).Team = TeamWin Then
               UserList(.Users(LoopC).Id).Stats.GLD = UserList(.Users(LoopC).Id).Stats.GLD + (.GldInscription * .Quotas)

              End If
              Next LoopC
180           SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Duelos " & .TeamCant & "vs" & .TeamCant & _
                  " Evento terminado. Felicitamos a " & strTemp & " por haber ganado el torneo." & vbCrLf & strReWard, FontTypeNames.FONTTYPE_GM)
190       Else
200           'If .GldInscription > 0 Or .DspInscription > 0 Then strReWard = "El participante recibió 10 CANJES y "
210           If .GldInscription > 0 Then strReWard = strReWard & .GldInscription * .Quotas & " Monedas de oro"
220           'If .DspInscription > 0 Then strReWard = strReWard & " y " & .DspInscription * .Quotas & " Monedas DSP."
              'If .CanjeInscription > 0 Then strReWard = strReWard & .CanjeInscription * .Quotas & " Canjes."
              For LoopC = LBound(.Users()) To UBound(.Users())
              If .Users(LoopC).Id > 0 And .Users(LoopC).Team = TeamWin Then
               UserList(.Users(LoopC).Id).Stats.GLD = UserList(.Users(LoopC).Id).Stats.GLD + (.GldInscription * .Quotas)

              End If
              Next LoopC
              SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Duelos " & .TeamCant & "vs" & .TeamCant & " Evento terminado. Felicitamos a " & strTemp & _
                  " por haber ganado el evento." & vbCrLf & strReWard, FontTypeNames.FONTTYPE_DIOS)
240       End If
            CloseEvent SlotEvent
          
250       End With
          
          'aqui
260   Exit Sub

error:
270       'LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : Fight_SearchTeamWin()"
End Sub


Private Sub EventWarpUser(ByVal UserIndex As Integer, ByVal map As Integer, ByVal X As Byte, ByVal Y As Byte)
10    On Error GoTo error

          ' Teletransportamos a cualquier usuario que cumpla con la regla de estar en un evento.
          
          Dim Pos As WorldPos
          
20        With UserList(UserIndex)
30            Pos.map = map
40            Pos.X = X
50            Pos.Y = Y
              
60            ClosestStablePos Pos, Pos
70            WarpUserChar UserIndex, Pos.map, Pos.X, Pos.Y, False
          
80        End With
          
90    Exit Sub

error:
100       'LogEventos "[" & Err.Number & "] " & Err.Description & ") PROCEDIMIENTO : EventWarpUser()"
End Sub





