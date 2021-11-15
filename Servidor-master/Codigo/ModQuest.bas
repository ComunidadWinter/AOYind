Attribute VB_Name = "ModQuest"
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
'along with this program; if not, you can find it at [url=http://www.affero.org/oagpl.html]http://www.affero.org/oagpl.html[/url]
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at [email=aaron@baronsoft.com]aaron@baronsoft.com[/email]
'for more information about ORE please visit [url=http://www.baronsoft.com/]http://www.baronsoft.com/[/url]
Option Explicit
 
'Constantes de las quests
Public Const MAXUSERQUESTS As Integer = 5    'Máxima cantidad de quests que puede tener un usuario al mismo tiempo.
 
Public Function TieneQuest(ByVal UserIndex As Integer, ByVal QuestNumber As Integer) As Byte
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Devuelve el slot de UserQuests en que tiene la quest QuestNumber. En caso contrario devuelve 0.
'Last modified: 27/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Dim i As Integer
 
    For i = 1 To MAXUSERQUESTS
        If UserList(UserIndex).QuestStats.Quests(i).QuestIndex = QuestNumber Then
            TieneQuest = i
            Exit Function
        End If
    Next i
    
    TieneQuest = 0
End Function
 
Public Function FreeQuestSlot(ByVal UserIndex As Integer) As Byte
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Devuelve el próximo slot de quest libre.
'Last modified: 27/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Dim i As Integer
 
    For i = 1 To MAXUSERQUESTS
        If UserList(UserIndex).QuestStats.Quests(i).QuestIndex = 0 Then
            FreeQuestSlot = i
            Exit Function
        End If
    Next i
    
    FreeQuestSlot = 0
End Function
 
Public Sub HandleQuestAccept(ByVal UserIndex As Integer)
        
        On Error GoTo HandleQuestAccept_Err
        

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Maneja el evento de aceptar una quest.
        'Last modified: 31/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        Dim NpcIndex  As Integer

        Dim QuestSlot As Byte
        
        Dim indice As Integer
 
100     Call UserList(UserIndex).incomingData.ReadByte

102     indice = UserList(UserIndex).incomingData.ReadInteger
 
104     NpcIndex = UserList(UserIndex).flags.TargetNPC
    
106     If NpcIndex = 0 Then Exit Sub
108     If indice = 0 Then Exit Sub
    
        'Esta el personaje en la distancia correcta?
110     If Distancia(UserList(UserIndex).Pos, Npclist(NpcIndex).Pos) > 5 Then
112         Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
114     If TieneQuest(UserIndex, Npclist(NpcIndex).QuestNumber(indice)) Then
116         Call WriteConsoleMsg(UserIndex, "La quest ya esta en curso.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        
        
        'El personaje completo la quest que requiere?
118     If QuestList(Npclist(NpcIndex).QuestNumber(indice)).RequiredQuest > 0 Then
120         If Not UserDoneQuest(UserIndex, QuestList(Npclist(NpcIndex).QuestNumber(indice)).RequiredQuest) Then
122             Call WriteChatOverHead(UserIndex, "Debes completas la quest " & QuestList(QuestList(Npclist(NpcIndex).QuestNumber(indice)).RequiredQuest).Nombre & " para emprender esta mision.", Npclist(NpcIndex).Char.CharIndex, vbYellow)
                Exit Sub
            End If
        End If
        

        'El personaje tiene suficiente nivel?
124     If UserList(UserIndex).Stats.ELV < QuestList(Npclist(NpcIndex).QuestNumber(indice)).RequiredLevel Then
126         Call WriteChatOverHead(UserIndex, "Debes ser por lo menos nivel " & QuestList(Npclist(NpcIndex).QuestNumber(indice)).RequiredLevel & " para emprender esta mision.", Npclist(NpcIndex).Char.CharIndex, vbYellow)
            Exit Sub
        End If
        
        
        'El personaje ya hizo la quest?
128     If UserDoneQuest(UserIndex, Npclist(NpcIndex).QuestNumber(indice)) Then
130         Call WriteChatOverHead(UserIndex, "QUESTNEXT*" & Npclist(NpcIndex).QuestNumber(indice), Npclist(NpcIndex).Char.CharIndex, vbYellow)
            Exit Sub
        End If
    
132     QuestSlot = FreeQuestSlot(UserIndex)


134     If QuestSlot = 0 Then
136         Call WriteChatOverHead(UserIndex, "Debes completar las misiones en curso para poder aceptar más misiones.", Npclist(NpcIndex).Char.CharIndex, vbYellow)
            Exit Sub
        End If
        
        
        



    
        'Agregamos la quest.
138     With UserList(UserIndex).QuestStats.Quests(QuestSlot)
140         .QuestIndex = Npclist(NpcIndex).QuestNumber(indice)
        
142         If QuestList(.QuestIndex).RequiredNPCs Then ReDim .NPCsKilled(1 To QuestList(.QuestIndex).RequiredNPCs)
144         If QuestList(.QuestIndex).RequiredTargetNPCs Then ReDim .NPCsTarget(1 To QuestList(.QuestIndex).RequiredTargetNPCs)
146         Call WriteConsoleMsg(UserIndex, "Has aceptado la mision " & Chr(34) & QuestList(.QuestIndex).Nombre & Chr(34) & ".", FontTypeNames.FONTTYPE_INFO)
148         Call WriteUpdateNPCSimbolo(UserIndex, NpcIndex, 6)
        
        End With

        
        Exit Sub

HandleQuestAccept_Err:
150    ' Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleQuestAccept", Erl)
152     Resume Next
        
End Sub
 
Public Sub FinishQuest(ByVal UserIndex As Integer, ByVal QuestIndex As Integer, ByVal QuestSlot As Byte)
        
        On Error GoTo FinishQuest_Err
        

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Maneja el evento de terminar una quest.
        'Last modified: 29/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        Dim i              As Integer

        Dim InvSlotsLibres As Byte

        Dim NpcIndex       As Integer
        

 
100     NpcIndex = UserList(UserIndex).flags.TargetNPC
    
102     With QuestList(QuestIndex)

            'Comprobamos que tenga los objetos.
104         If .RequiredOBJs > 0 Then

106             For i = 1 To .RequiredOBJs

108                 If TieneObjetos(.RequiredOBJ(i).ObjIndex, .RequiredOBJ(i).Amount, UserIndex) = False Then
110                     Call WriteChatOverHead(UserIndex, "No has conseguido todos los objetos que te he pedido.", Npclist(NpcIndex).Char.CharIndex, vbYellow)
                    
                        Exit Sub

                    End If

112             Next i

            End If
        
            'Comprobamos que haya matado todas las criaturas.
114         If .RequiredNPCs > 0 Then

116             For i = 1 To .RequiredNPCs

118                 If .RequiredNPC(i).Amount > UserList(UserIndex).QuestStats.Quests(QuestSlot).NPCsKilled(i) Then
120                     Call WriteChatOverHead(UserIndex, "No has matado todas las criaturas que te he pedido.", Npclist(NpcIndex).Char.CharIndex, vbYellow)
                        Exit Sub

                    End If

122             Next i

            End If
            
            'Comprobamos que haya targeteado todos los npc
124          If .RequiredTargetNPCs > 0 Then

126              For i = 1 To .RequiredTargetNPCs
    
128                  If .RequiredTargetNPC(i).Amount > UserList(UserIndex).QuestStats.Quests(QuestSlot).NPCsTarget(i) Then
130                      Call WriteChatOverHead(UserIndex, "No has visitado al npc que te pedi.", Npclist(NpcIndex).Char.CharIndex, vbYellow)
                        Exit Sub
    
                        End If
    
132              Next i

            End If
    
            'Comprobamos que el usuario tenga espacio para recibir los items.
134         If .RewardOBJs > 0 Then

                'Buscamos la cantidad de slots de inventario libres.
136             For i = 1 To UserList(UserIndex).CurrentInventorySlots

138                 If UserList(UserIndex).Invent.Object(i).ObjIndex = 0 Then InvSlotsLibres = InvSlotsLibres + 1
140             Next i
            
                'Nos fijamos si entra
142             If InvSlotsLibres < .RewardOBJs Then
144                 Call WriteChatOverHead(UserIndex, "No tienes suficiente espacio en el inventario para recibir la recompensa. Vuelve cuando hayas hecho mas espacio.", Npclist(NpcIndex).Char.CharIndex, vbYellow)
                    Exit Sub

                End If

            End If
    
            'A esta altura ya cumplio los objetivos, entonces se le entregan las recompensas.
146         Call WriteChatOverHead(UserIndex, "QUESTFIN*" & QuestIndex, Npclist(NpcIndex).Char.CharIndex, vbYellow)
        

            'Si la quest pedia objetos, se los saca al personaje.
148         If .RequiredOBJs Then

150             For i = 1 To .RequiredOBJs
152                 Call QuitarObjetos(.RequiredOBJ(i).ObjIndex, .RequiredOBJ(i).Amount, UserIndex)
154             Next i

            End If
        
            'Se entrega la experiencia.
156         If .RewardEXP Then
158             If UserList(UserIndex).Stats.ELV < STAT_MAXELV Then
160                 UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + .RewardEXP
162                 Call WriteUpdateExp(UserIndex)
164                 Call CheckUserLevel(UserIndex)
166                  Call WriteConsoleMsg(UserIndex, "Has ganado " & .RewardEXP & " puntos de experiencia como recompensa.", FontTypeNames.FONTTYPE_INFO)
                Else
168                 Call WriteConsoleMsg(UserIndex, "No se te ha dado experiencia porque eres nivel máximo.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If
        
            'Se entrega el oro.
170         If .RewardGLD Then
172             UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + .RewardGLD
174             Call WriteConsoleMsg(UserIndex, "Has ganado " & .RewardGLD & " monedas de oro como recompensa.", FontTypeNames.FONTTYPE_INFO)

            End If
        
            'Si hay recompensa de objetos, se entregan.
176         If .RewardOBJs > 0 Then

178             For i = 1 To .RewardOBJs

180                 If .RewardOBJ(i).Amount Then
182                     Call MeterItemEnInventario(UserIndex, .RewardOBJ(i))
184                     Call WriteConsoleMsg(UserIndex, "Has recibido " & QuestList(QuestIndex).RewardOBJ(i).Amount & " " & ObjData(QuestList(QuestIndex).RewardOBJ(i).ObjIndex).Name & " como recompensa.", FontTypeNames.FONTTYPE_INFO)

                    End If

186             Next i

            End If
        
188         Call WriteUpdateGold(UserIndex)
    
            'Actualizamos el personaje
190         Call UpdateUserInv(True, UserIndex, 0)
    
            'Limpiamos el slot de quest.
192         Call CleanQuestSlot(UserIndex, QuestSlot)
        
            'Ordenamos las quests
194         Call ArrangeUserQuests(UserIndex)
        
196         If .Repetible = 0 Then
                'Se agrega que el usuario ya hizo esta quest.
198             Call AddDoneQuest(UserIndex, QuestIndex)
200             Call WriteUpdateNPCSimbolo(UserIndex, NpcIndex, 4)
            End If
        
        End With

        
        Exit Sub

FinishQuest_Err:
202     'Call RegistrarError(Err.Number, Err.Description, "ModQuest.FinishQuest", Erl)
204     Resume Next
        
End Sub
 
Public Sub AddDoneQuest(ByVal UserIndex As Integer, ByVal QuestIndex As Integer)
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Agrega la quest QuestIndex a la lista de quests hechas.
'Last modified: 28/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    With UserList(UserIndex).QuestStats
        .NumQuestsDone = .NumQuestsDone + 1
        ReDim Preserve .QuestsDone(1 To .NumQuestsDone)
        .QuestsDone(.NumQuestsDone) = QuestIndex
    End With
End Sub
 
Public Function UserDoneQuest(ByVal UserIndex As Integer, ByVal QuestIndex As Integer) As Boolean
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Verifica si el usuario hizo la quest QuestIndex.
'Last modified: 28/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Dim i As Integer
    With UserList(UserIndex).QuestStats
        If .NumQuestsDone Then
            For i = 1 To .NumQuestsDone
                If .QuestsDone(i) = QuestIndex Then
                    UserDoneQuest = True
                    Exit Function
                End If
            Next i
        End If
    End With
    
    UserDoneQuest = False
        
End Function
 
Public Sub CleanQuestSlot(ByVal UserIndex As Integer, ByVal QuestSlot As Integer)
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Limpia un slot de quest de un usuario.
'Last modified: 28/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Dim i As Integer
 
    With UserList(UserIndex).QuestStats.Quests(QuestSlot)
        If .QuestIndex Then
            If QuestList(.QuestIndex).RequiredNPCs Then
                For i = 1 To QuestList(.QuestIndex).RequiredNPCs
                    .NPCsKilled(i) = 0
                Next i
            End If
        End If
        .QuestIndex = 0
    End With
End Sub
 
Public Sub ResetQuestStats(ByVal UserIndex As Integer)
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Limpia todos los QuestStats de un usuario
'Last modified: 28/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Dim i As Integer
 
    For i = 1 To MAXUSERQUESTS
        Call CleanQuestSlot(UserIndex, i)
    Next i
    
    With UserList(UserIndex).QuestStats
        .NumQuestsDone = 0
        Erase .QuestsDone
    End With
End Sub
 
Public Sub HandleQuest(ByVal UserIndex As Integer)
        
        On Error GoTo HandleQuest_Err
        

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Maneja el paquete Quest.
        'Last modified: 28/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        Dim NpcIndex As Integer

        Dim tmpByte  As Byte
 
        'Leemos el paquete
    
100     Call UserList(UserIndex).incomingData.ReadInteger
 
102     NpcIndex = UserList(UserIndex).flags.TargetNPC
    
104     If NpcIndex = 0 Then Exit Sub
    
        'Esta el personaje en la distancia correcta?
106     If Distancia(UserList(UserIndex).Pos, Npclist(NpcIndex).Pos) > 5 Then
108         Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
        'El NPC hace quests?
110     If Npclist(NpcIndex).NumQuest = 0 Then
112         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tengo ninguna mision para ti.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
            Exit Sub

        End If
    
        'El personaje ya hizo la quest?
      '  If UserDoneQuest(UserIndex, Npclist(NpcIndex).QuestNumber) Then
        '    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("Ya has hecho una mision para mi.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
         '   Exit Sub

       ' End If
        
        
        
        
        
        
 
        'El personaje tiene suficiente nivel?
       ' If UserList(UserIndex).Stats.ELV < QuestList(Npclist(NpcIndex).QuestNumber).RequiredLevel Then
          '  Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("Debes ser por lo menos nivel " & QuestList(Npclist(NpcIndex).QuestNumber).RequiredLevel & " para emprender esta mision.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
         '   Exit Sub

        'End If
    
        'A esta altura ya analizo todas las restricciones y esta preparado para el handle propiamente dicho
 
       ' tmpByte = TieneQuest(UserIndex, Npclist(NpcIndex).QuestNumber)
    
      '  If tmpByte Then
            'El usuario esta haciendo la quest, entonces va a hablar con el NPC para recibir la recompensa.
         '   Call FinishQuest(UserIndex, Npclist(NpcIndex).QuestNumber, tmpByte)
      '  Else
            'El usuario no esta haciendo la quest, entonces primero recibe un informe con los detalles de la mision.
         '   tmpByte = FreeQuestSlot(UserIndex)
        
            'El personaje tiene algun slot de quest para la nueva quest?
         '   If tmpByte = 0 Then
114             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("Estas haciendo demasiadas misiones. Vuelve cuando hayas completado alguna.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
         ''       Exit Sub

       '     End If
        
            'Enviamos los detalles de la quest
         '   Call WriteQuestDetails(UserIndex, Npclist(NpcIndex).QuestNumber)

       ' End If

        
        Exit Sub

HandleQuest_Err:
116     'Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleQuest", Erl)
118     Resume Next
        
End Sub
 
Public Sub LoadQuests()

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Carga el archivo QUESTS.DAT en el array QuestList.
        'Last modified: 27/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        On Error GoTo ErrorHandler

        Dim Reader    As clsIniReader

        Dim NumQuests As Integer

        Dim tmpStr    As String

        Dim i         As Integer

        Dim j         As Integer
    
        'Cargamos el clsIniManager en memoria
100     Set Reader = New clsIniReader
    
        'Lo inicializamos para el archivo Quests.DAT
102     Call Reader.Initialize(DatPath & "Quests.DAT")
    
        'Redimensionamos el array
104     NumQuests = Reader.GetValue("INIT", "NumQuests")
106     ReDim QuestList(1 To NumQuests)
    
        'Cargamos los datos
108     For i = 1 To NumQuests

110         With QuestList(i)
112             .Nombre = Reader.GetValue("QUEST" & i, "Nombre")
114             .desc = Reader.GetValue("QUEST" & i, "Desc")
116             .RequiredLevel = val(Reader.GetValue("QUEST" & i, "RequiredLevel"))
            
118             .RequiredQuest = val(Reader.GetValue("QUEST" & i, "RequiredQuest"))
            
120             .DescFinal = Reader.GetValue("QUEST" & i, "DescFinal")
            
122             .NextQuest = Reader.GetValue("QUEST" & i, "NextQuest")
            
                'CARGAMOS OBJETOS REQUERIDOS
124             .RequiredOBJs = val(Reader.GetValue("QUEST" & i, "RequiredOBJs"))

126             If .RequiredOBJs > 0 Then
128                 ReDim .RequiredOBJ(1 To .RequiredOBJs)

130                 For j = 1 To .RequiredOBJs
132                     tmpStr = Reader.GetValue("QUEST" & i, "RequiredOBJ" & j)
                    
134                     .RequiredOBJ(j).ObjIndex = val(ReadField(1, tmpStr, 45))
136                     .RequiredOBJ(j).Amount = val(ReadField(2, tmpStr, 45))
138                 Next j

                End If
            
                'CARGAMOS NPCS REQUERIDOS
140             .RequiredNPCs = val(Reader.GetValue("QUEST" & i, "RequiredNPCs"))

142             If .RequiredNPCs > 0 Then
144                 ReDim .RequiredNPC(1 To .RequiredNPCs)

146                 For j = 1 To .RequiredNPCs
148                     tmpStr = Reader.GetValue("QUEST" & i, "RequiredNPC" & j)
                    
150                     .RequiredNPC(j).NpcIndex = val(ReadField(1, tmpStr, 45))
152                     .RequiredNPC(j).Amount = val(ReadField(2, tmpStr, 45))
154                 Next j

                End If
            
            
            
                'CARGAMOS NPCS TARGET REQUERIDOS
156             .RequiredTargetNPCs = val(Reader.GetValue("QUEST" & i, "RequiredTargetNPCs"))

158             If .RequiredTargetNPCs > 0 Then
160                 ReDim .RequiredTargetNPC(1 To .RequiredTargetNPCs)

162                 For j = 1 To .RequiredTargetNPCs
164                     tmpStr = Reader.GetValue("QUEST" & i, "RequiredTargetNPC" & j)
                    
166                     .RequiredTargetNPC(j).NpcIndex = val(ReadField(1, tmpStr, 45))
168                     .RequiredTargetNPC(j).Amount = 1
170                 Next j

                End If
            
            
            
            
            
172             .RewardGLD = val(Reader.GetValue("QUEST" & i, "RewardGLD"))
174             .RewardEXP = val(Reader.GetValue("QUEST" & i, "RewardEXP"))
176             .Repetible = val(Reader.GetValue("QUEST" & i, "Repetible"))
            
                'CARGAMOS OBJETOS DE RECOMPENSA
178             .RewardOBJs = val(Reader.GetValue("QUEST" & i, "RewardOBJs"))

180             If .RewardOBJs > 0 Then
182                 ReDim .RewardOBJ(1 To .RewardOBJs)

184                 For j = 1 To .RewardOBJs
186                     tmpStr = Reader.GetValue("QUEST" & i, "RewardOBJ" & j)
                    
188                     .RewardOBJ(j).ObjIndex = val(ReadField(1, tmpStr, 45))
190                     .RewardOBJ(j).Amount = val(ReadField(2, tmpStr, 45))
192                 Next j

                End If

            End With

194     Next i
    
        'Eliminamos la clase
196     Set Reader = Nothing
        Exit Sub
                    
ErrorHandler:
198     MsgBox "Error cargando el archivo QUESTS.DAT.", vbOKOnly + vbCritical

End Sub

Public Sub LoadQuestStats(ByVal UserIndex As Integer, ByRef UserFile As clsMySQLRecordSet)
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Carga las QuestStats del usuario.
'Last modified: 28/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim i As Integer
    Dim j As Integer
    Dim tmpStr As String
    Dim fields() As String
    For i = 1 To MAXUSERQUESTS
        With UserList(UserIndex).QuestStats.Quests(i)
            tmpStr = UserFile("Q" & i)
            If tmpStr = vbNullString Then
                .QuestIndex = 0

            Else
                fields = Split(tmpStr, "-")
                .QuestIndex = val(fields(0))
                If .QuestIndex Then
                    If QuestList(.QuestIndex).RequiredNPCs Then
                        ReDim .NPCsKilled(1 To QuestList(.QuestIndex).RequiredNPCs)

                        For j = 1 To QuestList(.QuestIndex).RequiredNPCs
                            .NPCsKilled(j) = val(ReadField(j + 1, tmpStr, 45))
                        Next j
                    End If
                End If
            End If
        End With
    Next i

    With UserList(UserIndex).QuestStats
        tmpStr = UserFile("QuestsDone")
        If tmpStr = vbNullString Then
            .NumQuestsDone = 0

        Else

            fields = Split(tmpStr, "-")

            .NumQuestsDone = val(fields(0))

            If .NumQuestsDone Then
                ReDim .QuestsDone(1 To .NumQuestsDone)
                For i = 1 To .NumQuestsDone
                    .QuestsDone(i) = val(fields(i))
                Next i
            End If
        End If
    End With

End Sub
 
Public Sub SaveQuestStats(ByVal UserIndex As Integer)
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Guarda las QuestStats del usuario.
'Last modified: 29/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Dim i As Integer
Dim j As Integer
Dim tmpStr As String
 
    For i = 1 To MAXUSERQUESTS
        With UserList(UserIndex).QuestStats.Quests(i)
            tmpStr = .QuestIndex
            
            If .QuestIndex Then
                If QuestList(.QuestIndex).RequiredNPCs Then
                    For j = 1 To QuestList(.QuestIndex).RequiredNPCs
                        tmpStr = tmpStr & "-" & .NPCsKilled(j)
                    Next j
                End If
            End If
           Call Execute("UPDATE pjs SET Q" & "i" & "= " & tmpStr & " WHERE Id=" & UserList(UserIndex).MySQLId)
            'Call WriteVar(UserFile, "QUESTS", "Q" & i, tmpStr)
        End With
    Next i
    
    With UserList(UserIndex).QuestStats
        tmpStr = .NumQuestsDone
        
        If .NumQuestsDone Then
            For i = 1 To .NumQuestsDone
                tmpStr = tmpStr & "-" & .QuestsDone(i)
            Next i
        End If
        Call Execute("UPDATE pjs SET QuestsDone= " & tmpStr & " WHERE Id=" & UserList(UserIndex).MySQLId)
        'Call WriteVar(UserFile, "QUESTS", "QuestsDone", tmpStr)
    End With
End Sub
 
Public Sub HandleQuestListRequest(ByVal UserIndex As Integer)
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Maneja el paquete QuestListRequest.
        'Last modified: 30/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        
        On Error GoTo HandleQuestListRequest_Err
        
 
        'Leemos el paquete
100     Call UserList(UserIndex).incomingData.ReadByte
    
102    ' If UserList(UserIndex).flags.BattleModo = 0 Then
104         Call WriteQuestListSend(UserIndex)
        'Else
106      '   Call WriteConsoleMsg(UserIndex, "No disponible aquí.", FontTypeNames.FONTTYPE_INFOIAO)

        'End If

        
        Exit Sub

HandleQuestListRequest_Err:
108     'Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleQuestListRequest", Erl)
110     Resume Next
        
End Sub
 
Public Sub ArrangeUserQuests(ByVal UserIndex As Integer)
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Ordena las quests del usuario de manera que queden todas al principio del arreglo.
'Last modified: 30/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Dim i As Integer
Dim j As Integer
 
    With UserList(UserIndex).QuestStats
        For i = 1 To MAXUSERQUESTS - 1
            If .Quests(i).QuestIndex = 0 Then
                For j = i + 1 To MAXUSERQUESTS
                    If .Quests(j).QuestIndex Then
                        .Quests(i) = .Quests(j)
                        Call CleanQuestSlot(UserIndex, j)
                        Exit For
                    End If
                Next j
            End If
        Next i
    End With
End Sub
 
Public Sub HandleQuestDetailsRequest(ByVal UserIndex As Integer)
        
        On Error GoTo HandleQuestDetailsRequest_Err
        

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Maneja el paquete QuestInfoRequest.
        'Last modified: 30/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        Dim QuestSlot As Byte
 
        'Leemos el paquete
100     Call UserList(UserIndex).incomingData.ReadByte
    
102     QuestSlot = UserList(UserIndex).incomingData.ReadByte
    
104     Call WriteQuestDetails(UserIndex, UserList(UserIndex).QuestStats.Quests(QuestSlot).QuestIndex, QuestSlot)

        
        Exit Sub

HandleQuestDetailsRequest_Err:
106    ' Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleQuestDetailsRequest", Erl)
108     Resume Next
        
End Sub
 
Public Sub HandleQuestAbandon(ByVal UserIndex As Integer)
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Maneja el paquete QuestAbandon.
        'Last modified: 31/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Leemos el paquete.
        
        On Error GoTo HandleQuestAbandon_Err
        
100     Call UserList(UserIndex).incomingData.ReadByte
    
        'Borramos la quest.
102     Call CleanQuestSlot(UserIndex, UserList(UserIndex).incomingData.ReadByte)
    
        'Ordenamos la lista de quests del usuario.
104     Call ArrangeUserQuests(UserIndex)
    
        'Enviamos la lista de quests actualizada.
106     Call WriteQuestListSend(UserIndex)

        
        Exit Sub

HandleQuestAbandon_Err:
108     'Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleQuestAbandon", Erl)
110     Resume Next
        
End Sub


Public Sub EnviarQuest(ByVal UserIndex As Integer)

    On Error GoTo EnviarQuest_Err


    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Maneja el paquete Quest.
    'Last modified: 28/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim NpcIndex As Integer

    Dim tmpByte As Byte

100 NpcIndex = UserList(UserIndex).flags.TargetNPC

102 If NpcIndex = 0 Then Exit Sub

    'Esta el personaje en la distancia correcta?
104 If Distancia(UserList(UserIndex).Pos, Npclist(NpcIndex).Pos) > 5 Then
106     Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    'El NPC hace quests?
108 If Npclist(NpcIndex).NumQuest = 0 Then
110     Call WriteChatOverHead(UserIndex, "No tengo ninguna mision para ti.", Npclist(NpcIndex).Char.CharIndex, vbYellow)
        Exit Sub

    End If


    'Hago un for para chequear si alguna de las misiones que da el NPC ya se completo.
    Dim q As Byte


112 For q = 1 To Npclist(NpcIndex).NumQuest
114     tmpByte = TieneQuest(UserIndex, Npclist(NpcIndex).QuestNumber(q))

116     If tmpByte Then
            'El usuario esta haciendo la quest, entonces va a hablar con el NPC para recibir la recompensa.
118         If FinishQuestCheck(UserIndex, Npclist(NpcIndex).QuestNumber(q), tmpByte) Then
120             Call FinishQuest(UserIndex, Npclist(NpcIndex).QuestNumber(q), tmpByte)
                Exit Sub
            End If

        End If

122 Next q
    ' Else
    'El usuario no esta haciendo la quest, entonces primero recibe un informe con los detalles de la mision.
    'tmpByte = FreeQuestSlot(UserIndex)

    'El personaje tiene algun slot de quest para la nueva quest?
    'If tmpByte = 0 Then
    '  Call WriteChatOverHead(UserIndex, "Estas haciendo demasiadas misiones. Vuelve cuando hayas completado alguna.", Npclist(NpcIndex).Char.CharIndex, vbWhite)
    '  Exit Sub

    ' End If

    'Enviamos los detalles de la quest
    'Call WriteQuestDetails(UserIndex, Npclist(NpcIndex).QuestNumber(1))

    '  End If

124 Call WriteNpcQuestListSend(UserIndex, NpcIndex)


    Exit Sub

EnviarQuest_Err:
126 'Call RegistrarError(Err.Number, Err.Description, "ModQuest.EnviarQuest", Erl)
128 Resume Next

End Sub


Public Function FinishQuestCheck(ByVal UserIndex As Integer, ByVal QuestIndex As Integer, ByVal QuestSlot As Byte) As Boolean
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Funcion para chequear si finalizo una quest
        'Ladder
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

        On Error GoTo FinishQuestCheck_Err
        
        Dim i              As Integer

        Dim InvSlotsLibres As Byte

        Dim NpcIndex       As Integer
 
100     NpcIndex = UserList(UserIndex).flags.TargetNPC
    
102     With QuestList(QuestIndex)

            'Comprobamos que tenga los objetos.
104         If .RequiredOBJs > 0 Then

106             For i = 1 To .RequiredOBJs

108                 If TieneObjetos(.RequiredOBJ(i).ObjIndex, .RequiredOBJ(i).Amount, UserIndex) = False Then
110                     FinishQuestCheck = False
                    
                        Exit Function

                    End If

112             Next i

            End If
        
            'Comprobamos que haya matado todas las criaturas.
114         If .RequiredNPCs > 0 Then

116             For i = 1 To .RequiredNPCs

118                 If .RequiredNPC(i).Amount > UserList(UserIndex).QuestStats.Quests(QuestSlot).NPCsKilled(i) Then
120                     FinishQuestCheck = False
                        Exit Function

                    End If

122             Next i

            End If
            
            'Comprobamos que haya targeteado todas las criaturas.
124      If .RequiredTargetNPCs > 0 Then

126          For i = 1 To .RequiredTargetNPCs

128              If .RequiredTargetNPC(i).Amount > UserList(UserIndex).QuestStats.Quests(QuestSlot).NPCsTarget(i) Then
130                  FinishQuestCheck = False
                        Exit Function

                    End If

132          Next i

            End If
            
        End With
        
        
134     FinishQuestCheck = True

        

        Exit Function

FinishQuestCheck_Err:
        'Call RegistrarError(Err.Number, Err.Description, "ModQuest.FinishQuestCheck", Erl)

End Function

