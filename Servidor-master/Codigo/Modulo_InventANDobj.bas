Attribute VB_Name = "InvNpc"
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
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo Inv & Obj
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Modulo para controlar los objetos y los inventarios.
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
Public Function TirarItemAlPiso(Pos As WorldPos, Obj As Obj, Optional NotPirata As Boolean = True) As WorldPos
On Error GoTo errhandler

    Dim NuevaPos As WorldPos
    NuevaPos.X = 0
    NuevaPos.Y = 0
    
    Tilelibre Pos, NuevaPos, Obj, NotPirata, True
    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
        Call MakeObj(Obj, Pos.map, NuevaPos.X, NuevaPos.Y)
    End If
    TirarItemAlPiso = NuevaPos

Exit Function
errhandler:

End Function

Public Sub NPC_TIRAR_ITEMS(ByRef NPC As NPC)
'TIRA TODOS LOS ITEMS DEL NPC
On Error Resume Next

If NPC.Invent.NroItems > 0 Then
    
    Dim i As Byte
    Dim MiObj As Obj
        
    For i = 1 To MAX_INVENTORY_SLOTS
        
        If NPC.Invent.Object(i).ObjIndex > 0 Then
            MiObj.Amount = NPC.Invent.Object(i).Amount
            MiObj.ObjIndex = NPC.Invent.Object(i).ObjIndex
            Call TirarItemAlPiso(NPC.Pos, MiObj)
        End If
          
    Next i
End If

End Sub

Function QuedanItems(ByVal NpcIndex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error Resume Next
'Call LogTarea("Function QuedanItems npcindex:" & NpcIndex & " objindex:" & ObjIndex)

Dim i As Integer
If Npclist(NpcIndex).Invent.NroItems > 0 Then
    For i = 1 To MAX_INVENTORY_SLOTS
        If Npclist(NpcIndex).Invent.Object(i).ObjIndex = ObjIndex Then
            QuedanItems = True
            Exit Function
        End If
    Next
End If
QuedanItems = False
End Function

''
' Gets the amount of a certain item that an npc has.
'
' @param npcIndex Specifies reference to npcmerchant
' @param ObjIndex Specifies reference to object
' @return   The amount of the item that the npc has
' @remarks This function reads the Npc.dat file
Function EncontrarCant(ByVal NpcIndex As Integer, ByVal ObjIndex As Integer) As Integer
'***************************************************
'Author: Unknown
'Last Modification: 03/09/08
'Last Modification By: Marco Vanotti (Marco)
' - 03/09/08 EncontrarCant now returns 0 if the npc doesn't have it (Marco)
'***************************************************
On Error Resume Next
'Devuelve la cantidad original del obj de un npc

Dim ln As String, npcfile As String
Dim i As Integer

npcfile = DatPath & "NPCs.dat"
 
For i = 1 To MAX_INVENTORY_SLOTS
    ln = GetVar(npcfile, "NPC" & Npclist(NpcIndex).Numero, "Obj" & i)
    If ObjIndex = val(ReadField(1, ln, 45)) Then
        EncontrarCant = val(ReadField(2, ln, 45))
        Exit Function
    End If
Next
                   
EncontrarCant = 0

End Function

Sub ResetNpcInv(ByVal NpcIndex As Integer)
On Error Resume Next

Dim i As Integer

Npclist(NpcIndex).Invent.NroItems = 0

For i = 1 To MAX_INVENTORY_SLOTS
   Npclist(NpcIndex).Invent.Object(i).ObjIndex = 0
   Npclist(NpcIndex).Invent.Object(i).Amount = 0
Next i

Npclist(NpcIndex).InvReSpawn = 0

End Sub

''
' Removes a certain amount of items from a slot of an npc's inventory
'
' @param npcIndex Specifies reference to npcmerchant
' @param Slot Specifies reference to npc's inventory's slot
' @param antidad Specifies amount of items that will be removed
Sub QuitarNpcInvItem(ByVal NpcIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 03/09/08
'Last Modification By: Marco Vanotti (Marco)
' - 03/09/08 Now this sub checks that te npc has an item before respawning it (Marco)
'***************************************************
Dim ObjIndex As Integer
Dim iCant As Integer
ObjIndex = Npclist(NpcIndex).Invent.Object(Slot).ObjIndex

    'Quita un Obj
    If ObjData(Npclist(NpcIndex).Invent.Object(Slot).ObjIndex).Crucial = 0 Then
        Npclist(NpcIndex).Invent.Object(Slot).Amount = Npclist(NpcIndex).Invent.Object(Slot).Amount - Cantidad
        
        If Npclist(NpcIndex).Invent.Object(Slot).Amount <= 0 Then
            Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems - 1
            Npclist(NpcIndex).Invent.Object(Slot).ObjIndex = 0
            Npclist(NpcIndex).Invent.Object(Slot).Amount = 0
            If Npclist(NpcIndex).Invent.NroItems = 0 And Npclist(NpcIndex).InvReSpawn <> 1 Then
               Call CargarInvent(NpcIndex) 'Reponemos el inventario
            End If
        End If
    Else
        Npclist(NpcIndex).Invent.Object(Slot).Amount = Npclist(NpcIndex).Invent.Object(Slot).Amount - Cantidad
        
        If Npclist(NpcIndex).Invent.Object(Slot).Amount <= 0 Then
            Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems - 1
            Npclist(NpcIndex).Invent.Object(Slot).ObjIndex = 0
            Npclist(NpcIndex).Invent.Object(Slot).Amount = 0
            
            If Not QuedanItems(NpcIndex, ObjIndex) Then
                'Check if the item is in the npc's dat.
                iCant = EncontrarCant(NpcIndex, ObjIndex)
                If iCant Then
                    Npclist(NpcIndex).Invent.Object(Slot).ObjIndex = ObjIndex
                    Npclist(NpcIndex).Invent.Object(Slot).Amount = iCant
                    Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems + 1
                End If
            End If
            
            If Npclist(NpcIndex).Invent.NroItems = 0 And Npclist(NpcIndex).InvReSpawn <> 1 Then
               Call CargarInvent(NpcIndex) 'Reponemos el inventario
            End If
        End If
    
    
    
    End If
End Sub

Sub CargarInvent(ByVal NpcIndex As Integer)

'Vuelve a cargar el inventario del npc NpcIndex
Dim LoopC As Integer
Dim ln As String
Dim npcfile As String

npcfile = DatPath & "NPCs.dat"

Npclist(NpcIndex).Invent.NroItems = val(GetVar(npcfile, "NPC" & Npclist(NpcIndex).Numero, "NROITEMS"))

For LoopC = 1 To Npclist(NpcIndex).Invent.NroItems
    ln = GetVar(npcfile, "NPC" & Npclist(NpcIndex).Numero, "Obj" & LoopC)
    Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
    Npclist(NpcIndex).Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
    
Next LoopC

End Sub


Public Sub DropObjQuest(ByRef NPC As NPC, ByRef UserIndex As Integer)
    'Dropeo por Quest
    'Ladder
    '3/12/2020
        On Error GoTo errhandler

100     If NPC.NumDropQuest = 0 Then Exit Sub
        
    
        Dim Dropeo       As Obj
        Dim QuestIndex As Integer
        Dim ObjIndex As Integer
        Dim Amount As Integer
        Dim Probabilidad As Byte

        Dim i As Byte
    
    
102     For i = 1 To NPC.NumDropQuest
    
104         QuestIndex = val(ReadField(1, NPC.DropQuest(i), Asc("-")))
106         ObjIndex = val(ReadField(2, NPC.DropQuest(i), Asc("-")))
108         Amount = val(ReadField(3, NPC.DropQuest(i), Asc("-")))
110         Probabilidad = val(ReadField(4, NPC.DropQuest(i), Asc("-")))
        
112         If QuestIndex <> 0 Then
114             If TieneQuest(UserIndex, QuestIndex) <> 0 Then
116                 Probabilidad = RandomNumber(1, Probabilidad) 'Tiro Item?
118                 If Probabilidad = 1 Then
120                     Dropeo.Amount = Amount
122                     Dropeo.ObjIndex = ObjIndex
124                     Call TirarItemAlPiso(NPC.Pos, Dropeo)
126                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SWING, NPC.Pos.X, NPC.Pos.Y))
                    End If
                End If
            End If
128     Next i

        Exit Sub

errhandler:
130     Call LogError("Error DropObjQuest al dropear el item " & ObjData(ObjIndex).Name & ", al usuario " & UserList(UserIndex).Name & ". " & Err.Description & ".")

End Sub

Public Sub DropNPCaNPC(ByRef NPC As NPC, ByRef UserIndex As Integer)

    On Error GoTo errhandler

100 If NPC.NumDropNPC = 0 Then Exit Sub



    Dim SpawnedNpc As Integer
    Dim NpcIndex As Integer
    Dim Amount As Integer
    Dim Probabilidad As Byte

    Dim i As Byte
    Dim Cant As Byte

102 For i = 1 To NPC.NumDropNPC

104     NpcIndex = val(ReadField(1, NPC.DropNPC(i), Asc("-")))
106     Amount = val(ReadField(2, NPC.DropNPC(i), Asc("-")))
108     Probabilidad = val(ReadField(3, NPC.DropNPC(i), Asc("-")))
110

112     If NpcIndex <> 0 Then
114
116         Probabilidad = RandomNumber(1, Probabilidad)    'Spaw NPC?
118         If Probabilidad = 1 Then
120             For Cant = 1 To Amount
                    SpawnedNpc = SpawnNpc(NpcIndex, NPC.Pos, True, False, NPC.zona)
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SWING, NPC.Pos.X, NPC.Pos.Y))

                Next Cant
            End If
        End If
128 Next i

    Exit Sub

errhandler:
130 Call LogError("Error al crear el NPC: " & NpcIndex & " en la Posicion Mapa: " & NPC.Pos.map & ", X: " & NPC.Pos.X & ", Y: " & NPC.Pos.Y & " Zona: " & NPC.zona)

End Sub
