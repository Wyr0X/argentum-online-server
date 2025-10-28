Attribute VB_Name = "modCastles"
' Argentum 20 Game Server
'
'    Copyright (C) 2023 Noland Studios LTD
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'    This program was based on Argentum Online 0.11.6
'    Copyright (C) 2002 M·rquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'
Option Explicit

Enum e_CastleState
    Built = 0
    Building
    Upgrading
    Destroyed
    Rebuilding
End Enum

Type t_Castle
    Location As t_WorldPos
    LocationInside As t_WorldPos
    LocationWarp As t_WorldPos
    CastleType As Byte ' 0 = inactive
    State As e_CastleState
    GuildOwner As Integer
    HP As Long
    InsideHP As Long
    InCombatCounter As Integer
    WaitTime As Long
End Type

Type t_CastleType
    GrhIndexBuilt As Long
    GrhIndexBuilding As Long
    GrhIndexDestroyed As Long
    GrhOffsetX As Integer
    GrhOffsetY As Integer
    Width As Byte
    Height As Byte
    MaxHP As Long
End Type

Private Castles() As t_Castle
Private CastleTypes() As t_CastleType

Private CastleInsideHP As Long
Private CastleInCombatTimeSeconds As Integer
Private CastleRegenerationPerSecond As Long
Private CastleInsideRegenerationPerSecond As Long
Private CastleBuildTimeSeconds As Long
Private CastleCollapseTimeSeconds As Long
Private CastleNotifyCollapseSecondsBefore As Long
Private CastleRebuildTimeSeconds As Long
Private CastleDefense As Integer
Public CastleWarpCooldown As Integer
Private CastleWarpTimeSeconds As Integer
Private CastleWarpMaxHPPercentage As Byte
Private CastleMaxInvasionTimeSeconds As Long

Const CASTLE_INSIDE_MASK As Long = 128

Public Sub LoadCastlesData()
    On Error GoTo LoadCastlesData_Err

    Dim Reader As New clsIniManager

    Call Reader.Initialize(DatPath & "Castles.dat")

    CastleInsideHP = val(Reader.GetValue("CastleData", "CastleInsideHP"))
    CastleInCombatTimeSeconds = val(Reader.GetValue("CastleData", "InCombatTimeSeconds"))
    CastleRegenerationPerSecond = val(Reader.GetValue("CastleData", "RegenerationPerSecond"))
    CastleInsideRegenerationPerSecond = val(Reader.GetValue("CastleData", "InsideRegenerationPerSecond"))
    CastleDefense = val(Reader.GetValue("CastleData", "Defense"))
    CastleWarpCooldown = val(Reader.GetValue("CastleData", "WarpCooldownSeconds"))
    CastleBuildTimeSeconds = val(Reader.GetValue("CastleData", "BuildTimeSeconds"))
    CastleCollapseTimeSeconds = val(Reader.GetValue("CastleData", "CollapseTimeSeconds"))
    CastleNotifyCollapseSecondsBefore = val(Reader.GetValue("CastleData", "NotifyCollapseSecondsBefore"))
    CastleRebuildTimeSeconds = val(Reader.GetValue("CastleData", "RebuildTimeSeconds"))
    CastleWarpTimeSeconds = val(Reader.GetValue("CastleData", "WarpTimeSeconds"))
    CastleWarpMaxHPPercentage = val(Reader.GetValue("CastleData", "WarpMaxHPPercentage"))
    CastleMaxInvasionTimeSeconds = val(Reader.GetValue("CastleData", "MaxInvasionTimeSeconds"))
    
    Dim NumLocations As Byte
    NumLocations = val(Reader.GetValue("CastleLocations", "NumLocations"))
    
    If NumLocations <= 0 Then
        ReDim Castles(0)
        Set Reader = Nothing
        Exit Sub
    End If
    
    ReDim Preserve Castles(1 To NumLocations)
    
    Dim Fields() As String
    
    Dim i As Byte
    For i = 1 To NumLocations
        Call RemoveCastleFromMap(i)
    
        Fields = Split(Reader.GetValue("CastleLocations", "LocationOut" & i), "-")
        If UBound(Fields) >= 2 Then
            Castles(i).Location.Map = Fields(0)
            Castles(i).Location.X = Fields(1)
            Castles(i).Location.Y = Fields(2)
        End If

        Fields = Split(Reader.GetValue("CastleLocations", "LocationIn" & i), "-")
        If UBound(Fields) >= 2 Then
            Castles(i).LocationInside.Map = Fields(0)
            Castles(i).LocationInside.X = Fields(1)
            Castles(i).LocationInside.Y = Fields(2)
        End If
        
        Fields = Split(Reader.GetValue("CastleLocations", "LocationWarp" & i), "-")
        If UBound(Fields) >= 2 Then
            Castles(i).LocationWarp.Map = Fields(0)
            Castles(i).LocationWarp.X = Fields(1)
            Castles(i).LocationWarp.Y = Fields(2)
        End If
    Next
    
    Dim NumCastleTypes As Byte
    NumCastleTypes = val(Reader.GetValue("CastleTypes", "NumCastleTypes"))
    
    If NumCastleTypes <= 0 Then
        ReDim CastleTypes(0)
        Set Reader = Nothing
        Exit Sub
    End If
    
    ReDim CastleTypes(1 To NumCastleTypes)
    
    For i = 1 To NumCastleTypes
        CastleTypes(i).GrhIndexBuilt = val(Reader.GetValue("CastleType" & i, "GrhIndexBuilt"))
        CastleTypes(i).GrhIndexBuilding = val(Reader.GetValue("CastleType" & i, "GrhIndexBuilding"))
        CastleTypes(i).GrhIndexDestroyed = val(Reader.GetValue("CastleType" & i, "GrhIndexDestroyed"))
        CastleTypes(i).GrhOffsetX = val(Reader.GetValue("CastleType" & i, "GrhOffsetX"))
        CastleTypes(i).GrhOffsetY = val(Reader.GetValue("CastleType" & i, "GrhOffset"))
        CastleTypes(i).Width = val(Reader.GetValue("CastleType" & i, "Width"))
        CastleTypes(i).Height = val(Reader.GetValue("CastleType" & i, "Height"))
        CastleTypes(i).MaxHP = val(Reader.GetValue("CastleType" & i, "MaxHP"))
    Next

    Set Reader = Nothing

    Call LoadCastlesDB

    Exit Sub
LoadCastlesData_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.LoadCastlesData", Erl)
End Sub

' Returns Castle index (0 if not found - only one castle per map)
Private Function FindCastleByMap(ByVal Map As Integer) As Byte
    Dim i As Byte
    For i = 1 To UBound(Castles)
        If Castles(i).Location.Map = Map Then
            FindCastleByMap = i
            Exit Function
        End If
    Next
End Function

' Returns Castle index (0 if not found)
Private Function FindCastle(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal CastleType As Byte) As Byte
    If CastleType = 0 Then Exit Function

    Dim CastleId As Byte
    CastleId = FindCastleByMap(Map)
    If CastleId = 0 Then Exit Function
    
    Dim Width As Byte
    Dim Height As Byte
    With CastleTypes(CastleType)
        Width = .Width
        Height = .Height
    End With
    
    With Castles(CastleId)
        If X < .Location.X - (Width - 1) \ 2 Or X > .Location.X + Width \ 2 Then Exit Function
        If Y <= .Location.Y - Height Or Y > .Location.Y Then Exit Function
        FindCastle = CastleId
    End With
    
End Function

' Returns Castle index (0 if not found)
Private Function FindCastleByGuild(ByVal GuildIndex As Integer) As Byte
    Dim i As Byte
    For i = 1 To UBound(Castles)
        If Castles(i).GuildOwner = GuildIndex Then
            If Castles(i).CastleType <> 0 Then
                FindCastleByGuild = i
            End If
            Exit Function
        End If
    Next
End Function

' Returns Castle index (0 if not found - only one castle per map)
Private Function FindCastleInsideByMap(ByVal Map As Integer) As Byte
    Dim i As Byte
    For i = 1 To UBound(Castles)
        If Castles(i).LocationInside.Map = Map Then
            FindCastleInsideByMap = i
            Exit Function
        End If
    Next
End Function

Public Function CanBuildCastle(ByVal UserIndex As Integer, ByVal CastleType As Byte) As Boolean
    On Error GoTo CanBuildCastle_Err
    
    With UserList(UserIndex)
    
        If .flags.Muerto = 1 Then
            ' Estas muerto
            Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        If .GuildIndex = 0 Then
            ' No perteneces a ningun clan.
            Call WriteLocaleMsg(UserIndex, "720", e_FontTypeNames.FONTTYPE_INFOIAO)
            Exit Function
        End If
        
        Dim CastleId As Byte
        CastleId = FindCastleByMap(.Pos.Map)
        
        If CastleId = 0 Then
            Call WriteLocaleMsg(UserIndex, MsgCantPlaceCastleOnThisMap, e_FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        If Castles(CastleId).CastleType <> 0 Then
            Call WriteLocaleMsg(UserIndex, MsgAlreadyACastleOnThisMap, e_FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        If CastleType < LBound(CastleTypes) Or CastleType > UBound(CastleTypes) Then
            Call LogError("Error in modCastles.CanBuildCastle: invalid CastleType for castle permit")
            Exit Function
        End If
        
        If FindCastleByGuild(.GuildIndex) <> 0 Then
            Call WriteLocaleMsg(UserIndex, MsgYourGuildAlreadyHasCastle, e_FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If

    End With
    
    CanBuildCastle = True
    
    Exit Function
CanBuildCastle_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.CanBuildCastle", Erl)
End Function

Public Sub UseCastlePermit(ByVal UserIndex As Integer, ByRef obj As t_ObjData)
    On Error GoTo UseCastlePermit_Err

    With UserList(UserIndex)
        If Not CanBuildCastle(UserIndex, obj.cdType) Then Exit Sub
        
        Dim CastleId As Byte
        CastleId = FindCastleByMap(.Pos.Map)

        Dim GrhIndex As Long:       GrhIndex = CastleTypes(obj.cdType).GrhIndexBuilt
        Dim GrhOffsetX As Integer:  GrhOffsetX = CastleTypes(obj.cdType).GrhOffsetX
        Dim GrhOffsetY As Integer:  GrhOffsetY = CastleTypes(obj.cdType).GrhOffsetY
        Dim Width As Byte:          Width = CastleTypes(obj.cdType).Width
        Dim Height As Byte:         Height = CastleTypes(obj.cdType).Height

        Call WriteCastleBuildPosition(UserIndex, Castles(CastleId).Location.X, Castles(CastleId).Location.Y, _
            GrhIndex, GrhOffsetX, GrhOffsetY, Width, Height)

        .flags.UsingItemSlot = .flags.TargetObjInvSlot
        Call WriteWorkRequestTarget(UserIndex, e_Skill.TargetableItem)
    End With
    
    Exit Sub
UseCastlePermit_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.UseCastlePermit", Erl)
End Sub

Public Sub BuildCastle(ByVal UserIndex As Integer, ByVal X As Integer, ByVal Y As Integer)
    On Error GoTo BuildCastle_Err
    
    With UserList(UserIndex)

        Dim PermitObjIndex As Integer: PermitObjIndex = .invent.Object(.flags.UsingItemSlot).ObjIndex
    
        If PermitObjIndex = 0 Then Exit Sub
        If ObjData(PermitObjIndex).OBJType <> e_OBJType.otCastlePermit Then Exit Sub
    
        Dim CastleType As Byte
        CastleType = ObjData(PermitObjIndex).cdType
        If CastleType = 0 Then Exit Sub

        If Not CanBuildCastle(UserIndex, CastleType) Then Exit Sub
    
        Dim CastleId As Byte
        CastleId = FindCastle(.Pos.Map, X, Y, CastleType)
        
        If CastleId = 0 Then
            Call WriteLocaleMsg(UserIndex, MsgCantPlaceCastleOnThisPos, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Call RemoveItemFromInventory(UserIndex, .flags.UsingItemSlot)

        Call LogCastles(.Name & " <" & GuildName(.GuildIndex) & "> built a castle on map " & .Pos.Map)

    End With

    With Castles(CastleId)

        .CastleType = CastleType
        .GuildOwner = UserList(UserIndex).GuildIndex

        Call SetCastleState(CastleId, e_CastleState.Building)

        Call PutCastleOnMap(CastleId)

    End With
    
    With UserList(UserIndex)
    
        Call WriteLocaleMsg(UserIndex, MsgYouStartedBuildingTheCastle, e_FontTypeNames.FONTTYPE_INFO)
        Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageLocaleMsg(MsgGuildMemberStartedBuildingTheCastle, .Name, e_FontTypeNames.FONTTYPE_GUILD))
        Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave(43, NO_3D_SOUND, NO_3D_SOUND))
    
    End With
    
    Exit Sub
BuildCastle_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.BuildCastle", Erl)
End Sub

Private Function CanAttackCastle(ByVal UserIndex As Integer, ByVal CastleId As Byte, ByVal Inside As Boolean) As Boolean
    On Error GoTo CanAttackCastle_Err

    If Castles(CastleId).CastleType = 0 Then Exit Function

    With UserList(UserIndex)

        If .GuildIndex = 0 Then
            Call WriteLocaleMsg(UserIndex, MsgToAttackCastleYouMustBeInAGuild, e_FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If

        If Castles(CastleId).GuildOwner = .GuildIndex And Not EsGM(UserIndex) Then
            Call WriteLocaleMsg(UserIndex, MsgCantAttackYourOwnCastle, e_FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If

        If FindCastleByGuild(.GuildIndex) <> 0 And Not EsGM(UserIndex) Then
            Call WriteLocaleMsg(UserIndex, MsgYourGuildAlreadyHasCastle, e_FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If

        If .flags.EnConsulta Then
            Call WriteLocaleMsg(UserIndex, "1047", e_FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If

        If .flags.Maldicion = 1 Then
            Call WriteLocaleMsg(UserIndex, "1049", e_FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If

        If .flags.Montado = 1 Then
            Call WriteLocaleMsg(UserIndex, "1050", e_FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If

    End With

    With Castles(CastleId)

        If .State = e_CastleState.Building Or (Not Inside And .State = e_CastleState.Rebuilding) Then
            Call WriteLocaleMsg(UserIndex, MsgCantAttackCastleBuilding, e_FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If

        If Not Inside Then
            If .State = e_CastleState.Destroyed Then
                Call WriteLocaleMsg(UserIndex, MsgCantAttackCastleDestroyed, e_FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If
        End If

    End With

    CanAttackCastle = True

    Exit Function
CanAttackCastle_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.CanAttackCastle", Erl)
End Function

Public Function UserPhysicalAttackCastle(ByVal UserIndex As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
    On Error GoTo UserPhysicalAttackCastle_Err

    With UserList(UserIndex)

        Dim CastleId As Byte
        Dim Inside As Boolean

        CastleId = FindCastleByMap(.Pos.Map)

        If CastleId = 0 Then
            CastleId = FindCastleInsideByMap(.Pos.Map)
            If CastleId = 0 Then Exit Function

            Inside = True
        End If

        If Not CanAttackCastle(UserIndex, CastleId, Inside) Then Exit Function

        Dim DamageBase As Long
        DamageBase = GetUserDamage(UserIndex)

        Dim Damage As Long
        Damage = DamageBase - max(0, CastleDefense - GetArmorPenetration(UserIndex, CastleDefense))

        Dim Color As Long
        Color = vbRed

        Damage = Damage * UserMod.GetPhysicalDamageModifier(UserList(UserIndex))
        If Damage < 0 Then Damage = 0

        Dim DamageExtra As Long
        ' Critical hit
        If PuedeGolpeCritico(UserIndex) Then
            If RandomNumber(1, 100) <= ProbabilidadGolpeCritico(UserIndex) Then
                DamageExtra = DamageBase * 0.33
                DamageExtra = DamageExtra * UserMod.GetPhysicalDamageModifier(UserList(UserIndex))

                If .ChatCombate = 1 Then
                    Call WriteLocaleMsg(UserIndex, 383, e_FontTypeNames.FONTTYPE_INFOBOLD, PonerPuntos(Damage) & "¨" & PonerPuntos(DamageExtra))
                End If

                Color = RGB(225, 165, 0)
            End If

        ' Stab
        ElseIf PuedeApuÒalar(UserIndex) Then
            If RandomNumber(1, 100) <= ProbabilidadApuÒalar(UserIndex, 1) Then
                Dim min_stab_npc As Double
                Dim max_stab_npc As Double
                min_stab_npc = GetStabbingNPCMinForClass(UserList(UserIndex).clase)
                max_stab_npc = GetStabbingNPCMaxForClass(UserList(UserIndex).clase)

                DamageExtra = Damage * (Rnd * (max_stab_npc - min_stab_npc) + min_stab_npc)

                If .ChatCombate = 1 Then
                    Call WriteLocaleMsg(UserIndex, MsgYouHaveStabbedFor, e_FontTypeNames.FONTTYPE_INFOBOLD, PonerPuntos(Damage) & "¨" & PonerPuntos(DamageExtra))
                End If

                Color = vbYellow
            End If

            Call SubirSkill(UserIndex, Apu√±alar)
        End If
        If DamageExtra > 0 Then
            Damage = Damage + DamageExtra
        End If

        If .ChatCombate = 1 Then
            Call WriteLocaleMsg(UserIndex, MsgCastleAttacked, e_FontTypeNames.FONTTYPE_FIGHT, GuildName(Castles(CastleId).GuildOwner) & "¨" & PonerPuntos(Damage))
        End If

        Call SendData(SendTarget.ToCastleArea, GetCastleSndIndex(CastleId, Inside), PrepareMessagePlayWave(IIf(RandomNumber(0, 1) = 0, 15, 42), X, Y))

        Call UserAttackedCastle(UserIndex, CastleId, Inside, Damage, Color, X, Y)

        Call SubirSkillDeArmaActual(UserIndex)

    End With

    UserPhysicalAttackCastle = True
    
    Exit Function
UserPhysicalAttackCastle_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.UserPhysicalAttackCastle", Erl)
End Function

Public Function UserMagicAttackCastle(ByVal UserIndex As Integer, ByVal Spell As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
    On Error GoTo UserMagicAttackCastle_Err
    
    With UserList(UserIndex)
    
        Dim CastleId As Byte
        Dim Inside As Boolean

        CastleId = FindCastleByMap(.Pos.Map)

        If CastleId = 0 Then
            CastleId = FindCastleInsideByMap(.Pos.Map)
            If CastleId = 0 Then Exit Function

            Inside = True
        End If

        If Not CanAttackCastle(UserIndex, CastleId, Inside) Then Exit Function
        
        If Not IsSet(Hechizos(Spell).Effects, e_SpellEffects.eDoDamage) Then
            Call WriteLocaleMsg(UserIndex, MsgOnlyDamageSpellsToCastle, e_FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If

        Dim Damage As Long
        Damage = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHP)
        Damage = Damage + Porcentaje(Damage, 3 * .Stats.ELV)
            
        Dim MagicPenetration As Integer
        If UserList(UserIndex).invent.WeaponEqpObjIndex > 0 Then
            Damage = Damage + Porcentaje(Damage, ObjData(.invent.WeaponEqpObjIndex).MagicDamageBonus)
            MagicPenetration = ObjData(.invent.WeaponEqpObjIndex).MagicPenetration
            Damage = Damage + ObjData(.invent.WeaponEqpObjIndex).MagicAbsoluteBonus
        End If
        If .invent.MagicoObjIndex > 0 Then
            Damage = Damage + Porcentaje(Damage, ObjData(.invent.MagicoObjIndex).MagicDamageBonus)
            Damage = Damage + ObjData(.invent.MagicoObjIndex).MagicAbsoluteBonus
            MagicPenetration = MagicPenetration + ObjData(.invent.MagicoObjIndex).MagicPenetration
        End If
        ' Magic Damage ring
        If .invent.Da√±oMagicoEqpObjIndex > 0 Then
            Damage = Damage + Porcentaje(Damage, ObjData(.invent.Da√±oMagicoEqpObjIndex).MagicDamageBonus)
            Damage = Damage + ObjData(.invent.Da√±oMagicoEqpObjIndex).MagicAbsoluteBonus
            MagicPenetration = MagicPenetration + ObjData(.invent.Da√±oMagicoEqpObjIndex).MagicPenetration
        End If

        Damage = Damage * UserMod.GetMagicDamageModifier(UserList(UserIndex))
        If Damage < 0 Then Damage = 0

        If Hechizos(Spell).FXgrh > 0 Then 'Envio Fx?
            Call SendData(SendTarget.ToCastleArea, GetCastleSndIndex(CastleId, Inside), PrepareMessageFxPiso(Hechizos(Spell).FXgrh, X, Y))
        End If
    
        If Hechizos(Spell).Particle > 0 Then 'Envio Particula?
            Call SendData(SendTarget.ToCastleArea, GetCastleSndIndex(CastleId, Inside), PrepareMessageParticleFXToFloor(X, Y, Hechizos(Spell).Particle, Hechizos(Spell).TimeParticula))
        End If
    
        If Hechizos(Spell).wav <> 0 Then
            Call SendData(SendTarget.ToCastleArea, GetCastleSndIndex(CastleId, Inside), PrepareMessagePlayWave(Hechizos(Spell).wav, X, Y))
        End If
        
        If .ChatCombate = 1 Then
            Call WriteConsoleMsg(UserIndex, "HecMSGC*" & Spell, e_FontTypeNames.FONTTYPE_FIGHT)
        End If
        
        Call UserAttackedCastle(UserIndex, CastleId, Inside, Damage, vbRed, X, Y)

        If Not IsSet(Hechizos(Spell).SpellRequirementMask, eIsSkill) Then
            Call SubirSkill(UserIndex, e_Skill.Magia)
        End If
        
        .Stats.MinMAN = .Stats.MinMAN - ManaHechizoPorClase(UserIndex, Hechizos(Spell), Spell)
        If .Stats.MinMAN < 0 Then .Stats.MinMAN = 0
    
        If Hechizos(Spell).RequiredHP > 0 Then
            Call UserMod.ModifyHealth(UserIndex, -Hechizos(Spell).RequiredHP, 1)
        End If

        .Stats.MinSta = .Stats.MinSta - Hechizos(Spell).StaRequerido
        If .Stats.MinSta < 0 Then .Stats.MinSta = 0

        Call WriteUpdateMana(UserIndex)
        Call WriteUpdateSta(UserIndex)
        
    End With

    UserMagicAttackCastle = True
    
    Exit Function
UserMagicAttackCastle_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.UserMagicAttackCastle", Erl)
End Function

Public Sub UserAttackedCastle(ByVal UserIndex As Integer, ByVal CastleId As Byte, ByVal Inside As Boolean, ByVal Damage As Long, ByVal Color As Long, ByVal X As Integer, ByVal Y As Integer)
    On Error GoTo UserAttackedCastle_Err

    With Castles(CastleId)

        .InCombatCounter = CastleInCombatTimeSeconds

        Call WriteLocaleMsg(UserIndex, MsgDamageToCastle, e_FontTypeNames.FONTTYPE_FIGHT, PonerPuntos(Damage))

        Call SendData(SendTarget.ToCastleArea, GetCastleSndIndex(CastleId, Inside), PrepareMessageTextOverTile(PonerPuntos(Damage), X, Y, Color))

        If Inside Then
            .InsideHP = .InsideHP - Damage

            If .InsideHP <= 0 Then
                .InsideHP = 0

                Call ConqueredCastle(UserIndex, CastleId)
            End If
        Else
            .HP = .HP - Damage

            If .HP <= 0 Then
                .HP = 0

                Call LogCastles(UserList(UserIndex).Name & " <" & GuildName(UserList(UserIndex).GuildIndex) & "> destroyed the gates of the castle from guild " & GuildName(.GuildOwner) & " on map " & .Location.Map)

                Call SetCastleState(CastleId, e_CastleState.Destroyed)
                Call OpenCastleGates(CastleId)
            End If
        End If

        Call SendData(SendTarget.ToCastleArea, GetCastleSndIndex(CastleId, Inside), PrepareMessage_UpdateCastleHP(GetCastleHPPercentage(CastleId, Inside)))

    End With

    Exit Sub
UserAttackedCastle_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.UserAttackedCastle", Erl)
End Sub

Private Sub ConqueredCastle(ByVal UserIndex As Integer, ByVal CastleId As Byte)
    On Error GoTo ConqueredCastle_Err

    With Castles(CastleId)

        ' Notify loser guild
        Dim NewOwnerName As String
        NewOwnerName = GuildName(UserList(UserIndex).GuildIndex)
        Call SendData(SendTarget.ToGuildMembers, .GuildOwner, PrepareMessageLocaleMsg(MsgCastleLostToGuild, NewOwnerName, e_FontTypeNames.FONTTYPE_GUILD))
        Call SendData(SendTarget.ToGuildMembers, .GuildOwner, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))

        ' Notify winner guild
        Dim OldOwnerName As String
        OldOwnerName = GuildName(.GuildOwner)
        Call SendData(SendTarget.ToGuildMembers, UserList(UserIndex).GuildIndex, PrepareMessageLocaleMsg(MsgYourGuildConqueredTheCastle, OldOwnerName, e_FontTypeNames.FONTTYPE_GUILD))
        Call SendData(SendTarget.ToGuildMembers, UserList(UserIndex).GuildIndex, PrepareMessagePlayWave(43, NO_3D_SOUND, NO_3D_SOUND))

        Call LogCastles(UserList(UserIndex).Name & " <" & NewOwnerName & "> conquered the castle from guild " & OldOwnerName & " on map " & .Location.Map)

        .GuildOwner = UserList(UserIndex).GuildIndex

        Call CloseCastleGates(CastleId)

        .InsideHP = CastleInsideHP
        .InCombatCounter = 0

        Call SetCastleState(CastleId, e_CastleState.Rebuilding)
        
        Call SendData(SendTarget.toMap, .LocationInside.Map, PrepareMessage_ShowCastleInside(CastleId))

        ' Kick all but owners
        Dim LoopC As Integer
        Dim tUser As Integer

        For LoopC = 1 To ConnGroups(.LocationInside.Map).CountEntrys
            tUser = ConnGroups(.LocationInside.Map).UserEntrys(LoopC)
            If UserList(tUser).GuildIndex = .GuildOwner Then
                ' Refresh owners' names
                Call RefreshCharStatus(tUser)
            Else
                Call WarpToLegalPos(tUser, .Location.Map, .Location.X, .Location.Y + 1, True)
            End If
        Next
        For LoopC = 1 To ConnGroups(.LocationWarp.Map).CountEntrys
            tUser = ConnGroups(.LocationWarp.Map).UserEntrys(LoopC)
            If UserList(tUser).GuildIndex <> .GuildOwner Then
                Call WarpToLegalPos(tUser, .Location.Map, .Location.X, .Location.Y + 1, True)
            End If
        Next

    End With

    Exit Sub
ConqueredCastle_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.ConqueredCastle", Erl)
End Sub

Private Sub OpenCastleGates(ByVal CastleId As Byte)
    On Error GoTo OpenCastleGates_Err
    
    With Castles(CastleId)
    
        ' Notify guild
        Call SendData(SendTarget.ToGuildMembers, .GuildOwner, PrepareMessageLocaleMsg(MsgCastleGatesOpened, "", e_FontTypeNames.FONTTYPE_GUILD))
    
        ' Unlock door positions
        Call BlockAndInform(.Location.Map, .Location.X, .Location.Y, 0)

        If CastleTypes(.CastleType).Width Mod 2 = 0 Then
            Call BlockAndInform(.Location.Map, .Location.X + 1, .Location.Y, 0)
        End If

    End With
    
    Exit Sub
OpenCastleGates_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.OpenCastleGates", Erl)
End Sub

Private Sub CloseCastleGates(ByVal CastleId As Byte)
    On Error GoTo CloseCastleGates_Err
    
    With Castles(CastleId)
    
        ' Lock door positions
        Call BlockAndInform(.Location.Map, .Location.X, .Location.Y, 1)

        If CastleTypes(.CastleType).Width Mod 2 = 0 Then
            Call BlockAndInform(.Location.Map, .Location.X + 1, .Location.Y, 1)
        End If

    End With
    
    Exit Sub
CloseCastleGates_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.CloseCastleGates", Erl)
End Sub

Public Sub SendCastleToUser(ByVal UserIndex As Integer, ByVal Map As Integer)
    On Error GoTo SendCastleToUser_Err
    
    Dim CastleId As Byte
    CastleId = FindCastleByMap(Map)
        
    If CastleId = 0 Then Exit Sub
    
    With Castles(CastleId)
        If .CastleType = 0 Then Exit Sub

        Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessage_ShowCastle(CastleId))
    End With
    
    Exit Sub
SendCastleToUser_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.SendCastleToUser", Erl)
End Sub

Private Function GetCastleHPPercentage(ByVal CastleId As Byte, ByVal Inside As Boolean) As Byte
    On Error GoTo GetCastleHPPercentage_Err

    With Castles(CastleId)
        If .CastleType = 0 Then Exit Function

        If Inside Then
            If CastleInsideHP = 0 Then Exit Function
            GetCastleHPPercentage = .InsideHP / CastleInsideHP * 100
            Exit Function
        End If

        Select Case .State
            Case e_CastleState.Built
                If CastleTypes(.CastleType).MaxHP = 0 Then Exit Function
                GetCastleHPPercentage = .HP / CastleTypes(.CastleType).MaxHP * 100

            Case e_CastleState.Building
                GetCastleHPPercentage = 100 - 100 * .WaitTime / CastleBuildTimeSeconds
                
            Case e_CastleState.Upgrading
                GetCastleHPPercentage = 0 ' TODO

            Case e_CastleState.Destroyed
                GetCastleHPPercentage = 100
                
            Case e_CastleState.Rebuilding
                GetCastleHPPercentage = 100 - 100 * .WaitTime / CastleRebuildTimeSeconds
        End Select

    End With

    Exit Function
GetCastleHPPercentage_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.GetCastleHPPercentage", Erl)
End Function

Public Function IsCastle(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
    On Error GoTo IsCastle_Err

    With MapData(Map, X, Y)
        IsCastle = .trigger = e_Trigger.CASTLE Or .trigger = e_Trigger.CASTLE_CENTER
    End With

    Exit Function
IsCastle_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.IsCastle", Erl)
End Function

Public Sub InfoCastle(ByVal UserIndex As Integer, ByVal Map As Integer)
    On Error GoTo InfoCastle_Err

    Dim CastleId As Byte
    CastleId = FindCastleByMap(Map)

    If CastleId = 0 Then
        CastleId = FindCastleInsideByMap(Map)
        If CastleId = 0 Then Exit Sub
        Call InfoCastleInside(UserIndex, CastleId)
        Exit Sub
    End If
    
    With Castles(CastleId)
        If .CastleType = 0 Then Exit Sub
        
        Dim State As Integer
        Select Case .State
            Case e_CastleState.Built
                If .HP = CastleTypes(.CastleType).MaxHP Then
                    State = MsgCastleIntact
                Else
                    State = MsgCastleDamaged
                End If
            Case e_CastleState.Building, e_CastleState.Rebuilding
                State = MsgCastleBuilding
            Case e_CastleState.Upgrading
                ' TODO
            Case e_CastleState.Destroyed
                If .WaitTime > 0 Then
                    State = MsgCastleDestroyed
                Else
                    State = MsgCastleCollapsed
                End If
        End Select

        Call WriteLocaleMsg(UserIndex, MsgCastleInfo, e_FontTypeNames.FONTTYPE_EJECUCION, GuildName(.GuildOwner) & "¨^" & State)
    End With

    Exit Sub
InfoCastle_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.InfoCastle", Erl)
End Sub

Public Sub InfoCastleInside(ByVal UserIndex As Integer, ByVal CastleId As Byte)
    On Error GoTo InfoCastleInside_Err

    With Castles(CastleId)
        If .CastleType = 0 Then Exit Sub
        
        Dim State As Integer
        If .InsideHP = CastleInsideHP Then
            State = MsgCastleIntact
        Else
            State = MsgCastleDamaged
        End If

        Call WriteLocaleMsg(UserIndex, MsgCastleInsideInfo, e_FontTypeNames.FONTTYPE_EJECUCION, "^" & State)
    End With

    Exit Sub
InfoCastleInside_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.InfoCastleInside", Erl)
End Sub

Private Sub PutCastleOnMap(ByVal CastleId As Byte, Optional ByVal NotifyUsers As Boolean = True)
    On Error GoTo PutCastleOnMap_Err

    With Castles(CastleId)
    
        Dim Width As Byte:              Width = CastleTypes(.CastleType).Width
        Dim Height As Byte:             Height = CastleTypes(.CastleType).Height
    
        Dim CastleTop As Integer:       CastleTop = .Location.Y - Height + 1
        Dim CastleBottom As Integer:    CastleBottom = .Location.Y
        Dim CastleLeft As Integer:      CastleLeft = .Location.X - (Width - 1) \ 2
        Dim CastleRight As Integer:     CastleRight = .Location.X + Width \ 2
    
        ' Add trigger and block
        Dim X As Integer: Dim Y As Integer
        For Y = CastleTop To CastleBottom
            For X = CastleLeft To CastleRight
                MapData(.Location.Map, X, Y).trigger = e_Trigger.CASTLE
                If NotifyUsers Then
                    Call BlockAndInform(.Location.Map, X, Y, 1)
                Else
                    MapData(.Location.Map, X, Y).Blocked = e_Block.ALL_SIDES Or e_Block.DYNAMIC
                End If
            Next
        Next

        ' Move any player and respawn npcs (after blocking all tiles)
        For Y = CastleTop To CastleBottom
            For X = CastleLeft To CastleRight
                If MapData(.Location.Map, X, Y).NpcIndex > 0 Then
                    Dim tNPC As t_Npc
                    tNPC = NpcList(MapData(.Location.Map, X, Y).NpcIndex)
                    Call QuitarNPC(MapData(.Location.Map, X, Y).NpcIndex, eBuildCastle)
                    Call ReSpawnNpc(tNPC)
                
                ElseIf MapData(.Location.Map, X, Y).UserIndex > 0 Then
                    Dim tUser As Integer
                    tUser = MapData(.Location.Map, X, Y).UserIndex
                    Call WarpToLegalPos(tUser, UserList(tUser).Pos.Map, UserList(tUser).Pos.X, UserList(tUser).Pos.Y)
                End If
            Next
        Next

        ' Castle center
        MapData(.Location.Map, .Location.X, .Location.Y).trigger = e_Trigger.CASTLE_CENTER
        
        ' Tile Exit
        MapData(.Location.Map, .Location.X, .Location.Y).TileExit.Map = .LocationInside.Map
        MapData(.Location.Map, .Location.X, .Location.Y).TileExit.X = .LocationInside.X
        MapData(.Location.Map, .Location.X, .Location.Y).TileExit.Y = .LocationInside.Y
        
        ' Tile Exit for double door
        If Width Mod 2 = 0 Then
            MapData(.Location.Map, .Location.X + 1, .Location.Y).TileExit.Map = .LocationInside.Map
            MapData(.Location.Map, .Location.X + 1, .Location.Y).TileExit.X = .LocationInside.X + 1
            MapData(.Location.Map, .Location.X + 1, .Location.Y).TileExit.Y = .LocationInside.Y
        End If
    
    End With

    Exit Sub
PutCastleOnMap_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.PutCastleOnMap", Erl)
End Sub

Private Sub RemoveCastleFromMap(ByVal CastleId As Byte)
    On Error GoTo RemoveCastleFromMap_Err

    With Castles(CastleId)
    
        If .CastleType = 0 Then Exit Sub
    
        Dim Width As Byte:              Width = CastleTypes(.CastleType).Width
        Dim Height As Byte:             Height = CastleTypes(.CastleType).Height
    
        Dim CastleTop As Integer:       CastleTop = .Location.Y - Height + 1
        Dim CastleBottom As Integer:    CastleBottom = .Location.Y
        Dim CastleLeft As Integer:      CastleLeft = .Location.X - (Width - 1) \ 2
        Dim CastleRight As Integer:     CastleRight = .Location.X + Width \ 2
    
        ' Remove trigger and block
        Dim X As Integer: Dim Y As Integer
        For Y = CastleTop To CastleBottom
            For X = CastleLeft To CastleRight
                MapData(.Location.Map, X, Y).trigger = e_Trigger.nada
                Call BlockAndInform(.Location.Map, X, Y, 0)
            Next
        Next

        ' Tile Exit
        MapData(.Location.Map, .Location.X, .Location.Y).TileExit.Map = 0
        
        ' Tile Exit for double door
        If Width Mod 2 = 0 Then
            MapData(.Location.Map, .Location.X + 1, .Location.Y).TileExit.Map = 0
        End If
    
    End With

    Exit Sub
RemoveCastleFromMap_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.RemoveCastleFromMap", Erl)
End Sub

Private Sub LoadCastlesDB()
    On Error GoTo LoadCastlesDB_Err
    
    Dim RS As Recordset
    Set RS = Query("SELECT id, guild_id, castle_type, state, max(timestamp - CAST(strftime('%s') as INT), 0) as wait_time FROM castle")
    If RS Is Nothing Then Exit Sub
    If RS.RecordCount = 0 Then Exit Sub
    
    Dim CastleId As Byte
    
    While Not RS.EOF

        CastleId = RS!ID

        With Castles(CastleId)

            .GuildOwner = RS!guild_id
            .CastleType = RS!castle_type
            .State = RS!State
            .WaitTime = RS!wait_Time
            .HP = CastleTypes(.CastleType).MaxHP
    
            Call PutCastleOnMap(CastleId, False)

        End With
        
        RS.MoveNext

    Wend

    Exit Sub
LoadCastlesDB_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.LoadCastlesDB", Erl)
End Sub

Private Sub SaveCastleDB(ByVal CastleId As Byte)
    On Error GoTo SaveCastleDB_Err

    If CastleId = 0 Then Exit Sub

    With Castles(CastleId)
        Call Query( _
            "INSERT OR REPLACE INTO castle (id, guild_id, castle_type, state, timestamp) VALUES (?, ?, ?, ?, CAST(strftime('%s') as INT) + ?)", _
            CastleId, .GuildOwner, .CastleType, .State, .WaitTime _
        )
    End With

    Exit Sub
SaveCastleDB_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.SaveCastleDB", Erl)
End Sub

Private Sub SetCastleState(ByVal CastleId As Byte, State As e_CastleState)
    On Error GoTo SetCastleState_Err

    With Castles(CastleId)

        .State = State

        Select Case State
            Case e_CastleState.Built
                .HP = CastleTypes(.CastleType).MaxHP
                .InsideHP = CastleInsideHP
                .WaitTime = 0
                .InCombatCounter = 0

            Case e_CastleState.Building
                .HP = 0
                .InsideHP = CastleInsideHP
                .WaitTime = CastleBuildTimeSeconds

            Case e_CastleState.Upgrading
                ' TODO
                .WaitTime = 123
                .InCombatCounter = 0

            Case e_CastleState.Destroyed
                .InsideHP = CastleInsideHP
                .WaitTime = CastleMaxInvasionTimeSeconds

            Case e_CastleState.Rebuilding
                .InsideHP = CastleInsideHP
                .InCombatCounter = 0
                .WaitTime = CastleRebuildTimeSeconds
        End Select

        Call SaveCastleDB(CastleId)

        Call SendData(SendTarget.ToCastleArea, CastleId, PrepareMessage_ShowCastle(CastleId))
        
    End With
    
    Exit Sub
SetCastleState_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.SetCastleState", Erl)
End Sub

Public Sub UpdateCastlesTimer()
    On Error GoTo UpdateCastlesTimer_Err
    
    Dim CastleId As Byte
    For CastleId = 1 To UBound(Castles)

        With Castles(CastleId)
            If .CastleType <> 0 Then
                Call UpdateInCombat(CastleId)
                Call RegenerateCastleInside(CastleId)
            
                Select Case .State
                    Case e_CastleState.Built
                        Call UpdateBuiltCastle(CastleId)
                    Case e_CastleState.Building, e_CastleState.Upgrading, e_CastleState.Rebuilding
                        Call UpdateConstructingCastle(CastleId)
                    Case e_CastleState.Destroyed
                        Call UpdateDestroyedCastle(CastleId)
                End Select
            End If
        End With

    Next

    Exit Sub
UpdateCastlesTimer_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.UpdateCastlesTimer", Erl)
End Sub

Private Sub UpdateBuiltCastle(ByVal CastleId As Byte)
    On Error GoTo UpdateBuiltCastle_Err

    Call RegenerateCastle(CastleId)

    Exit Sub
UpdateBuiltCastle_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.UpdateBuiltCastle", Erl)
End Sub

Private Sub UpdateConstructingCastle(ByVal CastleId As Byte)
    On Error GoTo UpdateConstructingCastle_Err

    With Castles(CastleId)

        If .WaitTime > 0 Then
            .WaitTime = .WaitTime - 1
            
            Call SendData(SendTarget.ToCastleArea, CastleId, PrepareMessage_UpdateCastleHP(GetCastleHPPercentage(CastleId, False)))
        Else
            ' Construction done!
            Call SendData(SendTarget.ToGuildMembers, .GuildOwner, PrepareMessageLocaleMsg(MsgCastleDoneBuilding, "", e_FontTypeNames.FONTTYPE_GUILD))
            
            Call SetCastleState(CastleId, e_CastleState.Built)
        End If

    End With

    Exit Sub
UpdateConstructingCastle_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.UpdateConstructingCastle", Erl)
End Sub

Private Sub UpdateDestroyedCastle(ByVal CastleId As Byte)
    On Error GoTo UpdateConstructingCastle_Err

    With Castles(CastleId)

        If .WaitTime > 0 Then
            .WaitTime = .WaitTime - 1

            ' Notify people X sec before closing
            If CastleNotifyCollapseSecondsBefore > 0 And .WaitTime = CastleMaxInvasionTimeSeconds + CastleNotifyCollapseSecondsBefore - CastleCollapseTimeSeconds Then
                Call SendData(SendTarget.toMap, .LocationInside.Map, PrepareMessageLocaleMsg(MsgCastleEntranceCollapsing, "", e_FontTypeNames.FONTTYPE_FIGHT))

            ' Entrance collapsed!
            ElseIf .WaitTime = CastleMaxInvasionTimeSeconds - CastleCollapseTimeSeconds Then
                ' Block door tiles
                Call CloseCastleGates(CastleId)
                ' Notify owner guild
                Call SendData(SendTarget.toMap, .LocationInside.Map, PrepareMessageLocaleMsg(MsgCastleEntranceCollapsed, "", e_FontTypeNames.FONTTYPE_FIGHT))
                
            ' After entrance collapsed
            ElseIf .WaitTime < CastleMaxInvasionTimeSeconds - CastleCollapseTimeSeconds Then
                ' Count if there are more enemies inside the castle
                Dim LoopC As Integer
                Dim tUser As Integer

                For LoopC = 1 To ConnGroups(.LocationInside.Map).CountEntrys
                    tUser = ConnGroups(.LocationInside.Map).UserEntrys(LoopC)
                    If UserList(tUser).GuildIndex <> .GuildOwner Then
                        Exit Sub
                    End If
                Next

                ' No more enemies! invasion is over
                Call CastleInvasionFailed(CastleId)
            End If

            Exit Sub
        End If

        ' Castle invasion time is over!
        Call CastleInvasionFailed(CastleId)

    End With

    Exit Sub
UpdateConstructingCastle_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.UpdateConstructingCastle", Erl)
End Sub

Private Sub CastleInvasionFailed(ByVal CastleId As Byte)
    On Error GoTo CastleInvasionFailed_Err

    With Castles(CastleId)

        ' Notify all in map
        Call SendData(SendTarget.toMap, .LocationInside.Map, PrepareMessageLocaleMsg(MsgCastleInvasionFailed, "", e_FontTypeNames.FONTTYPE_GUILD))
        Call SendData(SendTarget.toMap, .LocationInside.Map, PrepareMessagePlayWave(48, NO_3D_SOUND, NO_3D_SOUND))

        Call LogCastles("Invasion failed: castle from guild " & GuildName(.GuildOwner) & " on map " & .Location.Map)

        Call CloseCastleGates(CastleId)

        Call SetCastleState(CastleId, e_CastleState.Rebuilding)

        ' Kick all but owners
        Dim LoopC As Integer
        Dim tUser As Integer

        For LoopC = 1 To ConnGroups(.LocationInside.Map).CountEntrys
            tUser = ConnGroups(.LocationInside.Map).UserEntrys(LoopC)
            If UserList(tUser).GuildIndex = .GuildOwner Then
                ' Refresh owners' names
                Call RefreshCharStatus(tUser)
            Else
                Call WarpToLegalPos(tUser, .Location.Map, .Location.X, .Location.Y + 1, True)
            End If
        Next
        For LoopC = 1 To ConnGroups(.LocationWarp.Map).CountEntrys
            tUser = ConnGroups(.LocationWarp.Map).UserEntrys(LoopC)
            If UserList(tUser).GuildIndex <> .GuildOwner Then
                Call WarpToLegalPos(tUser, .Location.Map, .Location.X, .Location.Y + 1, True)
            End If
        Next

        Call SendData(SendTarget.toMap, .LocationInside.Map, PrepareMessage_ShowCastleInside(CastleId))
    
    End With

    Exit Sub
CastleInvasionFailed_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.CastleInvasionFailed", Erl)
End Sub

Private Sub UpdateInCombat(ByVal CastleId As Byte)
    On Error GoTo UpdateInCombat_Err

    With Castles(CastleId)
        If .InCombatCounter > 0 Then
            .InCombatCounter = .InCombatCounter - 1
            
            ' In combat! Notify guild members
            If .InCombatCounter Mod 2 = 1 Then
                Call SendData(SendTarget.ToGuildMembers, .GuildOwner, PrepareMessageLocaleMsg(MsgCastleIsUnderAttack, "", e_FontTypeNames.FONTTYPE_GUILD))
            End If
        End If
    End With

    Exit Sub
UpdateInCombat_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.UpdateInCombat", Erl)
End Sub

Private Sub RegenerateCastle(ByVal CastleId As Byte)
    On Error GoTo RegenerateCastle_Err

    With Castles(CastleId)
        If .HP < CastleTypes(.CastleType).MaxHP Then
            .HP = .HP + CastleRegenerationPerSecond
            If .HP > CastleTypes(.CastleType).MaxHP Then
                .HP = CastleTypes(.CastleType).MaxHP
            End If
    
            Call SendData(SendTarget.ToCastleArea, CastleId, PrepareMessage_UpdateCastleHP(GetCastleHPPercentage(CastleId, False)))
        End If
    End With
    
    Exit Sub
RegenerateCastle_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.RegenerateCastle", Erl)
End Sub

Private Sub RegenerateCastleInside(ByVal CastleId As Byte)
    On Error GoTo RegenerateCastle_Err

    With Castles(CastleId)
        If .InsideHP < CastleInsideHP Then
            .InsideHP = .InsideHP + CastleInsideRegenerationPerSecond
            If .InsideHP > CastleInsideHP Then
                .InsideHP = CastleInsideHP
            End If
    
            Call SendData(SendTarget.toMap, .LocationInside.Map, PrepareMessage_UpdateCastleHP(GetCastleHPPercentage(CastleId, True)))
        End If
    End With
    
    Exit Sub
RegenerateCastle_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.RegenerateCastle", Erl)
End Sub

Public Sub StartWarpToCastle(ByVal UserIndex As Integer)
    On Error GoTo StartWarpToCastle_Err
    
    With UserList(UserIndex)
    
        If .flags.EnReto Then
            Call WriteLocaleMsg(UserIndex, MsgCantWarpInThisMoment, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .flags.EnConsulta Then
            Call WriteLocaleMsg(UserIndex, MsgCantWarpInThisMoment, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .flags.jugando_captura = 1 Then
            Call WriteLocaleMsg(UserIndex, MsgCantWarpInThisMoment, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .flags.EnTorneo Then
            Call WriteLocaleMsg(UserIndex, MsgCantWarpInThisMoment, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = CARCEL Then
            Call WriteLocaleMsg(UserIndex, MsgCantWarpInThisMoment, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    
        If .GuildIndex = 0 Then
            Call WriteLocaleMsg(UserIndex, MsgYouNeedToBeInAGuild, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Dim CastleId As Byte
        CastleId = FindCastleByGuild(.GuildIndex)
        
        If CastleId = 0 Then
            Call WriteLocaleMsg(UserIndex, MsgYourGuildDoesntHaveCastle, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Castles(CastleId).LocationWarp.Map = .Pos.Map Then
            Call WriteLocaleMsg(UserIndex, MsgYouAreAlreadyOnCastle, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Castles(CastleId).InCombatCounter = 0 Then
            If .Counters.WarpCastleCooldown > 0 Then
                Call WriteLocaleMsg(UserIndex, MsgMustWaitXSecondsForWarpCastle, e_FontTypeNames.FONTTYPE_INFO, .Counters.WarpCastleCooldown)
                Exit Sub
            ElseIf CastleWarpCooldown < 0 Then
                Call WriteLocaleMsg(UserIndex, MsgCantWarpToCastleIfNotAttacked, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        ElseIf CastleWarpMaxHPPercentage > 0 Then
            If GetCastleHPPercentage(CastleId, True) > CastleWarpMaxHPPercentage And _
                ((Castles(CastleId).State <> e_CastleState.Built And Castles(CastleId).State <> e_CastleState.Upgrading) Or GetCastleHPPercentage(CastleId, False) > CastleWarpMaxHPPercentage) _
            Then
                Call WriteLocaleMsg(UserIndex, MsgCantWarpToCastleIfHpIsMoreThan, e_FontTypeNames.FONTTYPE_INFO, CastleWarpMaxHPPercentage)
                Exit Sub
            End If
        End If
        
        ' Show progress bar and fx
        .Counters.TimerBarra = CastleWarpTimeSeconds
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(.Char.charindex, e_ParticulasIndex.Runa, .Counters.TimerBarra * 100, False, , .Pos.X, .Pos.Y))
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(.Char.charindex, .Counters.TimerBarra, e_AccionBarra.Hogar))
        Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(MsgWillArriveInCastleInXSeconds, .Counters.TimerBarra, e_FontTypeNames.FONTTYPE_New_Gris))
        
        .Accion.Particula = e_ParticulasIndex.Runa
        .Accion.TipoAccion = e_AccionBarra.GoCastle
        .Accion.AccionPendiente = True
    
    End With
    
    Exit Sub
StartWarpToCastle_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.StartWarpToCastle", Erl)
End Sub

Public Sub WarpToCastle(ByVal UserIndex As Integer)
    On Error GoTo WarpToCastle_Err
    
    With UserList(UserIndex)
    
        If .flags.EnReto Then
            Call WriteLocaleMsg(UserIndex, MsgCantWarpInThisMoment, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .flags.EnConsulta Then
            Call WriteLocaleMsg(UserIndex, MsgCantWarpInThisMoment, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .flags.jugando_captura = 1 Then
            Call WriteLocaleMsg(UserIndex, MsgCantWarpInThisMoment, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .flags.EnTorneo Then
            Call WriteLocaleMsg(UserIndex, MsgCantWarpInThisMoment, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = CARCEL Then
            Call WriteLocaleMsg(UserIndex, MsgCantWarpInThisMoment, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    
        If .GuildIndex = 0 Then
            Call WriteLocaleMsg(UserIndex, MsgYouNeedToBeInAGuild, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Dim CastleId As Byte
        CastleId = FindCastleByGuild(.GuildIndex)
        
        If CastleId = 0 Then
            Call WriteLocaleMsg(UserIndex, MsgYourGuildDoesntHaveCastle, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Castles(CastleId).LocationWarp.Map = .Pos.Map Then
            Call WriteLocaleMsg(UserIndex, MsgYouAreAlreadyOnCastle, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        .Counters.WarpCastleCooldown = CastleWarpCooldown

        If .flags.Nadando <> 0 Or .flags.Navegando <> 0 Then
            Call DoNavega(UserIndex, ObjData(.invent.BarcoObjIndex), .invent.BarcoSlot)
        End If

        Call WriteLocaleMsg(UserIndex, MsgYouWarpedToCastle, e_FontTypeNames.FONTTYPE_INFO)

    End With

    With Castles(CastleId)
        Call WarpToLegalPos(UserIndex, .LocationWarp.Map, .LocationWarp.X, .LocationWarp.Y, True)
    End With

    Exit Sub
WarpToCastle_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.WarpToCastle", Erl)
End Sub

Private Function GetCastleGrhIndex(ByVal CastleId As Byte) As Long
    On Error GoTo GetCastleGrhIndex_Err

    With CastleTypes(Castles(CastleId).CastleType)
    
        Select Case Castles(CastleId).State
            Case e_CastleState.Built
                GetCastleGrhIndex = .GrhIndexBuilt
            Case e_CastleState.Building, e_CastleState.Upgrading, e_CastleState.Rebuilding
                GetCastleGrhIndex = .GrhIndexBuilding
            Case e_CastleState.Destroyed
                GetCastleGrhIndex = .GrhIndexDestroyed
        End Select
    
    End With

    Exit Function
GetCastleGrhIndex_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.GetCastleGrhIndex", Erl)
End Function

Public Sub TryOpenCastleGate(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal Entering As Boolean)
    On Error GoTo TryOpenCastleGate_Err

    Dim CastleId As Byte
    If Entering Then
        CastleId = FindCastleByMap(Map)
    Else
        CastleId = FindCastleInsideByMap(Map)
    End If
    If CastleId = 0 Then Exit Sub

    With Castles(CastleId)
        If .CastleType = 0 Then Exit Sub
        If .GuildOwner <> UserList(UserIndex).GuildIndex Then Exit Sub
        If .State = e_CastleState.Building Then Exit Sub

        If Entering Then
            Call WarpToLegalPos(UserIndex, .LocationInside.Map, .LocationInside.X, .LocationInside.Y, False)
            Call WriteLocaleMsg(UserIndex, MsgEnteredCastle, e_FontTypeNames.FONTTYPE_INFO)
        Else
            Call WarpToLegalPos(UserIndex, .Location.Map, .Location.X, .Location.Y + 1, False)
            Call WriteLocaleMsg(UserIndex, MsgExitedCastle, e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With

    Call ClearAttackerNpc(UserIndex)

    Exit Sub
TryOpenCastleGate_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.TryOpenCastleGate", Erl)
End Sub

Private Function GetCastleSndIndex(ByVal CastleId As Byte, ByVal Inside As Boolean) As Integer
    On Error GoTo GetCastleSndIndex_Err
    
    GetCastleSndIndex = CastleId
    If Inside Then
        Call SetIntMask(GetCastleSndIndex, CASTLE_INSIDE_MASK)
    End If
    
    Exit Function
GetCastleSndIndex_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.GetCastleSndIndex", Erl)
End Function

Private Sub GetCastleIdAndInside(ByVal CastleSndIndex As Integer, ByRef CastleId As Byte, ByRef Inside As Boolean)
    On Error GoTo GetCastleIdAndInside_Err
    
    CastleId = CastleSndIndex And Not CASTLE_INSIDE_MASK
    Inside = IsIntSet(CastleSndIndex, CASTLE_INSIDE_MASK)
    
    Exit Sub
GetCastleIdAndInside_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.GetCastleIdAndInside", Erl)
End Sub

Public Sub GetCastleArea(ByVal CastleSndIndex As Integer, ByRef Map As Integer, ByRef AreaX As Integer, ByRef AreaY As Integer)
    On Error GoTo GetCastleArea_Err
    
    Dim CastleId As Byte
    Dim Inside As Boolean
    Call GetCastleIdAndInside(CastleSndIndex, CastleId, Inside)
    
    If CastleId = 0 Then Exit Sub
    
    With Castles(CastleId)
        If .CastleType = 0 Then Exit Sub

        If Inside Then
            Map = .LocationInside.Map
            Call GetAreaByPos(.LocationInside.X, .LocationInside.Y, AreaX, AreaY)
        Else
            Map = .Location.Map
            Call GetAreaByPos(.Location.X, .Location.Y, AreaX, AreaY)
        End If
    End With
    
    Exit Sub
GetCastleArea_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.GetCastleArea", Erl)
End Sub

Public Sub CheckConnectInsideCastle(ByVal UserIndex As Integer)
    On Error GoTo CheckConnectInsideCastle_Err

    Dim CastleId As Byte
    CastleId = FindCastleInsideByMap(UserList(UserIndex).Pos.Map)
    If CastleId = 0 Then Exit Sub
    If Castles(CastleId).CastleType = 0 Then Exit Sub
    
    ' If not owner, warp him outside
    If Castles(CastleId).GuildOwner <> UserList(UserIndex).GuildIndex Then
        With Castles(CastleId)
            Call WarpToLegalPos(UserIndex, .Location.Map, .Location.X, .Location.Y + 1, True)
        End With
        Exit Sub
    End If
    
    ' Show HP bar
    Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessage_ShowCastleInside(CastleId))

    Exit Sub
CheckConnectInsideCastle_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.CheckConnectInsideCastle", Erl)
End Sub

Public Sub CheckEnteredCastle(ByVal UserIndex As Integer)
    On Error GoTo CheckEnteredCastle_Err

    Dim CastleId As Byte
    CastleId = FindCastleInsideByMap(UserList(UserIndex).Pos.Map)
    If CastleId = 0 Then Exit Sub
    If Castles(CastleId).CastleType = 0 Then Exit Sub
    
    ' Show HP bar
    Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessage_ShowCastleInside(CastleId))

    Exit Sub
CheckEnteredCastle_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.CheckEnteredCastle", Erl)
End Sub

Public Sub CheckDieInsideCastle(ByVal UserIndex As Integer)
    On Error GoTo CheckDieInsideCastle_Err

    Dim CastleId As Byte
    CastleId = FindCastleInsideByMap(UserList(UserIndex).Pos.Map)
    If CastleId = 0 Then Exit Sub

    With Castles(CastleId)
        If .CastleType = 0 Then Exit Sub

        If .GuildOwner = UserList(UserIndex).GuildIndex Then
            ' If owner, warp him inside
            Call WarpToLegalPos(UserIndex, .LocationWarp.Map, .LocationWarp.X, .LocationWarp.Y, True)
        Else
            ' If invader, warp him ouside
            Call WarpToLegalPos(UserIndex, .Location.Map, .Location.X, .Location.Y + 1, True)
        End If

    End With

    Exit Sub
CheckDieInsideCastle_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.CheckDieInsideCastle", Erl)
End Sub

Public Function CanEnterCastle(ByVal UserIndex As Integer, ByVal Map As Integer) As Boolean
    On Error GoTo CanEnterCastle_Err
    
    CanEnterCastle = True

    Dim CastleId As Byte
    CastleId = FindCastleInsideByMap(Map)
    If CastleId = 0 Then Exit Function
    If Castles(CastleId).CastleType = 0 Then Exit Function

    With UserList(UserIndex)
        If .flags.Muerto = 1 Then
            If .GuildIndex <> Castles(CastleId).GuildOwner Then
                CanEnterCastle = False
                
                If .flags.UltimoMensaje <> 107 Then
                    Call WriteLocaleMsg(UserIndex, MsgYouMustBeAliveToEnterTheCastle, e_FontTypeNames.FONTTYPE_INFO)
                    .flags.UltimoMensaje = 107
                End If
            End If
        End If
    End With

    Exit Function
CanEnterCastle_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.CanEnterCastle", Erl)
End Function

Public Function GetCastleStatus(ByVal UserIndex As Integer) As Byte
    On Error GoTo GetCastleStatus_Err

    With UserList(UserIndex)

        GetCastleStatus = 0

        Dim CastleId As Byte
        CastleId = FindCastleInsideByMap(.Pos.Map)

        If CastleId = 0 Then Exit Function
        If Castles(CastleId).CastleType = 0 Then Exit Function
        If Castles(CastleId).State <> e_CastleState.Destroyed Then Exit Function

        If .GuildIndex = Castles(CastleId).GuildOwner Then
            GetCastleStatus = 5
        Else
            GetCastleStatus = 4
        End If

    End With

    Exit Function
GetCastleStatus_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.GetCastleStatus", Erl)
End Function

Public Function CastleNpcCanAttack(ByVal UserIndex As Integer, ByVal Map As Integer) As Boolean
    On Error GoTo CastleNpcCanAttack_Err
    
    CastleNpcCanAttack = True

    Dim CastleId As Byte
    CastleId = FindCastleInsideByMap(Map)
    If CastleId = 0 Then Exit Function
    If Castles(CastleId).CastleType = 0 Then Exit Function

    CastleNpcCanAttack = UserList(UserIndex).GuildIndex <> Castles(CastleId).GuildOwner

    Exit Function
CastleNpcCanAttack_Err:
    Call TraceError(Err.Number, Err.Description, "modCastles.CastleNpcCanAttack", Erl)
End Function

#If DIRECT_PLAY = 0 Then
Public Sub WriteCastleData(ByRef Writer As Network.Writer, ByVal CastleId As Byte)
#Else
Public Sub WriteCastleData(ByRef Writer As clsNetWriter, ByVal CastleId As Byte)
#End If
        On Error GoTo WriteCastleData_Err

102     With Castles(CastleId)
104         Call Writer.WriteInt16(.Location.X)
106         Call Writer.WriteInt16(.Location.Y)
108         Call Writer.WriteInt8(.State)
        End With

110     With CastleTypes(Castles(CastleId).CastleType)
112         Call Writer.WriteInt32(GetCastleGrhIndex(CastleId))
114         Call Writer.WriteInt16(.GrhOffsetX)
116         Call Writer.WriteInt16(.GrhOffsetY)
118         Call Writer.WriteBool(.Width Mod 2 = 0) ' Is tiled even?
        End With

120     Call Writer.WriteInt8(GetCastleHPPercentage(CastleId, False))

        Exit Sub
WriteCastleData_Err:
        Call TraceError(Err.Number, Err.Description, "modCastles.WriteCastleData", Erl)
        ' Rethrow
        Err.raise Err.Number
End Sub

#If DIRECT_PLAY = 0 Then
Public Sub WriteCastleInsideData(ByRef Writer As Network.Writer, ByVal CastleId As Byte)
#Else
Public Sub WriteCastleInsideData(ByRef Writer As clsNetWriter, ByVal CastleId As Byte)
#End If
        On Error GoTo WriteCastleInsideData_Err

118     Call Writer.WriteInt8(GetCastleHPPercentage(CastleId, True))

        Exit Sub
WriteCastleInsideData_Err:
        Call TraceError(Err.Number, Err.Description, "modCastles.WriteCastleInsideData", Erl)
        ' Rethrow
        Err.raise Err.Number
End Sub
